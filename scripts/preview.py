"""Render local Office files into a self-contained Antigravity HTML artifact."""
import argparse
import base64
import csv
import hashlib
import html
import json
from pathlib import Path
import subprocess
import sys
import zipfile
import shutil

KINDS = {'.doc': 'Word', '.docx': 'Word', '.ppt': 'PowerPoint',
         '.pptx': 'PowerPoint', '.xlsx': 'Excel'}
IMAGES = {'Word': 'WINWORD.EXE', 'PowerPoint': 'POWERPNT.EXE', 'Excel': 'EXCEL.EXE'}
PLACEHOLDER = '<html><!-- OFFICE_PREVIEW_PENDING --></html>'


def digest(path):
    with path.open('rb') as stream:
        return hashlib.file_digest(stream, 'sha256').hexdigest()


def check_closed(kind):
    result = subprocess.run(['tasklist', '/FI', f'IMAGENAME eq {IMAGES[kind]}',
                             '/FO', 'CSV', '/NH'], capture_output=True, check=True)
    rows = csv.reader(result.stdout.decode(errors='replace').splitlines())
    if any(row and row[0].casefold() == IMAGES[kind].casefold() for row in rows):
        raise RuntimeError(f'{kind} is running. Close it yourself and retry; no process was stopped.')


def export_office(source, target):
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = document = None
    old_links = old_security = old_alerts = None
    old_ask_links = old_events = None
    kind = KINDS[source.suffix.lower()]
    sections = []
    try:
        check_closed(kind)
        app = win32com.client.DispatchEx(kind + '.Application')
        old_security = app.AutomationSecurity
        old_alerts = app.DisplayAlerts
        app.AutomationSecurity = 3
        if kind == 'Word':
            app.Visible = False
            app.DisplayAlerts = 0
            old_links = app.Options.UpdateLinksAtOpen
            app.Options.UpdateLinksAtOpen = False
            document = app.Documents.Open(FileName=str(source), ReadOnly=True,
                                          AddToRecentFiles=False, Visible=False)
            pdf = target / 'document.pdf'
            document.ExportAsFixedFormat(str(pdf), 17, OpenAfterExport=False)
            sections.append({'title': source.name, 'pdf': str(pdf)})
        elif kind == 'PowerPoint':
            app.DisplayAlerts = 1
            document = app.Presentations.Open(str(source), ReadOnly=True,
                                              Untitled=False, WithWindow=False)
            pdf = target / 'presentation.pdf'
            document.SaveCopyAs(str(pdf), 32)
            sections.append({'title': source.name, 'pdf': str(pdf)})
        else:
            app.Visible = False
            app.DisplayAlerts = False
            old_ask_links = app.AskToUpdateLinks
            old_events = app.EnableEvents
            app.AskToUpdateLinks = False
            app.EnableEvents = False
            document = app.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True,
                                         IgnoreReadOnlyRecommended=True, AddToMru=False)
            for index in range(1, document.Worksheets.Count + 1):
                sheet = document.Worksheets.Item(index)
                if sheet.Visible != -1:
                    continue
                pdf = target / f'sheet-{index:03d}.pdf'
                sheet.ExportAsFixedFormat(0, str(pdf), OpenAfterPublish=False)
                sections.append({'title': str(sheet.Name), 'pdf': str(pdf)})
        if not sections:
            raise RuntimeError('No visible pages or worksheets could be exported.')
        (target / 'sections.json').write_text(json.dumps(sections, ensure_ascii=False), encoding='utf-8')
    finally:
        try:
            if document is not None:
                if kind == 'PowerPoint':
                    document.Close()
                else:
                    document.Close(False)
        finally:
            if app is not None:
                if old_links is not None:
                    app.Options.UpdateLinksAtOpen = old_links
                if old_security is not None:
                    app.AutomationSecurity = old_security
                if old_alerts is not None:
                    app.DisplayAlerts = old_alerts
                if old_ask_links is not None:
                    app.AskToUpdateLinks = old_ask_links
                if old_events is not None:
                    app.EnableEvents = old_events
                remaining = (app.Presentations.Count if kind == 'PowerPoint' else
                             app.Documents.Count if kind == 'Word' else app.Workbooks.Count)
                if remaining == 0:
                    app.Quit()
            document = app = None
            pythoncom.CoUninitialize()


def make_html(source, target, max_pages):
    import pymupdf as fitz
    sections = json.loads((target / 'sections.json').read_text(encoding='utf-8'))
    pages = []
    pdf_paths = []
    total = 0
    for section in sections:
        pdf_paths.append(section['pdf'])
        with fitz.open(section['pdf']) as pdf:
            total += len(pdf)
            if total > max_pages:
                raise RuntimeError(f'{total}+ pages exceed --max-pages {max_pages}; PDFs retained in {target}')
            for index, page in enumerate(pdf):
                pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                pages.append({'section': section['title'], 'number': index + 1,
                              'image': base64.b64encode(pixmap.tobytes('png')).decode('ascii')})
    payload = json.dumps(pages, ensure_ascii=True).replace('<', '\\u003c').replace('>', '\\u003e').replace('&', '\\u0026')
    title = html.escape(source.name)
    template = '''<!doctype html><html lang="zh-CN"><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>__TITLE__</title>
<style>body{margin:0;padding:12px;background:var(--background);color:var(--foreground);font:14px system-ui}header{position:sticky;top:0;padding:12px;background:var(--card);border:1px solid var(--border);border-radius:10px}h1{font-size:16px;margin:0 0 10px;overflow-wrap:anywhere}nav{display:flex;gap:8px;flex-wrap:wrap;align-items:center}button,select{font:inherit;padding:7px;border:1px solid var(--border);border-radius:6px;color:var(--foreground);background:var(--card)}#stage{overflow:auto;margin-top:12px;text-align:center}img{width:100%;max-width:none;display:block;margin:auto}small{display:block;margin-top:8px;font-size:12px;color:var(--muted-foreground)}</style>
<header><h1>__TITLE__</h1><nav><select id="sections" aria-label="工作表"></select><button id="prev">上一页</button><span id="count"></span><button id="next">下一页</button><button id="less">缩小</button><button id="more">放大</button></nav><small>本地静态打印预览 · 原文件未修改</small></header><main id="stage"><img id="page" alt="文档页面"></main>
<script>const pages=__DATA__;let at=0,zoom=100;const sel=document.getElementById('sections'),img=document.getElementById('page');for(const s of [...new Set(pages.map(p=>p.section))]){const o=document.createElement('option');o.textContent=s;o.value=s;sel.append(o)}function draw(){const p=pages[at];img.src='data:image/png;base64,'+p.image;img.alt=p.section+' 第'+p.number+'页';img.style.width=zoom+'%';sel.value=p.section;document.getElementById('count').textContent=(at+1)+' / '+pages.length;document.getElementById('prev').disabled=at===0;document.getElementById('next').disabled=at===pages.length-1}document.getElementById('prev').onclick=()=>{if(at>0)at--;draw()};document.getElementById('next').onclick=()=>{if(at<pages.length-1)at++;draw()};document.getElementById('less').onclick=()=>{zoom=Math.max(25,zoom-25);draw()};document.getElementById('more').onclick=()=>{zoom=Math.min(300,zoom+25);draw()};sel.onchange=()=>{at=pages.findIndex(p=>p.section===sel.value);draw()};draw();</script></html>'''
    output = target / 'index.html'
    output.write_text(template.replace('__TITLE__', title).replace('__DATA__', payload), encoding='utf-8')
    return {'html': str(output), 'pdf': pdf_paths, 'pages': len(pages), 'sha256': digest(source)}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--input', required=True)
    parser.add_argument('--output-dir', required=True)
    parser.add_argument('--max-pages', type=int, default=100)
    parser.add_argument('--artifact-html', help='Previously registered placeholder HTML artifact')
    parser.add_argument('--export-worker', action='store_true', help=argparse.SUPPRESS)
    args = parser.parse_args()
    source = Path(args.input).resolve(strict=True)
    target = Path(args.output_dir).resolve()
    if not source.is_file() or source.suffix.lower() not in KINDS:
        raise ValueError('Supported local files: doc/docx, ppt/pptx, xlsx. Export legacy Excel files manually.')
    if source.suffix.lower() == '.xlsx':
        with zipfile.ZipFile(source) as archive:
            if any('macrosheet' in entry.lower() or 'vbaproject' in entry.lower()
                   for entry in archive.namelist()):
                raise ValueError('Macro-bearing Excel files require manual PDF export.')
    if args.export_worker:
        export_office(source, target)
        return
    if args.max_pages < 1:
        raise ValueError('--max-pages must be positive.')
    artifact = Path(args.artifact_html).resolve() if args.artifact_html else None
    if artifact and (artifact.suffix.lower() != '.html' or
                     artifact.read_text(encoding='utf-8').strip() != PLACEHOLDER):
        raise ValueError('--artifact-html must contain exactly the documented preview placeholder.')
    check_closed(KINDS[source.suffix.lower()])
    original = digest(source)
    target.mkdir(parents=True, exist_ok=False)
    try:
        worker = subprocess.Popen([sys.executable, str(Path(__file__).resolve()), '--export-worker',
                                   '--input', str(source), '--output-dir', str(target)])
        code = worker.wait(timeout=120)
        if code:
            raise RuntimeError(f'Office export worker exited with status {code}.')
    except subprocess.TimeoutExpired as error:
        raise RuntimeError('Office export is still pending after 120 seconds. Worker and Office were left running for cleanup; check Office before retrying.') from error
    if digest(source) != original:
        raise RuntimeError('Source file changed during conversion; preview not published.')
    result = make_html(source, target, args.max_pages)
    if artifact:
        if artifact.read_text(encoding='utf-8').strip() != PLACEHOLDER:
            raise RuntimeError('Artifact placeholder changed; refusing to overwrite it.')
        shutil.copyfile(result['html'], artifact)
        result['html'] = str(artifact)
    print(json.dumps(result, ensure_ascii=True))


if __name__ == '__main__':
    try:
        main()
    except Exception as error:
        print(f'{type(error).__name__}: {error}', file=sys.stderr)
        sys.exit(1)
