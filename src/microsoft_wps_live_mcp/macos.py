"""macOS desktop Office and WPS Office automation via AppleScript and native OOXML."""
from __future__ import annotations

import math
import os
from pathlib import Path
import re
import shutil
import struct
import subprocess
import sys
import tempfile
import uuid
import xml.etree.ElementTree as ET
import zipfile

from .discovery import MACOS_APPS

_EMU = 12700
_P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"


def run_osascript(script: str, timeout: float = 10.0) -> tuple[int, str, str]:
    """Execute AppleScript via osascript."""
    if sys.platform != "darwin":
        return 1, "", "Not running on macOS"
    try:
        proc = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        return proc.returncode, proc.stdout.strip(), proc.stderr.strip()
    except Exception as exc:
        return 1, "", str(exc)


def is_process_running(process_name: str) -> bool:
    """Check if process is running on macOS."""
    if sys.platform != "darwin":
        return False
    code, out, _ = run_osascript(f'tell application "System Events" to exists (process "{process_name}")')
    return code == 0 and out.lower() == "true"


def get_app_name(vendor: str, kind: str) -> str:
    """Return the macOS AppleScript application name."""
    if vendor == "microsoft":
        names = {"word": "Microsoft Word", "excel": "Microsoft Excel", "ppt": "Microsoft PowerPoint"}
        return names.get(kind, "Microsoft Word")
    return "WPS Office"


def _as_quote(text: str) -> str:
    return text.replace("\\", "\\\\").replace('"', '\\"')


def open_in_app(app_name: str, path: Path, visible: bool = True) -> bool:
    """Open document in macOS application."""
    if sys.platform != "darwin":
        return False
    path_str = _as_quote(str(path.resolve()))
    if "Microsoft" in app_name:
        cmd = f'tell application "{app_name}"\n'
        if visible:
            cmd += "    activate\n"
        cmd += f'    open POSIX file "{path_str}"\nend tell'
        code, _, _ = run_osascript(cmd)
        if code == 0:
            if visible:
                activate_app(app_name)
            return True
    open_args = ["open", "-a", app_name, str(path.resolve())]
    if not visible:
        open_args.append("-g")
    res = subprocess.run(open_args, capture_output=True, text=True)
    if res.returncode == 0 and visible:
        activate_app(app_name)
    return res.returncode == 0


def activate_app(app_name: str) -> bool:
    """Bring macOS application window to frontmost top view."""
    if sys.platform != "darwin":
        return False
    q_name = _as_quote(app_name)
    script = (
        f'tell application "{q_name}" to activate\n'
        'tell application "System Events"\n'
        f'    if exists (process "{q_name}") then\n'
        f'        set frontmost of process "{q_name}" to true\n'
        "    end if\n"
        "end tell"
    )
    code, _, _ = run_osascript(script)
    return code == 0


def focus_ppt_slide(app_name: str, slide: int) -> bool:
    """Navigate the frontmost Microsoft PowerPoint window on macOS to a 1-based slide."""
    if sys.platform != "darwin":
        return False
    if "Microsoft" not in app_name:
        return activate_app(app_name)
    q_name = _as_quote(app_name)
    script = (
        f'tell application "{q_name}"\n'
        "    activate\n"
        "    if (count of document windows) > 0 then\n"
        "        try\n"
        f"            set view type of document window 1 to slide view\n"
        f"            go to slide (slide {int(slide)} of active presentation) in view of document window 1\n"
        "        end try\n"
        "    end if\n"
        "end tell"
    )
    code, _, _ = run_osascript(script)
    return code == 0


def save_in_app(app_name: str, path: Path) -> bool:
    """Request active document save via AppleScript."""
    if sys.platform != "darwin":
        return False
    if "Microsoft" in app_name:
        if "PowerPoint" in app_name:
            code, _, _ = run_osascript(f'tell application "{_as_quote(app_name)}" to save active presentation')
        elif "Excel" in app_name:
            code, _, _ = run_osascript(f'tell application "{_as_quote(app_name)}" to save active workbook')
        else:
            code, _, _ = run_osascript(f'tell application "{_as_quote(app_name)}" to save active document')
        return code == 0
    return False


def close_in_app(app_name: str, path: Path) -> bool:
    """Close document via AppleScript."""
    if sys.platform != "darwin":
        return False
    q_name = _as_quote(app_name)
    q_path = _as_quote(str(path.resolve()))
    q_base = _as_quote(path.name)
    if "Microsoft" in app_name:
        if "PowerPoint" in app_name:
            script = (
                f'tell application "{q_name}"\n'
                f'    try\n        close (first presentation whose full name is "{q_path}") saving no\n'
                f'    on error\n        close (first presentation whose name is "{q_base}") saving no\n'
                "    end try\nend tell"
            )
        elif "Excel" in app_name:
            script = f'tell application "{q_name}" to close workbook "{q_base}" saving no'
        else:
            script = (
                f'tell application "{q_name}"\n'
                f'    try\n        close (first document whose full name is "{q_path}") saving no\n'
                f'    on error\n        close document "{q_base}" saving no\n'
                "    end try\nend tell"
            )
        code, _, _ = run_osascript(script)
        return code == 0
    return False


def reload_ppt_in_app(app_name: str, path: Path, slide: int = 1) -> bool:
    """Reload updated .pptx cleanly into frontmost PowerPoint/WPS on macOS."""
    if sys.platform != "darwin":
        return False
    close_in_app(app_name, path)
    ok = open_in_app(app_name, path, visible=True)
    if ok and "Microsoft" in app_name:
        focus_ppt_slide(app_name, slide)
    return ok


def export_ppt_slide_image(app_name: str, pptx_path: Path, out_path: Path, slide: int = 1) -> bool:
    """Export a slide PNG/JPG on macOS via PowerPoint AppleScript (with sips format conversion if JPG)."""
    if sys.platform != "darwin":
        return False
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp = Path(tmp_dir)
        if "PowerPoint" in app_name:
            open_in_app(app_name, pptx_path, visible=True)
            script = (
                f'tell application "{_as_quote(app_name)}"\n'
                f'    save active presentation in POSIX file "{_as_quote(str(tmp))}" as save as PNG\n'
                "end tell"
            )
            code, _, _ = run_osascript(script, timeout=30.0)
            if code == 0:
                pngs = sorted(tmp.rglob("*.png")) + sorted(tmp.rglob("*.PNG"))
                picked = None
                for p in pngs:
                    m = re.search(r"(\d+)\.png$", p.name, re.I)
                    if m and int(m.group(1)) == int(slide):
                        picked = p
                        break
                if picked is None and 1 <= int(slide) <= len(pngs):
                    picked = pngs[int(slide) - 1]
                if picked is not None and picked.is_file():
                    if out_path.suffix.lower() == ".png":
                        shutil.copy2(picked, out_path)
                    else:
                        subprocess.run(
                            ["sips", "-s", "format", "jpeg", str(picked), "--out", str(out_path)],
                            capture_output=True, text=True, check=False,
                        )
                    return out_path.is_file() and out_path.stat().st_size > 0
        # Fallback for WPS / headless macOS using QuickLook
        if shutil.which("qlmanage"):
            subprocess.run(["qlmanage", "-t", "-s", "1920", "-o", str(tmp), str(pptx_path)],
                           capture_output=True, text=True, check=False)
            candidates = list(tmp.glob("*.png"))
            if candidates:
                picked = candidates[0]
                if out_path.suffix.lower() == ".png":
                    shutil.copy2(picked, out_path)
                else:
                    subprocess.run(
                        ["sips", "-s", "format", "jpeg", str(picked), "--out", str(out_path)],
                        capture_output=True, text=True, check=False,
                    )
                return out_path.is_file() and out_path.stat().st_size > 0
    return False


def export_pdf_in_app(app_name: str, src_path: Path, out_path: Path) -> bool:
    """Export an open Office document to PDF on macOS via AppleScript."""
    if sys.platform != "darwin":
        return False
    q_name = _as_quote(app_name)
    q_out = _as_quote(str(out_path.resolve()))
    if "Microsoft" in app_name:
        open_in_app(app_name, src_path, visible=True)
        if "PowerPoint" in app_name:
            script = f'tell application "{q_name}" to save active presentation in POSIX file "{q_out}" as save as PDF'
        elif "Excel" in app_name:
            script = f'tell application "{q_name}" to save workbook as active workbook filename POSIX file "{q_out}" file format PDF file format'
        else:
            script = f'tell application "{q_name}" to save as active document file name POSIX file "{q_out}" file format format PDF'
        code, _, _ = run_osascript(script, timeout=30.0)
        if code == 0 and out_path.is_file() and out_path.stat().st_size > 0:
            return True
    soffice = shutil.which("soffice") or shutil.which("/Applications/LibreOffice.app/Contents/MacOS/soffice")
    if soffice:
        with tempfile.TemporaryDirectory() as tmp:
            subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp, str(src_path)],
                           capture_output=True, text=True, check=False, timeout=30.0)
            produced = Path(tmp) / f"{src_path.stem}.pdf"
            if produced.is_file():
                shutil.copy2(produced, out_path)
                return out_path.is_file() and out_path.stat().st_size > 0
    return False


def list_open_docs(kind: str | None = None) -> list[dict[str, str]]:
    """Query currently open document paths/names on macOS Microsoft Office apps."""
    if sys.platform != "darwin":
        return []
    items: list[dict[str, str]] = []
    targets = [("word", "Microsoft Word", "documents"), ("excel", "Microsoft Excel", "workbooks"), ("ppt", "Microsoft PowerPoint", "presentations")]
    for k, app_name, coll in targets:
        if kind and k != kind:
            continue
        if not is_process_running(app_name):
            continue
        script = (
            f'tell application "{app_name}"\n'
            "    set outList to {}\n"
            f"    repeat with d in {coll}\n"
            "        set p to \"\"\n"
            "        try\n"
            "            set p to POSIX path of (full name of d as alias)\n"
            "        on error\n"
            "            try\n"
            "                set p to (full name of d as string)\n"
            "            end try\n"
            "        end try\n"
            '        set end of outList to (name of d & "|" & p)\n'
            "    end repeat\n"
            '    set AppleScript\'s text item delimiters to linefeed\n'
            "    return outList as string\n"
            "end tell"
        )
        code, out, _ = run_osascript(script)
        if code == 0 and out:
            for line in out.splitlines():
                if "|" not in line:
                    continue
                name, full = line.split("|", 1)
                items.append({"kind": k, "engine": "microsoft", "name": name.strip(), "full_name": full.strip()})
    return items


class MacAppProxy:
    """Lightweight Application proxy for macOS matching Windows COM attribute conventions."""

    def __init__(self, vendor: str, kind: str):
        self.vendor = vendor
        self.kind = kind
        self.app_name = get_app_name(vendor, kind)
        self.Name = self.app_name
        self.Path = f"/Applications/{self.app_name}.app"
        self.Visible = True
        self.ScreenUpdating = True
        self.macos = True

    def Activate(self):
        return activate_app(self.app_name)


# ---------------------------------------------------------------------------
# Minimal valid OOXML templates (Windows/macOS parity, zero extra dependencies)
# ---------------------------------------------------------------------------

_SLIDE_MASTER_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<p:sldMaster xmlns:a="{_A_NS}" xmlns:r="{_R_NS}" xmlns:p="{_P_NS}">'
    '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
    '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '</p:spTree></p:cSld>'
    '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
    '</p:sldMaster>'
)

_SLIDE_LAYOUT_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<p:sldLayout xmlns:a="{_A_NS}" xmlns:r="{_R_NS}" xmlns:p="{_P_NS}" type="blank" preserve="1">'
    '<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
    '</p:spTree></p:cSld></p:sldLayout>'
)

_THEME_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    f'<a:theme xmlns:a="{_A_NS}" name="Office Theme"><a:themeElements>'
    '<a:clrScheme name="Office">'
    '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
    '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
    '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2>'
    '<a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2>'
    '<a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>'
    '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6>'
    '<a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink>'
    '</a:clrScheme>'
    '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>'
    '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
    '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
    '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>'
    '</a:themeElements></a:theme>'
)


def _blank_slide_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        f'<p:sld xmlns:a="{_A_NS}" xmlns:r="{_R_NS}" xmlns:p="{_P_NS}">'
        '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    )


def create_blank_ooxml(path: Path, kind: str):
    """Create a completely valid, standard blank OOXML file (.docx, .xlsx, or .pptx)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        if kind == "word":
            z.writestr("[Content_Types].xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                '</Types>'
            ))
            z.writestr("_rels/.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                '</Relationships>'
            ))
            z.writestr("word/document.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:body><w:p><w:r><w:t></w:t></w:r></w:p></w:body></w:document>'
            ))
        elif kind == "excel":
            z.writestr("[Content_Types].xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '</Types>'
            ))
            z.writestr("_rels/.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                '</Relationships>'
            ))
            z.writestr("xl/workbook.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
            ))
            z.writestr("xl/_rels/workbook.xml.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                '</Relationships>'
            ))
            z.writestr("xl/worksheets/sheet1.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData></sheetData></worksheet>'
            ))
        elif kind == "ppt":
            z.writestr("[Content_Types].xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<Types xmlns="{_CT_NS}">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Default Extension="png" ContentType="image/png"/>'
                '<Default Extension="jpg" ContentType="image/jpeg"/>'
                '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
                '<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
                '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
                '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
                '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
                '</Types>'
            ))
            z.writestr("_rels/.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<Relationships xmlns="{_REL_NS}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
                '</Relationships>'
            ))
            z.writestr("ppt/presentation.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<p:presentation xmlns:a="{_A_NS}" xmlns:r="{_R_NS}" xmlns:p="{_P_NS}">'
                '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
                '<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>'
                '<p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>'
                '<p:notesSz cx="6858000" cy="9144000"/>'
                '</p:presentation>'
            ))
            z.writestr("ppt/_rels/presentation.xml.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<Relationships xmlns="{_REL_NS}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>'
                '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
                '</Relationships>'
            ))
            z.writestr("ppt/slides/slide1.xml", _blank_slide_xml())
            z.writestr("ppt/slides/_rels/slide1.xml.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<Relationships xmlns="{_REL_NS}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
                '</Relationships>'
            ))
            z.writestr("ppt/slideLayouts/slideLayout1.xml", _SLIDE_LAYOUT_XML)
            z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<Relationships xmlns="{_REL_NS}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
                '</Relationships>'
            ))
            z.writestr("ppt/slideMasters/slideMaster1.xml", _SLIDE_MASTER_XML)
            z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<Relationships xmlns="{_REL_NS}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
                '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
                '</Relationships>'
            ))
            z.writestr("ppt/theme/theme1.xml", _THEME_XML)


# ---------------------------------------------------------------------------
# Cross-platform PresentationML / DrawingML proxy (powers ppt_draw_scene on macOS)
# ---------------------------------------------------------------------------

def _rgb_int_to_hex(val: int) -> str:
    return f"{val & 0xFF:02X}{(val >> 8) & 0xFF:02X}{(val >> 16) & 0xFF:02X}"


def _image_pixel_size(path: Path) -> tuple[float, float]:
    data = path.read_bytes()
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        w, h = struct.unpack(">II", data[16:24])
        return float(w) * 0.75, float(h) * 0.75
    if data[:2] == b"\xff\xd8":
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                h, w = struct.unpack(">HH", data[i + 5:i + 9])
                return float(w) * 0.75, float(h) * 0.75
            if marker in (0xD8, 0xD9):
                i += 2
            else:
                seg_len = struct.unpack(">H", data[i + 2:i + 4])[0]
                i += 2 + max(2, seg_len)
    return 100.0, 100.0


class _OoxmlColor:
    def __init__(self, owner):
        self._owner = owner
        self._rgb = 0

    @property
    def RGB(self) -> int:
        return self._rgb

    @RGB.setter
    def RGB(self, value: int):
        self._rgb = int(value)
        self._owner._dirty_hex = _rgb_int_to_hex(self._rgb)


class _OoxmlFill:
    def __init__(self):
        self.Visible = 0
        self.ForeColor = _OoxmlColor(self)
        self._dirty_hex = "FFFFFF"
        self.Transparency = 0.0

    def Solid(self):
        self.Visible = -1


class _OoxmlLine:
    def __init__(self):
        self.Visible = 0
        self.ForeColor = _OoxmlColor(self)
        self._dirty_hex = "262626"
        self.Weight = 1.0
        self.DashStyle = 1
        self.BeginArrowheadStyle = 1
        self.EndArrowheadStyle = 1
        self.EndArrowheadLength = 2
        self.EndArrowheadWidth = 2


class _OoxmlShadow:
    def __init__(self):
        self.Visible = 0
        self.Blur = 4.0
        self.Transparency = 0.78
        self.OffsetX = 1.5
        self.OffsetY = 2.0


class _OoxmlFont:
    def __init__(self):
        self.Name = "Times New Roman"
        self.Size = 10.0
        self.Bold = 0
        self.Italic = 0
        self.Subscript = 0
        self.Superscript = 0
        self.Color = _OoxmlColor(self)
        self._dirty_hex = "111111"


class _OoxmlParagraphFormat:
    def __init__(self):
        self.Alignment = 2
        self.SpaceWithin = None


class _OoxmlTextRange:
    def __init__(self):
        self.Text = ""
        self.Font = _OoxmlFont()
        self.ParagraphFormat = _OoxmlParagraphFormat()
        self.runs: list[tuple[int, int, _OoxmlFont]] = []

    def Characters(self, pos: int, length: int):
        sub = _OoxmlTextRange()
        sub.Text = self.Text[pos - 1:pos - 1 + length]
        self.runs.append((pos - 1, pos - 1 + length, sub.Font))
        return sub


class _OoxmlTextFrame:
    def __init__(self):
        self.MarginLeft = 1.0
        self.MarginRight = 1.0
        self.MarginTop = 0.5
        self.MarginBottom = 0.5
        self.WordWrap = -1
        self.VerticalAnchor = 3
        self.TextRange = _OoxmlTextRange()


class _OoxmlAdjustments:
    def __init__(self):
        self.values: dict[int, float] = {}

    def Item(self, idx: int, val: float):
        self.values[idx] = float(val)

    def __setitem__(self, idx: int, val: float):
        self.values[idx] = float(val)


class _OoxmlPictureFormat:
    def __init__(self):
        self.CropLeft = 0.0
        self.CropTop = 0.0
        self.CropRight = 0.0
        self.CropBottom = 0.0


class _OoxmlShape:
    def __init__(self, shape_id: int, kind: str, **meta):
        self.Id = shape_id
        self.Name = f"Shape {shape_id}"
        self.kind = kind
        self.meta = meta
        self.Left = float(meta.get("left", 0.0))
        self.Top = float(meta.get("top", 0.0))
        self.Width = float(meta.get("w", 1.0))
        self.Height = float(meta.get("h", 1.0))
        self.Rotation = 0.0
        self.LockAspectRatio = -1
        self.Fill = _OoxmlFill()
        self.Line = _OoxmlLine()
        self.Shadow = _OoxmlShadow()
        self.TextFrame = _OoxmlTextFrame()
        self.HasTextFrame = -1 if kind in ("shape", "textbox", "freeform") else 0
        self.Adjustments = _OoxmlAdjustments()
        self.PictureFormat = _OoxmlPictureFormat()

    def to_xml(self, rel_map: list[tuple[str, str, bytes]]) -> ET.Element:
        if self.kind == "connector":
            return self._connector_xml()
        if self.kind == "picture":
            return self._picture_xml(rel_map)
        return self._sp_xml()

    def _ln_el(self) -> ET.Element | None:
        if self.Line.Visible == 0 and self.kind != "connector":
            return None
        ln = ET.Element(f"{{{_A_NS}}}ln", w=str(max(1270, int(round(self.Line.Weight * _EMU)))))
        fill = ET.SubElement(ln, f"{{{_A_NS}}}solidFill")
        ET.SubElement(fill, f"{{{_A_NS}}}srgbClr", val=self.Line._dirty_hex)
        dash_map = {2: "sysDot", 3: "dot", 4: "dash", 5: "dashDot", 7: "lgDash"}
        if self.Line.DashStyle in dash_map:
            ET.SubElement(ln, f"{{{_A_NS}}}prstDash", val=dash_map[self.Line.DashStyle])
        arr_map = {2: "triangle", 3: "arrow", 4: "stealth", 5: "diamond", 6: "oval"}
        len_map = {1: "sm", 2: "med", 3: "lg"}
        if self.Line.BeginArrowheadStyle in arr_map:
            ET.SubElement(ln, f"{{{_A_NS}}}headEnd", type=arr_map[self.Line.BeginArrowheadStyle], w="med", len="med")
        if self.Line.EndArrowheadStyle in arr_map:
            ET.SubElement(
                ln, f"{{{_A_NS}}}tailEnd", type=arr_map[self.Line.EndArrowheadStyle],
                w=len_map.get(self.Line.EndArrowheadWidth, "med"),
                len=len_map.get(self.Line.EndArrowheadLength, "med"),
            )
        return ln

    def _fill_el(self) -> ET.Element:
        if self.Fill.Visible == 0:
            return ET.Element(f"{{{_A_NS}}}noFill")
        sf = ET.Element(f"{{{_A_NS}}}solidFill")
        clr = ET.SubElement(sf, f"{{{_A_NS}}}srgbClr", val=self.Fill._dirty_hex)
        if self.Fill.Transparency:
            alpha = max(0, min(100000, int(round((1.0 - float(self.Fill.Transparency)) * 100000))))
            ET.SubElement(clr, f"{{{_A_NS}}}alpha", val=str(alpha))
        return sf

    def _effect_el(self) -> ET.Element | None:
        if not self.Shadow.Visible:
            return None
        eff = ET.Element(f"{{{_A_NS}}}effectLst")
        dist = int(round(math.hypot(self.Shadow.OffsetX, self.Shadow.OffsetY) * _EMU)) if "math" in globals() else int(round(self.Shadow.OffsetY * _EMU))
        blur = int(round(self.Shadow.Blur * _EMU))
        shdw = ET.SubElement(eff, f"{{{_A_NS}}}outerShdw", blurRad=str(blur), dist=str(dist), dir="2700000", algn="tl")
        clr = ET.SubElement(shdw, f"{{{_A_NS}}}srgbClr", val="000000")
        alpha = max(5000, min(95000, int(round((1.0 - float(self.Shadow.Transparency)) * 100000))))
        ET.SubElement(clr, f"{{{_A_NS}}}alpha", val=str(alpha))
        return eff

    def _txbody_el(self) -> ET.Element:
        tf = self.TextFrame
        tr = tf.TextRange
        tb = ET.Element(f"{{{_P_NS}}}txBody")
        anchor = {1: "t", 3: "ctr", 4: "b"}.get(tf.VerticalAnchor, "ctr")
        ET.SubElement(
            tb, f"{{{_A_NS}}}bodyPr",
            wrap="square" if tf.WordWrap else "none",
            lIns=str(int(round(tf.MarginLeft * _EMU))),
            tIns=str(int(round(tf.MarginTop * _EMU))),
            rIns=str(int(round(tf.MarginRight * _EMU))),
            bIns=str(int(round(tf.MarginBottom * _EMU))),
            anchor=anchor,
        )
        ET.SubElement(tb, f"{{{_A_NS}}}lstStyle")
        algn = {1: "l", 2: "ctr", 3: "r", 4: "just"}.get(tr.ParagraphFormat.Alignment, "ctr")
        text = tr.Text or ""
        lines = text.split("\n")
        cursor = 0
        for line in lines:
            p = ET.SubElement(tb, f"{{{_A_NS}}}p")
            pPr = ET.SubElement(p, f"{{{_A_NS}}}pPr", algn=algn)
            if tr.ParagraphFormat.SpaceWithin:
                lnSpc = ET.SubElement(pPr, f"{{{_A_NS}}}lnSpc")
                ET.SubElement(lnSpc, f"{{{_A_NS}}}spcPct", val=str(int(round(float(tr.ParagraphFormat.SpaceWithin) * 100000))))
            line_start, line_end = cursor, cursor + len(line)
            if tr.runs:
                matched = [(max(s, line_start), min(e, line_end), f) for s, e, f in tr.runs if e > line_start and s < line_end]
                if matched:
                    for s, e, f in matched:
                        self._append_run(p, text[s:e], f)
                else:
                    self._append_run(p, line, tr.Font)
            else:
                self._append_run(p, line, tr.Font)
            cursor = line_end + 1
        return tb

    @staticmethod
    def _append_run(p: ET.Element, txt: str, font: _OoxmlFont):
        r = ET.SubElement(p, f"{{{_A_NS}}}r")
        rPr = ET.SubElement(
            r, f"{{{_A_NS}}}rPr", lang="en-US",
            sz=str(max(400, int(round(float(font.Size) * 100)))),
            b="1" if font.Bold else "0",
            i="1" if font.Italic else "0",
            dirty="0",
        )
        if font.Subscript:
            rPr.set("baseline", "-25000")
        elif font.Superscript:
            rPr.set("baseline", "30000")
        fill = ET.SubElement(rPr, f"{{{_A_NS}}}solidFill")
        ET.SubElement(fill, f"{{{_A_NS}}}srgbClr", val=font._dirty_hex)
        ET.SubElement(rPr, f"{{{_A_NS}}}latin", typeface=str(font.Name))
        ET.SubElement(rPr, f"{{{_A_NS}}}ea", typeface=str(font.Name))
        t = ET.SubElement(r, f"{{{_A_NS}}}t")
        t.text = txt

    def _sp_xml(self) -> ET.Element:
        sp = ET.Element(f"{{{_P_NS}}}sp")
        nv = ET.SubElement(sp, f"{{{_P_NS}}}nvSpPr")
        ET.SubElement(nv, f"{{{_P_NS}}}cNvPr", id=str(self.Id), name=self.Name)
        ET.SubElement(nv, f"{{{_P_NS}}}cNvSpPr", txBox="1" if self.kind == "textbox" else "0")
        ET.SubElement(nv, f"{{{_P_NS}}}nvPr")
        spPr = ET.SubElement(sp, f"{{{_P_NS}}}spPr")

        if self.kind == "freeform":
            self._build_freeform_geom(spPr)
        else:
            xfrm = ET.SubElement(spPr, f"{{{_A_NS}}}xfrm")
            if self.Rotation:
                xfrm.set("rot", str(int(round(self.Rotation * 60000))))
            ET.SubElement(xfrm, f"{{{_A_NS}}}off", x=str(int(round(self.Left * _EMU))), y=str(int(round(self.Top * _EMU))))
            ET.SubElement(xfrm, f"{{{_A_NS}}}ext", cx=str(max(1, int(round(self.Width * _EMU)))), cy=str(max(1, int(round(self.Height * _EMU)))))
            prst = {1: "rect", 5: "roundRect", 32: "rightBrace"}.get(int(self.meta.get("stype", 1)), "rect")
            pg = ET.SubElement(spPr, f"{{{_A_NS}}}prstGeom", prst=prst)
            av = ET.SubElement(pg, f"{{{_A_NS}}}avLst")
            if 1 in self.Adjustments.values and prst == "roundRect":
                ET.SubElement(av, f"{{{_A_NS}}}gd", name="adj", fmla=f"val {int(round(self.Adjustments.values[1] * 100000))}")

        spPr.append(self._fill_el())
        ln = self._ln_el()
        if ln is not None:
            spPr.append(ln)
        elif self.kind == "textbox":
            ln_no = ET.SubElement(spPr, f"{{{_A_NS}}}ln")
            ET.SubElement(ln_no, f"{{{_A_NS}}}noFill")
        eff = self._effect_el()
        if eff is not None:
            spPr.append(eff)
        if self.kind != "freeform" or self.TextFrame.TextRange.Text:
            sp.append(self._txbody_el())
        return sp

    def _build_freeform_geom(self, spPr: ET.Element):
        start = self.meta["start"]
        nodes = self.meta["nodes"]
        all_pts = [start]
        for n in nodes:
            seg_type = n[0]
            if seg_type == 0:
                all_pts.append((n[2], n[3]))
            else:
                all_pts.extend([(n[2], n[3]), (n[4], n[5]), (n[6], n[7])])
        xs = [p[0] for p in all_pts]
        ys = [p[1] for p in all_pts]
        x0, y0 = min(xs), min(ys)
        w = max(0.5, max(xs) - x0)
        h = max(0.5, max(ys) - y0)
        w_emu = max(6350, int(round(w * _EMU)))
        h_emu = max(6350, int(round(h * _EMU)))
        xfrm = ET.SubElement(spPr, f"{{{_A_NS}}}xfrm")
        ET.SubElement(xfrm, f"{{{_A_NS}}}off", x=str(int(round(x0 * _EMU))), y=str(int(round(y0 * _EMU))))
        ET.SubElement(xfrm, f"{{{_A_NS}}}ext", cx=str(w_emu), cy=str(h_emu))
        cg = ET.SubElement(spPr, f"{{{_A_NS}}}custGeom")
        ET.SubElement(cg, f"{{{_A_NS}}}avLst")
        ET.SubElement(cg, f"{{{_A_NS}}}gdLst")
        ET.SubElement(cg, f"{{{_A_NS}}}ahLst")
        ET.SubElement(cg, f"{{{_A_NS}}}cxnLst")
        ET.SubElement(cg, f"{{{_A_NS}}}rect", l="0", t="0", r="r", b="b")
        pl = ET.SubElement(cg, f"{{{_A_NS}}}pathLst")
        path_el = ET.SubElement(pl, f"{{{_A_NS}}}path", w=str(w_emu), h=str(h_emu))

        def px(pt):
            return str(int(round((pt[0] - x0) * _EMU))), str(int(round((pt[1] - y0) * _EMU)))

        mv = ET.SubElement(path_el, f"{{{_A_NS}}}moveTo")
        sx, sy = px(start)
        ET.SubElement(mv, f"{{{_A_NS}}}pt", x=sx, y=sy)
        last_pt = start
        has_bezier = False
        for n in nodes:
            if n[0] == 0:
                lt = ET.SubElement(path_el, f"{{{_A_NS}}}lnTo")
                x_s, y_s = px((n[2], n[3]))
                ET.SubElement(lt, f"{{{_A_NS}}}pt", x=x_s, y=y_s)
                last_pt = (n[2], n[3])
            else:
                has_bezier = True
                bz = ET.SubElement(path_el, f"{{{_A_NS}}}cubicBezTo")
                for pt in ((n[2], n[3]), (n[4], n[5]), (n[6], n[7])):
                    x_s, y_s = px(pt)
                    ET.SubElement(bz, f"{{{_A_NS}}}pt", x=x_s, y=y_s)
                last_pt = (n[6], n[7])
        if (not has_bezier and len(nodes) >= 3
                and abs(last_pt[0] - start[0]) < 1e-2 and abs(last_pt[1] - start[1]) < 1e-2):
            ET.SubElement(path_el, f"{{{_A_NS}}}close")

    def _connector_xml(self) -> ET.Element:
        x1, y1 = self.meta["x1"], self.meta["y1"]
        x2, y2 = self.meta["x2"], self.meta["y2"]
        cxn = ET.Element(f"{{{_P_NS}}}cxnSp")
        nv = ET.SubElement(cxn, f"{{{_P_NS}}}nvCxnSpPr")
        ET.SubElement(nv, f"{{{_P_NS}}}cNvPr", id=str(self.Id), name=self.Name)
        ET.SubElement(nv, f"{{{_P_NS}}}cNvCxnSpPr")
        ET.SubElement(nv, f"{{{_P_NS}}}nvPr")
        spPr = ET.SubElement(cxn, f"{{{_P_NS}}}spPr")
        xfrm = ET.SubElement(spPr, f"{{{_A_NS}}}xfrm")
        if x2 < x1:
            xfrm.set("flipH", "1")
        if y2 < y1:
            xfrm.set("flipV", "1")
        ET.SubElement(xfrm, f"{{{_A_NS}}}off", x=str(int(round(min(x1, x2) * _EMU))), y=str(int(round(min(y1, y2) * _EMU))))
        ET.SubElement(xfrm, f"{{{_A_NS}}}ext", cx=str(int(round(abs(x2 - x1) * _EMU))), cy=str(int(round(abs(y2 - y1) * _EMU))))
        pg = ET.SubElement(spPr, f"{{{_A_NS}}}prstGeom", prst="straightConnector1")
        ET.SubElement(pg, f"{{{_A_NS}}}avLst")
        ln = self._ln_el()
        if ln is not None:
            spPr.append(ln)
        return cxn

    def _picture_xml(self, rel_map: list[tuple[str, str, bytes]]) -> ET.Element:
        img_path = Path(self.meta["path"])
        rid = f"rImg{len(rel_map) + 10}"
        ext = img_path.suffix.lower().lstrip(".") or "png"
        media_name = f"media_{self.Id}_{uuid.uuid4().hex[:6]}.{ext}"
        rel_map.append((rid, media_name, img_path.read_bytes()))
        orig_w, orig_h = float(self.meta.get("orig_w", 100.0)), float(self.meta.get("orig_h", 100.0))

        pic = ET.Element(f"{{{_P_NS}}}pic")
        nv = ET.SubElement(pic, f"{{{_P_NS}}}nvPicPr")
        ET.SubElement(nv, f"{{{_P_NS}}}cNvPr", id=str(self.Id), name=self.Name)
        cNvPic = ET.SubElement(nv, f"{{{_P_NS}}}cNvPicPr")
        ET.SubElement(cNvPic, f"{{{_A_NS}}}picLocks", noChangeAspect="0")
        ET.SubElement(nv, f"{{{_P_NS}}}nvPr")
        blipFill = ET.SubElement(pic, f"{{{_P_NS}}}blipFill")
        ET.SubElement(blipFill, f"{{{_A_NS}}}blip", **{f"{{{_R_NS}}}embed": rid})
        cl = int(round((self.PictureFormat.CropLeft / max(1.0, orig_w)) * 100000))
        ct = int(round((self.PictureFormat.CropTop / max(1.0, orig_h)) * 100000))
        cr = int(round((self.PictureFormat.CropRight / max(1.0, orig_w)) * 100000))
        cb = int(round((self.PictureFormat.CropBottom / max(1.0, orig_h)) * 100000))
        if cl or ct or cr or cb:
            ET.SubElement(blipFill, f"{{{_A_NS}}}srcRect", l=str(cl), t=str(ct), r=str(cr), b=str(cb))
        stretch = ET.SubElement(blipFill, f"{{{_A_NS}}}stretch")
        ET.SubElement(stretch, f"{{{_A_NS}}}fillRect")

        spPr = ET.SubElement(pic, f"{{{_P_NS}}}spPr")
        xfrm = ET.SubElement(spPr, f"{{{_A_NS}}}xfrm")
        ET.SubElement(xfrm, f"{{{_A_NS}}}off", x=str(int(round(self.Left * _EMU))), y=str(int(round(self.Top * _EMU))))
        ET.SubElement(xfrm, f"{{{_A_NS}}}ext", cx=str(max(1, int(round(self.Width * _EMU)))), cy=str(max(1, int(round(self.Height * _EMU)))))
        pg = ET.SubElement(spPr, f"{{{_A_NS}}}prstGeom", prst="rect")
        ET.SubElement(pg, f"{{{_A_NS}}}avLst")
        return pic


class _OoxmlFreeformBuilder:
    def __init__(self, shapes, x0: float, y0: float):
        self._shapes = shapes
        self._start = (float(x0), float(y0))
        self._nodes: list[tuple] = []

    def AddNodes(self, seg_type: int, edit_type: int, *coords):
        self._nodes.append((int(seg_type), int(edit_type), *[float(c) for c in coords]))

    def ConvertToShape(self) -> _OoxmlShape:
        s = self._shapes._new_shape("freeform", start=self._start, nodes=list(self._nodes))
        return s


class _OoxmlShapes:
    def __init__(self):
        self.items: list[_OoxmlShape] = []
        self._next_id = 2

    @property
    def Count(self) -> int:
        return len(self.items)

    def Item(self, idx: int):
        return self.items[int(idx) - 1]

    def _new_shape(self, kind: str, **meta) -> _OoxmlShape:
        shape = _OoxmlShape(self._next_id, kind, **meta)
        self._next_id += 1
        self.items.append(shape)
        return shape

    def AddShape(self, stype: int, left: float, top: float, w: float, h: float):
        return self._new_shape("shape", stype=int(stype), left=left, top=top, w=w, h=h)

    def AddTextbox(self, orient: int, left: float, top: float, w: float, h: float):
        return self._new_shape("textbox", orient=int(orient), left=left, top=top, w=w, h=h)

    def AddConnector(self, ctype: int, x1: float, y1: float, x2: float, y2: float):
        return self._new_shape("connector", ctype=int(ctype), x1=float(x1), y1=float(y1), x2=float(x2), y2=float(y2),
                               left=min(x1, x2), top=min(y1, y2), w=abs(x2 - x1), h=abs(y2 - y1))

    def AddPicture(self, path: str, link: bool, save: bool, left: float, top: float, w: float, h: float):
        orig_w, orig_h = _image_pixel_size(Path(path))
        rw = orig_w if float(w) < 0 else float(w)
        rh = orig_h if float(h) < 0 else float(h)
        return self._new_shape("picture", path=path, left=left, top=top, w=rw, h=rh, orig_w=orig_w, orig_h=orig_h)

    def BuildFreeform(self, editing_type: int, x: float, y: float):
        return _OoxmlFreeformBuilder(self, x, y)


class _OoxmlSlideProxy:
    def __init__(self, slide_idx: int):
        self.SlideIndex = slide_idx
        self.SlideID = 255 + slide_idx
        self.Shapes = _OoxmlShapes()


class _OoxmlSlidesCollection:
    def __init__(self, doc, count: int = 1):
        self._doc = doc
        self._slides: list[_OoxmlSlideProxy] = [_OoxmlSlideProxy(i) for i in range(1, max(1, count) + 1)]

    @property
    def Count(self) -> int:
        return len(self._slides)

    def Item(self, idx: int) -> _OoxmlSlideProxy:
        return self._slides[int(idx) - 1]

    def Add(self, idx: int, layout: int = 12) -> _OoxmlSlideProxy:
        s = _OoxmlSlideProxy(len(self._slides) + 1)
        self._slides.insert(max(0, min(len(self._slides), int(idx) - 1)), s)
        for i, item in enumerate(self._slides, 1):
            item.SlideIndex = i
            item.SlideID = 255 + i
        self._doc.Saved = False
        return s


class MacDocProxy:
    """Native-compatible document proxy backed by macOS AppleScript & OOXML."""

    def __init__(self, path: str | Path, kind: str, engine: str):
        p = Path(path).resolve()
        self._path = p
        self.FullName = str(p)
        self.Name = p.name
        self.kind = kind
        self.engine = engine
        self.Saved = True
        self.ReadOnly = False
        self.Windows = type("MacWindows", (), {
            "Count": 1,
            "Item": lambda _self, i: type("MacWin", (), {
                "Visible": True, "ViewType": 9,
                "Activate": lambda _w: activate_app(get_app_name(engine, kind)),
                "View": type("MacView", (), {"GotoSlide": lambda _v, idx: focus_ppt_slide(get_app_name(engine, kind), idx), "ZoomToFit": True})(),
            })(),
        })()
        slide_w, slide_h, slide_count = 960.0, 540.0, 1
        self._base_slide_xmls: dict[int, bytes] = {}
        self._base_slide_rels: dict[int, bytes] = {}
        if kind == "ppt" and p.is_file():
            try:
                with zipfile.ZipFile(p, "r") as zin:
                    names = zin.namelist()
                    if "ppt/presentation.xml" in names:
                        pres_root = ET.fromstring(zin.read("ppt/presentation.xml"))
                        sld_sz = pres_root.find(f"{{{_P_NS}}}sldSz")
                        if sld_sz is not None:
                            slide_w = round(int(sld_sz.get("cx", str(int(960 * _EMU)))) / _EMU, 2)
                            slide_h = round(int(sld_sz.get("cy", str(int(540 * _EMU)))) / _EMU, 2)
                        sld_ids = pres_root.findall(f".//{{{_P_NS}}}sldId")
                        if sld_ids:
                            slide_count = len(sld_ids)
                    for idx in range(1, slide_count + 1):
                        sname = f"ppt/slides/slide{idx}.xml"
                        rname = f"ppt/slides/_rels/slide{idx}.xml.rels"
                        if sname in names:
                            self._base_slide_xmls[idx] = zin.read(sname)
                        if rname in names:
                            self._base_slide_rels[idx] = zin.read(rname)
            except Exception:
                pass
        self.PageSetup = type("PageSetup", (), {"SlideWidth": slide_w, "SlideHeight": slide_h})()
        self.Slides = _OoxmlSlidesCollection(self, count=slide_count)

    def _flush_pptx(self, target: Path):
        if self.kind != "ppt":
            return
        with zipfile.ZipFile(self._path, "r") as zin:
            existing_files = {name: zin.read(name) for name in zin.namelist()}

        ET.register_namespace("p", _P_NS)
        ET.register_namespace("a", _A_NS)
        ET.register_namespace("r", _R_NS)
        ET.register_namespace("", _CT_NS)
        pres_xml = ET.fromstring(existing_files["ppt/presentation.xml"])
        sld_sz = pres_xml.find(f"{{{_P_NS}}}sldSz")
        if sld_sz is not None:
            sld_sz.set("cx", str(int(round(float(self.PageSetup.SlideWidth) * _EMU))))
            sld_sz.set("cy", str(int(round(float(self.PageSetup.SlideHeight) * _EMU))))

        # Ensure [Content_Types].xml knows image extensions and all slides
        ct_xml = ET.fromstring(existing_files["[Content_Types].xml"])
        seen_ext = {el.get("Extension", "").lower() for el in ct_xml.findall(f"{{{_CT_NS}}}Default")}
        for ext, ctype in (("png", "image/png"), ("jpg", "image/jpeg"), ("jpeg", "image/jpeg")):
            if ext not in seen_ext:
                ET.SubElement(ct_xml, f"{{{_CT_NS}}}Default", Extension=ext, ContentType=ctype)
        seen_parts = {el.get("PartName", "") for el in ct_xml.findall(f"{{{_CT_NS}}}Override")}

        pres_rels = ET.fromstring(existing_files["ppt/_rels/presentation.xml.rels"])
        seen_targets = {el.get("Target", "") for el in pres_rels.findall(f"{{{_REL_NS}}}Relationship")}
        used_rids = {el.get("Id", "") for el in pres_rels.findall(f"{{{_REL_NS}}}Relationship")}
        sld_id_lst = pres_xml.find(f"{{{_P_NS}}}sldIdLst")
        if sld_id_lst is None:
            sld_id_lst = ET.SubElement(pres_xml, f"{{{_P_NS}}}sldIdLst")
        seen_sld_rids = {el.get(f"{{{_R_NS}}}id", "") for el in sld_id_lst.findall(f"{{{_P_NS}}}sldId")}

        for slide in self.Slides._slides:
            slide_name = f"ppt/slides/slide{slide.SlideIndex}.xml"
            rels_name = f"ppt/slides/_rels/slide{slide.SlideIndex}.xml.rels"
            rel_target = f"slides/slide{slide.SlideIndex}.xml"
            part_name = f"/{slide_name}"
            if part_name not in seen_parts:
                ET.SubElement(
                    ct_xml, f"{{{_CT_NS}}}Override",
                    PartName=part_name,
                    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
                )
                seen_parts.add(part_name)
            if rel_target not in seen_targets:
                rid_num = len(used_rids) + 1
                while f"rId{rid_num}" in used_rids:
                    rid_num += 1
                new_rid = f"rId{rid_num}"
                used_rids.add(new_rid)
                ET.SubElement(
                    pres_rels, f"{{{_REL_NS}}}Relationship",
                    Id=new_rid,
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                    Target=rel_target,
                )
                seen_targets.add(rel_target)
                if new_rid not in seen_sld_rids:
                    ET.SubElement(
                        sld_id_lst, f"{{{_P_NS}}}sldId",
                        id=str(255 + slide.SlideIndex),
                        **{f"{{{_R_NS}}}id": new_rid},
                    )
                    seen_sld_rids.add(new_rid)

            if not slide.Shapes.items and slide_name in existing_files:
                continue
            base_xml = self._base_slide_xmls.get(slide.SlideIndex) or _blank_slide_xml().encode("utf-8")
            sld_root = ET.fromstring(base_xml)
            sp_tree = sld_root.find(f".//{{{_P_NS}}}spTree")
            rel_media: list[tuple[str, str, bytes]] = []
            for shape in slide.Shapes.items:
                sp_tree.append(shape.to_xml(rel_media))
            existing_files[slide_name] = ET.tostring(sld_root, encoding="utf-8", xml_declaration=True)

            default_rels = (
                f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<Relationships xmlns="{_REL_NS}">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
                '</Relationships>'
            ).encode("utf-8")
            rels_bytes = self._base_slide_rels.get(slide.SlideIndex) or default_rels
            if rel_media or rels_name not in existing_files:
                rels_root = ET.fromstring(rels_bytes)
                for rid, mname, mbytes in rel_media:
                    ET.SubElement(
                        rels_root, f"{{{_REL_NS}}}Relationship",
                        Id=rid,
                        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        Target=f"../media/{mname}",
                    )
                    existing_files[f"ppt/media/{mname}"] = mbytes
                existing_files[rels_name] = ET.tostring(rels_root, encoding="utf-8", xml_declaration=True)

        existing_files["ppt/presentation.xml"] = ET.tostring(pres_xml, encoding="utf-8", xml_declaration=True)
        existing_files["ppt/_rels/presentation.xml.rels"] = ET.tostring(pres_rels, encoding="utf-8", xml_declaration=True)
        existing_files["[Content_Types].xml"] = ET.tostring(ct_xml, encoding="utf-8", xml_declaration=True)

        tmp = target.with_name(f".{target.name}.{uuid.uuid4().hex[:6]}.tmp")
        try:
            with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
                for fname, payload in existing_files.items():
                    zout.writestr(fname, payload)
            os.replace(tmp, target)
        finally:
            if tmp.exists():
                tmp.unlink()

    def Save(self):
        if self.kind == "ppt" and any(s.Shapes.items for s in self.Slides._slides):
            self._flush_pptx(self._path)
            if sys.platform == "darwin":
                reload_ppt_in_app(get_app_name(self.engine, self.kind), self._path, slide=1)
        elif sys.platform == "darwin":
            save_in_app(get_app_name(self.engine, self.kind), self._path)
        self.Saved = True

    def SaveAs(self, target_path: str, _fmt: int = 0):
        dest = Path(target_path).resolve()
        if self.kind == "ppt" and any(s.Shapes.items for s in self.Slides._slides):
            self._flush_pptx(dest)
        elif dest != self._path:
            shutil.copy2(self._path, dest)
        self._path = dest
        self.FullName = str(dest)
        self.Name = dest.name
        self.Saved = True

    def Close(self, _save_changes: bool = False):
        if sys.platform == "darwin":
            close_in_app(get_app_name(self.engine, self.kind), self._path)
