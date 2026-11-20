"""Show real staged Word/Excel/PPT construction through stdio MCP, using test documents only."""
from __future__ import annotations

import argparse
import asyncio
import datetime as dt
import hashlib
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

from mcp import Client, StdioServerParameters

ROOT = Path(__file__).resolve().parents[1]
EXAMPLES = {"word": (".docx", "live_word_document.json"),
            "excel": (".xlsx", "live_excel_workbook.json"),
            "ppt": (".pptx", "live_ppt_flowchart.json")}


def unpack(result):
    if result.is_error:
        raise RuntimeError(str(result.content))
    data = result.structured_content or json.loads(result.content[0].text)
    if data.get("state") in ("partial_error", "display_error", "pending"):
        raise RuntimeError(json.dumps(data, ensure_ascii=False))
    return data


def snapshot(path, kind):
    payload = path.read_bytes()
    result = {"sha256": hashlib.sha256(payload).hexdigest()}
    with zipfile.ZipFile(path) as archive:
        if kind == "ppt":
            root = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
            p = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
            result["shapes"] = len(root.findall(".//" + p + "sp")) + len(root.findall(".//" + p + "cxnSp"))
        elif kind == "word":
            root = ET.fromstring(archive.read("word/document.xml"))
            result["paragraphs"] = len(root.findall(".//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"))
        else:
            root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
            result["cells"] = len(root.findall(".//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c"))
    return result


async def run(engine, kind, folder, keep_open=False, installed=False, mode="legacy"):
    suffix, example = EXAMPLES[kind]
    path = folder / ("live-" + engine + "-" + kind + suffix)
    server = ["-m", "microsoft_wps_live_mcp"] if installed else [str(ROOT / "scripts/run_mcp.py")]
    parameters = StdioServerParameters(command=sys.executable, args=["-X", "utf8", "-u", *server,
                                       "--data-dir", str(folder / "runtime")])
    events = []
    async with Client(parameters, read_timeout_seconds=120, mode=mode) as client:
        listed = await client.list_tools()
        assert len(listed.tools) == 35
        info = unpack(await client.call_tool("office_create_document", {"path": str(path), "engine": engine}))
        document = info["document_id"]
        success = False
        try:
            assert info["live_view"]["enabled"] and info["live_view"]["window_open"], info
            stages = json.loads((ROOT / "examples" / example).read_text(encoding="utf-8"))

            async def progress(number, total, message):
                record = {"stage": int(number), "total": int(total), "label": message, **snapshot(path, kind)}
                events.append(record)
                (folder / (engine + '-' + kind + '-progress.json')).write_text(
                    json.dumps(events, ensure_ascii=False, indent=2), encoding='utf-8')
                print(json.dumps({"engine": engine, "kind": kind, **record}), flush=True)

            result = unpack(await client.call_tool("office_live_steps", {"document": document, "stages": stages,
                                                                         "pause_ms": 650}, progress_callback=progress))
            assert result["completed_stages"] == len(stages), result
            assert len(events) == len(stages), events
            assert len({e["sha256"] for e in events}) >= 3, events
            assert all(s["disk_saved"] for s in result["stages"]), result
            assert not any(s["live_view"].get("focus_error") for s in result["stages"]), result
            if kind == "word":
                edited = unpack(await client.call_tool("office_word_insert_text", {"document": document,
                    "text": "Supplement A\nSupplement B 🧪\n"}))
                assert edited["operation"]["visible_steps"] == 2 and edited["disk_saved"], edited
                read = unpack(await client.call_tool("office_word_read", {"document": document}))
                assert "Supplement B 🧪" in "".join(p["text"] for p in read["paragraphs"])
            elif kind == "excel":
                values = [["00123", "=literal"], ["2026-01-02", "text"], [3, 4]]
                edited = unpack(await client.call_tool("office_excel_write_range", {"document": document,
                    "sheet": 1, "address": "E10:F12", "values": values}))
                assert edited["disk_saved"] and edited["live_view"]["focus"]["address"] == "E12:F12", edited
                read = unpack(await client.call_tool("office_excel_read_range", {"document": document,
                    "sheet": 1, "address": "E10:F12"}))
                assert read["values"] == values, read
            shown = unpack(await client.call_tool("office_show_document", {"document": document,
                "focus": {"slide": 1} if kind == "ppt" else {"sheet": 1, "address": "A1:D8"} if kind == "excel" else {"start": 0, "end": 13}}))
            if kind == "ppt":
                assert [e["shapes"] for e in events] == [1, 2, 3, 4, 5, 6], events
            copied = unpack(await client.call_tool("office_create_copy", {"document": document}))
            assert copied["copy_sha256"] == copied["sha256"]
            success = True
            return {"engine": engine, "kind": kind, "passed": True, "path": str(path),
                    "protocol": client.protocol_version, "live_view": shown["live_view"],
                    "stages": events, "copy_path": copied["copy_path"]}
        finally:
            if not (keep_open and success):
                try:
                    unpack(await client.call_tool("office_close_document", {"document": document}))
                except Exception:
                    if success:
                        raise
                    print(f"Cleanup could not close {path}; retained for inspection.", file=sys.stderr)


def error_message(exc):
    if isinstance(exc, BaseExceptionGroup):
        return '; '.join(error_message(child) for child in exc.exceptions)
    return str(exc)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--engine", choices=("microsoft", "wps", "all"), default="all")
    parser.add_argument("--kind", choices=("word", "excel", "ppt", "all"), default="all")
    parser.add_argument("--keep-open", action="store_true")
    parser.add_argument("--installed", action="store_true")
    parser.add_argument("--mode", choices=("auto", "legacy"), default="legacy")
    args = parser.parse_args()
    folder = ROOT / "data/live-smoke" / dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    folder.mkdir(parents=True)
    results = []
    for engine in (("microsoft", "wps") if args.engine == "all" else (args.engine,)):
        for kind in (EXAMPLES if args.kind == "all" else (args.kind,)):
            try:
                result = asyncio.run(run(engine, kind, folder, args.keep_open, args.installed, args.mode))
            except Exception as exc:
                result = {"engine": engine, "kind": kind, "passed": False, "error": error_message(exc)}
                print(json.dumps(result, ensure_ascii=False), flush=True)
            results.append(result)
            (folder / "report.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Artifacts: " + str(folder))
    return 0 if all(r['passed'] for r in results) else 1


if __name__ == "__main__":
    sys.exit(main())
