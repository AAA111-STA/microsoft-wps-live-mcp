"""Opt-in integration checks against real desktop apps, using newly created test files only."""
from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from microsoft_wps_live_mcp.com_dispatch import dispatch
from microsoft_wps_live_mcp.discovery import PROGIDS, vendor_of
from microsoft_wps_live_mcp.native import NativeOffice, file_hash


def disk_text(path):
    with zipfile.ZipFile(path) as archive:
        return "\n".join("".join(ET.fromstring(archive.read(name)).itertext())
                         for name in archive.namelist() if name.endswith(".xml")
                         and name.startswith(("word/", "xl/", "ppt/slides/")))


def check(engine, root):
    import pythoncom

    pythoncom.CoInitialize()
    native = NativeOffice(root / "backups")
    result = []
    try:
        for kind, suffix in (("word", ".docx"), ("excel", ".xlsx"), ("ppt", ".pptx")):
            path = root / (engine + "-" + kind + suffix)
            document = None
            external = None
            try:
                info = native.create(str(path), engine)
                document = info["document_id"]
                assert info["disk_saved"] and info["engine"] == engine, info
                if kind == "word":
                    native.word_insert(document, "initial paragraph\n")
                    native.word_add_table(document, [["A", "B"], [1, 2]])
                elif kind == "excel":
                    native.excel_write(document, 1, "A1:B2", [["initial", 2], [3, 4]])
                    native.excel_write(document, 1, "C1", [["=SUM(B1:B2)"]], formulas=True)
                    native.excel_format(document, 1, "A1:B2", bold=True, fill="#DDEEFF")
                else:
                    native.ppt_add_slide(document, "initial title", "initial body")
                    native.ppt_format(document, 1, 1, bold=True, color="#4472C4")
                native.close(document)
                document = None
                assert "initial" in disk_text(path), path
                result.append({"kind": kind, "case": "create_and_save", "passed": True})

                # Simulate a document opened independently by the user, including unsaved text.
                app = dispatch(pythoncom.CoCreateInstanceEx(PROGIDS[engine][kind][0], None,
                    pythoncom.CLSCTX_SERVER, None, (pythoncom.IID_IDispatch,))[0])
                assert vendor_of(app) == engine, 'External fixture opened in the wrong product'
                if kind == "word":
                    external = app.Documents.Open(str(path))
                    external.Range(0, 0).InsertBefore("USER_UNSAVED\r")
                elif kind == "excel":
                    external = app.Workbooks.Open(str(path))
                    external.Worksheets.Item(1).Range("D1").Value2 = "USER_UNSAVED"
                else:
                    external = app.Presentations.Open(str(path), False, False, False)
                    external.Slides.Item(1).Shapes.Item(2).TextFrame.TextRange.Text = "USER_UNSAVED"
                before = file_hash(path)
                info = native.open(str(path), engine)
                document = info["document_id"]
                assert info["already_open"], info
                assert not info["opened_by_server"], info
                if kind == "word":
                    native.word_insert(document, "AGENT_LIVE\n", save=False)
                    assert "AGENT_LIVE" in str(external.Content.Text)
                    native.word_replace(document, "initial", "replaced", save=False)
                    native.word_format(document, 0, 4, bold=True, color="#4472C4", save=False)
                    assert native.word_read(document)["tables"] == 1
                elif kind == "excel":
                    native.excel_write(document, 1, "A3", [["AGENT_LIVE"]], save=False)
                    assert external.Worksheets.Item(1).Range("A3").Value2 == "AGENT_LIVE"
                    native.excel_write(document, 1, "E1", [["=literal"]], save=False)
                    assert native.excel_read(document, 1, "E1")["values"] == [["=literal"]]
                    native.excel_write(document, 1, "E2:F2", [["00123", "2026-01-02"]], save=False)
                    assert native.excel_read(document, 1, "E2:F2")["values"] == [["00123", "2026-01-02"]]
                    native.excel_sheets(document, "add", "Smoke", save=False)
                    native.excel_sheets(document, "rename", "Smoke", "Renamed", save=False)
                else:
                    native.ppt_set_text(document, 1, 1, "AGENT_LIVE", save=False)
                    assert external.Slides.Item(1).Shapes.Item(1).TextFrame.TextRange.Text == "AGENT_LIVE"
                    native.ppt_add_textbox(document, 1, "extra textbox", save=False)
                    assert native.ppt_read(document)["total_slides"] == 1
                assert file_hash(path) == before, "save=False unexpectedly changed the disk file"
                native.save(document)
                assert "AGENT_LIVE" in disk_text(path) and "USER_UNSAVED" in disk_text(path)
                native.close(document)
                document = None
                external = None
                result.append({"kind": kind, "case": "attach_live_preserve_user_edits_and_save", "passed": True})

                # Explicitly select WPS on dual-installed hosts; auto intentionally prefers
                # Microsoft for closed files. Both products must load a closed file correctly.
                reopened = native.open(str(path), engine)
                assert not reopened["already_open"] and reopened["engine"] == engine, reopened
                document = reopened["document_id"]
                if kind == "word":
                    native.word_insert(document, "CLOSED_FILE_EDIT\n")
                elif kind == "excel":
                    native.excel_write(document, 1, "A4", [["CLOSED_FILE_EDIT"]])
                else:
                    native.ppt_set_text(document, 1, 1, "CLOSED_FILE_EDIT")
                assert "CLOSED_FILE_EDIT" in disk_text(path)
                native.close(document)
                document = None
                result.append({"kind": kind, "case": "closed_file_auto_open_and_edit", "passed": True})
            except Exception as exc:
                result.append({"kind": kind, "case": "integration", "passed": False,
                               "error": f"{type(exc).__name__}: {exc}"})
            finally:
                # Only test documents created in this run are eligible for cleanup.
                if path.exists():
                    for record in list(native.documents.values()):
                        if record.path and Path(record.path).resolve() == path.resolve():
                            try:
                                native.close(record.id)
                            except Exception:
                                pass
                if external is not None:
                    try:
                        external.Close()
                    except Exception:
                        pass
            print(json.dumps({"engine": engine, "kind": kind, "results": result[-3:]}, ensure_ascii=False), flush=True)
    finally:
        native = None
        pythoncom.CoUninitialize()
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--engine", choices=("microsoft", "wps", "all"), required=True)
    parser.add_argument("--output-dir", type=Path)
    args = parser.parse_args()
    root = args.output_dir or Path(__file__).resolve().parents[1] / "data" / "native-smoke" / dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    root = root.resolve()
    root.mkdir(parents=True, exist_ok=False)
    results = {}
    for engine in (("microsoft", "wps") if args.engine == "all" else (args.engine,)):
        results[engine] = check(engine, root)
    report = {"results": results, "output_dir": str(root)}
    (root / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if all(r["passed"] for rows in results.values() for r in rows) else 1


if __name__ == "__main__":
    sys.exit(main())

