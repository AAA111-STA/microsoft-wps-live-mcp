"""Opt-in verification of review components, automatic user-edit saving, and independent copies."""
from __future__ import annotations

import argparse
import asyncio
import datetime as dt
import json
from pathlib import Path
import sys
import time

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from microsoft_wps_live_mcp.native import NativeOffice, file_hash
from microsoft_wps_live_mcp.worker import OfficeWorker


class ManualInputProbe(NativeOffice):
    def simulate_manual_input(self, document):
        # Direct native buffer mutation, deliberately bypassing all MCP edit/save
        # helpers. The idle observer must notice Saved=False by itself.
        r = self.resolve(document)
        if r.kind == "word":
            end = int(r.doc.Content.End) - 1
            r.doc.Range(end, end).InsertBefore("MANUAL_EDIT_AUTOSAVED\r")
        elif r.kind == "excel":
            r.doc.Worksheets.Item(1).Range("A1").Value2 = "MANUAL_EDIT_AUTOSAVED"
        else:
            slide = r.doc.Slides.Add(1, 12)
            shape = slide.Shapes.AddTextbox(1, 10, 10, 200, 50)
            shape.TextFrame.TextRange.Text = "MANUAL_EDIT_AUTOSAVED"
        return {"buffer_mutated": True}


async def run(engine, root):
    import pythoncom
    pythoncom.CoInitialize()
    worker = OfficeWorker(root / "runtime", factory=ManualInputProbe)
    results = []
    document = None
    try:
        path = root / (engine + "-review.docx")
        info = await worker.call("create", str(path), engine)
        document = info["document_id"]
        await worker.call("word_insert", document, "First paragraph\nSecond paragraph\n")
        await worker.call("word_snapshot", document, "take")
        await worker.call("word_insert", document, "New text\n")
        assert (await worker.call("word_snapshot", document, "diff"))["changed_blocks"] > 0
        results.append({"feature": "snapshot_diff", "passed": True})
        await worker.call("word_comments", document, "add", None, 0, 5, "Please review")
        assert len((await worker.call("word_comments", document))["comments"]) == 1
        await worker.call("word_comments", document, "delete", 1)
        results.append({"feature": "comments_add_read_delete", "passed": True})
        await worker.call("word_revisions", document, "track", True)
        await worker.call("word_insert", document, "Tracked text\n")
        assert (await worker.call("word_revisions", document))["revisions"]
        await worker.call("word_revisions", document, "accept")
        await worker.call("word_revisions", document, "track", False)
        results.append({"feature": "track_list_accept_revisions", "passed": True})
        for component, options in (("footnote", {"text": "Footnote content"}),
                                   ("endnote", {"text": "Endnote content"}),
                                   ("bookmark", {"name": "CopyAnchor"}),
                                   ("field", {"code": "PAGE"}),
                                   ("cross_reference", {"reference_type": 2, "reference_kind": 7, "reference_item": "CopyAnchor"}),
                                   ("equation", {"text": "x^2+y^2=z^2"})):
            try:
                await worker.call("word_component", document, component, options)
                if component == "equation":
                    info = await worker.call("inspect_document", document)
                    assert info["components"]["OMaths"] > 0, "No editable native equation was created"
                results.append({"feature": component, "passed": True})
            except Exception as exc:
                results.append({"feature": component, "passed": False, "error": str(exc)})
                # Retain and explicitly save any partial native edit before continuing the test.
                await worker.call("save", document)
        pdf = root / (engine + "-review.pdf")
        await worker.call("export_pdf", document, str(pdf))
        results.append({"feature": "pdf_export", "passed": True})

        before = file_hash(path)
        await worker.call("simulate_manual_input", document)
        for _ in range(15):
            await asyncio.sleep(0.2)
            if file_hash(path) != before:
                break
        assert file_hash(path) != before, "Observer did not save manual input"
        info = await worker.call("inspect_document", document)
        assert file_hash(Path(info["recovery_copy"])) == file_hash(path), info
        results.append({"feature": "manual_edit_autosave_and_recovery_copy", "passed": True})
        copied = await worker.call("create_copy", document)
        assert file_hash(Path(copied["copy_path"])) == file_hash(path)
        assert Path(copied["path"]).resolve() == path.resolve()
        results.append({"feature": "independent_final_copy", "passed": True})
        # Even an explicit staging flag on the public worker cannot bypass autosave.
        forced = await worker.call("word_insert", document, "FORCED_AUTOSAVE\n", "end", False)
        assert forced["disk_saved"] and forced["autosave_enforced"]
        results.append({"feature": "agent_edit_always_saved", "passed": True})
        await worker.call("close", document)
        document = None
        for kind, suffix in (("excel", ".xlsx"), ("ppt", ".pptx")):
            other_path = root / (engine + "-autosave-" + kind + suffix)
            other = await worker.call("create", str(other_path), engine)
            document = other["document_id"]
            before = file_hash(other_path)
            await worker.call("simulate_manual_input", document)
            for _ in range(20):
                await asyncio.sleep(0.2)
                if file_hash(other_path) != before:
                    break
            assert file_hash(other_path) != before, kind + " observer failed to save"
            info = await worker.call("inspect_document", document)
            assert info["saved"] and file_hash(Path(info["recovery_copy"])) == file_hash(other_path), info
            copied = await worker.call("create_copy", document)
            assert file_hash(Path(copied["copy_path"])) == file_hash(other_path)
            results.append({"feature": kind + "_manual_autosave_recovery_and_final_copy", "passed": True})
            await worker.call("close", document)
            document = None
    finally:
        if document:
            try:
                await worker.call("close", document)
            except Exception:
                pass
        worker.stop()
        worker.thread.join(3)
        pythoncom.CoUninitialize()
    print(json.dumps({"engine": engine, "results": results}, ensure_ascii=False), flush=True)
    return {"engine": engine, "results": results}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--engine", choices=("microsoft", "wps", "all"), default="all")
    args = parser.parse_args()
    root = ROOT / "data" / "feature-smoke" / dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    root.mkdir(parents=True)
    results = []
    for engine in (("microsoft", "wps") if args.engine == "all" else (args.engine,)):
        results.append(asyncio.run(run(engine, root)))
    (root / "report.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Artifacts: " + str(root))


if __name__ == "__main__":
    main()
