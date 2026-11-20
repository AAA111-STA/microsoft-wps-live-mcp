"""Exercise real stdio MCP calls and general component editing in both desktop suites."""
from __future__ import annotations

import argparse
import asyncio
import datetime as dt
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

from mcp import Client, StdioServerParameters

ROOT = Path(__file__).resolve().parents[1]


def unpack(result):
    if result.is_error:
        raise RuntimeError(str(result.content))
    data = result.structured_content
    if data is None:
        data = json.loads(result.content[0].text)
    if data.get("state") in ("partial_error", "pending"):
        raise RuntimeError(json.dumps(data))
    return data


async def run(engine, root, mode, installed=False):
    server_args = ["-m", "microsoft_wps_live_mcp"] if installed else [str(ROOT / "scripts" / "run_mcp.py")]
    parameters = StdioServerParameters(command=sys.executable,
                                      args=["-X", "utf8", "-u", *server_args,
                                            "--data-dir", str(root / "backups")])
    async with Client(parameters, mode=mode, read_timeout_seconds=120) as client:
        listed = await client.list_tools()
        names = {tool.name for tool in listed.tools}
        assert {"office_api_batch", "office_api_describe", "office_api_constants", "office_word_read",
                "office_excel_read_range", "office_ppt_read"} <= names
        print(json.dumps({"protocol": client.protocol_version, "tools": len(names), "engine": engine}), flush=True)
        results = []
        for kind, suffix, example in (("word", ".docx", "word_layout.json"),
                                       ("excel", ".xlsx", "excel_chart.json"),
                                       ("ppt", ".pptx", "ppt_flowchart.json")):
            path = root / (engine + "-" + kind + suffix)
            info = unpack(await client.call_tool("office_create_document", {"path": str(path), "engine": engine}))
            doc = info["document_id"]
            try:
                described = unpack(await client.call_tool("office_api_describe", {"document": doc, "limit": 10}))
                assert described["members"]
                steps = json.loads((ROOT / "examples" / example).read_text(encoding="utf-8"))
                batch = unpack(await client.call_tool("office_api_batch", {"document": doc, "steps": steps}))
                assert batch["disk_saved"] and batch["state"] == "completed", batch
                with zipfile.ZipFile(path) as archive:
                    if kind == "excel":
                        assert any(n.startswith("xl/charts/chart") for n in archive.namelist()), archive.namelist()
                    elif kind == "ppt":
                        xml = ET.fromstring(archive.read("ppt/slides/slide1.xml"))
                        ns = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
                        assert len(xml.findall(".//p:cxnSp", ns)) == 2
                        assert len(xml.findall(".//p:sp", ns)) == 4
                if kind == "ppt":
                    # Export an actual render through the same general MCP object model.
                    image = root / (engine + "-workflow.png")
                    unpack(await client.call_tool("office_api_batch", {"document": doc, "steps": [
                        {"action": "call", "target": "document.Slides", "member": "Item", "args": [1], "as": "slide", "read_only": True},
                        {"action": "call", "target": "slide", "member": "Export", "args": [str(image), "PNG", 960, 540]}
                    ]}))
                    assert image.is_file()
                results.append({"engine": engine, "kind": kind, "case": "stdio_generic_components", "passed": True})
                print(json.dumps(results[-1]), flush=True)
            finally:
                unpack(await client.call_tool("office_close_document", {"document": doc}))
        return results


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--engine", choices=("microsoft", "wps", "all"), default="all")
    parser.add_argument("--mode", default="auto", help="SDK connection mode: auto or legacy")
    parser.add_argument("--installed", action="store_true", help="Launch the installed package instead of the source entrypoint")
    args = parser.parse_args()
    root = ROOT / "data" / "mcp-smoke" / dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    root.mkdir(parents=True)
    results = []
    for engine in (("microsoft", "wps") if args.engine == "all" else (args.engine,)):
        results.extend(asyncio.run(run(engine, root, args.mode, args.installed)))
    (root / "report.json").write_text(json.dumps(results, indent=2), encoding="utf-8")
    print("Artifacts: " + str(root))


if __name__ == "__main__":
    main()
