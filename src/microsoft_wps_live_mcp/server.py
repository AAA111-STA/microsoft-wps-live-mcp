"""Standard stdio MCP server. stdout is reserved exclusively for MCP messages."""
from __future__ import annotations

import argparse
import asyncio
import atexit
import json
import os
from pathlib import Path
import sys
import threading
from typing import Any, Literal

from mcp.server.mcpserver import Context, MCPServer
from mcp.types import ToolAnnotations

from . import __title__, __version__
from .discovery import capabilities
from .worker import OfficeWorker
from .clients import PROFILES, client_config, write_configs


INSTRUCTIONS = """Edit native Microsoft Office and WPS Office documents on Windows.
ALL tasks default to visible, progressive editing in the actual document window, across Word, Excel and PPT.
Use these MCP tools directly. No separate browser editor, takeover page, or "let the agent operate"
button is required for native editing. Do not add a per-step user confirmation flow.
Create/open visibly; only use visible=False when the user explicitly requests background work.
Plan short meaningful stages (a paragraph/section, data block, chart, shape group or layout change).
Use office_live_steps for a multi-stage object-model workflow; it paints and saves each named stage.
For convenience tools, use several small calls rather than a single opaque whole-document replacement.
office_show_document navigates the native window to the exact paragraph, range or slide being edited.
Do not disable screen updating, hide the document, silently switch to offline file generation, or close
the completed document unless requested. Keep the native output visible and editable for the user.
If live_view reports unavailable or focus_error, inspect and fix display before claiming visible success.
Call office_status first. Reuse existing open documents; do not edit their OOXML files directly.
Use office_open_document(engine='wps' or 'microsoft') to choose a product explicitly.
Every document argument accepts a document_id or an absolute .docx/.xlsx/.pptx path.
Paths are automatically opened when needed. IDs are scoped to this MCP process.
Paragraph, worksheet, slide and shape indices are 1-based. Word character offsets are native 0-based offsets.
Read before editing; do not guess sheet names or shape indices. Successful mutations automatically save.
Automatic saving also observes manual edits in explicitly connected documents while this MCP is running.
Any internal save=False staging is followed by enforced save at the public MCP operation boundary.
Each save updates a separate recovery copy; office_create_copy produces an additional user-visible final copy.
Before finishing an editing task, call office_create_copy and report both original and final-copy paths.
Report disk_saved=True only when the tool confirms it. Report recovery-copy failures separately.
Existing user edits are part of the live document and will be included when it is saved.
Timed-out operations may complete later. Query office_operation_status; never blindly replay a modification.
The convenience tools are examples, not a task whitelist. office_api_describe and office_api_constants
discover native members/enums; office_api_batch composes any public native get/set/call operations.
This includes shapes, connectors, tables, charts, images, themes, layouts, formulas, fields,
comments, animations and other components when the installed application exposes them.
Use the document root for the selected document, application only for requested app-level actions.
Do not close other documents, quit apps, change global settings or run macros unless requested.
Generic batches are not transactional; inspect partial_error results before deciding how to continue.
Do not promise every Office feature, WPS edition, OS, protected file, or unattended service is supported.
"""


def build_server(worker=None):
    worker = worker or OfficeWorker()
    server = MCPServer("microsoft-wps-live", title=__title__, version=__version__, instructions=INSTRUCTIONS)
    read = ToolAnnotations(read_only_hint=True, destructive_hint=False, open_world_hint=False)
    edit = ToolAnnotations(read_only_hint=False, destructive_hint=False, open_world_hint=False)

    @server.tool(annotations=read)
    async def office_status() -> dict:
        """Detect desktop Office/WPS COM registrations and list open documents without starting apps."""
        if sys.platform not in ("win32", "darwin"):
            return {**capabilities(), "documents": []}
        return await worker.call("status")

    @server.tool(annotations=read)
    async def office_list_open_documents(engine: Literal["auto", "microsoft", "wps"] = "auto") -> dict:
        """List already-open Word, Excel, PPT documents, including untitled documents; returns document IDs."""
        return await worker.call("list_open", engine)

    @server.tool(annotations=edit)
    async def office_open_document(path: str, engine: Literal["auto", "microsoft", "wps"] = "auto",
                                   visible: bool = True) -> dict:
        """Attach to an existing document, or load a closed file in native Office/WPS. Never hides existing windows.

        auto follows the product already editing this path; for closed files prefers Microsoft, then WPS.
        Visible editing is the default for every task. Only use visible=False on an explicit background-work request.
        Returns document_id for subsequent calls.
        """
        return await worker.call("open", path, engine, visible)

    @server.tool(annotations=edit)
    async def office_create_document(path: str, engine: Literal["auto", "microsoft", "wps"] = "auto",
                                     visible: bool = True) -> dict:
        """Create and save a new .docx/.xlsx/.pptx using native Office/WPS. Never overwrites an existing file."""
        return await worker.call("create", path, engine, visible)

    @server.tool(annotations=edit)
    async def office_show_document(document: str, focus: dict[str, Any] | None = None) -> dict:
        """Show the real native editing window and navigate to the work in progress. No file content is changed.

        Word focus={start:0,end:20} (native character offsets);
        Excel focus={sheet:'Sheet1',address:'A1:D10'}; PPT focus={slide:2,shape:1} (shape optional).
        Can make a previously hidden native document visible without rebuilding it. Returns observed live_view status.
        """
        return await worker.call("show_document", document, focus)

    @server.tool(annotations=ToolAnnotations(read_only_hint=False, destructive_hint=True, open_world_hint=False))
    async def office_live_steps(document: str, stages: list[dict[str, Any]], ctx: Context,
                                 pause_ms: int = 450) -> dict:
        """Build or edit ANY Office task in named visible stages, saving each stage as the user watches.

        stages=[{label:'Create title',steps:[...native get/set/call steps...],focus:{slide:1}}, ...].
        Steps use office_api_batch syntax. Aliases carry across stages within this call.
        Focus can address Word start/end offsets, Excel sheet/address, or PPT slide/shape.
        pause_ms is 0..2000 ms between stages, default 450. Keep stages short and meaningful.
        Emits standard MCP progress where supported. Stages are not transactions; stop on partial_error/display_error.
        For long calls returning pending, query office_operation_status and do not replay the workflow.
        """
        loop = asyncio.get_running_loop()
        events = asyncio.Queue()
        active = True
        cancelled = threading.Event()

        def on_progress(event):
            if active and not loop.is_closed():
                loop.call_soon_threadsafe(events.put_nowait, event)

        running = asyncio.create_task(worker.call("live_steps", document, stages, pause_ms,
                                                  _on_progress=on_progress, _cancel_event=cancelled))
        waiting = None
        try:
            while not running.done():
                waiting = asyncio.create_task(events.get())
                done, _ = await asyncio.wait({running, waiting}, return_when=asyncio.FIRST_COMPLETED)
                if waiting in done:
                    event = waiting.result()
                    try:
                        await ctx.report_progress(event["stage"], event["total"], event["label"])
                    except Exception:
                        pass  # Native visibility does not depend on a host's progress UI.
                else:
                    waiting.cancel()
                    await asyncio.gather(waiting, return_exceptions=True)
            while not events.empty():
                event = events.get_nowait()
                try:
                    await ctx.report_progress(event["stage"], event["total"], event["label"])
                except Exception:
                    pass
            return await running
        finally:
            active = False
            if waiting is not None and not waiting.done():
                waiting.cancel()
                await asyncio.gather(waiting, return_exceptions=True)
            if not running.done():
                cancelled.set()
                running.cancel()  # Stops awaiting; never cancels or replays an in-flight native call.

    @server.tool(annotations=edit)
    async def office_save_document(document: str, save_as: str | None = None) -> dict:
        """Save the live document, including user edits, and verify disk output. Save As requires an unused absolute path."""
        return await worker.call("save", document, save_as)

    @server.tool(annotations=edit)
    async def office_create_copy(document: str, path: str | None = None) -> dict:
        """Save and produce an independent verified final copy, keeping the original document path unchanged.

        path omitted creates a uniquely named sibling .copy-TIMESTAMP file. Never overwrites an existing file.
        Use at the end of editing tasks and return the copy path to the user. Automatic recovery copies also update on every save.
        """
        return await worker.call("create_copy", document, path)

    @server.tool(annotations=edit)
    async def office_close_document(document: str) -> dict:
        """Save and close this document only. Never quits the app or discards user edits."""
        return await worker.call("close", document, True)

    @server.tool(annotations=read)
    async def office_operation_status(operation_id: str) -> dict:
        """Retrieve the result of a pending/timed-out native operation. Do not replay modifications after timeout."""
        return worker.operation_status(operation_id)

    @server.tool(annotations=read)
    async def office_inspect_document(document: str) -> dict:
        """Inventory native components in Word, Excel or PPT: tables, charts, shapes, notes, revisions, masters/animations, etc.

        Unsupported native inventory properties are marked available=false individually.
        """
        return await worker.call("inspect_document", document)

    @server.tool(annotations=edit)
    async def office_export_pdf(document: str, path: str) -> dict:
        """Export the current live Word/Excel/PPT document to a new PDF using its native application. Never overwrites output."""
        return await worker.call("export_pdf", document, path)

    @server.tool(annotations=read)
    async def office_api_describe(document: str, target: str = "document", search: str = "", limit: int = 100) -> dict:
        """Discover public native COM properties/methods and argument names from installed Office/WPS type information.

        target: document, application, returned object_id, or dotted property path such as document.PageSetup.
        Use office_api_batch to obtain indexed child objects and store their returned object_id.
        """
        return await worker.call("api_describe", document, target, search, limit)

    @server.tool(annotations=read)
    async def office_api_constants(document: str, search: str = "", limit: int = 100) -> dict:
        """Discover native enum values from the installed product's type libraries, including shared Office shape/arrow enums."""
        return await worker.call("api_constants", document, search, limit)

    @server.tool(annotations=ToolAnnotations(read_only_hint=False, destructive_hint=True, open_world_hint=False))
    async def office_api_batch(document: str, steps: list[dict[str, Any]]) -> dict:
        """Compose ANY public native Office/WPS object-model operations, beyond the convenience tools.

        Each step: action=get/set/call, target (default document), member, optional args array,
        kwargs object, value for set, as alias. Example: call document.Slides.Add(1,12) with
        {"action":"call","target":"document.Slides","member":"Add","args":[1,12],"as":"slide"}.
        Then target="slide.Shapes", member="AddShape", args=[1,40,40,200,80], as="box";
        set target="box.TextFrame.TextRange", member="Text", value="Data processing".
        Arguments can reference previous aliases or returned object IDs using {"$ref":"box"}.
        Output COM objects become object_id references usable across calls in this document.
        Indexed-property calls (Item, Range, etc.) can set read_only=true when genuinely read-only.
        Batches stop at the first error, keep earlier live edits, and return partial_error with completed results.
        Successful modifying batches automatically save, including existing user edits.
        Any installed native component is addressable; unsupported members fail explicitly.
        Uses native COM, not Python eval/exec. This grants the same desktop access as the native API.
        """
        return await worker.call("api_batch", document, steps, True)

    @server.tool(annotations=read)
    async def office_word_read(document: str, start: int = 1, count: int = 100) -> dict:
        """Read current live Word/WPS paragraphs, table count, and native character offsets. Paragraphs are 1-based."""
        return await worker.call("word_read", document, start, count)

    @server.tool(annotations=read)
    async def office_word_find(document: str, text: str, match_case: bool = False,
                                whole_word: bool = False, wildcards: bool = False, limit: int = 100) -> dict:
        """Find native Word/WPS text and return exact start/end offsets and context. Supports native Word wildcard syntax."""
        return await worker.call("word_find", document, text, match_case, whole_word, wildcards, limit)

    @server.tool(annotations=read)
    async def office_word_snapshot(document: str, action: Literal["take", "status", "diff"] = "take") -> dict:
        """Take an in-memory paragraph snapshot, check snapshot status, or compare current live text with the snapshot.

        diff reports inserted/deleted/replaced text blocks and does not update the baseline; formatting is not compared.
        """
        return await worker.call("word_snapshot", document, action)

    @server.tool(annotations=edit)
    async def office_word_comments(document: str, action: Literal["list", "add", "reply", "resolve", "reopen", "delete"] = "list",
                                    index: int | None = None, start: int | None = None, end: int | None = None,
                                    text: str | None = None) -> dict:
        """Read or edit native comments. Add uses start/end character offsets; other edits use 1-based comment index.

        Threaded replies and resolution depend on installed product/version. list is read-only and never saves.
        """
        return await worker.call("word_comments", document, action, index, start, end, text, True)

    @server.tool(annotations=edit)
    async def office_word_revisions(document: str, action: Literal["list", "track", "accept", "reject"] = "list",
                                     enabled: bool | None = None, indices: list[int] | None = None,
                                     author: str | None = None) -> dict:
        """List native tracked changes, explicitly enable/disable tracking, or accept/reject changes by indices/author.

        indices omitted means all matching changes; all indices are 1-based. list is read-only and never saves.
        """
        return await worker.call("word_revisions", document, action, enabled, indices, author, True)

    @server.tool(annotations=edit)
    async def office_word_insert_component(document: str,
            component: Literal["image", "equation", "footnote", "endnote", "bookmark", "field", "hyperlink",
                               "page_break", "section_break", "toc", "cross_reference"],
            options: dict[str, Any], position: str | int = "end") -> dict:
        """Insert native Word components. Other component types remain accessible via office_api_batch.

        options: image={path,width?,height?} (points); equation={text} (UnicodeMath);
        footnote/endnote={text}; bookmark={name}; field={code}; hyperlink={address,text?,subaddress?};
        section_break={type:2}; toc={upper_level:1,lower_level:3}; page_break={};
        cross_reference={reference_type:2,reference_kind:-1,reference_item:'BookmarkName',as_hyperlink:true}.
        Query native constants for other reference types/kinds; caption reference types use the actual localized label.
        position: start/end or a native character offset. Native feature support varies by product/version.
        """
        return await worker.call("word_component", document, component, options, position, True)

    @server.tool(annotations=edit)
    async def office_word_undo(document: str, times: int = 1) -> dict:
        """Undo native Word/WPS undo entries. An entry is not guaranteed to equal one MCP call, especially in batches."""
        return await worker.call("word_undo", document, times, True)

    @server.tool(annotations=edit)
    async def office_word_insert_text(document: str, text: str, position: str | int = "end") -> dict:
        """Insert text without rebuilding the document. position is start, end, or a native 0-based character offset."""
        return await worker.call("word_insert", document, text, position, True)

    @server.tool(annotations=edit)
    async def office_word_replace_text(document: str, search: str, replacement: str, match_case: bool = False,
                                       whole_word: bool = False, replace_all: bool = True) -> dict:
        """Use native Find/Replace in the main document body; retains unrelated fields and formatting.

        Search and replacement each have a native 255-character limit. Native Word control codes are supported.
        """
        return await worker.call("word_replace", document, search, replacement, match_case, whole_word, replace_all, True)

    @server.tool(annotations=edit)
    async def office_word_format_text(document: str, start: int, end: int, font_name: str | None = None,
                                      font_size: float | None = None, bold: bool | None = None,
                                      italic: bool | None = None, color: str | None = None) -> dict:
        """Format a native Word character range; get start/end offsets from office_word_read. Colors use #RRGGBB."""
        return await worker.call("word_format", document, start, end, font_name, font_size, bold, italic, color, True)

    @server.tool(annotations=edit)
    async def office_word_add_table(document: str, values: list[list[str | int | float | bool | None]],
                                    position: str | int = "end") -> dict:
        """Insert a rectangular native Word table at end or a character offset. Does not rebuild other document content."""
        return await worker.call("word_add_table", document, values, position, True)

    @server.tool(annotations=read)
    async def office_excel_read_range(document: str, sheet: str | int, address: str) -> dict:
        """Read current native Excel/WPS values and formulas. sheet is its actual name or 1-based index; address is A1:B3."""
        return await worker.call("excel_read", document, sheet, address)

    @server.tool(annotations=edit)
    async def office_excel_write_range(document: str, sheet: str | int, address: str,
                                       values: list[list[str | int | float | bool | None]],
                                       formulas: bool = False) -> dict:
        """Write an exact-sized 2D array into a native range. Strings are literal by default; set formulas=True for formulas."""
        return await worker.call("excel_write", document, sheet, address, values, formulas, True)

    @server.tool(annotations=edit)
    async def office_excel_sheets(document: str, action: Literal["list", "add", "rename"] = "list",
                                   name: str | None = None, new_name: str | None = None) -> dict:
        """List, add, or rename worksheets. add uses name; rename uses name and new_name. No worksheet deletion."""
        return await worker.call("excel_sheets", document, action, name, new_name, True)

    @server.tool(annotations=edit)
    async def office_excel_format_range(document: str, sheet: str | int, address: str,
                                        font_name: str | None = None, font_size: float | None = None,
                                        bold: bool | None = None, italic: bool | None = None,
                                        color: str | None = None, fill: str | None = None,
                                        number_format: str | None = None, column_width: float | None = None) -> dict:
        """Format range fonts, fill, number format, and column width. Colors use #RRGGBB."""
        return await worker.call("excel_format", document, sheet, address, font_name, font_size, bold,
                                 italic, color, fill, number_format, column_width, True)

    @server.tool(annotations=read)
    async def office_ppt_read(document: str, start: int = 1, count: int = 20) -> dict:
        """Inspect live PowerPoint/WPS slides and text shapes. Slide and shape indices are 1-based, unlike the old web bridge."""
        return await worker.call("ppt_read", document, start, count)

    @server.tool(annotations=edit)
    async def office_ppt_set_text(document: str, slide: int, shape: int, text: str) -> dict:
        """Replace text in an existing text shape. Read shape indices first; mixed character formatting may change."""
        return await worker.call("ppt_set_text", document, slide, shape, text, True)

    @server.tool(annotations=edit)
    async def office_ppt_add_slide(document: str, title: str = "", text: str = "",
                                   position: int | None = None) -> dict:
        """Add a blank-layout slide with optional title/body textboxes. position is 1-based, omitted means append."""
        return await worker.call("ppt_add_slide", document, title, text, position, True)

    @server.tool(annotations=edit)
    async def office_ppt_add_textbox(document: str, slide: int, text: str, left: float = 36,
                                     top: float = 100, width: float = 400, height: float = 100) -> dict:
        """Add a native editable textbox to a slide. Position and size are in points (72 points per inch)."""
        return await worker.call("ppt_add_textbox", document, slide, text, left, top, width, height, True)

    @server.tool(annotations=edit)
    async def office_ppt_format_text(document: str, slide: int, shape: int, font_name: str | None = None,
                                     font_size: float | None = None, bold: bool | None = None,
                                     italic: bool | None = None, color: str | None = None) -> dict:
        """Format a text shape in native PowerPoint/WPS; colors use #RRGGBB. Inspect slide/shape indices first."""
        return await worker.call("ppt_format", document, slide, shape, font_name, font_size, bold, italic, color, True)

    return server


def main():
    parser = argparse.ArgumentParser(description="Live MCP tools for native Microsoft Office and WPS Office (Windows).")
    parser.add_argument("--version", action="version", version=__version__)
    parser.add_argument("--doctor", action="store_true", help="Print compatibility and open-document status; no apps are started")
    parser.add_argument("--config", choices=PROFILES, help="Print installation-specific MCP client configuration")
    parser.add_argument("--write-configs", type=Path, help="Generate all supported client profiles in this local directory")
    parser.add_argument("--data-dir", type=Path, help="Local backup location (default: user application data directory)")
    args = parser.parse_args()
    if args.write_configs:
        print(json.dumps({"files": write_configs(args.write_configs)}, indent=2))
        return
    if args.config:
        print(client_config(args.config))
        return
    worker = OfficeWorker(args.data_dir)
    atexit.register(worker.stop)
    if args.doctor:
        result = asyncio.run(worker.call("status")) if sys.platform in ("win32", "darwin") else capabilities()
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return
    build_server(worker).run(transport="stdio")


if __name__ == "__main__":
    main()

