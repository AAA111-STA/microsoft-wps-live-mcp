"""Regressions for visible defaults, staged saving, shared references, and cancellation."""
import asyncio
from pathlib import Path
import sys
import threading
from types import SimpleNamespace
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from mcp import Client
from microsoft_wps_live_mcp.live_view import LiveView
from microsoft_wps_live_mcp.native import OfficeError
from microsoft_wps_live_mcp.server import build_server


def stages():
    return [
        {"label": "Frame", "steps": [{"action": "get", "member": "Content", "as": "text"}]},
        {"label": "Write", "steps": [{"action": "set", "target": "text", "member": "Text", "value": "Draft"}]},
    ]


class FakeView(LiveView):
    def __init__(self, fail=None):
        self.record = SimpleNamespace(id="doc", doc=SimpleNamespace(Saved=True), live_enabled=False)
        self.calls = []
        self.fail = fail

    def resolve(self, _):
        return self.record

    def _ensure_live_window(self, record, foreground=False):
        return None

    def _info(self, record):
        return {"document_id": record.id}

    def _paint(self, record, pause_ms=0):
        pass

    def live_checkpoint(self, record, focus=None, pause_ms=None, label=None):
        return {"enabled": True, "state": "ready", "label": label}

    def api_batch(self, document, steps, save, _aliases=None):
        self.calls.append((steps, save, _aliases))
        if len(self.calls) == 1:
            _aliases["text"] = "native object"
        else:
            assert _aliases["text"] == "native object"
        if self.fail == len(self.calls):
            return {"state": "partial_error", "disk_saved": False, "results": []}
        return {"state": "completed", "disk_saved": True, "results": []}


class LiveViewTests(unittest.TestCase):
    def test_each_stage_saves_and_keeps_aliases_on_the_same_native_workflow(self):
        engine = FakeView()
        progress = []
        result = engine.live_steps("doc", stages(), 0, _on_progress=progress.append)
        self.assertEqual(result["completed_stages"], 2)
        self.assertEqual([p["stage"] for p in progress], [1, 2])
        self.assertTrue(all(call[1] for call in engine.calls))
        self.assertIs(engine.calls[0][2], engine.calls[1][2])

    def test_bad_later_stage_is_rejected_before_any_native_edit(self):
        engine = FakeView()
        plan = stages()
        plan[1]["steps"][0]["action"] = "invalid"
        with self.assertRaises(OfficeError):
            engine.live_steps("doc", plan, 0)
        self.assertEqual(engine.calls, [])

    def test_failed_stage_does_not_run_later_stages(self):
        engine = FakeView(fail=1)
        result = engine.live_steps("doc", stages(), 0)
        self.assertEqual(result["state"], "partial_error")
        self.assertEqual(result["failed_stage"], 1)
        self.assertEqual(len(engine.calls), 1)

    def test_cancel_stops_between_saved_stages(self):
        engine = FakeView()
        stop = threading.Event()
        result = engine.live_steps("doc", stages(), 0, _on_progress=lambda _: stop.set(), _cancel_event=stop)
        self.assertEqual(result["state"], "cancelled")
        self.assertEqual(result["completed_stages"], 1)
        self.assertTrue(result["disk_saved"])

    def test_busy_at_later_stage_preserves_saved_history(self):
        engine = FakeView()
        original = engine.api_batch

        def run(*args, **kwargs):
            if engine.calls:
                raise OfficeError('Native application is busy')
            return original(*args, **kwargs)

        engine.api_batch = run
        result = engine.live_steps('doc', stages(), 0)
        self.assertEqual(result['state'], 'partial_error')
        self.assertEqual(result['completed_stages'], 1)
        self.assertEqual(result['failed_stage'], 2)
        self.assertTrue(result['stages'][0]['disk_saved'])
        self.assertTrue(result['inspect_before_retry'])
        self.assertTrue(engine.record.autosave_paused)

    def test_mcp_visible_defaults_and_live_progress(self):
        class Worker:
            async def call(self, method, *args, **kwargs):
                for index, stage in enumerate(args[1], 1):
                    kwargs["_on_progress"]({"stage": index, "total": len(args[1]), "label": stage["label"]})
                    await asyncio.sleep(0)
                return {"state": "completed", "completed_stages": len(args[1])}

        async def check():
            progress = []

            async def notify(number, total, message):
                progress.append((number, total, message))

            async with Client(build_server(Worker()), mode="legacy") as client:
                listed = await client.list_tools()
                schemas = {t.name: t.input_schema for t in listed.tools}
                for name in ("office_open_document", "office_create_document"):
                    self.assertTrue(schemas[name]["properties"]["visible"]["default"])
                self.assertNotIn("ctx", schemas["office_live_steps"]["properties"])
                result = await client.call_tool("office_live_steps", {"document": "doc", "stages": stages()}, progress_callback=notify)
                self.assertFalse(result.is_error)
            self.assertEqual([p[0] for p in progress], [1, 2])
        asyncio.run(check())

    def test_expected_native_error_reaches_the_agent(self):
        class Worker:
            async def call(self, *args, **kwargs):
                raise OfficeError('Native application is busy. Inspect before retrying.')

        async def check():
            async with Client(build_server(Worker()), mode='legacy') as client:
                result = await client.call_tool('office_open_document', {'path': 'F:/test.docx'})
                self.assertTrue(result.is_error)
                self.assertIn('Native application is busy', result.content[0].text)
                self.assertIn('Inspect before retrying', result.content[0].text)
        asyncio.run(check())


if __name__ == "__main__":
    unittest.main()
