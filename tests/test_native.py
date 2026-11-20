"""Regression checks for live Office state, disk conflicts, and MCP concurrency."""
import asyncio
import json
from pathlib import Path
import sys
import tempfile
import threading
import time
import unittest
import zipfile
from unittest.mock import MagicMock, patch
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from microsoft_wps_live_mcp.native import NativeOffice, Document, OfficeError, a1_size, file_hash, rgb
from microsoft_wps_live_mcp.server import build_server, client_config
from microsoft_wps_live_mcp.worker import OfficeWorker
from microsoft_wps_live_mcp.discovery import canonical


class NativeTests(unittest.TestCase):
    def test_unready_external_document_does_not_hide_healthy_open_documents(self):
        native = NativeOffice.__new__(NativeOffice)
        native.documents = {}
        app = MagicMock()
        app.Documents.Count = 2
        native._apps = lambda kind: [('wps', app)] if kind == 'word' else []
        records = [SimpleNamespace(id='bad'), SimpleNamespace(id='good')]
        native._remember = MagicMock(side_effect=records)
        native._info = MagicMock(side_effect=[AttributeError('Item.Name'), {'document_id':'good'}])
        result = native.list_open()
        self.assertEqual(result['documents'], [{'document_id':'good'}])
        self.assertEqual(result['discovery_warnings'][0]['index'], 1)
        self.assertIn('Item.Name', result['discovery_warnings'][0]['error'])

    def test_new_document_activates_registered_server_without_compatibility_rot(self):
        native = NativeOffice.__new__(NativeOffice)
        native.apps = []
        native.com = MagicMock(CLSCTX_SERVER=21, IID_IDispatch='IDispatch')
        app = SimpleNamespace(Path='C:/Program Files/Microsoft Office/Office16')
        with patch('microsoft_wps_live_mcp.native.registration', return_value={'registered': True}), \
                patch('microsoft_wps_live_mcp.native.retry_busy', side_effect=lambda operation: operation()), \
                patch('microsoft_wps_live_mcp.native.dispatch', return_value=app) as wrapper:
            chosen, vendor = native._choose_app('excel', 'microsoft')
        self.assertIs(chosen, app)
        self.assertEqual(vendor, 'microsoft')
        native.com.CoCreateInstanceEx.assert_called_once_with('Excel.Application', None, 21, None, ('IDispatch',))
        wrapper.assert_called_once_with(native.com.CoCreateInstanceEx.return_value[0])

    def test_microsoft_request_does_not_accept_wps_advertising_office_identity(self):
        native = NativeOffice.__new__(NativeOffice)
        native.apps = []
        native.com = MagicMock()
        app = SimpleNamespace(Path='C:/Kingsoft/WPS Office/office6', Name='Microsoft Excel')
        with patch('microsoft_wps_live_mcp.native.registration', return_value={'registered': True}), \
                patch('microsoft_wps_live_mcp.native.retry_busy', side_effect=lambda operation: operation()), \
                patch('microsoft_wps_live_mcp.native.dispatch', return_value=app):
            with self.assertRaisesRegex(OfficeError, 'belongs to wps'):
                native._choose_app('excel', 'microsoft')
        self.assertEqual(native.apps, [])

    def test_save_produces_verified_recovery_copy_and_final_copy_without_renaming_original(self):
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            path = root / "sample.docx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", "<document/>")
            doc = MagicMock()
            doc.FullName = str(path)
            doc.Name = path.name
            doc.Saved = True
            doc.ReadOnly = False
            native = NativeOffice.__new__(NativeOffice)
            native.data_dir = root / "runtime"
            native.documents = {}
            record = Document("test", MagicMock(), doc, "microsoft", "word", canonical(path), file_hash(path), managed=True)
            native.documents[record.id] = record
            saved = native.save(record.id)
            self.assertEqual(file_hash(Path(saved["recovery_copy"])), file_hash(path))
            copied = native.create_copy(record.id)
            self.assertEqual(file_hash(Path(copied["copy_path"])), file_hash(path))
            self.assertEqual(copied["path"], canonical(path))
            with self.assertRaises(OfficeError):
                native.create_copy(record.id, copied["copy_path"])

    def test_autosave_observer_ignores_unselected_documents_and_detects_selected_dirty_buffer(self):
        native = NativeOffice.__new__(NativeOffice)
        doc = MagicMock()
        doc.Saved = False
        record = Document("test", MagicMock(), doc, "microsoft", "word", "example.docx", None)
        native.documents = {record.id: record}
        native.save = MagicMock()
        native.autosave_tick()
        native.save.assert_not_called()
        record.managed = True
        native.autosave_tick()
        native.save.assert_called_once_with(record.id)

    def test_ranges_validate_boundaries_and_reject_expressions(self):
        self.assertEqual(a1_size("$A$1:B2"), (2, 2))
        self.assertEqual(a1_size("XFD1048576"), (1, 1))
        for address in ("A0", "XFE1", "A1048577", "B2:A1", "A1:A10001", "[external.xlsx]S!A1", "A1,B2"):
            with self.subTest(address=address), self.assertRaises(OfficeError):
                a1_size(address)

    def test_color_uses_office_byte_order(self):
        self.assertEqual(rgb("#112233"), 0x332211)
        with self.assertRaises(OfficeError):
            rgb("red")

    def test_disk_conflict_refuses_save_even_if_native_buffer_claims_saved(self):
        with tempfile.TemporaryDirectory() as folder:
            path = Path(folder) / "test.docx"
            path.write_bytes(b"original")
            doc = MagicMock()
            doc.FullName = str(path)
            doc.Name = path.name
            doc.Saved = True
            doc.ReadOnly = False
            native = NativeOffice.__new__(NativeOffice)
            native.documents = {}
            record = Document("test", MagicMock(), doc, "microsoft", "word", canonical(path), file_hash(path), managed=True)
            native.documents[record.id] = record
            path.write_bytes(b"external modification")
            with self.assertRaisesRegex(OfficeError, "Disk file changed"):
                native.save(record.id)
            doc.Save.assert_not_called()
            self.assertEqual(path.read_bytes(), b"external modification")

    def test_close_cannot_discard_user_edits(self):
        native = NativeOffice.__new__(NativeOffice)
        native.documents = {}
        doc = MagicMock()
        doc.FullName = "untitled"
        doc.Name = "untitled"
        doc.Saved = False
        record = Document("test", MagicMock(), doc, "wps", "word", None, None, managed=True)
        native.documents[record.id] = record
        with self.assertRaisesRegex(OfficeError, "unsaved edits"):
            native.close(record.id, save=False)
        doc.Close.assert_not_called()

    def test_config_uses_current_python_and_valid_json(self):
        config = json.loads(client_config("json"))["mcpServers"]["microsoft-wps-live"]
        self.assertEqual(config["command"], sys.executable)
        self.assertEqual(config["args"], ["-m", "microsoft_wps_live_mcp"])

    def test_same_sta_thread_owns_engine_and_all_operations(self):
        ids = []

        class Fake:
            def __init__(self, _):
                ids.append(threading.get_ident())

            def work(self):
                ids.append(threading.get_ident())
                return {"ok": True}

        worker = OfficeWorker(factory=Fake)
        try:
            asyncio.run(worker.call("work"))
            asyncio.run(worker.call("work"))
            self.assertEqual(len(set(ids)), 1)
            self.assertNotEqual(ids[0], threading.get_ident())
        finally:
            worker.stop()
            worker.thread.join(2)

    def test_timeout_is_queryable_and_cannot_be_replayed_while_pending(self):
        gate = threading.Event()

        class Fake:
            def __init__(self, _):
                pass

            def work(self):
                gate.wait(2)
                return {"applied": True}

        worker = OfficeWorker(factory=Fake, timeout=0.01)
        try:
            result = asyncio.run(worker.call("work"))
            self.assertEqual(result["state"], "pending")
            with self.assertRaisesRegex(OfficeError, "still running"):
                asyncio.run(worker.call("work"))
            gate.set()
            worker.history[result["operation_id"]].result(timeout=2)
            completed = worker.operation_status(result["operation_id"])
            self.assertEqual(completed["state"], "completed")
            self.assertTrue(completed["result"]["applied"])
        finally:
            gate.set()
            worker.stop()
            worker.thread.join(2)


    def test_macos_blank_ooxml_generation(self):
        from microsoft_wps_live_mcp.macos import create_blank_ooxml
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            for kind in ("word", "excel", "ppt"):
                ext = {"word": ".docx", "excel": ".xlsx", "ppt": ".pptx"}[kind]
                path = root / f"blank_{kind}{ext}"
                create_blank_ooxml(path, kind)
                self.assertTrue(path.is_file())
                with zipfile.ZipFile(path) as z:
                    names = z.namelist()
                    self.assertIn("[Content_Types].xml", names)
                    self.assertIn("_rels/.rels", names)

    def test_macos_capabilities_and_initialization(self):
        from microsoft_wps_live_mcp.discovery import capabilities
        with patch("sys.platform", "darwin"):
            caps = capabilities()
            self.assertEqual(caps["platform"], "darwin")
            self.assertTrue(caps["native_supported"])
            self.assertIn("microsoft", caps["engines"])
            self.assertIn("wps", caps["engines"])

            with tempfile.TemporaryDirectory() as folder:
                native = NativeOffice(data_dir=folder)
                self.assertIsNone(native.com)
                self.assertEqual(native.documents, {})


if __name__ == "__main__":
    unittest.main()
