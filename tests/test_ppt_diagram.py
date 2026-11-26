"""Tests for PowerPoint diagram scene primitives, image export, and fluency optimizations."""
import asyncio
import json
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from mcp import Client
from microsoft_wps_live_mcp.native import NativeOffice, OfficeError
from microsoft_wps_live_mcp.ppt_diagram import _CoordMapper, _shade_hex
from microsoft_wps_live_mcp.server import build_server


def _fake_shape():
    shape = MagicMock()
    shape.TextFrame.TextRange.Characters.side_effect = lambda pos, length: MagicMock()
    return shape


class _FakeShapes:
    def __init__(self):
        self.created = []

    def AddShape(self, stype, left, top, w, h):
        s = _fake_shape()
        self.created.append(("AddShape", stype, left, top, w, h, s))
        return s

    def AddTextbox(self, orient, left, top, w, h):
        s = _fake_shape()
        self.created.append(("AddTextbox", orient, left, top, w, h, s))
        return s

    def AddConnector(self, ctype, x1, y1, x2, y2):
        s = _fake_shape()
        self.created.append(("AddConnector", ctype, x1, y1, x2, y2, s))
        return s

    def AddPicture(self, path, link, save, left, top, w, h):
        s = _fake_shape()
        s.Width = 400.0
        s.Height = 200.0
        self.created.append(("AddPicture", path, left, top, w, h, s))
        return s

    def BuildFreeform(self, editing_type, x, y):
        nodes = [(x, y)]
        ff = MagicMock()
        ff.AddNodes.side_effect = lambda *args: nodes.append(args)

        def convert():
            s = _fake_shape()
            self.created.append(("Freeform", editing_type, nodes, s))
            return s

        ff.ConvertToShape.side_effect = convert
        return ff


def _build_scene_mock():
    native = NativeOffice.__new__(NativeOffice)
    native.documents = {}
    native.live_windows = {}
    shapes = _FakeShapes()
    slide = SimpleNamespace(Shapes=shapes)
    doc = SimpleNamespace(
        Saved=True,
        PageSetup=SimpleNamespace(SlideWidth=960.0, SlideHeight=540.0),
        Slides=SimpleNamespace(Count=1, Item=lambda i: slide),
    )
    record = SimpleNamespace(
        id="doc1", doc=doc, app=MagicMock(), engine="microsoft", kind="ppt",
        path=None, disk_hash=None, managed=True, autosave_paused=False,
        autosave_state=None, recovery_copy=None, live_enabled=True,
        live_pause_ms=0, live_state={"enabled": True, "state": "ready"},
    )
    native.resolve = lambda d, k=None: record
    native._ensure_live_window = lambda r, foreground=False: None
    native._slide = lambda r, i: slide
    native._before_edit = lambda r: None
    native.save = lambda did: {"disk_saved": True}
    native.live_checkpoint = lambda r, f=None, pause_ms=None, label=None: {"enabled": True, "state": "ready"}
    native._paint = lambda r, p=0: None
    native._info = lambda r: {"document_id": r.id, "saved": True}
    return native, shapes


class PptDiagramTests(unittest.TestCase):
    def test_coord_mapper_contain_and_stretch(self):
        contain = _CoordMapper(960, 540, {"width": 1000, "height": 500, "margin": 20, "fit": "contain"})
        self.assertEqual(contain.w(1000), 920.0)
        self.assertEqual(contain.h(500), 460.0)
        self.assertEqual(contain.x(0), 20.0)
        self.assertEqual(contain.y(0), 40.0)  # vertically centered in 500pt available height

        stretch = _CoordMapper(960, 540, {"width": 970, "height": 531, "origin_x": 38, "fit": "stretch"})
        self.assertEqual(stretch.x(38), 0.0)
        self.assertEqual(stretch.x(1008), 960.0)
        self.assertEqual(stretch.y(531), 540.0)

    def test_shade_hex_bounds(self):
        self.assertEqual(_shade_hex("#808080", 1.5), "#C0C0C0")
        self.assertEqual(_shade_hex("#808080", 0.5), "#404040")

    def test_draw_scene_anchors_prism_grid_and_runs(self):
        native, shapes = _build_scene_mock()

        stages = [
            {
                "label": "Draw nodes",
                "items": [
                    {"type": "prism3d", "id": "p1", "x": 100, "y": 100, "w": 80, "h": 40,
                     "extrude": [10, -6], "fill": "#9AC77E", "border": "#4E7A42", "text": "Trans"},
                    {"type": "shear_box", "id": "s1", "cx": 300, "cy": 120, "w": 90, "h": 24,
                     "tilt_deg": -12, "fill": "#FFFFFF", "border": "#4F6D8A",
                     "runs": [{"text": "1x1"}, {"text": "Conv", "bold": True}]},
                    {"type": "grid3d", "id": "g1", "x": 450, "y": 100, "rows": 2, "cols": 2, "depth": 2},
                    {"type": "feature_stack", "id": "f1", "x": 600, "y": 100, "w": 50, "h": 50, "count": 3},
                ],
            },
            {
                "label": "Connect nodes",
                "items": [
                    {"type": "arrow", "start": {"node": "p1", "anchor": "right"},
                     "end": {"node": "s1", "anchor": "left"}, "dash": "dash", "arrow": "stealth"},
                    {"type": "curve_arrow", "start": {"node": "f1", "anchor": "center"},
                     "c1": [650, 40], "c2": [700, 40], "end": [720, 120], "color": "#3F6AA8"},
                ],
            },
        ]
        res = native.ppt_draw_scene("doc1", stages, coord_space={"width": 1000, "height": 500, "fit": "stretch"},
                                    pause_ms=0, item_pause_ms=0)
        self.assertEqual(res["state"], "completed")
        self.assertEqual(res["completed_stages"], 2)
        self.assertTrue(res["disk_saved"])
        # 1 prism(3 faces + 1 tb) + 1 shear_box(1 poly + 1 tb) + grid3d(8*3=24 faces) + feature_stack(3) + 1 connector + 1 curve
        self.assertGreaterEqual(len(shapes.created), 34)
        # Connector should attach to s1's left anchor (around x=300*0.96 - 45*0.96 = 244.8)
        conn = [c for c in shapes.created if c[0] == "AddConnector"][0]
        self.assertAlmostEqual(conn[4], 244.8, delta=0.5)

    def test_junction_aligned_chain_arc_height_and_layout_audit(self):
        native, shapes = _build_scene_mock()
        stages = [
            {
                "label": "Junction and chain",
                "items": [
                    {"type": "junction", "id": "j1", "on_line": {"start": [100, 100], "end": [500, 300], "t": 0.5}},
                    {"type": "junction", "id": "j2", "point": [500, 120]},
                    {
                        "type": "aligned_chain",
                        "start": {"node": "j1"},
                        "end": {"node": "j2"},
                        "color": "#3F6AA8",
                        "nodes": [
                            {"id": "c1", "t": 0.3, "w": 50, "h": 20, "text": "1x1Conv"},
                            {"id": "c2", "t": 0.7, "w": 50, "h": 20, "text": "Down"},
                        ],
                    },
                    {"type": "curve_arrow", "start": {"node": "j1"}, "end": {"node": "j2"}, "arc_height": -45},
                ],
            },
            {
                "label": "Intentional collisions for audit",
                "items": [
                    {"type": "box", "id": "b1", "x": 20, "y": 20, "w": 80, "h": 40, "text": "B1"},
                    {"type": "box", "id": "b2", "x": 50, "y": 30, "w": 80, "h": 40, "text": "B2"},
                    {"type": "grid3d", "id": "g_bad", "x": 180, "y": 60, "rows": 2, "cols": 2, "depth": 1,
                     "cell_w": 10, "cell_h": 10, "extrude": [-20, -15]},
                    {"type": "textbox", "x": 165, "y": 48, "w": 40, "h": 18, "text": "Collide"},
                    {"type": "textbox", "x": -18, "y": 10, "w": 30, "h": 16, "text": "EdgeOut"},
                ],
            },
        ]
        res = native.ppt_draw_scene("doc1", stages, coord_space={"width": 960, "height": 540, "fit": "stretch"},
                                    pause_ms=0, item_pause_ms=0)
        self.assertEqual(res["state"], "completed")
        # aligned_chain produced 3 connectors between j1=(300,200) and j2=(500,120)
        conns = [c for c in shapes.created if c[0] == "AddConnector"]
        self.assertEqual(len(conns), 3)
        self.assertAlmostEqual(conns[0][2], 300.0, delta=0.2)
        self.assertAlmostEqual(conns[0][3], 200.0, delta=0.2)
        self.assertAlmostEqual(conns[-1][4], 500.0, delta=0.2)
        self.assertAlmostEqual(conns[-1][5], 120.0, delta=0.2)
        # Verify collinear slope dy/dx == -80/200 == -0.4 on every segment
        for _, _, x1, y1, x2, y2, _ in conns:
            self.assertAlmostEqual((y2 - y1) / (x2 - x1), -0.4, delta=0.02)

        warning_types = {w["type"] for w in res["layout_warnings"]}
        self.assertIn("box_overlap", warning_types)
        self.assertIn("text_on_object", warning_types)
        self.assertIn("out_of_bounds", warning_types)

    def test_api_batch_batches_checkpoints_per_created_shape(self):
        native = NativeOffice.__new__(NativeOffice)
        shapes = _FakeShapes()
        slide = SimpleNamespace(Shapes=shapes)
        doc = SimpleNamespace(Saved=True, Slides=SimpleNamespace(Item=lambda i: slide))
        record = SimpleNamespace(
            id="doc1", doc=doc, app=MagicMock(), engine="microsoft", kind="ppt",
            path=None, disk_hash=None, managed=True, autosave_paused=False,
            autosave_state=None, recovery_copy=None, live_enabled=True,
            live_pause_ms=0, live_state={"enabled": True, "state": "ready"},
        )
        checkpoints = []
        native.resolve = lambda d, k=None: record
        native._before_edit = lambda r: None
        native._result = lambda r, op, save, backup: {"disk_saved": True, "operation": op}
        native._focus_for_object = lambda r, o: {"slide": 1}
        native.live_checkpoint = lambda r, f=None, pause_ms=None, label=None: checkpoints.append(f) or {"state": "ready"}
        native._info = lambda r: {"document_id": r.id}

        steps = [
            {"action": "call", "target": "document.Slides", "member": "Item", "args": [1], "read_only": True, "as": "s"},
            {"action": "call", "target": "s.Shapes", "member": "AddShape", "args": [1, 10, 10, 50, 20], "as": "b1"},
            {"action": "set", "target": "b1.Fill", "member": "Visible", "value": -1},
            {"action": "set", "target": "b1.Line", "member": "Weight", "value": 1.5},
            {"action": "set", "target": "b1.TextFrame.TextRange", "member": "Text", "value": "A"},
            {"action": "call", "target": "s.Shapes", "member": "AddShape", "args": [1, 80, 10, 50, 20], "as": "b2"},
            {"action": "set", "target": "b2.TextFrame.TextRange", "member": "Text", "value": "B"},
        ]
        res = native.api_batch("doc1", steps, save=True)
        self.assertEqual(res["state"], "completed")
        # 2 shapes -> only 2 shape checkpoints (not 6!)
        self.assertEqual(len(checkpoints), 2)

    def test_ppt_export_image_validates_and_exports(self):
        with tempfile.TemporaryDirectory() as folder:
            out = Path(folder) / "slide.png"
            native = NativeOffice.__new__(NativeOffice)
            page = MagicMock()

            def fake_export(dst, fmt, w, h):
                Path(dst).write_bytes(b"\x89PNG\r\n\x1a\nfake")

            page.Export.side_effect = fake_export
            record = SimpleNamespace(id="d", kind="ppt", doc=SimpleNamespace(Slides=SimpleNamespace(Count=1)))
            native.resolve = lambda d, k=None: record
            native._slide = lambda r, i: page
            native._info = lambda r: {"document_id": r.id}
            res = native.ppt_export_image("d", str(out), slide=1, width=1920, height=1080)
            self.assertTrue(res["exported"])
            self.assertEqual(res["format"], "PNG")
            with self.assertRaises(OfficeError):
                native.ppt_export_image("d", str(out))  # already exists

    def test_new_mcp_tools_registered_without_save_param(self):
        class StubWorker:
            async def call(self, method, *args, **kwargs):
                return {"method": method}

        async def check():
            async with Client(build_server(StubWorker())) as client:
                tools = await client.list_tools()
                names = {t.name for t in tools.tools}
                self.assertIn("office_ppt_draw_scene", names)
                self.assertIn("office_ppt_export_image", names)
                for t in tools.tools:
                    self.assertNotIn("save", t.input_schema.get("properties", {}), t.name)

        asyncio.run(check())

    def test_macos_ooxml_proxy_draw_scene_multi_stage_no_duplicate_shapes(self):
        import xml.etree.ElementTree as ET
        import zipfile
        from unittest.mock import patch
        from microsoft_wps_live_mcp.macos import create_blank_ooxml, MacDocProxy

        with tempfile.TemporaryDirectory() as folder:
            pptx = Path(folder) / "arch.pptx"
            create_blank_ooxml(pptx, "ppt")
            native = NativeOffice.__new__(NativeOffice)
            native.com = None
            native.data_dir = Path(folder) / "data"
            native.documents = {}
            native.apps = []
            native.live_windows = {}

            doc = MacDocProxy(pptx, "ppt", "microsoft")
            app = SimpleNamespace(macos=True, app_name="Microsoft PowerPoint", Visible=True, ScreenUpdating=True, Activate=lambda: True)
            with patch("sys.platform", "darwin"), \
                 patch("microsoft_wps_live_mcp.macos.activate_app", return_value=True), \
                 patch("microsoft_wps_live_mcp.macos.reload_ppt_in_app", return_value=True):
                record = native._remember(app, doc, "microsoft", "ppt", opened=True)
                record.live_enabled = True
                stages = [
                    {"label": "Stage 1", "items": [{"type": "box", "id": "b1", "x": 100, "y": 100, "w": 120, "h": 40, "text": "Input"}]},
                    {"label": "Stage 2", "items": [
                        {"type": "box", "id": "b2", "x": 300, "y": 100, "w": 120, "h": 40, "text": "Output"},
                        {"type": "arrow", "start": {"node": "b1", "anchor": "right"}, "end": {"node": "b2", "anchor": "left"}},
                    ]},
                ]
                res = native.ppt_draw_scene(record.id, stages, pause_ms=0, item_pause_ms=0)
                self.assertEqual(res["state"], "completed")
                self.assertEqual(res["total_items_drawn"], 3)
                self.assertTrue(res["disk_saved"])

            with zipfile.ZipFile(pptx) as z:
                sld = ET.fromstring(z.read("ppt/slides/slide1.xml"))
                sp_tree = sld.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}spTree")
                shapes_els = [c for c in list(sp_tree) if c.tag.endswith(("}sp", "}cxnSp", "}pic"))]
                self.assertEqual(len(shapes_els), 3)


if __name__ == "__main__":
    unittest.main()
