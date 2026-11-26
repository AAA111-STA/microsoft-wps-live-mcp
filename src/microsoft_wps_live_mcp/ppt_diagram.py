"""High-precision, smooth PowerPoint diagramming primitives for 1:1 reproduction and generative design."""
from __future__ import annotations

import math
from pathlib import Path
from typing import Any


def _fail(message: str):
    from .native import OfficeError
    return OfficeError(message)


def _rgb(value: str) -> int:
    from .native import rgb
    return rgb(value)


def _shade_hex(hex_str: str, factor: float) -> str:
    """Scale RGB brightness by factor (>1 lightens, <1 darkens) and return #RRGGBB."""
    raw = hex_str.lstrip("#")
    r, g, b = int(raw[0:2], 16), int(raw[2:4], 16), int(raw[4:6], 16)
    if factor >= 1.0:
        t = min(1.0, factor - 1.0)
        nr = int(round(r + (255 - r) * t))
        ng = int(round(g + (255 - g) * t))
        nb = int(round(b + (255 - b) * t))
    else:
        nr = int(round(max(0, min(255, r * factor))))
        ng = int(round(max(0, min(255, g * factor))))
        nb = int(round(max(0, min(255, b * factor))))
    return f"#{nr:02X}{ng:02X}{nb:02X}"


class _CoordMapper:
    """Maps input coordinates (reference pixels or logical space) onto PowerPoint points."""

    def __init__(self, slide_w: float, slide_h: float, coord_space: dict[str, Any] | None = None):
        if not coord_space:
            self.ox = 0.0
            self.oy = 0.0
            self.dx = 0.0
            self.dy = 0.0
            self.sx = 1.0
            self.sy = 1.0
            return
        cw = float(coord_space.get("width", 0))
        ch = float(coord_space.get("height", 0))
        if not (math.isfinite(cw) and math.isfinite(ch) and cw > 0 and ch > 0):
            raise _fail("coord_space.width and coord_space.height must be positive numbers.")
        margin = float(coord_space.get("margin", 0.0))
        if not math.isfinite(margin) or margin < 0 or margin * 2 >= min(slide_w, slide_h):
            raise _fail("coord_space.margin must be nonnegative and smaller than half the slide size.")
        self.ox = float(coord_space.get("origin_x", 0.0))
        self.oy = float(coord_space.get("origin_y", 0.0))
        avail_w = slide_w - 2.0 * margin
        avail_h = slide_h - 2.0 * margin
        fit = coord_space.get("fit", "contain")
        if fit == "stretch":
            self.sx = avail_w / cw
            self.sy = avail_h / ch
            self.dx = margin
            self.dy = margin
        elif fit == "contain":
            scale = min(avail_w / cw, avail_h / ch)
            self.sx = scale
            self.sy = scale
            self.dx = margin + (avail_w - cw * scale) / 2.0
            self.dy = margin + (avail_h - ch * scale) / 2.0
        else:
            raise _fail("coord_space.fit must be 'contain' or 'stretch'.")

    def x(self, val: float) -> float:
        return round(self.dx + (float(val) - self.ox) * self.sx, 2)

    def y(self, val: float) -> float:
        return round(self.dy + (float(val) - self.oy) * self.sy, 2)

    def w(self, val: float) -> float:
        return round(float(val) * self.sx, 2)

    def h(self, val: float) -> float:
        return round(float(val) * self.sy, 2)

    def pt(self, p: list[float] | tuple[float, float]) -> tuple[float, float]:
        return (self.x(p[0]), self.y(p[1]))


def _rect_from_item(item: dict[str, Any], mapper: _CoordMapper,
                    nodes: dict[str, dict[str, Any]] | None = None) -> tuple[float, float, float, float]:
    w = mapper.w(item["w"])
    h = mapper.h(item["h"])
    if w <= 0 or h <= 0:
        raise _fail("Item width w and height h must be positive.")
    if "center" in item and nodes is not None:
        cx_pt, cy_pt = _resolve_point(item["center"], mapper, nodes)
        left = cx_pt - w / 2.0
        top = cy_pt - h / 2.0
    elif "cx" in item and "cy" in item:
        left = mapper.x(item["cx"]) - w / 2.0
        top = mapper.y(item["cy"]) - h / 2.0
    elif "x" in item and "y" in item:
        left = mapper.x(item["x"])
        top = mapper.y(item["y"])
    else:
        raise _fail("Item requires (x, y, w, h), (cx, cy, w, h), or (center, w, h).")
    return left, top, w, h


def _register_anchors(nodes: dict[str, dict[str, Any]], item: dict[str, Any],
                      left: float, top: float, w: float, h: float,
                      extra: dict[str, tuple[float, float]] | None = None) -> None:
    node_id = item.get("id")
    if not node_id:
        return
    anchors = {
        "center": (left + w / 2.0, top + h / 2.0),
        "left": (left, top + h / 2.0),
        "right": (left + w, top + h / 2.0),
        "top": (left + w / 2.0, top),
        "bottom": (left + w / 2.0, top + h),
        "top_left": (left, top),
        "top_right": (left + w, top),
        "bottom_left": (left, top + h),
        "bottom_right": (left + w, top + h),
    }
    if extra:
        anchors.update(extra)
    nodes[str(node_id)] = {"left": left, "top": top, "width": w, "height": h, "anchors": anchors}


def _resolve_point(spec: Any, mapper: _CoordMapper, nodes: dict[str, dict[str, Any]]) -> tuple[float, float]:
    if isinstance(spec, (list, tuple)) and len(spec) == 2:
        return mapper.pt(spec)
    if isinstance(spec, dict) and "node" in spec:
        node_id = str(spec["node"])
        if node_id not in nodes:
            raise _fail(f"Unknown node reference '{node_id}'. Define nodes before referencing them.")
        anchor = str(spec.get("anchor", "center"))
        anchors = nodes[node_id]["anchors"]
        if anchor not in anchors:
            raise _fail(f"Unknown anchor '{anchor}' on node '{node_id}'.")
        bx, by = anchors[anchor]
        off = spec.get("offset")
        if off is not None:
            bx += mapper.w(off[0])
            by += mapper.h(off[1])
        return (round(bx, 2), round(by, 2))
    if isinstance(spec, dict) and "on_line" in spec:
        line_spec = spec["on_line"]
        p0 = _resolve_point(line_spec["start"], mapper, nodes)
        p1 = _resolve_point(line_spec["end"], mapper, nodes)
        t = float(line_spec.get("t", spec.get("t", 0.5)))
        return (round(p0[0] + (p1[0] - p0[0]) * t, 2), round(p0[1] + (p1[1] - p0[1]) * t, 2))
    raise _fail("Point must be [x, y], {'node': '<id>', ...}, or {'on_line': {'start': ..., 'end': ..., 't': ...}}.")


_DASH_STYLES = {
    "solid": 1,
    "square_dot": 2,
    "dot": 3,
    "dash": 4,
    "dash_dot": 5,
    "long_dash": 7,
}

_ARROW_STYLES = {
    "none": 1,
    "triangle": 2,
    "open": 3,
    "stealth": 4,
    "diamond": 5,
    "oval": 6,
}


def _dash_value(dash: int | str | None) -> int | None:
    if dash is None:
        return None
    if isinstance(dash, str):
        key = dash.strip().casefold()
        if key not in _DASH_STYLES:
            raise _fail(f"Unknown dash style '{dash}'.")
        return _DASH_STYLES[key]
    return int(dash)


def _arrow_value(spec: bool | str | int | None) -> int:
    if spec is False or spec is None:
        return 1
    if spec is True:
        return 2
    if isinstance(spec, str):
        key = spec.strip().casefold()
        if key not in _ARROW_STYLES:
            raise _fail(f"Unknown arrow style '{spec}'.")
        return _ARROW_STYLES[key]
    return int(spec)


def _apply_fill(shape, fill: str | None, transparent: bool = False, transparency: float | None = None) -> None:
    if transparent or fill is None:
        shape.Fill.Visible = 0
        return
    shape.Fill.Visible = -1
    shape.Fill.Solid()
    shape.Fill.ForeColor.RGB = _rgb(fill)
    if transparency is not None:
        shape.Fill.Transparency = max(0.0, min(1.0, float(transparency)))


def _apply_line(shape, border: str | None, weight: float = 1.0, dash: int | str | None = None) -> None:
    if border is None or weight <= 0:
        shape.Line.Visible = 0
        return
    shape.Line.Visible = -1
    shape.Line.ForeColor.RGB = _rgb(border)
    shape.Line.Weight = float(weight)
    dash_val = _dash_value(dash)
    if dash_val is not None:
        shape.Line.DashStyle = dash_val


def _apply_shadow(shape, shadow: bool | dict[str, Any] | None) -> None:
    if not shadow:
        return
    sh = shape.Shadow
    sh.Visible = -1
    if isinstance(shadow, dict):
        if "blur" in shadow:
            sh.Blur = float(shadow["blur"])
        if "transparency" in shadow:
            sh.Transparency = float(shadow["transparency"])
        if "offset_x" in shadow:
            sh.OffsetX = float(shadow["offset_x"])
        if "offset_y" in shadow:
            sh.OffsetY = float(shadow["offset_y"])


def _apply_text(shape, text: str = "", *, font_name: str = "Times New Roman", font_size: float = 10.0,
                bold: bool = False, italic: bool = False, text_color: str = "#111111",
                align: int = 2, valign: int = 3, margins: list[float] | tuple[float, ...] = (1.0, 1.0, 0.5, 0.5),
                word_wrap: bool = True, runs: list[dict[str, Any]] | None = None,
                line_spacing: float | None = None) -> None:
    tf = shape.TextFrame
    tf.MarginLeft = float(margins[0])
    tf.MarginRight = float(margins[1])
    tf.MarginTop = float(margins[2])
    tf.MarginBottom = float(margins[3])
    tf.WordWrap = -1 if word_wrap else 0
    tf.VerticalAnchor = int(valign)
    tr = tf.TextRange
    if runs:
        tr.Text = "".join(str(r.get("text", "")) for r in runs)
        pos = 1
        for run_spec in runs:
            chunk = str(run_spec.get("text", ""))
            if not chunk:
                continue
            sub = tr.Characters(pos, len(chunk))
            f = sub.Font
            f.Name = str(run_spec.get("font_name", font_name))
            f.Size = float(run_spec.get("font_size", font_size))
            f.Bold = -1 if run_spec.get("bold", bold) else 0
            f.Italic = -1 if run_spec.get("italic", italic) else 0
            f.Color.RGB = _rgb(str(run_spec.get("color", text_color)))
            if run_spec.get("subscript"):
                f.Subscript = -1
            if run_spec.get("superscript"):
                f.Superscript = -1
            pos += len(chunk)
    else:
        tr.Text = text
        font = tr.Font
        font.Name = font_name
        font.Size = float(font_size)
        font.Bold = -1 if bold else 0
        font.Italic = -1 if italic else 0
        font.Color.RGB = _rgb(text_color)
    tr.ParagraphFormat.Alignment = int(align)
    if line_spacing is not None:
        tr.ParagraphFormat.SpaceWithin = float(line_spacing)


def _add_polygon(slide, points: list[tuple[float, float]], fill: str | None,
                 border: str | None = None, border_weight: float = 1.0, dash: int | None = None):
    if len(points) < 3:
        raise _fail("Polygon requires at least 3 points.")
    ff = slide.Shapes.BuildFreeform(0, float(points[0][0]), float(points[0][1]))
    for x, y in points[1:]:
        ff.AddNodes(0, 0, float(x), float(y))
    ff.AddNodes(0, 0, float(points[0][0]), float(points[0][1]))
    shape = ff.ConvertToShape()
    _apply_fill(shape, fill, transparent=(fill is None))
    _apply_line(shape, border, border_weight, dash)
    return shape


def _draw_prism(slide, left: float, top: float, w: float, h: float, ex: float, ey: float,
                skew_x: float, skew_y: float, fill: str, top_fill: str | None,
                side_fill: str | None, border: str | None, border_weight: float):
    """Draw a true 3-face 3D oblique prism using native vector freeform polygons."""
    # Front face vertices clockwise from top-left
    f0 = (left, top)
    f1 = (left + w, top + skew_y)
    f2 = (left + w + skew_x, top + h + skew_y)
    f3 = (left + skew_x, top + h)
    b0 = (f0[0] + ex, f0[1] + ey)
    b1 = (f1[0] + ex, f1[1] + ey)
    b2 = (f2[0] + ex, f2[1] + ey)
    b3 = (f3[0] + ex, f3[1] + ey)

    top_c = top_fill or _shade_hex(fill, 1.18)
    side_c = side_fill or _shade_hex(fill, 0.84)
    shapes = []
    # Vertical extrusion face (top if ey < 0 else bottom)
    if abs(ey) > 0.1:
        if ey < 0:
            shapes.append(_add_polygon(slide, [f0, b0, b1, f1], top_c, border, border_weight))
        else:
            shapes.append(_add_polygon(slide, [f3, f2, b2, b3], side_c, border, border_weight))
    # Horizontal extrusion face (right if ex > 0 else left)
    if abs(ex) > 0.1:
        if ex > 0:
            shapes.append(_add_polygon(slide, [f1, b1, b2, f2], side_c, border, border_weight))
        else:
            shapes.append(_add_polygon(slide, [b0, f0, f3, b3], side_c, border, border_weight))
    front = _add_polygon(slide, [f0, f1, f2, f3], fill, border, border_weight)
    shapes.append(front)
    return front, shapes


def _boxes_overlap(a: tuple[float, float, float, float], b: tuple[float, float, float, float], pad: float = 1.0) -> bool:
    ax1, ay1, ax2, ay2 = a[0] + pad, a[1] + pad, a[0] + a[2] - pad, a[1] + a[3] - pad
    bx1, by1, bx2, by2 = b[0] + pad, b[1] + pad, b[0] + b[2] - pad, b[1] + b[3] - pad
    return ax1 < bx2 and ax2 > bx1 and ay1 < by2 and ay2 > by1


def _audit_layout(audit_items: list[dict[str, Any]], slide_w: float, slide_h: float) -> list[dict[str, Any]]:
    """Lightweight in-memory geometric self-check returned after drawing."""
    warnings: list[dict[str, Any]] = []
    for entry in audit_items:
        if entry.get("allow_overflow"):
            continue
        x, y, w, h = entry["rect"]
        if x < -2.0 or y < -2.0 or x + w > slide_w + 2.0 or y + h > slide_h + 2.0:
            warnings.append({"type": "out_of_bounds", "label": entry["label"],
                             "rect": [round(v, 1) for v in (x, y, w, h)]})

    texts = [e for e in audit_items if e["kind"] == "textbox" and not e.get("allow_overlap")]
    for i in range(len(texts)):
        for j in range(i + 1, len(texts)):
            a, b = texts[i], texts[j]
            if a["slide"] == b["slide"] and _boxes_overlap(a["rect"], b["rect"], pad=1.5):
                warnings.append({"type": "text_overlap", "items": [a["label"], b["label"]]})

    solids = [e for e in audit_items
              if e["kind"] in ("box", "shear_box", "prism3d", "grid3d", "feature_stack", "image")
              and not e.get("transparent") and not e.get("allow_overlap")]
    for i in range(len(solids)):
        for j in range(i + 1, len(solids)):
            a, b = solids[i], solids[j]
            # Multi-strip 3D token stacks intentionally layer prism3d bars tightly
            if a["kind"] == "prism3d" and b["kind"] == "prism3d":
                continue
            if a["slide"] == b["slide"] and _boxes_overlap(a["rect"], b["rect"], pad=2.0):
                warnings.append({"type": "box_overlap", "items": [a["label"], b["label"]]})

    # Catch standalone textboxes colliding with 3D tensors/stacks/images
    media_objs = [e for e in solids if e["kind"] in ("prism3d", "grid3d", "feature_stack", "image")]
    for tb in texts:
        for obj in media_objs:
            if tb["slide"] == obj["slide"] and _boxes_overlap(tb["rect"], obj["rect"], pad=2.0):
                warnings.append({"type": "text_on_object", "items": [tb["label"], obj["label"]]})
    return warnings[:20]


def _bezier_from_arc_height(p0: tuple[float, float], p3: tuple[float, float],
                            arc_height: float, arc_skew: float = 0.0) -> tuple[tuple[float, float], tuple[float, float]]:
    dx, dy = p3[0] - p0[0], p3[1] - p0[1]
    length = math.hypot(dx, dy) or 1.0
    nx, ny = -dy / length, dx / length
    t1 = max(0.1, min(0.45, 0.30 + arc_skew))
    t2 = max(0.55, min(0.9, 0.70 + arc_skew))
    c1 = (round(p0[0] + dx * t1 + nx * arc_height, 2), round(p0[1] + dy * t1 + ny * arc_height, 2))
    c2 = (round(p0[0] + dx * t2 + nx * arc_height, 2), round(p0[1] + dy * t2 + ny * arc_height, 2))
    return c1, c2


class PptDiagramMixin:
    """High-level scene drawing tools mixed into NativeOffice."""

    def ppt_draw_scene(self, document: str, stages: list[dict[str, Any]],
                       coord_space: dict[str, Any] | None = None,
                       slide_size: list[float] | None = None,
                       pause_ms: int = 320, item_pause_ms: int = 32,
                       _on_progress=None, _cancel_event=None):
        if not isinstance(stages, list) or not 1 <= len(stages) <= 60:
            raise _fail("stages must contain 1 to 60 named drawing stages.")
        if isinstance(pause_ms, bool) or not isinstance(pause_ms, int) or not 0 <= pause_ms <= 2000:
            raise _fail("pause_ms must be an integer between 0 and 2000.")
        if isinstance(item_pause_ms, bool) or not isinstance(item_pause_ms, int) or not 0 <= item_pause_ms <= 1000:
            raise _fail("item_pause_ms must be an integer between 0 and 1000.")
        for st in stages:
            if (not isinstance(st, dict) or not isinstance(st.get("label"), str) or not st["label"].strip()
                    or not isinstance(st.get("items"), list) or not st["items"]):
                raise _fail("Each stage requires a non-empty 'label' and non-empty 'items' list.")

        r = self.resolve(document, "ppt")
        r.live_enabled = True
        self._ensure_live_window(r, foreground=True)

        if slide_size is not None:
            if not isinstance(slide_size, (list, tuple)) or len(slide_size) != 2:
                raise _fail("slide_size must be [width, height] in points.")
            sw, sh = float(slide_size[0]), float(slide_size[1])
            if not (math.isfinite(sw) and math.isfinite(sh) and sw > 0 and sh > 0):
                raise _fail("slide_size values must be positive finite numbers.")
            r.doc.PageSetup.SlideWidth = sw
            r.doc.PageSetup.SlideHeight = sh
        else:
            sw = float(r.doc.PageSetup.SlideWidth)
            sh = float(r.doc.PageSetup.SlideHeight)

        mapper = _CoordMapper(sw, sh, coord_space)
        nodes: dict[str, dict[str, Any]] = {}
        audit_items: list[dict[str, Any]] = []
        completed: list[dict[str, Any]] = []
        total_items = 0

        for index, stage in enumerate(stages, 1):
            if _cancel_event is not None and _cancel_event.is_set():
                return {**self._info(r), "state": "cancelled", "completed_stages": len(completed),
                        "stages": completed, "disk_saved": bool(r.doc.Saved)}
            slide_idx = int(stage.get("slide", 1))
            page = self._slide(r, slide_idx)
            backup = self._before_edit(r)
            created_in_stage = 0
            try:
                for item in stage["items"]:
                    if not isinstance(item, dict) or "type" not in item:
                        raise _fail("Each item must be an object with a 'type' field.")
                    created = self._draw_item(page, item, mapper, nodes, audit_items, slide_idx)
                    created_in_stage += created
                    if created > 0:
                        self.live_checkpoint(r, {"slide": slide_idx}, pause_ms=item_pause_ms, label=stage["label"])
                total_items += created_in_stage
                saved = self.save(r.id)
            except Exception as exc:
                r.autosave_paused = True
                r.autosave_state = {"state": "needs_inspection", "message": str(exc)}
                return {**self._info(r), "state": "partial_error", "failed_stage": index,
                        "completed_stages": len(completed), "stages": completed,
                        "error": str(exc), "disk_saved": False, "backup": backup,
                        "inspect_before_retry": True}

            focus = stage.get("focus") or {"slide": slide_idx}
            view = self.live_checkpoint(r, focus, pause_ms=0, label=stage["label"])
            event = {"stage": index, "total": len(stages), "label": stage["label"],
                     "items_drawn": created_in_stage, "state": "completed",
                     "disk_saved": saved.get("disk_saved", False), "live_view": view}
            completed.append(event)
            if _on_progress:
                _on_progress(event)
            if view.get("state") == "unavailable":
                return {**self._info(r), "state": "display_error", "completed_stages": index,
                        "stages": completed, "disk_saved": bool(r.doc.Saved),
                        "message": "Stage saved; live display became unavailable."}
            self._paint(r, pause_ms)

        layout_warnings = _audit_layout(audit_items, sw, sh)
        return {**self._info(r), "state": "completed", "disk_saved": bool(r.doc.Saved),
                "completed_stages": len(completed), "total_items_drawn": total_items,
                "slide_size_points": {"width": sw, "height": sh},
                "layout_warnings": layout_warnings, "stages": completed}

    def _draw_item(self, slide, item: dict[str, Any], mapper: _CoordMapper,
                   nodes: dict[str, dict[str, Any]],
                   audit_items: list[dict[str, Any]] | None = None,
                   slide_idx: int = 1) -> int:
        kind = str(item["type"]).casefold()

        def record_audit(rect: tuple[float, float, float, float], label_fallback: str = ""):
            if audit_items is not None:
                rx, ry, rw, rh = rect
                rot = float(item.get("rotation") or 0.0)
                if rot:
                    rad = math.radians(rot)
                    ca, sa = abs(math.cos(rad)), abs(math.sin(rad))
                    cx_r, cy_r = rx + rw / 2.0, ry + rh / 2.0
                    nrw = rw * ca + rh * sa
                    nrh = rw * sa + rh * ca
                    rx, ry, rw, rh = cx_r - nrw / 2.0, cy_r - nrh / 2.0, nrw, nrh
                audit_items.append({
                    "kind": kind,
                    "slide": slide_idx,
                    "rect": (rx, ry, rw, rh),
                    "transparent": bool(item.get("transparent", False)),
                    "allow_overlap": bool(item.get("allow_overlap", False)),
                    "allow_overflow": bool(item.get("allow_overflow", False)),
                    "label": str(item.get("id") or item.get("text") or label_fallback or kind),
                })

        if kind == "junction":
            pt = _resolve_point(item.get("point", item), mapper, nodes)
            node_id = item.get("id")
            if not node_id:
                raise _fail("junction item requires an 'id' field.")
            nodes[str(node_id)] = {
                "left": pt[0], "top": pt[1], "width": 0.0, "height": 0.0,
                "anchors": {k: pt for k in ("center", "left", "right", "top", "bottom")},
            }
            return 0

        if kind == "aligned_chain":
            p_start = _resolve_point(item["start"], mapper, nodes)
            p_end = _resolve_point(item["end"], mapper, nodes)
            chain_nodes = item.get("nodes", [])
            if not isinstance(chain_nodes, list) or not chain_nodes:
                raise _fail("aligned_chain requires a non-empty 'nodes' list.")
            dx, dy = p_end[0] - p_start[0], p_end[1] - p_start[1]
            auto_tilt = round(math.degrees(math.atan2(dy, dx)), 2)
            # Keep text upright (-90..90) when drawing right-to-left
            if auto_tilt > 90:
                auto_tilt -= 180
            elif auto_tilt < -90:
                auto_tilt += 180
            tilt_deg = auto_tilt if item.get("tilt_deg", "auto") == "auto" else float(item["tilt_deg"])
            shape_kind = str(item.get("shape", "shear_box")).casefold()
            if shape_kind not in ("box", "shear_box"):
                raise _fail("aligned_chain.shape must be 'shear_box' or 'box'.")
            seg_arrows = bool(item.get("segment_arrows", True))
            color = str(item.get("color", "#262626"))
            weight = float(item.get("weight", 1.6))
            first_weight = float(item.get("first_weight", weight))
            dash_val = item.get("dash")
            arrow_style = item.get("arrow", True)

            count_shapes = 0
            n = len(chain_nodes)
            created_ids: list[str] = []
            # If segment_arrows=False, draw the underlying trunk connector first so boxes sit cleanly on top
            if not seg_arrows:
                count_shapes += self._draw_item(slide, {
                    "type": "arrow", "start": [0, 0], "end": [0, 0],  # unused when _raw_pts given below
                    "_raw_start": p_start, "_raw_end": p_end,
                    "color": color, "weight": weight, "arrow": arrow_style, "dash": dash_val,
                }, mapper, nodes, audit_items, slide_idx)

            for idx, node_spec in enumerate(chain_nodes):
                t = float(node_spec.get("t", (idx + 1) / (n + 1)))
                cx_pt = p_start[0] + dx * t
                cy_pt = p_start[1] + dy * t
                nid = str(node_spec.get("id") or f"__chain_{id(item)}_{idx}")
                created_ids.append(nid)
                w_pt = mapper.w(node_spec.get("w", item.get("node_w", 64)))
                h_pt = mapper.h(node_spec.get("h", item.get("node_h", 21)))
                sub = {
                    "type": shape_kind, "id": nid,
                    "cx": 0, "cy": 0, "w": 1, "h": 1,
                    "_raw_rect": (cx_pt - w_pt / 2.0, cy_pt - h_pt / 2.0, w_pt, h_pt),
                    "text": node_spec.get("text", ""),
                    "runs": node_spec.get("runs"),
                    "fill": node_spec.get("fill", item.get("fill", "#FFFFFF")),
                    "border": node_spec.get("border", item.get("border", color)),
                    "border_weight": float(node_spec.get("border_weight", item.get("border_weight", 1.1))),
                    "font_name": node_spec.get("font_name", item.get("font_name", "Times New Roman")),
                    "font_size": float(node_spec.get("font_size", item.get("font_size", 9.5))),
                    "bold": bool(node_spec.get("bold", item.get("bold", True))),
                    "text_color": node_spec.get("text_color", item.get("text_color", "#111111")),
                    "tilt_deg": tilt_deg,
                    "rotation": tilt_deg,
                    "skew_x": float(node_spec.get("skew_x", item.get("skew_x", 0.0))),
                    "rounded": bool(node_spec.get("rounded", item.get("rounded", False))),
                    "corner_radius": node_spec.get("corner_radius", item.get("corner_radius")),
                }
                count_shapes += self._draw_item(slide, sub, mapper, nodes, audit_items, slide_idx)

            if seg_arrows:
                entry_anchor = "left" if dx >= 0 else "right"
                exit_anchor = "right" if dx >= 0 else "left"
                # start -> first node
                count_shapes += self._draw_item(slide, {
                    "type": "arrow", "_raw_start": p_start,
                    "end": {"node": created_ids[0], "anchor": entry_anchor},
                    "color": color, "weight": first_weight, "arrow": arrow_style, "dash": dash_val,
                }, mapper, nodes, audit_items, slide_idx)
                for idx in range(len(created_ids) - 1):
                    count_shapes += self._draw_item(slide, {
                        "type": "arrow",
                        "start": {"node": created_ids[idx], "anchor": exit_anchor},
                        "end": {"node": created_ids[idx + 1], "anchor": entry_anchor},
                        "color": color, "weight": weight, "arrow": arrow_style, "dash": dash_val,
                    }, mapper, nodes, audit_items, slide_idx)
                count_shapes += self._draw_item(slide, {
                    "type": "arrow",
                    "start": {"node": created_ids[-1], "anchor": exit_anchor},
                    "_raw_end": p_end,
                    "color": color, "weight": weight, "arrow": arrow_style, "dash": dash_val,
                }, mapper, nodes, audit_items, slide_idx)
            return count_shapes

        if kind == "polygon":
            pts = [mapper.pt(p) for p in item["points"]]
            _add_polygon(slide, pts, item.get("fill"), item.get("border"),
                         float(item.get("border_weight", 1.0)), item.get("dash"))
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            _register_anchors(nodes, item, min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys))
            return 1

        if kind == "box":
            left, top, w, h = item["_raw_rect"] if "_raw_rect" in item else _rect_from_item(item, mapper, nodes)
            stype = item.get("shape_type")
            if stype is None:
                stype = 5 if item.get("rounded", False) else 1
            shape = slide.Shapes.AddShape(int(stype), left, top, w, h)
            if item.get("corner_radius") is not None and int(stype) == 5:
                try:
                    shape.Adjustments.Item(1, float(item["corner_radius"]))
                except Exception:
                    try:
                        shape.Adjustments[1] = float(item["corner_radius"])
                    except Exception:
                        pass
            _apply_fill(shape, item.get("fill", "#FFFFFF"), bool(item.get("transparent", False)),
                        item.get("fill_transparency"))
            _apply_line(shape, item.get("border", "#4F6D8A"),
                        float(item.get("border_weight", 1.1)), item.get("dash"))
            _apply_shadow(shape, item.get("shadow"))
            if item.get("rotation"):
                shape.Rotation = float(item["rotation"])
            if item.get("text") or item.get("runs"):
                _apply_text(
                    shape, str(item.get("text", "")),
                    font_name=str(item.get("font_name", "Times New Roman")),
                    font_size=float(item.get("font_size", 9.5)),
                    bold=bool(item.get("bold", True)),
                    italic=bool(item.get("italic", False)),
                    text_color=str(item.get("text_color", "#111111")),
                    align=int(item.get("align", 2)),
                    valign=int(item.get("valign", 3)),
                    margins=item.get("margins", (1.0, 1.0, 0.5, 0.5)),
                    runs=item.get("runs"),
                    line_spacing=item.get("line_spacing"),
                )
            _register_anchors(nodes, item, left, top, w, h)
            record_audit((left, top, w, h))
            return 1

        if kind == "shear_box":
            left, top, w, h = item["_raw_rect"] if "_raw_rect" in item else _rect_from_item(item, mapper, nodes)
            tilt_deg = float(item.get("tilt_deg", -12.0))
            skew_x = mapper.w(item.get("skew_x", 0.0))
            dy_tilt = math.tan(math.radians(tilt_deg)) * w
            # Top-left -> top-right -> bottom-right -> bottom-left shifted around box center
            cx, cy = left + w / 2.0, top + h / 2.0
            hw, hh = w / 2.0, h / 2.0
            p0 = (cx - hw + skew_x / 2.0, cy - hh - dy_tilt / 2.0)
            p1 = (cx + hw + skew_x / 2.0, cy - hh + dy_tilt / 2.0)
            p2 = (cx + hw - skew_x / 2.0, cy + hh + dy_tilt / 2.0)
            p3 = (cx - hw - skew_x / 2.0, cy + hh - dy_tilt / 2.0)
            _add_polygon(slide, [p0, p1, p2, p3],
                         None if item.get("transparent") else item.get("fill", "#FFFFFF"),
                         item.get("border", "#4F6D8A"),
                         float(item.get("border_weight", 1.1)), item.get("dash"))
            if item.get("text") or item.get("runs"):
                tb = slide.Shapes.AddTextbox(1, left, top, w, h)
                tb.Rotation = float(item.get("text_rotation", tilt_deg))
                _apply_text(
                    tb, str(item.get("text", "")),
                    font_name=str(item.get("font_name", "Times New Roman")),
                    font_size=float(item.get("font_size", 9.5)),
                    bold=bool(item.get("bold", True)),
                    italic=bool(item.get("italic", False)),
                    text_color=str(item.get("text_color", "#111111")),
                    align=int(item.get("align", 2)),
                    valign=int(item.get("valign", 3)),
                    margins=item.get("margins", (0.5, 0.5, 0.5, 0.5)),
                    runs=item.get("runs"),
                    line_spacing=item.get("line_spacing"),
                )
            extra = {
                "left": ((p0[0] + p3[0]) / 2.0, (p0[1] + p3[1]) / 2.0),
                "right": ((p1[0] + p2[0]) / 2.0, (p1[1] + p2[1]) / 2.0),
                "top": ((p0[0] + p1[0]) / 2.0, (p0[1] + p1[1]) / 2.0),
                "bottom": ((p2[0] + p3[0]) / 2.0, (p2[1] + p3[1]) / 2.0),
            }
            _register_anchors(nodes, item, left, top, w, h, extra=extra)
            record_audit((left, top, w, h))
            return 1

        if kind == "prism3d":
            left, top, w, h = _rect_from_item(item, mapper, nodes)
            ex, ey = item.get("extrude", [10, -6])
            ex_pt, ey_pt = mapper.w(ex), mapper.h(ey)
            sx_pt = mapper.w(item.get("skew_x", 0.0))
            sy_pt = mapper.h(item.get("skew_y", 0.0))
            front, shapes = _draw_prism(
                slide, left, top, w, h, ex_pt, ey_pt, sx_pt, sy_pt,
                fill=str(item.get("fill", "#9AC77E")),
                top_fill=item.get("top_fill"),
                side_fill=item.get("side_fill"),
                border=item.get("border", "#4E7A42"),
                border_weight=float(item.get("border_weight", 0.85)),
            )
            if item.get("text"):
                tb = slide.Shapes.AddTextbox(1, left + sx_pt / 2.0, top + sy_pt / 2.0, w, h)
                if item.get("rotation"):
                    tb.Rotation = float(item["rotation"])
                _apply_text(
                    tb, str(item["text"]),
                    font_name=str(item.get("font_name", "Times New Roman")),
                    font_size=float(item.get("font_size", 8.5)),
                    bold=bool(item.get("bold", False)),
                    italic=bool(item.get("italic", False)),
                    text_color=str(item.get("text_color", "#111111")),
                    align=int(item.get("align", 2)),
                    valign=int(item.get("valign", 3)),
                    margins=item.get("margins", (0.5, 0.5, 0.5, 0.5)),
                )
            _register_anchors(nodes, item, left, top, w, h, extra={
                "center": (left + sx_pt / 2.0 + w / 2.0, top + sy_pt / 2.0 + h / 2.0),
                "left": (left + sx_pt / 2.0, top + h / 2.0),
                "right": (left + w + sx_pt / 2.0, top + sy_pt + h / 2.0),
                "top": (left + w / 2.0, top + sy_pt / 2.0),
                "bottom": (left + sx_pt + w / 2.0, top + h + sy_pt / 2.0),
            })
            px0 = min(left, left + sx_pt, left + ex_pt, left + sx_pt + ex_pt)
            px1 = max(left + w, left + w + sx_pt, left + w + ex_pt, left + w + sx_pt + ex_pt)
            py0 = min(top, top + sy_pt, top + ey_pt, top + sy_pt + ey_pt)
            py1 = max(top + h, top + h + sy_pt, top + h + ey_pt, top + h + sy_pt + ey_pt)
            record_audit((px0, py0, px1 - px0, py1 - py0))
            return len(shapes)

        if kind == "grid3d":
            x0, y0 = mapper.pt([item["x"], item["y"]])
            rows = int(item.get("rows", 4))
            cols = int(item.get("cols", 4))
            depth = int(item.get("depth", 3))
            cw = mapper.w(item.get("cell_w", 12))
            ch = mapper.h(item.get("cell_h", 12))
            gx = mapper.w(item.get("gap_x", 3))
            gy = mapper.h(item.get("gap_y", 3))
            ex, ey = item.get("extrude", [4.5, -3.5])
            ex_pt, ey_pt = mapper.w(ex), mapper.h(ey)
            ldx, ldy = item.get("layer_dx", [7, -5])
            ldx_pt, ldy_pt = mapper.w(ldx), mapper.h(ldy)
            sx_pt = mapper.w(item.get("skew_x", 0.0))
            sy_pt = mapper.h(item.get("skew_y", 0.0))
            col_dy = mapper.h(item.get("col_dy", 0.0))
            fill = str(item.get("fill", "#A6C3DE"))
            border = item.get("border", "#4A6580")
            bw = float(item.get("border_weight", 0.7))
            total = 0
            row_order = range(rows) if ey_pt <= 0 else range(rows - 1, -1, -1)
            col_order = range(cols) if ex_pt <= 0 else range(cols - 1, -1, -1)
            for d in range(depth - 1, -1, -1):
                for r in row_order:
                    for c in col_order:
                        lx = x0 + c * (cw + gx) + d * ldx_pt
                        ty = y0 + r * (ch + gy) + c * col_dy + d * ldy_pt
                        _, shapes = _draw_prism(
                            slide, lx, ty, cw, ch, ex_pt, ey_pt, sx_pt, sy_pt,
                            fill=fill, top_fill=item.get("top_fill"),
                            side_fill=item.get("side_fill"), border=border, border_weight=bw,
                        )
                        total += len(shapes)
            full_w = cols * (cw + gx) - gx + (depth - 1) * abs(ldx_pt) + abs(ex_pt)
            full_h = rows * (ch + gy) - gy + (depth - 1) * abs(ldy_pt) + abs(ey_pt)
            gx0 = min(x0, x0 + ex_pt, x0 + (depth - 1) * ldx_pt, x0 + (depth - 1) * ldx_pt + ex_pt)
            gy0 = min(y0, y0 + ey_pt, y0 + (depth - 1) * ldy_pt, y0 + (depth - 1) * ldy_pt + ey_pt,
                      y0 + (cols - 1) * col_dy)
            _register_anchors(nodes, item, x0, y0, full_w, full_h)
            record_audit((gx0, gy0, full_w, full_h))
            return total

        if kind == "feature_stack":
            x0, y0 = mapper.pt([item["x"], item["y"]])
            w = mapper.w(item.get("w", 62))
            h = mapper.h(item.get("h", 58))
            count = int(item.get("count", 3))
            skew_y = mapper.h(item.get("skew_y", -7.0))
            dx = mapper.w(item.get("dx", 8.0))
            dy = mapper.h(item.get("dy", 7.0))
            fill = str(item.get("fill", "#DCEAF7"))
            border = str(item.get("border", "#4F6D8A"))
            bw = float(item.get("border_weight", 1.0))
            front_c = (x0, y0)
            for k in range(count):
                lx = x0 + k * dx
                ty = y0 + k * dy
                pts = [(lx, ty), (lx + w, ty + skew_y), (lx + w, ty + h + skew_y), (lx, ty + h)]
                _add_polygon(slide, pts, fill, border, bw)
                front_c = (lx + w / 2.0, ty + (h + skew_y) / 2.0)
            _register_anchors(nodes, item, x0, y0, w + (count - 1) * dx, h + (count - 1) * dy,
                              extra={"center": front_c, "front_center": front_c})
            fx0 = min(x0, x0 + (count - 1) * dx)
            fy0 = min(y0, y0 + skew_y, y0 + (count - 1) * dy, y0 + (count - 1) * dy + skew_y)
            record_audit((fx0, fy0, w + abs((count - 1) * dx), h + abs((count - 1) * dy) + abs(skew_y)))
            return count

        if kind == "textbox":
            left, top, w, h = _rect_from_item(item, mapper, nodes)
            shape = slide.Shapes.AddTextbox(1, left, top, w, h)
            if item.get("fill"):
                _apply_fill(shape, str(item["fill"]), transparency=item.get("fill_transparency"))
            if item.get("rotation"):
                shape.Rotation = float(item["rotation"])
            _apply_text(
                shape, str(item.get("text", "")),
                font_name=str(item.get("font_name", "Times New Roman")),
                font_size=float(item.get("font_size", 11.0)),
                bold=bool(item.get("bold", True)),
                italic=bool(item.get("italic", False)),
                text_color=str(item.get("color", item.get("text_color", "#111111"))),
                align=int(item.get("align", 2)),
                valign=int(item.get("valign", 3)),
                margins=item.get("margins", (0.0, 0.0, 0.0, 0.0)),
                runs=item.get("runs"),
                line_spacing=item.get("line_spacing"),
            )
            _register_anchors(nodes, item, left, top, w, h)
            record_audit((left, top, w, h))
            return 1

        if kind == "arrow":
            p_start = item["_raw_start"] if "_raw_start" in item else _resolve_point(item["start"], mapper, nodes)
            p_end = item["_raw_end"] if "_raw_end" in item else _resolve_point(item["end"], mapper, nodes)
            via = [_resolve_point(v, mapper, nodes) for v in item.get("via", [])]
            color = str(item.get("color", "#262626"))
            weight = float(item.get("weight", 1.5))
            end_style = _arrow_value(item.get("arrow", True))
            begin_style = _arrow_value(item.get("begin_arrow", False))
            dash_val = _dash_value(item.get("dash"))
            if not via:
                shape = slide.Shapes.AddConnector(1, p_start[0], p_start[1], p_end[0], p_end[1])
            else:
                ff = slide.Shapes.BuildFreeform(1, p_start[0], p_start[1])
                for vx, vy in via:
                    ff.AddNodes(0, 0, vx, vy)
                ff.AddNodes(0, 0, p_end[0], p_end[1])
                shape = ff.ConvertToShape()
            shape.Line.ForeColor.RGB = _rgb(color)
            shape.Line.Weight = weight
            shape.Line.EndArrowheadStyle = end_style
            shape.Line.BeginArrowheadStyle = begin_style
            if item.get("arrow_length") is not None:
                shape.Line.EndArrowheadLength = int(item["arrow_length"])
            if item.get("arrow_width") is not None:
                shape.Line.EndArrowheadWidth = int(item["arrow_width"])
            if dash_val is not None:
                shape.Line.DashStyle = dash_val
            return 1

        if kind == "curve_arrow":
            p0 = _resolve_point(item["start"], mapper, nodes)
            p3 = _resolve_point(item["end"], mapper, nodes)
            if "arc_height" in item:
                ah = mapper.h(float(item["arc_height"]))
                c1, c2 = _bezier_from_arc_height(p0, p3, ah, float(item.get("arc_skew", 0.0)))
            else:
                c1 = _resolve_point(item["c1"], mapper, nodes)
                c2 = _resolve_point(item["c2"], mapper, nodes)
            ff = slide.Shapes.BuildFreeform(1, p0[0], p0[1])
            ff.AddNodes(1, 1, c1[0], c1[1], c2[0], c2[1], p3[0], p3[1])
            shape = ff.ConvertToShape()
            shape.Line.ForeColor.RGB = _rgb(str(item.get("color", "#262626")))
            shape.Line.Weight = float(item.get("weight", 1.6))
            shape.Line.EndArrowheadStyle = _arrow_value(item.get("arrow", True))
            shape.Line.BeginArrowheadStyle = _arrow_value(item.get("begin_arrow", False))
            if item.get("arrow_length") is not None:
                shape.Line.EndArrowheadLength = int(item["arrow_length"])
            if item.get("arrow_width") is not None:
                shape.Line.EndArrowheadWidth = int(item["arrow_width"])
            dash_val = _dash_value(item.get("dash"))
            if dash_val is not None:
                shape.Line.DashStyle = dash_val
            return 1

        if kind == "brace":
            left, top, w, h = _rect_from_item(item, mapper, nodes)
            shape = slide.Shapes.AddShape(32, left, top, w, h)
            _apply_fill(shape, None, transparent=True)
            _apply_line(shape, str(item.get("color", "#595959")), float(item.get("weight", 1.0)))
            if item.get("rotation") is not None:
                shape.Rotation = float(item["rotation"])
            _register_anchors(nodes, item, left, top, w, h)
            record_audit((left, top, w, h))
            return 1

        if kind == "image":
            img_path = Path(str(item["path"])).expanduser()
            if not img_path.is_absolute() or not img_path.is_file():
                raise _fail(f"Image path must be an existing absolute file: {img_path}")
            left, top, w, h = _rect_from_item(item, mapper, nodes)
            crop = item.get("crop_frac")
            if crop is not None:
                if not isinstance(crop, (list, tuple)) or len(crop) != 4:
                    raise _fail("crop_frac must be [left_frac, top_frac, right_frac, bottom_frac] in 0..1.")
                cl, ct, cr, cb = [float(v) for v in crop]
                if min(cl, ct, cr, cb) < 0 or cl + cr >= 0.99 or ct + cb >= 0.99:
                    raise _fail("Invalid crop_frac values; remaining width and height must be positive.")
                pic = slide.Shapes.AddPicture(str(img_path), False, True, left, top, -1, -1)
                pic.LockAspectRatio = 0
                orig_w = float(pic.Width)
                orig_h = float(pic.Height)
                pf = pic.PictureFormat
                pf.CropLeft = orig_w * cl
                pf.CropTop = orig_h * ct
                pf.CropRight = orig_w * cr
                pf.CropBottom = orig_h * cb
                pic.Left = left
                pic.Top = top
                pic.Width = w
                pic.Height = h
            else:
                slide.Shapes.AddPicture(str(img_path), False, True, left, top, w, h)
            _register_anchors(nodes, item, left, top, w, h)
            record_audit((left, top, w, h))
            return 1

        raise _fail(f"Unknown scene item type '{kind}'.")
