"""1-to-1 live reproduction of the Conformer architecture diagram in Microsoft PowerPoint via microsoft-wps-live-mcp."""
from __future__ import annotations

import asyncio
import json
from pathlib import Path
import sys
import uuid

from mcp import Client, StdioServerParameters

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = ROOT / "data" / "conformer_architecture.pptx"
PEACOCK_PATH = ROOT / "data" / "peacock_input.jpg"

# Reference image is 1010x531; white slide area spans roughly px=38..1008, py=5..528.
# Map cleanly onto PowerPoint's 960x540 widescreen canvas.
SX = 960.0 / 970.0
SY = 540.0 / 531.0


def X(px: float) -> float:
    return round((px - 38.0) * SX, 2)


def Y(py: float) -> float:
    return round(py * SY, 2)


def W(pw: float) -> float:
    return round(pw * SX, 2)


def H(ph: float) -> float:
    return round(ph * SY, 2)


def to_bgr(hex_str: str) -> int:
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r | (g << 8) | (b << 16)


def unpack(res):
    if res.is_error:
        raise RuntimeError(str(res.content))
    data = res.structured_content or json.loads(res.content[0].text)
    if data.get("state") in ("partial_error", "display_error", "pending"):
        raise RuntimeError(json.dumps(data, ensure_ascii=False))
    return data


def uid(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:8]}"


def add_polygon_steps(points: list[tuple[float, float]], fill: str, line_color: str | None = None,
                      line_weight: float = 1.0, dash: int | None = None) -> list[dict]:
    name = uid("poly")
    ff = uid("ff")
    x0, y0 = points[0]
    steps = [
        {"action": "call", "target": "slide.Shapes", "member": "BuildFreeform", "args": [0, x0, y0], "as": ff},
    ]
    for x, y in points[1:]:
        steps.append({"action": "call", "target": ff, "member": "AddNodes", "args": [0, 0, x, y]})
    steps.extend([
        {"action": "call", "target": ff, "member": "AddNodes", "args": [0, 0, x0, y0]},
        {"action": "call", "target": ff, "member": "ConvertToShape", "as": name},
        {"action": "set", "target": f"{name}.Fill", "member": "Visible", "value": -1},
        {"action": "call", "target": f"{name}.Fill", "member": "Solid", "args": []},
        {"action": "set", "target": f"{name}.Fill.ForeColor", "member": "RGB", "value": to_bgr(fill)},
    ])
    if line_color is None:
        steps.append({"action": "set", "target": f"{name}.Line", "member": "Visible", "value": 0})
    else:
        steps.extend([
            {"action": "set", "target": f"{name}.Line", "member": "Visible", "value": -1},
            {"action": "set", "target": f"{name}.Line.ForeColor", "member": "RGB", "value": to_bgr(line_color)},
            {"action": "set", "target": f"{name}.Line", "member": "Weight", "value": line_weight},
        ])
        if dash is not None:
            steps.append({"action": "set", "target": f"{name}.Line", "member": "DashStyle", "value": dash})
    return steps


def add_feature_map_stack(px_left: float, py_top: float, pw: float = 64, ph: float = 58,
                          skew: float = -7.0, dx: float = 8.0, dy: float = 7.0) -> list[dict]:
    steps = []
    for k in range(3):
        lx = X(px_left + k * dx)
        ty = Y(py_top + k * dy)
        rx = lx + W(pw)
        by = ty + H(ph)
        sk = H(skew)
        pts = [
            (lx, ty),
            (rx, ty + sk),
            (rx, by + sk),
            (lx, by),
        ]
        steps.extend(add_polygon_steps(pts, fill="#DCEAF7", line_color="#4F6D8A", line_weight=1.0))
    return steps


def add_box_steps(cx: float, cy: float, w: float, h: float, text: str, fill: str, border: str,
                  text_color: str = "#111111", font_size: float = 9.5, rotation: float = 0.0,
                  rounded: bool = False, bold: bool = True, border_weight: float = 1.1,
                  dash: int | None = None, transparent_fill: bool = False,
                  shape_type: int | None = None, font_name: str = "Times New Roman") -> list[dict]:
    name = uid("box")
    left = X(cx) - W(w) / 2.0
    top = Y(cy) - H(h) / 2.0
    stype = shape_type if shape_type is not None else (5 if rounded else 1)
    steps = [
        {"action": "call", "target": "slide.Shapes", "member": "AddShape",
         "args": [stype, left, top, W(w), H(h)], "as": name},
    ]
    if transparent_fill:
        steps.append({"action": "set", "target": f"{name}.Fill", "member": "Visible", "value": 0})
    else:
        steps.extend([
            {"action": "set", "target": f"{name}.Fill", "member": "Visible", "value": -1},
            {"action": "call", "target": f"{name}.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": f"{name}.Fill.ForeColor", "member": "RGB", "value": to_bgr(fill)},
        ])
    if border_weight <= 0:
        steps.append({"action": "set", "target": f"{name}.Line", "member": "Visible", "value": 0})
    else:
        steps.extend([
            {"action": "set", "target": f"{name}.Line", "member": "Visible", "value": -1},
            {"action": "set", "target": f"{name}.Line.ForeColor", "member": "RGB", "value": to_bgr(border)},
            {"action": "set", "target": f"{name}.Line", "member": "Weight", "value": border_weight},
        ])
    if dash is not None:
        steps.append({"action": "set", "target": f"{name}.Line", "member": "DashStyle", "value": dash})
    if rotation != 0.0:
        steps.append({"action": "set", "target": name, "member": "Rotation", "value": rotation})
    if text:
        steps.extend([
            {"action": "set", "target": f"{name}.TextFrame", "member": "MarginLeft", "value": 1.0},
            {"action": "set", "target": f"{name}.TextFrame", "member": "MarginRight", "value": 1.0},
            {"action": "set", "target": f"{name}.TextFrame", "member": "MarginTop", "value": 0.5},
            {"action": "set", "target": f"{name}.TextFrame", "member": "MarginBottom", "value": 0.5},
            {"action": "set", "target": f"{name}.TextFrame", "member": "WordWrap", "value": -1},
            {"action": "set", "target": f"{name}.TextFrame", "member": "VerticalAnchor", "value": 3},
            {"action": "set", "target": f"{name}.TextFrame.TextRange", "member": "Text", "value": text},
            {"action": "set", "target": f"{name}.TextFrame.TextRange.Font", "member": "Name", "value": font_name},
            {"action": "set", "target": f"{name}.TextFrame.TextRange.Font", "member": "Size", "value": font_size},
            {"action": "set", "target": f"{name}.TextFrame.TextRange.Font", "member": "Bold", "value": -1 if bold else 0},
            {"action": "set", "target": f"{name}.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr(text_color)},
            {"action": "set", "target": f"{name}.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},
        ])
    return steps


def add_text_steps(left_px: float, top_px: float, w_px: float, h_px: float, text: str,
                   size: float = 11.0, color: str = "#111111", bold: bool = True,
                   italic: bool = False, rotation: float = 0.0, align: int = 2,
                   font_name: str = "Times New Roman") -> list[dict]:
    name = uid("txt")
    steps = [
        {"action": "call", "target": "slide.Shapes", "member": "AddTextbox",
         "args": [1, X(left_px), Y(top_px), W(w_px), H(h_px)], "as": name},
        {"action": "set", "target": f"{name}.TextFrame", "member": "MarginLeft", "value": 0.0},
        {"action": "set", "target": f"{name}.TextFrame", "member": "MarginRight", "value": 0.0},
        {"action": "set", "target": f"{name}.TextFrame", "member": "MarginTop", "value": 0.0},
        {"action": "set", "target": f"{name}.TextFrame", "member": "MarginBottom", "value": 0.0},
        {"action": "set", "target": f"{name}.TextFrame.TextRange", "member": "Text", "value": text},
        {"action": "set", "target": f"{name}.TextFrame.TextRange.Font", "member": "Name", "value": font_name},
        {"action": "set", "target": f"{name}.TextFrame.TextRange.Font", "member": "Size", "value": size},
        {"action": "set", "target": f"{name}.TextFrame.TextRange.Font", "member": "Bold", "value": -1 if bold else 0},
        {"action": "set", "target": f"{name}.TextFrame.TextRange.Font", "member": "Italic", "value": -1 if italic else 0},
        {"action": "set", "target": f"{name}.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr(color)},
        {"action": "set", "target": f"{name}.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": align},
    ]
    if rotation != 0.0:
        steps.append({"action": "set", "target": name, "member": "Rotation", "value": rotation})
    return steps


def add_line_steps(x1: float, y1: float, x2: float, y2: float, color: str = "#262626",
                   weight: float = 1.4, arrow: bool = True, begin_arrow: bool = False,
                   dash: int | None = None) -> list[dict]:
    name = uid("ln")
    steps = [
        {"action": "call", "target": "slide.Shapes", "member": "AddConnector",
         "args": [1, X(x1), Y(y1), X(x2), Y(y2)], "as": name},
        {"action": "set", "target": f"{name}.Line.ForeColor", "member": "RGB", "value": to_bgr(color)},
        {"action": "set", "target": f"{name}.Line", "member": "Weight", "value": weight},
        {"action": "set", "target": f"{name}.Line", "member": "EndArrowheadStyle", "value": 2 if arrow else 1},
        {"action": "set", "target": f"{name}.Line", "member": "BeginArrowheadStyle", "value": 2 if begin_arrow else 1},
    ]
    if dash is not None:
        steps.append({"action": "set", "target": f"{name}.Line", "member": "DashStyle", "value": dash})
    return steps


def add_curve_arrow_steps(x0: float, y0: float, cx1: float, cy1: float, cx2: float, cy2: float,
                          x3: float, y3: float, color: str, weight: float = 1.6) -> list[dict]:
    ff = uid("cff")
    arc = uid("arc")
    return [
        {"action": "call", "target": "slide.Shapes", "member": "BuildFreeform", "args": [1, X(x0), Y(y0)], "as": ff},
        {"action": "call", "target": ff, "member": "AddNodes",
         "args": [1, 1, X(cx1), Y(cy1), X(cx2), Y(cy2), X(x3), Y(y3)]},
        {"action": "call", "target": ff, "member": "ConvertToShape", "as": arc},
        {"action": "set", "target": f"{arc}.Line.ForeColor", "member": "RGB", "value": to_bgr(color)},
        {"action": "set", "target": f"{arc}.Line", "member": "Weight", "value": weight},
        {"action": "set", "target": f"{arc}.Line", "member": "EndArrowheadStyle", "value": 2},
    ]


def add_patch_bar_stack(px_left: float, py_top: float, pw: float = 92, ph: float = 9.0,
                        rotation: float = -10.0) -> list[dict]:
    steps = []
    # 3 green bars back-to-front + 1 cream/orange class-token bar in front
    for k in range(4):
        fill = "#FCE3BE" if k == 3 else "#9AC77E"
        border = "#C99A5B" if k == 3 else "#5B8C4F"
        cx = px_left + pw / 2.0 + k * 4.0
        cy = py_top + k * 8.0
        steps.extend(add_box_steps(
            cx=cx, cy=cy, w=pw, h=ph, text="", fill=fill, border=border,
            rotation=rotation, shape_type=14, border_weight=0.8
        ))
    return steps


def build_groups() -> list[list[dict]]:
    # Group 1: Backdrop (FCU peach region + inset frames + labels)
    fcu_pts = [
        (X(266), Y(246)),
        (X(562), Y(180)),
        (X(766), Y(336)),
        (X(422), Y(416)),
    ]
    g1 = [
        {
            "label": "Setup widescreen slide & FCU coupling backdrop",
            "focus": {"slide": 1},
            "steps": [
                {"action": "set", "target": "document.PageSetup", "member": "SlideWidth", "value": 960},
                {"action": "set", "target": "document.PageSetup", "member": "SlideHeight", "value": 540},
                {"action": "call", "target": "document.Slides", "member": "Item", "args": [1], "as": "slide"},
                *add_polygon_steps(fcu_pts, fill="#FDECE4", line_color=None),
                *add_text_steps(475, 262, 65, 24, "FCU", size=14, rotation=32),
            ],
        },
        {
            "label": "Draw inset (a) & (c) dashed boundaries and figure labels",
            "focus": {"slide": 1},
            "steps": [
                # Inset (a) box + Up/Down legend
                *add_box_steps(210, 442, 312, 114, "", "#FFFFFF", "#7F7F7F",
                               rounded=True, border_weight=1.1, dash=4, transparent_fill=True),
                *add_line_steps(92, 366, 124, 366, color="#5B8C4F", weight=1.6, arrow=True),
                *add_text_steps(126, 357, 28, 18, "Up", size=10.5, bold=True, align=1),
                *add_line_steps(190, 366, 158, 366, color="#3F6AA8", weight=1.6, arrow=True),
                *add_text_steps(194, 357, 40, 18, "Down", size=10.5, bold=True, align=1),
                *add_text_steps(192, 504, 36, 20, "(a)", size=12.5, bold=False),
                # Inset (c) box + inner N-1 dashed box
                *add_box_steps(824, 118, 340, 212, "", "#FFFFFF", "#7F7F7F",
                               rounded=True, border_weight=1.1, dash=4, transparent_fill=True),
                *add_box_steps(892, 98, 62, 144, "", "#FFFFFF", "#595959",
                               rounded=True, border_weight=1.0, dash=4, transparent_fill=True),
                *add_text_steps(874, 12, 38, 18, "N-1", size=10.5, bold=False, italic=True),
                *add_text_steps(808, 228, 36, 20, "(c)", size=12.5, bold=False),
                *add_text_steps(758, 502, 36, 20, "(b)", size=12.5, bold=False),
            ],
        },
    ]

    # Group 2: CNN Branch (Feature Maps + Upper & Lower Conv-BN Trios + Blue Skip Curves)
    g2 = [
        {
            "label": "Draw top Feature Maps stack and dimensions W, H, C",
            "focus": {"slide": 1},
            "steps": [
                *add_feature_map_stack(82, 108, pw=62, ph=58),
                *add_text_steps(102, 83, 24, 18, "W", size=11, bold=True, italic=True),
                *add_text_steps(52, 138, 22, 18, "H", size=11, bold=True, italic=True),
                *add_text_steps(69, 176, 22, 18, "C", size=11, bold=True, italic=True),
                *add_text_steps(156, 86, 105, 22, "Feature Maps", size=11.5, bold=True, align=1),
            ],
        },
        {
            "label": "Draw upper CNN Conv-BN blocks & blue residual curve",
            "focus": {"slide": 1},
            "steps": [
                *add_line_steps(130, 146, 306, 304, color="#3F6AA8", weight=1.6, arrow=True),
                *add_box_steps(172, 198, 82, 22, "1x1Conv-BN", "#FFFFFF", "#4F6D8A", font_size=9.5, rotation=-12),
                *add_box_steps(212, 234, 82, 22, "3x3Conv-BN", "#FFFFFF", "#4F6D8A", font_size=9.5, rotation=-12),
                *add_box_steps(252, 268, 82, 22, "1x1Conv-BN", "#FFFFFF", "#4F6D8A", font_size=9.5, rotation=-12),
                *add_curve_arrow_steps(132, 144, 235, 58, 308, 125, 316, 286, color="#3F6AA8", weight=1.6),
            ],
        },
        {
            "label": "Draw middle Feature Maps, lower CNN Conv-BN blocks & branch title",
            "focus": {"slide": 1},
            "steps": [
                *add_feature_map_stack(310, 284, pw=62, ph=56),
                *add_line_steps(364, 324, 536, 468, color="#3F6AA8", weight=1.6, arrow=True),
                *add_box_steps(414, 372, 82, 22, "1x1Conv-BN", "#FFFFFF", "#4F6D8A", font_size=9.5, rotation=-12),
                *add_box_steps(452, 405, 82, 22, "3x3Conv-BN", "#FFFFFF", "#4F6D8A", font_size=9.5, rotation=-12),
                *add_box_steps(490, 438, 82, 22, "1x1Conv-BN", "#FFFFFF", "#4F6D8A", font_size=9.5, rotation=-12),
                *add_curve_arrow_steps(364, 322, 480, 255, 536, 322, 540, 452, color="#3F6AA8", weight=1.6),
                *add_text_steps(384, 458, 105, 22, "CNN Branch", size=12, bold=True, rotation=40),
                *add_feature_map_stack(536, 452, pw=58, ph=54),
            ],
        },
    ]

    # Group 3: Transformer Branch (Patch Embeddings + LayerNorm / MHSA / MLP + Green Skip Curves)
    g3 = [
        {
            "label": "Draw top Patch Embeddings and dimensions E, K+1",
            "focus": {"slide": 1},
            "steps": [
                *add_text_steps(428, 96, 125, 22, "Patch Embeddings", size=11.5, bold=True),
                *add_text_steps(482, 118, 22, 18, "E", size=11, bold=True, italic=True),
                *add_text_steps(404, 152, 36, 18, "K+1", size=11, bold=True, italic=True),
                *add_patch_bar_stack(444, 140, pw=92, ph=9, rotation=-10),
            ],
        },
        {
            "label": "Draw Transformer blocks, green residual curves & bottom embeddings",
            "focus": {"slide": 1},
            "steps": [
                *add_line_steps(508, 168, 804, 392, color="#5B8C4F", weight=1.6, arrow=True),
                *add_box_steps(574, 192, 74, 22, "LayerNorm", "#FFFFFF", "#5B8C4F", font_size=9.5, rotation=-12),
                *add_box_steps(624, 232, 96, 30, "Multi-head\nSelf-attention", "#FFFFFF", "#5B8C4F", font_size=9.0, rotation=-12),
                *add_curve_arrow_steps(554, 178, 645, 132, 694, 174, 666, 254, color="#5B8C4F", weight=1.6),
                *add_box_steps(674, 272, 74, 22, "LayerNorm", "#FFFFFF", "#5B8C4F", font_size=9.5, rotation=-12),
                *add_box_steps(716, 306, 48, 22, "MLP", "#FFFFFF", "#5B8C4F", font_size=9.5, rotation=-12),
                *add_curve_arrow_steps(662, 254, 764, 222, 804, 274, 796, 380, color="#5B8C4F", weight=1.6),
                *add_text_steps(754, 266, 132, 22, "Transformer Branch", size=12, bold=True, rotation=42),
                *add_patch_bar_stack(784, 388, pw=92, ph=9, rotation=-10),
            ],
        },
    ]

    # Group 4: FCU Coupling Bridges (Blue Down Path & Green Up Path)
    g4 = [
        {
            "label": "Draw top FCU Down-coupling bridge (CNN -> Transformer)",
            "focus": {"slide": 1},
            "steps": [
                *add_line_steps(288, 265, 536, 203, color="#3F6AA8", weight=1.6, arrow=True),
                *add_box_steps(340, 252, 62, 21, "1x1Conv", "#FFFFFF", "#4A75A8", font_size=9.5, rotation=-12),
                *add_box_steps(408, 235, 48, 21, "Down", "#9EC3E6", "#4A75A8", font_size=9.5, rotation=-12),
                *add_box_steps(484, 216, 74, 21, "LayerNorm", "#FFFFFF", "#4A75A8", font_size=9.5, rotation=-12),
            ],
        },
        {
            "label": "Draw bottom FCU Up-coupling bridge (Transformer -> CNN)",
            "focus": {"slide": 1},
            "steps": [
                *add_line_steps(768, 344, 444, 392, color="#5B8C4F", weight=1.6, arrow=True),
                *add_box_steps(710, 354, 42, 21, "Up", "#84C475", "#4E8A42", font_size=9.5, rotation=-12),
                *add_box_steps(640, 368, 62, 21, "1x1Conv", "#FFFFFF", "#5B8C4F", font_size=9.5, rotation=-12),
                *add_box_steps(560, 382, 74, 21, "BatchNorm", "#FFFFFF", "#5B8C4F", font_size=9.5, rotation=-12),
            ],
        },
    ]

    # Group 5: Inset (a) 3D Cubes + Inset (c) Overall Pipeline
    a_left_mid = []
    # Left 4 green stacked 3D cubes
    for idx in range(4):
        a_left_mid.extend(add_box_steps(
            90, 412 + idx * 17, 32, 11, "", "#9AC77E", "#4E7A42",
            shape_type=14, border_weight=0.8, rotation=-8
        ))
    # Middle tilted blue 3D cubes (2 pairs)
    for row in range(2):
        for col in range(2):
            a_left_mid.extend(add_box_steps(
                182 + col * 18 + row * 6, 434 + row * 16 - col * 4, 24, 10, "", "#A6C3DE", "#4A6580",
                shape_type=14, border_weight=0.8, rotation=-18
            ))

    # Right 4x4x3 spatial blue 3D feature map grid (split into depth slices so each stage <= 220 steps)
    a_grid_slices = []
    for d in range(2, -1, -1):
        slice_steps = []
        for r in range(4):
            for c in range(4):
                cx = 282 + c * 15 + d * 6
                cy = 420 + r * 15 - d * 5
                slice_steps.extend(add_box_steps(
                    cx, cy, 13, 13, "", "#A6C3DE", "#4A6580",
                    shape_type=14, border_weight=0.7
                ))
        a_grid_slices.append(slice_steps)

    a_arrows = [
        *add_text_steps(112, 396, 56, 18, "Reshaping", size=9.0, bold=False),
        *add_line_steps(156, 422, 122, 422, color="#3F6AA8", weight=1.6, arrow=True),
        *add_text_steps(208, 396, 62, 18, "Avgpooling", size=9.0, bold=False),
        *add_line_steps(260, 422, 224, 422, color="#3F6AA8", weight=1.6, arrow=True),
        *add_line_steps(122, 464, 156, 464, color="#5B8C4F", weight=1.6, arrow=True),
        *add_text_steps(112, 476, 56, 18, "Reshaping", size=9.0, bold=False),
        *add_line_steps(224, 464, 260, 464, color="#5B8C4F", weight=1.6, arrow=True),
        *add_text_steps(208, 476, 66, 18, "Interpolating", size=9.0, bold=False),
    ]

    c_steps_1 = [
        # Peacock input thumbnail
        {"action": "call", "target": "slide.Shapes", "member": "AddPicture",
         "args": [str(PEACOCK_PATH), False, True, X(666), Y(88), W(25), H(58)]},
        *add_line_steps(691, 117, 704, 117, color="#262626", weight=1.3, arrow=True),
        # Stem (vertical rounded box)
        *add_box_steps(715, 117, 66, 22, "Stem", "#FFFFFF", "#7F7F7F", font_size=9.5,
                       rotation=-90, rounded=True, bold=False, border_weight=1.0),
        # Split lines from Stem
        *add_line_steps(726, 117, 740, 117, color="#262626", weight=1.2, arrow=False),
        *add_line_steps(740, 52, 740, 148, color="#262626", weight=1.2, arrow=False),
        *add_line_steps(740, 52, 756, 52, color="#262626", weight=1.2, arrow=True),
        *add_line_steps(740, 148, 794, 148, color="#262626", weight=1.2, arrow=True),
        # Project (vertical green bar)
        *add_box_steps(768, 52, 54, 22, "Project", "#9AC77E", "#5B8C4F", font_size=9.5,
                       rotation=-90, rounded=True, bold=False, border_weight=1.0),
        *add_line_steps(780, 52, 794, 52, color="#262626", weight=1.2, arrow=True),
        # Trans Block 1 & 2 (green 3D cubes)
        *add_box_steps(818, 52, 42, 42, "Trans\nBlock", "#C8E0B4", "#5B8C4F", font_size=8.5,
                       shape_type=14, bold=False, border_weight=1.0),
        *add_line_steps(840, 52, 870, 52, color="#262626", weight=1.2, arrow=True),
        *add_box_steps(894, 52, 42, 42, "Trans\nBlock", "#C8E0B4", "#5B8C4F", font_size=8.5,
                       shape_type=14, bold=False, border_weight=1.0),
    ]

    c_steps_2 = [
        # Conv Block 1 & 2 (blue 3D cubes)
        *add_box_steps(818, 148, 42, 42, "Conv\nBlock", "#BDD7EE", "#4A6580", font_size=8.5,
                       shape_type=14, bold=False, border_weight=1.0),
        *add_line_steps(840, 148, 870, 148, color="#262626", weight=1.2, arrow=True),
        *add_box_steps(894, 148, 42, 42, "Conv\nBlock", "#BDD7EE", "#4A6580", font_size=8.5,
                       shape_type=14, bold=False, border_weight=1.0),
        # Coupling arrows inside N-1 box
        *add_line_steps(885, 126, 885, 76, color="#5B8C4F", weight=1.5, arrow=True, dash=4),
        *add_line_steps(902, 76, 902, 126, color="#3F6AA8", weight=1.5, arrow=True, dash=4),
        # Bottom brace & N
        *add_box_steps(856, 178, 14, 120, "", "#FFFFFF", "#595959", rotation=90,
                       shape_type=32, border_weight=1.0, transparent_fill=True),
        *add_text_steps(846, 190, 22, 18, "N", size=10.5, bold=False, italic=True),
        # Outputs to Classifiers
        *add_line_steps(915, 52, 956, 52, color="#262626", weight=1.2, arrow=True, dash=4),
        *add_box_steps(936, 52, 42, 18, "Class\ntoken", "#FFFFFF", "#FFFFFF", font_size=7.0,
                       rotation=-90, bold=False, border_weight=0.0),
        *add_box_steps(970, 52, 58, 22, "Classifier", "#E2EFDA", "#5B8C4F", font_size=9.5,
                       rotation=-90, rounded=True, bold=False, border_weight=1.0),
        *add_line_steps(915, 148, 956, 148, color="#262626", weight=1.2, arrow=True, dash=4),
        *add_box_steps(936, 148, 38, 14, "Pooling", "#FFFFFF", "#FFFFFF", font_size=7.5,
                       rotation=-90, bold=False, border_weight=0.0),
        *add_box_steps(970, 148, 58, 22, "Classifier", "#DDEBF7", "#4A6580", font_size=9.5,
                       rotation=-90, rounded=True, bold=False, border_weight=1.0),
    ]

    g5 = [
        {"label": "Draw inset (a) patch & reshaped 3D cubes", "focus": {"slide": 1}, "steps": a_left_mid},
        {"label": "Draw inset (a) spatial cube grid layer 1", "focus": {"slide": 1}, "steps": a_grid_slices[0]},
        {"label": "Draw inset (a) spatial cube grid layer 2", "focus": {"slide": 1}, "steps": a_grid_slices[1]},
    ]
    g6 = [
        {"label": "Draw inset (a) spatial cube grid front layer & labels", "focus": {"slide": 1},
         "steps": a_grid_slices[2] + a_arrows},
        {"label": "Draw inset (c) input image, Stem & Transformer branch", "focus": {"slide": 1}, "steps": c_steps_1},
        {"label": "Draw inset (c) CNN branch, coupling & Classifiers", "focus": {"slide": 1}, "steps": c_steps_2},
    ]

    return [g1, g2, g3, g4, g5, g6]


async def main():
    if not PEACOCK_PATH.is_file():
        raise FileNotFoundError(f"Missing cropped input image: {PEACOCK_PATH}")

    out_path = OUTPUT_PATH
    # Clean up old output file so office_create_document creates fresh
    if out_path.exists():
        try:
            out_path.unlink()
        except Exception:
            out_path = ROOT / "data" / f"conformer_architecture_{uuid.uuid4().hex[:6]}.pptx"

    params = StdioServerParameters(
        command=sys.executable,
        args=["-X", "utf8", "-u", "-m", "microsoft_wps_live_mcp"]
    )
    print("[*] Launching microsoft-wps-live-mcp over stdio...", flush=True)

    async with Client(params, read_timeout_seconds=120) as client:
        status = unpack(await client.call_tool("office_status", {}))
        print(f"[+] MCP connected. Platform={status['platform']}, PPT registered={status['engines']['microsoft']['ppt']['registered']}", flush=True)

        # Close any previous unfinished conformer_architecture draft in PowerPoint
        open_docs = unpack(await client.call_tool("office_list_open_documents", {"engine": "microsoft"}))
        for doc in open_docs.get("documents", []):
            if doc.get("path") and Path(doc["path"]).name.startswith("conformer_architecture"):
                await client.call_tool("office_close_document", {"document": doc["document_id"]})
        if out_path.exists():
            try:
                out_path.unlink()
            except Exception:
                out_path = ROOT / "data" / f"conformer_architecture_{uuid.uuid4().hex[:6]}.pptx"

        # Create visible native Microsoft PowerPoint presentation
        created = unpack(await client.call_tool("office_create_document", {
            "path": str(out_path),
            "engine": "microsoft",
            "visible": True,
        }))
        doc_id = created["document_id"]
        print(f"[+] Created PowerPoint document: {created['path']} (hwnd={created['live_view'].get('hwnd')})", flush=True)

        # Raise the real PowerPoint window immediately before drawing
        shown = unpack(await client.call_tool("office_show_document", {
            "document": doc_id,
            "focus": {"slide": 1},
        }))
        print(f"[+] Raised PowerPoint window: live_view={shown['live_view']}", flush=True)

        groups = build_groups()
        for g_idx, stages in enumerate(groups, 1):
            # Each office_live_steps call has its own alias scope, so ensure 'slide'
            # is bound in the first stage of every call (except group 1 which already binds it).
            if g_idx > 1:
                stages[0]["steps"].insert(
                    0, {"action": "call", "target": "document.Slides", "member": "Item", "args": [1], "as": "slide"}
                )
            print(f"[*] Running live batch group {g_idx}/{len(groups)} ({len(stages)} stages)...", flush=True)

            async def on_progress(number, total, message):
                print(f"    -> Stage {int(number)}/{int(total)}: {message}", flush=True)

            res = unpack(await client.call_tool("office_live_steps", {
                "document": doc_id,
                "stages": stages,
                "pause_ms": 250,
                "step_pause_ms": 12,
            }, progress_callback=on_progress))
            print(f"[+] Group {g_idx} finished: completed_stages={res['completed_stages']}, disk_saved={res['disk_saved']}", flush=True)

        # Create verified independent final copy (per MCP instructions) and leave window open!
        copied = unpack(await client.call_tool("office_create_copy", {"document": doc_id}))
        print(f"[+] Original:   {copied['path']}", flush=True)
        print(f"[+] Final copy: {copied['copy_path']}", flush=True)
        print(f"[+] Recovery:   {copied['recovery_copy']}", flush=True)


if __name__ == "__main__":
    asyncio.run(main())
