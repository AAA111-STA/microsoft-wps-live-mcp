"""True 1-to-1 Conformer architecture diagram using the upgraded office_ppt_draw_scene MCP tool."""
from __future__ import annotations

import asyncio
import json
from pathlib import Path
import sys
import time
import uuid

from mcp import Client, StdioServerParameters

ROOT = Path(__file__).resolve().parents[1]
PEACOCK_PATH = ROOT / "data" / "peacock_input.jpg"


def unpack(res):
    if res.is_error:
        raise RuntimeError(str(res.content))
    data = res.structured_content or json.loads(res.content[0].text)
    if data.get("state") in ("partial_error", "display_error", "pending"):
        raise RuntimeError(json.dumps(data, ensure_ascii=False))
    return data


def build_conformer_scene() -> tuple[dict, list[dict]]:
    coord_space = {"width": 970, "height": 531, "origin_x": 38, "origin_y": 0, "fit": "stretch"}

    # Exact trunk endpoints calibrated to reference pixel geometry
    cnn_up_a, cnn_up_b = [130.0, 146.0], [306.0, 300.0]
    cnn_lo_a, cnn_lo_b = [362.0, 326.0], [536.0, 468.0]
    tr_a, tr_b = [504.0, 156.0], [778.0, 376.0]

    # Stage 1: Backdrop (exact FCU region matching original!), dashed inset frames & legends
    s1 = {
        "label": "Setup backdrop, FCU zone & dashed inset boundaries",
        "items": [
            {
                "type": "polygon",
                "points": [[284, 246], [552, 184], [742, 336], [448, 404]],
                "fill": "#FDECE4",
                "border": None,
            },
            {
                "type": "textbox", "x": 475, "y": 262, "w": 65, "h": 24,
                "text": "FCU", "font_size": 14, "bold": True, "rotation": 32,
            },
            # Inset (a) outer dashed frame + Up/Down legend
            {
                "type": "box", "cx": 210, "cy": 442, "w": 312, "h": 114,
                "transparent": True, "border": "#7F7F7F", "border_weight": 1.1,
                "dash": "dash", "rounded": True, "corner_radius": 0.14,
            },
            {"type": "arrow", "start": [92, 366], "end": [124, 366], "color": "#5B8C4F", "weight": 1.6},
            {"type": "textbox", "x": 126, "y": 357, "w": 28, "h": 18, "text": "Up", "font_size": 10.5, "align": 1},
            {"type": "arrow", "start": [190, 366], "end": [158, 366], "color": "#3F6AA8", "weight": 1.6},
            {"type": "textbox", "x": 194, "y": 357, "w": 40, "h": 18, "text": "Down", "font_size": 10.5, "align": 1},
            {"type": "textbox", "x": 192, "y": 504, "w": 36, "h": 20, "text": "(a)", "font_size": 12.5, "bold": False},
            # Inset (c) outer dashed frame + inner N-1 dashed frame
            {
                "type": "box", "cx": 824, "cy": 112, "w": 340, "h": 202,
                "transparent": True, "border": "#7F7F7F", "border_weight": 1.1,
                "dash": "dash", "rounded": True, "corner_radius": 0.10,
            },
            {
                "type": "box", "cx": 890, "cy": 96, "w": 64, "h": 144,
                "transparent": True, "border": "#595959", "border_weight": 1.0,
                "dash": "dash", "rounded": True, "corner_radius": 0.16,
            },
            {"type": "textbox", "x": 872, "y": 10, "w": 38, "h": 18, "text": "N-1", "font_size": 10.5, "bold": False, "italic": True},
            {"type": "textbox", "x": 808, "y": 218, "w": 36, "h": 20, "text": "(c)", "font_size": 12.5, "bold": False},
            {"type": "textbox", "x": 758, "y": 502, "w": 36, "h": 20, "text": "(b)", "font_size": 12.5, "bold": False},
        ],
    }

    # Stage 2: CNN Branch — draw trunks & skip curves FIRST so Conv-BN boxes sit cleanly ON TOP (matching original z-order!)
    s2 = {
        "label": "Draw CNN Branch feature maps, residual curves & Conv-BN blocks",
        "items": [
            {
                "type": "feature_stack", "id": "fm_top", "x": 82, "y": 108, "w": 62, "h": 58,
                "count": 3, "skew_y": -7.0, "dx": 8.0, "dy": 7.0, "fill": "#DCEAF7", "border": "#4F6D8A",
            },
            {"type": "textbox", "x": 102, "y": 83, "w": 24, "h": 18, "text": "W", "font_size": 11, "italic": True},
            {"type": "textbox", "x": 52, "y": 138, "w": 22, "h": 18, "text": "H", "font_size": 11, "italic": True},
            {"type": "textbox", "x": 64, "y": 176, "w": 16, "h": 18, "text": "C", "font_size": 11, "italic": True},
            {"type": "textbox", "x": 156, "y": 86, "w": 105, "h": 22, "text": "Feature Maps", "font_size": 11.5, "align": 1},

            # Trunks & residual curves drawn BEFORE middle/bottom Feature Maps and Conv-BN blocks
            {"type": "arrow", "start": cnn_up_a, "end": cnn_up_b, "color": "#3F6AA8", "weight": 1.6},
            {"type": "junction", "id": "junc_down_in", "on_line": {"start": cnn_up_a, "end": cnn_up_b, "t": 0.64}},
            {"type": "junction", "id": "cnn_u1", "on_line": {"start": cnn_up_a, "end": cnn_up_b, "t": 0.27}},
            {"type": "junction", "id": "cnn_u2", "on_line": {"start": cnn_up_a, "end": cnn_up_b, "t": 0.49}},
            {"type": "junction", "id": "cnn_u3", "on_line": {"start": cnn_up_a, "end": cnn_up_b, "t": 0.76}},
            # Upper blue skip curve: passes cleanly through open gap at x~290 (between 1x1Conv-BN and FCU 1x1Conv) -> top-left of fm_mid!
            {"type": "curve_arrow", "start": cnn_up_a, "c1": [228, 70], "c2": [296, 150], "end": [308, 294], "color": "#3F6AA8", "weight": 1.6},

            {
                "type": "feature_stack", "id": "fm_mid", "x": 310, "y": 284, "w": 62, "h": 56,
                "count": 3, "skew_y": -7.0, "dx": 8.0, "dy": 7.0, "fill": "#DCEAF7", "border": "#4F6D8A",
            },
            {"type": "arrow", "start": cnn_lo_a, "end": cnn_lo_b, "color": "#3F6AA8", "weight": 1.6},
            {"type": "junction", "id": "junc_up_out", "on_line": {"start": cnn_lo_a, "end": cnn_lo_b, "t": 0.43}},
            {"type": "junction", "id": "cnn_l1", "on_line": {"start": cnn_lo_a, "end": cnn_lo_b, "t": 0.27}},
            {"type": "junction", "id": "cnn_l2", "on_line": {"start": cnn_lo_a, "end": cnn_lo_b, "t": 0.57}},
            {"type": "junction", "id": "cnn_l3", "on_line": {"start": cnn_lo_a, "end": cnn_lo_b, "t": 0.79}},
            # Lower blue skip curve: passes just left of BatchNorm (x~514) and behind right edge of 3x3Conv-BN / 1x1Conv-BN -> top-left of fm_bot!
            {"type": "curve_arrow", "start": cnn_lo_a, "c1": [468, 268], "c2": [518, 336], "end": [534, 460], "color": "#3F6AA8", "weight": 1.6},

            {
                "type": "feature_stack", "id": "fm_bot", "x": 536, "y": 452, "w": 58, "h": 54,
                "count": 3, "skew_y": -7.0, "dx": 8.0, "dy": 7.0, "fill": "#DCEAF7", "border": "#4F6D8A",
            },

            # Conv-BN boxes cleanly placed on exact trunk junctions
            {"type": "shear_box", "center": {"node": "cnn_u1"}, "w": 82, "h": 21, "tilt_deg": -11, "skew_x": 0, "text": "1x1Conv-BN", "border": "#4F6D8A", "font_size": 9.5},
            {"type": "shear_box", "center": {"node": "cnn_u2"}, "w": 82, "h": 21, "tilt_deg": -11, "skew_x": 0, "text": "3x3Conv-BN", "border": "#4F6D8A", "font_size": 9.5},
            {"type": "shear_box", "center": {"node": "cnn_u3"}, "w": 82, "h": 21, "tilt_deg": -11, "skew_x": 0, "text": "1x1Conv-BN", "border": "#4F6D8A", "font_size": 9.5},

            {"type": "shear_box", "center": {"node": "cnn_l1"}, "w": 82, "h": 21, "tilt_deg": -11, "skew_x": 0, "text": "1x1Conv-BN", "border": "#4F6D8A", "font_size": 9.5},
            {"type": "shear_box", "center": {"node": "cnn_l2"}, "w": 82, "h": 21, "tilt_deg": -11, "skew_x": 0, "text": "3x3Conv-BN", "border": "#4F6D8A", "font_size": 9.5},
            {"type": "shear_box", "center": {"node": "cnn_l3"}, "w": 82, "h": 21, "tilt_deg": -11, "skew_x": 0, "text": "1x1Conv-BN", "border": "#4F6D8A", "font_size": 9.5},
            {"type": "textbox", "x": 384, "y": 458, "w": 105, "h": 22, "text": "CNN Branch", "font_size": 12, "rotation": 40},
        ],
    }

    # Stage 3: Transformer Branch (gentler -6.5 deg Patch Embeddings + exact junctions + cleared green skip arcs)
    top_prisms = []
    bot_prisms = []
    for k in range(4):
        fill = "#FCE3BE" if k == 3 else "#9AC77E"
        border = "#C99A5B" if k == 3 else "#5B8C4F"
        top_prisms.append({
            "type": "prism3d", "x": 454 + k * 4.0, "y": 134 + k * 7.0, "w": 76, "h": 5.5,
            "skew_y": -6.5, "extrude": [-6, -4], "fill": fill, "border": border, "border_weight": 0.75,
        })
        bot_prisms.append({
            "type": "prism3d", "x": 756 + k * 4.0, "y": 374 + k * 7.0, "w": 76, "h": 5.5,
            "skew_y": -6.5, "extrude": [-6, -4], "fill": fill, "border": border, "border_weight": 0.75,
        })

    s3 = {
        "label": "Draw Transformer Branch 3D embeddings, skip curves & attention blocks",
        "items": [
            {"type": "textbox", "x": 432, "y": 86, "w": 125, "h": 22, "text": "Patch Embeddings", "font_size": 11.5},
            {"type": "textbox", "x": 486, "y": 107, "w": 22, "h": 18, "text": "E", "font_size": 11, "italic": True},
            {"type": "textbox", "x": 410, "y": 146, "w": 36, "h": 18, "text": "K+1", "font_size": 11, "italic": True},
            *top_prisms,
            {"type": "arrow", "start": tr_a, "end": tr_b, "color": "#5B8C4F", "weight": 1.6},
            {"type": "junction", "id": "junc_skip_start", "on_line": {"start": tr_a, "end": tr_b, "t": 0.10}},
            {"type": "junction", "id": "junc_down_out", "on_line": {"start": tr_a, "end": tr_b, "t": 0.14}},
            {"type": "junction", "id": "tr_ln1", "on_line": {"start": tr_a, "end": tr_b, "t": 0.23}},
            {"type": "junction", "id": "tr_mhsa", "on_line": {"start": tr_a, "end": tr_b, "t": 0.39}},
            {"type": "junction", "id": "junc_skip_mid", "on_line": {"start": tr_a, "end": tr_b, "t": 0.46}},
            {"type": "junction", "id": "tr_ln2", "on_line": {"start": tr_a, "end": tr_b, "t": 0.53}},
            {"type": "junction", "id": "tr_mlp", "on_line": {"start": tr_a, "end": tr_b, "t": 0.66}},
            {"type": "junction", "id": "junc_up_in", "on_line": {"start": tr_a, "end": tr_b, "t": 0.79}},
            # Draw green skip curves clearing outside Multi-head Self-attention & MLP and staying below inset (c)
            {"type": "curve_arrow", "start": {"node": "junc_skip_start"}, "c1": [640, 144], "c2": [692, 196], "end": {"node": "junc_skip_mid"}, "color": "#5B8C4F", "weight": 1.6},
            {"type": "curve_arrow", "start": {"node": "junc_skip_mid"}, "c1": [734, 212], "c2": [762, 266], "end": {"node": "junc_up_in"}, "color": "#5B8C4F", "weight": 1.6},
            {"type": "shear_box", "center": {"node": "tr_ln1"}, "w": 72, "h": 20, "tilt_deg": -11, "skew_x": 0, "text": "LayerNorm", "border": "#5B8C4F", "font_size": 9.5},
            {"type": "shear_box", "center": {"node": "tr_mhsa"}, "w": 88, "h": 26, "tilt_deg": -11, "skew_x": 0, "text": "Multi-head\nSelf-attention", "border": "#5B8C4F", "font_size": 8.5},
            {"type": "shear_box", "center": {"node": "tr_ln2"}, "w": 72, "h": 20, "tilt_deg": -11, "skew_x": 0, "text": "LayerNorm", "border": "#5B8C4F", "font_size": 9.5},
            {"type": "shear_box", "center": {"node": "tr_mlp"}, "w": 46, "h": 20, "tilt_deg": -11, "skew_x": 0, "text": "MLP", "border": "#5B8C4F", "font_size": 9.5},
            {"type": "textbox", "x": 728, "y": 254, "w": 132, "h": 22, "text": "Transformer Branch", "font_size": 12, "rotation": 42},
            *bot_prisms,
        ],
    }

    # Stage 4: FCU Coupling Bridges via aligned_chain (auto atan2 tilt_deg, strictly collinear & segmented!)
    s4 = {
        "label": "Draw FCU Down and Up coupling bridges",
        "items": [
            {
                "type": "aligned_chain",
                "start": {"node": "junc_down_in"},
                "end": {"node": "junc_down_out"},
                "color": "#3F6AA8",
                "border": "#4A75A8",
                "first_weight": 2.0,
                "weight": 1.6,
                "tilt_deg": "auto",
                "nodes": [
                    {"id": "d_conv", "t": 0.34, "w": 56, "h": 21, "text": "1x1Conv"},
                    {"id": "d_down", "t": 0.56, "w": 46, "h": 21, "text": "Down", "fill": "#9EC3E6"},
                    {"id": "d_ln", "t": 0.80, "w": 70, "h": 21, "text": "LayerNorm"},
                ],
            },
            {
                "type": "aligned_chain",
                "start": {"node": "junc_up_in"},
                "end": {"node": "junc_up_out"},
                "color": "#5B8C4F",
                "border": "#5B8C4F",
                "first_weight": 2.1,
                "weight": 1.6,
                "tilt_deg": "auto",
                "nodes": [
                    {"id": "u_up", "t": 0.13, "w": 38, "h": 21, "text": "Up", "fill": "#84C475", "border": "#4E8A42"},
                    {"id": "u_conv", "t": 0.34, "w": 54, "h": 21, "text": "1x1Conv"},
                    {"id": "u_bn", "t": 0.57, "w": 64, "h": 21, "text": "BatchNorm"},
                ],
            },
        ],
    }

    # Stage 5: Inset (a)
    a_items = []
    for idx in range(4):
        a_items.append({
            "type": "prism3d", "x": 96, "y": 426 + idx * 16, "w": 11.0, "h": 11.0,
            "extrude": [-24, -19], "fill": "#9AC77E", "border": "#4E7A42", "border_weight": 0.75,
        })
    a_items.append({
        "type": "grid3d", "x": 182, "y": 442, "rows": 2, "cols": 2, "depth": 1,
        "cell_w": 10.5, "cell_h": 10.5, "gap_x": 4.0, "gap_y": 4.0,
        "extrude": [-24, -19], "fill": "#A6C3DE", "border": "#4A6580", "border_weight": 0.75,
    })
    a_items.append({
        "type": "grid3d", "x": 288, "y": 424, "rows": 4, "cols": 4, "depth": 1,
        "cell_w": 10.0, "cell_h": 10.0, "gap_x": 3.8, "gap_y": 3.8,
        "extrude": [-24, -19], "fill": "#A6C3DE", "border": "#4A6580", "border_weight": 0.75,
    })
    a_items.extend([
        {"type": "textbox", "x": 112, "y": 396, "w": 56, "h": 18, "text": "Reshaping", "font_size": 9.0, "bold": False},
        {"type": "arrow", "start": [152, 426], "end": [120, 426], "color": "#3F6AA8", "weight": 1.6},
        {"type": "textbox", "x": 196, "y": 396, "w": 62, "h": 18, "text": "Avgpooling", "font_size": 9.0, "bold": False},
        {"type": "arrow", "start": [254, 426], "end": [220, 426], "color": "#3F6AA8", "weight": 1.6},
        {"type": "arrow", "start": [120, 464], "end": [152, 464], "color": "#5B8C4F", "weight": 1.6},
        {"type": "textbox", "x": 112, "y": 478, "w": 56, "h": 18, "text": "Reshaping", "font_size": 9.0, "bold": False},
        {"type": "arrow", "start": [220, 464], "end": [254, 464], "color": "#5B8C4F", "weight": 1.6},
        {"type": "textbox", "x": 204, "y": 478, "w": 66, "h": 18, "text": "Interpolating", "font_size": 9.0, "bold": False},
    ])
    s5 = {"label": "Draw inset (a) 3D patch/spatial transforms", "items": a_items}

    # Stage 6: Inset (c)
    s6 = {
        "label": "Draw inset (c) overall Conformer pipeline",
        "items": [
            {"type": "image", "path": str(PEACOCK_PATH), "x": 666, "y": 88, "w": 25, "h": 58},
            {"type": "arrow", "start": [691, 117], "end": [704, 117], "color": "#262626", "weight": 1.3},
            {"type": "box", "cx": 715, "cy": 117, "w": 66, "h": 22, "text": "Stem", "fill": "#FFFFFF", "border": "#7F7F7F",
             "font_size": 9.5, "rotation": -90, "rounded": True, "corner_radius": 0.35, "bold": False, "border_weight": 1.0},
            {"type": "arrow", "start": [726, 117], "end": [740, 117], "color": "#262626", "weight": 1.2, "arrow": False},
            {"type": "arrow", "start": [740, 52], "end": [740, 148], "color": "#262626", "weight": 1.2, "arrow": False},
            {"type": "arrow", "start": [740, 52], "end": [756, 52], "color": "#262626", "weight": 1.2},
            {"type": "arrow", "start": [740, 148], "end": [794, 148], "color": "#262626", "weight": 1.2},
            {"type": "box", "cx": 768, "cy": 52, "w": 54, "h": 22, "text": "Project", "fill": "#9AC77E", "border": "#5B8C4F",
             "font_size": 9.5, "rotation": -90, "rounded": True, "corner_radius": 0.35, "bold": False, "border_weight": 1.0},
            {"type": "arrow", "start": [780, 52], "end": [794, 52], "color": "#5B8C4F", "weight": 1.3},

            {"type": "prism3d", "id": "tb1", "cx": 815, "cy": 54, "w": 36, "h": 36, "extrude": [7, -5],
             "fill": "#C8E0B4", "border": "#5B8C4F", "border_weight": 1.0, "text": "Trans\nBlock", "font_size": 8.5},
            {"type": "arrow", "start": [840, 52], "end": [868, 52], "color": "#5B8C4F", "weight": 1.3, "dash": "dash"},
            {"type": "prism3d", "id": "tb2", "cx": 890, "cy": 54, "w": 36, "h": 36, "extrude": [7, -5],
             "fill": "#C8E0B4", "border": "#5B8C4F", "border_weight": 1.0, "text": "Trans\nBlock", "font_size": 8.5},

            {"type": "prism3d", "id": "cb1", "cx": 815, "cy": 150, "w": 36, "h": 36, "extrude": [7, -5],
             "fill": "#BDD7EE", "border": "#4A6580", "border_weight": 1.0, "text": "Conv\nBlock", "font_size": 8.5},
            {"type": "arrow", "start": [840, 148], "end": [868, 148], "color": "#3F6AA8", "weight": 1.3, "dash": "dash"},
            {"type": "prism3d", "id": "cb2", "cx": 890, "cy": 150, "w": 36, "h": 36, "extrude": [7, -5],
             "fill": "#BDD7EE", "border": "#4A6580", "border_weight": 1.0, "text": "Conv\nBlock", "font_size": 8.5},

            {"type": "arrow", "start": [884, 127], "end": [884, 76], "color": "#5B8C4F", "weight": 1.5},
            {"type": "arrow", "start": [901, 76], "end": [901, 127], "color": "#3F6AA8", "weight": 1.5},
            {"type": "brace", "cx": 854, "cy": 178, "w": 12, "h": 116, "rotation": 90, "color": "#595959", "weight": 1.0},
            {"type": "textbox", "x": 844, "y": 190, "w": 22, "h": 18, "text": "N", "font_size": 10.5, "bold": False, "italic": True},

            {"type": "arrow", "start": [915, 52], "end": [958, 52], "color": "#5B8C4F", "weight": 1.3, "dash": "dash"},
            {"type": "textbox", "cx": 935, "cy": 52, "w": 56, "h": 14, "text": "Class token", "font_size": 8.0, "bold": False, "rotation": 90},
            {"type": "box", "cx": 970, "cy": 52, "w": 58, "h": 22, "text": "Classifier", "fill": "#E2EFDA", "border": "#5B8C4F",
             "font_size": 9.5, "rotation": -90, "rounded": True, "corner_radius": 0.35, "bold": False, "border_weight": 1.0},

            {"type": "arrow", "start": [915, 148], "end": [958, 148], "color": "#3F6AA8", "weight": 1.3, "dash": "dash"},
            {"type": "textbox", "cx": 935, "cy": 148, "w": 46, "h": 14, "text": "Pooling", "font_size": 8.0, "bold": False, "rotation": 90},
            {"type": "box", "cx": 970, "cy": 148, "w": 58, "h": 22, "text": "Classifier", "fill": "#DDEBF7", "border": "#4A6580",
             "font_size": 9.5, "rotation": -90, "rounded": True, "corner_radius": 0.35, "bold": False, "border_weight": 1.0},
        ],
    }

    return coord_space, [s1, s2, s3, s4, s5, s6]


async def main():
    out_path = ROOT / "data" / f"conformer_v7_{uuid.uuid4().hex[:6]}.pptx"
    preview_path = ROOT / "data" / f"conformer_v7_preview.png"
    if preview_path.exists():
        preview_path.unlink()
    coord_space, stages = build_conformer_scene()
    params = StdioServerParameters(
        command=sys.executable,
        args=["-X", "utf8", "-u", "-m", "microsoft_wps_live_mcp"],
    )
    t0 = time.monotonic()
    async with Client(params, read_timeout_seconds=120) as client:
        created = unpack(await client.call_tool("office_create_document", {
            "path": str(out_path), "engine": "microsoft", "visible": True,
        }))
        doc_id = created["document_id"]
        await client.call_tool("office_show_document", {"document": doc_id, "focus": {"slide": 1}})

        async def on_progress(num, tot, msg):
            print(f"  -> Stage {int(num)}/{int(tot)}: {msg}", flush=True)

        t_draw = time.monotonic()
        res = unpack(await client.call_tool("office_ppt_draw_scene", {
            "document": doc_id,
            "stages": stages,
            "coord_space": coord_space,
            "slide_size": [960, 540],
            "pause_ms": 150,
            "item_pause_ms": 12,
        }, progress_callback=on_progress))
        draw_elapsed = time.monotonic() - t_draw
        print(f"[+] Draw finished in {draw_elapsed:.2f}s: items={res['total_items_drawn']}, saved={res['disk_saved']}")
        print(f"[+] Layout warnings ({len(res.get('layout_warnings', []))}): {res.get('layout_warnings', [])}")

        exp = unpack(await client.call_tool("office_ppt_export_image", {
            "document": doc_id, "path": str(preview_path), "slide": 1, "width": 1920, "height": 1080,
        }))
        print(f"[+] Exported preview: {exp['output']} ({exp['size_bytes']} bytes)")
        copied = unpack(await client.call_tool("office_create_copy", {"document": doc_id}))
        print(f"[+] Final copy: {copied['copy_path']}")
        await client.call_tool("office_close_document", {"document": doc_id})
    print(f"[+] Total elapsed: {time.monotonic() - t0:.2f}s")


if __name__ == "__main__":
    asyncio.run(main())
