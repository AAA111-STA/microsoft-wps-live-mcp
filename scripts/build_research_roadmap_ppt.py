"""Build the research roadmap presentation in Microsoft PowerPoint via microsoft-wps-live-mcp."""
import asyncio
import json
from pathlib import Path
import sys

from mcp import Client, StdioServerParameters

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = ROOT / "data" / "research_roadmap.pptx"

def to_bgr(hex_str: str) -> int:
    h = hex_str.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r | (g << 8) | (b << 16)

def unpack(res):
    if res.is_error:
        raise RuntimeError(str(res.content))
    return res.structured_content or json.loads(res.content[0].text)

async def main():
    if OUTPUT_PATH.exists():
        try:
            OUTPUT_PATH.unlink()
        except Exception:
            pass

    server_params = StdioServerParameters(command=sys.executable, args=["-m", "microsoft_wps_live_mcp"])
    print("[*] Connecting to microsoft-wps-live-mcp server...")
    
    async with Client(server_params, read_timeout_seconds=180) as client:
        print("[*] Calling office_create_document (engine='microsoft', visible=True)...")
        res = unpack(await client.call_tool("office_create_document", {
            "path": str(OUTPUT_PATH),
            "engine": "microsoft",
            "visible": True
        }))
        doc_id = res["document_id"]
        print(f"[+] Created native PowerPoint presentation! Document ID: {doc_id}")

        # Add a blank slide (layout 12 is blank)
        print("[*] Adding blank slide...")
        await client.call_tool("office_ppt_add_slide", {
            "document": doc_id,
            "title": "",
            "position": 1
        })

        # Helper to run batch steps on slide 1
        async def run_batch(steps):
            full_steps = [
                {"action": "call", "target": "document.Slides", "member": "Item", "args": [1], "as": "slide"}
            ] + steps
            return unpack(await client.call_tool("office_api_batch", {
                "document": doc_id,
                "steps": full_steps
            }))

        # -------------------------------------------------------------
        # STAGE 1: Outer Container Boxes & Subtitle Labels
        # -------------------------------------------------------------
        print("[*] Stage 1: Creating main modular containers & subtitles...")
        stage1_steps = [
            # Top-left: (d) Graph Refiner
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 30, 30, 230, 195], "as": "box_d"},
            {"action": "set", "target": "box_d.Fill", "member": "Visible", "value": True},
            {"action": "call", "target": "box_d.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "box_d.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FFFFFF")},
            {"action": "set", "target": "box_d.Line", "member": "Weight", "value": 1.5},
            {"action": "set", "target": "box_d.Line.ForeColor", "member": "RGB", "value": to_bgr("#333333")},
            
            # Label (d) title
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 35, 32, 220, 25], "as": "txt_d"},
            {"action": "set", "target": "txt_d.TextFrame.TextRange", "member": "Text", "value": "(d)   Graph Refiner"},
            {"action": "set", "target": "txt_d.TextFrame.TextRange.Font", "member": "Name", "value": "Segoe UI"},
            {"action": "set", "target": "txt_d.TextFrame.TextRange.Font", "member": "Size", "value": 13},
            {"action": "set", "target": "txt_d.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "txt_d.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # Red subtitle for (d)
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 35, 12, 220, 20], "as": "sub_d"},
            {"action": "set", "target": "sub_d.TextFrame.TextRange", "member": "Text", "value": "3.  稀疏化邻接矩阵融合邻接节点特征"},
            {"action": "set", "target": "sub_d.TextFrame.TextRange.Font", "member": "Name", "value": "Microsoft YaHei"},
            {"action": "set", "target": "sub_d.TextFrame.TextRange.Font", "member": "Size", "value": 8},
            {"action": "set", "target": "sub_d.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#C00000")},

            # Bottom-left: (a) Variational Generator
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 30, 240, 230, 175], "as": "box_a"},
            {"action": "set", "target": "box_a.Fill", "member": "Visible", "value": True},
            {"action": "call", "target": "box_a.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "box_a.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FFFFFF")},
            {"action": "set", "target": "box_a.Line", "member": "Weight", "value": 1.5},
            {"action": "set", "target": "box_a.Line.ForeColor", "member": "RGB", "value": to_bgr("#333333")},
            
            # Label (a) title
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 35, 242, 220, 25], "as": "txt_a"},
            {"action": "set", "target": "txt_a.TextFrame.TextRange", "member": "Text", "value": "(a) Variational Generator"},
            {"action": "set", "target": "txt_a.TextFrame.TextRange.Font", "member": "Name", "value": "Segoe UI"},
            {"action": "set", "target": "txt_a.TextFrame.TextRange.Font", "member": "Size", "value": 13},
            {"action": "set", "target": "txt_a.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "txt_a.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # Red subtitle for (a)
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 35, 225, 220, 20], "as": "sub_a"},
            {"action": "set", "target": "sub_a.TextFrame.TextRange", "member": "Text", "value": "1.  VAE进行粗略预测"},
            {"action": "set", "target": "sub_a.TextFrame.TextRange.Font", "member": "Name", "value": "Microsoft YaHei"},
            {"action": "set", "target": "sub_a.TextFrame.TextRange.Font", "member": "Size", "value": 8},
            {"action": "set", "target": "sub_a.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#C00000")},

            # Center-right: (b) Graph Structure Aligner
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 475, 30, 290, 385], "as": "box_b"},
            {"action": "set", "target": "box_b.Fill", "member": "Visible", "value": True},
            {"action": "call", "target": "box_b.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "box_b.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FFFFFF")},
            {"action": "set", "target": "box_b.Line", "member": "Weight", "value": 1.5},
            {"action": "set", "target": "box_b.Line.ForeColor", "member": "RGB", "value": to_bgr("#333333")},
            
            # Label (b) title
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 480, 32, 280, 25], "as": "txt_b"},
            {"action": "set", "target": "txt_b.TextFrame.TextRange", "member": "Text", "value": "(b) Graph Structure Aligner"},
            {"action": "set", "target": "txt_b.TextFrame.TextRange.Font", "member": "Name", "value": "Segoe UI"},
            {"action": "set", "target": "txt_b.TextFrame.TextRange.Font", "member": "Size", "value": 13},
            {"action": "set", "target": "txt_b.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "txt_b.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # Red subtitle for (b)
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 480, 12, 280, 20], "as": "sub_b"},
            {"action": "set", "target": "sub_b.TextFrame.TextRange", "member": "Text", "value": "2.  利用Graph VAE生成patch之间的邻接矩阵"},
            {"action": "set", "target": "sub_b.TextFrame.TextRange.Font", "member": "Name", "value": "Microsoft YaHei"},
            {"action": "set", "target": "sub_b.TextFrame.TextRange.Font", "member": "Size", "value": 8},
            {"action": "set", "target": "sub_b.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#C00000")},

            # Far-right: (c) Graph VAE
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 780, 30, 150, 385], "as": "box_c"},
            {"action": "set", "target": "box_c.Fill", "member": "Visible", "value": True},
            {"action": "call", "target": "box_c.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "box_c.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FFFFFF")},
            {"action": "set", "target": "box_c.Line", "member": "Weight", "value": 1.5},
            {"action": "set", "target": "box_c.Line.ForeColor", "member": "RGB", "value": to_bgr("#333333")},
            
            # Label (c) title
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 785, 32, 140, 25], "as": "txt_c"},
            {"action": "set", "target": "txt_c.TextFrame.TextRange", "member": "Text", "value": "(c) Graph VAE"},
            {"action": "set", "target": "txt_c.TextFrame.TextRange.Font", "member": "Name", "value": "Segoe UI"},
            {"action": "set", "target": "txt_c.TextFrame.TextRange.Font", "member": "Size", "value": 13},
            {"action": "set", "target": "txt_c.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "txt_c.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # Bottom Legend Panel Container
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 30, 428, 900, 95], "as": "box_legend"},
            {"action": "set", "target": "box_legend.Fill", "member": "Visible", "value": True},
            {"action": "call", "target": "box_legend.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "box_legend.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FDFDFD")},
            {"action": "set", "target": "box_legend.Line", "member": "Weight", "value": 1.0},
            {"action": "set", "target": "box_legend.Line", "member": "DashStyle", "value": 4}, # msoLineDash
            {"action": "set", "target": "box_legend.Line.ForeColor", "member": "RGB", "value": to_bgr("#888888")},
        ]
        await run_batch(stage1_steps)
        await client.call_tool("office_save_document", {"document": doc_id})
        print("[+] Stage 1 completed & saved.")

        # -------------------------------------------------------------
        # STAGE 2: Central Pipeline Architecture
        # -------------------------------------------------------------
        print("[*] Stage 2: Building Center Main Pipeline...")
        stage2_steps = [
            # Top formula title: \hat{Y}^{endo}
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 290, 10, 160, 28], "as": "txt_yhat"},
            {"action": "set", "target": "txt_yhat.TextFrame.TextRange", "member": "Text", "value": "Ŷ endo"},
            {"action": "set", "target": "txt_yhat.TextFrame.TextRange.Font", "member": "Name", "value": "Cambria Math"},
            {"action": "set", "target": "txt_yhat.TextFrame.TextRange.Font", "member": "Size", "value": 15},
            {"action": "set", "target": "txt_yhat.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "txt_yhat.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # Red subtitle 4
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 280, 36, 180, 20], "as": "sub_pipe"},
            {"action": "set", "target": "sub_pipe.TextFrame.TextRange", "member": "Text", "value": "4.  展开所有patch特征线性预测"},
            {"action": "set", "target": "sub_pipe.TextFrame.TextRange.Font", "member": "Name", "value": "Microsoft YaHei"},
            {"action": "set", "target": "sub_pipe.TextFrame.TextRange.Font", "member": "Size", "value": 8},
            {"action": "set", "target": "sub_pipe.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#C00000")},
            {"action": "set", "target": "sub_pipe.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # 1. Flatten Head (top box)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 305, 58, 140, 32], "as": "p_flatten"},
            {"action": "call", "target": "p_flatten.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "p_flatten.Fill.ForeColor", "member": "RGB", "value": to_bgr("#E9ECEF")},
            {"action": "set", "target": "p_flatten.Line.ForeColor", "member": "RGB", "value": to_bgr("#ADB5BD")},
            {"action": "set", "target": "p_flatten.TextFrame.TextRange", "member": "Text", "value": "Flatten Head"},
            {"action": "set", "target": "p_flatten.TextFrame.TextRange.Font", "member": "Name", "value": "Segoe UI"},
            {"action": "set", "target": "p_flatten.TextFrame.TextRange.Font", "member": "Size", "value": 11},
            {"action": "set", "target": "p_flatten.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "p_flatten.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#212529")},

            # Label H
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 385, 92, 40, 20], "as": "txt_h"},
            {"action": "set", "target": "txt_h.TextFrame.TextRange", "member": "Text", "value": "H"},
            {"action": "set", "target": "txt_h.TextFrame.TextRange.Font", "member": "Name", "value": "Cambria Math"},
            {"action": "set", "target": "txt_h.TextFrame.TextRange.Font", "member": "Size", "value": 12},
            {"action": "set", "target": "txt_h.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # 2. Graph Refiner (gold box)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 305, 115, 140, 32], "as": "p_refiner"},
            {"action": "call", "target": "p_refiner.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "p_refiner.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FFE599")},
            {"action": "set", "target": "p_refiner.Line.ForeColor", "member": "RGB", "value": to_bgr("#D6B656")},
            {"action": "set", "target": "p_refiner.TextFrame.TextRange", "member": "Text", "value": "Graph Refiner"},
            {"action": "set", "target": "p_refiner.TextFrame.TextRange.Font", "member": "Name", "value": "Segoe UI"},
            {"action": "set", "target": "p_refiner.TextFrame.TextRange.Font", "member": "Size", "value": 11},
            {"action": "set", "target": "p_refiner.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "p_refiner.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#5A4A00")},

            # Label (\tilde{S}^P, \hat{A})
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 355, 148, 80, 22], "as": "txt_spa"},
            {"action": "set", "target": "txt_spa.TextFrame.TextRange", "member": "Text", "value": "(S̃ᴾ, Â)"},
            {"action": "set", "target": "txt_spa.TextFrame.TextRange.Font", "member": "Name", "value": "Segoe UI"},
            {"action": "set", "target": "txt_spa.TextFrame.TextRange.Font", "member": "Size", "value": 11},
            {"action": "set", "target": "txt_spa.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # 3. Graph Structure Aligner (orange box)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 305, 172, 140, 42], "as": "p_aligner"},
            {"action": "call", "target": "p_aligner.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "p_aligner.Fill.ForeColor", "member": "RGB", "value": to_bgr("#F4B183")},
            {"action": "set", "target": "p_aligner.Line.ForeColor", "member": "RGB", "value": to_bgr("#C55A11")},
            {"action": "set", "target": "p_aligner.TextFrame.TextRange", "member": "Text", "value": "Graph\nStructure Aligner"},
            {"action": "set", "target": "p_aligner.TextFrame.TextRange.Font", "member": "Name", "value": "Segoe UI"},
            {"action": "set", "target": "p_aligner.TextFrame.TextRange.Font", "member": "Size", "value": 10},
            {"action": "set", "target": "p_aligner.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "p_aligner.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#FFFFFF")},
            {"action": "set", "target": "p_aligner.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # Dual input labels: \tilde{S} and S
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 325, 216, 35, 20], "as": "txt_stilde"},
            {"action": "set", "target": "txt_stilde.TextFrame.TextRange", "member": "Text", "value": "S̃"},
            {"action": "set", "target": "txt_stilde.TextFrame.TextRange.Font", "member": "Name", "value": "Cambria Math"},
            {"action": "set", "target": "txt_stilde.TextFrame.TextRange.Font", "member": "Size", "value": 13},
            {"action": "set", "target": "txt_stilde.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 395, 216, 35, 20], "as": "txt_s"},
            {"action": "set", "target": "txt_s.TextFrame.TextRange", "member": "Text", "value": "S"},
            {"action": "set", "target": "txt_s.TextFrame.TextRange.Font", "member": "Name", "value": "Cambria Math"},
            {"action": "set", "target": "txt_s.TextFrame.TextRange.Font", "member": "Size", "value": 13},
            {"action": "set", "target": "txt_s.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # 4. Variational Generator (peach box)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 305, 238, 140, 42], "as": "p_var"},
            {"action": "call", "target": "p_var.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "p_var.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FCE4D6")},
            {"action": "set", "target": "p_var.Line.ForeColor", "member": "RGB", "value": to_bgr("#F4B183")},
            {"action": "set", "target": "p_var.TextFrame.TextRange", "member": "Text", "value": "Variational\nGenerator"},
            {"action": "set", "target": "p_var.TextFrame.TextRange.Font", "member": "Name", "value": "Segoe UI"},
            {"action": "set", "target": "p_var.TextFrame.TextRange.Font", "member": "Size", "value": 10},
            {"action": "set", "target": "p_var.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "p_var.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#843C0C")},
            {"action": "set", "target": "p_var.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # 5. Bottom 2x2 Matrix Input Cards (S: X^exo, Y^exo, X^endo, Y^endo)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 300, 310, 150, 75], "as": "card_input_bg"},
            {"action": "call", "target": "card_input_bg.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "card_input_bg.Fill.ForeColor", "member": "RGB", "value": to_bgr("#F8F9FA")},
            {"action": "set", "target": "card_input_bg.Line.ForeColor", "member": "RGB", "value": to_bgr("#CCCCCC")},

            # Sub-card: X^exo (top-left, green)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 306, 316, 65, 30], "as": "c_xexo"},
            {"action": "call", "target": "c_xexo.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "c_xexo.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "c_xexo.Line.ForeColor", "member": "RGB", "value": to_bgr("#70AD47")},
            {"action": "set", "target": "c_xexo.TextFrame.TextRange", "member": "Text", "value": "X exo"},
            {"action": "set", "target": "c_xexo.TextFrame.TextRange.Font", "member": "Size", "value": 10},
            {"action": "set", "target": "c_xexo.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # Sub-card: Y^exo (top-right, green)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 377, 316, 65, 30], "as": "c_yexo"},
            {"action": "call", "target": "c_yexo.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "c_yexo.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "c_yexo.Line.ForeColor", "member": "RGB", "value": to_bgr("#70AD47")},
            {"action": "set", "target": "c_yexo.TextFrame.TextRange", "member": "Text", "value": "Y exo"},
            {"action": "set", "target": "c_yexo.TextFrame.TextRange.Font", "member": "Size", "value": 10},
            {"action": "set", "target": "c_yexo.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # Sub-card: X^endo (bottom-left, blue)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 306, 350, 65, 30], "as": "c_xendo"},
            {"action": "call", "target": "c_xexo.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "c_xendo.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "c_xendo.Line.ForeColor", "member": "RGB", "value": to_bgr("#4472C4")},
            {"action": "set", "target": "c_xendo.TextFrame.TextRange", "member": "Text", "value": "X endo"},
            {"action": "set", "target": "c_xendo.TextFrame.TextRange.Font", "member": "Size", "value": 10},
            {"action": "set", "target": "c_xendo.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # Sub-card: Y^endo (bottom-right, blue)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 377, 350, 65, 30], "as": "c_yendo"},
            {"action": "call", "target": "c_yendo.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "c_yendo.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "c_yendo.Line.ForeColor", "member": "RGB", "value": to_bgr("#4472C4")},
            {"action": "set", "target": "c_yendo.TextFrame.TextRange", "member": "Text", "value": "Y endo"},
            {"action": "set", "target": "c_yendo.TextFrame.TextRange.Font", "member": "Size", "value": 10},
            {"action": "set", "target": "c_yendo.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # Connectors in pipeline
            {"action": "call", "target": "slide.Shapes", "member": "AddConnector", "args": [1, 375, 115, 375, 90], "as": "con1"},
            {"action": "call", "target": "slide.Shapes", "member": "AddConnector", "args": [1, 375, 172, 375, 147], "as": "con2"},
            {"action": "call", "target": "slide.Shapes", "member": "AddConnector", "args": [1, 375, 58, 375, 36], "as": "con0"},
        ]
        await run_batch(stage2_steps)
        await client.call_tool("office_save_document", {"document": doc_id})
        print("[+] Stage 2 completed & saved.")

        # -------------------------------------------------------------
        # STAGE 3: Submodule (d) Graph Refiner Details (Top-Left)
        # -------------------------------------------------------------
        print("[*] Stage 3: Building (d) Graph Refiner Stack & Sparse Graph...")
        stage3_steps = [
            # Top output H
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 140, 52, 30, 20], "as": "d_h"},
            {"action": "set", "target": "d_h.TextFrame.TextRange", "member": "Text", "value": "H"},
            {"action": "set", "target": "d_h.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "d_h.TextFrame.TextRange.Font", "member": "Size", "value": 12},

            # L x text
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 38, 75, 30, 25], "as": "txt_lx"},
            {"action": "set", "target": "txt_lx.TextFrame.TextRange", "member": "Text", "value": "L ×"},
            {"action": "set", "target": "txt_lx.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "txt_lx.TextFrame.TextRange.Font", "member": "Size", "value": 12},

            # Stacked GCN layers (Layer 3 background, Layer 2, Layer 1 foreground)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 75, 70, 165, 26], "as": "gcn3"},
            {"action": "call", "target": "gcn3.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "gcn3.Fill.ForeColor", "member": "RGB", "value": to_bgr("#C3E6CB")},
            {"action": "set", "target": "gcn3.Line.ForeColor", "member": "RGB", "value": to_bgr("#70AD47")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 72, 74, 165, 26], "as": "gcn2"},
            {"action": "call", "target": "gcn2.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "gcn2.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "gcn2.Line.ForeColor", "member": "RGB", "value": to_bgr("#70AD47")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 69, 78, 165, 26], "as": "gcn1"},
            {"action": "call", "target": "gcn1.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "gcn1.Fill.ForeColor", "member": "RGB", "value": to_bgr("#E2F0D9")},
            {"action": "set", "target": "gcn1.Line.ForeColor", "member": "RGB", "value": to_bgr("#70AD47")},
            {"action": "set", "target": "gcn1.TextFrame.TextRange", "member": "Text", "value": "GCN"},
            {"action": "set", "target": "gcn1.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "gcn1.TextFrame.TextRange.Font", "member": "Size", "value": 11},
            {"action": "set", "target": "gcn1.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # Sparsify box
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 69, 114, 165, 24], "as": "b_sparsify"},
            {"action": "call", "target": "b_sparsify.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_sparsify.Fill.ForeColor", "member": "RGB", "value": to_bgr("#DCE4EC")},
            {"action": "set", "target": "b_sparsify.Line.ForeColor", "member": "RGB", "value": to_bgr("#95A5A6")},
            {"action": "set", "target": "b_sparsify.TextFrame.TextRange", "member": "Text", "value": "Sparsify"},
            {"action": "set", "target": "b_sparsify.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "b_sparsify.TextFrame.TextRange.Font", "member": "Size", "value": 10},
            {"action": "set", "target": "b_sparsify.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # Graph container (dashed)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 45, 145, 200, 68], "as": "b_graph"},
            {"action": "call", "target": "b_graph.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_graph.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FAFAFA")},
            {"action": "set", "target": "b_graph.Line", "member": "Weight", "value": 1.0},
            {"action": "set", "target": "b_graph.Line", "member": "DashStyle", "value": 4},
            {"action": "set", "target": "b_graph.Line.ForeColor", "member": "RGB", "value": to_bgr("#888888")},

            # Graph label
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 48, 146, 70, 20], "as": "txt_graph_lbl"},
            {"action": "set", "target": "txt_graph_lbl.TextFrame.TextRange", "member": "Text", "value": "(S̃ᴾ, Â)"},
            {"action": "set", "target": "txt_graph_lbl.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "txt_graph_lbl.TextFrame.TextRange.Font", "member": "Size", "value": 9},

            # Graph nodes: row 1 (green)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 120, 155, 18, 14], "as": "gn1"},
            {"action": "call", "target": "gn1.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "gn1.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "gn1.Line.ForeColor", "member": "RGB", "value": to_bgr("#70AD47")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 155, 155, 18, 14], "as": "gn2"},
            {"action": "call", "target": "gn2.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "gn2.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "gn2.Line.ForeColor", "member": "RGB", "value": to_bgr("#70AD47")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 190, 155, 18, 14], "as": "gn3"},
            {"action": "call", "target": "gn3.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "gn3.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "gn3.Line.ForeColor", "member": "RGB", "value": to_bgr("#70AD47")},

            # Graph nodes: row 2 (blue)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 100, 182, 18, 14], "as": "gn4"},
            {"action": "call", "target": "gn4.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "gn4.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "gn4.Line.ForeColor", "member": "RGB", "value": to_bgr("#4472C4")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 140, 182, 18, 14], "as": "gn5"},
            {"action": "call", "target": "gn5.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "gn5.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "gn5.Line.ForeColor", "member": "RGB", "value": to_bgr("#4472C4")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 175, 182, 18, 14], "as": "gn6"},
            {"action": "call", "target": "gn6.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "gn6.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "gn6.Line.ForeColor", "member": "RGB", "value": to_bgr("#4472C4")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 210, 182, 18, 14], "as": "gn7"},
            {"action": "call", "target": "gn7.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "gn7.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "gn7.Line.ForeColor", "member": "RGB", "value": to_bgr("#4472C4")},
        ]
        await run_batch(stage3_steps)
        await client.call_tool("office_save_document", {"document": doc_id})
        print("[+] Stage 3 completed & saved.")

        # -------------------------------------------------------------
        # STAGE 4: Submodule (a) Variational Generator Details
        # -------------------------------------------------------------
        print("[*] Stage 4: Building (a) Variational Generator Pipeline...")
        stage4_steps = [
            # Input waveform cards
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 45, 280, 52, 34], "as": "va_xexo"},
            {"action": "call", "target": "va_xexo.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "va_xexo.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "va_xexo.Line.ForeColor", "member": "RGB", "value": to_bgr("#70AD47")},
            {"action": "set", "target": "va_xexo.TextFrame.TextRange", "member": "Text", "value": "X exo"},
            {"action": "set", "target": "va_xexo.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "va_xexo.TextFrame.TextRange.Font", "member": "Size", "value": 9},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 45, 322, 52, 34], "as": "va_xendo"},
            {"action": "call", "target": "va_xendo.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "va_xendo.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "va_xendo.Line.ForeColor", "member": "RGB", "value": to_bgr("#4472C4")},
            {"action": "set", "target": "va_xendo.TextFrame.TextRange", "member": "Text", "value": "X endo"},
            {"action": "set", "target": "va_xendo.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "va_xendo.TextFrame.TextRange.Font", "member": "Size", "value": 9},

            # Vertical bar 1: Instance Norm
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 115, 275, 24, 82], "as": "bar_norm"},
            {"action": "call", "target": "bar_norm.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "bar_norm.Fill.ForeColor", "member": "RGB", "value": to_bgr("#F2F2F2")},
            {"action": "set", "target": "bar_norm.Line.ForeColor", "member": "RGB", "value": to_bgr("#BBBBBB")},
            {"action": "set", "target": "bar_norm.TextFrame.TextRange", "member": "Text", "value": "Instance\nNorm"},
            {"action": "set", "target": "bar_norm.TextFrame.TextRange.Font", "member": "Size", "value": 8},
            {"action": "set", "target": "bar_norm.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # Vertical bar 2: VAE (Pink)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 150, 275, 26, 82], "as": "bar_vae"},
            {"action": "call", "target": "bar_vae.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "bar_vae.Fill.ForeColor", "member": "RGB", "value": to_bgr("#F8CBAD")},
            {"action": "set", "target": "bar_vae.Line.ForeColor", "member": "RGB", "value": to_bgr("#C55A11")},
            {"action": "set", "target": "bar_vae.TextFrame.TextRange", "member": "Text", "value": "VAE"},
            {"action": "set", "target": "bar_vae.TextFrame.TextRange.Font", "member": "Size", "value": 9},
            {"action": "set", "target": "bar_vae.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # Vertical bar 3: Concat
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 188, 275, 26, 82], "as": "bar_concat"},
            {"action": "call", "target": "bar_concat.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "bar_concat.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "bar_concat.Line.ForeColor", "member": "RGB", "value": to_bgr("#4472C4")},
            {"action": "set", "target": "bar_concat.TextFrame.TextRange", "member": "Text", "value": "Concat"},
            {"action": "set", "target": "bar_concat.TextFrame.TextRange.Font", "member": "Size", "value": 8},
            {"action": "set", "target": "bar_concat.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # Output label S̃
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 222, 305, 30, 25], "as": "txt_va_out"},
            {"action": "set", "target": "txt_va_out.TextFrame.TextRange", "member": "Text", "value": "S̃"},
            {"action": "set", "target": "txt_va_out.TextFrame.TextRange.Font", "member": "Name", "value": "Cambria Math"},
            {"action": "set", "target": "txt_va_out.TextFrame.TextRange.Font", "member": "Size", "value": 14},
            {"action": "set", "target": "txt_va_out.TextFrame.TextRange.Font", "member": "Bold", "value": True},
        ]
        await run_batch(stage4_steps)
        await client.call_tool("office_save_document", {"document": doc_id})
        print("[+] Stage 4 completed & saved.")

        # -------------------------------------------------------------
        # STAGE 5: Submodule (b) Graph Structure Aligner & Heatmaps
        # -------------------------------------------------------------
        print("[*] Stage 5: Building (b) Structure Aligner Heatmaps & Loss...")
        stage5_steps = [
            # Heatmap A (Left)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 500, 68, 95, 95], "as": "hm_a"},
            {"action": "call", "target": "hm_a.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "hm_a.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FFF2CC")},
            {"action": "set", "target": "hm_a.Line", "member": "Weight", "value": 1.0},
            {"action": "set", "target": "hm_a.Line", "member": "DashStyle", "value": 4},
            {"action": "set", "target": "hm_a.Line.ForeColor", "member": "RGB", "value": to_bgr("#888888")},

            # Label A
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 535, 50, 25, 20], "as": "lbl_a"},
            {"action": "set", "target": "lbl_a.TextFrame.TextRange", "member": "Text", "value": "A"},
            {"action": "set", "target": "lbl_a.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "lbl_a.TextFrame.TextRange.Font", "member": "Size", "value": 13},

            # Heatmap Â (Right)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 645, 68, 95, 95], "as": "hm_ahat"},
            {"action": "call", "target": "hm_ahat.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "hm_ahat.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FFF2CC")},
            {"action": "set", "target": "hm_ahat.Line", "member": "Weight", "value": 1.0},
            {"action": "set", "target": "hm_ahat.Line", "member": "DashStyle", "value": 4},
            {"action": "set", "target": "hm_ahat.Line.ForeColor", "member": "RGB", "value": to_bgr("#888888")},

            # Label Â
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 680, 50, 25, 20], "as": "lbl_ahat"},
            {"action": "set", "target": "lbl_ahat.TextFrame.TextRange", "member": "Text", "value": "Â"},
            {"action": "set", "target": "lbl_ahat.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "lbl_ahat.TextFrame.TextRange.Font", "member": "Size", "value": 13},

            # Curved Loss Arrow & Label
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 595, 45, 50, 22], "as": "txt_loss"},
            {"action": "set", "target": "txt_loss.TextFrame.TextRange", "member": "Text", "value": "loss"},
            {"action": "set", "target": "txt_loss.TextFrame.TextRange.Font", "member": "Name", "value": "Segoe UI"},
            {"action": "set", "target": "txt_loss.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "txt_loss.TextFrame.TextRange.Font", "member": "Size", "value": 12},
            {"action": "set", "target": "txt_loss.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # Teal Box: Graph VAE
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 495, 185, 250, 30], "as": "b_gvae"},
            {"action": "call", "target": "b_gvae.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_gvae.Fill.ForeColor", "member": "RGB", "value": to_bgr("#A2D9CE")},
            {"action": "set", "target": "b_gvae.Line.ForeColor", "member": "RGB", "value": to_bgr("#1ABC9C")},
            {"action": "set", "target": "b_gvae.TextFrame.TextRange", "member": "Text", "value": "Graph VAE"},
            {"action": "set", "target": "b_gvae.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "b_gvae.TextFrame.TextRange.Font", "member": "Size", "value": 11},
            {"action": "set", "target": "b_gvae.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#0E6251")},

            # Labels S^P and S̃^P
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 535, 215, 30, 20], "as": "lbl_sp"},
            {"action": "set", "target": "lbl_sp.TextFrame.TextRange", "member": "Text", "value": "S ᴾ"},
            {"action": "set", "target": "lbl_sp.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 680, 215, 30, 20], "as": "lbl_stildep"},
            {"action": "set", "target": "lbl_stildep.TextFrame.TextRange", "member": "Text", "value": "S̃ ᴾ"},
            {"action": "set", "target": "lbl_stildep.TextFrame.TextRange.Font", "member": "Bold", "value": True},

            # Orange Box: Patch Embedding
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 495, 235, 250, 30], "as": "b_patch"},
            {"action": "call", "target": "b_patch.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_patch.Fill.ForeColor", "member": "RGB", "value": to_bgr("#F4B183")},
            {"action": "set", "target": "b_patch.Line.ForeColor", "member": "RGB", "value": to_bgr("#C55A11")},
            {"action": "set", "target": "b_patch.TextFrame.TextRange", "member": "Text", "value": "Patch Embedding"},
            {"action": "set", "target": "b_patch.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "b_patch.TextFrame.TextRange.Font", "member": "Size", "value": 11},
            {"action": "set", "target": "b_patch.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#FFFFFF")},

            # Dual input cards at bottom of (b)
            # Card 1: S (Left)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 495, 310, 115, 65], "as": "b_in_s"},
            {"action": "call", "target": "b_in_s.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_in_s.Fill.ForeColor", "member": "RGB", "value": to_bgr("#F8F9FA")},
            {"action": "set", "target": "b_in_s.Line.ForeColor", "member": "RGB", "value": to_bgr("#CCCCCC")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 499, 314, 52, 26], "as": "b_x1"},
            {"action": "call", "target": "b_x1.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_x1.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "b_x1.TextFrame.TextRange", "member": "Text", "value": "X exo"},
            {"action": "set", "target": "b_x1.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 553, 314, 52, 26], "as": "b_y1"},
            {"action": "call", "target": "b_y1.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_y1.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "b_y1.TextFrame.TextRange", "member": "Text", "value": "Y exo"},
            {"action": "set", "target": "b_y1.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 499, 344, 52, 26], "as": "b_x2"},
            {"action": "call", "target": "b_x2.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_x2.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "b_x2.TextFrame.TextRange", "member": "Text", "value": "X endo"},
            {"action": "set", "target": "b_x2.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 553, 344, 52, 26], "as": "b_y2"},
            {"action": "call", "target": "b_y2.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_y2.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "b_y2.TextFrame.TextRange", "member": "Text", "value": "Y endo"},
            {"action": "set", "target": "b_y2.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            # Card 2: S̃ (Right)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 630, 310, 115, 65], "as": "b_in_stilde"},
            {"action": "call", "target": "b_in_stilde.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_in_stilde.Fill.ForeColor", "member": "RGB", "value": to_bgr("#F8F9FA")},
            {"action": "set", "target": "b_in_stilde.Line.ForeColor", "member": "RGB", "value": to_bgr("#CCCCCC")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 634, 314, 52, 26], "as": "b_x3"},
            {"action": "call", "target": "b_x3.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_x3.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "b_x3.TextFrame.TextRange", "member": "Text", "value": "X exo"},
            {"action": "set", "target": "b_x3.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 688, 314, 52, 26], "as": "b_y3"},
            {"action": "call", "target": "b_y3.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_y3.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "b_y3.Line", "member": "DashStyle", "value": 4},
            {"action": "set", "target": "b_y3.TextFrame.TextRange", "member": "Text", "value": "Ỹ exo"},
            {"action": "set", "target": "b_y3.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 634, 344, 52, 26], "as": "b_x4"},
            {"action": "call", "target": "b_x4.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_x4.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "b_x4.TextFrame.TextRange", "member": "Text", "value": "X endo"},
            {"action": "set", "target": "b_x4.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 688, 344, 52, 26], "as": "b_y4"},
            {"action": "call", "target": "b_y4.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "b_y4.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "b_y4.Line", "member": "DashStyle", "value": 4},
            {"action": "set", "target": "b_y4.TextFrame.TextRange", "member": "Text", "value": "Ỹ endo"},
            {"action": "set", "target": "b_y4.TextFrame.TextRange.Font", "member": "Size", "value": 8},
        ]
        await run_batch(stage5_steps)
        await client.call_tool("office_save_document", {"document": doc_id})
        print("[+] Stage 5 completed & saved.")

        # -------------------------------------------------------------
        # STAGE 6: Submodule (c) Graph VAE Details (Far-Right)
        # -------------------------------------------------------------
        print("[*] Stage 6: Building (c) Graph VAE & Graph Learner...")
        stage6_steps = [
            # Stacked Adjacency Matrix at top of (c)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 805, 65, 95, 65], "as": "c_hm_stack"},
            {"action": "call", "target": "c_hm_stack.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "c_hm_stack.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FFF2CC")},
            {"action": "set", "target": "c_hm_stack.Line", "member": "DashStyle", "value": 4},
            {"action": "set", "target": "c_hm_stack.Line.ForeColor", "member": "RGB", "value": to_bgr("#D6B656")},

            # VAE box (pink)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 810, 145, 90, 24], "as": "c_vae"},
            {"action": "call", "target": "c_vae.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "c_vae.Fill.ForeColor", "member": "RGB", "value": to_bgr("#F8CBAD")},
            {"action": "set", "target": "c_vae.Line.ForeColor", "member": "RGB", "value": to_bgr("#C55A11")},
            {"action": "set", "target": "c_vae.TextFrame.TextRange", "member": "Text", "value": "VAE"},
            {"action": "set", "target": "c_vae.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "c_vae.TextFrame.TextRange.Font", "member": "Size", "value": 10},

            # Graph Learner container
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 795, 185, 120, 130], "as": "c_glearner"},
            {"action": "call", "target": "c_glearner.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "c_glearner.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FCE4D6")},
            {"action": "set", "target": "c_glearner.Line.ForeColor", "member": "RGB", "value": to_bgr("#F4B183")},

            # Inside heatmap of Graph Learner
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [1, 820, 205, 70, 70], "as": "c_inner_hm"},
            {"action": "call", "target": "c_inner_hm.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "c_inner_hm.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FFE599")},
            {"action": "set", "target": "c_inner_hm.Line.ForeColor", "member": "RGB", "value": to_bgr("#B45F06")},

            # Label Graph Learner
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 800, 290, 110, 20], "as": "txt_glearn"},
            {"action": "set", "target": "txt_glearn.TextFrame.TextRange", "member": "Text", "value": "Graph Learner"},
            {"action": "set", "target": "txt_glearn.TextFrame.TextRange.Font", "member": "Bold", "value": True},
            {"action": "set", "target": "txt_glearn.TextFrame.TextRange.Font", "member": "Size", "value": 9},
            {"action": "set", "target": "txt_glearn.TextFrame.TextRange.ParagraphFormat", "member": "Alignment", "value": 2},

            # Bottom patch embedding array
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 805, 345, 100, 35], "as": "c_patch_arr"},
            {"action": "call", "target": "c_patch_arr.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "c_patch_arr.Fill.ForeColor", "member": "RGB", "value": to_bgr("#F2F2F2")},
            {"action": "set", "target": "c_patch_arr.Line.ForeColor", "member": "RGB", "value": to_bgr("#CCCCCC")},
        ]
        await run_batch(stage6_steps)
        await client.call_tool("office_save_document", {"document": doc_id})
        print("[+] Stage 6 completed & saved.")

        # -------------------------------------------------------------
        # STAGE 7: Bottom Legend / Notation Panel
        # -------------------------------------------------------------
        print("[*] Stage 7: Adding Bottom Legend & Symbols...")
        stage7_steps = [
            # History / Future header
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 135, 432, 60, 20], "as": "leg_h"},
            {"action": "set", "target": "leg_h.TextFrame.TextRange", "member": "Text", "value": "history"},
            {"action": "set", "target": "leg_h.TextFrame.TextRange.Font", "member": "Size", "value": 9},

            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 195, 432, 60, 20], "as": "leg_f"},
            {"action": "set", "target": "leg_f.TextFrame.TextRange", "member": "Text", "value": "future"},
            {"action": "set", "target": "leg_f.TextFrame.TextRange.Font", "member": "Size", "value": 9},

            # Row 1: 外生变量 exo
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 40, 452, 90, 20], "as": "leg_lbl_exo"},
            {"action": "set", "target": "leg_lbl_exo.TextFrame.TextRange", "member": "Text", "value": "外生变量 exo"},
            {"action": "set", "target": "leg_lbl_exo.TextFrame.TextRange.Font", "member": "Name", "value": "Microsoft YaHei"},
            {"action": "set", "target": "leg_lbl_exo.TextFrame.TextRange.Font", "member": "Size", "value": 9},
            {"action": "set", "target": "leg_lbl_exo.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#C00000")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 135, 452, 50, 22], "as": "leg_xexo"},
            {"action": "call", "target": "leg_xexo.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "leg_xexo.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "leg_xexo.TextFrame.TextRange", "member": "Text", "value": "X exo"},
            {"action": "set", "target": "leg_xexo.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 195, 452, 50, 22], "as": "leg_yexo"},
            {"action": "call", "target": "leg_yexo.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "leg_yexo.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "leg_yexo.TextFrame.TextRange", "member": "Text", "value": "Y exo"},
            {"action": "set", "target": "leg_yexo.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            # Row 2: 内生变量 endo
            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 40, 482, 90, 20], "as": "leg_lbl_endo"},
            {"action": "set", "target": "leg_lbl_endo.TextFrame.TextRange", "member": "Text", "value": "内生变量 endo"},
            {"action": "set", "target": "leg_lbl_endo.TextFrame.TextRange.Font", "member": "Name", "value": "Microsoft YaHei"},
            {"action": "set", "target": "leg_lbl_endo.TextFrame.TextRange.Font", "member": "Size", "value": 9},
            {"action": "set", "target": "leg_lbl_endo.TextFrame.TextRange.Font.Color", "member": "RGB", "value": to_bgr("#C00000")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 135, 482, 50, 22], "as": "leg_xendo"},
            {"action": "call", "target": "leg_xendo.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "leg_xendo.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "leg_xendo.TextFrame.TextRange", "member": "Text", "value": "X endo"},
            {"action": "set", "target": "leg_xendo.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 195, 482, 50, 22], "as": "leg_yendo"},
            {"action": "call", "target": "leg_yendo.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "leg_yendo.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "leg_yendo.TextFrame.TextRange", "member": "Text", "value": "Y endo"},
            {"action": "set", "target": "leg_yendo.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            # Optional (dashed green)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 275, 462, 50, 26], "as": "leg_opt"},
            {"action": "call", "target": "leg_opt.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "leg_opt.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},
            {"action": "set", "target": "leg_opt.Line", "member": "DashStyle", "value": 4},
            {"action": "set", "target": "leg_opt.Line.ForeColor", "member": "RGB", "value": to_bgr("#70AD47")},

            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 330, 465, 65, 20], "as": "txt_opt"},
            {"action": "set", "target": "txt_opt.TextFrame.TextRange", "member": "Text", "value": "optional"},
            {"action": "set", "target": "txt_opt.TextFrame.TextRange.Font", "member": "Size", "value": 9},

            # Predict (dashed blue)
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 405, 462, 50, 26], "as": "leg_pred"},
            {"action": "call", "target": "leg_pred.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "leg_pred.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},
            {"action": "set", "target": "leg_pred.Line", "member": "DashStyle", "value": 4},
            {"action": "set", "target": "leg_pred.Line.ForeColor", "member": "RGB", "value": to_bgr("#4472C4")},

            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 460, 465, 60, 20], "as": "txt_pred"},
            {"action": "set", "target": "txt_pred.TextFrame.TextRange", "member": "Text", "value": "predict"},
            {"action": "set", "target": "txt_pred.TextFrame.TextRange.Font", "member": "Size", "value": 9},

            # Patch embedding chips
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 545, 465, 16, 18], "as": "chip1"},
            {"action": "call", "target": "chip1.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "chip1.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 565, 465, 16, 18], "as": "chip2"},
            {"action": "call", "target": "chip2.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "chip2.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D4EDDA")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 585, 465, 16, 18], "as": "chip3"},
            {"action": "call", "target": "chip3.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "chip3.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},

            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [5, 605, 465, 16, 18], "as": "chip4"},
            {"action": "call", "target": "chip4.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "chip4.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D9E1F2")},

            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 625, 465, 110, 20], "as": "txt_patch_lbl"},
            {"action": "set", "target": "txt_patch_lbl.TextFrame.TextRange", "member": "Text", "value": "patch embedding"},
            {"action": "set", "target": "txt_patch_lbl.TextFrame.TextRange.Font", "member": "Size", "value": 9},

            # Adjacency matrix mini heatmap
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [1, 745, 455, 32, 32], "as": "leg_hm"},
            {"action": "call", "target": "leg_hm.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "leg_hm.Fill.ForeColor", "member": "RGB", "value": to_bgr("#FFE599")},
            {"action": "set", "target": "leg_hm.Line.ForeColor", "member": "RGB", "value": to_bgr("#B45F06")},

            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 785, 455, 65, 35], "as": "txt_adj_lbl"},
            {"action": "set", "target": "txt_adj_lbl.TextFrame.TextRange", "member": "Text", "value": "adjacency\nmatrix"},
            {"action": "set", "target": "txt_adj_lbl.TextFrame.TextRange.Font", "member": "Size", "value": 8},

            # Weight color bar
            {"action": "call", "target": "slide.Shapes", "member": "AddShape", "args": [1, 855, 455, 12, 35], "as": "leg_bar"},
            {"action": "call", "target": "leg_bar.Fill", "member": "Solid", "args": []},
            {"action": "set", "target": "leg_bar.Fill.ForeColor", "member": "RGB", "value": to_bgr("#D6B656")},

            {"action": "call", "target": "slide.Shapes", "member": "AddTextbox", "args": [1, 870, 462, 50, 20], "as": "txt_wt_lbl"},
            {"action": "set", "target": "txt_wt_lbl.TextFrame.TextRange", "member": "Text", "value": "weight"},
            {"action": "set", "target": "txt_wt_lbl.TextFrame.TextRange.Font", "member": "Size", "value": 9},
        ]
        await run_batch(stage7_steps)
        await client.call_tool("office_save_document", {"document": doc_id})
        print("[+] Stage 7 completed & saved.")

        # -------------------------------------------------------------
        # STAGE 8: Create Independent Safe Copy
        # -------------------------------------------------------------
        print("[*] Stage 8: Calling office_create_copy to create verified final copy...")
        copy_res = unpack(await client.call_tool("office_create_copy", {
            "document": doc_id
        }))
        print("[+] Final Copy generated:", copy_res.get("copy_path"))
        
        # Keep open for the user to view in PowerPoint
        print(f"[+] All stages successfully constructed in Microsoft PowerPoint!")
        print(f"[+] Presentation is now live and visible on your desktop: {OUTPUT_PATH}")

if __name__ == "__main__":
    asyncio.run(main())
