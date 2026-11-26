---
name: ppt-architecture-diagram
description: >-
  Design, generate, and 1:1 reproduce academic architecture diagrams, neural
  network figures, research roadmaps, system pipelines, and flowcharts live in
  Microsoft PowerPoint or WPS Presentation using microsoft-wps-live MCP
  (`office_ppt_draw_scene`, `office_ppt_export_image`, `office_api_batch`).
  Use whenever the user provides a reference diagram to replicate 1:1 or asks
  to design any diagram/architecture figure from natural language.
---

# PPT Architecture & Scientific Diagram Skill

This skill governs two workflows over `microsoft-wps-live` MCP so that every
diagram renders smoothly on the user's visible desktop window, stays 100%
editable as native vector PowerPoint shapes, and achieves pixel-level style and
geometric fidelity.

## Mandatory Pre-Flight: Open & Raise Native App to Topmost View First

Whenever you are ready to **reproduce a reference diagram** OR **generate/edit a diagram or document from natural language** (across **PowerPoint**, **Word**, or **Excel**):
1. **Open / Create the native app FIRST** via `office_create_document` or `office_open_document` with `visible=true`.
2. **Raise it to the topmost foreground view BEFORE starting any drawing or staged edits** by immediately calling `office_show_document` (`focus: {"slide": 1}` for PPT, `{"start": 0}` for Word, or `{"sheet": "Sheet1", "address": "A1"}` for Excel).
   Never begin `office_ppt_draw_scene` or `office_live_steps` while the Office/WPS window is still closed, minimized, or buried behind the chat/IDE window — the user must see the live native application right on top from the very first step.

## Preferred Toolset

1. **`office_create_document`** / **`office_open_document`** (`visible=true`) → **`office_show_document`**
   immediately to open PowerPoint/WPS (or Word/Excel when also active) and bring
   the native window to the very top of the desktop view before any drawing begins.
2. **`office_ppt_draw_scene`** as the primary drawing engine for multi-element
   diagrams:
   - Maps reference-image pixel coordinates or logical coordinates automatically
     via `coord_space: {width, height, origin_x?, origin_y?, margin?, fit?}`.
   - Renders each completed primitive cleanly (`item_pause_ms`, default `20` ms)
     without per-property flicker, and saves every stage to disk.
   - Supports node `id` + `anchor` attachment (`start`/`end` as
     `{"node": "<id>", "anchor": "center|left|right|top|bottom|..."}`), exact
     line-intercept `junction`s (`on_line: {start, end, t}`), collinear
     multi-block chains (`aligned_chain` with automatic `atan2` slope and
     segmented arrows), `curve_arrow` with `arc_height`, and returns
     `layout_warnings` (overlapping textboxes/boxes/3D shapes or out-of-bounds).
3. **`office_ppt_export_image`** after drawing to export a PNG of the live
   slide, visually inspect it (plus zoom-in inspection of dense junctions) using
   your image/vision inspection tool (e.g., `read_file` / `Read` / `view_image`),
   fix any `layout_warnings` or misalignment via `office_api_batch`,
   and finish by calling **`office_create_copy`**.

---

## Mode A: Reference-to-PPT (1:1 Reproduction)

Follow this strict 5-phase protocol whenever the user supplies a reference image:

### Phase 1: Visual & Geometric Audit (Before Writing Any Coordinates)
Inspect the reference image carefully and write down:
1. **Canvas & Sub-regions**: Total pixel dimensions `(W, H)` (or crop bounds),
   main panel vs. inset panels `(a)`, `(b)`, `(c)`, legend regions, and outer
   dashed/solid container boxes.
2. **Projection & Axes**:
   - Identify whether modules are **flat 2D**, **tilted/rotated**, **sheared
     parallelograms** (`shear_box` along a diagonal axis angle such as `-12°`),
     **3D elongated channel bars** (`prism3d` / `grid3d` with `depth: 1` and
     long `extrude`, e.g., `[-24, -19]`), or **multi-depth 3D tensor lattices**
     (`grid3d` with `depth > 1`).
   - Measure the exact 3D extrusion direction `(extrude_x, extrude_y)` and
     face skew (`skew_x`, `skew_y`) from the reference image! Never substitute
     generic flat rectangles or wrong-direction `AddShape(14)` cubes when the
     reference shows elongated diagonal bars/prisms. Note that `grid3d` handles
     Painter's occlusion order automatically for both positive and negative
     `extrude`.
3. **Typography & Mixed Runs**:
   - Font family (typically `Times New Roman` for papers, `Arial`/`Calibri` for
     engineering), exact case (`1x1Conv-BN` vs `1×1 Conv-BN`), bold vs italic
     math symbols ($W, H, C, E, K+1, N, N-1$), subscripts/superscripts (`runs`),
     and text rotation angles.
4. **Colors, Borders, Dashes, Shadows & Corner Radii**:
   - Sample exact fill hexes (`#RRGGBB`), border hexes, line weights, dash
     patterns (`"dash"`, `"dot"`, `"dash_dot"`), rounded-rect corner radii
     (`corner_radius`: small subtle vs pill-shaped), fill transparency, and
     subtle shadows if present.
5. **Arrow Topology, Junctions & Layering**:
   - Record every trunk's start/end and every 3-way split/merge `junction`
     (`on_line: {start, end, t}`) so residual curves, trunks, and coupling
     bridges meet at **one exact shared coordinate** (never floating or
     overshooting).
   - For tilted or horizontal multi-block branches/bridges (e.g., `1x1Conv ->
     Down -> LayerNorm`), use `aligned_chain` (`tilt_deg: "auto"`,
     `segment_arrows: true`) instead of manual box placement + single long line,
     so box slant matches `atan2(dy, dx)` cleanly and segmented arrows stay
     visible between blocks.
   - Always draw background zones, trunks, and skip `curve_arrow`s *before*
     foreground blocks so blocks cleanly mask underlying lines (never cut
     through text), and use `arc_height` when routing skip arcs cleanly past
     tall intermediate boxes.
   - Keep labels transparent-backed and offset beside dashed arrows rather than
     placing opaque white rectangles that clip lines.

### Phase 2: Build via `office_ppt_draw_scene` in Layered Stages
Always structure `stages` strictly from back to front (Z-order):
1. **Stage 1 — Backdrop & Containers**: Shaded zones (`polygon`), dashed inset
   borders (`box` with `transparent: true, dash: "dash"`), legends, and panel
   labels `(a)`, `(b)`, `(c)`.
2. **Stage 2 — Trunk & Backbone Geometry**: Define shared `junction`s first,
   draw continuous trunk arrows and `curve_arrow` skips, then place backbone
   `feature_stack`, `prism3d`, `grid3d`, and `shear_box` modules *on top*.
3. **Stage 3 — Multi-Block Coupling Bridges**: Use `aligned_chain` connecting
   exact `junction` nodes so every bridge module is strictly collinear and
   properly tilted with visible segmented arrowheads.
4. **Stage 4 — Insets, Cropped Thumbnails & Annotations**: Cropped input images
   (`image` with `crop_frac`), braces (`brace`), and foreground typography.

### Phase 3: Geometric & Visual Verification Loop (Mandatory for 1:1)
1. Inspect `layout_warnings` returned by `office_ppt_draw_scene` — resolve any
   unexpected `text_overlap`, `box_overlap`, `text_on_object`, or `out_of_bounds`
   before calling the diagram done.
2. Call `office_ppt_export_image` to save a fresh PNG preview and inspect it with
   your image/vision tool, zooming closely on tilted bridges, 3-way junctions, and 3D
   prisms/grids to verify against the reference image across 6 checks:
   - Overall proportions & whitespace margins
   - 3D prism direction, bar elongation, and 3-face shading contrast
   - Strictly parallel tilted bridges & exact 3-way junction convergence
   - Curve clearance around intermediate blocks and neighbor inset boxes
   - Text readability, rotation alignment, italic math symbols, transparent labels
   - Color accuracy (fills, borders, coupling accents)
3. If *any* discrepancy is visible, fix only the affected shapes (or rebuild
   cleanly into a fresh slide/scene) and re-export to verify before finishing.

---

## Mode B: Prompt-to-Diagram (Natural Language Free Design)

When the user describes a diagram in natural language (paper roadmap, neural
architecture, agent workflow, system block diagram, timeline, comparison grid):

1. **Clarify Structure Silently or Concisely**:
   - Determine the best layout archetype from [references/design-patterns.md](references/design-patterns.md):
     - **Dual-Branch / Coupled Architecture** (like Conformer / multimodal fusion)
     - **Encoder-Decoder / Stage Pipeline** (left-to-right or top-to-bottom blocks)
     - **Layered System Stack** (bottom infrastructure → middleware → agent/UI)
     - **Research Roadmap / Experiment Workflow** (phase swimlanes + feedback loops)
2. **Enforce Anti-Overlap & Visual Hierarchy Rules**:
   - Use `coord_space: {"width": 1600, "height": 900, "margin": 28}` (16:9 high-res
     logical grid) mapped onto a `960×540` widescreen slide.
   - Maintain consistent spacing (`>= 24` logical units between sibling boxes,
     `>= 48` units between swimlanes/tiers).
   - Pick **one cohesive academic palette** (max 2–3 primary hue families + neutral
     slate grays; light pastel fills `#DCEAF7` / `#E2EFDA` / `#FDECE4` / `#FFF2CC`
     with dark crisp `#1F2937` text and darker accent borders).
   - Give every module an `id` and connect arrows via `{"node": ..., "anchor": ...}`
     so layout stays clean and maintainable.
   - Group into 3–6 meaningful `stages` so the user watches the diagram grow
     section by section (`item_pause_ms: 25`, `pause_ms: 300`).
3. **Finish**: Export preview PNG with `office_ppt_export_image` to verify zero
   text overflow or overlapping arrows, then call `office_create_copy` and keep
   the native PowerPoint window open for the user's manual edits.
