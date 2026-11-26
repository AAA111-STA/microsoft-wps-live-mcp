# Diagram Design Patterns, Palettes & Primitive Recipes

## 1. Academic & Engineering Color Palettes

### Palette A: Paper Architecture (Blue + Sage Green + Warm Peach Accent)
- **CNN / Spatial Stream**: Fill `#DCEAF7`, Border `#4F6D8A`, Accent Arrow `#3F6AA8`
- **Transformer / Sequence Stream**: Fill `#E2EFDA` / `#9AC77E`, Border `#5B8C4F`, Accent Arrow `#4E8A42`
- **Token / Special Highlight**: Fill `#FCE3BE`, Border `#C99A5B`
- **Coupling / Attention Zone Backdrop**: Fill `#FDECE4`, Border `#E8B8A2` (or borderless)
- **Neutral Containers & Text**: Border `#7F7F7F`, Connector `#262626`, Text `#111111`

### Palette B: Modern Systems / AI Agent Stack (Slate + Indigo + Teal)
- **User / Interface Layer**: Fill `#EEF2FF`, Border `#4F46E5`, Accent `#4338CA`
- **Orchestration / Agent Core**: Fill `#ECFDF5`, Border `#059669`, Accent `#047857`
- **Tool / MCP / Execution Layer**: Fill `#FFF7ED`, Border `#EA580C`, Accent `#C2410C`
- **Storage / Infra**: Fill `#F1F5F9`, Border `#475569`, Accent `#334155`

### Palette C: Classic Clean Monograph (Warm Gray + Steel Blue + Gold)
- **Primary Box**: Fill `#F8FAFC`, Border `#334155`
- **Highlight Box**: Fill `#DBEAFE`, Border `#1D4ED8`
- **Key Metric / Gate**: Fill `#FEF3C7`, Border `#D97706`

---

## 2. Choosing the Right Primitive in `office_ppt_draw_scene`

| Visual Element | Recommended `type` | Key Parameters |
| :--- | :--- | :--- |
| Standard module / card / phase box | `box` | `rounded: true`, `corner_radius: 0.12`, `fill`, `border`, `shadow` |
| Outer dashed grouping boundary | `box` | `transparent: true`, `rounded: true`, `corner_radius: 0.08`, `dash: "dash"` |
| Diagonal 2.5D tilted block (`Conv-BN`, `MLP`) | `shear_box` | `tilt_deg: -12`, `skew_x`, `text_rotation: -12` |
| 3D tensor bar, token strip, or 3D block | `prism3d` | `extrude: [ex, ey]`, `skew_x`, `skew_y`, `fill`, `top_fill`, `side_fill` |
| 3D spatial tensor / channel bars | `grid3d` | `rows`, `cols`, `depth` (`1` for channel strips), `extrude`, `layer_dx` |
| Stacked feature maps | `feature_stack` | `count: 3`, `skew_y: -7`, `dx: 8`, `dy: 7` |
| Shaded coupling / attention region | `polygon` | `points: [[x,y], ...]`, `fill`, `border: null` |
| Shared 3-way split/merge anchor | `junction` | `id`, `on_line: {start, end, t}` or `point: [x, y]` |
| Collinear multi-block branch/bridge | `aligned_chain` | `start`, `end`, `nodes: [{id, t, w, h, text}]`, `tilt_deg: "auto"`, `segment_arrows: true` |
| Straight or orthogonal routed arrow | `arrow` | `start`, `end`, `via: [[x,y], ...]`, `dash: "dash"`, `arrow: "triangle"|"stealth"` |
| Residual skip / feedback arc | `curve_arrow` | `start`, `end`, `arc_height` (or explicit `c1`, `c2`) |
| Mixed bold/italic/subscript math label | `textbox` / `box` | `runs: [{"text": "x", "italic": true}, {"text": "i", "subscript": true}]` |
| Repetition span | `brace` | `cx`, `cy`, `w`, `h`, `rotation: 90` |
| Cropped sub-image from reference | `image` | `path`, `crop_frac: [left, top, right, bottom]` |

---

## 3. Seven Common Pitfalls in 1:1 Reproduction (And How to Avoid Them)

1. **Non-Parallel Tilted Branches & Bridges (`aligned_chain`)**: Never guess box
   `tilt_deg` independently of the trunk/bridge endpoints. Use `aligned_chain`
   with `tilt_deg: "auto"` so box slant equals `atan2(dy, dx)` and all nodes stay
   strictly collinear along the line.
2. **Long Continuous Arrow Hidden Behind Boxes**: Drawing one end-to-end arrow
   first and stacking opaque boxes on top hides intermediate arrowheads. Keep
   `segment_arrows: true` (default in `aligned_chain`) whenever arrowheads
   must remain visible between blocks.
3. **3-Way Junction Gaps or Overshoots (`junction`)**: Residual curves, trunks,
   and branch bridges often miss by a few pixels when typed separately. Define
   one `junction` (`on_line: {start, end, t}`) and reference `{"node": "<id>"}`
   across all converging shapes.
4. **3D Channel Strips vs. Voxel Arrays (`grid3d` / `prism3d`)**: Check whether a
   3D figure depicts elongated 1-deep channel strips (`depth: 1`, long
   `extrude: [-24, -19]`) or stacked `depth > 1` cubes, and always trace negative
   vs. positive extrusion direction.
5. **Wrong Z-Order & Skip-Curve Collisions (`arc_height`)**: Draw shaded zones,
   trunks, and `curve_arrow`s *before* foreground blocks so boxes cleanly mask
   crossing lines. Use `arc_height` on `curve_arrow` when routing around tall
   modules (`Multi-head Self-attention`) without clipping adjacent frames.
6. **White Labels Clipping Arrows**: Do not put opaque white-filled rectangles on
   top of dashed arrows for labels like `Class token` or `Pooling`; keep
   `textbox` transparent and offset cleanly beside or across lines.
7. **Unverified Dense Regions (`layout_warnings` + Zoom Audit)**: Check
   `layout_warnings` (`text_overlap`, `box_overlap`, `text_on_object`,
   `out_of_bounds`) after `office_ppt_draw_scene`, and visually inspect zoomed
   crops around tilted bridges and 3D blocks in `office_ppt_export_image` output.
