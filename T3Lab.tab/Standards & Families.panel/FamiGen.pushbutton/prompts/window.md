# T3Lab FamiGen — Window Spec

Category guidance for **Window** (Revit template `Metric Window.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES (hosted in a wall — the template owns the opening)
- Origin = centre of the opening width at the level; the template places the sill by its
  `Sill Height`. X = width, Y = through the wall (exterior −Y), Z up.
- The template has `Width`, `Height` and planes Left, Right, Sill, Head: set `Width`/`Height`
  in `parameters[]`; do not create planes with those names.

# ANATOMY
| Part | Form | Notes |
|---|---|---|
| Outer frame (4 members) | Extrusion | 50–70 × 60–100 aluminium/uPVC, 50 × 100 timber |
| Sash per opening | Extrusion | 50–70 face width, set inside the frame |
| Mullion / transom | Extrusion | same depth as the frame |
| Glazing | Extrusion, transparency 60–80 | double glazing 24–28 thick, one solid |
| Sill (external) | Extrusion | 30–50 thick, projects 30–60 beyond the wall face |
| Handle | Cylinder | on the interior face |

# TYPICAL SIZES
| Window | Width × Height |
|---|---|
| Casement 1-pane | 600 × 1200 |
| Casement 2-pane | 1200 × 1500 |
| Fixed light | 900–2400 × 600–2100 |
| Sliding | 1500–2400 × 1200–1500 |
# DETAIL LEVELS
- coarse: frame + glass. medium/fine: sashes, mullions, sill, handles.
# PITFALLS
- Model the frame in the wall thickness band (Y between the wall faces); a frame centred on
  Y = 0 is usual. Glass inside sash/frame, never overlapping it.
