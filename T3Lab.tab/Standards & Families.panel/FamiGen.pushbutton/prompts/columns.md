# T3Lab FamiGen — Columns Spec

Category guidance for **Columns** (Revit template `Metric Column.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES
- Origin = column centre at the base level (Ref. Level). The template has `Width`/`Depth`
  for rectangular columns; for round columns create `Diameter`.
- Height in the project comes from base/top levels: model the shaft from z = 0 to a nominal
  height (3000) and, for flexing, lock its top to a plane "Top" labelled `Height`.

# ANATOMY
| Part | Form |
|---|---|
| Shaft (square) | Extrusion |
| Shaft (round / tapered / fluted) | Revolution (taper) or Extrusion of a Circle |
| Capital / base | Extrusion or Revolution, 1.3–1.6 × the shaft width |
| Plinth | Extrusion |

# TYPICAL SIZES
| Column | Size |
|---|---|
| Square concrete | 300 × 300 … 600 × 600 |
| Round concrete | Ø300 … Ø600 |
| Capital | height 200–400 |
# PARAMETRIC
- Planes "Left"/"Right" (x = ±W/2), "Front"/"Back" (y = ±D/2), "Top" (z = H); dimensions
  Width, Depth, Height; lock the shaft faces. Capital/base: lock to the same planes only if
  they share faces; otherwise leave them unlocked.
