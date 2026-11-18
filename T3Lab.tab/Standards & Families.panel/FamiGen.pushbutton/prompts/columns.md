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


# JOINTS & SMOOTH FORMS (FamiGen checks these: loose parts are sent back to you)
- Build from the base up and give every part a parent it overlaps by 1-5 mm on all
  three axes; the chain must reach the lowest part. Floating parts are reported as
  "Not connected" and returned for one more round.
- Take a child's start coordinate from its parent's face (top z, side x/y) - do not
  type an approximate number.
- Curved rods, rails, handles, pipes, cables and tube slides: one `Tube` through the
  bend points (it adds smooth spline segments with ball joints). Never several
  Cylinders meeting at an angle.
- Turned, domed or round-ended parts: `Revolution` with Arc3P / Spline profile edges.
- Visible top edges of thick slabs and worktops: chamfer or round them (Blend, or Arc
  corners in plan) when the reference shows it.
- Base, shaft and capital stack exactly (top z of one = bottom z of the next).
