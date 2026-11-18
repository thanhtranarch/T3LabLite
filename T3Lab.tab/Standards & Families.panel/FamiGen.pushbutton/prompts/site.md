# T3Lab FamiGen — Site Spec

Category guidance for **Site** (Revit template `Metric Site.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES
- Origin = centre of the footprint on finished ground (z = 0); anchoring below ground
  (footings, spikes) goes to negative z and should be fine-only.

# ANATOMY & TYPICAL SIZES
| Object | Size | Forms |
|---|---|---|
| Bollard | Ø150–Ø250 × 900–1000 | Revolution (chamfered top) |
| Bike rack hoop | Ø48 tube, 800 high, 600 wide, 600–800 pitch | Sweep of a Circle along an inverted U, `repeat` |
| Bench | 1600–2000 × 600–700 × 800, seat 450 | slats Extrusion + `repeat`, cast ends Extrusion in YZ |
| Litter bin | Ø400–Ø500 × 900–1000 | Revolution |
| Planter | 600–1200 square, 450–800 high | Extrusion with inner_loop |
| Fence panel | 2000–2500 × 1200–2000 | posts Extrusion + bars `repeat` |
# MATERIALS
- Galvanised steel #B7BCC0, powder-coat black #2B2B2E, concrete #A7A39C, timber #9C6B3E,
  corten #8B4A2B.
# DETAIL LEVELS
- coarse: envelope. fine: fixings, slats, chamfers.


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
- Hoops, rails, frames and benches' armrests are Tubes; posts run from z = 0 (or the
  footing) into what they carry; slats rest on the frames.
- Play equipment: platforms rest on posts; railings stand on the platform edge and
  their top rail is one closed Tube or Revolution ring; slides and tube slides are
  Tubes from the platform edge down to the ground; stairs climb from the ground
  to the platform with each step touching the next.
