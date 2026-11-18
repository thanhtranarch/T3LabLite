# T3Lab FamiGen — Electrical Equipment Spec

Category guidance for **Electrical Equipment** (Revit template `Metric Electrical Equipment.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES
- Floor-standing: origin = centre of the footprint at floor level, front face toward −Y.
- Wall-mounted boards: back face on y = 0 (the wall), body toward −Y; give
  `Mounting Height` (instance) instead of modelling the height above floor.

# ANATOMY
| Part | Form | Notes |
|---|---|---|
| Enclosure body | Extrusion | sheet steel, RAL 7035 light grey (#C8C9C7) common |
| Door(s) | Extrusion 2–3 mm proud of the body | 3 mm gap between doors |
| Handle / lock | Cylinder / small Extrusion | 1100–1500 above floor |
| Plinth / base frame | Extrusion | 100 high, set back 20 |
| Louvres | Extrusion + `repeat` | 3–5 mm slots, 10–15 pitch |
| Gland plate / cable entry | Extrusion | top or bottom |
| Lifting eyes (transformer) | Revolution / Sweep | |

# TYPICAL SIZES
| Item | W × D × H |
|---|---|
| Distribution board (wall) | 400–800 × 150–250 × 500–1200 |
| Switchboard section | 600–1000 × 600–800 × 2000–2200 |
| Dry-type transformer | 1200–2000 × 800–1200 × 1500–2200 |
| Safety switch | 200–300 × 120–200 × 300–450 |
| UPS cabinet | 600 × 800 × 1900 |
# PARAMETERS / TYPES
- Width, Depth, Height (parametric: planes + locks on the body), optional `Rated Current`
  (number) and `Voltage` (text). Types per size.
# DETAIL LEVELS
- coarse: body only. medium: doors, plinth. fine: handles, louvres, glands, nameplate.


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
- Doors and plinths touch the enclosure; handles and glands overlap their faces;
  louvres are cut voids or thin parts touching the door.
