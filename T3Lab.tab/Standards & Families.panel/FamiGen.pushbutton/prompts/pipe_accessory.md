# T3Lab FamiGen — Pipe Accessory Spec

Category guidance for **Pipe Accessory** (Revit template `Metric Pipe Accessory.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES (in-line items)
- Pipe axis along X through the origin: inlet face at x = −L/2, outlet at x = +L/2
  (L = face-to-face). Operator (stem, handwheel, lever, gearbox) points +Z.
- FamiGen creates no pipe connectors: geometry only. Say so in `assumptions` when the user
  expects connectors.

# ANATOMY
| Part | Form | Notes |
|---|---|---|
| Body | Revolution about X (profile in XZ via `sketch_plane_y: 0`) | globe/gate bodies bulge 1.3–1.6 × flange OD |
| Flanges ×2 | Cylinder along X | thickness 16–30; bolt holes as voids at fine |
| Bonnet / stem | Cylinder along Z | |
| Handwheel | Revolution (torus) or Sweep of a circle | spokes = Cylinders + `repeat` |
| Lever (ball valve) | Extrusion | 1.2–1.5 × L long, red/yellow |
| Strainer leg (Y-strainer) | Cylinder at 45° down | |

# FLANGE OUTSIDE DIAMETERS (EN 1092-1 PN16, mm)
| DN | 25 | 50 | 80 | 100 | 150 | 200 |
|---|---|---|---|---|---|---|
| Flange OD | 115 | 165 | 200 | 220 | 285 | 340 |
| Bolts | 4 | 4 | 8 | 8 | 8 | 12 |
# TYPICAL FACE-TO-FACE (EN 558 series 14/15 style, approximate)
- Gate valve DN100 ≈ 190–230; butterfly (wafer) DN150 ≈ 56; ball valve DN25 threaded ≈ 85–95;
  Y-strainer DN80 ≈ 310.
# PARAMETERS / TYPES
- `Nominal Diameter` (length), `Face to Face` (length), `Manufacturer` (text). Types per DN.
# DETAIL LEVELS
- coarse: body + flanges as one simple envelope. fine: bolts, handwheel spokes, nameplate.


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
- Body, flanges and bonnet overlap along the pipe axis; handwheel rim is a closed
  Tube (points around the circle, closed: true) with spoke Cylinders into the hub.

## MEANING, SOLID / HOLLOW, JOINTS

- `concept.what`: an in-line pipe part: valve, strainer, flow meter, flexible joint - say which one, its main parts, and its size class. `concept.use`: who or what uses it and how. `concept.placement`: what it stands on, hangs from or connects to.
- Roles and solid / hollow: The body is a conduit: HOLLOW with the pipe's bore open end to end (Cylinder / Tube with "wall", or a Revolution of the wall section about the pipe axis). Flanges = fixing rings at each end (Extrusion / Revolution with the bore as an inner loop). Bonnet, stem, handwheel or lever = control, on top of the body, never inside the bore.
- Joints to declare: each flange to the body at its end; bonnet to the body; stem to bonnet and handwheel. The bore between the two ends must stay clear.
- FamiGen checks the model against this: parts that do not really touch ("Not connected"), joints that miss ("Joint"), containers and conduits left solid ("Hollow"), access parts without both landings ("Access") and blocked routes or bores ("Crossing") come back for correction.
