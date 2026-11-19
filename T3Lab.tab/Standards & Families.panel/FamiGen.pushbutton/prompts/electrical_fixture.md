# T3Lab FamiGen — Electrical Fixture Spec

Category guidance for **Electrical Fixture** (Revit template `Metric Electrical Fixture.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES (face-mounted wiring devices)
- Back of the device on the mounting face: back plane y = 0, device grows toward −Y
  (into the room); centre of the plate at x = 0, z = 0. Mounting height is an instance
  parameter in the project (`Mounting Height`, e.g. 300 sockets, 1100 switches).

# ANATOMY
| Part | Form | Notes |
|---|---|---|
| Face plate | Extrusion with 1–2 mm chamfer look via a Blend | 86 × 86 (single), 146 × 86 (double) |
| Socket face / rocker | Extrusion 2–4 proud | |
| Pin holes | small voids | fine only |
| Back box (flush) | Extrusion behind the plate | 35–47 deep |
| Junction box | Extrusion + knockout cylinders | 100 × 100 × 50 … 150 × 150 × 100 |
# TYPICAL SIZES
- BS 1363 socket plate 86 × 86 (single) / 146 × 86 (double), 8–10 thick plate.
- Schuko plate 80 × 80. US duplex 70 × 115.
# DETAIL LEVELS
- coarse: plate only. fine: sockets, switches, screws.


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
- Face plate on the back plane (y = 0); sockets and rockers overlap the plate front.

## MEANING, SOLID / HOLLOW, JOINTS

- `concept.what`: a socket, switch, junction box or data outlet on a wall or ceiling - say which one, its main parts, and its size class. `concept.use`: who or what uses it and how. `concept.placement`: what it stands on, hangs from or connects to.
- Roles and solid / hollow: Back box = container (HOLLOW, open to the front, sits in the wall); faceplate = surface covering it and overlapping the wall face by 2-5; rocker / socket inserts = control. Socket holes are voids in the faceplate.
- Joints to declare: faceplate to the back box; each rocker or insert to the faceplate; screws (fixing) to the faceplate.
- FamiGen checks the model against this: parts that do not really touch ("Not connected"), joints that miss ("Joint"), containers and conduits left solid ("Hollow"), access parts without both landings ("Access") and blocked routes or bores ("Crossing") come back for correction.
