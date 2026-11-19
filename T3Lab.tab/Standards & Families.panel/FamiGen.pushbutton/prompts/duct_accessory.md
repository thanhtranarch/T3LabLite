# T3Lab FamiGen — Duct Accessory Spec

Category guidance for **Duct Accessory** (Revit template `Metric Duct Accessory.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES (in-line duct items)
- Airflow along X through the origin: inlet at x = −L/2, outlet at x = +L/2. Rectangular
  duct sizes are W (along Y) × H (along Z), centred on the axis. No duct connectors are made.

# ANATOMY
| Part | Form | Notes |
|---|---|---|
| Casing / sleeve | Extrusion with an inner_loop (hollow) | galvanised #B7BCC0 |
| Flanges | Extrusion ring, 30–40 wide | |
| Blades (dampers) | Extrusion + `repeat` | pitch 100–150, 1–2 mm thick |
| Actuator / fusible link box | Extrusion on the side or top | |
| Round spigot | Cylinder along X | VAV inlet Ø100–Ø400 |
| Reheat coil (VAV) | Extrusion | 100–200 long, headers as Cylinders |
# TYPICAL SIZES
| Item | Size |
|---|---|
| Fire damper | duct size + 300–400 casing length |
| VAV box | 400–900 L, inlet Ø150–Ø300 |
| Attenuator | 900–1800 L |
# PARAMETERS
- `Duct Width`, `Duct Height`, `Length`; types per duct size.
# DETAIL LEVELS
- coarse: casing envelope. fine: blades, actuator, flanges.


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
- Casing, flanges and actuator box touch; blades span the full inside width and
  touch the casing sides.

## MEANING, SOLID / HOLLOW, JOINTS

- `concept.what`: an in-line duct part: damper, silencer, grille box, access door - say which one, its main parts, and its size class. `concept.use`: who or what uses it and how. `concept.placement`: what it stands on, hangs from or connects to.
- Roles and solid / hollow: The body is a conduit: HOLLOW sheet-metal shell (wall 1-2 mm: Extrusion with inner_loops for rectangular, Cylinder / Tube with "wall" for round), open at both duct ends. Blades, baffles and actuators are inside or on the shell (control / fixing). Flanges are fixing rings around each open end.
- Joints to declare: each flange to the body at its end; actuator to the body; blade shaft to both side walls. The air path between the two open ends stays clear except for blades / baffles that belong there (give them role control).
- FamiGen checks the model against this: parts that do not really touch ("Not connected"), joints that miss ("Joint"), containers and conduits left solid ("Hollow"), access parts without both landings ("Access") and blocked routes or bores ("Crossing") come back for correction.
