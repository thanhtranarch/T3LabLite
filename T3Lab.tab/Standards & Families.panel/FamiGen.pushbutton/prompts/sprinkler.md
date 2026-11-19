# T3Lab FamiGen — Sprinkler Spec

Category guidance for **Sprinkler** (Revit template `Metric Sprinkler.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES
- Origin = thread connection point on the pipe; pendent heads hang toward −Z, upright
  heads point +Z, sidewall heads toward −Y. Overall heights are small (50–80).

# ANATOMY (pendent)
| Part | Form | Notes |
|---|---|---|
| Thread / body | Revolution about Z | Ø20–25, 15 mm (½") or 20 mm (¾") thread |
| Frame arms ×2 | Cylinder / Sweep | meet at the deflector boss |
| Glass bulb | Revolution, transparency 40, colour by temperature | Ø3–5 × 20 |
| Deflector | Extrusion disc Ø25–Ø55, 1–2 thick, slots as voids at fine | |
| Escutcheon (recessed) | Revolution | Ø60–Ø80 |
# BULB COLOURS (temperature rating)
- 57 °C orange, 68 °C red, 79 °C yellow, 93 °C green, 141 °C blue.
# PARAMETERS
- `K Factor` (number, e.g. 80), `Temperature Rating` (text), `Response` (text: Standard/Quick).
# DETAIL LEVELS
- coarse: body + deflector. fine: arms, bulb, deflector slots.


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
- Frame arms are Tubes from the body to the deflector boss; the bulb touches both
  the seat and the boss.

## MEANING, SOLID / HOLLOW, JOINTS

- `concept.what`: a sprinkler head: pendent, upright or sidewall - say which one, its main parts, and its size class. `concept.use`: who or what uses it and how. `concept.placement`: what it stands on, hangs from or connects to.
- Roles and solid / hollow: Body / frame = support on a threaded inlet. Inlet = conduit: a short HOLLOW Cylinder ("wall" 3) on the pipe axis. Deflector = thin surface (1-2 mm) at the free end; frame arms = two Tubes from the body to the deflector boss; bulb = decor between orifice and boss.
- Joints to declare: each frame arm to the body and to the deflector boss; deflector to the boss; escutcheon (fixing) to the body at the ceiling plane.
- FamiGen checks the model against this: parts that do not really touch ("Not connected"), joints that miss ("Joint"), containers and conduits left solid ("Hollow"), access parts without both landings ("Access") and blocked routes or bores ("Crossing") come back for correction.
