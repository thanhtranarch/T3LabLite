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
