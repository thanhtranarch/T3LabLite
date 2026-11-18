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
