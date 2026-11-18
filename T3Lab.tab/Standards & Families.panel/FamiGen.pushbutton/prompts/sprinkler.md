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
