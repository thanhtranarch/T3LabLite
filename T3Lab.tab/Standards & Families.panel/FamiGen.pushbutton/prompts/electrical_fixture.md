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
