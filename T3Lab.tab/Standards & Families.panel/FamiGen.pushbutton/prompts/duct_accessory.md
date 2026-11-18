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
