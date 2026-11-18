# T3Lab FamiGen — Site Spec

Category guidance for **Site** (Revit template `Metric Site.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES
- Origin = centre of the footprint on finished ground (z = 0); anchoring below ground
  (footings, spikes) goes to negative z and should be fine-only.

# ANATOMY & TYPICAL SIZES
| Object | Size | Forms |
|---|---|---|
| Bollard | Ø150–Ø250 × 900–1000 | Revolution (chamfered top) |
| Bike rack hoop | Ø48 tube, 800 high, 600 wide, 600–800 pitch | Sweep of a Circle along an inverted U, `repeat` |
| Bench | 1600–2000 × 600–700 × 800, seat 450 | slats Extrusion + `repeat`, cast ends Extrusion in YZ |
| Litter bin | Ø400–Ø500 × 900–1000 | Revolution |
| Planter | 600–1200 square, 450–800 high | Extrusion with inner_loop |
| Fence panel | 2000–2500 × 1200–2000 | posts Extrusion + bars `repeat` |
# MATERIALS
- Galvanised steel #B7BCC0, powder-coat black #2B2B2E, concrete #A7A39C, timber #9C6B3E,
  corten #8B4A2B.
# DETAIL LEVELS
- coarse: envelope. fine: fixings, slats, chamfers.
