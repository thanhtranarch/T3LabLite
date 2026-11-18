# T3Lab FamiGen — Entourage Spec

Category guidance for **Entourage** (Revit template `Metric Entourage.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES
- Origin = centre of the footprint on the ground plane.

# GUIDANCE
- Entourage is context, not product: keep it light (fewer than 20 parts) and recognisable
  in silhouette from plan and elevation.
- Shrubs and trees: crown as a Revolution (squashed sphere profile, Arc3P/Spline) or a
  stack of Blends; trunk as a Cylinder. Typical: shrub Ø800–Ø1500 × 600–1200; tree crown
  Ø3000–Ø6000, height 5000–10000, trunk Ø200–Ø400.
- People (stylised): 1700–1800 tall, shoulder width 450, Revolution torso + Cylinder limbs.
- Cars (stylised): 4500 × 1800 × 1500; body Extrusion with a profile drawn in XZ
  (`sketch_plane_y`), wheels Cylinders Ø650 along Y.
- Colours: foliage #6E8B3D / #4F6B2E, bark #6B4F3A, glazing transparency 60.
- Use `transparency` 30–50 on large crowns so the model behind stays readable.
