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
- Trunk into crown; people's limbs into the torso; car wheels touch the ground and
  the body.
