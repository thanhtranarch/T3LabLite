# T3Lab FamiGen — Door Spec

Category guidance for **Door** (Revit template `Metric Door.rft`). It adds design knowledge only:
the JSON contract in the system prompt (forms, fields, units) always wins. Millimetres,
radians, Z up. Sizes below are typical catalogue values for when neither the request nor a
reference gives one — list every such choice in `assumptions`.

# ORIGIN & AXES (hosted in a wall — the template owns the opening)
- Origin = centre of the opening at floor level. X = opening width (left −, right +),
  Y = through the wall (exterior −Y, interior +Y), Z up from the level.
- The template already has `Width` and `Height` parameters, the wall cut and planes named
  Left, Right, Exterior, Interior. SET `Width`/`Height` in `parameters[]` to the opening size;
  never create reference planes with those names. Name your own planes e.g. "Leaf Left".
- Frame sits inside the opening: frame outer edge = ±Width/2, head = Height.

# ANATOMY
| Part | Form | Notes |
|---|---|---|
| Frame jambs ×2 + head | Extrusion (profile in XZ via `sketch_plane_y`) or Extrusion in plan | Section typically 50–100 × 100–150 |
| Leaf (per leaf) | Extrusion | 40–54 thick (interior), 54–68 (external), 3 mm gap to frame |
| Glazing | Extrusion, `transparency` 60–80 | 6–12 thick, centred in the leaf/stile |
| Stiles / rails (glazed leaf) | Extrusion | stile 100–125, bottom rail 150–250 |
| Handle / pull | Cylinder + small Extrusion rosette | lever 900–1050 above floor, both faces |
| Threshold (external) | Extrusion | 15–20 high |

# TYPICAL SIZES
| Door | Width × Height | Leaf |
|---|---|---|
| Single interior | 800 / 900 / 1000 × 2100 | 40 |
| Single external | 900 / 1000 × 2100–2400 | 54 |
| Double | 1500–1800 × 2100 | 2 leaves |
| Aluminium storefront pair | 1800 × 2100–2400 | stiles 100 |

# TYPES & PARAMETRIC
- Types per standard width (e.g. 800, 900, 1000). With planes "Leaf Left"/"Leaf Right" at
  ±(Width/2 − frame) you may label a leaf-width dimension; keep the template's Width.
# DETAIL LEVELS
- coarse: frame + leaf only. medium/fine: handles, glazing beads, threshold, hinges.
# MATERIALS
- Frame, Leaf, Glass, Handle (e.g. "Door Leaf", "Door Frame", "Glass", "Brushed Steel").
# PITFALLS
- Parts outside ±Width/2 or above Height stick out of the wall cut.
- Glazing must be thinner than the leaf and fully inside it (no z-fighting faces).


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
- Frame jambs run from the floor to the head and overlap the head 5 mm; the leaf
  keeps a 3 mm gap to the frame but its hinge knuckles overlap both; handles touch
  the leaf face.
