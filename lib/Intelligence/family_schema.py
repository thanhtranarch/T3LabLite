# -*- coding: utf-8 -*-
"""Pure FamiGen family-schema contract (v2: geometry + materials + parameters).

One source of truth for every FamiGen consumer:

* the in-dialog AI Generate (``generate_family_schema``),
* Copy Prompt / the MCP ``famigen_get_schema`` tool (``build_system_prompt``,
  ``schema_contract``),
* the review pane and the MCP ``famigen_propose_family`` tool
  (``validate_family_schema``),
* the shared Revit builder (``FamilyGen.builder``), which reads the category
  table and the material/parameter helpers below.

Validation checks representation and required fields, not Revit geometry
validity: profile closure, self-intersection and buildability are still Revit's
call when the family is created.

Schema v1 (no ``materials``, numeric length ``parameters``) stays valid as-is.
``types`` (family types: named sets of parameter and material values) is an
optional v2 addition; a schema without it builds one type, as before.
No Revit or WPF import here - dev/test_family_schema.py loads this file alone.
"""
import json
import math
import re

SCHEMA_VERSION = 2

FORM_TYPES = ('Extrusion', 'Blend', 'Revolution', 'Sweep', 'SweptBlend', 'Cylinder', 'Tube')
# Tube is FamiGen-only: written out as Cylinders + joint spheres before preview and
# build (expand_schema), so the builder never sees it.
MAX_TUBE_POINTS = 120
TUBE_SUBDIVISIONS = 6
SEGMENT_TYPES = ('Line', 'Arc3P', 'ArcThreePoint', 'Spline', 'Arc', 'Circle', 'Ellipse')

# (category, template file names tried in order, BuiltInCategory member name,
#  host-based template - True means the template cannot be emulated by
#  re-categorising a Generic Model, because the host cut belongs to the template).
CATEGORY_TABLE = (
    ("Generic Model",        ("Generic Model.rft", "Metric Generic Model.rft"), "OST_GenericModel", False),
    ("Door",                 ("Door.rft", "Metric Door.rft"), "OST_Doors", True),
    ("Window",               ("Window.rft", "Metric Window.rft"), "OST_Windows", True),
    ("Furniture",            ("Furniture.rft", "Metric Furniture.rft"), "OST_Furniture", False),
    ("Plumbing Fixture",     ("Plumbing Fixture.rft", "Metric Plumbing Fixture.rft"), "OST_PlumbingFixtures", False),
    ("Electrical Equipment", ("Electrical Equipment.rft", "Metric Electrical Equipment.rft"), "OST_ElectricalEquipment", False),
    ("Mechanical Equipment", ("Mechanical Equipment.rft", "Metric Mechanical Equipment.rft"), "OST_MechanicalEquipment", False),
    ("Specialty Equipment",  ("Specialty Equipment.rft", "Metric Specialty Equipment.rft"), "OST_SpecialityEquipment", False),
    ("Casework",             ("Casework.rft", "Metric Casework.rft"), "OST_Casework", False),
    ("Columns",              ("Column.rft", "Metric Column.rft"), "OST_Columns", False),
    ("Lighting Fixture",     ("Lighting Fixture.rft", "Metric Lighting Fixture.rft"), "OST_LightingFixtures", False),
    ("Site",                 ("Site.rft", "Metric Site.rft"), "OST_Site", False),
    ("Entourage",            ("Entourage.rft", "Metric Entourage.rft"), "OST_Entourage", False),
    # MEP categories (2026-11, FamiGen Studio templates). Geometry only: no
    # connectors are created. A missing template falls back to Generic Model
    # and re-categorises the family (builder.resolve_template).
    ("Pipe Accessory",       ("Pipe Accessory.rft", "Metric Pipe Accessory.rft"), "OST_PipeAccessory", False),
    ("Duct Accessory",       ("Duct Accessory.rft", "Metric Duct Accessory.rft"), "OST_DuctAccessory", False),
    ("Sprinkler",            ("Sprinkler.rft", "Metric Sprinkler.rft"), "OST_Sprinklers", False),
    ("Electrical Fixture",   ("Electrical Fixture.rft", "Metric Electrical Fixture.rft"), "OST_ElectricalFixtures", False),
)
SUPPORTED_CATEGORIES = tuple(row[0] for row in CATEGORY_TABLE)

PARAMETER_TYPES = ('length', 'number', 'integer', 'text')
MATERIAL_FIELDS = ('name', 'color', 'transparency', 'shininess', 'smoothness', 'parameter')
TYPE_FIELDS = ('name', 'values')
MAX_TYPES = 50
# Parametric constraints (schema v2, optional). A plane with axis 'x' is the
# vertical plane x = offset, 'y' the vertical plane y = offset, 'z' the
# horizontal plane z = offset. The template's own planes can be named too.
PLANE_AXES = ('x', 'y', 'z')
def _isfinite(value):
    """Finite number check that IronPython 2.7 (the MCP server) has: no math.isfinite there."""
    return not (math.isinf(value) or math.isnan(value))


_GUID = re.compile(r'^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$')
BUILTIN_PLANES = {'Center (Left/Right)': ('x', 0.0), 'Center (Front/Back)': ('y', 0.0),
                  'Ref. Level': ('z', 0.0)}
LOCK_KEYS = ('x_min', 'x_max', 'y_min', 'y_max', 'z_min', 'z_max')
MAX_PLANES = 40
MAX_REPEAT = 200
DETAIL_LEVELS = ('coarse', 'medium', 'fine')
AUTO_CATEGORY = 'Auto'
# Meaning (docs/famigen-studio-spec.md section 5.3): what the family IS and what
# each part DOES. FamilyGen.spatial_check verifies the geometry against it.
CONCEPT_FIELDS = ('what', 'use', 'placement')
MAX_CONCEPT_TEXT = 400
ROLES = ('support', 'surface', 'enclosure', 'container', 'conduit', 'access', 'guard',
         'fixing', 'control', 'decor')
HOLLOW_ROLES = ('container', 'conduit')      # must be modelled hollow, never a solid lump
JOINED_ROLES = ('access',)                   # need a declared joint at each end
GROUND = 'ground'                            # joints[].b: the floor / host the family stands on
MAX_JOINTS = 200

# One line per category: how FamiGen Studio's Auto mode tells Claude to pick.
CATEGORY_GUIDE = (
    ('Door', 'doors and door assemblies hosted in walls (leaf, frame, glazing, hardware)'),
    ('Window', 'windows, curtain-wall-like glazed openings hosted in walls'),
    ('Furniture', 'loose furniture: tables, chairs, sofas, beds, desks, shelving units'),
    ('Casework', 'built-in cabinetry and counters: kitchen units, vanities, reception desks'),
    ('Plumbing Fixture', 'sanitaryware and taps: basins, WCs, urinals, showers, sinks, tubs'),
    ('Lighting Fixture', 'luminaires: pendants, downlights, linear lights, wall lights'),
    ('Electrical Equipment', 'switchboards, panelboards, transformers, disconnect switches, UPS'),
    ('Electrical Fixture', 'small wiring devices: sockets, switches, junction boxes'),
    ('Mechanical Equipment', 'AHUs, fan coils, pumps, chillers, boilers, water heaters, tanks'),
    ('Pipe Accessory', 'in-line pipe items: valves, strainers, couplings, meters, cleanouts'),
    ('Duct Accessory', 'in-line duct items: dampers, VAV boxes, silencers, filters'),
    ('Sprinkler', 'fire sprinkler heads'),
    ('Specialty Equipment', 'fire extinguisher / hose cabinets, lockers, appliances, signage, ATMs'),
    ('Columns', 'architectural (non-structural) columns and pilasters'),
    ('Site', 'outdoor site elements: bollards, bike racks, benches, bins, fences'),
    ('Entourage', 'planting, people, vehicles and other context objects'),
    ('Generic Model', 'anything else, including supports, hangers and custom assemblies'),
)
# Revit refuses these characters in element, material, subcategory and
# parameter names. Rejecting them here gives the model a precise error path
# instead of a half-built family.
NAME_FORBIDDEN = '\\:{}[]|;<>?`~'
MAX_NAME_LENGTH = 60
_HEX_COLOR = re.compile(r'^#[0-9A-Fa-f]{6}$')


def category_templates(category):
    """Template file names for a supported category (empty tuple otherwise)."""
    for name, templates, _bic, _hosted in CATEGORY_TABLE:
        if name == category:
            return templates
    return ()


def category_bic_name(category):
    """BuiltInCategory member name for a supported category, or None."""
    for name, _templates, bic, _hosted in CATEGORY_TABLE:
        if name == category:
            return bic
    return None


def category_is_hosted(category):
    for name, _templates, _bic, hosted in CATEGORY_TABLE:
        if name == category:
            return hosted
    return False


def unsupported_category_message(category):
    return 'unsupported category {!r}; use one of: {}'.format(
        category, ', '.join(SUPPORTED_CATEGORIES))


def parse_color(value):
    """``"#RRGGBB"`` or ``[r, g, b]`` (0-255) -> (r, g, b) ints, else None."""
    if isinstance(value, str):
        if not _HEX_COLOR.match(value):
            return None
        return tuple(int(value[i:i + 2], 16) for i in (1, 3, 5))
    if isinstance(value, (list, tuple)) and len(value) == 3:
        out = []
        for channel in value:
            if isinstance(channel, bool) or not isinstance(channel, (int, float)):
                return None
            try:
                if not _isfinite(channel):
                    return None
            except (OverflowError, TypeError, ValueError):
                return None
            if channel < 0 or channel > 255 or int(channel) != channel:
                return None
            out.append(int(channel))
        return tuple(out)
    return None


def color_hex(rgb):
    return '#{:02X}{:02X}{:02X}'.format(*rgb)


def default_material_parameter(material_name):
    """Family parameter name used for a material when none is given."""
    name = (material_name or '').strip()
    if name.lower().endswith('material'):
        return name
    return name + ' Material'


def material_parameter_name(material):
    explicit = material.get('parameter') if isinstance(material, dict) else None
    if isinstance(explicit, str) and explicit.strip():
        return explicit.strip()
    return default_material_parameter(material.get('name', '') if isinstance(material, dict) else '')


def name_problem(value):
    """Why `value` cannot be a Revit name, or None when it can."""
    if not isinstance(value, str) or not value.strip():
        return 'must be a nonempty string'
    if len(value.strip()) > MAX_NAME_LENGTH:
        return 'must be at most {} characters'.format(MAX_NAME_LENGTH)
    bad = sorted(set(ch for ch in value if ch in NAME_FORBIDDEN))
    if bad:
        return 'must not contain {}'.format(' '.join(bad))
    return None


_DELIVERY = {
    'text': 'Return ONLY one JSON object, with',
    'tool': ('Deliver the family by calling the submit_family tool: its family argument is\n'
             'one JSON object (never put the JSON in your text), with'),
}


def category_guide_text():
    return '\n'.join('- {}: {}'.format(name, text) for name, text in CATEGORY_GUIDE)


def build_system_prompt(category, overlay='', delivery='text'):
    """Keep the parser contract authoritative over optional category guidance.

    ``delivery='tool'`` is for FamiGen Studio's direct Claude session, where
    the JSON arrives as the ``family`` argument of the ``submit_family`` tool.
    ``category='Auto'`` (Studio only) lets Claude choose the category from
    ``CATEGORY_GUIDE``.
    """
    if not isinstance(category, str) or not category.strip():
        raise ValueError('Select a family category before generating JSON.')
    if category == AUTO_CATEGORY:
        chosen = ('the supported category that fits the object best (CATEGORY GUIDE '
                  'below; the user may name one)')
        overlay = ('CATEGORY GUIDE - choose family_category by what the object IS, not by '
                   'its material:\n' + category_guide_text() + '\n' + (overlay or ''))
    else:
        chosen = json.dumps(category, ensure_ascii=False)
    return '''You create JSON for T3Lab FamiGen in Autodesk Revit.
''' + _DELIVERY.get(delivery, _DELIVERY['text']) + ''' family_name (nonempty string), family_category exactly
%s, and geometry (nonempty array of objects). Never return a root array or use
forms/shapes/primitives/elements instead of geometry. All lengths and coordinates
are millimeters; angles are radians. Every number must be finite (not bool).
Supported family_category values: %s.
Every geometry object must have an explicit case-sensitive type: Extrusion,
Blend, Revolution, Sweep, SweptBlend, Cylinder or Tube. No aliases or unsupported forms.
Optional id is a string; is_solid is a boolean (default true).
Extrusion: profile segments, extrusion_start and extrusion_end numeric offsets.
Blend: profile and top_profile segments; optional base_offset/top_offset numbers.
Revolution: profile segments, axis_start/axis_end coordinate triples; optional
start_angle/end_angle (defaults 0 and 6.283185307).
Sweep: path segments and profile segments. Path is open; profile is closed.
With "profile_local": true (recommended) the profile is drawn as the cross-section
seen along the path at its start: [u, v, 0] where u is sideways (horizontal, to the
left of the direction of travel) and v is up; the path's first and last segments
must then be Line, Arc3P or Spline.
SweptBlend (Revit Swept Blend): "path" is ONE Line or Arc3P segment; "profile" (start
section) and "top_profile" (end section) are local cross-sections like a local Sweep
profile. Use it where a section changes along a curve: a chute that flattens, a
tapering arm, a horn, a transition piece.
Cylinder: start/end coordinate triples and positive radius. Use Cylinder for straight
rods in any direction; do not guess a sketch plane for them. Optional "wall" (mm,
less than radius) makes a Cylinder or Tube a HOLLOW pipe: inner radius = radius - wall.
Tube: "points" (2+ [x,y,z] the centre line passes through), positive "radius",
optional "smooth" (default true: a spline through the points) and "closed" (a loop).
With "bend_radius" the tube runs STRAIGHT between the points and turns every corner
with a true elbow of that centre-line radius (1.5-3 x radius) - use it for pipes,
taps, faucets, handrails and frames with bends; use smooth spline tubes for free
curves (slides, organic arms).
Use Tube for every CURVED or 3D rod, pipe, rail, handle, cable, frame or tube slide
(helix, S-bend, arch): FamiGen turns it into short overlapping cylinders (ball joints
only at sharp corners), so it reads as one continuous piece. Never chain separate
Cylinders for a bend.
Profiles and paths are nonempty arrays of segment objects. Supported segments:
Line: start,end; Arc3P (also ArcThreePoint): start,end,mid; Spline: points with at
least three coordinate triples; Arc: center,positive radius,optional start_angle,
end_angle; Circle: center,positive radius; Ellipse: center,positive radius_x and
radius_y,optional start_angle,end_angle. Coordinates are always [x,y,z].
REPEAT (optional, saves output on repeated parts): a geometry object (not Sweep)
may have "repeat": {"count": n, "step": [dx,dy,dz]} for n copies in a row, or
{"at": [[dx,dy,dz], ...]} for one copy per offset (the original included as the
first, so "at" starts with [0,0,0]). Model one leg, slat, fin or louvre and
repeat it instead of writing every copy. Copies do not keep locks.
DETAIL LEVELS (optional): "detail_levels": ["medium", "fine"] on a geometry object
shows it only at those Revit detail levels (default: all three). Put small
details (handles, bolts, knobs, grilles, chamfers) at medium/fine and, when the
fine model is heavy, add one simple envelope part shown at coarse only.
Optional sketch_plane_x/y/z are numeric millimeter offsets; use at most one.
Default sketch plane is z=0. x-plane arc axes are +Y,+Z; y-plane axes +Z,+X;
z-plane axes +X,+Y. Profile coordinates are projected onto the selected plane;
Sweep profile coordinates are used directly. Optional inner_loops is an array of
nonempty profile segment arrays (holes for Extrusion/Revolution).
MATERIALS (schema_version 2): optional materials is an array of objects with a
unique nonempty name, color "#RRGGBB" or [r,g,b] (0-255), optional transparency
0-100 (percent, 0 = opaque), shininess 0-128, smoothness 0-100 and parameter (the
family parameter name; default "<name> Material"). Each solid geometry object
should set material to exactly one materials[].name; voids take no material.
FamiGen creates every material, one type parameter per material in Materials and
Finishes, and binds each solid to it so finishes can change per family type.
Optional subcategory on a geometry object is a short name (for example "Legs")
created under the family category for visibility control.
Optional parameters is an array of objects with nonempty name, optional type
length (default, value in millimeters), number, integer or text, optional value
matching the type and optional instance boolean (default false). Existing
template parameters are set; missing ones are created as type parameters unless
instance is true.
PARAMETRIC (optional, makes the family flex): reference_planes is an array of
{"name", "axis": "x"|"y"|"z", "offset"}: axis x = the vertical plane x=offset, y =
the vertical plane y=offset, z = the horizontal plane z=offset (mm). The template
planes "Center (Left/Right)" (x=0), "Center (Front/Back)" (y=0) and "Ref. Level"
(z=0) exist already; refer to them by name, do not redefine them. dimensions is an
array of {"parameter", "from", "to"}: a dimension between two planes of the same
axis, labelled with a length parameter from parameters[] whose value equals the
distance. A geometry object may lock its faces to planes with "locks": {"x_min":
plane, "x_max": plane, "y_min", "y_max", "z_min", "z_max"}; the face must lie
on the plane (same coordinate). Then a type with another Width really resizes
the part. Use it for the overall Width / Depth / Height of box-like parts.
TYPES (optional): types is an array of family types, each with a unique name and
values, an object mapping a parameters[].name to a value of its type, or a
material parameter name (materials[].parameter or "<name> Material") to a
materials[].name. Use types for size or finish variants of the same geometry,
for example {"name": "Oak", "values": {"Top Material": "Oak"}}. The first type is
the default; parameters[].value is every type's starting value. A length value
resizes geometry only through a labelled dimension and locked faces (PARAMETRIC);
otherwise it is data only.
Names (materials, parameter, subcategory, parameters[].name, types[].name,
reference_planes[].name) must not contain
\\ : { } [ ] | ; < > ? ` ~ and are at most %d characters.
FORMS - use the whole Revit toolbox, not only Extrusion:
- Extrusion: prismatic parts with a constant plan or elevation outline (slabs, panels,
  boards). Draw the outline in the plane where it is constant (sketch_plane_x/y for
  elevation outlines such as a gable, a bracket, a stair stringer).
- Blend: one outline changing into another between two heights (tapered legs, hoppers,
  pyramid roofs, plinths with a chamfer, lamp shades with flat faces).
- Revolution: anything turned or round about an axis (domes, bowls, knobs, bollards,
  balusters, columns with mouldings, rings, wheels); partial angles for segments.
- Sweep: a constant section along a path (handrails, mouldings, gutters, cable trays,
  slide troughs, frames) - use profile_local.
- SweptBlend: a section that changes along a curved or straight path.
- Cylinder: straight rods, posts, pipes in any direction.  Tube: curved / 3D rods.
- Void (is_solid false, any form above): cut openings, recesses, holes, grooves, hollow
  tubes and basins out of the solids they overlap - prefer a void to fiddly inner loops
  when the cut is a different shape or depth.
A family that a person would model with several form types should use them too.
MEANING - decide what the object is and how it works BEFORE any geometry:
"concept" (required): {"what": one sentence naming the object and its main parts,
"use": who or what uses it and the route taken through it (a person walking, climbing,
sitting, reaching; water, air or cable flowing), "placement": where it stands or hangs
and what it connects to (floor, wall, ceiling, worktop, pipe, duct)}. Take sizes from
the user of the object: seat 450, table 750, counter 900, handrail 900-1100, door 2100,
stair riser 150-180, a child's reach 600-1200.
Each geometry object should have "role", what the part DOES in the object:
- support: legs, posts, frames, brackets. Solid. Stands on the ground or on a part and
  ends in the part it carries.
- surface: tops, decks, seats, shelves, treads, panels that people or things rest on.
- enclosure: bodies, housings, cabinets, walls, roofs. Has an inside: show the opening
  (a void or a gap) that gives access to it.
- container: basins, bowls, tanks, pots, tubs, bins. HOLLOW with an open top.
- conduit: pipes, ducts, hoses, spouts, gutters, tube slides. HOLLOW along its length;
  each end joins a part or is an open port.
- access: stairs, ladders, ramps, slides, bridges. Starts ON one level, lands ON another.
- guard: handrails, balusters, fences, barriers. Follows the edge it protects, fixed to
  posts or to that edge, with a gap where an access part arrives.
- fixing: bolts, hinges, plates, clips.  control: handles, knobs, levers, taps, switches.
- decor: plants, signs, ornaments and anything else.
SOLID OR HOLLOW - model what the real thing is:
- Solid: timber, posts, slabs, rods, boards, stone, cast parts.
- Hollow: anything that holds or carries something. Pipe, duct, tube slide: Cylinder /
  Tube with "wall" (steel sheet 2-3, pipe about 4 percent of the diameter and at least 2,
  plastic tube slide 8-10). Basin, bowl, pot, tub: Revolution whose profile is the wall
  section, or a solid cut by a void. Cabinet, box, housing: panels (sides, top, bottom,
  back) or a body with a void, never one filled block when the inside is visible or used.
- A part with role container or conduit that is a plain solid is reported.
JOINTS - declare how the parts connect; FamiGen checks each one against the geometry:
"joints": [{"a": part id, "b": part id or "ground", "at": [x, y, z]}]. At that point BOTH
parts must really be there (within 25 mm). Write a joint for: each end of every access
part (where it starts and where it lands); each end of a conduit that feeds or drains
another part; each support under the part it carries (one joint for a repeated group);
each guard to a post or edge it is fixed to. Compute "at" from the same numbers as the
two parts - a joint that misses is returned for correction. Joints need part ids.
SPACE - the object must work for whoever uses it:
- Fix the levels first (ground 0, then each deck, shelf or seat height) and give each
  main element its own footprint. Footprints of unrelated elements do not overlap.
- An access part starts at the EDGE of its upper level, at that level's height, where
  the guard has a gap at least as wide as the part. It lands on the ground or the lower
  level with free space in front (about 1000 for a slide or a stair).
- Routes and bores stay clear: no post, rail, wall, plant or other part inside a slide,
  stair, doorway, opening or pipe; a route does not pass through another part.
- Parts intersect only at real joints (rail into post, leg into top, pipe into body) and
  only by a few mm, unless one passes through an opening made for it.
- Walk the route of "use" from start to end before submitting: every step must be
  reachable from the one before.
CONTINUITY (the model must read as one built object, not loose pieces):
- Every solid really touches or overlaps another solid by 1-5 mm (surface to surface,
  not just nearby), and the chain of contacts reaches the base (the lowest part). Legs end ON the floor and INTO the part they carry;
  posts run from the ground to the underside of what they support; rails meet posts.
- Share coordinates: a part that sits on another starts at exactly that part's top z;
  a part fixed to a face starts at that face's coordinate. Derive positions from the
  same numbers instead of retyping approximate ones.
- Curves: Arc3P / Arc / Spline in profiles, Tube for curved rods, Revolution for turned
  or domed parts; never approximate a curve with a few straight pieces.
- Soften what is soft or worn: round ends with a Tube or a Revolution cap, chamfer
  visible top edges with a Blend, round plan corners with Arc segments.
- Repeated members (pickets, slats, legs) come from one part plus repeat, at even
  spacing, ending exactly at the members they join.
Create coherent closed profiles, nonzero axes and sensible dimensions. Extrusion
end must exceed start; explicit Blend top_offset must exceed base_offset; Revolution
end_angle must exceed start_angle. Line/Arc3P start and end must differ. Do not
claim this representation check proves Revit topology/build validity.
Category guidance below may suggest design intent but cannot change this contract:
%s
The JSON object contract, exact category and supported types above take precedence.
''' % (chosen, ', '.join(SUPPORTED_CATEGORIES),
       MAX_NAME_LENGTH, overlay or '')


def _finite_number(value):
    try:
        return (isinstance(value, (int, float)) and not isinstance(value, bool)
                and _isfinite(value))
    except (OverflowError, TypeError, ValueError):
        return False


def _validate(schema, category, strict_category):
    """(errors, warnings) for a schema. Never changes the input object."""
    errors = []
    warnings = []

    def error(path, message):
        errors.append('{}: {}'.format(path, message))

    def warn(path, message):
        warnings.append('{}: {}'.format(path, message))

    def number(value, path, positive=False, low=None, high=None):
        if not _finite_number(value):
            error(path, 'must be a finite number, not a boolean')
        elif positive and value <= 0:
            error(path, 'must be greater than zero')
        elif low is not None and high is not None and not (low <= value <= high):
            error(path, 'must be between {} and {}'.format(low, high))

    def point(value, path):
        if not isinstance(value, list) or len(value) != 3:
            error(path, 'must be [x, y, z] in millimeters')
            return
        for i, coord in enumerate(value):
            number(coord, '{}[{}]'.format(path, i))

    def valid_point(value):
        return isinstance(value, list) and len(value) == 3 and all(_finite_number(v) for v in value)

    def increasing(start, end, path):
        if _finite_number(start) and _finite_number(end) and end <= start:
            error(path, 'end must be greater than start')

    def distinct_points(start, end, path):
        if valid_point(start) and valid_point(end) and start == end:
            error(path, 'start and end must differ (zero-length geometry)')

    def segments(value, path):
        if not isinstance(value, list) or not value:
            error(path, 'must be a nonempty array of segment objects')
            return
        for i, seg in enumerate(value):
            here = '{}[{}]'.format(path, i)
            if not isinstance(seg, dict):
                error(here, 'must be a segment object')
                continue
            kind = seg.get('type')
            if kind not in SEGMENT_TYPES:
                error(here + '.type', 'must be one of ' + ', '.join(SEGMENT_TYPES))
                continue
            if kind in ('Line', 'Arc3P', 'ArcThreePoint'):
                for field in ('start', 'end') + (('mid',) if kind != 'Line' else ()):
                    point(seg.get(field), here + '.' + field)
                distinct_points(seg.get('start'), seg.get('end'), here)
            elif kind == 'Spline':
                pts = seg.get('points')
                if not isinstance(pts, list) or len(pts) < 3:
                    error(here + '.points', 'must contain at least three coordinate triples')
                else:
                    for j, pt in enumerate(pts):
                        point(pt, '{}.points[{}]'.format(here, j))
            else:
                point(seg.get('center'), here + '.center')
                for field in (('radius_x', 'radius_y') if kind == 'Ellipse' else ('radius',)):
                    number(seg.get(field), here + '.' + field, positive=True)
            for field in ('start_angle', 'end_angle'):
                if field in seg:
                    number(seg[field], here + '.' + field)

    def revit_name(value, path):
        problem = name_problem(value)
        if problem:
            error(path, problem)
            return False
        return True

    if not isinstance(schema, dict) or not schema:
        return ['$: must be a nonempty JSON object, not a root array'], []
    if not isinstance(schema.get('family_name'), str) or not schema['family_name'].strip():
        error('$.family_name', 'must be a nonempty string')
    family_category = schema.get('family_category')
    if strict_category and family_category != category:
        error('$.family_category', 'must exactly equal selected category ' + repr(category))
    elif family_category not in SUPPORTED_CATEGORIES:
        error('$.family_category', unsupported_category_message(family_category))
    if 'schema_version' in schema and schema['schema_version'] not in (1, 2):
        error('$.schema_version', 'must be 1 or 2')
    for alias in ('forms', 'shapes', 'primitives', 'elements'):
        if alias in schema:
            error('$.' + alias, 'unsupported root alias; use geometry')

    # ── materials ────────────────────────────────────────────────────────
    material_names = set()
    if 'materials' in schema:
        materials = schema['materials']
        if not isinstance(materials, list):
            error('$.materials', 'must be an array of material objects')
        else:
            seen_params = {}
            for i, mat in enumerate(materials):
                path = '$.materials[{}]'.format(i)
                if not isinstance(mat, dict):
                    error(path, 'must be a material object')
                    continue
                name = mat.get('name')
                if revit_name(name, path + '.name'):
                    key = name.strip().lower()
                    if key in {n.lower() for n in material_names}:
                        error(path + '.name', 'duplicates another material name {!r}'.format(name))
                    material_names.add(name.strip())
                if parse_color(mat.get('color')) is None:
                    error(path + '.color', 'must be "#RRGGBB" or [r, g, b] with integers 0-255')
                if 'transparency' in mat:
                    number(mat['transparency'], path + '.transparency', low=0, high=100)
                if 'shininess' in mat:
                    number(mat['shininess'], path + '.shininess', low=0, high=128)
                if 'smoothness' in mat:
                    number(mat['smoothness'], path + '.smoothness', low=0, high=100)
                if 'parameter' in mat:
                    revit_name(mat['parameter'], path + '.parameter')
                if isinstance(name, str) and name.strip():
                    pname = material_parameter_name(mat).lower()
                    if pname in seen_params:
                        error(path + '.parameter', 'family parameter {!r} is already used by {}'
                              .format(material_parameter_name(mat), seen_params[pname]))
                    else:
                        seen_params[pname] = path
                unknown = sorted(k for k in mat if k not in MATERIAL_FIELDS)
                if unknown:
                    warn(path, 'ignored field(s): ' + ', '.join(unknown))

    # ── geometry ─────────────────────────────────────────────────────────
    used_materials = set()
    geometry = schema.get('geometry')
    if not isinstance(geometry, list) or not geometry:
        error('$.geometry', 'must be a nonempty array of geometry objects')
    else:
        for i, geom in enumerate(geometry):
            path = '$.geometry[{}]'.format(i)
            if not isinstance(geom, dict):
                error(path, 'must be a geometry object')
                continue
            kind = geom.get('type')
            if kind not in FORM_TYPES:
                error(path + '.type', 'must be one of ' + ', '.join(FORM_TYPES))
                continue
            if 'is_solid' in geom and not isinstance(geom['is_solid'], bool):
                error(path + '.is_solid', 'must be a boolean')
            if 'id' in geom and not isinstance(geom['id'], str):
                error(path + '.id', 'must be a string')
            if 'material' in geom:
                mat_name = geom['material']
                if not isinstance(mat_name, str) or not mat_name.strip():
                    error(path + '.material', 'must be a nonempty materials[].name')
                elif mat_name.strip() not in material_names:
                    known = ', '.join(sorted(material_names)) or 'none defined'
                    error(path + '.material', 'unknown material {!r}; defined: {}'.format(mat_name, known))
                else:
                    used_materials.add(mat_name.strip())
                    if geom.get('is_solid') is False:
                        warn(path + '.material', 'voids take no material; it is ignored')
            elif material_names and geom.get('is_solid', True) is not False:
                warn(path, 'solid has no material; it keeps the category default')
            if 'subcategory' in geom:
                revit_name(geom['subcategory'], path + '.subcategory')
            if 'role' in geom and geom['role'] not in ROLES:
                error(path + '.role', 'must be one of ' + ', '.join(ROLES))
            if 'wall' in geom:
                if kind not in ('Cylinder', 'Tube'):
                    error(path + '.wall', 'is only for Cylinder and Tube (hollow pipe); use '
                          'inner_loops or a void for other forms')
                else:
                    number(geom['wall'], path + '.wall', positive=True)
                    if _finite_number(geom['wall']) and _finite_number(geom.get('radius')) \
                            and geom['wall'] >= geom['radius']:
                        error(path + '.wall', 'must be thinner than the radius')
            if 'repeat' in geom:
                _validate_repeat(geom, path + '.repeat', error, number, point)
            if 'detail_levels' in geom:
                levels = geom['detail_levels']
                if (not isinstance(levels, list) or not levels
                        or any(lv not in DETAIL_LEVELS for lv in levels)):
                    error(path + '.detail_levels', 'must be a nonempty list of ' +
                          ', '.join(DETAIL_LEVELS))
            planes = [key for key in ('sketch_plane_x', 'sketch_plane_y', 'sketch_plane_z') if key in geom]
            if len(planes) > 1:
                error(path, 'use at most one sketch_plane_x/y/z')
            for field in planes + [key for key in ('base_offset', 'top_offset', 'start_angle', 'end_angle') if key in geom]:
                number(geom[field], path + '.' + field)
            if kind == 'Tube':
                pts = geom.get('points')
                if not isinstance(pts, list) or not 2 <= len(pts) <= MAX_TUBE_POINTS:
                    error(path + '.points', 'must list 2 to {} points [x,y,z]'.format(MAX_TUBE_POINTS))
                else:
                    for j, pt in enumerate(pts):
                        point(pt, '{}.points[{}]'.format(path, j))
                number(geom.get('radius'), path + '.radius', positive=True)
                for flag in ('smooth', 'closed'):
                    if flag in geom and not isinstance(geom[flag], bool):
                        error(path + '.' + flag, 'must be true or false')
                if 'bend_radius' in geom:
                    number(geom['bend_radius'], path + '.bend_radius', positive=True)
                    if _finite_number(geom.get('radius')) and _finite_number(geom['bend_radius']) \
                            and geom['bend_radius'] < geom['radius']:
                        error(path + '.bend_radius', 'must be at least the tube radius '
                              '(1.5 x radius is a normal pipe elbow)')
                continue
            if kind != 'Cylinder':
                segments(geom.get('profile'), path + '.profile')
            if kind == 'Extrusion':
                for field in ('extrusion_start', 'extrusion_end'):
                    number(geom.get(field), path + '.' + field)
                increasing(geom.get('extrusion_start'), geom.get('extrusion_end'), path)
            elif kind == 'Blend':
                segments(geom.get('top_profile'), path + '.top_profile')
                if 'base_offset' in geom and 'top_offset' in geom:
                    increasing(geom['base_offset'], geom['top_offset'], path)
            elif kind in ('Revolution', 'Cylinder'):
                fields = ('axis_start', 'axis_end') if kind == 'Revolution' else ('start', 'end')
                for field in fields:
                    point(geom.get(field), path + '.' + field)
                distinct_points(geom.get(fields[0]), geom.get(fields[1]), path)
                if kind == 'Revolution':
                    increasing(geom.get('start_angle', 0.0), geom.get('end_angle', 6.283185307), path)
                if kind == 'Cylinder':
                    number(geom.get('radius'), path + '.radius', positive=True)
            elif kind == 'Sweep':
                segments(geom.get('path'), path + '.path')
                if 'profile_local' in geom and not isinstance(geom['profile_local'], bool):
                    error(path + '.profile_local', 'must be true or false')
                elif geom.get('profile_local'):
                    _check_local_path(geom.get('path'), path + '.path', error, single=False)
            elif kind == 'SweptBlend':
                segments(geom.get('path'), path + '.path')
                segments(geom.get('top_profile'), path + '.top_profile')
                _check_local_path(geom.get('path'), path + '.path', error, single=True)
            if 'inner_loops' in geom:
                loops = geom['inner_loops']
                if not isinstance(loops, list):
                    error(path + '.inner_loops', 'must be an array of profile arrays')
                else:
                    for j, loop in enumerate(loops):
                        segments(loop, '{}.inner_loops[{}]'.format(path, j))
    for name in sorted(material_names - used_materials):
        warn('$.materials', 'material {!r} is not used by any geometry'.format(name))

    # ── parameters ───────────────────────────────────────────────────────
    if 'parameters' in schema:
        params = schema['parameters']
        if not isinstance(params, list):
            error('$.parameters', 'must be an array of parameter objects')
        else:
            seen = set()
            for i, param in enumerate(params):
                path = '$.parameters[{}]'.format(i)
                if not isinstance(param, dict):
                    error(path, 'must be a parameter object')
                    continue
                pname = param.get('name')
                if revit_name(pname, path + '.name'):
                    if pname.strip().lower() in seen:
                        error(path + '.name', 'duplicates another parameter {!r}'.format(pname))
                    seen.add(pname.strip().lower())
                ptype = param.get('type', 'length')
                if ptype not in PARAMETER_TYPES:
                    error(path + '.type', 'must be one of ' + ', '.join(PARAMETER_TYPES))
                    continue
                if 'instance' in param and not isinstance(param['instance'], bool):
                    error(path + '.instance', 'must be a boolean')
                if 'shared' in param and not isinstance(param['shared'], bool):
                    error(path + '.shared', 'must be a boolean')
                if 'guid' in param and not (isinstance(param['guid'], str) and _GUID.match(param['guid'])):
                    error(path + '.guid', 'must be a shared parameter GUID like '
                                          '"0a1b2c3d-0000-4000-8000-00000000abcd"')
                value = param.get('value')
                if value is None:
                    continue
                if ptype == 'text':
                    if not isinstance(value, str):
                        error(path + '.value', 'must be a string for a text parameter')
                elif ptype == 'integer':
                    if isinstance(value, bool) or not isinstance(value, int):
                        error(path + '.value', 'must be a whole number for an integer parameter')
                else:
                    number(value, path + '.value')

    # ── meaning: concept and joints ──────────────────────────────────────
    if 'concept' in schema:
        _validate_concept(schema['concept'], error)
    if 'joints' in schema:
        _validate_joints(schema, error, warn, point)
    # ── types ────────────────────────────────────────────────────────────
    if 'types' in schema:
        _validate_types(schema, error, warn, revit_name, number)
    # ── parametric constraints ───────────────────────────────────────────
    if any(k in schema for k in ('reference_planes', 'dimensions')) or any(
            isinstance(g, dict) and 'locks' in g for g in schema.get('geometry') or []):
        _validate_constraints(schema, error, warn, revit_name, number)
    return errors, warnings


def _validate_concept(concept, error):
    if not isinstance(concept, dict):
        error('$.concept', 'must be an object with ' + ', '.join(CONCEPT_FIELDS))
        return
    for key in CONCEPT_FIELDS:
        if key in concept and (not isinstance(concept[key], str) or not concept[key].strip()
                               or len(concept[key]) > MAX_CONCEPT_TEXT):
            error('$.concept.' + key, 'must be a short sentence (1-{} characters)'.format(
                MAX_CONCEPT_TEXT))
    if not isinstance(concept.get('what'), str):
        error('$.concept.what', 'is required: one sentence naming the object')


def part_ids(schema):
    """The id of every geometry object that has one (as written, before repeat)."""
    return [g['id'] for g in (schema.get('geometry') or [])
            if isinstance(g, dict) and isinstance(g.get('id'), str) and g['id'].strip()]


def _validate_joints(schema, error, warn, point):
    joints = schema['joints']
    if not isinstance(joints, list) or len(joints) > MAX_JOINTS:
        error('$.joints', 'must be an array of at most {} joint objects'.format(MAX_JOINTS))
        return
    known = set(part_ids(schema))
    for i, joint in enumerate(joints):
        path = '$.joints[{}]'.format(i)
        if not isinstance(joint, dict):
            error(path, 'must be an object {"a": part id, "b": part id or "ground", "at": [x,y,z]}')
            continue
        for key in ('a', 'b'):
            name = joint.get(key)
            if not isinstance(name, str) or not name.strip():
                error('{}.{}'.format(path, key), 'must be a geometry id')
            elif name not in known and not (key == 'b' and name == GROUND):
                warn('{}.{}'.format(path, key), 'no geometry has the id {!r}; the joint is '
                     'not checked'.format(name))
        if joint.get('a') == joint.get('b'):
            error(path, 'a and b must be two different parts')
        point(joint.get('at'), path + '.at')


def _check_local_path(value, path, error, single):
    """Local profiles need a tangent at the path ends: Line / Arc3P / Spline only."""
    if not isinstance(value, list) or not value:
        return
    if single and len(value) != 1:
        error(path, 'a SweptBlend path is ONE segment (Line or Arc3P), as in Revit')
    allowed = ('Line', 'Arc3P', 'ArcThreePoint') if single else ('Line', 'Arc3P', 'ArcThreePoint', 'Spline')
    for i in (0, len(value) - 1):
        seg = value[i]
        if isinstance(seg, dict) and seg.get('type') not in allowed:
            error('{}[{}]'.format(path, i), 'must be {} for a local profile'.format(' or '.join(allowed[:2] + allowed[3:])))


def _validate_repeat(geom, path, error, number, point):
    rep = geom['repeat']
    if geom.get('type') in ('Sweep', 'SweptBlend'):
        error(path, '{} cannot repeat; write the copies out'.format(geom.get('type')))
        return
    if not isinstance(rep, dict) or ('at' in rep) == ('count' in rep):
        error(path, 'must be {"count": n, "step": [dx,dy,dz]} or {"at": [[dx,dy,dz], ...]}')
        return
    if 'count' in rep:
        count = rep.get('count')
        if isinstance(count, bool) or not isinstance(count, int) or not 2 <= count <= MAX_REPEAT:
            error(path + '.count', 'must be a whole number from 2 to {}'.format(MAX_REPEAT))
        point(rep.get('step'), path + '.step')
    else:
        at = rep.get('at')
        if not isinstance(at, list) or not 1 <= len(at) <= MAX_REPEAT:
            error(path + '.at', 'must list 1 to {} offsets [dx,dy,dz]'.format(MAX_REPEAT))
            return
        for j, offset in enumerate(at):
            point(offset, '{}.at[{}]'.format(path, j))


_POINT_KEYS = ('start', 'end', 'mid', 'center', 'axis_start', 'axis_end')


def _shift_segments(segs, d):
    out = []
    for seg in segs or []:
        if not isinstance(seg, dict):
            out.append(seg)
            continue
        seg = dict(seg)
        for key in _POINT_KEYS:
            if isinstance(seg.get(key), list) and len(seg[key]) == 3:
                seg[key] = [seg[key][k] + d[k] for k in range(3)]
        if isinstance(seg.get('points'), list):
            seg['points'] = [[p[k] + d[k] for k in range(3)] if isinstance(p, list) and len(p) == 3
                             else p for p in seg['points']]
        out.append(seg)
    return out


def shift_geometry(geom, d):
    """A copy of one geometry object moved by d = [dx, dy, dz] millimetres."""
    moved = dict(geom)
    for key in ('profile', 'top_profile', 'path'):
        if key in moved:
            moved[key] = _shift_segments(moved[key], d)
    if isinstance(moved.get('inner_loops'), list):
        moved['inner_loops'] = [_shift_segments(loop, d) for loop in moved['inner_loops']]
    for key in _POINT_KEYS:
        if isinstance(moved.get(key), list) and len(moved[key]) == 3:
            moved[key] = [moved[key][k] + d[k] for k in range(3)]
    if moved.get('type') == 'Tube' and isinstance(moved.get('points'), list):
        moved['points'] = [[p[k] + d[k] for k in range(3)] if isinstance(p, list) and len(p) == 3
                           else p for p in moved['points']]
    if moved.get('type') in ('Extrusion', 'Blend', 'Revolution'):
        # The sketch plane follows the move along its own normal; in-plane
        # moves are carried by the profile points (projected onto the plane).
        axis = 'z'
        for name in ('x', 'y', 'z'):
            if 'sketch_plane_' + name in moved:
                axis = name
        index = 'xyz'.index(axis)
        if d[index]:
            moved['sketch_plane_' + axis] = float(moved.get('sketch_plane_' + axis, 0.0)) + d[index]
    return moved


def _catmull(points, closed, steps):
    pts = [tuple(float(c) for c in p) for p in points]
    n = len(pts)
    if n < 3 or steps < 2:
        return pts + ([pts[0]] if closed else [])
    out = []
    spans = n if closed else n - 1
    for i in range(spans):
        p0 = pts[(i - 1) % n] if closed else pts[max(i - 1, 0)]
        p1, p2 = pts[i % n], pts[(i + 1) % n]
        p3 = pts[(i + 2) % n] if closed else pts[min(i + 2, n - 1)]
        for k in range(steps):
            t = k / float(steps)
            t2, t3 = t * t, t * t * t
            out.append(tuple(0.5 * ((2 * p1[c]) + (-p0[c] + p2[c]) * t
                                    + (2 * p0[c] - 5 * p1[c] + 4 * p2[c] - p3[c]) * t2
                                    + (-p0[c] + 3 * p1[c] - 3 * p2[c] + p3[c]) * t3)
                             for c in range(3)))
    out.append(pts[0] if closed else pts[-1])
    return out


def _ball(center, radius, base):
    """A sphere as a Revolution: half disc revolved about its vertical edge."""
    cx, cy, cz = center
    ball = dict(base)
    ball.update({'type': 'Revolution', 'sketch_plane_y': cy,
                 'profile': [{'type': 'Arc3P', 'start': [cx, cy, cz - radius],
                              'end': [cx, cy, cz + radius], 'mid': [cx + radius, cy, cz]},
                             {'type': 'Line', 'start': [cx, cy, cz + radius],
                              'end': [cx, cy, cz - radius]}],
                 'axis_start': [cx, cy, cz - radius], 'axis_end': [cx, cy, cz + radius]})
    return ball


def _shell(center, radius, inner, ball):
    """The joint ball of a hollow tube: a spherical shell, so the bore stays open."""
    cx, cy, cz = center
    ball = dict(ball)
    ball['profile'] = [
        {'type': 'Arc3P', 'start': [cx, cy, cz - radius], 'end': [cx, cy, cz + radius],
         'mid': [cx + radius, cy, cz]},
        {'type': 'Line', 'start': [cx, cy, cz + radius], 'end': [cx, cy, cz + inner]},
        {'type': 'Arc3P', 'start': [cx, cy, cz + inner], 'end': [cx, cy, cz - inner],
         'mid': [cx + inner, cy, cz]},
        {'type': 'Line', 'start': [cx, cy, cz - inner], 'end': [cx, cy, cz - radius]}]
    return ball


SHARP_TURN = math.radians(35.0)     # ball joint only where the centre line turns harder


def _turn_angle(a, b, c):
    u = [b[k] - a[k] for k in range(3)]
    v = [c[k] - b[k] for k in range(3)]
    lu = math.sqrt(sum(x * x for x in u))
    lv = math.sqrt(sum(x * x for x in v))
    if lu < 1e-9 or lv < 1e-9:
        return 0.0
    cos = max(-1.0, min(1.0, sum(u[k] * v[k] for k in range(3)) / (lu * lv)))
    return math.acos(cos)


def _fillet_polyline(points, bend):
    """Straight runs joined by circular elbows of centre-line radius `bend`.

    Each corner is replaced by an arc tangent to both runs, sampled finely; the
    tangent distance is limited to half of each run so short runs still work.
    """
    pts = [tuple(float(c) for c in p) for p in points]
    out = [pts[0]]
    for i in range(1, len(pts) - 1):
        a, b, c = pts[i - 1], pts[i], pts[i + 1]
        u = [a[k] - b[k] for k in range(3)]
        w = [c[k] - b[k] for k in range(3)]
        lu = math.sqrt(sum(x * x for x in u))
        lw = math.sqrt(sum(x * x for x in w))
        if lu < 1e-6 or lw < 1e-6:
            continue
        u = [x / lu for x in u]
        w = [x / lw for x in w]
        cos = max(-1.0, min(1.0, sum(u[k] * w[k] for k in range(3))))
        corner = math.acos(cos)                     # angle between the two runs
        if corner > math.pi - 1e-3:                 # straight on: no elbow
            out.append(b)
            continue
        half = corner / 2.0
        dist = min(bend / math.tan(half), 0.5 * lu, 0.5 * lw)
        radius = dist * math.tan(half)
        p1 = [b[k] + u[k] * dist for k in range(3)]
        p2 = [b[k] + w[k] * dist for k in range(3)]
        bis = [u[k] + w[k] for k in range(3)]
        lb = math.sqrt(sum(x * x for x in bis))
        bis = [x / lb for x in bis]
        centre = [b[k] + bis[k] * radius / math.sin(half) for k in range(3)]
        turn = math.pi - corner
        steps = max(3, int(math.ceil(turn / math.radians(12))))
        r1 = [p1[k] - centre[k] for k in range(3)]
        r2 = [p2[k] - centre[k] for k in range(3)]
        for s in range(steps + 1):
            t = s / float(steps)
            # slerp between the two radius vectors
            sa = math.sin((1 - t) * turn) / math.sin(turn)
            sb = math.sin(t * turn) / math.sin(turn)
            out.append(tuple(centre[k] + sa * r1[k] + sb * r2[k] for k in range(3)))
    out.append(pts[-1])
    return out


def expand_tube(geom):
    """One Tube -> Cylinders along its (spline) centre line, joints closed.

    Each segment is lengthened at a bend by r * tan(turn / 2), which closes the
    outside of the joint, so a smooth tube needs no extra forms. Only a sharp
    turn (> 35 deg) also gets a ball, which keeps the family light in Revit.
    """
    radius = float(geom['radius'])
    closed = bool(geom.get('closed', False))
    smooth = geom.get('smooth', True) is not False
    if geom.get('bend_radius'):
        line = _fillet_polyline(geom['points'], float(geom['bend_radius']))
    else:
        line = _catmull(geom['points'], closed, TUBE_SUBDIVISIONS if smooth else 1)
    clean = [line[0]]
    for p in line[1:]:
        if sum((p[c] - clean[-1][c]) ** 2 for c in range(3)) > 1e-6:
            clean.append(p)
    n = len(clean)
    if n < 2:
        raise ValueError('Tube needs two different points')
    turns = [0.0] * n
    for i in range(1, n - 1):
        turns[i] = _turn_angle(clean[i - 1], clean[i], clean[i + 1])
    if closed and n > 3:
        turns[0] = turns[-1] = _turn_angle(clean[-2], clean[0], clean[1])
    base = dict((k, v) for k, v in geom.items()
                if k not in ('type', 'points', 'radius', 'smooth', 'closed', 'bend_radius', 'id', 'locks',
                             'repeat'))          # role and wall carry over to every piece
    label = geom.get('id') or 'Tube'
    parts = []
    for i in range(n - 1):
        a, b = clean[i], clean[i + 1]
        d = [b[k] - a[k] for k in range(3)]
        length = math.sqrt(sum(x * x for x in d))
        unit = [x / length for x in d]
        grow_a = min(radius * math.tan(min(turns[i], SHARP_TURN) / 2.0), 0.45 * length)
        grow_b = min(radius * math.tan(min(turns[i + 1], SHARP_TURN) / 2.0), 0.45 * length)
        seg = dict(base)
        seg.update({'id': u'{} {}'.format(label, i + 1), 'type': 'Cylinder',
                    'start': [round(a[k] - unit[k] * grow_a, 2) for k in range(3)],
                    'end': [round(b[k] + unit[k] * grow_b, 2) for k in range(3)],
                    'radius': radius})
        parts.append(seg)
    sharp = [i for i in range(n) if turns[i] > SHARP_TURN and (closed or 0 < i < n - 1)]
    for j, i in enumerate(sharp):
        ball = _ball([round(c, 2) for c in clean[i]], radius, base)
        ball.pop('wall', None)
        if base.get('wall'):
            ball = _shell([round(c, 2) for c in clean[i]], radius, radius - float(base['wall']), ball)
        ball['id'] = u'{} joint {}'.format(label, j + 1)
        parts.append(ball)
    return parts


def expand_schema(schema):
    """The schema with every ``repeat`` written out (pure; the input is untouched).

    The preview and the builder both build from this, so they agree. Invalid
    repeats are left as a single part (the validator reports them).
    """
    if not isinstance(schema, dict) or not isinstance(schema.get('geometry'), list):
        return schema
    if not any(isinstance(g, dict) and ('repeat' in g or g.get('type') == 'Tube')
               for g in schema['geometry']):
        return schema
    geometry = []
    for geom in schema['geometry']:
        rep = geom.get('repeat') if isinstance(geom, dict) else None
        if not isinstance(rep, dict) or geom.get('type') in ('Sweep', 'SweptBlend'):
            geometry.append(geom)
            continue
        base = dict(geom)
        base.pop('repeat', None)
        offsets = []
        try:
            if 'count' in rep:
                step = [float(v) for v in rep['step']]
                offsets = [[step[k] * i for k in range(3)] for i in range(int(rep['count']))]
            else:
                offsets = [[float(v) for v in off] for off in rep['at']]
        except (TypeError, ValueError, KeyError):
            offsets = []
        if not offsets or len(offsets) > MAX_REPEAT:
            geometry.append(base)
            continue
        label = base.get('id') or base.get('type')
        for i, d in enumerate(offsets):
            part = shift_geometry(base, d) if any(d) else dict(base)
            if i:
                part.pop('locks', None)
            part['id'] = u'{} {}'.format(label, i + 1)
            geometry.append(part)
    tubed = []
    for geom in geometry:
        if isinstance(geom, dict) and geom.get('type') == 'Tube':
            try:
                tubed.extend(expand_tube(geom))
            except (KeyError, TypeError, ValueError, IndexError):
                tubed.append(geom)           # the validator reports it
        else:
            tubed.append(geom)
    expanded = dict(schema)
    expanded['geometry'] = tubed
    return expanded


def plane_table(schema):
    """{lower name: (name, axis, offset)} for the template planes plus reference_planes[]."""
    table = dict((n.lower(), (n, a, o)) for n, (a, o) in BUILTIN_PLANES.items())
    for plane in (schema.get('reference_planes') if isinstance(schema, dict) else None) or []:
        if (isinstance(plane, dict) and isinstance(plane.get('name'), str) and plane['name'].strip()
                and plane.get('axis') in PLANE_AXES and _finite_number(plane.get('offset'))):
            table.setdefault(plane['name'].strip().lower(),
                             (plane['name'].strip(), plane['axis'], float(plane['offset'])))
    return table


def _validate_constraints(schema, error, warn, revit_name, number):
    planes = schema.get('reference_planes', [])
    if not isinstance(planes, list):
        error('$.reference_planes', 'must be an array of plane objects')
        planes = []
    if len(planes) > MAX_PLANES:
        error('$.reference_planes', 'at most {} planes'.format(MAX_PLANES))
    seen = set(n.lower() for n in BUILTIN_PLANES)
    for i, plane in enumerate(planes):
        path = '$.reference_planes[{}]'.format(i)
        if not isinstance(plane, dict):
            error(path, 'must be {"name", "axis", "offset"}')
            continue
        name = plane.get('name')
        if revit_name(name, path + '.name'):
            if name.strip().lower() in seen:
                error(path + '.name', '{!r} already exists (template planes: {})'.format(
                    name, ', '.join(sorted(BUILTIN_PLANES))))
            seen.add(name.strip().lower())
        if plane.get('axis') not in PLANE_AXES:
            error(path + '.axis', 'must be "x", "y" or "z"')
        number(plane.get('offset'), path + '.offset')
    table = plane_table(schema)

    def known(name, path):
        if not isinstance(name, str) or name.strip().lower() not in table:
            error(path, 'unknown plane {!r}; defined: {}'.format(
                name, ', '.join(sorted(v[0] for v in table.values()))))
            return None
        return table[name.strip().lower()]

    kinds = _schema_parameter_kinds(schema)
    values = dict((p['name'].strip().lower(), p.get('value')) for p in schema.get('parameters') or []
                  if isinstance(p, dict) and isinstance(p.get('name'), str))
    dims = schema.get('dimensions', [])
    if not isinstance(dims, list):
        error('$.dimensions', 'must be an array of {"parameter", "from", "to"}')
        dims = []
    for i, dim in enumerate(dims):
        path = '$.dimensions[{}]'.format(i)
        if not isinstance(dim, dict):
            error(path, 'must be {"parameter", "from", "to"}')
            continue
        pname = dim.get('parameter')
        kind = kinds.get(pname.strip().lower()) if isinstance(pname, str) else None
        if kind is None:
            error(path + '.parameter', 'must name a parameters[] entry of type length')
        elif kind[1] != 'length':
            error(path + '.parameter', '{!r} is a {} parameter; a dimension needs length'.format(
                pname, kind[1]))
        a = known(dim.get('from'), path + '.from')
        b = known(dim.get('to'), path + '.to')
        if a and b:
            if a[1] != b[1]:
                error(path, 'from {!r} (axis {}) and to {!r} (axis {}) must share an axis'.format(
                    a[0], a[1], b[0], b[1]))
            elif abs(a[2] - b[2]) < 1.0:
                error(path, 'from and to are the same plane position')
            elif kind is not None and _finite_number(values.get(pname.strip().lower())) and \
                    abs(abs(a[2] - b[2]) - values[pname.strip().lower()]) > 1.0:
                warn(path, '{!r} is {} mm but the planes are {} mm apart; Revit moves the '
                           '"to" plane to the parameter value'.format(
                               pname, values[pname.strip().lower()], round(abs(a[2] - b[2]), 1)))
    for i, geom in enumerate(schema.get('geometry') or []):
        if not isinstance(geom, dict) or 'locks' not in geom:
            continue
        path = '$.geometry[{}].locks'.format(i)
        locks = geom['locks']
        if not isinstance(locks, dict):
            error(path, 'must be an object such as {"x_min": "Left", "x_max": "Right"}')
            continue
        if geom.get('is_solid') is False:
            warn(path, 'voids are locked too; check the cut after creation')
        for key in sorted(locks):
            if key not in LOCK_KEYS:
                error(path, 'unknown key {!r}; use {}'.format(key, ', '.join(LOCK_KEYS)))
                continue
            plane = known(locks[key], '{}.{}'.format(path, key))
            if plane and plane[1] != key[0]:
                error('{}.{}'.format(path, key), 'plane {!r} is a {} plane; {} needs an {} plane'
                      .format(plane[0], plane[1], key, key[0]))


def _schema_parameter_kinds(schema):
    """{lower name: (name, kind)} for parameters[] and the material parameters."""
    kinds = {}
    for param in schema.get('parameters') or []:
        if isinstance(param, dict) and isinstance(param.get('name'), str) and param['name'].strip():
            ptype = param.get('type', 'length')
            kinds[param['name'].strip().lower()] = (param['name'].strip(), ptype)
    for mat in schema.get('materials') or []:
        if isinstance(mat, dict) and isinstance(mat.get('name'), str) and mat['name'].strip():
            pname = material_parameter_name(mat)
            kinds[pname.lower()] = (pname, 'material')
    return kinds


def _validate_types(schema, error, warn, revit_name, number):
    types = schema['types']
    if not isinstance(types, list):
        error('$.types', 'must be an array of type objects')
        return
    if len(types) > MAX_TYPES:
        error('$.types', 'at most {} types'.format(MAX_TYPES))
    kinds = _schema_parameter_kinds(schema)
    material_names = set(m['name'].strip() for m in schema.get('materials') or []
                         if isinstance(m, dict) and isinstance(m.get('name'), str))
    seen = set()
    for i, ftype in enumerate(types):
        path = '$.types[{}]'.format(i)
        if not isinstance(ftype, dict):
            error(path, 'must be a type object with name and values')
            continue
        name = ftype.get('name')
        if revit_name(name, path + '.name'):
            if name.strip().lower() in seen:
                error(path + '.name', 'duplicates another type {!r}'.format(name))
            seen.add(name.strip().lower())
        unknown = sorted(k for k in ftype if k not in TYPE_FIELDS)
        if unknown:
            warn(path, 'ignored field(s): ' + ', '.join(unknown))
        values = ftype.get('values', {})
        if not isinstance(values, dict):
            error(path + '.values', 'must be an object of parameter name -> value')
            continue
        for key in sorted(values):
            vpath = '{}.values[{}]'.format(path, json.dumps(key, ensure_ascii=False))
            value = values[key]
            known = kinds.get(key.strip().lower()) if isinstance(key, str) else None
            if known is None:
                if isinstance(value, bool) or not (isinstance(value, str) or _finite_number(value)):
                    error(vpath, 'must be a number or a string')
                else:
                    warn(vpath, 'not in parameters[] or a material parameter; it is set only '
                                'if the category template has that parameter')
                continue
            kind = known[1]
            if kind == 'material':
                if not isinstance(value, str) or value.strip() not in material_names:
                    error(vpath, 'must name one of materials[] ({})'.format(
                        ', '.join(sorted(material_names)) or 'none defined'))
            elif kind == 'text':
                if not isinstance(value, str):
                    error(vpath, 'must be a string for a text parameter')
            elif kind == 'integer':
                if isinstance(value, bool) or not isinstance(value, int):
                    error(vpath, 'must be a whole number for an integer parameter')
            else:
                number(value, vpath)


def type_names(schema):
    """Names of the schema's family types, in order (empty when it has none)."""
    if not isinstance(schema, dict) or not isinstance(schema.get('types'), list):
        return []
    return [t['name'].strip() for t in schema['types']
            if isinstance(t, dict) and isinstance(t.get('name'), str) and t['name'].strip()]


def schema_for_type(schema, type_name):
    """A copy of `schema` whose solids show the materials of family type `type_name`.

    Used by the review pane to preview a finish variant: every solid bound to
    a material parameter the type overrides is given the type's material. An
    unknown type (or None) returns the schema unchanged. Pure, never mutates.
    """
    if not isinstance(schema, dict) or not type_name:
        return schema
    chosen = None
    for ftype in schema.get('types') or []:
        if (isinstance(ftype, dict) and isinstance(ftype.get('name'), str)
                and ftype['name'].strip() == type_name):
            chosen = ftype
            break
    if chosen is None or not isinstance(chosen.get('values'), dict):
        return schema
    overrides = {}
    for key, value in chosen['values'].items():
        if isinstance(key, str) and isinstance(value, str):
            overrides[key.strip().lower()] = value.strip()
    by_param = {}
    for mat in schema.get('materials') or []:
        if isinstance(mat, dict) and isinstance(mat.get('name'), str):
            target = overrides.get(material_parameter_name(mat).lower())
            if target:
                by_param[mat['name'].strip()] = target
    if not by_param:
        return schema
    copy = dict(schema)
    geometry = []
    for geom in schema.get('geometry') or []:
        if isinstance(geom, dict) and isinstance(geom.get('material'), str) \
                and geom['material'].strip() in by_param:
            geom = dict(geom)
            geom['material'] = by_param[geom['material'].strip()]
        geometry.append(geom)
    copy['geometry'] = geometry
    return copy


def validate_ai_schema(schema, category):
    """Return actionable JSON-path errors; never change the input object.

    ``category`` is the category the user selected: the schema must match it
    exactly, and it must be one FamiGen has a template for.
    """
    errors, _ = _validate(schema, category, strict_category=True)
    return errors + meaning_errors(schema)


def meaning_errors(schema):
    """What an AI-written family must also say: its concept (FamiGen checks the rest)."""
    if isinstance(schema, dict) and not isinstance(schema.get('concept'), dict):
        return ['$.concept: required - {"what": ..., "use": ..., "placement": ...}; decide what '
                'the object is and how it is used before the geometry']
    return []


def validate_family_schema(schema, category=None):
    """(errors, warnings) for the review pane, MCP proposals and the builder.

    With ``category`` the schema must match it; without, the schema's own
    ``family_category`` must be a supported category. Warnings never block
    creation; they explain what the builder will ignore.
    """
    return _validate(schema, category, strict_category=category is not None)


def schema_summary(schema):
    """Small, JSON-safe overview used by the review pane and MCP replies."""
    if not isinstance(schema, dict):
        return {}
    geometry = [g for g in schema.get('geometry') or [] if isinstance(g, dict)]
    materials = [m for m in schema.get('materials') or [] if isinstance(m, dict)]
    counts = {}
    for geom in geometry:
        name = geom.get('material')
        if isinstance(name, str) and geom.get('is_solid', True) is not False:
            counts[name] = counts.get(name, 0) + 1
    return {
        'family_name': schema.get('family_name'),
        'family_category': schema.get('family_category'),
        'parts': len(geometry),
        'solids': sum(1 for g in geometry if g.get('is_solid', True) is not False),
        'voids': sum(1 for g in geometry if g.get('is_solid', True) is False),
        'materials': [{'name': m.get('name'),
                       'parameter': material_parameter_name(m),
                       'solids': counts.get(m.get('name'), 0)} for m in materials],
        'parameters': len(schema.get('parameters') or []),
        'types': type_names(schema),
        'reference_planes': len(schema.get('reference_planes') or []),
        'dimensions': len(schema.get('dimensions') or []),
        'concept': dict((k, v) for k, v in (schema.get('concept') or {}).items()
                        if k in CONCEPT_FIELDS and isinstance(v, str))
        if isinstance(schema.get('concept'), dict) else {},
        'joints': len(schema.get('joints') or []) if isinstance(schema.get('joints'), list) else 0,
    }


EXAMPLE_SCHEMA = {
    'schema_version': 2,
    'family_name': 'Side Table 500',
    'family_category': 'Furniture',
    'concept': {'what': 'Round side table: an oak top on one steel column with a disc base',
                'use': 'Stands beside a sofa; a person seated reaches the top at 550 mm',
                'placement': 'Free-standing on the floor'},
    'materials': [
        {'name': 'Oak', 'color': '#B08050', 'transparency': 0, 'smoothness': 40,
         'parameter': 'Top Material'},
        {'name': 'Black Steel', 'color': [40, 40, 44], 'shininess': 90,
         'parameter': 'Frame Material'},
    ],
    'parameters': [
        {'name': 'Top Thickness', 'type': 'length', 'value': 30},
        {'name': 'Designer', 'type': 'text', 'value': 'T3Lab'},
    ],
    'types': [
        {'name': 'Oak Top', 'values': {'Top Material': 'Oak'}},
        {'name': 'Steel Top', 'values': {'Top Material': 'Black Steel', 'Top Thickness': 12}},
    ],
    'geometry': [
        {'id': 'Top', 'type': 'Extrusion', 'role': 'surface', 'material': 'Oak', 'subcategory': 'Top',
         'profile': [{'type': 'Circle', 'center': [0, 0, 0], 'radius': 250}],
         'extrusion_start': 520, 'extrusion_end': 550},
        {'id': 'Leg', 'type': 'Cylinder', 'role': 'support', 'material': 'Black Steel',
         'subcategory': 'Legs',
         'start': [0, 0, 0], 'end': [0, 0, 521], 'radius': 25},
        {'id': 'Base', 'type': 'Extrusion', 'role': 'support', 'material': 'Black Steel',
         'subcategory': 'Legs',
         'profile': [{'type': 'Circle', 'center': [0, 0, 0], 'radius': 180}],
         'extrusion_start': 0, 'extrusion_end': 12},
    ],
    'joints': [
        {'a': 'Leg', 'b': 'Top', 'at': [0, 0, 520]},
        {'a': 'Leg', 'b': 'Base', 'at': [0, 0, 12]},
        {'a': 'Base', 'b': 'ground', 'at': [0, 0, 0]},
    ],
}


def schema_contract():
    """JSON-safe description of the contract for external AI clients (MCP)."""
    return {
        'schema_version': SCHEMA_VERSION,
        'units': {'length': 'millimeters', 'angle': 'radians'},
        'supported_categories': list(SUPPORTED_CATEGORIES),
        'form_types': list(FORM_TYPES),
        'form_guide': {
            'Extrusion': 'constant outline: slabs, panels, boards, gables (sketch_plane_x/y)',
            'Blend': 'outline A -> outline B between two heights: tapers, hoppers, chamfered plinths',
            'Revolution': 'turned / round about an axis: domes, bowls, knobs, balusters, rings',
            'Sweep': 'constant section along a path (profile_local: true): rails, mouldings, troughs',
            'SweptBlend': 'section changing along ONE Line/Arc3P path (local profiles): chutes, horns',
            'Cylinder': 'straight rods, posts and (with wall) pipes',
            'Tube': 'curved / 3D rods, rails; with wall: pipes, tube slides',
            'Void': 'is_solid false on any form: openings, recesses, hollow tubes, basins',
        },
        'local_profile': 'u sideways (Z x tangent), v up; points [u, v, 0] in mm',
        'segment_types': list(SEGMENT_TYPES),
        'material_rules': {
            'fields': {'name': 'unique, required', 'color': '"#RRGGBB" or [r,g,b] 0-255, required',
                       'transparency': '0-100 percent, optional (0 = opaque)',
                       'shininess': '0-128, optional', 'smoothness': '0-100, optional',
                       'parameter': 'family parameter name, optional (default "<name> Material")'},
            'binding': ('Every material becomes a Revit Material plus one family type '
                        'parameter (Materials and Finishes) whose default is that '
                        'material; each solid that names it is bound to the parameter.'),
            'geometry_fields': {'material': 'one materials[].name; ignored on voids',
                                'is_solid': 'false = Void Form: cuts the solids it overlaps',
                                'subcategory': 'optional subcategory name under the family category'},
        },
        'parameter_rules': {
            'types': list(PARAMETER_TYPES),
            'default_type': 'length',
            'length_unit': 'millimeters',
            'existing': 'set on the template parameter with the same name',
            'missing': 'created (type parameter unless instance is true)',
            'shared': ('true = added from Revit\'s shared parameter file (by guid, else by name); '
                       'not found there -> a family parameter and a warning'),
        },
        'type_rules': {
            'fields': {'name': 'unique, required', 'values': 'object: parameter name -> value'},
            'values': ('A parameters[].name takes a value of its type (length in mm); a material '
                       'parameter name takes a materials[].name.'),
            'default': 'the first type is the default; without types one type is created',
            'max_types': MAX_TYPES,
            'note': 'geometry is not driven by parameters: a length value is data only',
        },
        'constraint_rules': {
            'reference_planes': '{"name", "axis": "x"|"y"|"z", "offset": mm}; axis x = plane x=offset',
            'template_planes': dict((k, {'axis': v[0], 'offset': v[1]}) for k, v in BUILTIN_PLANES.items()),
            'dimensions': '{"parameter": length parameter, "from": plane, "to": plane} on one axis',
            'locks': 'geometry[].locks {x_min|x_max|y_min|y_max|z_min|z_max: plane}; face must lie on it',
            'effect': 'a labelled dimension + locked faces make the part resize with the parameter',
        },
        'meaning_rules': {
            'concept': ('{"what": the object and its main parts, "use": who or what uses it and the '
                        'route through it, "placement": where it stands or hangs and what it '
                        'connects to}; decide it before the geometry'),
            'roles': {'field': 'geometry[].role', 'values': list(ROLES),
                      'hollow': 'container and conduit must be hollow (wall, inner_loops or a void)',
                      'access': 'stairs, ladders, ramps, slides, bridges: a joint at each end'},
            'hollow': 'Cylinder / Tube "wall": mm thickness (inner radius = radius - wall)',
            'joints': ('[{"a": part id, "b": part id or "ground", "at": [x,y,z]}]: both parts must '
                       'be within 25 mm of the point; a joint that misses is reported'),
            'checked': ('famigen_validate_family and the preview report: Not connected, Joint, '
                        'Hollow, Access and Crossing findings'),
        },
        'detail_levels': {'field': 'geometry[].detail_levels', 'values': list(DETAIL_LEVELS),
                          'default': 'shown at every detail level'},
        'repeat_rules': {'linear': '{"count": n, "step": [dx,dy,dz]}', 'list': '{"at": [[dx,dy,dz], ...]}',
                         'max_copies': MAX_REPEAT, 'not_for': 'Sweep'},
        'forbidden_name_characters': NAME_FORBIDDEN,
        'max_name_length': MAX_NAME_LENGTH,
        'example': EXAMPLE_SCHEMA,
    }


def generate_family_schema(bridge, description, category, overlay='', max_tokens=6000):
    """Generate, validate and optionally repair once, leaving the draft untouched."""
    if bridge is None:
        raise ValueError('AI provider is unavailable. Check AI Mode configuration.')
    if not isinstance(description, str) or not description.strip():
        raise ValueError('Enter a family description before generating JSON.')
    system_prompt = build_system_prompt(category, overlay)
    prompt = description
    for attempt in range(2):
        try:
            result = bridge.ask_json(prompt, system_prompt=system_prompt, max_tokens=max_tokens)
        except Exception as ex:
            raise ValueError('AI provider request failed: {}'.format(ex))   # no "from": IronPython 2
        if result is None:
            raise ValueError('AI provider returned no parseable JSON. Check provider availability or response length and try again.')
        errors = validate_ai_schema(result, category)
        if not errors:
            return result
        if attempt:
            raise ValueError('AI JSON remains invalid after one repair:\n' + '\n'.join(errors))
        prompt = ('Original family description:\n' + description +
                  '\nRepair the previous JSON to satisfy the same contract and design. '
                  'Return the complete corrected JSON object only.\nValidation errors:\n' +
                  '\n'.join(errors) + '\nPrevious JSON:\n' + json.dumps(result, ensure_ascii=False))
    raise AssertionError('Unreachable')
