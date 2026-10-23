# -*- coding: utf-8 -*-
"""
Model Meet — walk camera, screen projection, picking and mesh batching.

Pure Python: no Revit, no WPF. GUI/ModelMeetDialog.py turns the numbers into
a WPF Viewport3D; dev/test_model_meet.py tests them outside Revit.

Coordinates
-----------
Everything here is in SCENE coordinates: Revit internal feet, Z up, minus the
scene origin (the centre of the loaded geometry). WPF renders in single
precision, so a model placed thousands of feet from the internal origin
would shake on screen without that shift. `to_model` / `to_scene` convert at
the edges — presence (other people) and "Open View in Revit" always use model
coordinates, so two people whose scenes have a different origin still meet
at the same spot.

Camera
------
yaw   degrees, 0 = looking along +X, 90 = along +Y (counter-clockwise, plan).
pitch degrees, + looks up, clamped to +/- MAX_PITCH so `forward x up` never
      degenerates.
Walk mode keeps the eye EYE_HEIGHT_FT above the current level; Fly mode moves
freely along the view direction.

Author: T3Lab
"""
from __future__ import division

import math
import zlib
from array import array

FT_PER_M = 1.0 / 0.3048

EYE_HEIGHT_FT = 1.6 * FT_PER_M          # eye above the level in Walk mode
WALK_SPEED_FT = 1.6 * FT_PER_M          # ft/s, a brisk walk
FLY_SPEED_FT = 4.0 * FT_PER_M           # ft/s
RUN_FACTOR = 3.0                        # Shift held
LOOK_DEG_PER_PX = 0.25                  # mouse drag sensitivity
MAX_PITCH = 85.0
FIELD_OF_VIEW = 70.0                    # horizontal, degrees (WPF PerspectiveCamera)
NEAR_PLANE_FT = 0.25                    # ~8 cm; keeps the depth buffer sharp far away
WHEEL_STEP_FT = 1.0 * FT_PER_M          # one mouse-wheel notch
MAX_DT = 0.1                            # s; a stalled frame never teleports you

WALK = "walk"
FLY = "fly"

# Movement intents read from the keyboard (GUI maps keys to these).
FORWARD, BACK, LEFT, RIGHT, UP, DOWN, RUN = (
    "forward", "back", "left", "right", "up", "down", "run")

KEY_INTENTS = {
    "W": FORWARD, "Up": FORWARD,
    "S": BACK, "Down": BACK,
    "A": LEFT, "Left": LEFT,
    "D": RIGHT, "Right": RIGHT,
    "E": UP, "Q": DOWN,
    "LeftShift": RUN, "RightShift": RUN,
}

PLOT_COLOURS = 8                        # T3.Plot.1 .. T3.Plot.8

# How much of an element's mesh is drawn (Snippets/model_meet_proxy.py).
FULL_DETAIL = "full"
PROXY_DETAIL = "proxy"


# ── vectors ──────────────────────────────────────────────────────────────────
def vsub(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def vdot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def vcross(a, b):
    return (a[1] * b[2] - a[2] * b[1],
            a[2] * b[0] - a[0] * b[2],
            a[0] * b[1] - a[1] * b[0])


def vlen(a):
    return math.sqrt(vdot(a, a))


def vnorm(a):
    length = math.sqrt(vdot(a, a))
    if length < 1e-12:
        return (0.0, 0.0, 0.0)
    return (a[0] / length, a[1] / length, a[2] / length)


def direction(yaw, pitch):
    """Unit view vector for yaw/pitch in degrees."""
    cy, sy = math.cos(math.radians(yaw)), math.sin(math.radians(yaw))
    cp, sp = math.cos(math.radians(pitch)), math.sin(math.radians(pitch))
    return (cy * cp, sy * cp, sp)


def view_axes(yaw, pitch):
    """(forward, up) unit vectors, perpendicular — what ViewOrientation3D needs."""
    return Camera((0.0, 0.0, 0.0), yaw, pitch).basis()[0::2]


def yaw_pitch(vector):
    """(yaw, pitch) in degrees of a direction vector; (0, 0) for a zero vector."""
    x, y, z = vector
    horizontal = math.sqrt(x * x + y * y)
    if horizontal < 1e-12 and abs(z) < 1e-12:
        return 0.0, 0.0
    yaw = math.degrees(math.atan2(y, x)) if horizontal > 1e-12 else 0.0
    pitch = math.degrees(math.atan2(z, horizontal))
    return yaw, max(-MAX_PITCH, min(MAX_PITCH, pitch))


def wrap_degrees(angle):
    """Angle folded into (-180, 180]."""
    angle = math.fmod(angle, 360.0)
    if angle <= -180.0:
        angle += 360.0
    elif angle > 180.0:
        angle -= 360.0
    return angle


def to_scene(point, origin):
    return (point[0] - origin[0], point[1] - origin[1], point[2] - origin[2])


def to_model(point, origin):
    return (point[0] + origin[0], point[1] + origin[1], point[2] + origin[2])


# ── levels ───────────────────────────────────────────────────────────────────
class Level(object):
    """A Revit level: name + elevation in scene feet."""

    def __init__(self, name, elevation):
        self.name = name
        self.elevation = float(elevation)

    def __repr__(self):
        return "Level({!r}, {:.2f})".format(self.name, self.elevation)


def level_at(levels, eye_z, eye=EYE_HEIGHT_FT):
    """The level someone whose eye is at `eye_z` stands on, or None.

    Highest level at or below the feet, with half an eye height of slack so a
    camera that sinks a little does not drop a floor. Below every level: the
    lowest one.
    """
    if not levels:
        return None
    ordered = sorted(levels, key=lambda lv: lv.elevation)
    feet = eye_z - eye
    found = ordered[0]
    for lv in ordered:
        if lv.elevation <= feet + eye * 0.5:
            found = lv
        else:
            break
    return found


def neighbour_level(levels, current, step):
    """Level `step` (+1 up / -1 down) from `current` in elevation order."""
    if not levels:
        return None
    ordered = sorted(levels, key=lambda lv: lv.elevation)
    if current is None:
        return ordered[0]
    names = [lv.name for lv in ordered]
    try:
        i = names.index(current.name)
    except ValueError:
        return ordered[0]
    return ordered[max(0, min(len(ordered) - 1, i + step))]


# ── camera ───────────────────────────────────────────────────────────────────
class Camera(object):
    """First-person camera in scene coordinates."""

    def __init__(self, position=(0.0, 0.0, EYE_HEIGHT_FT), yaw=0.0, pitch=0.0, mode=FLY):
        self.x, self.y, self.z = (float(v) for v in position)
        self.yaw = wrap_degrees(float(yaw))
        self.pitch = max(-MAX_PITCH, min(MAX_PITCH, float(pitch)))
        self.mode = mode
        self.ground = None              # level elevation Walk mode stands on
        self.eye = EYE_HEIGHT_FT        # eye above the floor in Walk mode (character height)

    @property
    def position(self):
        return (self.x, self.y, self.z)

    def copy(self):
        cam = Camera(self.position, self.yaw, self.pitch, self.mode)
        cam.ground = self.ground
        cam.eye = self.eye
        return cam

    def forward(self):
        return direction(self.yaw, self.pitch)

    def basis(self):
        """(forward, right, up) unit vectors of the view."""
        f = self.forward()
        r = vnorm(vcross(f, (0.0, 0.0, 1.0)))
        if r == (0.0, 0.0, 0.0):
            r = (math.sin(math.radians(self.yaw)), -math.cos(math.radians(self.yaw)), 0.0)
        u = vcross(r, f)
        return f, r, u

    def look(self, dx_px, dy_px):
        """Mouse drag: right/left turns, up/down tilts (drag up = look up)."""
        self.yaw = wrap_degrees(self.yaw - dx_px * LOOK_DEG_PER_PX)
        self.pitch = max(-MAX_PITCH, min(MAX_PITCH, self.pitch - dy_px * LOOK_DEG_PER_PX))

    def look_at(self, target):
        self.yaw, self.pitch = yaw_pitch(vsub(target, self.position))

    def stand_on(self, level):
        """Walk mode on `level`: eyes `self.eye` above it."""
        self.mode = WALK
        self.ground = level.elevation if level is not None else self.z - self.eye
        self.z = self.ground + self.eye

    def fly(self):
        self.mode = FLY
        self.ground = None

    def step(self, intents, dt):
        """Move for `dt` seconds with the held `intents`. Returns True when moved."""
        dt = max(0.0, min(MAX_DT, dt))
        if not intents or dt <= 0.0:
            return False
        speed = WALK_SPEED_FT if self.mode == WALK else FLY_SPEED_FT
        if RUN in intents:
            speed *= RUN_FACTOR
        distance = speed * dt

        if self.mode == WALK:
            fwd = (math.cos(math.radians(self.yaw)), math.sin(math.radians(self.yaw)), 0.0)
        else:
            fwd = self.forward()
        right = (math.sin(math.radians(self.yaw)), -math.cos(math.radians(self.yaw)), 0.0)

        mx = my = mz = 0.0
        if FORWARD in intents:
            mx, my, mz = mx + fwd[0], my + fwd[1], mz + fwd[2]
        if BACK in intents:
            mx, my, mz = mx - fwd[0], my - fwd[1], mz - fwd[2]
        if RIGHT in intents:
            mx, my = mx + right[0], my + right[1]
        if LEFT in intents:
            mx, my = mx - right[0], my - right[1]
        if self.mode == FLY:
            if UP in intents:
                mz += 1.0
            if DOWN in intents:
                mz -= 1.0
        move = vnorm((mx, my, mz))
        if move == (0.0, 0.0, 0.0):
            return False
        self.x += move[0] * distance
        self.y += move[1] * distance
        self.z += move[2] * distance
        if self.mode == WALK and self.ground is not None:
            self.z = self.ground + self.eye
        return True

    def dolly(self, notches):
        """Mouse wheel: move along the view (Walk: along the floor)."""
        distance = WHEEL_STEP_FT * notches
        if self.mode == WALK:
            self.x += math.cos(math.radians(self.yaw)) * distance
            self.y += math.sin(math.radians(self.yaw)) * distance
        else:
            f = self.forward()
            self.x += f[0] * distance
            self.y += f[1] * distance
            self.z += f[2] * distance


def overview_camera(bbox):
    """Fly camera outside the model's SE corner, looking at its centre.

    bbox = (minx, miny, minz, maxx, maxy, maxz) in scene feet.
    """
    minx, miny, minz, maxx, maxy, maxz = bbox
    cx, cy, cz = (minx + maxx) / 2.0, (miny + maxy) / 2.0, (minz + maxz) / 2.0
    span = max(maxx - minx, maxy - miny, maxz - minz, 10.0)
    cam = Camera((cx + span * 0.75, cy - span * 0.75, cz + span * 0.45), mode=FLY)
    cam.look_at((cx, cy, cz))
    return cam


def meet_point(peer_position, peer_yaw, distance=2.0 * FT_PER_M):
    """Where to stand to greet someone: `distance` in front of them, facing them."""
    f = direction(peer_yaw, 0.0)
    x = peer_position[0] + f[0] * distance
    y = peer_position[1] + f[1] * distance
    cam = Camera((x, y, peer_position[2]), mode=FLY)
    cam.look_at(peer_position)
    cam.pitch = 0.0
    return cam


# ── projection & rays ────────────────────────────────────────────────────────
def project(camera, point, width, height, fov=FIELD_OF_VIEW):
    """Screen (x, y) in pixels of a scene point, or None when behind the eye."""
    if width <= 0 or height <= 0:
        return None
    f, r, u = camera.basis()
    d = vsub(point, camera.position)
    depth = vdot(d, f)
    if depth <= NEAR_PLANE_FT:
        return None
    tan_h = math.tan(math.radians(fov) / 2.0)
    aspect = float(width) / float(height)
    ndc_x = vdot(d, r) / (depth * tan_h)
    ndc_y = vdot(d, u) * aspect / (depth * tan_h)
    return ((ndc_x + 1.0) * 0.5 * width, (1.0 - ndc_y) * 0.5 * height)


def screen_ray(camera, sx, sy, width, height, fov=FIELD_OF_VIEW):
    """(origin, unit direction) of the ray under pixel (sx, sy)."""
    f, r, u = camera.basis()
    tan_h = math.tan(math.radians(fov) / 2.0)
    aspect = float(width) / float(height) if height else 1.0
    ndc_x = 2.0 * sx / float(width) - 1.0 if width else 0.0
    ndc_y = 1.0 - 2.0 * sy / float(height) if height else 0.0
    hx = ndc_x * tan_h
    hy = ndc_y * tan_h / aspect
    d = (f[0] + r[0] * hx + u[0] * hy,
         f[1] + r[1] * hx + u[1] * hy,
         f[2] + r[2] * hx + u[2] * hy)
    return camera.position, vnorm(d)


def ray_box(origin, direction_, bbox):
    """Entry distance of the ray into an axis-aligned box, or None if it misses."""
    tmin, tmax = -1e30, 1e30
    for axis in range(3):
        o, d = origin[axis], direction_[axis]
        lo, hi = bbox[axis], bbox[axis + 3]
        if abs(d) < 1e-12:
            if o < lo or o > hi:
                return None
            continue
        t1, t2 = (lo - o) / d, (hi - o) / d
        if t1 > t2:
            t1, t2 = t2, t1
        tmin, tmax = max(tmin, t1), min(tmax, t2)
        if tmin > tmax:
            return None
    if tmax < 0.0:
        return None
    return max(tmin, 0.0)


def ray_triangle(origin, direction_, a, b, c):
    """Möller–Trumbore: distance to the triangle along the ray, or None."""
    e1 = vsub(b, a)
    e2 = vsub(c, a)
    p = vcross(direction_, e2)
    det = vdot(e1, p)
    if abs(det) < 1e-12:
        return None
    inv = 1.0 / det
    s = vsub(origin, a)
    u = vdot(s, p) * inv
    if u < 0.0 or u > 1.0:
        return None
    q = vcross(s, e1)
    v = vdot(direction_, q) * inv
    if v < 0.0 or u + v > 1.0:
        return None
    t = vdot(e2, q) * inv
    return t if t > NEAR_PLANE_FT else None


class SceneElement(object):
    """One Revit element's triangles, in scene feet.

    positions: array('f') x,y,z,...   indices: array('i') a,b,c,...
    detail: FULL_DETAIL, or PROXY_DETAIL when the mesh is a simplified stand-in;
    source_triangles: triangles of the full mesh (what streaming would bring back).
    source: which model of the project it comes from — 0 the open model, n a
    linked / added one (Snippets/model_meet_scene.py: Source, scene_id).
    """

    __slots__ = ("element_id", "category", "bic", "name", "level", "glass",
                 "positions", "indices", "bbox", "detail", "source_triangles", "source")

    def __init__(self, element_id, category=u"", bic=u"", name=u"", level=u"", glass=False,
                 source=0):
        self.element_id = element_id
        self.source = source
        self.category = category
        self.bic = bic
        self.name = name
        self.level = level
        self.glass = glass
        self.positions = array("f")
        self.indices = array("i")
        self.bbox = None
        self.detail = FULL_DETAIL
        self.source_triangles = 0

    @property
    def triangle_count(self):
        return len(self.indices) // 3

    def add_mesh(self, points, triangles, source=None):
        """Append a mesh: points [(x,y,z)], triangles [(i,j,k)] into `points`.

        `source`: triangles of the full mesh when this one is a proxy of it.
        Triangles touching a non-finite point are dropped: 'nan' in the render
        text makes Point3DCollection.Parse throw on .NET Framework. The point
        itself is kept (indices stay valid) as a copy of a finite one, so it
        never stretches the bounding box to the scene origin."""
        base = len(self.positions) // 3
        bad = set()
        fill = None
        for n, (x, y, z) in enumerate(points):
            if not (math.isfinite(x) and math.isfinite(y) and math.isfinite(z)):
                bad.add(n)
            elif fill is None:
                fill = (x, y, z)
        if bad and fill is None:
            return                      # nothing finite to draw
        for n, p in enumerate(points):
            self.positions.extend(fill if n in bad else p)
        for i, j, k in triangles:
            if bad and (i in bad or j in bad or k in bad):
                continue
            self.indices.extend((base + i, base + j, base + k))
        self.source_triangles += len(triangles) if source is None else source
        if source is not None and source > len(triangles):
            self.detail = PROXY_DETAIL

    def finish(self):
        """Compute the bounding box; False when there is nothing to draw."""
        if not self.indices:
            return False
        xs = self.positions[0::3]
        ys = self.positions[1::3]
        zs = self.positions[2::3]
        self.bbox = (min(xs), min(ys), min(zs), max(xs), max(ys), max(zs))
        return True

    def hit(self, origin, direction_):
        """Nearest triangle hit distance, or None."""
        pos, idx = self.positions, self.indices
        best = None
        for t in range(0, len(idx), 3):
            ia, ib, ic = idx[t] * 3, idx[t + 1] * 3, idx[t + 2] * 3
            d = ray_triangle(origin, direction_,
                             (pos[ia], pos[ia + 1], pos[ia + 2]),
                             (pos[ib], pos[ib + 1], pos[ib + 2]),
                             (pos[ic], pos[ic + 1], pos[ic + 2]))
            if d is not None and (best is None or d < best):
                best = d
        return best


def bbox_centre(bbox):
    return ((bbox[0] + bbox[3]) / 2.0, (bbox[1] + bbox[4]) / 2.0, (bbox[2] + bbox[5]) / 2.0)


def pick(elements, origin, direction_, skip_glass=True):
    """(element, distance) of the nearest element under the ray, or (None, None).

    Broad phase on bounding boxes, nearest box first; a box farther than the
    best triangle hit so far ends the search.
    """
    candidates = []
    for e in elements:
        if e.bbox is None or (skip_glass and e.glass):
            continue
        t = ray_box(origin, direction_, e.bbox)
        if t is not None:
            candidates.append((t, e))
    candidates.sort(key=lambda c: c[0])
    best, best_t = None, None
    for t_box, e in candidates:
        if best_t is not None and t_box > best_t:
            break
        t = e.hit(origin, direction_)
        if t is not None and (best_t is None or t < best_t):
            best, best_t = e, t
    return best, best_t


# ── render text & chunks ─────────────────────────────────────────────────────
# WPF parses Point3DCollection text with the invariant culture, so '.' decimals
# and spaces are safe on every Windows locale, and ONE Parse call replaces
# hundreds of thousands of pythonnet calls (Point3D(...) + Add per vertex).
# The text is a "triangle soup" — three points per triangle, no index list —
# so a chunk is a plain join of cached per-element strings: re-colouring an
# element never re-numbers anyone's indices. WPF shades each triangle flat,
# which suits walls and slabs.
CHUNK_TRIANGLES = 20000         # a chunk rebuilds in a few tens of milliseconds


def soup_text(element):
    """'x y z x y z ...' — the element's triangles, three points each."""
    pos, idx = element.positions, element.indices
    parts = []
    for i in idx:
        j = i * 3
        parts.append(u"%.3f %.3f %.3f" % (pos[j], pos[j + 1], pos[j + 2]))
    return u" ".join(parts)


def mesh_text(element):
    """(positions text, vertex count, indices) of an element for an INDEXED chunk.

    Revit triangulates each face on its own, so the element's vertices are
    shared inside a face and never across two: WPF averages the normals of
    shared vertices, which shades a round column, a curved wall or a terrain
    smooth while every edge between two faces stays crisp. A proxy's vertices
    were merged across faces by clustering — averaged it would look like a
    pebble — so a proxy is written unshared (flat), like the triangle soup."""
    pos, idx = element.positions, element.indices
    if element.detail == PROXY_DETAIL:
        parts = []
        for i in idx:
            j = i * 3
            parts.append(u"%.3f %.3f %.3f" % (pos[j], pos[j + 1], pos[j + 2]))
        return u" ".join(parts), len(idx), list(range(len(idx)))
    parts = [u"%.3f %.3f %.3f" % (pos[j], pos[j + 1], pos[j + 2]) for j in range(0, len(pos) - 2, 3)]
    return u" ".join(parts), len(pos) // 3, list(idx)


def chunk_indices(items):
    """'a b c ...' of a chunk: each element's indices shifted past the vertices
    before it. `items`: [(vertex count, indices)]."""
    out, base = [], 0
    for count, indices in items:
        if base:
            out.append(u" ".join([str(i + base) for i in indices]))
        else:
            out.append(u" ".join(map(str, indices)))
        base += count
    return u" ".join(t for t in out if t)


class Chunker(object):
    """Splits elements, in arrival order, into chunks of <= CHUNK_TRIANGLES.

    add() returns the chunk index the element went to; `of[element_id]` maps
    back. An element bigger than a chunk gets a chunk of its own.
    """

    def __init__(self, limit=CHUNK_TRIANGLES):
        self.limit = limit
        self.chunks = []            # [[SceneElement, ...], ...]
        self.triangles = []         # triangles per chunk
        self.of = {}

    def add(self, element):
        n = element.triangle_count
        if not self.chunks or (self.triangles[-1] and self.triangles[-1] + n > self.limit):
            self.chunks.append([])
            self.triangles.append(0)
        i = len(self.chunks) - 1
        self.chunks[i].append(element)
        self.triangles[i] += n
        self.of[element.element_id] = i
        return i

    def resize(self, element, old_count):
        """The element's mesh changed from `old_count` triangles (detail streaming).

        Returns the chunk indices to redraw: its own, plus the one it moved to
        when it grew its chunk past the limit (a chunk of its own stays put)."""
        i = self.of[element.element_id]
        n = element.triangle_count
        self.triangles[i] += n - old_count
        if n > old_count and self.triangles[i] > self.limit and len(self.chunks[i]) > 1:
            self.chunks[i].remove(element)
            self.triangles[i] -= n
            if i == len(self.chunks) - 1:
                self.chunks.append([])      # never back into the chunk it just left
                self.triangles.append(0)
            return {i, self.add(element)}
        return {i}


# ── how the model is drawn ───────────────────────────────────────────────────
# A clay model in a dark studio, like a product shot (the owner's choice,
# 2026-10-12): every element one soft light grey with a little glow, so no side
# goes murky, shaded smooth on curved faces and crisp between faces (mesh_text);
# a dark sky and floor with a grid. Colour is kept for meaning: people, issues,
# the daylight heat map, the glass. GUI/ModelMeetPanels.WorldTool draws it; its
# other colour modes (By material, Plain) use Snippets/model_meet_world classes.
CLAY = "T3.Canvas"                      # the model (colour mode Clay)
SCENE_GLOW = "T3.Text"                  # emissive lift of the clay
GLASS_TOKEN = "T3.Plot.6"               # translucent glass
SKY_TOP, SKY_BOTTOM = "T3.TextSecondary", "T3.Ink"     # the studio backdrop
SCENE_BACKGROUND = "T3.Text"            # the viewport before the sky is set (XAML)
FLOOR_TOKEN = "T3.Text"                 # the studio floor
GROUND_TOKEN = "T3.TextSecondary"       # its grid lines
SCENE_TOKENS = (CLAY, SCENE_GLOW, GLASS_TOKEN, SKY_TOP, SKY_BOTTOM, SCENE_BACKGROUND, FLOOR_TOKEN,
                GROUND_TOKEN)


# ── the ground ───────────────────────────────────────────────────────────────
GRID_STEPS_M = (1.0, 2.0, 5.0, 10.0, 20.0, 50.0, 100.0, 200.0, 500.0)
GRID_LINES = 40                 # about this many lines across the ground, each way
GRID_MARGIN = 0.5               # the ground runs this share of the model's size past it
GRID_LINE_WIDTH = 0.006         # of the grid step


def grid_step_ft(span_ft, lines=GRID_LINES):
    """A round grid step (1, 2, 5 × 10^n m) that cuts `span_ft` into <= `lines`."""
    for step in GRID_STEPS_M:
        if span_ft / (step * FT_PER_M) <= lines:
            return step * FT_PER_M
    return GRID_STEPS_M[-1] * FT_PER_M


def ground_grid(bbox, z=None):
    """(points, triangles, step_ft) of the ground grid under the model: thin flat
    quads every round step over its footprint plus a margin, at the model's
    bottom. Like a drawing board, it tells up from down and how big things are."""
    x0, y0, z0, x1, y1, _z1 = bbox
    span = max(x1 - x0, y1 - y0, 10.0 * FT_PER_M)
    margin = span * GRID_MARGIN
    step = grid_step_ft(span + 2.0 * margin)
    gx0 = math.floor((x0 - margin) / step) * step
    gx1 = math.ceil((x1 + margin) / step) * step
    gy0 = math.floor((y0 - margin) / step) * step
    gy1 = math.ceil((y1 + margin) / step) * step
    z = (z0 if z is None else z) - 0.03
    half = step * GRID_LINE_WIDTH / 2.0
    points, triangles = [], []

    def quad(ax, ay, bx, by):
        base = len(points)
        points.extend([(ax, ay, z), (bx, ay, z), (bx, by, z), (ax, by, z)])
        triangles.extend([(base, base + 1, base + 2), (base, base + 2, base + 3)])

    nx = int(round((gx1 - gx0) / step))
    ny = int(round((gy1 - gy0) / step))
    for i in range(nx + 1):
        x = gx0 + i * step
        quad(x - half, gy0 - half, x + half, gy1 + half)
    for j in range(ny + 1):
        y = gy0 + j * step
        quad(gx0 - half, y - half, gx1 + half, y + half)
    return points, triangles, step


def compass_heading(yaw):
    """'N', 'NE', ... where a camera with this yaw looks (project north = +Y)."""
    names = (u"E", u"NE", u"N", u"NW", u"W", u"SW", u"S", u"SE")
    return names[int(((wrap_degrees(yaw) % 360.0) + 22.5) // 45.0) % 8]


# ── avatars ──────────────────────────────────────────────────────────────────
def box_mesh(cx, cy, cz, sx, sy, sz):
    """Axis-aligned box centred on (cx, cy, cz) with sizes sx, sy, sz.

    Returns (points, triangles) with outward counter-clockwise winding.
    """
    hx, hy, hz = sx / 2.0, sy / 2.0, sz / 2.0
    p = [(cx - hx, cy - hy, cz - hz), (cx + hx, cy - hy, cz - hz),
         (cx + hx, cy + hy, cz - hz), (cx - hx, cy + hy, cz - hz),
         (cx - hx, cy - hy, cz + hz), (cx + hx, cy - hy, cz + hz),
         (cx + hx, cy + hy, cz + hz), (cx - hx, cy + hy, cz + hz)]
    t = [(0, 2, 1), (0, 3, 2),      # bottom
         (4, 5, 6), (4, 6, 7),      # top
         (0, 1, 5), (0, 5, 4),      # -Y
         (1, 2, 6), (1, 6, 5),      # +X
         (2, 3, 7), (2, 7, 6),      # +Y
         (3, 0, 4), (3, 4, 7)]      # -X
    return p, t


# ── characters ───────────────────────────────────────────────────────────────
BUILDS = (u"standard", u"slim", u"broad")
# Headwear, all Vietnamese: nón lá (conical hat), mũ cối (pith helmet), mũ bảo
# hiểm (the scooter helmet every street wears). Older looks map onto them.
HATS = (u"none", u"nonla", u"mucoi", u"helmet")
LEGACY_HATS = {u"hardhat": u"helmet", u"cap": u"nonla"}
HEIGHTS_M = (1.60, 1.75, 1.90)
# The plush (Snippets/model_meet_plush.py): the fabric of the block, and three 0..1
# settings in steps of 0.1 — how firm it is stuffed, how long its pile, how busy its waddle.
FABRICS = (u"classic", u"bronze", u"lagoon", u"sunset")
FABRIC_LABELS = {u"classic": u"Classic", u"bronze": u"Bronze", u"lagoon": u"Lagoon", u"sunset": u"Sunset"}
# fabric -> (diffuse token, emissive lift token): classic is the white t3lab.space block.
FABRIC_TOKENS = {u"classic": ("T3.Surface", "T3.TextSecondary"), u"bronze": ("T3.Plot.2", "T3.Text"),
                 u"lagoon": ("T3.Plot.6", "T3.Text"), u"sunset": ("T3.Plot.4", "T3.Text")}
PLUSH_DEFAULTS = {"stuffing": 0.5, "pile": 0.4, "waddle": 0.5}


def _unit_setting(value, default):
    """A 0..1 setting from untrusted data, in steps of 0.1 (a peer's slider sends few changes)."""
    try:
        value = float(value)
    except (TypeError, ValueError, OverflowError):
        return default
    if not math.isfinite(value):
        return default
    return round(min(1.0, max(0.0, value)) * 10.0) / 10.0


def clean_look(data, name=u""):
    """A character's look from untrusted data: colour 1..8, build, hat, height (m),
    fabric, stuffing / pile / waddle (0..1). Keys a peer on an older version does
    not send get their defaults; keys it does not know it drops."""
    data = data if isinstance(data, dict) else {}
    try:
        colour = int(data.get("colour"))
    except (TypeError, ValueError):
        colour = 0
    if not 1 <= colour <= PLOT_COLOURS:
        colour = colour_index(name)
    build = data.get("build") if data.get("build") in BUILDS else BUILDS[0]
    hat = data.get("hat")
    try:
        hat = LEGACY_HATS.get(hat, hat)
    except TypeError:                   # unhashable junk from the wire
        hat = None
    hat = hat if hat in HATS else HATS[0]
    try:
        height = float(data.get("height"))
    except (TypeError, ValueError):
        height = HEIGHTS_M[1]
    height = min(HEIGHTS_M, key=lambda h: abs(h - height))
    fabric = data.get("fabric")
    fabric = fabric if isinstance(fabric, str) and fabric in FABRICS else FABRICS[0]
    look = {"colour": colour, "build": build, "hat": hat, "height": height, "fabric": fabric}
    for key, default in PLUSH_DEFAULTS.items():
        look[key] = _unit_setting(data.get(key), default)
    return look


def eye_height_ft(height_m):
    """Eyes sit 12 cm below the top of the head."""
    return (float(height_m) - 0.12) * FT_PER_M


def _merge(parts):
    """One mesh of several: (points, triangles), or (points, triangles, normals)
    when every part carries vertex normals (smooth shading)."""
    points, triangles, normals = [], [], []
    smooth = bool(parts) and all(len(part) == 3 for part in parts)
    for part in parts:
        pts, tris = part[0], part[1]
        base = len(points)
        points.extend(pts)
        triangles.extend((a + base, b + base, c + base) for a, b, c in tris)
        if smooth:
            normals.extend(part[2])
    return (points, triangles, normals) if smooth else (points, triangles)


def _ring(cx, cy, z, rx, ry, segments):
    step = 2.0 * math.pi / segments
    return [(cx + rx * math.cos(i * step), cy + ry * math.sin(i * step), z) for i in range(segments)]


def prism_mesh(cx, cy, z0, z1, r0, r1, segments=12, caps=True):
    """Upright tapered elliptic prism from z0 (radii r0 = (rx, ry)) to z1 (radii r1).
    (points, triangles), outward counter-clockwise winding; caps close both ends."""
    n = segments
    points = _ring(cx, cy, z0, r0[0], r0[1], n) + _ring(cx, cy, z1, r1[0], r1[1], n)
    triangles = []
    for i in range(n):
        j = (i + 1) % n
        triangles.extend([(i, j, n + j), (i, n + j, n + i)])
    if caps:
        bottom, top = len(points), len(points) + 1
        points.extend([(cx, cy, z0), (cx, cy, z1)])
        for i in range(n):
            j = (i + 1) % n
            triangles.extend([(bottom, j, i), (top, n + i, n + j)])
    return points, triangles


def ellipsoid_mesh(cx, cy, cz, rx, ry, rz, rings=6, segments=10, dome=False):
    """Ellipsoid centred on (cx, cy, cz) — or its upper half closed by a flat
    bottom when `dome`. (points, triangles), outward counter-clockwise."""
    n = segments
    points = [(cx, cy, cz + rz)]
    step = (90.0 if dome else 180.0) / rings
    count = rings if dome else rings - 1
    for k in range(1, count + 1):
        phi = math.radians(90.0 - k * step)
        points.extend(_ring(cx, cy, cz + rz * math.sin(phi), rx * math.cos(phi), ry * math.cos(phi), n))
    triangles = []
    for i in range(n):
        triangles.append((0, 1 + i, 1 + (i + 1) % n))
    for k in range(count - 1):
        a, b = 1 + k * n, 1 + (k + 1) * n
        for i in range(n):
            j = (i + 1) % n
            triangles.extend([(a + i, b + i, b + j), (a + i, b + j, a + j)])
    last = 1 + (count - 1) * n
    pole = len(points)
    points.append((cx, cy, cz) if dome else (cx, cy, cz - rz))
    for i in range(n):
        triangles.append((pole, last + (i + 1) % n, last + i))
    return points, triangles


# ── smooth meshes (characters) ───────────────────────────────────────────────
# The model is drawn flat (a triangle soup — every face its own shade, crisp
# walls). A character is drawn SMOOTH: shared vertices with their true surface
# normal, so WPF blends the light across a rounded edge like the soft toy it
# is. Each builder returns (points, triangles, normals).
def _unit(v):
    n = math.sqrt(v[0] * v[0] + v[1] * v[1] + v[2] * v[2])
    return (v[0] / n, v[1] / n, v[2] / n) if n > 0 else (0.0, 0.0, 1.0)


def smooth_rounded_box(cx, cy, cz, sx, sy, sz, radius, steps=5):
    """Box centred on (cx, cy, cz), sizes sx, sy, sz, every edge and corner
    rounded with `radius`: a sphere split into 8 octants pushed out to the box's
    corners — the duplicated rows / columns between octants become the flat
    faces, and every vertex keeps its sphere direction as its normal."""
    r = max(0.0, min(radius, sx / 2.0, sy / 2.0, sz / 2.0))
    hx, hy, hz = sx / 2.0 - r, sy / 2.0 - r, sz / 2.0 - r
    signs = ((1, 1), (-1, 1), (-1, -1), (1, -1))           # quadrants, counter-clockwise
    columns = []
    for q, (xs, ys) in enumerate(signs):
        for k in range(steps + 1):
            a = math.radians(q * 90.0 + k * 90.0 / steps)
            columns.append((math.cos(a), math.sin(a), xs, ys))
    rows = []
    for zs, start_deg in ((1, 90.0), (-1, 0.0)):
        for k in range(steps + 1):
            phi = math.radians(start_deg - k * 90.0 / steps)
            rows.append((math.cos(phi), math.sin(phi), zs))
    points, normals = [], []
    for cp, sp, zs in rows:
        for ca, sa, xs, ys in columns:
            points.append((cx + r * cp * ca + xs * hx, cy + r * cp * sa + ys * hy, cz + r * sp + zs * hz))
            normals.append((cp * ca, cp * sa, sp))
    n = len(columns)
    triangles = []
    for i in range(len(rows) - 1):
        a, b = i * n, (i + 1) * n
        for j in range(n):
            k = (j + 1) % n
            triangles.extend([(a + j, b + j, b + k), (a + j, b + k, a + k)])
    top = [q * (steps + 1) for q in range(4)]                       # the four corners of the top face
    bottom = [(len(rows) - 1) * n + q * (steps + 1) for q in range(4)]
    triangles.extend([(top[0], top[1], top[2]), (top[0], top[2], top[3]),
                      (bottom[0], bottom[2], bottom[1]), (bottom[0], bottom[3], bottom[2])])
    return points, triangles, normals


def rounded_box_mesh(cx, cy, cz, sx, sy, sz, radius, steps=3):
    """smooth_rounded_box without the normals: (points, triangles)."""
    points, triangles, _normals = smooth_rounded_box(cx, cy, cz, sx, sy, sz, radius, steps)
    return points, triangles


def smooth_ellipsoid(cx, cy, cz, rx, ry, rz, rings=8, segments=16, dome=False):
    """ellipsoid_mesh with true normals. A dome's flat bottom is left open — it
    sits on something."""
    points, triangles = ellipsoid_mesh(cx, cy, cz, rx, ry, rz, rings, segments, dome)
    if dome:
        triangles = triangles[:-segments]                       # no bottom cap
    normals = [_unit(((x - cx) / (rx * rx), (y - cy) / (ry * ry), (z - cz) / (rz * rz)))
               if (x, y, z) != (cx, cy, cz) else (0.0, 0.0, -1.0) for x, y, z in points]
    return points, triangles, normals


def smooth_frustum(cx, cy, z0, z1, r0, r1, segments=24):
    """Open cone / frustum: radius r0 at z0, r1 at z1 (r1 = 0 is a cone with a
    sharp tip and smooth sides). Seen from both sides (BackMaterial)."""
    n = segments
    dz, dr = z1 - z0, r1 - r0
    points, normals = [], []
    for z, r in ((z0, r0), (z1, r1)):
        for i in range(n):
            a = 2.0 * math.pi * i / n
            points.append((cx + r * math.cos(a), cy + r * math.sin(a), z))
            normals.append(_unit((math.cos(a) * dz, math.sin(a) * dz, -dr)))
    triangles = []
    for i in range(n):
        j = (i + 1) % n
        triangles.extend([(i, j, n + j), (i, n + j, n + i)])
    return points, triangles, normals


# The T3 figure — a soft white block with two stubby legs, two dot eyes and a
# flat line for a mouth: a friendly building block, drawn the t3lab.space way
# (white, black detail, ONE colour). The colour is the person's T3.Plot colour,
# on a scarf knotted at the side (khăn rằn) and on the hat, so you know who it
# is from afar. Hats are Vietnamese: nón lá, mũ cối, mũ bảo hiểm. Everything is
# smooth and simple — no part has a sharp edge but the tip of the nón lá.
# It is smaller than the person (FIGURE_SCALE of their height): your eye height
# in Walk mode still follows your height, the block only stands for you.
FIGURE_KINDS = ("body", "face", "accent", "hat")
FIGURE_SCALE = 0.62                     # block height / person height
_BLOCK_WIDTH = {u"standard": 1.0, u"slim": 0.84, u"broad": 1.16}    # x the block's height
_HAT_RISE = {u"none": 0.0, u"nonla": 0.40, u"mucoi": 0.36, u"helmet": 0.24}   # x block width


def _figure_frame(look):
    total = FIGURE_SCALE * look["height"]
    width = (total * 0.86) * _BLOCK_WIDTH[look["build"]]
    return total, width


def figure_face(look):
    """Where the face of a figure sits, feet: {'total', 'width', 'depth', 'legs',
    'body_h', 'front', 'eye_z', 'eye_y', 'eye_r'} — figure_parts' own numbers, so
    cheeks and the plush cage line up with the eyes and the block."""
    total, width = _figure_frame(look)
    legs = 0.14 * total
    body_h = total - legs
    m = FT_PER_M
    return {"total": total * m, "width": width * m, "depth": total * 0.80 * m, "legs": legs * m,
            "body_h": body_h * m, "front": total * 0.40 * m, "eye_z": (legs + body_h * 0.44) * m,
            "eye_y": width * 0.17 * m, "eye_r": body_h * 0.038 * m}


def figure_size(look):
    """(height incl. hat, half width incl. a nón lá's brim) of a figure, metres."""
    total, width = _figure_frame(look)
    half = width * (0.64 if look["hat"] == u"nonla" else 0.5)
    return total + _HAT_RISE.get(look["hat"], 0.0) * width, half


def figure_parts(look):
    """{'body', 'face', 'accent', 'hat': (points, triangles, normals)} of a
    character, feet at the origin, facing +X, feet units. 'hat' may be empty.
    body = block and legs (T3.Surface) · face = eyes and mouth (T3.Ink) ·
    accent = the scarf, hat = the hat — both in the person's T3.Plot colour."""
    m = FT_PER_M
    total, width = _figure_frame(look)
    legs = 0.14 * total
    body_h = total - legs
    depth = total * 0.80
    radius = 0.09 * min(width, depth, body_h)        # tight, like a die — not a pillow
    body_z = legs + body_h / 2.0
    leg_w = width * 0.40
    leg_y = width / 2.0 - leg_w / 2.0 - width * 0.03
    leg_h = legs + radius
    body = _merge([
        smooth_rounded_box(0.0, 0.0, body_z * m, depth * m, width * m, body_h * m, radius * m, 5),
        smooth_rounded_box(0.0, leg_y * m, leg_h / 2.0 * m, depth * 0.84 * m, leg_w * m, leg_h * m,
                           radius * 0.8 * m, 3),
        smooth_rounded_box(0.0, -leg_y * m, leg_h / 2.0 * m, depth * 0.84 * m, leg_w * m, leg_h * m,
                           radius * 0.8 * m, 3),
    ])
    front = depth / 2.0
    eye_z = legs + body_h * 0.44
    eye_r = body_h * 0.038
    line = body_h * 0.022
    face = _merge([
        smooth_ellipsoid(front * m, width * 0.17 * m, eye_z * m, eye_r * 0.5 * m, eye_r * m, eye_r * m, 5, 10),
        smooth_ellipsoid(front * m, -width * 0.17 * m, eye_z * m, eye_r * 0.5 * m, eye_r * m, eye_r * m, 5, 10),
        smooth_rounded_box(front * m, 0.0, (eye_z - body_h * 0.075) * m, line * m, width * 0.42 * m,
                           line * m, line * 0.49 * m, 2),                    # the mouth: one flat line
    ])
    # The scarf: a soft band under the face, knotted at the front corner, two tails.
    band_z = legs + body_h * 0.2
    band_h = body_h * 0.09
    knot_y = width * 0.33
    accent = _merge([
        smooth_rounded_box(0.0, 0.0, band_z * m, (depth + 0.02) * m, (width + 0.02) * m, band_h * m,
                           (radius + 0.01) * m, 4),
        smooth_ellipsoid((front + 0.01) * m, knot_y * m, band_z * m, band_h * 0.45 * m, band_h * 0.7 * m,
                         band_h * 0.6 * m, 5, 10),
        smooth_rounded_box((front + 0.012) * m, (knot_y - width * 0.035) * m, (band_z - band_h * 1.1) * m,
                           0.02 * m, width * 0.07 * m, band_h * 1.6 * m, 0.009 * m, 2),
        smooth_rounded_box((front + 0.012) * m, (knot_y + width * 0.045) * m, (band_z - band_h * 0.9) * m,
                           0.02 * m, width * 0.07 * m, band_h * 1.3 * m, 0.009 * m, 2),
    ])
    top = total
    w = width
    if look["hat"] == u"nonla":                     # nón lá: a wide, shallow cone
        hat = _merge([
            smooth_frustum(0.0, 0.0, (top - 0.01) * m, (top + _HAT_RISE[u"nonla"] * w) * m, w * 0.64 * m,
                           0.0, 32),
            smooth_frustum(0.0, 0.0, (top - 0.02) * m, (top - 0.01) * m, w * 0.645 * m, w * 0.64 * m, 32),
        ])
    elif look["hat"] == u"mucoi":                   # mũ cối: a tall dome, a sloping brim, a knob
        dome = 0.30 * w
        base = top + 0.03                           # the brim floats just over the block's top
        hat = _merge([
            smooth_ellipsoid(0.0, 0.0, base * m, w * 0.36 * m, w * 0.36 * m, dome * m, 8, 24, dome=True),
            smooth_frustum(0.0, 0.0, (top + 0.008) * m, (base + 0.012) * m, w * 0.56 * m, w * 0.35 * m, 24),
            smooth_ellipsoid(0.0, 0.0, (base + dome) * m, w * 0.05 * m, w * 0.05 * m, w * 0.03 * m, 4, 12),
        ])
    elif look["hat"] == u"helmet":                  # mũ bảo hiểm: a round shell and a short peak
        hat = _merge([
            smooth_ellipsoid(0.0, 0.0, (top - 0.005) * m, w * 0.41 * m, w * 0.39 * m,
                             (_HAT_RISE[u"helmet"] * w + 0.005) * m, 8, 24, dome=True),
            smooth_ellipsoid(w * 0.40 * m, 0.0, (top + 0.012) * m, w * 0.17 * m, w * 0.30 * m, 0.014 * m, 4, 16),
        ])
    else:
        hat = ([], [], [])
    return {"body": body, "face": face, "accent": accent, "hat": hat}


def avatar_parts(look):
    """((figure points, triangles), (hat points, triangles)) — figure_parts with
    block, face and scarf in one mesh, without normals. The hat may be empty."""
    parts = figure_parts(look)
    figure = _merge([parts["body"], parts["face"], parts["accent"]])
    return (figure[0], figure[1]), (parts["hat"][0], parts["hat"][1])


def index_text(triangles):
    """'a b c a b c ...' — triangle indices for Int32Collection.Parse."""
    return u" ".join(u"%d %d %d" % t for t in triangles)


def points_text(points):
    """'x y z x y z ...' — Point3DCollection / Vector3DCollection.Parse."""
    return u" ".join(u"%.4f %.4f %.4f" % p for p in points)


def soup_from(points, triangles):
    """Triangle-soup text (see soup_text) of a (points, triangles) mesh."""
    return u" ".join(u"%.3f %.3f %.3f" % points[i] for tri in triangles for i in tri)


def colour_index(name):
    """Stable 1..PLOT_COLOURS for a user name (T3.Plot.<n>), same on every PC."""
    key = (name or u"").strip().lower().encode("utf-8")
    return zlib.crc32(key) % PLOT_COLOURS + 1


def lerp_pose(current, target, factor):
    """Ease (x, y, z, yaw) toward target; yaw turns the short way round."""
    x = current[0] + (target[0] - current[0]) * factor
    y = current[1] + (target[1] - current[1]) * factor
    z = current[2] + (target[2] - current[2]) * factor
    dyaw = wrap_degrees(target[3] - current[3])
    return (x, y, z, wrap_degrees(current[3] + dyaw * factor))
