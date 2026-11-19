# -*- coding: utf-8 -*-
"""Does the family make sense in space? (pure Python, no Revit)

``spatial_issues(schema, expanded, model)`` reads the preview meshes and the
meaning the schema declares (``concept``, ``geometry[].role``, ``joints``,
docs/famigen-studio-spec.md section 5.3) and returns findings, each starting
with one of ``PREFIXES``:

* ``Not connected`` - solids that do not reach the base through parts that
  REALLY touch (surface to surface, not bounding boxes: a diagonal rail next
  to a stair no longer counts as fixed to it).
* ``Joint``    - a declared joint where one of the two parts is not there.
* ``Hollow``   - a container or conduit modelled as a solid lump.
* ``Access``   - a stair / ladder / ramp / slide / bridge without a joint at
  each end.
* ``Crossing`` - something in the bore of a conduit (a post inside a tube
  slide), a conduit whose centre line runs inside a solid, or something in
  the space a person needs above an access part - with no joint between them.

``build_preview`` adds them to its warnings, so the Studio's Issues tab, the
Claude polish round and the MCP tool ``famigen_validate_family`` all see the
same list. Mesh based: it finds what is wrong in the model that will be built;
it does not prove the Revit solids are valid.
"""
import json
import math
import re

from Intelligence.family_schema import GROUND, HOLLOW_ROLES, JOINED_ROLES

PREFIXES = (u'Not connected', u'Joint', u'Hollow', u'Access', u'Crossing')

CONTACT_MM = 10.0        # surfaces closer than this touch (48-gon circles are ~3 mm off)
CONTACT_SCALE = 0.002    # ... or this share of the model's largest size, if larger
JOINT_MM = 25.0          # how near a part must be to a declared joint point
CROSS_MARGIN = 30.0      # deeper than this inside another solid = passing through it
HEAD_FROM, HEAD_TO = 150.0, 1400.0   # the space someone on an access part needs, mm above it
HEAD_STEP = 100.0
HEAD_EDGE = 80.0         # ... measured this far in from the edges of each face
CROSS_SAMPLES = 9
MAX_REPORTED = 8
MAX_EDGES = 600          # per part, for the crossing and the second contact pass
PAIR_BUDGET = 60000      # point-triangle tests per pair in the second contact pass
_SUFFIX = re.compile(r'( joint)? \d+$')
_CACHE = {}


# ── small vector helpers ─────────────────────────────────────────────────────

def _sub(a, b):
    return (a[0] - b[0], a[1] - b[1], a[2] - b[2])


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _dist2_point_triangle(p, a, b, c):
    """Squared distance from p to triangle abc (Ericson, Real-Time Collision Detection)."""
    ab = _sub(b, a)
    ac = _sub(c, a)
    ap = _sub(p, a)
    d1 = _dot(ab, ap)
    d2 = _dot(ac, ap)
    if d1 <= 0.0 and d2 <= 0.0:
        return _dot(ap, ap)
    bp = _sub(p, b)
    d3 = _dot(ab, bp)
    d4 = _dot(ac, bp)
    if d3 >= 0.0 and d4 <= d3:
        return _dot(bp, bp)
    vc = d1 * d4 - d3 * d2
    if vc <= 0.0 and d1 >= 0.0 and d3 <= 0.0:
        v = d1 / (d1 - d3) if d1 != d3 else 0.0
        q = (ap[0] - v * ab[0], ap[1] - v * ab[1], ap[2] - v * ab[2])
        return _dot(q, q)
    cp = _sub(p, c)
    d5 = _dot(ab, cp)
    d6 = _dot(ac, cp)
    if d6 >= 0.0 and d5 <= d6:
        return _dot(cp, cp)
    vb = d5 * d2 - d1 * d6
    if vb <= 0.0 and d2 >= 0.0 and d6 <= 0.0:
        w = d2 / (d2 - d6) if d2 != d6 else 0.0
        q = (ap[0] - w * ac[0], ap[1] - w * ac[1], ap[2] - w * ac[2])
        return _dot(q, q)
    va = d3 * d6 - d5 * d4
    if va <= 0.0 and (d4 - d3) >= 0.0 and (d5 - d6) >= 0.0:
        den = (d4 - d3) + (d5 - d6)
        w = (d4 - d3) / den if den else 0.0
        q = (bp[0] - w * (c[0] - b[0]), bp[1] - w * (c[1] - b[1]), bp[2] - w * (c[2] - b[2]))
        return _dot(q, q)
    den = va + vb + vc
    if not den:                                   # degenerate triangle
        return min(_dot(ap, ap), _dot(bp, bp), _dot(cp, cp))
    v = vb / den
    w = vc / den
    q = (ap[0] - v * ab[0] - w * ac[0], ap[1] - v * ab[1] - w * ac[1], ap[2] - v * ab[2] - w * ac[2])
    return _dot(q, q)


def _bbox(points):
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    zs = [p[2] for p in points]
    return (min(xs), min(ys), min(zs)), (max(xs), max(ys), max(zs))


def _boxes_touch(a, b, tol):
    return all(a[0][k] - tol <= b[1][k] and b[0][k] - tol <= a[1][k] for k in range(3))


def _clip(p, q, lo, hi):
    """The part of segment pq inside the box (Liang-Barsky), or None."""
    t0, t1 = 0.0, 1.0
    for k in range(3):
        d = q[k] - p[k]
        if abs(d) < 1e-9:
            if p[k] < lo[k] or p[k] > hi[k]:
                return None
            continue
        ta = (lo[k] - p[k]) / d
        tb = (hi[k] - p[k]) / d
        if ta > tb:
            ta, tb = tb, ta
        t0 = max(t0, ta)
        t1 = min(t1, tb)
        if t0 > t1:
            return None
    return (tuple(p[k] + (q[k] - p[k]) * t0 for k in range(3)),
            tuple(p[k] + (q[k] - p[k]) * t1 for k in range(3)))


def _along(p, q, count):
    if count < 2:
        return [tuple((p[k] + q[k]) / 2.0 for k in range(3))]
    return [tuple(p[k] + (q[k] - p[k]) * i / (count - 1.0) for k in range(3)) for i in range(count)]


# ── one part ─────────────────────────────────────────────────────────────────

class _Part(object):
    """A preview mesh with what the checks need, computed once."""

    def __init__(self, mesh, geom, owner):
        self.mesh = mesh
        self.geom = geom
        self.owner = owner                       # the id as written (before repeat / Tube)
        self.is_solid = mesh.is_solid
        pos = mesh.positions
        self.tris = [(pos[a], pos[b], pos[c]) for a, b, c in mesh.triangles]
        self.box = _bbox(pos)
        self._points = None
        self._edges = None
        self._inside = None

    @property
    def points(self):
        if self._points is None:
            seen = set()
            for p in self.mesh.positions:
                seen.add((round(p[0], 1), round(p[1], 1), round(p[2], 1)))
            self._points = list(seen)
        return self._points

    @property
    def edges(self):
        """Unique mesh edges (a Cylinder: its axis too), at most MAX_EDGES."""
        if self._edges is None:
            seen = set()
            out = []
            if self.geom.get('type') == 'Cylinder':
                try:
                    out.append((tuple(float(v) for v in self.geom['start']),
                                tuple(float(v) for v in self.geom['end'])))
                except (KeyError, TypeError, ValueError):
                    pass
            for a, b, c in self.mesh.triangles:
                for i, j in ((a, b), (b, c), (c, a)):
                    key = (i, j) if i < j else (j, i)
                    if key not in seen:
                        seen.add(key)
                        out.append((self.mesh.positions[i], self.mesh.positions[j]))
            if len(out) > MAX_EDGES:
                step = len(out) / float(MAX_EDGES)
                out = [out[int(i * step)] for i in range(MAX_EDGES)]
            self._edges = out
        return self._edges

    def contains(self, p):
        """p lies inside this part (known for Extrusion, Cylinder and full Revolution)."""
        if self._inside is None:
            self._inside = _solid_test(self.geom, 0.0) or False
        return bool(self._inside and self._inside(p))

    def distance(self, p, limit):
        """Distance from p to the part (0 inside it), or None when farther than ``limit``."""
        lo, hi = self.box
        if any(p[k] < lo[k] - limit or p[k] > hi[k] + limit for k in range(3)):
            return None
        if self.contains(p):
            return 0.0
        best = limit * limit
        found = False
        for a, b, c in self.tris:
            d2 = _dist2_point_triangle(p, a, b, c)
            if d2 <= best:
                best = d2
                found = True
        return math.sqrt(best) if found else None


# ── contact: do the solids form one object? ──────────────────────────────────

class _Grid(object):
    """Triangles of every part in a multi-level hash grid (big ones on coarse levels)."""

    def __init__(self, parts, tol):
        lo = [min(p.box[0][k] for p in parts) for k in range(3)]
        hi = [max(p.box[1][k] for p in parts) for k in range(3)]
        self.base = max(25.0, max(hi[k] - lo[k] for k in range(3)) / 64.0)
        self.cells = {}
        self.levels = 1
        for index, part in enumerate(parts):
            for a, b, c in part.tris:
                tlo = [min(a[k], b[k], c[k]) - tol for k in range(3)]
                thi = [max(a[k], b[k], c[k]) + tol for k in range(3)]
                level, size = 0, self.base
                while max(thi[k] - tlo[k] for k in range(3)) > 2.0 * size and level < 12:
                    level += 1
                    size *= 4.0
                self.levels = max(self.levels, level + 1)
                rng = [(int(math.floor(tlo[k] / size)), int(math.floor(thi[k] / size)))
                       for k in range(3)]
                item = (a, b, c, tlo, thi)
                for i in range(rng[0][0], rng[0][1] + 1):
                    for j in range(rng[1][0], rng[1][1] + 1):
                        for k in range(rng[2][0], rng[2][1] + 1):
                            cell = self.cells.setdefault((level, i, j, k), {})
                            cell.setdefault(index, []).append(item)

    def near(self, p):
        """[(part index, its triangles in the cell of p)], every level."""
        found = []
        size = self.base
        for level in range(self.levels):
            cell = self.cells.get((level, int(math.floor(p[0] / size)),
                                   int(math.floor(p[1] / size)), int(math.floor(p[2] / size))))
            if cell:
                found.extend(cell.items())
            size *= 4.0
        return found


class _Sets(object):
    def __init__(self, count):
        self.parent = list(range(count))

    def find(self, i):
        parent = self.parent
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def join(self, i, j):
        a, b = self.find(i), self.find(j)
        if a != b:
            self.parent[a] = b


def _crossing_contact(a, b, tol):
    """Second pass: one part set into the other, or faces that cross without a
    vertex near the other part (a plank across two joists; samples each part's
    edges inside the overlap of the two boxes). True when the work is too
    large to decide (lenient)."""
    lo = tuple(max(a.box[0][k], b.box[0][k]) - tol for k in range(3))
    hi = tuple(min(a.box[1][k], b.box[1][k]) + tol for k in range(3))
    tol2 = tol * tol
    for one, other in ((a, b), (b, a)):
        # A part set INTO another (glass in a wall, a net in a deck) is fixed to it
        # although no surface is near: one of its points lies inside the solid.
        if any(other.contains(p) for p in one.points):
            return True
    for one, other in ((a, b), (b, a)):
        tris = [t for t in other.tris
                if all(min(t[0][k], t[1][k], t[2][k]) <= hi[k]
                       and max(t[0][k], t[1][k], t[2][k]) >= lo[k] for k in range(3))]
        if not tris:
            continue
        samples = []
        for p, q in one.edges:
            piece = _clip(p, q, lo, hi)
            if piece:
                samples.extend(_along(piece[0], piece[1], 9))
        if len(samples) * len(tris) > PAIR_BUDGET:
            return True
        for s in samples:
            for ta, tb, tc in tris:
                if _dist2_point_triangle(s, ta, tb, tc) <= tol2:
                    return True
    return False


def _loose_parts(solids, tol):
    """Solids that do not reach the base through parts that really touch."""
    if len(solids) < 2:
        return []
    count = len(solids)
    sets = _Sets(count + 1)                       # the extra node is the base
    floor = min(p.box[0][2] for p in solids)
    for i, part in enumerate(solids):
        if part.box[0][2] <= floor + tol:
            sets.join(i, count)
    grid = _Grid(solids, tol)
    tol2 = tol * tol
    find = sets.find
    for i, part in enumerate(solids):
        mine = find(i)
        for p in part.points:
            for j, tris in grid.near(p):
                if j == i or find(j) == mine:
                    continue
                for a, b, c, tlo, thi in tris:
                    if (p[0] < tlo[0] or p[0] > thi[0] or p[1] < tlo[1] or p[1] > thi[1]
                            or p[2] < tlo[2] or p[2] > thi[2]):
                        continue
                    if _dist2_point_triangle(p, a, b, c) <= tol2:
                        sets.join(i, j)
                        mine = find(i)
                        break
    tried = set()
    changed = True
    while changed:
        changed = False
        base = sets.find(count)
        for i in range(count):
            if sets.find(i) == base:
                continue
            for j in range(count):
                if j == i or sets.find(j) == sets.find(i) or (min(i, j), max(i, j)) in tried:
                    continue
                if not _boxes_touch(solids[i].box, solids[j].box, tol):
                    continue
                tried.add((min(i, j), max(i, j)))
                if _crossing_contact(solids[i], solids[j], tol):
                    sets.join(i, j)
                    changed = True
    base = sets.find(count)
    return [solids[i] for i in range(count) if sets.find(i) != base]


# ── analytic solids for the crossing check ───────────────────────────────────

def _solid_test(geom, shrink=CROSS_MARGIN):
    """A function ``inside(p)`` for an Extrusion or Cylinder shrunk by a margin, or None."""
    from FamilyGen import preview_mesh as pm
    kind = geom.get('type')
    try:
        if kind == 'Cylinder':
            s = tuple(float(v) for v in geom['start'])
            e = tuple(float(v) for v in geom['end'])
            r = float(geom['radius'])
            axis = _sub(e, s)
            length = math.sqrt(_dot(axis, axis))
            if length < 1e-6:
                return None
            unit = tuple(v / length for v in axis)
            margin = min(shrink, r / 2.0, length / 4.0)

            def inside(p):
                rel = _sub(p, s)
                t = _dot(rel, unit)
                if t < margin or t > length - margin:
                    return False
                radial2 = _dot(rel, rel) - t * t
                return radial2 <= (r - margin) ** 2
            return inside
        if kind == 'Extrusion':
            plane = pm.plane_info(geom)
            outer = pm._to_uv(pm.sample_loop(geom.get('profile', []), plane), plane)
            holes = [pm._to_uv(pm.sample_loop(loop, plane), plane)
                     for loop in geom.get('inner_loops') or []]
            if len(outer) < 3:
                return None
            w0 = plane[1] + float(geom.get('extrusion_start', 0.0))
            w1 = plane[1] + float(geom.get('extrusion_end', 1.0))
            if w1 < w0:
                w0, w1 = w1, w0
            margin = min(shrink, (w1 - w0) / 4.0)
            loops = [outer] + [h for h in holes if len(h) >= 3]

            def inside(p):
                (u, v), = pm._to_uv([p], plane)
                w = p[pm._NORMAL_INDEX[plane[0]]]
                if w < w0 + margin or w > w1 - margin:
                    return False
                return _in_loops(u, v, loops)
            return inside
        if kind == 'Revolution':
            # Full turns only: the point's (radius, height) lies in the profile.
            if float(geom.get('end_angle', 6.283185307)) - float(geom.get('start_angle', 0.0)) < 6.1:
                return None
            plane = pm.plane_info(geom)
            a = tuple(float(v) for v in geom['axis_start'])
            b = tuple(float(v) for v in geom['axis_end'])
            axis = _sub(b, a)
            length = math.sqrt(_dot(axis, axis))
            unit = tuple(v / length for v in axis)

            def polar(p):
                rel = _sub(p, a)
                t = _dot(rel, unit)
                return (math.sqrt(max(_dot(rel, rel) - t * t, 0.0)), t)
            loops = [[polar(p) for p in pm.sample_loop(geom.get('profile', []), plane)]]
            loops += [[polar(p) for p in pm.sample_loop(loop, plane)]
                      for loop in geom.get('inner_loops') or []]
            if len(loops[0]) < 3:
                return None

            def inside(p):
                r, t = polar(p)
                return _in_loops(r, t, loops)
            return inside
    except (KeyError, TypeError, ValueError, IndexError, ZeroDivisionError):
        return None
    return None


def _in_loops(u, v, loops):
    """Even-odd: (u, v) is inside the outline and outside its holes."""
    odd = False
    for loop in loops:
        n = len(loop)
        j = n - 1
        for i in range(n):
            (ui, vi), (uj, vj) = loop[i], loop[j]
            if (vi > v) != (vj > v) and u < (uj - ui) * (v - vi) / (vj - vi) + ui:
                odd = not odd
            j = i
    return odd


def _bore_test(geom):
    """``inside(p)`` for the clear inside of a Cylinder pipe (a solid one: its core)."""
    try:
        s = tuple(float(v) for v in geom['start'])
        e = tuple(float(v) for v in geom['end'])
        r = float(geom['radius'])
        wall = float(geom.get('wall') or 0.0)
    except (KeyError, TypeError, ValueError):
        return None
    axis = _sub(e, s)
    length = math.sqrt(_dot(axis, axis))
    if length < 1e-6:
        return None
    unit = tuple(v / length for v in axis)
    clear = (r - wall - min(CROSS_MARGIN, r / 4.0)) if 0.0 < wall < r else 0.6 * r

    def inside(p):
        rel = _sub(p, s)
        t = _dot(rel, unit)
        return 0.0 <= t <= length and _dot(rel, rel) - t * t <= clear * clear
    return inside


def _headroom(part):
    """Vertical segments over the surfaces of an access part someone walks or slides
    on: from ``HEAD_FROM`` to ``HEAD_TO`` mm above interior points of each
    upward or downward facing triangle."""
    segments = []
    for a, b, c in part.tris:
        ab, ac = _sub(b, a), _sub(c, a)
        nz = ab[0] * ac[1] - ab[1] * ac[0]
        nx = ab[1] * ac[2] - ab[2] * ac[1]
        ny = ab[2] * ac[0] - ab[0] * ac[2]
        size = math.sqrt(nx * nx + ny * ny + nz * nz)
        if size < 1e-6 or abs(nz) < 0.5 * size:
            continue
        bc = _sub(c, b)
        sides = [math.sqrt(_dot(e, e)) for e in (bc, ac, ab)]   # opposite a, b, c
        longest = max(sides)
        n = max(3, min(24, int(longest / HEAD_STEP) + 1))
        # Stay HEAD_EDGE in from the triangle's edges: guards and kerbs stand there.
        least = [HEAD_EDGE * side / size if side else 1.0 for side in sides]
        picks = [(i / float(n), j / float(n)) for i in range(1, n) for j in range(1, n - i)]
        picks = [(u, v) for u, v in picks
                 if u >= least[1] and v >= least[2] and 1.0 - u - v >= least[0]]
        if not picks and longest < 3.0 * HEAD_EDGE:
            picks = [(1.0 / 3.0, 1.0 / 3.0)]      # a facet of a curved surface: its centre
        for u, v in picks:
            p = (a[0] + ab[0] * u + ac[0] * v, a[1] + ab[1] * u + ac[1] * v,
                 a[2] + ab[2] * u + ac[2] * v)
            segments.append(((p[0], p[1], p[2] + HEAD_FROM), (p[0], p[1], p[2] + HEAD_TO)))
    if len(segments) > MAX_EDGES:
        step = len(segments) / float(MAX_EDGES)
        segments = [segments[int(i * step)] for i in range(MAX_EDGES)]
    return segments


def _count_inside(segments, box, inside, skip=None, need=3):
    """Sample points of the segments (clipped to ``box``) that pass ``inside``."""
    count = 0
    lo, hi = box
    for p, q in segments:
        piece = _clip(p, q, lo, hi)
        if not piece:
            continue
        for s in _along(piece[0], piece[1], CROSS_SAMPLES):
            if skip and skip(s):
                continue
            if inside(s):
                count += 1
                if count >= need:
                    return count
    return count


# ── the checks ───────────────────────────────────────────────────────────────

def _owner(expanded_id, ids):
    name = expanded_id
    while name not in ids:
        shorter = _SUFFIX.sub('', name)
        if shorter == name:
            return expanded_id
        name = shorter
    return name


def _names(items):
    names = []
    for name in items:
        if name not in names:
            names.append(name)
    shown = u', '.join(u"'{}'".format(n) for n in names[:MAX_REPORTED])
    more = u' and {} more'.format(len(names) - MAX_REPORTED) if len(names) > MAX_REPORTED else u''
    return shown + more, len(names)


def _is_hollow(geom, part_boxes, void_boxes):
    from FamilyGen import preview_mesh as pm
    kind = geom.get('type')
    if kind in ('Sweep', 'SweptBlend'):
        return True                                   # the section can be a channel or a ring
    if geom.get('wall') or geom.get('inner_loops'):
        return True
    if any(_boxes_touch(box, void, 0.0) for box in part_boxes for void in void_boxes):
        return True
    if kind == 'Revolution':
        # A bowl is the wall section revolved: the profile leaves the axis. A
        # solid lump keeps to the axis from bottom to top.
        try:
            plane = pm.plane_info(geom)
            pts = pm.sample_loop(geom.get('profile', []), plane)
            a = tuple(float(v) for v in geom['axis_start'])
            b = tuple(float(v) for v in geom['axis_end'])
            axis = _sub(b, a)
            length = math.sqrt(_dot(axis, axis))
            unit = tuple(v / length for v in axis)
            along = [_dot(_sub(p, a), unit) for p in pts]
            on_axis = [t for p, t in zip(pts, along)
                       if _dot(_sub(p, a), _sub(p, a)) - t * t <= 4.0]
            height = max(along) - min(along)
            if height > 0 and on_axis and (max(on_axis) - min(on_axis)) < 0.6 * height:
                return True
            return not on_axis
        except (KeyError, TypeError, ValueError, ZeroDivisionError, IndexError):
            return True
    return False


def spatial_issues(schema, expanded, model):
    """Findings for a schema, its expanded form and its PreviewModel (see module doc)."""
    if not isinstance(schema, dict) or not model.meshes:
        return []
    geometry = [g for g in (expanded.get('geometry') or [])]
    try:
        key = json.dumps([geometry, schema.get('joints')], sort_keys=True, default=str)
    except (TypeError, ValueError):
        key = None
    if key is not None and key in _CACHE:
        return list(_CACHE[key])
    raw = [g for g in (schema.get('geometry') or []) if isinstance(g, dict)]
    by_id = dict((g['id'], g) for g in raw if isinstance(g.get('id'), str))
    ids = set(by_id)
    parts = []
    for mesh in model.meshes:
        if not mesh.positions or not mesh.triangles or mesh.index >= len(geometry):
            continue
        geom = geometry[mesh.index] if isinstance(geometry[mesh.index], dict) else {}
        parts.append(_Part(mesh, geom, _owner(geom.get('id') or mesh.label, ids)))
    solids = [p for p in parts if p.is_solid]
    if not solids:
        return []
    lo = [min(p.box[0][k] for p in solids) for k in range(3)]
    hi = [max(p.box[1][k] for p in solids) for k in range(3)]
    tol = max(CONTACT_MM, CONTACT_SCALE * max(hi[k] - lo[k] for k in range(3)))
    issues = []

    # 1. one built object
    loose = _loose_parts(solids, tol)
    if loose:
        shown, _ = _names(p.owner for p in loose)
        issues.append(u'Not connected: {} - {} part(s) float: their surfaces do not touch '
                      u'anything that reaches the base. Make every joint overlap 1-5 mm (legs '
                      u'into the floor and into what they carry, rails into posts).'.format(
                          shown, len(loose)))

    # 2. declared joints
    groups = {}
    for part in parts:
        groups.setdefault(part.owner, []).append(part)
    joints = [j for j in (schema.get('joints') or []) if isinstance(j, dict)]
    jointed = set()
    missed = []
    for joint in joints:
        a, b, at = joint.get('a'), joint.get('b'), joint.get('at')
        if not (isinstance(at, list) and len(at) == 3):
            continue
        try:
            at = tuple(float(v) for v in at)
        except (TypeError, ValueError):
            continue
        jointed.add((a, b))
        jointed.add((b, a))
        for name in (a, b):
            if name == GROUND:
                if min(abs(at[2] - lo[2]), abs(at[2])) > JOINT_MM:
                    missed.append(u"({:.0f}, {:.0f}, {:.0f}) is not on the ground (z {:.0f})".format(
                        at[0], at[1], at[2], lo[2]))
                continue
            members = groups.get(name)
            if not members:
                missed.append(u"no part has the id '{}'".format(name))
                continue
            near = [d for d in (m.distance(at, JOINT_MM) for m in members) if d is not None]
            if not near:
                far = [d for d in (m.distance(at, 1e9) for m in members) if d is not None]
                other = b if name == a else a
                missed.append(u"'{}' is {:.0f} mm from its joint with '{}' at ({:.0f}, {:.0f}, {:.0f})"
                              .format(name, min(far) if far else 0.0, other, at[0], at[1], at[2]))
    if missed:
        more = u'; and {} more'.format(len(missed) - MAX_REPORTED) if len(missed) > MAX_REPORTED else u''
        issues.append(u'Joint: {}{}. Both parts must reach the joint point: move or extend the '
                      u'part, or correct the point.'.format(u'; '.join(missed[:MAX_REPORTED]), more))

    # 3. hollow things are hollow
    void_boxes = [p.box for p in parts if not p.is_solid]
    lumps = []
    for geom in raw:
        if geom.get('role') in HOLLOW_ROLES and geom.get('is_solid', True) is not False:
            boxes = [p.box for p in groups.get(geom.get('id'), [])]
            if not _is_hollow(geom, boxes, void_boxes):
                lumps.append(geom.get('id') or geom.get('type'))
    if lumps:
        shown, _ = _names(lumps)
        issues.append(u'Hollow: {} - a container or conduit is modelled as a solid lump. Give a '
                      u'Cylinder / Tube a "wall", draw the wall section, add inner_loops, or cut '
                      u'the inside with a void.'.format(shown))

    # 4. access parts land at both ends
    open_ends = []
    for geom in raw:
        name = geom.get('id')
        if geom.get('role') in JOINED_ROLES and isinstance(name, str):
            if sum(1 for j in joints if name in (j.get('a'), j.get('b'))) < 2:
                open_ends.append(name)
    if open_ends:
        shown, _ = _names(open_ends)
        issues.append(u'Access: {} - a stair, ladder, ramp, slide or bridge needs a joint at each '
                      u'end: where it starts (the edge of the upper level) and where it lands '
                      u'(the ground or the lower level).'.format(shown))

    # 5. routes and bores stay clear
    crossings = []
    tests = {}
    role_of = dict((g['id'], g.get('role')) for g in raw if isinstance(g.get('id'), str))

    voids = [p for p in parts if not p.is_solid]

    def uncut(part, inside):
        """``inside`` less the voids that cut an opening through ``part``."""
        cuts = [v for v in voids if _boxes_touch(v.box, part.box, 0.0)]
        if not inside or not cuts:
            return inside
        return lambda p: inside(p) and not any(v.contains(p) for v in cuts)

    def test_for(part):
        if id(part) not in tests:
            tests[id(part)] = uncut(part, _solid_test(part.geom))
        return tests[id(part)]

    def near_end(ends, reach):
        reach2 = reach * reach
        return lambda p: any(_dot(_sub(p, e), _sub(p, e)) <= reach2 for e in ends)

    for geom in raw:
        name, role = geom.get('id'), geom.get('role')
        if role not in ('access', 'conduit') or not isinstance(name, str):
            continue
        members = [m for m in groups.get(name, []) if m.is_solid]
        others = [o for o in solids if o.owner != name and (name, o.owner) not in jointed
                  and role_of.get(o.owner) not in ('fixing', 'control', 'guard')]
        blocked = []
        if role == 'conduit':
            # Something in the bore, or the centre line inside a solid. The two
            # ends are where a pipe joins or comes out of things: not counted.
            ends = []
            try:
                if geom.get('type') == 'Tube' and not geom.get('closed'):
                    ends = [tuple(float(v) for v in geom['points'][0]),
                            tuple(float(v) for v in geom['points'][-1])]
                elif geom.get('type') == 'Cylinder':
                    ends = [tuple(float(v) for v in geom['start']),
                            tuple(float(v) for v in geom['end'])]
                skip = near_end(ends, max(2.0 * float(geom.get('radius', 0.0)), 100.0)) if ends else None
            except (KeyError, TypeError, ValueError, IndexError):
                skip = None
            for mine in members:
                if mine.geom.get('type') != 'Cylinder':
                    continue
                bore = _bore_test(mine.geom)
                axis = mine.edges[:1]
                for other in others:
                    if other.owner in blocked or not _boxes_touch(mine.box, other.box, 0.0):
                        continue
                    hits = _count_inside(other.edges, mine.box, uncut(other, bore), skip) if bore else 0
                    inside_other = test_for(other) if hits < 3 else None
                    if inside_other:
                        hits += _count_inside(axis, other.box, inside_other, skip)
                    if hits >= 3:
                        blocked.append(other.owner)
        else:
            # Something in the space a person needs above the surface.
            for mine in members:
                segments = _headroom(mine)
                if not segments:
                    continue
                lo_h = (mine.box[0][0], mine.box[0][1], mine.box[0][2] + HEAD_FROM)
                hi_h = (mine.box[1][0], mine.box[1][1], mine.box[1][2] + HEAD_TO)
                for other in others:
                    if other.owner in blocked or not _boxes_touch((lo_h, hi_h), other.box, 0.0):
                        continue
                    inside_other = test_for(other)
                    if inside_other and _count_inside(segments, other.box, inside_other) >= 3:
                        blocked.append(other.owner)
        if blocked:
            shown, _ = _names(blocked)
            crossings.append(u"{} is in the way of '{}'".format(shown, name))
    if crossings:
        more = u'; and {} more'.format(len(crossings) - MAX_REPORTED) if len(crossings) > MAX_REPORTED else u''
        issues.append(u'Crossing: {}{}. The inside of a pipe or tube and the space above a stair, '
                      u'ramp, slide or bridge must stay clear: move or reroute one of them, or list '
                      u'a joint if the two really connect there.'.format(
                          u'; '.join(crossings[:MAX_REPORTED]), more))

    if key is not None:
        if len(_CACHE) > 16:
            _CACHE.clear()
        _CACHE[key] = list(issues)
    return issues
