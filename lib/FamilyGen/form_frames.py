# -*- coding: utf-8 -*-
"""Local cross-sections for Sweep and Swept Blend (pure Python, no Revit).

A profile drawn "looking along the path" is far easier for a person or an AI
than one drawn in world coordinates. FamiGen's convention, shared by the
preview (``preview_mesh``) and the Revit builder (``builder``):

* at a path point with unit tangent ``t``: ``u = normalize(Z x t)`` (sideways,
  horizontal) and ``v = t x u`` (up). For a vertical path, ``u = +X``.
* a local profile point ``[a, b, 0]`` sits at ``origin + a*u + b*v``.

``local_to_world`` turns local segments into world segments the builder can
draw without a sketch plane: Line, Arc3P and Spline only (a Circle becomes two
Arc3P halves, an Ellipse four Arc3P quarters - a close approximation).
"""
import math


def _sub(a, b):
    return [a[0] - b[0], a[1] - b[1], a[2] - b[2]]


def _add(a, b):
    return [a[0] + b[0], a[1] + b[1], a[2] + b[2]]


def _mul(a, s):
    return [a[0] * s, a[1] * s, a[2] * s]


def _cross(a, b):
    return [a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0]]


def _dot(a, b):
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _unit(a):
    length = math.sqrt(_dot(a, a))
    if length < 1e-12:
        raise ValueError('zero-length direction')
    return [a[0] / length, a[1] / length, a[2] / length]


def frame(tangent):
    """(u, v) for a unit path tangent: u sideways (horizontal), v up."""
    t = _unit(tangent)
    if abs(t[2]) > 0.999:
        u = [1.0, 0.0, 0.0]
    else:
        u = _unit(_cross([0.0, 0.0, 1.0], t))
    v = _unit(_cross(t, u))
    return u, v


def _arc_center(p0, p1, pm):
    """Centre of the circle through three 3D points (raises when collinear)."""
    a, b = _sub(p1, p0), _sub(pm, p0)
    n = _cross(a, b)
    nn = _dot(n, n)
    if nn < 1e-9:
        raise ValueError('Arc3P points are collinear')
    term = _add(_mul(_cross(n, a), _dot(b, b)), _mul(_cross(b, n), _dot(a, a)))
    return _add(p0, _mul(term, 0.5 / nn)), _unit(n)


def segment_ends(seg):
    """((start point, start tangent), (end point, end tangent)) of a path segment."""
    kind = seg.get('type')
    if kind == 'Line':
        s, e = [float(c) for c in seg['start']], [float(c) for c in seg['end']]
        t = _unit(_sub(e, s))
        return (s, t), (e, t)
    if kind in ('Arc3P', 'ArcThreePoint'):
        s, e, m = ([float(c) for c in seg[k]] for k in ('start', 'end', 'mid'))
        center, normal = _arc_center(s, e, m)
        ts = _unit(_cross(normal, _sub(s, center)))
        te = _unit(_cross(normal, _sub(e, center)))
        if _dot(ts, _sub(m, s)) < 0:          # orient along start -> mid -> end
            ts, te = _mul(ts, -1.0), _mul(te, -1.0)
        return (s, ts), (e, te)
    if kind == 'Spline':
        pts = [[float(c) for c in p] for p in seg['points']]
        return (pts[0], _unit(_sub(pts[1], pts[0]))), (pts[-1], _unit(_sub(pts[-1], pts[-2])))
    raise ValueError('path segments must be Line, Arc3P or Spline here')


def path_ends(path):
    """Start and end (point, tangent) of a path given as segments."""
    first, _ = segment_ends(path[0])
    _, last = segment_ends(path[-1])
    return first, last


def _place(p, origin, u, v):
    a, b = float(p[0]), float(p[1])
    return [round(origin[k] + a * u[k] + b * v[k], 4) for k in range(3)]


def local_to_world(segments, origin, u, v):
    """Local profile segments -> world Line / Arc3P / Spline segments."""
    out = []

    def arc3(c, rx, ry, a0, a1):
        am = (a0 + a1) / 2.0
        pt = lambda a: [c[0] + rx * math.cos(a), c[1] + ry * math.sin(a), 0.0]  # noqa: E731
        out.append({'type': 'Arc3P', 'start': _place(pt(a0), origin, u, v),
                    'end': _place(pt(a1), origin, u, v), 'mid': _place(pt(am), origin, u, v)})

    for seg in segments or []:
        kind = seg.get('type')
        if kind == 'Line':
            out.append({'type': 'Line', 'start': _place(seg['start'], origin, u, v),
                        'end': _place(seg['end'], origin, u, v)})
        elif kind in ('Arc3P', 'ArcThreePoint'):
            out.append({'type': 'Arc3P', 'start': _place(seg['start'], origin, u, v),
                        'end': _place(seg['end'], origin, u, v), 'mid': _place(seg['mid'], origin, u, v)})
        elif kind == 'Spline':
            out.append({'type': 'Spline', 'points': [_place(p, origin, u, v) for p in seg['points']]})
        elif kind in ('Circle', 'Arc', 'Ellipse'):
            c = [float(x) for x in seg['center']]
            rx = float(seg.get('radius_x', seg.get('radius')))
            ry = float(seg.get('radius_y', seg.get('radius')))
            if kind == 'Circle' or ('start_angle' not in seg and 'end_angle' not in seg):
                pieces = 2 if kind != 'Ellipse' else 4
                step = 2 * math.pi / pieces
                for i in range(pieces):
                    arc3(c, rx, ry, i * step, (i + 1) * step)
            else:
                a0, a1 = float(seg.get('start_angle', 0.0)), float(seg.get('end_angle', 2 * math.pi))
                pieces = max(1, int(math.ceil(abs(a1 - a0) / (math.pi / 2))))
                for i in range(pieces):
                    arc3(c, rx, ry, a0 + (a1 - a0) * i / pieces, a0 + (a1 - a0) * (i + 1) / pieces)
        else:
            raise ValueError('unsupported profile segment {!r}'.format(kind))
    return out


def profile_at(segments, point, tangent):
    """Local profile placed at a path point (world segments)."""
    u, v = frame(tangent)
    return local_to_world(segments, point, u, v)
