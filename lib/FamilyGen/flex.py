# -*- coding: utf-8 -*-
"""Flex the preview like Revit flexes a family type (pure Python, no Revit).

The schema's geometry is drawn at the parameters' default values. In Revit a
length parameter resizes the family only through a labelled dimension between
two reference planes, and only the faces locked to a moving plane follow. This
module does the same to the tessellated preview, so editing a parameter in the
Studio shows what Revit will do - and a parameter that drives nothing moves
nothing (lite_guard/test_famigen_studio.py, FlexTests).
"""
from Intelligence.family_schema import BUILTIN_PLANES, plane_table, type_names

TOLERANCE = 0.01          # mm: a plane that moved less than this did not move
MAX_PASSES = 50           # dimension chains are short; this only stops a cycle


def _number(value):
    try:
        out = float(value)
    except (TypeError, ValueError):
        return None
    return out if out == out and abs(out) != float('inf') else None


def parameter_values(schema, type_name=None):
    """{param lower: value} for ``type_name`` (type overrides over defaults)."""
    values = {}
    for param in schema.get('parameters') or []:
        if isinstance(param, dict) and isinstance(param.get('name'), str):
            values[param['name'].strip().lower()] = param.get('value')
    if type_name:
        for ftype in schema.get('types') or []:
            if (isinstance(ftype, dict) and isinstance(ftype.get('name'), str)
                    and ftype['name'].strip() == type_name and isinstance(ftype.get('values'), dict)):
                for key, value in ftype['values'].items():
                    if isinstance(key, str):
                        values[key.strip().lower()] = value
    return values


def plane_offsets(schema, values):
    """{plane lower: (old offset, new offset)} once every labelled dimension takes its value.

    Template planes never move. Of the two planes of a dimension, the one that is
    not a template moves (``to`` when both are free); planes dimensioned from a
    moved plane follow it, as in Revit.
    """
    table = plane_table(schema)
    fixed = set(n.lower() for n in BUILTIN_PLANES)
    links = []
    moving = set()
    for dim in schema.get('dimensions') or []:
        if not isinstance(dim, dict):
            continue
        a = u'{}'.format(dim.get('from') or '').strip().lower()
        b = u'{}'.format(dim.get('to') or '').strip().lower()
        if a not in table or b not in table or table[a][1] != table[b][1] or a == b:
            continue
        if b in fixed and a not in fixed:
            a, b = b, a
        if b in fixed:
            continue                          # two template planes: nothing can move
        value = _number(values.get(u'{}'.format(dim.get('parameter') or '').strip().lower()))
        old = table[b][2] - table[a][2]
        if value is None or value <= 0:
            value = abs(old)
        links.append((a, b, (1.0 if old >= 0 else -1.0) * value))
        moving.add(b)
    new = dict((k, v[2]) for k, v in table.items() if k not in moving)
    for _ in range(MAX_PASSES):
        progressed = False
        for a, b, signed in links:
            if a in new and b not in new:
                new[b] = new[a] + signed
                progressed = True
        if not progressed:
            break
    return dict((k, (v[2], new.get(k, v[2]))) for k, v in table.items())


def flex_model(model, expanded, schema, values):
    """Move the locked faces of ``model``'s meshes in place. Returns the planes that moved.

    Each part stretches linearly along an axis between its min and max faces:
    exact for extrusions and blends drawn along that axis, close enough for the
    rest. Parts without locks stay where they are, as in Revit.
    """
    offsets = plane_offsets(schema, values)
    moved = dict((k, new - old) for k, (old, new) in offsets.items() if abs(new - old) > TOLERANCE)
    if not moved or model is None:
        return {}
    geometry = (expanded.get('geometry') or []) if isinstance(expanded, dict) else []
    by_index = {}
    for mesh in model.meshes:
        by_index.setdefault(mesh.index, []).append(mesh)
    for index, meshes in by_index.items():
        geom = geometry[index] if 0 <= index < len(geometry) else None
        locks = geom.get('locks') if isinstance(geom, dict) else None
        if not isinstance(locks, dict):
            continue
        pts = [p for m in meshes for p in m.positions]
        if not pts:
            continue
        for k, axis in enumerate('xyz'):
            d_min = moved.get(u'{}'.format(locks.get(axis + '_min') or '').strip().lower(), 0.0)
            d_max = moved.get(u'{}'.format(locks.get(axis + '_max') or '').strip().lower(), 0.0)
            if not d_min and not d_max:
                continue
            lo = min(p[k] for p in pts)
            hi = max(p[k] for p in pts)
            span = hi - lo
            for mesh in meshes:
                out = []
                for p in mesh.positions:
                    t = (p[k] - lo) / span if span > 1e-9 else 0.0
                    q = list(p)
                    q[k] = p[k] + d_min + t * (d_max - d_min)
                    out.append(tuple(q))
                mesh.positions = out
            pts = [p for m in meshes for p in m.positions]
    solids = [m for m in model.meshes if m.is_solid] or model.meshes
    pts = [p for m in solids for p in m.positions]
    if pts:
        model.bbox_min = tuple(min(p[i] for p in pts) for i in range(3))
        model.bbox_max = tuple(max(p[i] for p in pts) for i in range(3))
    return moved


def flexed_schema_planes(schema, values):
    """The schema's reference planes at their flexed offsets (for the overlay)."""
    offsets = plane_offsets(schema, values)
    out = dict(schema)
    planes = []
    for plane in schema.get('reference_planes') or []:
        if isinstance(plane, dict) and isinstance(plane.get('name'), str):
            key = plane['name'].strip().lower()
            if key in offsets:
                plane = dict(plane, offset=offsets[key][1])
        planes.append(plane)
    if planes:
        out['reference_planes'] = planes
    return out


def default_type(schema):
    names = type_names(schema)
    return names[0] if names else None
