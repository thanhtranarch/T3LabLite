# -*- coding: utf-8 -*-
"""Which parameters each family category should expose, and how well a family does.

A profile lists the sizes a person expects to change on a family of that
category - the ones Revit users schedule, tag or flex - each with the axis it
resizes. ``evaluate()`` grades a family against its category's profile:

* ``driven``  - the parameter exists AND a labelled dimension on the right
  axis uses it: changing it resizes the family (in Revit and in the Studio);
* ``present`` - the parameter exists but drives nothing yet;
* ``missing`` - not in the family (``unused`` for an optional size).

Score = (2 × driven + present) / (2 × required), 0-100. ``brief()`` turns the
gaps into one message for Claude, so "Make it parametric" asks for exactly the
category's sizes. Pure Python; lite_guard/test_famigen_studio.py (ProfileTests)
checks every supported category has a profile.

Plan and rationale: docs/famigen-param-profiles.md.
"""
from Intelligence.family_schema import SUPPORTED_CATEGORIES

# (name, axis, why[, scope]). Axis: x = width (left/right), y = depth (front/back), z = height,
# None = not a single axis (diameters, counts) - graded on presence only. Scope: 'type'
# (default: every placed copy shares it) or 'instance' (set per placement: sill and
# mounting heights, suspension lengths).
_WDH = (('Width', 'x', 'overall width'), ('Depth', 'y', 'overall depth'),
        ('Height', 'z', 'overall height'))

PROFILES = {
    'Generic Model': {'required': _WDH, 'optional': ()},
    'Furniture': {'required': _WDH,
                  'optional': (('Seat Height', 'z', 'seating: top of the seat'),
                               ('Top Thickness', 'z', 'tables and desks'))},
    'Casework': {'required': _WDH,
                 'optional': (('Plinth Height', 'z', 'toe kick'),
                              ('Top Thickness', 'z', 'worktop'))},
    'Door': {'required': (('Width', 'x', 'rough / leaf width'), ('Height', 'z', 'leaf height'),
                          ('Thickness', 'y', 'leaf thickness')),
             'optional': (('Frame Width', 'x', 'frame / jamb width'),
                          ('Handle Height', 'z', 'handle above the floor'))},
    'Window': {'required': (('Width', 'x', 'opening width'), ('Height', 'z', 'opening height'),
                            ('Sill Height', 'z', 'sill above the level', 'instance')),
               'optional': (('Frame Width', 'x', 'frame profile'),
                            ('Frame Depth', 'y', 'frame depth in the wall'))},
    'Columns': {'required': (('Width', 'x', 'section width or diameter'),
                             ('Depth', 'y', 'section depth'), ('Height', 'z', 'column height')),
                'optional': (('Base Height', 'z', 'base / plinth'),
                             ('Capital Height', 'z', 'capital'))},
    'Plumbing Fixture': {'required': _WDH, 'optional': (('Mounting Height', 'z', 'rim / outlet height', 'instance'),)},
    'Electrical Equipment': {'required': _WDH,
                             'optional': (('Mounting Height', 'z', 'bottom above the floor', 'instance'),)},
    'Electrical Fixture': {'required': _WDH,
                           'optional': (('Mounting Height', 'z', 'centre above the floor', 'instance'),)},
    'Mechanical Equipment': {'required': (('Length', 'x', 'unit length'), ('Width', 'y', 'unit width'),
                                          ('Height', 'z', 'unit height')),
                             'optional': (('Connection Diameter', None, 'duct / pipe connection'),
                                          ('Base Height', 'z', 'frame / feet'))},
    'Specialty Equipment': {'required': _WDH,
                            'optional': (('Mounting Height', 'z', 'bottom above the floor', 'instance'),)},
    'Lighting Fixture': {'required': (('Diameter', None, 'shade / trim size (or Width)'),
                                      ('Height', 'z', 'body height or drop')),
                         'optional': (('Length', 'x', 'linear fixtures'),
                                      ('Suspension Length', 'z', 'pendants', 'instance'))},
    'Site': {'required': _WDH, 'optional': ()},
    'Entourage': {'required': (('Width', 'x', 'overall width'), ('Height', 'z', 'overall height')),
                  'optional': (('Depth', 'y', 'overall depth'),)},
    'Pipe Accessory': {'required': (('Nominal Diameter', None, 'DN of the pipe'),
                                    ('Length', 'x', 'face to face')),
                       'optional': (('Flange Diameter', None, 'flanged ends'),
                                    ('Operator Height', 'z', 'handwheel / lever above the axis'))},
    'Duct Accessory': {'required': (('Duct Width', 'x', 'duct width (or Diameter)'),
                                    ('Duct Height', 'z', 'duct height'),
                                    ('Length', 'y', 'casing length')),
                       'optional': (('Flange Width', None, 'flanges'),)},
    'Sprinkler': {'required': (('Thread Diameter', None, 'connection size'),
                               ('Height', 'z', 'overall height')),
                  'optional': (('Deflector Diameter', None, 'deflector'),)},
}

# Other spellings Claude or a user may use for the same size.
ALIASES = {
    'width': ('w', 'overall width', 'nominal width', 'leaf width', 'opening width', 'duct width'),
    'depth': ('d', 'overall depth', 'thickness'),
    'height': ('h', 'overall height', 'leaf height', 'opening height', 'duct height'),
    'length': ('l', 'overall length', 'face to face', 'face-to-face length'),
    'diameter': ('dia', 'shade diameter', 'trim diameter', 'size'),
    'nominal diameter': ('dn', 'nominal size', 'pipe size', 'diameter', 'size'),
    'thread diameter': ('thread size', 'connection diameter', 'diameter', 'size'),
    'thickness': ('leaf thickness', 'door thickness'),
}


def _office_profiles(path=None):
    """{category: profile} from the office's JSON, read on each call (small file).

    ``%APPDATA%\\T3LabAI\\famigen_profiles.json``: {"Door": {"required": [["Width", "x",
    "rough width"], ...], "optional": [...]}}. A category given there replaces the
    built-in one; a bad file is ignored.
    """
    import io
    import json
    import os
    path = path or os.path.join(os.environ.get('APPDATA') or os.path.expanduser('~'),
                                'T3LabAI', 'famigen_profiles.json')
    try:
        with io.open(path, 'r', encoding='utf-8') as handle:
            data = json.loads(handle.read())
    except (IOError, OSError, ValueError):
        return {}
    out = {}
    for category, spec in (data.items() if isinstance(data, dict) else []):
        if not isinstance(spec, dict):
            continue
        rows = {}
        for group in ('required', 'optional'):
            items = []
            for item in spec.get(group) or []:
                if isinstance(item, (list, tuple)) and item and isinstance(item[0], str):
                    axis = item[1] if len(item) > 1 and item[1] in ('x', 'y', 'z') else None
                    why = item[2] if len(item) > 2 and isinstance(item[2], str) else u''
                    scope = item[3] if len(item) > 3 and item[3] in SCOPES else 'type'
                    items.append((item[0], axis, why, scope))
            rows[group] = tuple(items)
        if rows.get('required'):
            out[category] = rows
    return out


SCOPES = ('type', 'instance')


def _entries(spec, group):
    """(name, axis, why, scope) for every size of one group; scope defaults to type."""
    for item in spec.get(group) or ():
        yield (item[0], item[1], item[2], item[3] if len(item) > 3 else 'type')


def profile(category):
    """The category's profile: the office's JSON first, then the built-in one
    (Generic Model's for unknown categories)."""
    office = _office_profiles()
    return office.get(category) or PROFILES.get(category) or PROFILES['Generic Model']


def _names(name):
    low = name.strip().lower()
    return (low,) + tuple(a for a in ALIASES.get(low, ()))


def evaluate(schema, shared_defs=None):
    """Grade ``schema`` against its category's profile. Pure; never raises on bad input.

    ``shared_defs``: FamilyGen/shared_params definitions of Revit's shared
    parameter file; rows then say whether a size is in that file and shared."""
    from FamilyGen import shared_params
    category = schema.get('family_category') if isinstance(schema, dict) else None
    spec = profile(category)
    params = {}
    for p in (schema.get('parameters') if isinstance(schema, dict) else None) or []:
        if isinstance(p, dict) and isinstance(p.get('name'), str):
            params[p['name'].strip().lower()] = p
    driven_axes = {}
    try:
        from Intelligence.family_schema import plane_table
        table = plane_table(schema)
    except Exception:
        table = {}
    for dim in (schema.get('dimensions') if isinstance(schema, dict) else None) or []:
        if not isinstance(dim, dict):
            continue
        a = table.get(u'{}'.format(dim.get('from') or '').strip().lower())
        name = u'{}'.format(dim.get('parameter') or '').strip().lower()
        if a and name:
            driven_axes.setdefault(name, set()).add(a[1])
    rows = []
    for group in ('required', 'optional'):
        for name, axis, why, scope in _entries(spec, group):
            found = next((n for n in _names(name) if n in params), None)
            if found is None:
                state = 'missing' if group == 'required' else 'unused'
            elif found in driven_axes and (axis is None or axis in driven_axes[found]):
                state = 'driven'
            else:
                state = 'present'
            param = params.get(found) if found else None
            label = param['name'] if param else name
            in_file = None
            if shared_defs is not None:
                in_file = shared_params.find(shared_defs, label, (param or {}).get('guid'),
                                             (param or {}).get('type', 'length')) is not None
            rows.append({'name': name, 'axis': axis, 'why': why, 'required': group == 'required',
                         'state': state, 'as': param['name'] if param else u'',
                         'scope': scope,
                         'instance': bool(param.get('instance')) if param else None,
                         'scope_ok': None if param is None else
                         bool(param.get('instance')) == (scope == 'instance'),
                         'shared': bool(param.get('shared')) if param else False,
                         'in_file': in_file})
    required = [r for r in rows if r['required']]
    points = sum(2 if r['state'] == 'driven' else 1 if r['state'] == 'present' else 0 for r in required)
    score = int(round(100.0 * points / (2 * len(required)))) if required else 100
    return {'category': category or u'Generic Model', 'score': score, 'rows': rows,
            'driven': sum(1 for r in required if r['state'] == 'driven'),
            'required': len(required),
            'scope_issues': sum(1 for r in rows if r['scope_ok'] is False),
            'shareable': sum(1 for r in rows if r['as'] and r['in_file'] and not r['shared']),
            'shared_file': shared_defs is not None}


def apply_standards(schema, shared_defs=None):
    """(new schema, changes): profile scopes on the category's sizes, and every
    parameter found in the shared parameter file marked ``shared`` with its GUID.

    Nothing else changes; no AI needed. Pure: ``schema`` is not modified."""
    import copy
    from FamilyGen import shared_params
    out = copy.deepcopy(schema)
    changes = []
    if not isinstance(out, dict):
        return out, changes
    params = [p for p in out.get('parameters') or [] if isinstance(p, dict)
              and isinstance(p.get('name'), str)]
    by_name = dict((p['name'].strip().lower(), p) for p in params)
    spec = profile(out.get('family_category'))
    for group in ('required', 'optional'):
        for name, _axis, _why, scope in _entries(spec, group):
            found = next((n for n in _names(name) if n in by_name), None)
            if found is None:
                continue
            param = by_name[found]
            want = scope == 'instance'
            if bool(param.get('instance')) != want:
                param['instance'] = want
                changes.append(u'{} -> {} parameter'.format(param['name'], scope))
    for param in params:
        if param.get('shared') or not shared_defs:
            continue
        item = shared_params.find(shared_defs, param['name'], param.get('guid'),
                                  param.get('type', 'length'))
        if item is not None:
            param['shared'] = True
            param['guid'] = item['guid']
            changes.append(u'{} -> shared parameter ({})'.format(param['name'], item['group'] or
                                                                 u'shared parameter file'))
    return out, changes


def brief(schema):
    """One message asking Claude to give the family its category's sizes (empty when done)."""
    result = evaluate(schema)
    todo = [r for r in result['rows'] if r['required'] and r['state'] != 'driven']
    if not todo:
        return u''
    parts = []
    for r in todo:
        axis = {'x': u'left/right (X)', 'y': u'front/back (Y)', 'z': u'up (Z)'}.get(r['axis'])
        label = r['as'] or r['name']
        parts.append(u'{} ({}{}, {} parameter)'.format(label, r['why'], u', along ' + axis if axis
                                                       else u'', r['scope']))
    return (u'Make this {} family parametric with its standard sizes: {}. Add reference planes on '
            u'the faces those sizes measure, lock the matching part faces to them, and label the '
            u'dimensions with these length parameters so every type flexes. Keep the current '
            u'look and sizes as the defaults.').format(result['category'], u'; '.join(parts))


def template_note(category):
    """The 'make it parametric' sentence a template brief ends with."""
    names = [entry[0] + (u' (instance)' if entry[3] == 'instance' else u'')
             for entry in _entries(profile(category), 'required')]
    return (u' Make it parametric: reference planes, face locks and labelled dimensions for '
            u'{}.'.format(u', '.join(names)))


def covers_every_category():
    return all(c in PROFILES for c in SUPPORTED_CATEGORIES)
