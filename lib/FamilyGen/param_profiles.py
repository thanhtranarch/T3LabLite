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

# (name, axis, why). Axis: x = width (left/right), y = depth (front/back), z = height,
# None = not a single axis (diameters, counts) - graded on presence only.
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
                            ('Sill Height', 'z', 'sill above the level')),
               'optional': (('Frame Width', 'x', 'frame profile'),
                            ('Frame Depth', 'y', 'frame depth in the wall'))},
    'Columns': {'required': (('Width', 'x', 'section width or diameter'),
                             ('Depth', 'y', 'section depth'), ('Height', 'z', 'column height')),
                'optional': (('Base Height', 'z', 'base / plinth'),
                             ('Capital Height', 'z', 'capital'))},
    'Plumbing Fixture': {'required': _WDH, 'optional': (('Mounting Height', 'z', 'rim / outlet height'),)},
    'Electrical Equipment': {'required': _WDH,
                             'optional': (('Mounting Height', 'z', 'bottom above the floor'),)},
    'Electrical Fixture': {'required': _WDH,
                           'optional': (('Mounting Height', 'z', 'centre above the floor'),)},
    'Mechanical Equipment': {'required': (('Length', 'x', 'unit length'), ('Width', 'y', 'unit width'),
                                          ('Height', 'z', 'unit height')),
                             'optional': (('Connection Diameter', None, 'duct / pipe connection'),
                                          ('Base Height', 'z', 'frame / feet'))},
    'Specialty Equipment': {'required': _WDH,
                            'optional': (('Mounting Height', 'z', 'bottom above the floor'),)},
    'Lighting Fixture': {'required': (('Diameter', None, 'shade / trim size (or Width)'),
                                      ('Height', 'z', 'body height or drop')),
                         'optional': (('Length', 'x', 'linear fixtures'),
                                      ('Suspension Length', 'z', 'pendants'))},
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


def profile(category):
    """The category's profile (Generic Model's for unknown categories)."""
    return PROFILES.get(category) or PROFILES['Generic Model']


def _names(name):
    low = name.strip().lower()
    return (low,) + tuple(a for a in ALIASES.get(low, ()))


def evaluate(schema):
    """Grade ``schema`` against its category's profile. Pure; never raises on bad input."""
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
        for name, axis, why in spec[group]:
            found = next((n for n in _names(name) if n in params), None)
            if found is None:
                state = 'missing' if group == 'required' else 'unused'
            elif found in driven_axes and (axis is None or axis in driven_axes[found]):
                state = 'driven'
            else:
                state = 'present'
            rows.append({'name': name, 'axis': axis, 'why': why, 'required': group == 'required',
                         'state': state, 'as': params[found]['name'] if found else u''})
    required = [r for r in rows if r['required']]
    points = sum(2 if r['state'] == 'driven' else 1 if r['state'] == 'present' else 0 for r in required)
    score = int(round(100.0 * points / (2 * len(required)))) if required else 100
    return {'category': category or u'Generic Model', 'score': score, 'rows': rows,
            'driven': sum(1 for r in required if r['state'] == 'driven'),
            'required': len(required)}


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
        parts.append(u'{} ({}{})'.format(label, r['why'], u', along ' + axis if axis else u''))
    return (u'Make this {} family parametric with its standard sizes: {}. Add reference planes on '
            u'the faces those sizes measure, lock the matching part faces to them, and label the '
            u'dimensions with these length parameters so every type flexes. Keep the current '
            u'look and sizes as the defaults.').format(result['category'], u'; '.join(parts))


def template_note(category):
    """The 'make it parametric' sentence a template brief ends with."""
    names = [n for n, _, _ in profile(category)['required']]
    return (u' Make it parametric: reference planes, face locks and labelled dimensions for '
            u'{}.'.format(u', '.join(names)))


def covers_every_category():
    return all(c in PROFILES for c in SUPPORTED_CATEGORIES)
