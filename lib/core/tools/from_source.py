# -*- coding: utf-8 -*-
"""Model elements from a source as specs: build_from_point_cloud (Point Cloud to
Model, Services/point_cloud_build) and cad_to_elements (CAD to Elements,
Snippets/_cad_plan + _cad_revit).
Doc: docs/agents/tool-specs.md. Test: dev/test_tools_create_domain.py."""
from __future__ import unicode_literals

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS

_TYPE_ORDER = ('Wall', 'Floor', 'Ceiling', 'Column', 'Door', 'Window', 'Stair', 'Roof')


# ── build_from_point_cloud ──────────────────────────────────────────────────

def type_counts(elements):
    out = {}
    for el in elements:
        out[el.Type] = out.get(el.Type, 0) + 1
    return dict((k, out[k]) for k in _TYPE_ORDER if k in out)


def build_from_point_cloud(doc, args):
    from Services import point_cloud_build as pcb
    from Services.point_cloud_analysis import DETECTABLE, mm_to_ft

    kinds = args.get('elements') or list(DETECTABLE)
    unknown = [k for k in kinds if k not in DETECTABLE]
    if unknown:
        return {'error': 'Unknown element kinds: {}.'.format(', '.join(unknown)),
                'accepted': list(DETECTABLE)}
    out = {}
    if args.get('detections'):
        elements, problems = pcb.from_detections(args['detections'])
        if problems:
            out['ignored_rows'] = problems[:MAX_ROWS]
        out['source'] = 'detections'
    else:
        elements, source = pcb.detect(doc, args)
        if elements is None:
            return source
        out['source'] = source
    elements = [e for e in elements if pcb.TYPE_KEYS.get(e.Type) in kinds]
    floor = args.get('min_confidence')
    if floor is not None:
        low = [e for e in elements if pcb.confidence(e) < float(floor)]
        elements = [e for e in elements if e not in low]
        if low:
            out['below_min_confidence'] = len(low)
    if not elements:
        out['error'] = ('Nothing to build: no {} detected{}. Widen region_mm / level_name, '
                        'lower min_confidence, or check the scan with analyze_point_cloud.'
                        .format('/'.join(kinds), ' above min_confidence' if floor else ''))
        return out
    height_mm = float(args.get('wall_height_mm') or 3000.0)
    if height_mm <= 0:
        return {'error': 'wall_height_mm must be more than 0.'}
    counts = type_counts(elements)
    out.update(planned=counts, wall_height_mm=height_mm)
    if counts.get('Stair'):
        out['note'] = ('{} stair(s) are not built (no stair creation yet); model them by hand.'
                       .format(counts['Stair']))
    builder = pcb.ElementBuilder(doc)
    levels = sorted(set(e.LevelName for e in elements))
    missing = [n for n in levels if n not in builder._levels]
    if missing:
        out['levels_not_in_model'] = missing
        out['level_note'] = 'Elements on these levels are built on the lowest level.'
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False, levels=levels)
        return out

    created = []
    summary = builder.build_all(elements, mm_to_ft(height_mm), created=created)
    made = dict((k, v) for k, v in summary['counts'].items() if v)
    out.update(applied=bool(created), success=bool(created), created=made,
               created_count=len(created), created_ids=created[:MAX_ROWS],
               failed=summary['errors'])
    if summary['errors']:
        out['hint'] = ('{} element(s) failed: usually no type loaded for that kind (wall, floor, '
                       'column, door, window or roof family). Load one and build those kinds '
                       'again with elements=[...].'.format(summary['errors']))
    if not created:
        out['error'] = 'No element was created; see failed and hint.'
    return out


# ── cad_to_elements ─────────────────────────────────────────────────────────

def pick_cad(cads, wanted):
    """(cad dict, None) or (None, error dict); by id or name, or the only one."""
    rows = [{'id': c['id'], 'name': c['name']} for c in cads]
    if not cads:
        return None, {'error': 'No imported or linked CAD file in this model. Import or link a '
                               'DWG (Insert > Import CAD / Link CAD), then retry.'}
    if wanted in (None, ''):
        if len(cads) == 1:
            return cads[0], None
        return None, {'error': '{} CAD files in this model: pass cad_id.'.format(len(cads)),
                      'cad_files': rows[:MAX_ROWS]}
    for c in cads:
        if '{}'.format(c['id']) == '{}'.format(wanted).strip():
            return c, None
    from Snippets._cad_plan import _fold
    hits = [c for c in cads if _fold(wanted) in _fold(c['name'])]
    if len(hits) == 1:
        return hits[0], None
    return None, {'error': 'No single CAD file matches "{}".'.format(wanted),
                  'cad_files': rows[:MAX_ROWS]}


def layer_rows(layers):
    """Scan summary per layer: straight segments, closed shapes, curves."""
    rows = []
    for name in sorted(layers, key=lambda s: s.lower()):
        lg = layers[name]
        rows.append({'layer': name, 'lines': len(lg.segments),
                     'closed': len(lg.closed_polys) + len(lg.circles), 'curves': lg.n_curves})
    return rows


def cad_to_elements(doc, args):
    from Snippets import _cad_revit as cr
    from Snippets import _cad_plan as cp
    from Snippets._compat import disposing

    op = '{}'.format(args.get('operation') or 'scan').lower()
    if op not in ('scan', 'create'):
        return {'error': 'operation must be scan or create.'}
    cad, err = pick_cad(cr.get_cad_instances(doc), args.get('cad_id'))
    if err:
        return err
    layers = cr.scan_cad(doc, cad['element'])
    out = {'cad': {'id': cad['id'], 'name': cad['name']}}
    if op == 'scan':
        rows = layer_rows(layers)
        out.update(applied=False, layers=[r for r in rows if r['curves']][:MAX_ROWS],
                   empty_layers=len([r for r in rows if not r['curves']]),
                   levels=[lv['name'] for lv in cr.get_levels(doc)],
                   elements=list(cp.ELEMENTS),
                   next=('cad_to_elements(operation="create", mapping=[{"layers": [...], '
                         '"element": "wall", "type": ..., "level": ..., "options": {...}}], '
                         'dry_run=true). A wrong option lists the options and defaults.'))
        return out

    mapping = args.get('mapping') or []
    if not mapping:
        return {'error': 'create needs mapping: [{"layers": [...], "element": ...}]. '
                         'Run operation="scan" first to see the layers.'}
    look = cp.Lookups(doc)
    plans, rows = [], []
    try:
        for entry, names in cp.group_mapping(mapping):
            missing = [n for n in names if n not in layers]
            if missing or not names:
                return {'error': 'Layer(s) not in {}: {}.'.format(
                    cad['name'], ', '.join(missing) or '(none given)'),
                    'layers': sorted(n for n in layers if layers[n].n_curves)[:MAX_ROWS]}
            plan = cp.plan_entry(doc, look, entry, [layers[n] for n in names])
            plans.append(plan)
            rows.append({'element': entry.get('element'), 'layers': names,
                         'count': plan.count, 'noun': plan.noun,
                         'details': plan.details.split('\n')})
            if plan.count <= 0:
                rows[-1]['empty'] = plan.empty_msg
    except cp.PlanError as ex:
        return {'error': '{}'.format(ex).replace('\n', ' '), 'planned': rows[:MAX_ROWS]}
    total = sum(p.count for p in plans)
    if args.get('dry_run') or not total:
        out.update(dry_run=bool(args.get('dry_run')), applied=False, planned=rows[:MAX_ROWS],
                   would_create=total)
        if not total:
            out['error'] = 'Nothing to create: no element found in the mapped layers; see planned.'
        return out

    from Autodesk.Revit.DB import TransactionGroup
    created = 0
    with disposing(TransactionGroup(doc, 'T3Lab: CAD to Elements')) as group:
        group.Start()
        for plan, row in zip(plans, rows):
            row.pop('details', None)
            if plan.count <= 0:
                row['status'] = 'nothing found'
                continue
            try:
                res = plan.execute(None)
            except Exception as ex:
                from Snippets._compat import short_error
                row.update(status='failed', error='{} (rolled back)'.format(short_error(ex)))
                continue
            row.update(status='ok' if res.created and not res.failed else
                       ('partial' if res.created else 'failed'),
                       created=res.created, failed=res.failed, skipped=res.skipped,
                       warnings_dismissed=res.warnings)
            if res.notes:
                row['notes'] = res.notes[:5]
            if res.errors:
                row['first_error'] = res.errors[0]
            created += res.created
        if created:
            group.Assimilate()
        else:
            group.RollBack()
    out.update(applied=bool(created), success=bool(created), created_count=created,
               results=rows[:MAX_ROWS])
    if not created:
        out['error'] = 'Nothing was created and the model is unchanged; see results.'
    return out


SPECS = [
    ToolSpec(
        'build_from_point_cloud',
        ('Create walls, floors, ceilings, columns, doors, windows and roofs from a point cloud '
         '(Point Cloud to Model, without the window), one undo step. Pass the detections from '
         'detect_point_cloud_elements (edit or filter them first), or the same arguments to '
         'detect again. Lengths in mm. dry_run returns counts per type only.'),
        build_from_point_cloud,
        properties={
            'detections': {'type': 'object', 'description': (
                'detect_point_cloud_elements result or its detections. Omit to detect now.')},
            'point_cloud_id': {'type': 'integer'},
            'elements': {'type': 'array', 'items': {'type': 'string', 'enum': [
                'walls', 'floors', 'ceilings', 'doors', 'windows', 'columns', 'stairs',
                'roof']}, 'description': 'Kinds to build (default all).'},
            'level_name': {'type': 'string'},
            'region_mm': {'type': 'object', 'description': '{"min": [x,y,z], "max": [x,y,z]}'},
            'max_points': {'type': 'integer'},
            'min_confidence': {'type': 'number', 'description': 'Skip below this %.'},
            'wall_height_mm': {'type': 'number', 'description': 'Default 3000.'},
            'dry_run': {'type': 'boolean'},
        },
        effect='write', placement='modeling', verify=('ids', 'must_change'),
        title_vi='Dựng mô hình từ point cloud',
        keywords=['point cloud', 'scan to bim', 'đám mây điểm', 'dựng tường từ scan'],
        ribbon='PointCloud.pushbutton'),
    ToolSpec(
        'cad_to_elements',
        ('Turn the layers of an imported or linked CAD file into walls, floors, ceilings, rooms, '
         'columns, beams, grids, lines, ducts, pipes, cable trays or conduits (CAD to Elements, '
         'without the window). operation=scan lists layers and levels; create takes a mapping, '
         'one undo step. Lengths in mm; options default to the window\'s. dry_run returns the '
         'count per mapping entry.'),
        cad_to_elements,
        properties={
            'operation': {'type': 'string', 'enum': ['scan', 'create'], 'default': 'scan'},
            'cad_id': {'type': ['integer', 'string'], 'description': 'Id or name.'},
            'mapping': {'type': 'array', 'items': {'type': 'object', 'properties': {
                'layers': {'type': 'array', 'items': {'type': 'string'}},
                'element': {'type': 'string'},
                'type': {'type': 'string'}, 'level': {'type': 'string'},
                'options': {'type': 'object'}}},
                'description': 'Entries with the same element, type, level and options '
                               'are planned together.'},
            'dry_run': {'type': 'boolean'},
        },
        # 'modify' too: a tool with an operation enum must reach the ACTION specialist.
        effect='write', placement=('modeling', 'modify'), verify=('must_change',),
        title_vi='Chuyển layer CAD thành đối tượng Revit',
        keywords=['cad', 'dwg', 'layer', 'chuyển cad', 'dựng tường từ cad'],
        ribbon='CADToElements.pushbutton'),
]
