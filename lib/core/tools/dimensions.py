# -*- coding: utf-8 -*-
"""Annotation tools as specs: auto_dimension (Auto Dimension window, headless;
Snippets/_auto_dimension). Doc: docs/agents/tool-specs.md.
Test: dev/test_tools_create_domain.py."""
from __future__ import unicode_literals

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS, norm as _norm

TARGETS = ('grids', 'walls', 'structural_columns', 'columns', 'windows', 'doors', 'lifts',
           'facade')
DEFAULT_TARGETS = TARGETS[:-1]          # the window's ticks: everything but facade chains
WALL_FACES = {'exterior': 'ext', 'interior': 'int', 'both': 'both'}
MM = 1.0 / 304.8


def read_options(args):
    """The window's run options from plain args: (dict, None) or (None, error).
    Lengths in feet, as Snippets/_auto_dimension.dim_one_view takes them."""
    targets = args.get('targets') or list(DEFAULT_TARGETS)
    bad = [t for t in targets if t not in TARGETS]
    if bad:
        return None, 'Unknown targets: {}. Use: {}.'.format(', '.join(bad), ', '.join(TARGETS))
    faces = WALL_FACES.get('{}'.format(args.get('wall_faces') or 'exterior').lower())
    direction = '{}'.format(args.get('direction') or 'both').lower()
    if faces is None or direction not in ('both', 'x', 'y'):
        return None, 'wall_faces is exterior|interior|both and direction is both|x|y.'
    try:
        offsets = [float(v) for v in (args.get('offsets_mm') or [1000.0])]
        min_seg = float(300.0 if args.get('min_segment_mm') is None else args['min_segment_mm'])
    except (TypeError, ValueError):
        return None, 'offsets_mm and min_segment_mm must be numbers (mm).'
    if len(offsets) not in (1, 3) or min(offsets) <= 0:
        return None, 'offsets_mm is [L2] or [L1, L2, L3], each more than 0 mm.'
    if len(offsets) == 1:                    # the window's single offset: L1/L2/L3 = 0.5/1/1.5
        l1, l2, l3 = offsets[0] * 0.5, offsets[0], offsets[0] * 1.5
    else:
        l1, l2, l3 = offsets
    check = bool(args.get('check_mode'))
    opts = {
        'do_walls': 'walls' in targets, 'do_struct_cols': 'structural_columns' in targets,
        'do_arch_cols': 'columns' in targets, 'do_grids': 'grids' in targets,
        'do_windows': 'windows' in targets, 'do_doors': 'doors' in targets,
        'do_lifts': 'lifts' in targets, 'do_facade': 'facade' in targets,
        'wall_mode': faces, 'check_mode': check, 'ck_off': 300.0 * MM,
        'is_3level': len(offsets) == 3,
        'l1_feet': l1 * MM, 'l2_feet': l2 * MM, 'l3_feet': l3 * MM,
        'run_x': direction != 'y', 'run_y': direction != 'x',
        'both_sides': (not check) and args.get('both_sides') is not False,
        'check_min': min_seg > 0, 'min_seg_feet': min_seg * MM,
        'facade_off': 1200.0 * MM, 'facade_tol': 1800.0 * MM,
    }
    return opts, None


def pick_dim_type(types, wanted):
    """((name, DimensionType), None) or (None, error); types = [(name, type)]."""
    if not types:
        return None, 'The model has no linear dimension type. Load one, then retry.'
    if not wanted:
        return types[0], None
    for name, dt in types:
        if _norm(name) == _norm(wanted):
            return (name, dt), None
    return None, 'No linear dimension type "{}". Available: {}.'.format(
        wanted, ', '.join(n for n, _t in types[:30]))


def auto_dimension(doc, args):
    from Snippets import _auto_dimension as ad
    from Snippets._compat import disposing, eid_value, make_eid

    opts, problem = read_options(args)
    if problem:
        return {'error': problem}
    picked, problem = pick_dim_type(ad.linear_dimension_types(doc), args.get('dim_type'))
    if problem:
        return {'error': problem}
    ids = args.get('view_ids') or []
    views, skipped = [], []
    for raw in ids:
        try:
            v = doc.GetElement(make_eid(int(raw)))
        except (TypeError, ValueError):
            v = None
        if v is None or not hasattr(v, 'ViewType'):
            skipped.append({'id': raw, 'reason': 'not a view'})
        elif getattr(v, 'IsTemplate', False) or not ad._is_valid_view(v):
            skipped.append({'id': raw, 'reason': 'not a plan, section, elevation or detail view'})
        else:
            views.append(v)
    if not ids:
        v = doc.ActiveView
        if v is not None and ad._is_valid_view(v):
            views.append(v)
    if not views:
        return {'error': 'No view to dimension. Pass view_ids of plan, section or elevation '
                         'views, or activate one.', 'skipped_views': skipped[:MAX_ROWS]}
    dim_type = picked[1]
    out = {'dim_type': picked[0],
           'targets': sorted(k for k in TARGETS if opts.get(_TARGET_FLAG[k]))}
    if skipped:
        out['skipped_views'] = skipped[:MAX_ROWS]

    if args.get('dry_run'):
        out.update(dry_run=True, applied=False, views=[preview(doc, ad, v, opts) for v in views][
            :MAX_ROWS], note=('Counts are the elements each view shows; how many strings come out '
                              'is known only after the run (references Revit rejects are dropped).'))
        return out

    errors, rows, made = [], [], []
    from Autodesk.Revit.DB import TransactionGroup
    with disposing(TransactionGroup(doc, 'T3Lab: Auto Dimension')) as group:
        group.Start()
        for v in views:
            n, attempts, fails, dims = ad.dim_one_view(
                doc, v, opts['do_walls'], opts['do_struct_cols'], opts['do_arch_cols'],
                opts['do_grids'], opts['do_windows'], opts['do_doors'], opts['do_lifts'],
                opts['do_facade'], opts['wall_mode'], dim_type, opts['check_mode'],
                opts['ck_off'], opts['is_3level'], opts['l1_feet'], opts['l2_feet'],
                opts['l3_feet'], opts['l1_feet'] * 0.5, opts['run_x'], opts['run_y'],
                opts['both_sides'], opts['check_min'], opts['min_seg_feet'],
                opts['facade_off'], opts['facade_tol'], on_error=errors.append)
            row = {'view': v.Name, 'id': eid_value(v.Id), 'dimensions': n,
                   'attempts': attempts}
            if fails:
                row['rejected'] = fails[:3]
            rows.append(row)
            made += dims
        if made:
            group.Assimilate()
        else:
            group.RollBack()
    short = 0
    if opts['check_min']:
        for d in made:
            try:
                short += len([s for s in d.Segments
                              if s.Value is not None and s.Value < opts['min_seg_feet']])
            except Exception:
                pass
    out.update(applied=bool(made), success=bool(made), created_count=len(made),
               dimension_ids=[eid_value(d.Id) for d in made][:MAX_ROWS], views=rows[:MAX_ROWS])
    if errors:
        out['errors'] = errors[:5]
    if short:
        out['warning'] = ('{} segment(s) shorter than {:g} mm: text may overlap. Run again with '
                          'larger offsets_mm after undoing.'.format(short,
                                                                   opts['min_seg_feet'] / MM))
    if not made:
        out['error'] = ('No dimension was created: a plan view needs at least two grids, and '
                        'the chosen targets must show in the view. See views.')
    return out


_TARGET_FLAG = {'grids': 'do_grids', 'walls': 'do_walls', 'structural_columns': 'do_struct_cols',
                'columns': 'do_arch_cols', 'windows': 'do_windows', 'doors': 'do_doors',
                'lifts': 'do_lifts', 'facade': 'do_facade'}


def preview(doc, ad, view, opts):
    """What a run would measure in one view: counts per chosen target."""
    from Snippets._compat import eid_value
    row = {'view': view.Name, 'id': eid_value(view.Id), 'view_type': '{}'.format(view.ViewType)}
    grids = ad.count_in_view(doc, view, 'OST_Grids')
    row['grids'] = grids
    section = row['view_type'] in ('Section', 'Elevation')
    if section:
        row['levels'] = ad.count_in_view(doc, view, 'OST_Levels')
        row['measures'] = ('grids and levels' if opts['do_grids'] else
                           'nothing: section and elevation views take only grids')
        return row
    for key, bic in ad.TARGET_CATEGORIES:
        if opts.get(_TARGET_FLAG[key]):
            row[key] = ad.count_in_view(doc, view, bic)
    if grids < 1:
        row['measures'] = 'nothing: a plan view needs grids'
    return row


SPECS = [
    ToolSpec(
        'auto_dimension',
        ('Add chain dimensions for grids, walls, columns, windows, doors and lifts in plan '
         'views, and grids / levels in sections (Auto Dimension, without the window); one undo '
         'step. Plan views need grids. Lengths in mm. dry_run lists what each view would '
         'measure.'),
        auto_dimension,
        properties={
            'view_ids': {'type': 'array', 'items': {'type': ['integer', 'string']},
                         'description': 'Default: active view.'},
            'targets': {'type': 'array', 'items': {'type': 'string', 'enum': list(TARGETS)},
                        'description': 'Default: all but facade.'},
            'dim_type': {'type': 'string', 'description': 'Linear dimension type name.'},
            'offsets_mm': {'type': 'array', 'items': {'type': 'number'},
                           'description': '[L2] (L1=L2/2, L3=1.5*L2) or [L1, L2, L3]. '
                                          'Default [1000].'},
            'wall_faces': {'type': 'string', 'enum': ['exterior', 'interior', 'both']},
            'direction': {'type': 'string', 'enum': ['both', 'x', 'y']},
            'both_sides': {'type': 'boolean', 'description': 'Default true.'},
            'check_mode': {'type': 'boolean', 'description': 'Check dims to the nearest grid.'},
            'min_segment_mm': {'type': 'number', 'description': 'Warn below (default 300).'},
            'dry_run': {'type': 'boolean'},
        },
        effect='write', placement='modify', verify=('ids', 'must_change'),
        title_vi='Ghi kích thước tự động',
        keywords=['dim', 'dimension', 'kích thước', 'đo kích thước', 'ghi kích thước'],
        ribbon='AutoDimension.pushbutton'),
]
