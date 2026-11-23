# -*- coding: utf-8 -*-
"""View sync as a spec: sync_views (Crop Sync + Datum Sync windows, headless;
Snippets/crop_sync.py, Snippets/datum_sync.py). Test: dev/test_tools_views_domain.py."""
from __future__ import unicode_literals

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS

OPERATIONS = ('crop', 'datums')
FLAGS = {'crop': ('shape', 'crop_active', 'crop_visible', 'annotation_crop'),
         'datums': ('grids', 'levels', 'bubbles', 'leaders')}


def to_ints(values):
    """(ints, rejected raw values)."""
    out, bad = [], []
    for v in values if isinstance(values, (list, tuple)) else [values]:
        try:
            out.append(int(v))
        except (TypeError, ValueError):
            bad.append(v)
    return out, bad


def flags(args, op):
    """Every flag of the operation, default True (the windows' defaults)."""
    return dict((f, args.get(f) is not False) for f in FLAGS[op])


def _label(view, eid_value):
    return '{} [{}]'.format(view.Name, eid_value(view.Id))


def _has_crop(view):
    """Crop Sync's own filter: views with a crop region shape manager."""
    if not hasattr(view, 'GetCropRegionShapeManager') or not hasattr(view, 'CropBoxActive'):
        return False
    try:
        return view.GetCropRegionShapeManager() is not None
    except Exception:
        return False


def datum_kind(view, ds):
    if ds.is_plan_view(view):
        return 'plan'
    if ds.is_elevation_or_section(view):
        return 'vertical'
    return None


def _crop_readback(view, data, f):
    lost = []
    if f['crop_active'] and not view.CropBoxActive:
        lost.append('crop region not active')
    if f['crop_visible'] and bool(view.CropBoxVisible) != bool(data.is_visible):
        lost.append('crop region visibility')
    if f['shape'] and data.is_shaped and data.curve_loops:
        # sync_crop_to_view skips a shape Revit calls invalid for this view.
        try:
            shaped = view.GetCropRegionShapeManager().ShapeSet
        except Exception:
            shaped = None
        if shaped is False:
            lost.append('crop shape not applied (not valid in this view)')
    return lost


def _bubble_readback(doc, view, src, ds):
    """Grid bubbles on the target that differ from the source after a sync."""
    lost = []
    try:
        now = ds.get_view_grids_data(doc, view)
    except Exception:
        return lost
    for name, s in src.items():
        t = now.get(name)
        if t is None or not s.curves:
            continue
        for has_s, vis_s, has_t, vis_t, end in ((s.bubble0, s.bubble0_vis, t.bubble0,
                                                 t.bubble0_vis, 'End0'),
                                                (s.bubble1, s.bubble1_vis, t.bubble1,
                                                 t.bubble1_vis, 'End1')):
            if has_s and has_t and bool(vis_s) != bool(vis_t):
                lost.append('grid {} bubble {}'.format(name, end))
    return lost


def sync_views(doc, args):
    from Snippets._compat import disposing, eid_value, make_eid, short_error

    op = '{}'.format(args.get('operation') or '').strip().lower()
    if op not in OPERATIONS:
        return {'error': 'operation must be crop or datums.'}
    f = flags(args, op)
    if args.get('source_view_id') is not None:
        src_ids, _bad = to_ints([args.get('source_view_id')])
        src = doc.GetElement(make_eid(src_ids[0])) if src_ids else None
    else:
        src = doc.ActiveView
    if src is None or not hasattr(src, 'IsTemplate') or src.IsTemplate:
        return {'error': 'source_view_id is not a view in this model (templates do not count).'}
    ids, bad = to_ints(args.get('view_ids') or [])
    if bad or not ids:
        return {'error': 'view_ids must be view element ids or "@set" handles: {}.'.format(
            ', '.join('{}'.format(b) for b in bad[:20]) or 'none given')}

    if op == 'crop':
        from Snippets import crop_sync as cs
        data = cs.get_view_crop_data(src) if _has_crop(src) else None
        if data is None:
            return {'error': 'Source view "{}" has no crop region (sheets, schedules and '
                             'legends have none). Pick a plan, section, elevation or 3D view.'
                    .format(src.Name)}
        summary = {'shape': 'non-rectangular' if data.is_shaped else 'rectangular',
                   'size_mm': [int(round(data.width)), int(round(data.height))]}

        def fits(v):
            return None if _has_crop(v) else 'has no crop region'
    else:
        from Snippets import datum_sync as ds
        kind = datum_kind(src, ds)
        if kind is None:
            return {'error': 'Source view "{}" is not a plan, elevation or section; grids and '
                             'levels are synced between those only.'.format(src.Name)}
        grids = ds.get_view_grids_data(doc, src) if f['grids'] else {}
        levels = ds.get_view_levels_data(doc, src) if f['levels'] and kind == 'vertical' else {}
        summary = {'grids': len(grids), 'levels': len(levels)}
        if not grids and not levels:
            return {'error': 'Source view "{}" shows no grid{} to copy.'.format(
                src.Name, ' or level' if kind == 'vertical' else '')}

        def fits(v):
            return None if datum_kind(v, ds) == kind else 'not a {} view like the source'.format(
                'plan' if kind == 'plan' else 'section / elevation')

    targets, skipped = [], []
    for i in ids:
        v = doc.GetElement(make_eid(i))
        if v is None or not hasattr(v, 'IsTemplate'):
            skipped.append({'view_id': i, 'reason': 'not a view in this model'})
            continue
        reason = ('is a view template' if v.IsTemplate else
                  'is the source view' if eid_value(v.Id) == eid_value(src.Id) else fits(v))
        if reason:
            skipped.append({'view': _label(v, eid_value), 'reason': reason})
        else:
            targets.append(v)
    out = {'operation': op, 'source': _label(src, eid_value), 'source_data': summary,
           'flags': f}
    if skipped:
        out['skipped'] = skipped[:MAX_ROWS]
    if not targets:
        out['error'] = 'No target view can take this sync; see skipped.'
        return out
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False,
                   targets=[_label(v, eid_value) for v in targets][:MAX_ROWS])
        return out

    from Autodesk.Revit.DB import Transaction, TransactionGroup
    updated, failed, lost, counts = [], [], [], {'grid_updates': 0, 'level_updates': 0}
    title = 'Crop Sync' if op == 'crop' else 'Datum Sync'
    try:
        with disposing(TransactionGroup(doc, 'T3Lab: ' + title)) as group:
            group.Start()
            for v in targets:
                label = _label(v, eid_value)
                try:
                    with disposing(Transaction(doc, '{} to {}'.format(title, v.Name))) as t:
                        t.Start()
                        if op == 'crop':
                            changed = cs.sync_crop_to_view(
                                doc, v, data, apply_shape=f['shape'],
                                set_active=f['crop_active'], set_visible=f['crop_visible'],
                                sync_annotation_crop=f['annotation_crop'])
                        else:
                            g = ds.sync_grids_to_view(doc, v, grids, sync_extents=True,
                                                      sync_bubbles=f['bubbles'],
                                                      sync_leaders=f['leaders']) if grids else 0
                            n = ds.sync_levels_to_view(doc, v, levels, sync_extents=True,
                                                       sync_bubbles=f['bubbles'],
                                                       sync_leaders=f['leaders']) if levels else 0
                            counts['grid_updates'] += g
                            counts['level_updates'] += n
                            changed = bool(g or n)
                        t.Commit()
                except Exception as exc:
                    failed.append({'view': label, 'reason': short_error(exc)})
                    continue
                if not changed:
                    skipped.append({'view': label, 'reason': 'nothing to change (no matching '
                                    'crop or datum in this view)'})
                    continue
                updated.append(eid_value(v.Id))
                if op == 'crop':
                    problems = _crop_readback(v, data, f)
                elif f['bubbles'] and grids:
                    problems = _bubble_readback(doc, v, grids, ds)
                else:
                    problems = []
                lost += ['{}: {}'.format(label, p) for p in problems]
            if updated:
                group.Assimilate()
            else:
                group.RollBack()
    except Exception as exc:
        return {'error': '{} stopped: {}. The whole run was undone.'.format(
            title, short_error(exc))}
    out.update(applied=bool(updated), success=bool(updated), updated=len(updated),
               updated_view_ids=updated[:MAX_ROWS])
    if op == 'datums':
        out.update(counts)
    if skipped:
        out['skipped'] = skipped[:MAX_ROWS]
    if failed:
        out['failed'] = failed[:MAX_ROWS]
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    if not updated:
        out['error'] = 'No view was changed; see skipped / failed.'
    return out


SPECS = [
    ToolSpec(
        'sync_views',
        ('Copy one view\'s setup to other views (Crop Sync / Datum Sync windows, headless). '
         'operation=crop: crop region shape, crop on/visible, annotation crop. '
         'operation=datums: grid and level 2D extents, bubbles and leaders (plan to plans, '
         'section/elevation to sections/elevations). Incompatible targets are skipped with a '
         'reason. dry_run lists targets.'),
        sync_views,
        properties={
            'operation': {'type': 'string', 'enum': list(OPERATIONS)},
            'source_view_id': {'type': 'integer', 'description': 'Default: active view.'},
            'view_ids': {'type': 'array', 'items': {'type': ['integer', 'string']},
                         'description': 'Target views; ids or "@set".'},
            'shape': {'type': 'boolean', 'description': 'crop: copy the shape. Flags default true.'},
            'crop_active': {'type': 'boolean'},
            'crop_visible': {'type': 'boolean'},
            'annotation_crop': {'type': 'boolean'},
            'grids': {'type': 'boolean'},
            'levels': {'type': 'boolean', 'description': 'datums: sections/elevations only.'},
            'bubbles': {'type': 'boolean'},
            'leaders': {'type': 'boolean'},
            'dry_run': {'type': 'boolean', 'description': 'List targets, change nothing.'},
        },
        required=['operation', 'view_ids'], effect='write', placement='modify',
        verify=('ids', 'self'),
        title_vi='Đồng bộ crop, grid, level giữa các view',
        keywords=['crop sync', 'datum sync', 'đồng bộ crop', 'đồng bộ grid', 'đồng bộ level',
                  'bubble'],
        ribbon=('CropSync.pushbutton', 'DatumSync.pushbutton')),
]
