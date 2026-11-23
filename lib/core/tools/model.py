# -*- coding: utf-8 -*-
"""Model-edit tools as specs: auto_join (AutoJoin), room_to_floor (Room to Floor),
rehost_to_level (Wall Adjust Base), cut_wall_openings (Wall Cut Profile),
create_door_thresholds (Door Threshold). Each wraps the logic its window runs
(Services/join_service.py, Snippets/finish_floors.py, Snippets/wall_tools.py).
Test: dev/test_tools_model_domain.py.
"""
from __future__ import unicode_literals

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS

_IDS = {'type': 'array', 'items': {'type': ['integer', 'string']}}


def int_ids(value):
    """(ids, bad) from ints or digit strings ("@set" is expanded by the server)."""
    if value is None:
        return [], []
    if not isinstance(value, (list, tuple)):
        value = [value]
    ids, bad = [], []
    for v in value:
        try:
            i = int('{}'.format(v).strip())
        except (TypeError, ValueError):
            bad.append('{}'.format(v))
            continue
        if i not in ids:
            ids.append(i)
    return ids, bad


def pick_name(items, wanted, name_of):
    """First item whose name equals `wanted` (case-insensitive), else whose name
    contains it, else whose id equals it; None when nothing matches."""
    from Snippets._compat import eid_value, fold
    key = fold('{}'.format(wanted).strip())
    for test in (lambda i: fold(name_of(i)) == key, lambda i: key in fold(name_of(i)),
                 lambda i: '{}'.format(eid_value(i.Id)) == key):
        hits = [i for i in items if test(i)]
        if hits:
            return hits[0]
    return None


def parse_rules(raw, categories, default):
    """(rules, problems): [{priority, join_with}] checked against the join categories."""
    if not raw:
        return [dict(r) for r in default], []
    by_key = dict((c.lower(), c) for c in categories)
    rules, problems = [], []
    for r in raw if isinstance(raw, (list, tuple)) else [raw]:
        r = r if isinstance(r, dict) else {}
        pair = [by_key.get('{}'.format(r.get(k) or '').strip().lower())
                for k in ('priority', 'join_with')]
        if None in pair:
            problems.append('{} cuts {}'.format(r.get('priority'), r.get('join_with')))
        else:
            rules.append({'priority': pair[0], 'join_with': pair[1]})
    return rules, problems


def _elements(doc, ids):
    from Snippets._compat import make_eid
    return [(i, doc.GetElement(make_eid(i))) for i in ids]


def _group(doc, name):
    from Autodesk.Revit.DB import TransactionGroup
    from Snippets._compat import disposing
    return disposing(TransactionGroup(doc, name))


def _transaction(doc, name):
    from Autodesk.Revit.DB import Transaction
    from Snippets._compat import disposing
    return disposing(Transaction(doc, name))


def _name(el):
    from Snippets._compat import elem_name
    try:
        return elem_name(el) or ''
    except Exception:
        return ''


# ── auto_join ───────────────────────────────────────────────────────────────

def auto_join(doc, args):
    from Services import join_service as js

    mode = '{}'.format(args.get('mode') or 'join').lower()
    if mode not in ('join', 'unjoin'):
        return {'error': 'mode must be join or unjoin.'}
    rules, problems = parse_rules(args.get('rules'), js.CATEGORY_NAMES, js.DEFAULT_RULES)
    if problems:
        return {'error': 'Unknown category in rule(s): {}.'.format('; '.join(problems)),
                'categories': js.CATEGORY_NAMES}
    ids, bad = int_ids(args.get('element_ids'))
    if bad:
        return {'error': 'element_ids must be element ids: {}.'.format(', '.join(bad[:10]))}
    if args.get('element_ids') is not None:
        scope = 'Selected Elements'
        if not ids:
            return {'error': 'element_ids is empty. Give the elements to {}, or omit it '
                             'and set scope.'.format(mode)}
    else:
        scope_arg = '{}'.format(args.get('scope') or 'project').lower()
        if scope_arg not in ('project', 'active_view'):
            return {'error': 'scope must be project or active_view.'}
        scope = 'Entire Project' if scope_arg == 'project' else 'Active View'
    title = mode.capitalize()
    out = {'mode': mode, 'scope': scope, 'rules': ['{} cuts {}'.format(r['priority'],
                                                                        r['join_with'])
                                                    for r in rules]}
    if args.get('dry_run'):
        rows = js.count_pairs(rules, scope, title, doc=doc, selected_ids=ids or None)
        out.update(dry_run=True, applied=False, per_rule=rows,
                   pairs_to_change=sum(r['to_change'] for r in rows))
        return out

    stats = {}
    changed, skipped, errors, message = js.run_join(
        rules, scope, title, args.get('switch_order', True) is not False, doc=doc,
        protect_embedded=args.get('protect_embedded', True) is not False, stats=stats,
        selected_ids=ids or None)
    out.update(applied=bool(changed), pairs_changed=changed, skipped=skipped, errors=errors)
    if stats.get('protected'):
        out['embedded_kept_visible'] = stats['protected']
    if stats.get('duplicates'):
        out['overlaps_not_joined'] = stats['duplicates']
        out['note'] = 'Exact overlaps were left unjoined: check for duplicate elements.'
    if message:
        out['success'] = False
        out['error'] = '{} {}'.format(message, 'Resolve any Revit failure dialog and check '
                                               'the model before retrying.')
    elif not changed:
        # No change, so the must_change read-back would fail: report it as nothing done.
        out['success'] = False
        out['message'] = ('Nothing to {}: no intersecting pair in {} needed it. Try another '
                          'scope or rules.'.format(mode, scope.lower()))
    else:
        out['success'] = True
    return out


# ── room_to_floor ───────────────────────────────────────────────────────────

def _floor_type(doc, wanted):
    from Snippets import finish_floors as ff
    types = ff.floor_types(doc)
    if not types:
        return None, None, {'error': 'This model has no floor type. Load or create one, '
                                     'then run again.'}
    ftype, label = ff.pick_floor_type(types, wanted)
    if ftype is None:
        return None, None, {'error': 'No floor type "{}".'.format(wanted),
                            'floor_types': sorted(types)[:MAX_ROWS]}
    return ftype, label, None


def room_to_floor(doc, args):
    from Snippets import finish_floors as ff
    from Snippets._compat import eid_value

    raw = args.get('room_ids')
    if raw is None and args.get('room_id') is not None:
        raw = [args['room_id']]
    ids, bad = int_ids(raw)
    if bad or not ids:
        return {'error': 'Give room_ids (or room_id): the rooms to put floors under.',
                'hint': 'ai_element_filter with category Rooms lists them.'}
    ftype, label, problem = _floor_type(doc, args.get('floor_type'))
    if problem:
        return problem
    use_finish = '{}'.format(args.get('boundary') or 'finish').lower() != 'center'
    offset = float(args.get('offset_mm') or 0.0)
    rooms, skipped = [], []
    for i, room in _elements(doc, ids):
        reason = ff.room_problem(room)
        if reason:
            skipped.append({'room': i, 'reason': reason})
        else:
            rooms.append(room)
    out = {'floor_type': label, 'offset_mm': offset, 'structural': bool(args.get('structural')),
           'boundary': 'finish' if use_finish else 'center', 'rooms': len(rooms)}
    if skipped:
        out['skipped'] = skipped[:MAX_ROWS]
    if not rooms:
        out['error'] = 'None of the rooms can get a floor; see skipped.'
        return out
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False, room_list=[
            '{} {} [{}]'.format(_param(r, 'Number'), _param(r, 'Name'), eid_value(r.Id))
            for r in rooms][:MAX_ROWS])
        return out

    gen = ff.FloorGenerator(doc)
    floors, created, _errors = gen.generate_floors(rooms, ftype, offset,
                                                   bool(args.get('structural')), use_finish)
    out.update(applied=bool(created), success=bool(created), floors_created=created,
               floor_ids=[eid_value(f.Id) for f in floors][:MAX_ROWS])
    if gen.failures:
        out['failed'] = [{'room': eid_value(r.Id), 'reason': why}
                         for r, why in gen.failures][:MAX_ROWS]
    if not created:
        out['error'] = 'No floor was created; see failed.'
    return out


def _param(el, name):
    try:
        p = el.LookupParameter(name)
        return (p.AsString() if p else '') or ''
    except Exception:
        return ''


# ── rehost_to_level ─────────────────────────────────────────────────────────

def rehost_to_level(doc, args):
    from Autodesk.Revit.DB import FilteredElementCollector, Level
    from Snippets import wall_tools as wt
    from Snippets._compat import eid_value, short_error

    ids, bad = int_ids(args.get('element_ids'))
    if bad or not ids:
        return {'error': 'Give element_ids: walls, floors, columns or beams to re-host.'}
    end = '{}'.format(args.get('end') or 'base').lower()
    if end not in ('base', 'top'):
        return {'error': 'end must be base or top.'}
    adjust_base = end == 'base'
    levels = sorted(FilteredElementCollector(doc).OfClass(Level), key=lambda l: l.Elevation)
    level = pick_name(levels, args.get('level') or '', _name) if args.get('level') else None
    if level is None:
        return {'error': 'No level "{}" in this model.'.format(args.get('level') or ''),
                'levels': [_name(l) for l in levels][:MAX_ROWS]}

    planned, skipped = [], []
    for i, el in _elements(doc, ids):
        plan = wt.plan_rehost(doc, el, level, adjust_base) if el is not None else \
            {'reason': 'not in the model'}
        if 'reason' in plan:
            skipped.append({'element': i, 'reason': plan['reason']})
            continue
        planned.append((i, el, plan))
    rows = [{'id': i, 'from_level': _name(p['old_level']),
             'offset_mm': round(p['old_offset'] * 304.8, 1),
             'new_offset_mm': round(p['new_offset'] * 304.8, 1)} for i, _e, p in planned]
    out = {'level': _name(level), 'end': end, 'elements': rows[:MAX_ROWS]}
    if skipped:
        out['skipped'] = skipped[:MAX_ROWS]
    if not planned:
        out['error'] = 'No element can be re-hosted; see skipped.'
        return out
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False)
        return out

    done, failed = [], []
    try:
        with _group(doc, 'T3Lab: Rehost to Level') as group:
            group.Start()
            with _transaction(doc, 'T3Lab: Auto Adjust Offset') as t:
                t.Start()
                wt.silence_warnings(t)
                for i, el, plan in planned:
                    (done if wt.adjust_element(doc, el, level, adjust_base) else failed).append(i)
                t.Commit()
            if done:
                group.Assimilate()
            else:
                group.RollBack()
    except Exception as exc:
        return {'error': 'Re-host stopped: {}. Nothing was changed.'.format(short_error(exc))}
    # Read-back: the level now holds the element at the same elevation.
    wrong = []
    for i, el, plan in planned:
        if i not in done:
            continue
        after = wt.plan_rehost(doc, el, level, adjust_base)
        if ('reason' in after or eid_value(after['p_lvl'].AsElementId()) != eid_value(level.Id)
                or abs(after['old_offset'] - plan['new_offset']) > 1e-6):
            wrong.append(i)
    out.update(applied=bool(done), success=bool(done), adjusted=len(done),
               element_ids=done[:MAX_ROWS])
    if failed:
        out['failed'] = failed[:MAX_ROWS]
    if wrong:
        out['not_in_model_after'] = wrong[:MAX_ROWS]
    if not done:
        out['error'] = 'No element was re-hosted; Revit refused every change.'
    return out


# ── cut_wall_openings ───────────────────────────────────────────────────────

CUT_DEFAULT = ['ducts', 'pipes', 'cable_trays', 'structural_framing']   # the window's ticks


def cut_wall_openings(doc, args):
    from Autodesk.Revit.DB import (BuiltInCategory, FilteredElementCollector,
                                   RevitLinkInstance)
    from Snippets import wall_tools as wt
    from Snippets._compat import eid_value, make_eid

    keys = [k for k, _l, _b in wt.CUT_CATEGORIES]
    wanted = [('{}'.format(c)).lower() for c in (args.get('categories') or CUT_DEFAULT)]
    unknown = [c for c in wanted if c not in keys]
    if unknown:
        return {'error': 'Unknown categories: {}. Use: {}.'.format(', '.join(unknown),
                                                                   ', '.join(keys))}
    method = '{}'.format(args.get('method') or 'opening_family').lower()
    if method not in ('opening_family', 'wall_opening'):
        return {'error': 'method must be opening_family or wall_opening.'}
    offset_ft = float(args.get('clearance_mm', 25) or 0.0) / 304.8

    walls, skipped = [], []
    if args.get('element_ids') is not None:
        ids, bad = int_ids(args.get('element_ids'))
        if bad or not ids:
            return {'error': 'element_ids must list wall ids.'}
        for i, el in _elements(doc, ids):
            if el is not None and wt.is_cut_wall(el):
                walls.append(el)
            else:
                skipped.append({'element': i, 'reason': 'not a basic wall with a location line'})
        where = '{} wall(s)'.format(len(walls))
    else:
        view = doc.GetElement(make_eid(int(args['view_id']))) if args.get('view_id') else \
            doc.ActiveView
        if view is None:
            return {'error': 'No view {}.'.format(args.get('view_id'))}
        coll = FilteredElementCollector(doc, view.Id).OfCategory(
            BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
        walls = [w for w in coll.ToElements() if wt.is_cut_wall(w)]
        where = 'view "{}"'.format(_name(view))
    if not walls:
        return {'error': 'No wall to cut in {} (curtain and stacked walls are skipped). '
                         'Pass element_ids or another view_id.'.format(where),
                'skipped': skipped[:MAX_ROWS]}

    sources = [(doc, None)]
    if args.get('include_links'):
        for inst in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            ldoc = inst.GetLinkDocument()
            if ldoc is not None:                 # unloaded links have nothing to cut with
                sources.append((ldoc, inst.GetTotalTransform()))
    symbol, labels = None, []
    if method == 'opening_family':
        symbols = sorted(wt.opening_symbols(doc), key=wt.symbol_label)
        labels = [wt.symbol_label(s) for s in symbols]
        if not symbols:
            return {'error': 'No opening family is loaded (wall-hosted Generic Model, Window '
                             'or Door, or a family named opening / void / sleeve). Load one, '
                             'or use method=wall_opening.'}
        if args.get('opening_family'):
            symbol = pick_name(symbols, args['opening_family'], wt.symbol_label)
            if symbol is None:
                return {'error': 'No opening family "{}".'.format(args['opening_family']),
                        'opening_families': labels[:MAX_ROWS]}
        else:
            symbol = symbols[0]
    cats = wt.categories(wanted)
    out = {'walls': len(walls), 'scanned': [getattr(d, 'Title', '') for d, _t in sources],
           'method': method, 'clearance_mm': round(offset_ft * 304.8, 1)}
    if symbol is not None:
        out['opening_family'] = wt.symbol_label(symbol)
    if skipped:
        out['skipped'] = skipped[:MAX_ROWS]
    if args.get('dry_run'):
        hits = wt.find_hits(walls, wt.collect_cutters(sources, cats), offset_ft)
        by_cat = {}
        for _w, item in hits:
            by_cat[item['cat']] = by_cat.get(item['cat'], 0) + 1
        out.update(dry_run=True, applied=False, openings_planned=len(hits), by_category=by_cat,
                   walls_hit=len(set(eid_value(w.Id) for w, _i in hits)))
        if labels:
            out['opening_families'] = labels[:MAX_ROWS]
        return out

    method_name = wt.METHOD_FAMILY if method == 'opening_family' else wt.METHOD_OPENING
    created_ids = []
    with _group(doc, 'T3Lab: Wall Cut Profile') as group:
        group.Start()
        created, errors = wt.execute_openings(doc, walls, sources, cats, method_name, offset_ft,
                                              symbol, created_ids)
        if created:
            group.Assimilate()
        else:
            group.RollBack()
    out.update(applied=bool(created), success=bool(created), created_count=created,
               opening_ids=created_ids[:MAX_ROWS])
    if errors:
        out['failed_count'] = errors
    if not created:
        out['error'] = ('No opening was created: {}.'.format(
            '{} failure(s), the run was undone'.format(errors) if errors
            else 'no cutting element crosses these walls'))
    return out


# ── create_door_thresholds ──────────────────────────────────────────────────

def create_door_thresholds(doc, args):
    from Snippets import finish_floors as ff
    from Snippets import wall_tools as wt
    from Snippets._compat import eid_value, short_error

    ids, bad = int_ids(args.get('element_ids'))
    if bad or not ids:
        return {'error': 'Give element_ids: the doors to put thresholds under.',
                'hint': 'ai_element_filter with category Doors lists them.'}
    ftype, label, problem = _floor_type(doc, args.get('floor_type'))
    if problem:
        return problem
    door_cat = wt.bic_int('OST_Doors')
    doors, skipped = [], []
    for i, el in _elements(doc, ids):
        if el is None or wt.category_int(el) != door_cat or getattr(el, 'Location', None) is None:
            skipped.append({'element': i, 'reason': 'not a placed door'})
        else:
            doors.append(el)
    offset = float(args.get('offset_mm') or 0.0)
    out = {'floor_type': label, 'offset_mm': offset, 'doors': len(doors)}
    if skipped:
        out['skipped'] = skipped[:MAX_ROWS]
    if not doors:
        out['error'] = 'None of element_ids is a placed door; see skipped.'
        return out
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False,
                   door_list=[ff.door_label(d) for d in doors][:MAX_ROWS])
        return out

    try:
        with _group(doc, 'T3Lab: Door Threshold') as group:
            group.Start()
            floors, created, _errors, messages = ff.ThresholdGenerator(doc).generate_thresholds(
                doors, ftype, offset)
            if created:
                group.Assimilate()
            else:
                group.RollBack()
    except Exception as exc:
        return {'error': 'Door Threshold stopped: {}. Nothing was created.'.format(
            short_error(exc))}
    out.update(applied=bool(created), success=bool(created), created_count=created,
               threshold_ids=[eid_value(f.Id) for f in floors][:MAX_ROWS])
    if messages:
        out['failed'] = messages[:MAX_ROWS]
    if not created:
        out['error'] = 'No threshold was created; see failed.'
    return out


# ── specs ───────────────────────────────────────────────────────────────────

_FLOOR_TYPE = {'type': 'string', 'description': 'Floor type name. Default: the first one.'}
_OFFSET = {'type': 'number', 'description': 'Height offset from level, mm. Default 0.'}
_DRY = {'type': 'boolean', 'description': 'Report the plan, change nothing.'}

SPECS = [
    ToolSpec(
        'auto_join',
        ('Join (or unjoin) intersecting elements by category rules, the AutoJoin way: in each '
         'rule the priority category cuts join_with. An element lying inside another always '
         'cuts it so it stays visible. Default: the 7 standard rules (walls, floors, columns, '
         'framing) on the whole project. dry_run counts the pairs per rule.'),
        auto_join,
        properties={
            'mode': {'type': 'string', 'enum': ['join', 'unjoin'], 'default': 'join'},
            'rules': {'type': 'array', 'items': {'type': 'object'}, 'description': (
                '[{"priority": "Walls", "join_with": "Floors"}]. Categories: Walls, Floors, '
                'Columns, Structural Columns, Structural Framing, Structural Foundations, '
                'Roofs, Ceilings, Generic Models.')},
            'scope': {'type': 'string', 'enum': ['project', 'active_view'], 'default': 'project'},
            'element_ids': dict(_IDS, description='Only these priority elements (ids or "@set").'),
            'switch_order': {'type': 'boolean', 'description': 'Rule order decides who cuts. '
                                                               'Default true.'},
            'protect_embedded': {'type': 'boolean', 'description': 'Default true.'},
            'dry_run': _DRY,
        },
        effect='write', placement=('modify', 'modeling'), verify=('must_change',),
        destructive=('mode', ('unjoin',)),          # removes joins across the whole scope
        title_vi='Tự động join hình học theo quy tắc',
        keywords=['auto join', 'join hàng loạt', 'nối tường sàn', 'unjoin', 'thứ tự cắt'],
        ribbon='AutoJoin.pushbutton'),
    ToolSpec(
        'room_to_floor',
        ('Create a floor under each room from its boundary (the Room to Floor window). '
         'dry_run lists the rooms that would get one.'),
        room_to_floor,
        properties={
            'room_ids': dict(_IDS, description='Room ids or "@set".'),
            'room_id': {'type': 'integer', 'description': 'One room (alternative to room_ids).'},
            'floor_type': _FLOOR_TYPE,
            'offset_mm': _OFFSET,
            'structural': {'type': 'boolean'},
            'boundary': {'type': 'string', 'enum': ['finish', 'center'], 'default': 'finish',
                         'description': 'Room boundary at wall finish or wall centre.'},
            'dry_run': _DRY,
        },
        effect='write', placement='modeling', verify=('ids', 'must_change'),
        title_vi='Tạo sàn từ phòng', keywords=['sàn', 'floor', 'phòng'],
        ribbon='RoomToFloor.pushbutton', background_doc=True),
    ToolSpec(
        'rehost_to_level',
        ('Move the base (or top) constraint of walls, floors, columns or beams to another '
         'level, adjusting the offset so nothing moves in 3D (Wall Adjust Base). dry_run '
         'lists old level and offsets.'),
        rehost_to_level,
        properties={
            'element_ids': dict(_IDS, description='Element ids or "@set".'),
            'level': {'type': 'string', 'description': 'Target level name or id.'},
            'end': {'type': 'string', 'enum': ['base', 'top'], 'default': 'base'},
            'dry_run': _DRY,
        },
        required=['element_ids', 'level'], effect='write', placement=('modify', 'modeling'),
        verify=('ids', 'self'),
        title_vi='Đổi level đặt đối tượng, giữ nguyên cao độ',
        keywords=['đổi level', 'base constraint', 'top constraint', 'offset', 'rehost'],
        ribbon='Wall_Adjust Base.pushbutton'),
    ToolSpec(
        'cut_wall_openings',
        ('Cut openings in walls where ducts, pipes, cable trays, beams or columns cross them '
         '(Wall Cut Profile): an opening family sized to the element plus clearance, or a '
         'wall opening. Walls: element_ids, else the walls in view_id / the active view. '
         'dry_run counts the crossings and lists opening families.'),
        cut_wall_openings,
        properties={
            'element_ids': dict(_IDS, description='Wall ids or "@set".'),
            'view_id': {'type': 'integer', 'description': 'Walls in this view. Default: active.'},
            'categories': {'type': 'array', 'items': {'type': 'string', 'enum': [
                'ducts', 'pipes', 'cable_trays', 'structural_framing', 'columns']},
                'description': 'Default: all but columns.'},
            'clearance_mm': {'type': 'number', 'description': 'Default 25.'},
            'method': {'type': 'string', 'enum': ['opening_family', 'wall_opening'],
                       'default': 'opening_family'},
            'opening_family': {'type': 'string', 'description': (
                '"Family : Type" or part of it. Default: the first one.')},
            'include_links': {'type': 'boolean', 'description': 'Also cut for loaded links.'},
            'dry_run': _DRY,
        },
        effect='write', placement='modeling', verify=('ids', 'must_change'),
        title_vi='Tạo lỗ mở tường chỗ ống, ống gió, máng cáp, dầm xuyên qua',
        keywords=['lỗ mở', 'opening', 'khoét tường', 'xuyên tường', 'sleeve'],
        ribbon='Wall Cut Profile.pushbutton'),
    ToolSpec(
        'create_door_thresholds',
        ('Create a threshold floor under each door, as wide as the door and as deep as its '
         'wall (Door Threshold). dry_run lists the doors.'),
        create_door_thresholds,
        properties={
            'element_ids': dict(_IDS, description='Door ids or "@set".'),
            'floor_type': _FLOOR_TYPE,
            'offset_mm': _OFFSET,
            'dry_run': _DRY,
        },
        required=['element_ids'], effect='write', placement='modeling',
        verify=('ids', 'must_change'),
        title_vi='Tạo ngạch cửa (sàn dưới cửa)',
        keywords=['ngạch cửa', 'threshold', 'sàn cửa'],
        ribbon='DoorThreshold.pushbutton'),
]
