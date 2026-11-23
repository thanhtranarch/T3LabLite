# -*- coding: utf-8 -*-
"""Room views as a spec: generate_room_views (SheetGen window, headless;
Snippets/sheetgen.py). Test: dev/test_tools_views_domain.py."""
from __future__ import unicode_literals

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS

LAYOUTS = ('none', 'combined', 'separate')
KINDS = ('floor', 'ceiling', 'elevation')


def pick_name(wanted, names, fold):
    """Exact (case-insensitive) name, else the one whose part after ': ' matches
    (title block "Family: Type" asked as "Type"). None when absent or ambiguous."""
    key = fold('{}'.format(wanted).strip())
    exact = [n for n in names if fold(n) == key]
    if exact:
        return exact[0]
    tail = [n for n in names if fold(n.split(': ', 1)[-1]) == key]
    return tail[0] if len(tail) == 1 else None


def _offset_ft(value):
    """Crop offset in feet from mm; None when not a finite number >= 0."""
    try:
        mm = float(1000.0 if value is None else value)
    except (TypeError, ValueError):
        return None
    if mm != mm or mm < 0 or mm == float('inf'):
        return None
    return mm / 304.8


def _names(views):
    out = []
    for v in views:
        try:
            out.append(v.Name)
        except Exception:
            continue
    return out


def generate_room_views(doc, args):
    from Snippets import sheetgen as sg
    from Snippets._compat import disposing, eid_value, fold, make_eid, short_error
    from core.tools.view_sync import to_ints
    DB = sg.api()

    want = {'floor': args.get('floor') is not False, 'ceiling': bool(args.get('ceiling')),
            'elevation': bool(args.get('elevations'))}
    if not any(want.values()):
        return {'error': 'Nothing to create: set floor, ceiling or elevations to true.'}
    layout = '{}'.format(args.get('layout') or 'none').strip().lower()
    if layout not in LAYOUTS:
        return {'error': 'layout must be none, combined or separate.'}
    offset = _offset_ft(args.get('offset_mm'))
    if offset is None:
        return {'error': 'offset_mm must be a number of millimetres, zero or more.'}

    ids, bad = to_ints(args.get('room_ids') or [])
    if bad or not ids:
        return {'error': 'room_ids must be room element ids or "@set" handles.'}
    in_model = set(eid_value(r.Id) for r in DB.FilteredElementCollector(doc)
                   .OfCategory(DB.BuiltInCategory.OST_Rooms).ToElements())
    rooms, skipped = [], []
    for i in ids:
        room = doc.GetElement(make_eid(i)) if i in in_model else None
        if room is None:
            skipped.append({'room_id': i, 'reason': 'not a room in this model'})
        elif room.Location is None:
            skipped.append({'room_id': i, 'reason': 'not placed'})
        else:
            rooms.append(sg.RoomInfo(room))
    if not rooms:
        return {'error': 'No placed room in room_ids.', 'skipped': skipped[:MAX_ROWS]}

    type_ids = dict(zip(KINDS, sg.find_view_types(doc)))
    missing = [k for k in KINDS if want[k] and type_ids[k] is None]
    if missing:
        return {'error': 'The model has no {} view type. Add one (Manage > Transfer Project '
                         'Standards > View Types) or leave it out.'.format(', '.join(missing))}
    maps, chosen = sg.template_maps(doc), {}
    asked = args.get('templates') or {}
    if not isinstance(asked, dict) or [k for k in asked if k not in KINDS]:
        return {'error': 'templates takes {"floor", "ceiling", "elevation"}: template names.'}
    for kind in KINDS:
        if not asked.get(kind):
            continue
        name = pick_name(asked[kind], list(maps[kind]), fold)
        if name is None:
            return {'error': 'No {} view template "{}".'.format(kind, asked[kind]),
                    'templates': sorted(maps[kind])[:30]}
        chosen[kind] = name
    tb_id = tb_name = None
    if layout != 'none':
        tbs = sg.titleblock_map(doc)
        if not tbs:
            return {'error': 'The model has no title block; load one or use layout="none".'}
        tb_name = pick_name(args['titleblock'], list(tbs), fold) if args.get('titleblock') \
            else sorted(tbs)[0]
        if tb_name is None:
            return {'error': 'No title block "{}".'.format(args['titleblock']),
                    'titleblocks': sorted(tbs)[:30]}
        tb_id = tbs[tb_name]

    gen = sg.RoomSheetGen(doc, {
        'floor': want['floor'], 'ceiling': want['ceiling'], 'elevations': want['elevation'],
        'floor_type_id': type_ids['floor'], 'ceiling_type_id': type_ids['ceiling'],
        'elevation_type_id': type_ids['elevation'],
        'floor_template_id': maps['floor'].get(chosen.get('floor')),
        'ceiling_template_id': maps['ceiling'].get(chosen.get('ceiling')),
        'elevation_template_id': maps['elevation'].get(chosen.get('elevation')),
        'offset_ft': offset, 'cropbox_visible': args.get('crop_visible') is not False,
    })
    gen.seed_names()
    out = {'layout': layout, 'rooms_asked': len(ids)}
    if tb_name:
        out['titleblock'] = tb_name
    if chosen:
        out['templates'] = chosen
    if skipped:
        out['skipped'] = skipped[:MAX_ROWS]
    if args.get('dry_run'):
        rows = [gen.plan(r, layout) for r in rooms]
        out.update(dry_run=True, applied=False, rooms=rows[:MAX_ROWS],
                   view_count=sum(len(r['views']) for r in rows),
                   sheet_count=sum(len(r['sheets']) for r in rows))
        return out

    rows, view_ids, sheet_ids, pending = [], [], [], None
    try:
        with disposing(DB.TransactionGroup(doc, 'T3Lab: Room views and sheets')) as group:
            group.Start()
            try:
                for room in rooms:
                    res = gen.create_room(room)
                    if res is None:
                        continue
                    views = [v for v in [res['floor_plan'], res['ceiling_plan']] +
                             res['elevations'] if v is not None]
                    row = {'room': '{} {}'.format(room.Number, room.Name).strip(),
                           'views': _names(views)}
                    view_ids.extend(eid_value(v.Id) for v in views)
                    if tb_id is not None and views:
                        try:
                            sheets = gen.layout(res, tb_id, layout == 'combined')
                        except Exception as exc:
                            gen._fail('Sheet layout for room {}: {}'.format(
                                room.Number, short_error(exc)))
                            sheets = []
                        row['sheets'] = ['{} - {}'.format(s.SheetNumber, s.Name) for s in sheets]
                        sheet_ids.extend(eid_value(s.Id) for s in sheets)
                    rows.append(row)
            except sg.SheetGenPending as exc:
                pending = exc
            if pending is None:
                if view_ids:
                    group.Assimilate()
                else:
                    group.RollBack()
    except Exception as exc:
        return {'error': 'Room views stopped: {}. The whole run was undone.'.format(
            short_error(exc))}
    if pending is not None:
        return {'error': '{}. Answer the Revit dialog, then check the model before running '
                         'again.'.format(pending), 'rooms': rows[:MAX_ROWS]}
    out.update(applied=bool(view_ids), success=bool(view_ids), rooms=rows[:MAX_ROWS],
               view_count=len(view_ids), sheet_count=len(sheet_ids),
               view_ids=view_ids[:MAX_ROWS], sheet_ids=sheet_ids[:MAX_ROWS])
    if gen.problems:
        out['failed'] = gen.problems[:MAX_ROWS]
    if not view_ids:
        out['error'] = 'No view was created; see failed. Nothing was changed.'
    return out


SPECS = [
    ToolSpec(
        'generate_room_views',
        ('Per room: enlarged floor plan, ceiling plan and 4 interior elevations cropped to the '
         'room + offset, optionally on new sheets (SheetGen window, headless). Names '
         '"ENLARGED PLAN - <name> - (#<number>)", sheets EPL-<number>; clashes get a suffix. '
         'dry_run lists view names and sheet numbers.'),
        generate_room_views,
        properties={
            'room_ids': {'type': 'array', 'items': {'type': ['integer', 'string']},
                         'description': 'Rooms; ids or "@set".'},
            'floor': {'type': 'boolean', 'description': 'Default true.'},
            'ceiling': {'type': 'boolean'},
            'elevations': {'type': 'boolean'},
            'templates': {'type': 'object', 'description': (
                'View template names: {"floor", "ceiling", "elevation"}.')},
            'titleblock': {'type': 'string', 'description': '"Family: Type". Default: first.'},
            'offset_mm': {'type': 'number', 'description': 'Crop margin. Default 1000.'},
            'layout': {'type': 'string', 'enum': list(LAYOUTS),
                       'description': 'Sheets: one per room (combined), plans + elevations '
                                      '(separate) or none (default).'},
            'crop_visible': {'type': 'boolean', 'description': 'Default true.'},
            'dry_run': {'type': 'boolean', 'description': 'List names, create nothing.'},
        },
        required=['room_ids'], effect='write', placement='modify',
        verify=('ids', 'must_change'),
        title_vi='Tạo mặt bằng, trần, mặt đứng nội thất và sheet theo phòng',
        keywords=['sheetgen', 'mặt bằng phóng to', 'enlarged plan', 'mặt đứng nội thất',
                  'interior elevation', 'sheet theo phòng', 'room plan'],
        ribbon='SheetGen.pushbutton'),
]
