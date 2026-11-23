# -*- coding: utf-8 -*-
"""Schedule data tools as specs: schedule_roundtrip (Schedule Manager's Schedule Data
and New Table pages, without the window: Snippets/schedule_data, excel_table_revit).
Test: dev/test_tools_data_domain.py."""
from __future__ import unicode_literals

import os

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS, names as _names, norm as _norm

OPERATIONS = ('export', 'preview', 'apply', 'table', 'update_tables')
TABLE_KINDS = ('drafting', 'schedule', 'legend')     # excel_table_revit.KIND_*


def pick_schedules(schedules, wanted, id_of):
    """(picked [(name, schedule)], not found) by name (case/space-insensitive) or id."""
    by_name, by_id = {}, {}
    for pair in schedules:
        by_name.setdefault(_norm(pair[0]), pair)
        by_id['{}'.format(id_of(pair[1]))] = pair
    picked, missing = [], []
    for w in wanted:
        pair = by_id.get('{}'.format(w).strip()) or by_name.get(_norm(w))
        if pair is None:
            missing.append(w)
        elif pair not in picked:
            picked.append(pair)
    return picked, missing


def default_xlsx_path(picked):
    stem = picked[0][0] if len(picked) == 1 else 'Schedules'
    for ch in '/\\*?[]:"<>|':
        stem = stem.replace(ch, '_')
    return os.path.join(os.path.expanduser('~'), 'Documents', 'Revit Exports', 'XLSX',
                        stem + '.xlsx')


def change_row(ch):
    return {'schedule': ch['schedule'], 'element_id': ch['element_id'], 'field': ch['field'],
            'old': ch['old'], 'new': ch['new']}


def _table(doc, args, path):
    from Snippets import excel_table_revit as etr
    from Snippets import xlsx_table as xt
    from Snippets._compat import eid_value

    kind = '{}'.format(args.get('kind') or 'drafting').lower()
    if kind not in TABLE_KINDS:
        return {'error': 'kind must be one of: {}.'.format(', '.join(TABLE_KINDS))}
    try:
        table = xt.read_table(path, args.get('sheet') or None, args.get('range') or None, False)
    except xt.XlsxError as ex:
        return {'error': '{}'.format(ex), 'hint': 'Check sheet and range (A1:F20 or a named range).'}
    if not table.rows or not table.cols:
        return {'error': 'The range is empty in {}.'.format(os.path.basename(path))}
    name = etr.unique_view_name(doc, args.get('name') or 'Excel Table')
    out = {'operation': 'table', 'xlsx_path': path, 'sheet': table.sheet, 'kind': kind,
           'range': xt.range_text(*table.bounds) if table.bounds else table.range_name,
           'rows': table.rows, 'cols': table.cols, 'view_name': name}
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False)
        return out
    try:
        view = etr.create_table(doc, path, table, kind, name,
                                {'scale': 1.0, 'formatting': True, 'hidden': False}, True)
    except etr.TableError as ex:
        return {'error': 'Could not create the table: {}'.format(ex)}
    out.update(applied=True, success=True, view_id=eid_value(view.Id), view_name=view.Name,
               linked=True)
    if doc.GetElement(view.Id) is None:
        out['not_in_model_after'] = ['view "{}"'.format(name)]
    return out


def _update_tables(doc, args):
    from Autodesk.Revit.DB import TransactionGroup
    from Snippets import excel_table_revit as etr
    from Snippets._compat import disposing, eid_value, short_error

    linked = etr.find_linked_views(doc)
    wanted = set(_norm(n) for n in _names(args.get('name')))
    picked = [(v, i) for v, i in linked if not wanted or _norm(v.Name) in wanted]
    if not picked:
        return {'error': 'No linked Excel table{}.'.format(
            ' named "{}"'.format(args.get('name')) if wanted else ' in this model'),
            'tables': [v.Name for v, _i in linked][:MAX_ROWS]}
    rows = [{'view': v.Name, 'kind': i.get('kind'), 'source': etr.source_label(i)}
            for v, i in picked]
    out = {'operation': 'update_tables', 'tables': rows[:MAX_ROWS]}
    if args.get('dry_run'):
        for (v, i), row in zip(picked, rows):
            row['status'] = etr.check_status(doc, i)[0]
        out.update(dry_run=True, applied=False)
        return out
    done = []
    try:
        with disposing(TransactionGroup(doc, 'T3Lab: Excel Table - Update')) as group:
            group.Start()
            for (v, i), row in zip(picked, rows):
                try:
                    view, warnings = etr.update_table(doc, v, i)
                except etr.TableError as ex:
                    row.update(status='failed', detail='{}'.format(ex))
                    continue
                row['status'] = 'updated'
                if warnings:
                    row['warnings'] = warnings
                done.append((view.Id, row['view']))
            if done:
                group.Assimilate()
            else:
                group.RollBack()
    except Exception as ex:
        return {'error': 'Update stopped: {}. Nothing changed.'.format(short_error(ex)),
                'tables': rows[:MAX_ROWS]}
    out.update(applied=bool(done), success=len(done) == len(rows), updated=len(done))
    lost = ['view "{}"'.format(n) for vid, n in done if doc.GetElement(vid) is None]
    if lost:
        out['not_in_model_after'] = lost
    if not done:
        out['error'] = 'No table was updated; see tables.'
    return out


def schedule_roundtrip(doc, args):
    op = '{}'.format(args.get('operation') or '').lower()
    if op not in OPERATIONS:
        return {'error': 'operation must be one of: {}.'.format(', '.join(OPERATIONS))}
    path = '{}'.format(args.get('xlsx_path') or '').strip()
    if op == 'update_tables':
        return _update_tables(doc, args)
    if op != 'export' and not os.path.isfile(path):
        return {'error': 'xlsx_path "{}" is not a file. Export first (operation=export), '
                         'edit it in Excel, then preview / apply.'.format(path)}
    if op == 'table':
        return _table(doc, args, path)

    from Snippets import schedule_data as sd
    from Snippets import xlsx_table as xt
    from Snippets._compat import eid_value

    every = sd.list_schedules(doc)
    picked, missing = pick_schedules(every, _names(args.get('schedules')),
                                     lambda s: eid_value(s.Id))
    if missing:
        return {'error': 'No schedule named: {}.'.format(', '.join(missing)),
                'schedules': [n for n, _s in every][:MAX_ROWS]}
    dry = bool(args.get('dry_run'))

    if op == 'export':
        if not picked:
            return {'error': 'Say which schedules to export (names or ids).',
                    'schedules': [n for n, _s in every][:MAX_ROWS]}
        path = path or default_xlsx_path(picked)
        if not path.lower().endswith('.xlsx'):
            path += '.xlsx'
        sheets = sd.sheet_names_for([n for n, _s in every])
        out = {'operation': op, 'xlsx_path': path,
               'sheets': dict((n, sheets[n]) for n, _s in picked)}
        if dry:
            out.update(dry_run=True, applied=False)
            return out
        folder = os.path.dirname(path)
        try:
            if folder and not os.path.isdir(folder):
                os.makedirs(folder)
            rows = sd.export_xlsx(doc, picked, path)
        except (IOError, OSError) as ex:
            return {'error': 'Could not write {}: {}. If it is open in Excel, close it and run '
                             'again.'.format(path, ex)}
        out.update(applied=True, success=True, rows=rows, files=[path], exported=1)
        return out

    fallback = picked[0] if len(picked) == 1 else None
    try:
        changes, report = sd.read_changes(doc, path, every, fallback)
    except xt.XlsxError as ex:
        return {'error': '{}'.format(ex)}
    if picked:
        keep = set(n for n, _s in picked)
        changes = [c for c in changes if c['schedule'] in keep]
    out = {'operation': op, 'xlsx_path': path, 'change_count': len(changes),
           'elements': len(set(c['element_id'] for c in changes)),
           'changes': [change_row(c) for c in changes][:MAX_ROWS], 'report': report[:MAX_ROWS]}
    if op == 'preview' or dry:
        out.update(dry_run=dry, applied=False)
        return out
    if not changes:
        out.update(applied=False, message='No value differs from the model; nothing written.')
        return out
    done, errors, failed = sd.apply_changes(doc, changes)
    failed_ids = set(id(c) for c in failed)
    lost = sd.read_back(doc, [c for c in changes if id(c) not in failed_ids]) if done else []
    out.update(applied=done > 0, success=done > 0, written=done)
    if errors:
        out['failed'] = errors[:MAX_ROWS]
    if lost:
        out['not_in_model_after'] = ['{} [{}] {}: wrote "{}", reads "{}"'.format(
            c['schedule'], c['element_id'], c['field'], c['new'], c['now'])
            for c in lost][:MAX_ROWS]
    if not done:
        out['error'] = 'No change was written; see failed.'
    return out


SPECS = [
    ToolSpec(
        'schedule_roundtrip',
        ('Schedule values through Excel (Schedule Manager). export: schedules to one xlsx '
         '(Element ID + visible fields). preview: what an edited xlsx would change. apply: '
         'write those changes, one undo, each cell read back. table: draw an xlsx range as a '
         'linked drafting view, header-only schedule or legend. update_tables: redraw linked '
         'tables (name, or all) from their file.'),
        schedule_roundtrip,
        properties={
            'operation': {'type': 'string', 'enum': list(OPERATIONS)},
            'schedules': {'type': 'array', 'items': {'type': ['integer', 'string']},
                          'description': 'Names or ids. preview/apply: limit to these.'},
            'xlsx_path': {'type': 'string', 'description': (
                'export default: Documents/Revit Exports/XLSX/<schedule>.xlsx.')},
            'sheet': {'type': 'string', 'description': 'table: worksheet (default first).'},
            'range': {'type': 'string', 'description': 'table: A1:F20 or a name (default used).'},
            'kind': {'type': 'string', 'enum': list(TABLE_KINDS)},
            'name': {'type': 'string', 'description': 'table / update_tables: view name.'},
            'dry_run': {'type': 'boolean'},
        },
        required=['operation'], effect='write', placement=('modify', 'export'),
        verify=('self', 'files'),
        title_vi='Xuất, sửa schedule qua Excel; tạo bảng từ Excel',
        keywords=['schedule', 'excel', 'xlsx', 'bảng thống kê', 'xuất excel', 'nhập excel',
                  'bảng excel'],
        ribbon='ManaSched.pushbutton'),
]
