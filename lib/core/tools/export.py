# -*- coding: utf-8 -*-
"""Export tools as specs: export_sheets (BatchOut PDF / DWG / DXF / DGN, Services/sheet_export).
Doc: docs/agents/sheet-export.md. Test: dev/test_sheet_export.py."""
from __future__ import unicode_literals

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS


def export_sheets(doc, args):
    from Services import sheet_export as se

    fmt = '{}'.format(args.get('format') or '').strip().lower()
    if fmt not in se.FORMATS:
        return {'error': 'format must be one of: {}.'.format(', '.join(se.FORMATS))}
    profile = None
    if args.get('profile'):
        profile = se.load_profile(args['profile'])
        if profile is None:
            return {'error': 'No BatchOut profile "{}".'.format(args['profile']),
                    'profiles': se.profile_names(), 'folder': se.PROFILES_FOLDER}
    s = se.settings_from(profile, pattern=args.get('naming'), combined=args.get('combined'),
                         export_setup=args.get('export_setup'),
                         report=None if args.get('report') is None else bool(args['report']))
    if args.get('output_folder'):
        s['folder'], s['split_by_format'] = args['output_folder'], False

    sets = se.sheet_sets(doc)
    picks = [k for k in ('sheet_set', 'sheet_numbers', 'sheet_ids', 'prefix') if args.get(k)]
    if not picks and not args.get('all'):
        return {'error': 'Say which sheets: sheet_set, sheet_numbers, sheet_ids, prefix, '
                         'or all=true.', 'sheet_sets': sets}
    member_ids, other_views = None, 0
    if args.get('sheet_set'):
        member_ids, other_views = se.set_member_ids(doc, args['sheet_set'])
        if member_ids is None:
            return {'error': 'No sheet set "{}" in this model.'.format(args['sheet_set']),
                    'sheet_sets': sets}
    picked, missing = se.select(se.sheet_rows(doc), args.get('sheet_numbers'),
                                args.get('sheet_ids'), args.get('prefix'), member_ids)
    if missing:
        return {'error': 'Not in this model: {}.'.format(', '.join(missing[:20]))}
    if not picked:
        return {'error': 'No sheet matches the selection.', 'sheet_sets': sets}

    folder = se.output_folder(s, fmt)
    combined = fmt == 'pdf' and bool(s['combined'])
    out = {'format': fmt, 'output_folder': folder, 'sheets': len(picked),
           'expected': 1 if combined else len(picked)}
    if args.get('sheet_set'):
        out['sheet_set'] = args['sheet_set']
        if other_views:
            out['note'] = ('The sheet set also holds {} view(s) that are not sheets; '
                           'only sheets are exported.'.format(other_views))
    if profile is not None:
        out['profile'] = profile.get('Name')
    if args.get('dry_run'):
        used, names = set(), []
        for sh in picked if not combined else picked[:1]:
            stem = se.file_stem(s['pattern'], se.sheet_values(doc, sh['sheet'], s['pattern']))
            if combined and len(picked) > 1:
                stem = '{}-{}_Combined'.format(picked[0]['number'], picked[-1]['number'])
            names.append(se.reserve(used, folder, stem, '.' + fmt) + '.' + fmt)
        out.update(dry_run=True, applied=False, files=names[:MAX_ROWS],
                   sheet_list=['{} - {}'.format(r['number'], r['name']) for r in picked][:MAX_ROWS],
                   naming=s['pattern'], sheet_sets=sets, profiles=se.profile_names())
        if fmt != 'pdf':
            out['export_setups'] = se.export_setups(doc)
            out['export_setup'] = s['export_setup'] or None
        return out

    res = se.export(doc, fmt, picked, folder, s)
    failed = [{'sheet': r['number'], 'reason': r['reason']} for r in res['rows']
              if r['status'] != 'Exported']
    out.update(applied=True, exported=len(res['files']), files=res['files'][:MAX_ROWS],
               success=len(res['files']) == out['expected'] and not failed)
    if failed:
        out['failed'] = failed[:MAX_ROWS]
    if not res['files']:
        out['error'] = 'No file was exported; see failed.'
    if res['notes']:
        out['notes'] = res['notes']
    if s.get('report'):
        path, problem = se.write_report(folder, res['rows'], s)
        out['report'] = path or problem
    return out


SPECS = [
    ToolSpec(
        'export_sheets',
        ('Export sheets to PDF, DWG, DXF or DGN the BatchOut way. Pick sheets by Revit sheet set, '
         'numbers, ids or number prefix. Optional BatchOut profile, CAD export setup and file '
         'naming ({SheetNumber}, {SheetName}, {Revision} or any sheet parameter). Every file is '
         'checked after export. dry_run lists sheets, file names, sheet sets, setups and '
         'profiles. Many sheets: run it through t3lab_job.'),
        export_sheets,
        properties={
            'format': {'type': 'string', 'enum': ['pdf', 'dwg', 'dxf', 'dgn']},
            'sheet_set': {'type': 'string', 'description': 'Saved Revit sheet set (V/S set).'},
            'sheet_numbers': {'type': 'array', 'items': {'type': 'string'}},
            'sheet_ids': {'type': 'array', 'items': {'type': ['integer', 'string']}},
            'prefix': {'type': 'string', 'description': 'Sheet numbers starting with this.'},
            'all': {'type': 'boolean', 'description': 'Every sheet.'},
            'profile': {'type': 'string', 'description': 'Saved BatchOut profile.'},
            'output_folder': {'type': 'string', 'description': (
                'Default: profile folder or Documents/Revit Exports/<FORMAT>.')},
            'naming': {'type': 'string', 'description': 'Default "{SheetNumber}-{SheetName}".'},
            'combined': {'type': 'boolean', 'description': 'PDF: one file.'},
            'export_setup': {'type': 'string', 'description': 'DWG/DXF/DGN setup in the model.'},
            'report': {'type': 'boolean', 'description': 'Write the BatchOut report file.'},
            'dry_run': {'type': 'boolean', 'description': 'List what would be exported.'},
        },
        required=['format'], effect='file', placement=('export', 'essential'),
        verify=('files',),
        title_vi='Xuất sheet PDF, DWG, DXF, DGN theo sheet set (BatchOut)',
        keywords=['xuất pdf', 'xuất cad', 'xuất dwg', 'sheet set', 'bộ bản vẽ', 'batchout',
                  'in bản vẽ'],
        ribbon='BatchOut.pushbutton'),
]
