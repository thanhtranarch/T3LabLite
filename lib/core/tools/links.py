# -*- coding: utf-8 -*-
"""Link tools as specs: link_models (BatchLink: link new models, reload from, worksets
inside a link; Snippets/_links) and cad_imports (ManaDWG; Snippets/_cad_imports).
List / reload / unload / delete / pin of existing links stay in manage_links (server.py).
Test: dev/test_tools_data_domain.py."""
from __future__ import unicode_literals

import os
import re

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS, names as _names, norm as _norm

LINK_OPERATIONS = ('link', 'reload_from', 'set_worksets')
PLACEMENTS = ('origin', 'shared', 'site', 'center')    # = Snippets._links.PLACEMENTS
_BACKUP = re.compile(r'\.\d{3,4}\.rvt$', re.IGNORECASE)  # Revit backups: A.0001.rvt


def expand_paths(paths):
    """(.rvt files, problems): a folder gives the .rvt files directly in it, backups left out."""
    files, problems = [], []
    for p in paths:
        p = os.path.normpath(p.strip().strip('"'))
        if os.path.isdir(p):
            found = sorted(os.path.join(p, n) for n in os.listdir(p)
                           if n.lower().endswith('.rvt') and not _BACKUP.search(n)
                           and os.path.isfile(os.path.join(p, n)))
            if not found:
                problems.append('{}: no .rvt file in this folder'.format(p))
            files += [f for f in found if f not in files]
        elif not p.lower().endswith('.rvt'):
            problems.append('{}: not a .rvt file'.format(p))
        elif not os.path.isfile(p):
            problems.append('{}: file not found'.format(p))
        elif p not in files:
            files.append(p)
    return files, problems


def match_links(records, wanted, id_of=lambda r: r.type_id):
    """(records in order, not found) by link name ('.rvt' optional) or type id."""
    def key(text):
        k = _norm(text).split(' : ')[0]
        return k[:-4] if k.endswith('.rvt') else k
    by_name, by_id = {}, {}
    for r in records:
        by_name.setdefault(key(r.name), r)
        by_id['{}'.format(id_of(r))] = r
    found, missing = [], []
    for w in wanted:
        r = by_id.get('{}'.format(w).strip()) or by_name.get(key(w))
        if r is None:
            missing.append(w)
        elif r not in found:
            found.append(r)
    return found, missing


def workset_plan(worksets, open_names, close_names):
    """Names to open / close in one link (case-insensitive), the names it lacks, and
    whether anything changes. worksets: [LinkWorkset]."""
    want_open = set(_norm(n) for n in open_names)
    want_close = set(_norm(n) for n in close_names)
    have = dict((_norm(w.name), w) for w in worksets)
    to_open = [w.name for k, w in sorted(have.items()) if k in want_open]
    to_close = [w.name for k, w in sorted(have.items()) if k in want_close]
    unknown = sorted(n for n in list(open_names) + list(close_names) if _norm(n) not in have)
    changes = [n for n in to_open if not have[_norm(n)].is_open] + \
              [n for n in to_close if have[_norm(n)].is_open]
    return to_open, to_close, unknown, changes


def _link(doc, args, L):
    from Autodesk.Revit.DB import TransactionGroup
    from Snippets._compat import disposing, short_error

    files, problems = expand_paths(_names(args.get('paths')))
    if problems or not files:
        return {'error': 'paths: {}.'.format('; '.join(problems[:20]) or
                                             'give .rvt files or folders')}
    placement = '{}'.format(args.get('placement') or 'origin').lower()
    attachment = '{}'.format(args.get('attachment') or 'overlay').lower()
    if placement not in PLACEMENTS or attachment not in ('overlay', 'attachment'):
        return {'error': 'placement must be one of {}; attachment overlay or attachment.'.format(
            ', '.join(PLACEMENTS))}
    prefix, ws_name = args.get('workset_prefix'), args.get('workset')
    out = {'operation': 'link', 'placement': placement, 'attachment': attachment}
    workset_for = None
    if prefix or ws_name:
        if not L.is_workshared(doc):
            out['note'] = 'This model is not workshared: the links get no workset.'
        elif prefix:
            workset_for = lambda label: L.ensure_workset(doc, L.link_workset_name(prefix, label))
        else:
            workset_for = lambda label: L.ensure_workset(doc, ws_name)
    paths_before, names_before = L.existing_link_keys(doc)
    skip = args.get('skip_existing', True) is not False
    rows = []
    for f in files:
        already = (os.path.normpath(f).lower() in paths_before
                   or os.path.basename(f).lower() in names_before)
        rows.append({'file': f, 'action': 'skip (already linked)' if skip and already
                     else 'link'})
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False, models=rows[:MAX_ROWS],
                   workshared=L.is_workshared(doc))
        return out

    type_ids, instance_ids, linked = [], [], []
    try:
        with disposing(TransactionGroup(doc, 'T3Lab: Link Models')) as group:
            group.Start()
            for row in rows:
                if row.pop('action') != 'link':
                    row['status'] = 'skipped (already linked)'
                    continue
                label = os.path.basename(row['file'])
                try:
                    res = L.link_one(doc, L.user_model_path(row['file']), label,
                                     relative=bool(args.get('relative')), placement=placement,
                                     overlay=attachment == 'overlay',
                                     pin=args.get('pin', True) is not False,
                                     workset_for=workset_for)
                except Exception as ex:
                    res = {'ok': False, 'status': short_error(ex)}
                row['status'] = res['status']
                if res['ok']:
                    linked.append(row['file'])
                    type_ids.append(res['type_id'])
                    instance_ids.append(res['instance_id'])
                    if res.get('workset_note'):
                        row['status'] += ', workset not set: ' + res['workset_note']
            if linked:
                group.Assimilate()
            else:
                group.RollBack()
    except Exception as ex:
        return {'error': 'Link Models stopped: {}. Nothing was linked.'.format(short_error(ex)),
                'models': rows[:MAX_ROWS]}
    paths_after, names_after = L.existing_link_keys(doc)
    lost = [f for f in linked if os.path.normpath(f).lower() not in paths_after
            and os.path.basename(f).lower() not in names_after]
    out.update(applied=bool(linked), success=bool(linked) or all(
        r['status'].startswith('skipped') for r in rows), linked=len(linked),
        link_type_ids=type_ids[:MAX_ROWS], link_instance_ids=[i for i in instance_ids if i > 0][
            :MAX_ROWS], models=rows[:MAX_ROWS])
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    if not linked and not out['success']:
        out['error'] = 'No model was linked; see models.'
    return out


def _reload_from(doc, args, L):
    records = L.collect_links(doc)
    found, missing = match_links(records, _names(args.get('links')), lambda r: L.eid_int(r.type_id))
    files, problems = expand_paths(_names(args.get('paths')))
    if missing or len(found) != 1 or problems or len(files) != 1:
        return {'error': 'reload_from needs one link in links and one .rvt file in paths. {}'.format(
            '; '.join(problems + ['no link "{}"'.format(m) for m in missing])).strip(),
            'links': [r.name for r in records][:MAX_ROWS]}
    rec, path = found[0], files[0]
    out = {'operation': 'reload_from', 'link': rec.name, 'from': rec.path, 'to': path,
           'link_type_id': L.eid_int(rec.type_id)}
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False)
        return out
    ok, message = L.reload_link_from(rec.link_type, path)
    out.update(applied=ok, success=ok, status=message)
    if not ok:
        out['error'] = 'Revit did not reload "{}" from {}: {}.'.format(rec.name, path, message)
        return out
    now = [r for r in L.collect_links(doc) if L.eid_int(r.type_id) == out['link_type_id']]
    if not now or os.path.normpath(now[0].path).lower() != os.path.normpath(path).lower():
        out['not_in_model_after'] = ['"{}" points to {}, not {}'.format(
            rec.name, now[0].path if now else 'nothing', path)]
    return out


def _set_worksets(doc, args, L):
    records = L.collect_links(doc)
    wanted = _names(args.get('links'))
    found, missing = match_links(records, wanted, lambda r: L.eid_int(r.type_id))
    if missing or not found:
        return {'error': 'No link named: {}.'.format(', '.join(missing) or '(none given)'),
                'links': [r.name for r in records][:MAX_ROWS]}
    open_names, close_names = _names(args.get('open_worksets')), _names(args.get('close_worksets'))
    both = set(_norm(n) for n in open_names) & set(_norm(n) for n in close_names)
    if both or not (open_names or close_names):
        return {'error': 'Give open_worksets and/or close_worksets (worksets inside the link), '
                         'no name in both.'}
    rows, done, lost = [], [], []
    for rec in found:
        row = {'link': rec.name}
        rows.append(row)
        worksets = L.get_link_worksets(rec.link_doc)
        if rec.link_doc is None or not worksets:
            row['status'] = ('not loaded: reload it first' if rec.link_doc is None
                             else 'the linked model is not workshared')
            continue
        to_open, to_close, unknown, changes = workset_plan(worksets, open_names, close_names)
        row.update(open=to_open, close=to_close, changes=len(changes))
        if unknown:
            row['not_in_link'] = unknown
        if args.get('dry_run'):
            row['worksets'] = ['{} ({})'.format(w.name, 'open' if w.is_open else 'closed')
                               for w in worksets][:MAX_ROWS]
            continue
        if not changes:
            row['status'] = 'already so'
            continue
        ids = L.workset_config_ids(worksets, set(to_open), set(to_close))
        ok, row['status'] = L.apply_link_worksets(rec.link_type, ids[0], ids[1])
        if ok:
            done.append((rec, to_open, to_close))
    out = {'operation': 'set_worksets', 'links': rows[:MAX_ROWS]}
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False)
        return out
    after = dict((L.eid_int(r.type_id), r) for r in L.collect_links(doc))
    for rec, to_open, to_close in done:
        now = after.get(L.eid_int(rec.type_id))
        state = dict((w.name, w.is_open) for w in L.get_link_worksets(
            now.link_doc if now is not None else None))
        lost += ['{}: {} not open'.format(rec.name, n) for n in to_open if not state.get(n)]
        lost += ['{}: {} not closed'.format(rec.name, n) for n in to_close
                 if state.get(n, True)]
    failed = [r for r in rows if r.get('status') not in (None, 'Reloaded', 'already so')]
    out.update(applied=bool(done), success=not failed, reloaded=len(done))
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    if failed and not done:
        out['error'] = 'No link was changed; see links.'
    return out


def link_models(doc, args):
    op = '{}'.format(args.get('operation') or '').lower()
    if op not in LINK_OPERATIONS:
        return {'error': 'operation must be one of: {}.'.format(', '.join(LINK_OPERATIONS))}
    from Snippets import _links as L
    return {'link': _link, 'reload_from': _reload_from,
            'set_worksets': _set_worksets}[op](doc, args, L)


def cad_row(item, eid_value):
    inst = item.Instance
    view = item.OwnerViewId
    view = eid_value(view) if view is not None else -1
    return {'id': eid_value(inst.Id), 'name': item.Name, 'kind': item.DWGType.lower(),
            'view': item.ViewNames, 'view_id': view if view > 0 else None,
            'pinned': bool(getattr(inst, 'Pinned', False))}


def cad_imports(doc, args):
    from Snippets import _cad_imports as C
    from Snippets._compat import eid_value, short_error

    op = '{}'.format(args.get('operation') or 'list').lower()
    if op not in ('list', 'delete'):
        return {'error': 'operation must be list or delete.'}
    items = [i for i in C.collect_dwg_items(doc) if i.Instance is not None]
    rows = [cad_row(i, eid_value) for i in items]
    if op == 'list':
        return {'count': len(rows), 'imports': sum(1 for r in rows if r['kind'] == 'import'),
                'links': sum(1 for r in rows if r['kind'] == 'link'), 'items': rows[:MAX_ROWS]}
    wanted = []
    for v in args.get('element_ids') or []:
        try:
            wanted.append(int(v))
        except (TypeError, ValueError):
            return {'error': 'element_ids: "{}" is not an element id.'.format(v)}
    by_id = dict((r['id'], (i, r)) for i, r in zip(items, rows))
    unknown = [w for w in wanted if w not in by_id]
    if unknown or not wanted:
        return {'error': 'Not a CAD import or link in this model: {}. Call operation=list.'.format(
            ', '.join('{}'.format(u) for u in unknown[:20]) or '(no element_ids)')}
    picked = [by_id[w] for n, w in enumerate(wanted) if w not in wanted[:n]]
    out = {'operation': 'delete', 'items': [r for _i, r in picked][:MAX_ROWS]}
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False, note=(
            'Pinned items are unpinned first. A CAD link type goes too when no instance is left.'))
        return out
    ids = [(i.Instance.Id, r) for i, r in picked]
    try:
        deleted, failures = C.delete_cad_items(doc, [i for i, _r in picked])
    except Exception as ex:
        return {'error': 'Nothing was deleted: {}. Check that nobody else has these items '
                         'borrowed, then try again.'.format(short_error(ex))}
    still = ['{} [{}] still in the model'.format(r['name'], r['id']) for eid, r in ids
             if doc.GetElement(eid) is not None]
    instance_failures = [f for f in failures if not '{}'.format(f[0]).endswith('(link type)')]
    out.update(applied=deleted > 0, success=deleted > 0, deleted=deleted)
    if failures:
        out['failed'] = [{'name': n, 'reason': short_error(e)} for n, e in failures][:MAX_ROWS]
    if len(still) > len(instance_failures):
        out['not_in_model_after'] = still[:MAX_ROWS]
    if not deleted:
        out['error'] = 'Nothing was deleted; see failed.'
    return out


SPECS = [
    ToolSpec(
        'link_models',
        ('Link new Revit models (BatchLink). link: .rvt files or folders, one undo, already '
         'linked ones skipped. reload_from: point one link at another file. set_worksets: '
         'open/close worksets inside links. Reloads clear Revit undo (as Manage Links). '
         'List/reload/unload/remove: manage_links.'),
        link_models,
        properties={
            'operation': {'type': 'string', 'enum': list(LINK_OPERATIONS)},
            'paths': {'type': 'array', 'items': {'type': 'string'},
                      'description': '.rvt files or folders.'},
            'links': {'type': 'array', 'items': {'type': ['integer', 'string']},
                      'description': 'Existing links: names or type ids.'},
            'placement': {'type': 'string', 'enum': list(PLACEMENTS)},
            'attachment': {'type': 'string', 'enum': ['overlay', 'attachment']},
            'relative': {'type': 'boolean'},
            'pin': {'type': 'boolean', 'description': 'Default true.'},
            'skip_existing': {'type': 'boolean', 'description': 'Default true.'},
            'workset': {'type': 'string', 'description': 'Host workset for new links.'},
            'workset_prefix': {'type': 'string', 'description': (
                'One host workset per link: prefix + file name, created if missing.')},
            'open_worksets': {'type': 'array', 'items': {'type': 'string'}},
            'close_worksets': {'type': 'array', 'items': {'type': 'string'}},
            'dry_run': {'type': 'boolean'},
        },
        required=['operation'], effect='write', placement='modify', verify=('ids', 'self'),
        # A reload clears the whole Revit undo history: the person confirms.
        destructive=('operation', ('reload_from', 'set_worksets')),
        title_vi='Link hàng loạt file Revit, mở đóng workset trong link',
        keywords=['link revit', 'liên kết mô hình', 'batch link', 'workset link', 'reload from'],
        ribbon='BatchLink.pushbutton'),
    ToolSpec(
        'cad_imports',
        ('CAD imports and CAD links (ManaDWG). list: every instance with import/link, owner '
         'view and pin. delete: element_ids from list, unpinned first; a link type goes with '
         'its last instance; one undo.'),
        cad_imports,
        properties={
            'operation': {'type': 'string', 'enum': ['list', 'delete']},
            'element_ids': {'type': 'array', 'items': {'type': ['integer', 'string']}},
            'dry_run': {'type': 'boolean'},
        },
        required=['operation'], effect='write', placement='modify', verify=('self',),
        destructive=('operation', ('delete',)),
        title_vi='Liệt kê, xoá CAD import và CAD link (DWG)',
        keywords=['dwg', 'cad', 'import cad', 'xoá dwg', 'link cad'],
        ribbon='ManaDWG.pushbutton'),
]
