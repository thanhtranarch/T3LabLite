# -*- coding: utf-8 -*-
"""Model standards tools as specs: manage_groups (ManaGroup), manage_worksets (ManaWorkset),
smart_purge (Model Auditor Smart Purge), all headless. Every write reads the model back
and lists what did not land in `not_in_model_after`.
Test: dev/test_tools_standards_domain.py."""
from __future__ import unicode_literals

import sys

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS, names as _names, norm as _norm


def _stopped(what, exc):
    from Snippets._compat import short_error
    return {'error': '{} stopped: {}. Nothing was changed.'.format(what, short_error(exc))}


def _rollback(group):
    try:
        if group.HasStarted() and not group.HasEnded():
            group.RollBack()
    except Exception:
        pass        # a Pending transaction belongs to Revit's failure processing


# ── manage_groups ───────────────────────────────────────────────────────────

GROUP_OPS = ('list', 'rename', 'set_workset', 'purge_unused', 'ungroup')


def pick_named(pool, names, name_of=lambda r: r.name):
    """(items named in `names`, all when empty; names not found). Case-insensitive."""
    if not names:
        return list(pool), []
    by_name = {}
    for r in pool:
        by_name.setdefault(_norm(name_of(r)), []).append(r)
    picked, missing = [], []
    for n in names:
        hits = by_name.get(_norm(n))
        if hits:
            picked += [r for r in hits if r not in picked]
        else:
            missing.append(n)
    return picked, missing


def select_groups(records, names, kind=None):
    """pick_named over the group types of `kind` (every kind when None)."""
    return pick_named([r for r in records if not kind or r.kind == kind], names)


def check_renames(writes, records):
    """Split planned (record, new name) writes into (ok, [(record, reason)]): a name another
    group type keeps, or two records asking for one name, would fail inside Revit."""
    moving = set(id(r) for r, _w in writes)
    taken = set(_norm(r.name) for r in records if id(r) not in moving)
    ok, bad, seen = [], [], set()
    for r, wanted in writes:
        key = _norm(wanted)
        if key in taken:
            bad.append((r, '"{}" is used by another group type'.format(wanted)))
        elif key in seen:
            bad.append((r, '"{}" is asked for twice'.format(wanted)))
        else:
            seen.add(key)
            ok.append((r, wanted))
    return ok, bad


def _group_row(r):
    row = {'name': r.name, 'kind': r.kind, 'instances': r.instance_count,
           'members': r.member_count}
    if r.workset_names:
        row['worksets'] = r.workset_names[:5]
    issues = r.audit_issues()
    if issues:
        row['issues'] = issues
    return row


def manage_groups(doc, args):
    from Snippets import _group_ops as gops

    op = '{}'.format(args.get('operation') or 'list').lower()
    if op not in GROUP_OPS:
        return {'error': 'operation must be one of: {}.'.format(', '.join(GROUP_OPS))}
    kinds = dict((k.lower(), k) for k in gops.KINDS)
    kind = None
    if args.get('kind'):
        kind = kinds.get('{}'.format(args['kind']).lower())
        if kind is None:
            return {'error': 'kind must be model, detail or attached.'}
    new_names = args.get('new_names') or {}
    if not isinstance(new_names, dict):
        return {'error': 'new_names is an object {"old name": "new name"}.'}
    asked = _names(args.get('group_names')) or list(new_names)
    records = gops.collect_group_types(doc)
    picked, missing = select_groups(records, asked, kind)
    if missing:
        return {'error': 'No group type named: {}.'.format(', '.join(missing)),
                'hint': 'manage_groups(operation="list") lists them.'}
    if op == 'list':
        return {'count': len(picked), 'unused': sum(1 for r in picked if not r.instance_count),
                'groups': [_group_row(r) for r in picked[:MAX_ROWS]]}
    dry = bool(args.get('dry_run'))
    out = {'operation': op}
    try:
        if op == 'rename':
            return _rename_groups(doc, gops, args, records, picked, new_names, dry, out)
        if op == 'set_workset':
            return _groups_to_workset(doc, gops, args, picked, asked or kind, dry, out)
        if op == 'purge_unused':
            return _purge_groups(doc, gops, picked, dry, out)
        return _ungroup(doc, gops, picked, asked, dry, out)
    except Exception as exc:
        return _stopped('manage_groups {}'.format(op), exc)


def _rename_groups(doc, gops, args, records, picked, new_names, dry, out):
    if new_names:
        wanted = dict((_norm(k), '{}'.format(v)) for k, v in new_names.items())
        pairs = [(r, wanted[_norm(r.name)]) for r in picked if _norm(r.name) in wanted]
    elif any(args.get(k) for k in ('find', 'prefix', 'suffix')):
        pairs = [(r, gops.build_new_name(r.name, find=args.get('find') or '',
                                         replace=args.get('replace') or '',
                                         prefix=args.get('prefix') or '',
                                         suffix=args.get('suffix') or '')) for r in picked]
    else:
        return {'error': 'rename needs new_names, or find / replace / prefix / suffix.'}
    rejected, _parked, writes = gops.plan_rename(pairs)
    writes, clashes = check_renames(writes, records)
    skipped = [{'group': r.name, 'reason': m} for r, m in rejected if m != 'Unchanged']
    skipped += [{'group': r.name, 'reason': m} for r, m in clashes]
    out['renames'] = [{'from': r.name, 'to': w} for r, w in writes][:MAX_ROWS]
    if skipped:
        out['skipped'] = skipped[:MAX_ROWS]
    if dry or not writes:
        out.update(dry_run=dry, applied=False)
        if not writes and not dry:
            out['error'] = 'Nothing to rename; see skipped.' if skipped else 'No name changes.'
        return out
    target = dict((id(r), w) for r, w in writes)
    results = gops.rename_group_types(doc, writes)
    after = set(_norm(r.name) for r in gops.collect_group_types(doc))
    done = [r for r, ok, _m in results if ok]
    failed = [{'group': r.name, 'reason': m} for r, ok, m in results if not ok]
    out.update(success=bool(done) and not failed, applied=bool(done), renamed=len(done))
    if failed:
        out['failed'] = failed[:MAX_ROWS]
    lost = [target[id(r)] for r in done if _norm(target[id(r)]) not in after]
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    return out


def _groups_to_workset(doc, gops, args, picked, scoped, dry, out):
    ws_name = '{}'.format(args.get('workset') or '').strip()
    if not ws_name:
        return {'error': 'set_workset needs workset (a user workset name).'}
    if not scoped:
        return {'error': 'Name the group types in group_names, or pick a kind.'}
    if not doc.IsWorkshared:
        return {'error': 'The model is not workshared, so it has no worksets.'}
    worksets = gops.user_worksets(doc)
    match = [(i, n) for i, n in worksets if _norm(n) == _norm(ws_name)]
    if not match:
        return {'error': 'No user workset "{}".'.format(ws_name),
                'worksets': [n for _i, n in worksets][:MAX_ROWS]}
    ws_id, ws_name = match[0]
    targets = [r for r in picked if r.instances]
    out['workset'] = ws_name
    if dry or not targets:
        out.update(dry_run=dry, applied=False, groups=[
            {'group': r.name, 'instances': r.instance_count, 'now_on': r.workset_names[:5]}
            for r in targets][:MAX_ROWS])
        if not targets and not dry:
            out['error'] = 'None of these group types is placed; nothing to move.'
        return out
    results = gops.apply_workset(doc, targets, ws_id,
                                 include_members=bool(args.get('include_members')))
    after = dict((r.type_id, r) for r in gops.collect_group_types(doc))
    rows, lost = [], []
    for r, moved, skipped, failed, message in results:
        rows.append({'group': r.name, 'moved': moved, 'failed': failed, 'detail': message})
        now = after.get(r.type_id)
        if not failed and now is not None and [w for w in now.workset_names if w != ws_name]:
            lost.append('{}: still on {}'.format(r.name, ', '.join(now.workset_names)))
    bad = sum(row['failed'] for row in rows)
    out.update(success=not bad, applied=any(row['moved'] for row in rows),
               moved=sum(row['moved'] for row in rows), groups=rows[:MAX_ROWS])
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    return out


def _purge_groups(doc, gops, picked, dry, out):
    unused = [r for r in picked if not r.instance_count]
    out['unused'] = [r.name for r in unused][:MAX_ROWS]
    if dry or not unused:
        out.update(dry_run=dry, applied=False, count=len(unused))
        if not unused and not dry:
            out['message'] = 'No unused group type; nothing was deleted.'
        return out
    results = gops.purge_group_types(doc, unused)
    after = set(r.type_id for r in gops.collect_group_types(doc))
    done = [r for r, ok, _m in results if ok]
    failed = [{'group': r.name, 'reason': m} for r, ok, m in results if not ok]
    out.update(success=not failed, applied=bool(done), deleted=len(done))
    if failed:
        out['failed'] = failed[:MAX_ROWS]
    lost = ['deleted {} (type still in the model)'.format(r.name) for r in done
            if r.type_id in after]
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    return out


def _ungroup(doc, gops, picked, asked, dry, out):
    if not asked:
        return {'error': 'ungroup needs group_names: it breaks apart every instance of '
                         'those group types.'}
    targets = [r for r in picked if r.instances]
    if dry or not targets:
        out.update(dry_run=dry, applied=False, groups=[
            {'group': r.name, 'instances': r.instance_count} for r in targets][:MAX_ROWS])
        if not targets and not dry:
            out['error'] = 'None of these group types is placed; nothing to ungroup.'
        return out
    results = gops.ungroup_instances(doc, targets)
    after = dict((r.type_id, r) for r in gops.collect_group_types(doc))
    rows, lost = [], []
    for r, done, failed, message in results:
        rows.append({'group': r.name, 'ungrouped': done, 'failed': failed, 'detail': message})
        now = after.get(r.type_id)
        if not failed and now is not None and now.instance_count:
            lost.append('ungrouped {} ({} instance(s) still grouped)'.format(
                r.name, now.instance_count))
    out.update(success=not any(row['failed'] for row in rows),
               applied=any(row['ungrouped'] for row in rows),
               ungrouped=sum(row['ungrouped'] for row in rows), groups=rows[:MAX_ROWS])
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    return out


# ── manage_worksets ─────────────────────────────────────────────────────────

WORKSET_OPS = ('rename', 'delete', 'create_views', 'enable_worksharing')


def _ws_name(w):
    return w.Name


def manage_worksets(doc, args):
    from Services import workset_service as ws

    op = '{}'.format(args.get('operation') or '').lower()
    if op not in WORKSET_OPS:
        return {'error': 'operation must be one of: {}.'.format(', '.join(WORKSET_OPS)),
                'hint': 'list_worksets lists them; create_workset adds one.'}
    dry = bool(args.get('dry_run'))
    out = {'operation': op}
    if op == 'enable_worksharing':
        if doc.IsWorkshared:
            return {'error': 'The model is already workshared.'}
        out['worksets'] = list(ws.WORKSHARING_WORKSETS)
        if dry:
            out.update(dry_run=True, applied=False,
                       note='Cannot be undone: the model stays workshared.')
            return out
        ok = ws.enable_worksharing(doc)
        out.update(success=bool(ok), applied=bool(ok))
        if not ok:
            out['error'] = ('Revit did not enable worksharing; the model is unchanged. '
                            'Check that it is saved and not read-only, then retry.')
        elif not doc.IsWorkshared:
            out['not_in_model_after'] = ['worksharing']
        return out
    if not doc.IsWorkshared:
        return {'error': 'The model is not workshared. Run operation="enable_worksharing" first.'}
    current = ws.get_user_worksets(doc)
    try:
        if op == 'create_views':
            return _workset_views(doc, ws, current, dry, out)
        found, missing = pick_named(current, _names(args.get('worksets')), _ws_name)
        if missing or not _names(args.get('worksets')):
            return {'error': 'Name existing user worksets in worksets{}.'.format(
                        ': not found ' + ', '.join(missing) if missing else ''),
                    'worksets': [w.Name for w in current][:MAX_ROWS]}
        if op == 'rename':
            return _rename_workset(doc, ws, current, found, args, dry, out)
        return _delete_worksets(doc, ws, current, found, args, dry, out)
    except Exception as exc:
        return _stopped('manage_worksets {}'.format(op), exc)


def _rename_workset(doc, ws, current, found, args, dry, out):
    from Services.workset_naming import validate_workset_rename
    if len(found) != 1 or not args.get('new_name'):
        return {'error': 'rename takes exactly one workset in worksets, and new_name.'}
    w = found[0]
    ok, clean, message = validate_workset_rename(
        w.Name, '{}'.format(args['new_name']), [x.Name for x in current],
        lambda n: ws.is_workset_name_unique(doc, n))
    if not ok:
        return {'error': message or 'new_name is the current name; nothing to rename.'}
    owner = ws.workset_blocking_owner(doc, w)
    if owner:
        return {'error': '"{}" is owned by {}. Ask them to relinquish it, then retry.'.format(
            w.Name, owner)}
    out.update(renames=[{'from': w.Name, 'to': clean}])
    if dry:
        out.update(dry_run=True, applied=False)
        return out
    done, error = ws.rename_workset(doc, w.Id, clean)
    if not done:
        return {'error': '{} Workset "{}" was not renamed.'.format(error, w.Name)}
    out.update(success=True, applied=True)
    if clean not in ws.get_workset_names(doc):
        out['not_in_model_after'] = [clean]
    return out


def _delete_worksets(doc, ws, current, found, args, dry, out):
    keep, _miss = pick_named(current, _names(args.get('reassign_to')), _ws_name)
    if len(keep) != 1:
        return {'error': 'delete needs reassign_to: the one existing workset that receives '
                         'the elements of the deleted worksets.',
                'worksets': [w.Name for w in current][:MAX_ROWS]}
    keep = keep[0]
    if keep in found:
        return {'error': 'reassign_to "{}" is also being deleted.'.format(keep.Name)}
    blocked = ['{} (owned by {})'.format(w.Name, ws.workset_blocking_owner(doc, w))
               for w in found + [keep] if ws.workset_blocking_owner(doc, w)]
    if blocked:
        return {'error': 'Owned by another user: {}. Ask them to relinquish, then retry.'.format(
            ', '.join(blocked))}
    out.update(delete=[w.Name for w in found], reassign_to=keep.Name)
    if dry:
        out.update(dry_run=True, applied=False)
        return out
    from Autodesk.Revit.DB import TransactionGroup
    from Snippets._compat import disposing
    deleted, failed = [], []
    with disposing(TransactionGroup(doc, 'T3Lab: Delete Worksets')) as group:
        group.Start()
        try:
            for w in found:
                ok, error = ws.delete_workset(doc, w.Name, keep.Name)
                if ok:
                    deleted.append(w.Name)
                else:
                    failed.append({'workset': w.Name, 'reason': error})
        except Exception:
            _rollback(group)
            raise
        if deleted:
            group.Assimilate()
        else:
            group.RollBack()
    out.update(success=not failed, applied=bool(deleted), deleted=deleted)
    if failed:
        out['failed'] = failed
    left = set(ws.get_workset_names(doc))
    lost = ['deleted {} (still in the model)'.format(n) for n in deleted if n in left]
    if lost:
        out['not_in_model_after'] = lost
    return out


def _workset_views(doc, ws, current, dry, out):
    existing = ws.existing_3d_view_names(doc)
    todo = [w.Name for w in current if w.Name not in existing]
    if dry:
        out.update(dry_run=True, applied=False, create=todo[:MAX_ROWS],
                   skipped=len(current) - len(todo))
        return out
    created, skipped, error = ws.create_workset_views(doc)
    if error:
        return {'error': '{} No view was created.'.format(error)}
    out.update(success=True, applied=bool(created), created=created[:MAX_ROWS],
               skipped=len(skipped or []))
    after = ws.existing_3d_view_names(doc)
    lost = [n for n in created if n not in after]
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    return out


# ── smart_purge ─────────────────────────────────────────────────────────────

class _Quiet(object):
    """Scanners print progress; keep it out of the server console / Assistant output."""

    class _Null(object):
        def write(self, *_a):
            pass

        def flush(self, *_a):
            pass

    def __enter__(self):
        self._saved = sys.stdout, sys.stderr
        sys.stdout = sys.stderr = self._Null()
        return self

    def __exit__(self, *_exc):
        sys.stdout, sys.stderr = self._saved


def pick_categories(categories, asked):
    """(categories, unknown names): ids or names; none = the default-checked ones; 'all'."""
    if not asked:
        return [c for c in categories if c.default_checked], []
    if 'all' in [_norm(a) for a in asked]:
        return list(categories), []
    index = {}
    for c in categories:
        index[_norm(c.id)] = index[_norm(c.name)] = c
    picked, unknown = [], []
    for a in asked:
        c = index.get(_norm(a)) or index.get(_norm(a).replace(' ', '_'))
        if c is None:
            unknown.append(a)
        elif c not in picked:
            picked.append(c)
    return picked, unknown


def _ids(values):
    out = []
    for v in values or []:
        try:
            out.append(int(v))
        except (TypeError, ValueError):
            raise ValueError('exclude_ids takes element ids (numbers), not "{}".'.format(v))
    return set(out)


def smart_purge(doc, args):
    from Services.ModelAuditor.smart_purge import purge_categories_v2 as pc
    from Services.ModelAuditor.smart_purge.purge_scanner import create_scanner

    dry = args.get('dry_run') is not False
    categories = pc.create_purge_categories()
    chosen, unknown = pick_categories(categories, _names(args.get('categories')))
    if unknown:
        return {'error': 'Unknown purge category: {}.'.format(', '.join(unknown)),
                'categories': [c.id for c in categories]}
    try:
        exclude = _ids(args.get('exclude_ids'))
    except ValueError as exc:
        return {'error': '{}'.format(exc)}
    rows, buckets, budget, scan_errors = [], [], [MAX_ROWS], []
    for c in chosen:
        try:
            with _Quiet():
                scanner = create_scanner(c.scanner_class, doc)
                items = list(scanner.scan() or []) if scanner else []
        except Exception as exc:
            scan_errors.append('{}: {}'.format(c.id, exc))
            continue
        ready = [i for i in items if i.get('can_delete', True) and i.get('id') not in exclude]
        row = {'category': c.id, 'found': len(items), 'deletable': len(ready)}
        if dry and ready and budget[0] > 0:
            row['items'] = ['{} [{}]'.format(i.get('name'), i.get('id'))
                            for i in ready[:budget[0]]]
            budget[0] -= len(row['items'])
        rows.append(row)
        if ready:
            buckets.append((c.name, ready))
    total = sum(len(items) for _n, items in buckets)
    out = {'categories': rows, 'deletable': total}
    if scan_errors:
        out['scan_errors'] = scan_errors
    if dry:
        out.update(dry_run=True, applied=False,
                   other_categories=[c.id for c in categories if c not in chosen])
        return out
    if not total:
        out.update(applied=False, message='Nothing to purge; the model is unchanged.')
        return out

    from Services.ModelAuditor.smart_purge.purge_executor import PurgeExecutor
    from Snippets._compat import make_eid

    class _Bucket(object):
        def __init__(self, name, items):
            self.name, self.unused_items = name, items

    try:
        deleted, failed, done, failures = PurgeExecutor(doc).execute_purge(
            [_Bucket(n, items) for n, items in buckets])
    except Exception as exc:
        return {'error': '{}. The whole purge was undone.'.format(exc), 'categories': rows}
    out.update(success=not failed, applied=bool(deleted), deleted=deleted)
    if failures:
        out['failed'] = [{'item': '{} [{}]'.format(f['item'].get('name'), f['item'].get('id')),
                          'reason': '{}'.format(f.get('reason'))} for f in failures][:MAX_ROWS]
    lost = ['deleted {} [{}] (still in the model)'.format(i.get('name'), i.get('id'))
            for i in done if doc.GetElement(make_eid(i.get('id'))) is not None]
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    return out


SPECS = [
    ToolSpec(
        'manage_groups',
        ('Model / detail group types (ManaGroup): list them with use and issues, rename '
         '(new_names, or find / replace / prefix / suffix), move their instances to a '
         'workset, delete unused types, or ungroup every instance. Picks group_names, else '
         'every type of kind. One undo step. dry_run previews.'),
        manage_groups,
        properties={
            'operation': {'type': 'string', 'enum': list(GROUP_OPS)},
            'group_names': {'type': 'array', 'items': {'type': 'string'}},
            'kind': {'type': 'string', 'enum': ['model', 'detail', 'attached']},
            'new_names': {'type': 'object', 'description': '{"old name": "new name"}'},
            'find': {'type': 'string'}, 'replace': {'type': 'string'},
            'prefix': {'type': 'string'}, 'suffix': {'type': 'string'},
            'workset': {'type': 'string', 'description': 'set_workset target.'},
            'include_members': {'type': 'boolean',
                                'description': 'set_workset: move the members too.'},
            'dry_run': {'type': 'boolean'},
        },
        required=['operation'], effect='write', placement='modify', verify=('self',),
        destructive=('operation', ('purge_unused', 'ungroup')),
        title_vi='Quản lý group: đổi tên, workset, xoá group thừa, ungroup',
        keywords=['group', 'nhóm', 'đổi tên group', 'ungroup', 'xoá group'],
        ribbon='ManaGroup.pushbutton'),
    ToolSpec(
        'manage_worksets',
        ('Rename a user workset, delete worksets (their elements move to reassign_to), '
         'create one isolated 3D view per workset, or enable worksharing (cannot be undone). '
         'list_worksets / create_workset / set_element_workset do the rest. dry_run previews.'),
        manage_worksets,
        properties={
            'operation': {'type': 'string', 'enum': list(WORKSET_OPS)},
            'worksets': {'type': 'array', 'items': {'type': 'string'},
                         'description': 'rename: one name; delete: the ones to delete.'},
            'new_name': {'type': 'string'},
            'reassign_to': {'type': 'string', 'description': 'delete: receiving workset.'},
            'dry_run': {'type': 'boolean'},
        },
        required=['operation'], effect='write', placement='modify', verify=('self',),
        destructive=('operation', ('delete', 'enable_worksharing')),
        title_vi='Quản lý workset: đổi tên, xoá, tạo view 3D, bật worksharing',
        keywords=['workset', 'đổi tên workset', 'xoá workset', 'worksharing'],
        ribbon='ManaWorkset.pushbutton'),
    ToolSpec(
        'smart_purge',
        ('Model Auditor Smart Purge: find unused materials, patterns, styles, types, '
         'families, groups, views, sheets, CAD imports and more by category, then delete them '
         'in one undo step. Dry run by default: counts and names per category; pass '
         'dry_run=false to delete. Protected items are never deleted.'),
        smart_purge,
        properties={
            'categories': {'type': 'array', 'items': {'type': 'string'}, 'description': (
                'Ids from a dry run (e.g. materials, unused_families), or "all". '
                'Default: the safe default set.')},
            'exclude_ids': {'type': 'array', 'items': {'type': ['integer', 'string']},
                            'description': 'Element ids to keep.'},
            'dry_run': {'type': 'boolean', 'default': True},
        },
        effect='write', placement=('qa', 'modify'), verify=('self',),
        destructive=lambda args: args.get('dry_run') is False,
        title_vi='Dọn model: xoá vật liệu, family, view, kiểu không dùng (Smart Purge)',
        keywords=['purge', 'dọn model', 'xoá không dùng', 'unused', 'model auditor'],
        ribbon='ModelAuditor.pushbutton'),
]
