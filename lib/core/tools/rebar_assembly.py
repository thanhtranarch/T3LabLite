# -*- coding: utf-8 -*-
"""Rebar and assembly tools as specs: inspect_rebar (Rebar Check, assembly list),
rebar_tools (Cast Unit Manager, Clone Drawing), export_bvbs (BVBS Export).
The logic is the ribbon tools' Snippets; this only maps arguments and results.
Test: dev/test_tools_rebar_domain.py."""
from __future__ import unicode_literals

import os
import re

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS, norm as _norm

SERIES_ERROR = ('series is not valid: start and step must be whole numbers (step 1 or more) '
                'and digits 0 to 8, e.g. {"prefix": "CU-", "start": 1, "step": 1, "digits": 3}.')
CLONE_POLICIES = ['skip', 'add_missing', 'new_set', 'replace_annotations']


# ── helpers (pure) ──────────────────────────────────────────────────────────

def ints(values):
    """Element ids as ints; anything not a number is dropped."""
    out = []
    for v in values if isinstance(values, (list, tuple)) else ([values] if values else []):
        try:
            out.append(int(v))
        except (TypeError, ValueError):
            continue
    return out


def natural(text):
    """'C-2' < 'C-10'."""
    parts = re.split(r'(\d+)', text or '')
    return [int(p) if i % 2 else p.lower() for i, p in enumerate(parts)]


def find_assemblies(records, wanted):
    """(AssemblyRecords in asked order, names not found). Each item is an id or a
    mark (case- and space-insensitive); a mark matches every instance of its type."""
    by_id = dict((r.id, r) for r in records)
    by_mark = {}
    for r in records:
        by_mark.setdefault(_norm(r.mark), []).append(r)
    found, missing = [], []
    for w in wanted if isinstance(wanted, (list, tuple)) else [wanted]:
        hits = []
        try:
            if int(w) in by_id:
                hits = [by_id[int(w)]]
        except (TypeError, ValueError):
            pass
        hits = hits or by_mark.get(_norm(w), [])
        if not hits:
            missing.append('{}'.format(w))
        found += [h for h in hits if h not in found]
    return found, missing


def mark_of(record):
    return record.mark or 'assembly {}'.format(record.id)


def parse_series(value):
    """((prefix, start, step, digits), None) or (None, error); defaults CU-, 1, 1, 3
    as in Cast Unit Manager."""
    value = value if isinstance(value, dict) else {}
    try:
        start, step, digits = (int(value.get('start', 1)), int(value.get('step', 1)),
                               int(value.get('digits', 3)))
    except (TypeError, ValueError):
        return None, SERIES_ERROR
    if step < 1 or digits < 0 or digits > 8:
        return None, SERIES_ERROR
    prefix = value.get('prefix', 'CU-')
    return ('{}'.format('' if prefix is None else prefix), start, step, digits), None


def result_rows(results):
    return [{'name': r.name, 'status': r.status, 'detail': r.detail}
            for r in results][:MAX_ROWS]


def type_split_note(split, merged, changed, marks):
    """Rule A4: never silent when Revit split or merged assembly types."""
    if not (split or merged):
        return None
    return ('Revit split {} assembly type(s) and merged {}; marks changed on: {}. Check the '
            'Assembly schedule.'.format(split, merged, ', '.join(marks[:10]) or 'none listed'))


def _counts(results):
    out = {'ok': 0, 'skipped': 0, 'failed': 0}
    for r in results:
        out[r.status] = out.get(r.status, 0) + 1
    return out


def _assembly_model(doc):
    from Snippets import _assembly as asm, _rebar
    return asm.collect_assemblies(doc), _rebar.build_rebar_index(doc)


def _group(doc, name):
    from Autodesk.Revit.DB import TransactionGroup
    return TransactionGroup(doc, 'T3Lab: ' + name)


def _run_grouped(doc, name, work):
    """work() inside ONE TransactionGroup (one undo); rolled back if it raises."""
    from Snippets._compat import disposing
    with disposing(_group(doc, name)) as group:
        group.Start()
        try:
            out = work()
            group.Assimilate()
        except Exception:
            if not group.HasEnded():
                group.RollBack()
            raise
    return out


# ── inspect_rebar (read) ────────────────────────────────────────────────────

def _issue_row(issue):
    row = {'check': issue.check, 'id': issue.element_id, 'detail': issue.detail,
           'fix': issue.fix}
    if issue.assembly_mark:
        row['assembly'] = issue.assembly_mark
    if len(issue.element_ids) > 1:
        row['ids'] = issue.element_ids[:20]
    return row


def check_rebar(doc, args):
    from Snippets import _rebar_check as rc

    check = args.get('check')
    if check and check not in rc.CHECK_ORDER:
        return {'error': 'check must be one of: {}.'.format(', '.join(rc.CHECK_ORDER))}
    scope = args.get('scope') or rc.SCOPE_ALL
    if scope not in (rc.SCOPE_ALL, rc.SCOPE_IN_ASSEMBLY, rc.SCOPE_LOOSE):
        return {'error': 'scope must be all, in or loose.'}
    data = rc.scan(doc)
    deep = bool(args.get('deep'))
    legs = boxes = None
    if deep and data.rows:
        ids = rc.legs_needed(data.rows)
        legs = rc.collect_legs(doc, ids)[0] if ids else {}
        if data.hosts is not None:
            boxes = rc.collect_boxes(doc, data.rows, data.hosts)[0]
    issues = rc.run_checks(data.rows, data.hosts, data.assemblies, legs=legs, boxes=boxes)
    shown = rc.filter_issues(issues, check=check, scope=scope)
    by_check = {}
    for issue in issues:
        by_check[issue.check] = by_check.get(issue.check, 0) + 1
    out = {'summary': rc.summarize(data.rows, issues, data.assemblies), 'by_check': by_check,
           'issues_shown': len(shown), 'issues': [_issue_row(i) for i in shown[:MAX_ROWS]]}
    if not deep:
        out['note'] = ('Geometry not read: bbox is not checked and dup compares shape, type and '
                       'quantity only. deep=true reads bar geometry (slow on big models).')
    plan = rc.plan_fix([i for i in issues if i.fixable])
    if plan.sync_assembly_ids or plan.assign_ids:
        out['fix'] = {}
    if plan.sync_assembly_ids:
        out['fix']['sync'] = {
            'assembly_ids': plan.sync_assembly_ids[:MAX_ROWS],
            'rebar': len(set(plan.sync_rebar_ids)),
            'next': 'rebar_tools operation=sync_assembly_rebar with these assembly_ids'}
    if plan.assign_ids:
        out['fix']['assign_partition'] = {
            'rebar': len(set(plan.assign_ids)),
            'next': 'Cast Unit Manager > Partition by rule (not automated here)'}
    if data.stopped:
        out['warning'] = 'The scan stopped early; results are partial.'
    return out


def list_assemblies(doc, args):
    from Snippets import _assembly as asm
    records, index = _assembly_model(doc)
    if args.get('assemblies'):
        records, missing = find_assemblies(records, args['assemblies'])
        if missing:
            return {'error': 'No assembly with id or mark: {}.'.format(', '.join(missing[:20]))}
    rows = []
    for r in sorted(records, key=lambda a: (natural(a.mark), a.id)):
        rows.append({'id': r.id, 'mark': r.mark, 'type_id': r.type_id, 'same_type': r.instances,
                     'category': r.naming_category, 'level': r.level,
                     'members': len(r.members), 'rebar': len(r.rebar_ids),
                     'unsynced_rebar': len(asm.unsynced_rebar_ids(r, index.by_host)),
                     'views': len(r.view_ids), 'sheets': len(r.sheet_ids)})
    return {'assemblies': rows[:MAX_ROWS], 'count': len(rows)}


def inspect_rebar(doc, args):
    op = '{}'.format(args.get('operation') or 'check').lower()
    if op == 'check':
        return check_rebar(doc, args)
    if op == 'assemblies':
        return list_assemblies(doc, args)
    return {'error': 'operation must be check or assemblies.'}


# ── rebar_tools (write) ─────────────────────────────────────────────────────

def create_cast_units(doc, args):
    from Snippets import _assembly as asm, _rebar

    ids = ints(args.get('element_ids'))
    if not ids:
        return {'error': 'element_ids is required: hosts (beams, columns, foundations, floors, '
                         'walls), their rebar or assemblies.'}
    series = None
    if args.get('series') is not None:
        series, error = parse_series(args.get('series'))
        if error:
            return {'error': error}
    hosts = asm.host_records_from_ids(doc, ids)
    if not hosts:
        return {'error': 'None of element_ids is a valid rebar host (or the rebar / assembly of '
                         'one). Pick beams, columns, foundations, floors or walls.'}
    index = _rebar.build_rebar_index(doc)
    plans = asm.plan_batch_create(hosts, index.by_host,
                                  one_per_host=not args.get('one_assembly'),
                                  include_rebar=args.get('include_rebar') is not False)
    creatable = [p for p in plans if p.host_ids]
    prefix, start, step, digits = series or ('', 1, 1, 3)
    names = asm.series_names(len(creatable), prefix, start, step, digits) if series else []
    rows, n = [], 0
    for p in plans:
        row = {'hosts': p.host_ids or [p.naming_host_id], 'label': p.label,
               'rebar': len(p.rebar_ids)}
        if not p.host_ids:
            row['skipped'] = p.skips[0][1] if p.skips else asm.SKIP_NO_HOST
        elif names:
            row['name'] = names[n]
            n += 1
        rows.append(row)
    out = {'assemblies': len(creatable), 'hosts': len(hosts)}
    if args.get('dry_run') or not creatable:
        out.update(dry_run=bool(args.get('dry_run')), applied=False, plans=rows[:MAX_ROWS])
        if not creatable and not args.get('dry_run'):
            out['error'] = ('Every host is skipped (in a group, from a link or already in an '
                            'assembly); nothing was created.')
        return out

    try:
        before = asm.assembly_type_map(doc)
    except Exception:
        before = {}
    created = []
    results = _run_grouped(doc, 'Create assemblies', lambda: asm.batch_create(
        doc, plans, prefix=prefix, start=start, step=step, digits=digits, created=created))
    after = asm.assembly_type_map(doc)
    split, merged, changed = asm.diff_type_split(
        before, dict((i, after[i]) for i in before if i in after))
    out.update(applied=bool(created), success=bool(created), created=len(created),
               created_assembly_ids=created[:MAX_ROWS], counts=_counts(results),
               results=result_rows(results))
    lost = [i for i in created if i not in after]
    if lost:
        out['not_in_model_after'] = ['assembly {}'.format(i) for i in lost]
    note = type_split_note(split, merged, changed, ['{}'.format(i) for i in changed])
    if note:
        out['type_split'] = note
    if not created:
        out['error'] = 'No assembly was created; see results for the reason per host.'
    return out


def sync_assembly_rebar(doc, args):
    from Snippets import _assembly as asm

    records, index = _assembly_model(doc)
    if args.get('assemblies'):
        records, missing = find_assemblies(records, args['assemblies'])
        if missing:
            return {'error': 'No assembly with id or mark: {}.'.format(', '.join(missing[:20]))}
    todo = [(r, asm.unsynced_rebar_ids(r, index.by_host)) for r in records]
    todo = [(r, loose) for r, loose in todo if loose]
    plan = [{'assembly': mark_of(r), 'id': r.id, 'unsynced_rebar': len(loose)}
            for r, loose in todo]
    if args.get('dry_run') or not todo:
        out = {'dry_run': bool(args.get('dry_run')), 'applied': False, 'plan': plan[:MAX_ROWS]}
        if not todo:
            out['message'] = 'No assembly has loose rebar on its hosts; nothing to sync.'
        return out
    loose_ids = [i for _r, loose in todo for i in loose]
    results = _run_grouped(doc, 'Sync rebar', lambda: asm.sync_rebar(
        doc, [r.id for r, _l in todo], index))
    added = dict((i, index.by_id[i].assembly_id) for i in loose_ids
                 if i in index.by_id and index.by_id[i].assembly_id >= 0)
    members = dict((r.id, set(r.members)) for r in asm.collect_assemblies(doc))
    lost = ['rebar {} in assembly {}'.format(i, a) for i, a in sorted(added.items())
            if i not in members.get(a, ())]
    out = {'applied': bool(added), 'success': bool(added), 'rebar_added': len(added),
           'counts': _counts(results), 'results': result_rows(results)}
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    if not added:
        out['error'] = 'No rebar was added; see results for the reason per assembly.'
    return out


def rename_assemblies(doc, args):
    from Snippets import _assembly as asm

    if not args.get('assemblies'):
        return {'error': 'assemblies is required: ids or marks of the assemblies to rename.'}
    series, error = parse_series(args.get('series'))
    if error:
        return {'error': error}
    records = asm.collect_assemblies(doc)
    picked, missing = find_assemblies(records, args['assemblies'])
    if missing:
        return {'error': 'No assembly with id or mark: {}.'.format(', '.join(missing[:20]))}
    type_ids, marks = [], {}
    for r in sorted(picked, key=lambda a: natural(a.mark)):
        if r.type_id not in type_ids:
            type_ids.append(r.type_id)
            marks[r.type_id] = r.mark
    names = asm.series_names(len(type_ids), *series)
    plan = [{'type_id': t, 'from': marks[t], 'to': n} for t, n in zip(type_ids, names)]
    if args.get('dry_run'):
        return {'dry_run': True, 'applied': False, 'types': len(type_ids),
                'plan': plan[:MAX_ROWS],
                'note': 'Identical assemblies share one type and keep one name.'}
    results, (split, merged, changed) = _run_grouped(
        doc, 'Rename assemblies', lambda: asm.rename_series(doc, type_ids, *series))
    after = asm.collect_assemblies(doc)
    now = set(r.mark for r in after)
    renamed = [r.name for r in results if r.status == asm.STATUS_OK]
    out = {'applied': bool(renamed), 'success': bool(renamed), 'renamed': len(renamed),
           'counts': _counts(results), 'results': result_rows(results)}
    lost = [n for n in renamed if n not in now]
    if lost:
        out['not_in_model_after'] = lost[:MAX_ROWS]
    by_id = dict((r.id, mark_of(r)) for r in after)
    note = type_split_note(split, merged, changed, [by_id.get(i, '{}'.format(i)) for i in changed])
    if note:
        out['type_split'] = note
    if not renamed:
        out['error'] = 'No assembly type was renamed; see results.'
    return out


def _clone_settings(dc, args):
    """(CloneSettings, preset name, error)."""
    preset = args.get('preset') or dc.PRESET_TEKLA
    builtins = dc.builtin_presets()
    user, _last, _warn = dc.load_presets(dc.preset_path())
    base = builtins.get(preset) or user.get(preset)
    if base is None:
        return None, preset, 'No Clone Drawing preset "{}". Presets: {}.'.format(
            preset, ', '.join(list(dc.BUILTIN_PRESET_NAMES) + sorted(user)))
    data = base.to_dict()
    overrides = args.get('settings') or {}
    if not isinstance(overrides, dict):
        return None, preset, 'settings must be an object, e.g. {"tags": "skip"}.'
    unknown = [k for k in overrides if k not in data]
    data.update(overrides)
    if args.get('existing'):
        data['existing'] = args['existing']
    settings, invalid = dc.normalize_settings(data)
    bad = unknown + invalid
    if bad:
        choices = dict((row[0], list(row[2])) for row in dc.SETTINGS_ROWS)
        choices['existing'] = list(dc.EXISTING_CHOICES)
        return None, preset, 'Invalid clone settings: {}. Choices: {}.'.format(
            ', '.join(bad), '; '.join('{}={}'.format(k, '|'.join(v))
                                      for k, v in sorted(choices.items())))
    if settings.existing == dc.EXISTING_REPLACE and args.get('existing') != dc.EXISTING_REPLACE:
        return None, preset, ('Preset "{}" replaces annotations (deletes them first). Pass '
                              'existing="replace_annotations" to confirm, or another existing '
                              'policy.'.format(preset))
    if not settings.any_layer:
        return None, preset, 'Every clone setting is skip; set at least one to clone.'
    return settings, preset, None


def _profile(dc, doc, record):
    from Snippets._compat import make_eid
    try:
        return dc.assembly_profile(doc, doc.GetElement(make_eid(record.id)))
    except Exception:
        return {'category': record.naming_category, 'bbox_mm': (0.0, 0.0, 0.0),
                'rebar_count': len(record.rebar_ids), 'member_count': len(record.members)}


def clone_drawing(doc, args):
    from Snippets import _assembly as asm, _drawing_clone as dc
    from Snippets._compat import make_eid

    settings, preset, error = _clone_settings(dc, args)
    if error:
        return {'error': error}
    records = asm.collect_assemblies(doc)
    sources = [r for r in records if r.view_ids or r.sheet_ids]
    found = find_assemblies(sources, [args['source']])[0] if args.get('source') else []
    if len(found) != 1:
        return {'error': 'source must be the id or mark of ONE assembly that has views '
                         '(Modify > Assembly > Create Views first).',
                'sources': ['{} ({})'.format(mark_of(r), r.id)
                            for r in sorted(sources, key=lambda a: natural(a.mark))][:50]}
    src = found[0]
    views_index = asm.collect_assembly_views(doc)
    specs = dc.read_source(doc, doc.GetElement(make_eid(src.id)), views_index)
    if not [s for s in specs if not s.skip_reason]:
        return {'error': 'No view of {} can be cloned: {}.'.format(mark_of(src), '; '.join(
            '{} ({})'.format(s.name or s.kind, s.skip_reason) for s in specs[:5]) or 'no views')}

    others = [r for r in records if r.id != src.id]
    src_profile = _profile(dc, doc, src)
    rows = []
    if args.get('assemblies'):
        others, missing = find_assemblies(others, args['assemblies'])
        if missing:
            return {'error': 'No other assembly with id or mark: {}.'.format(
                ', '.join(missing[:20]))}
    for r in others:
        score, why = dc.similarity(src_profile, _profile(dc, doc, r))
        text, _sev, clonable = dc.target_status(
            r.type_id == src.type_id, r.naming_category == src.naming_category,
            bool(r.view_ids or r.sheet_ids), settings.existing)
        rows.append((r, score, why, text, clonable))
    if not args.get('assemblies'):
        rows = [x for x in rows if x[4] and x[1] >= dc.SIMILAR_MIN_SCORE]
    rows.sort(key=lambda x: (-x[1], natural(x[0].mark)))
    targets = [x[0] for x in rows if x[4]]
    listed = []
    for r, score, why, text, _c in rows:
        listed.append({'assembly': mark_of(r), 'id': r.id, 'similarity': score, 'status': text})
        if why:
            listed[-1]['why'] = '; '.join(why)
    if not targets:
        return {'error': 'No target to clone to{}.'.format(
            '' if args.get('assemblies') else ' (no clonable assembly scores {}+ % similar; '
            'name targets in assemblies)'.format(dc.SIMILAR_MIN_SCORE)),
                'targets': listed[:MAX_ROWS]}
    replace_plan = {}
    if settings.existing == dc.EXISTING_REPLACE:
        for r in targets:
            if r.view_ids or r.sheet_ids:
                replace_plan[r.id] = dc.replaceable_ids(doc, doc.GetElement(make_eid(r.id)),
                                                        views_index)
    deletes = sum(len(ids) for plan in replace_plan.values() for _v, ids in plan)
    message, details, _ok = dc.confirm_text(mark_of(src), specs, len(targets), settings, deletes)
    out = {'source': mark_of(src), 'preset': preset, 'existing': settings.existing,
           'views': dc.summarize_specs([s for s in specs if not s.skip_reason])}
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False, targets=listed[:MAX_ROWS],
                   plan=message + ' ' + details)
        if deletes:
            out['annotations_to_delete'] = deletes
        return out

    outcomes = dc.run_clone(doc, src.id, [r.id for r in targets], settings,
                            replace_plan=replace_plan)
    totals = dc.summarize_outcomes(outcomes)
    after = asm.collect_assembly_views(doc)
    lost, result = [], []
    for o in outcomes:
        row = {'assembly': o.mark, 'id': o.assembly_id, 'status': o.status,
               'detail': o.detail if o.status != 'ok' else dc.outcome_summary(o)}
        if o.sheet_number:
            row['sheet'] = o.sheet_number
        fails = [t for lvl, t in o.log if lvl == dc.LEVEL_FAILED]
        if fails:
            row['failed'] = fails[:5]
        if o.unmatched:
            row['unmatched'] = ['{}: {}'.format(i, why) for i, why in o.unmatched[:5]]
        result.append(row)
        if o.sheet_id > 0 and doc.GetElement(make_eid(o.sheet_id)) is None:
            lost.append('{}: sheet {}'.format(o.mark, o.sheet_number or o.sheet_id))
        have = len(after.get(o.assembly_id, ([], []))[0])
        if o.status == 'ok' and o.views_created and have < o.views_created:
            lost.append('{}: {} of {} views'.format(o.mark, have, o.views_created))
    out.update(applied=totals['ok'] > 0, success=totals['ok'] > 0, totals=totals,
               targets=result[:MAX_ROWS], summary=dc.tally_text(outcomes))
    if lost:
        out['not_in_model_after'] = lost
    if not totals['ok']:
        out['error'] = 'Nothing was cloned; see targets for the reason per assembly.'
    return out


OPERATIONS = {'create_cast_units': create_cast_units, 'sync_assembly_rebar': sync_assembly_rebar,
              'rename_assemblies': rename_assemblies, 'clone_drawing': clone_drawing}


def rebar_tools(doc, args):
    op = '{}'.format(args.get('operation') or '').lower()
    if op not in OPERATIONS:
        return {'error': 'operation must be one of: {}.'.format(', '.join(sorted(OPERATIONS)))}
    out = OPERATIONS[op](doc, args)
    out['operation'] = op
    return out


# ── export_bvbs (file) ──────────────────────────────────────────────────────

def default_folder(doc):
    """Model folder, else Documents, else home (BVBS Export window defaults)."""
    try:
        if doc.PathName and os.path.isdir(os.path.dirname(doc.PathName)):
            return os.path.dirname(doc.PathName)
    except Exception:
        pass
    documents = os.path.join(os.path.expanduser('~'), 'Documents')
    return documents if os.path.isdir(documents) else os.path.expanduser('~')


def _bvbs_scope(rd, index, assemblies, args):
    """(rebar rows, forced group or None, error)."""
    marks, by_asm = rd.assembly_maps(index, assemblies)
    if args.get('element_ids'):
        return rd.selection_rows(index, assemblies, by_asm, ints(args['element_ids'])), None, None
    if args.get('assembly') not in (None, ''):
        found, _m = find_assemblies(assemblies, [args['assembly']])
        if len(found) != 1:
            return None, None, 'assembly must be the id or mark of one assembly.'
        return rd.assembly_rows(index, by_asm, found[0]), (found[0].mark or ''), None
    if args.get('partition'):
        rows = [r for r in index.rows if r.partition == args['partition']]
        if not rows:
            return None, None, 'No rebar with partition "{}". Partitions: {}.'.format(
                args['partition'], ', '.join(sorted(rd.partition_counts(index), key=natural)[:50]))
        return rows, None, None
    return list(index.rows), None, None


def export_bvbs(doc, args):
    from Snippets import _assembly as asm, _bvbs, _bvbs_read as rd, _rebar

    ok, message = _bvbs.self_test()
    if not ok:
        return {'error': message}
    index = _rebar.build_rebar_index(doc)
    try:
        assemblies = asm.collect_assemblies(doc)
    except Exception:
        assemblies = []
    rows, group, error = _bvbs_scope(rd, index, assemblies, args)
    if error:
        return {'error': error}
    sets, notes = rd.bar_sets(rows, True)          # shape-driven only, the window's default
    weights = {}
    items, _stopped = rd.read_sets(doc, sets, group, rd.assembly_maps(index, assemblies)[0],
                                   _bvbs.load_weight_table(), weights)
    flagged = bool(args.get('include_flagged'))
    take = [i for i in items if i.exportable and (i.selected or flagged)]
    held = [i for i in items if i.exportable and not i.selected and not flagged]
    try:
        project = doc.ProjectInformation.Number or ''
    except Exception:
        project = ''
    plan = rd.apply_fields([i.record for i in take], project, args.get('plan') or rd.DEFAULT_PLAN,
                           rd.DEFAULT_REVISION, args.get('grade') or rd.DEFAULT_GRADE)
    folder = args.get('output_folder') or default_folder(doc)
    files = _bvbs.plan_files([i.record for i in take], bool(args.get('per_assembly')), plan)
    paths = [(os.path.join(folder, name), recs) for name, recs in files]
    out = {'folder': folder, 'bars': sum(i.quantity for i in take), 'marks': len(take),
           'weight': _bvbs.weight_caption(weights)}
    skipped = rd.skipped_counts(items)
    if skipped:
        out['skipped'] = skipped
    if held:
        out['held_back'] = [{'mark': i.mark, 'status': i.status, 'detail': i.detail}
                            for i in held][:MAX_ROWS]
    if notes:
        out['notes'] = notes
    existing = [p for p, _r in paths if os.path.exists(p)]
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False,
                   files=['{} ({} marks)'.format(os.path.basename(p), len(r)) for p, r in paths],
                   rows=[{'mark': i.mark, 'group': i.group, 'dia': i.diameter_mm,
                          'qty': i.quantity, 'legs': i.legs, 'status': i.status}
                         for i in items][:MAX_ROWS])
        if existing:
            out['would_overwrite'] = [os.path.basename(p) for p in existing]
        return out
    if not paths:
        out['error'] = ('Nothing to export in this scope: no shape-driven rebar set with a mark '
                        'reads cleanly{}.'.format(' (see held_back / skipped)' if items else ''))
        return out
    if existing and not args.get('overwrite'):
        out['error'] = ('{} file(s) already exist in {}: {}. Pass overwrite=true or another '
                        'output_folder.'.format(len(existing), folder, ', '.join(
                            os.path.basename(p) for p in existing[:8])))
        return out
    if not os.path.isdir(folder):
        os.makedirs(folder)
    written, failed, bars = [], [], 0
    for path, recs in paths:
        try:
            problems = _bvbs.write_verified(path, recs)[2]
        except Exception as exc:
            problems = ['unexpected error: {}'.format(exc)]
        if problems:
            failed.append({'file': os.path.basename(path), 'problems': problems[:4]})
        else:
            written.append(path)
            bars += sum(r.quantity for r in recs)
    out.update(applied=bool(written), success=bool(written) and not failed,
               exported=len(written), files=written[:MAX_ROWS], bars=bars)
    if failed:
        out['failed'] = failed
        out['error'] = ('{} file(s) failed read-back verification and were not written; any '
                        'earlier file of that name is untouched. Do not send them to a '
                        'machine.'.format(len(failed)))
    return out


SPECS = [
    ToolSpec(
        'inspect_rebar',
        ('Rebar Check (read-only): rebar with no valid host, missing from its host\'s assembly, '
         'assemblies without drawings, duplicate numbers, bars outside their host, no partition, '
         'unknown shape; returns fix hints. operation=assemblies lists cast units with members, '
         'rebar, unsynced rebar, views and sheets.'),
        inspect_rebar,
        properties={
            'operation': {'type': 'string', 'enum': ['check', 'assemblies']},
            'check': {'type': 'string', 'enum': ['host', 'member', 'drawing', 'dup', 'bbox',
                                                 'partition', 'shape'],
                      'description': 'Only this check.'},
            'scope': {'type': 'string', 'enum': ['all', 'in', 'loose'],
                      'description': 'Rebar in assemblies, loose, or all.'},
            'deep': {'type': 'boolean', 'description': 'Read bar geometry for dup/bbox (slow).'},
            'assemblies': {'type': 'array', 'items': {'type': ['integer', 'string']},
                           'description': 'assemblies: ids or marks to list.'},
        },
        effect='read', placement=('read', 'qa'),
        main_thread=True,       # geometry reads off Revit's thread crashed it (check_bad_geometry)
        title_vi='Kiểm tra cốt thép và cấu kiện lắp ghép',
        keywords=['kiểm tra thép', 'rebar check', 'assembly', 'cast unit', 'số hiệu thép'],
        ribbon='RebarCheck.pushbutton'),
    ToolSpec(
        'rebar_tools',
        ('Cast Unit Manager + Clone Drawing. create_cast_units: an assembly per host with its '
         'rebar. sync_assembly_rebar: add loose hosted rebar (default: every assembly). '
         'rename_assemblies: types as a series. clone_drawing: source assembly views, sheet, '
         'annotations to similar ones. dry_run = plan.'),
        rebar_tools,
        properties={
            'operation': {'type': 'string', 'enum': ['create_cast_units', 'sync_assembly_rebar',
                                                     'rename_assemblies', 'clone_drawing']},
            'element_ids': {'type': 'array', 'items': {'type': ['integer', 'string']},
                            'description': 'create: hosts, rebar or assemblies.'},
            'assemblies': {'type': 'array', 'items': {'type': ['integer', 'string']},
                           'description': 'Ids or marks (sync, rename, clone targets).'},
            'series': {'type': 'object',
                       'description': '{"prefix":"CU-","start":1,"step":1,"digits":3}'},
            'one_assembly': {'type': 'boolean'},
            'include_rebar': {'type': 'boolean'},
            'source': {'type': ['integer', 'string'], 'description': 'clone: id or mark.'},
            'preset': {'type': 'string'},
            'settings': {'type': 'object', 'description': 'clone, e.g. {"tags":"skip"}'},
            'existing': {'type': 'string', 'enum': CLONE_POLICIES},
            'dry_run': {'type': 'boolean'},
        },
        required=['operation'], effect='write', placement='modify', verify=('self',),
        destructive=('existing', ('replace_annotations',)),    # deletes target annotations
        title_vi='Tạo, đồng bộ, đặt tên cấu kiện lắp ghép; nhân bản bản vẽ',
        keywords=['assembly', 'cast unit', 'cấu kiện', 'đồng bộ thép', 'clone drawing',
                  'nhân bản bản vẽ', 'đặt tên assembly'],
        ribbon=('CastUnit.pushbutton', 'CloneDrawing.pushbutton')),
    ToolSpec(
        'export_bvbs',
        ('BVBS (.abs, BF2D) files for bending machines from shape-driven rebar, each file read '
         'back and checksum-verified. Scope: element_ids, assembly, partition or whole model. '
         'Bars that need a check (legs vs Revit length, one mark with two shapes) are held back '
         'unless include_flagged. dry_run lists marks, legs and file names.'),
        export_bvbs,
        properties={
            'element_ids': {'type': 'array', 'items': {'type': ['integer', 'string']},
                            'description': 'Rebar, hosts or assemblies.'},
            'assembly': {'type': ['integer', 'string'], 'description': 'Id or mark.'},
            'partition': {'type': 'string'},
            'output_folder': {'type': 'string', 'description': 'Default: the model folder.'},
            'per_assembly': {'type': 'boolean', 'description': 'One file per assembly.'},
            'plan': {'type': 'string', 'description': 'Plan / file name. Default MODEL.'},
            'grade': {'type': 'string', 'description': 'Default B500B.'},
            'include_flagged': {'type': 'boolean'},
            'overwrite': {'type': 'boolean'},
            'dry_run': {'type': 'boolean'},
        },
        effect='file', placement='export', verify=('files',),
        destructive=lambda args: bool(args.get('overwrite')),
        title_vi='Xuất file BVBS cho máy uốn thép',
        keywords=['bvbs', 'abs', 'máy uốn thép', 'bending schedule', 'xuất thép'],
        ribbon='BVBSExport.pushbutton'),
]
