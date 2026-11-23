# -*- coding: utf-8 -*-
"""Family tools as specs: transfer_families (Family Transfer window, headless).
Test: dev/test_tools_standards_domain.py."""
from __future__ import unicode_literals

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS, names as _names, norm as _norm
from core.tools.views import match_documents


def pick_families(families, family_names, type_names):
    """([(FamilyInfo, [type names])], families not found, types not found).
    No type_names = every type of each family."""
    by_name = {}
    for f in families:
        by_name.setdefault(_norm(f.name), []).append(f)
    wanted = set(_norm(t) for t in type_names)
    selections, missing = [], []
    for n in family_names:
        hits = by_name.get(_norm(n))
        if not hits:
            missing.append(n)
            continue
        for f in hits:
            if f in [s[0] for s in selections]:
                continue
            selections.append((f, [t for t in f.type_names if not wanted or _norm(t) in wanted]))
    found = set(_norm(t) for _f, types in selections for t in types)
    return selections, missing, [t for t in type_names if _norm(t) not in found]


def plan_rows(actions):
    """{'copy': [...], 'overwrite': [...], 'skip': [...]} as 'Family (types)' / 'Family: reason'."""
    out = {'copy': [], 'overwrite': [], 'skip': []}
    for a in actions:
        if a.kind == 'skip':
            out['skip'].append('{}: {}'.format(a.family.name, a.reason))
        else:
            out.setdefault(a.kind, []).append('{} ({})'.format(a.family.name, ', '.join(a.types)))
    return dict((k, v[:MAX_ROWS]) for k, v in out.items())


def lost_after(title, actions, results, index_after):
    """'Doc: Family[: Type]' for every family / type reported ok but not in the target."""
    wanted = dict((a.family.name, a.types) for a in actions if a.kind != 'skip')
    lost = []
    for r in results or []:
        if r.status != 'ok':
            continue
        have = index_after.get(r.family)
        if have is None:
            lost.append('{}: {}'.format(title, r.family))
            continue
        lost += ['{}: {}: {}'.format(title, r.family, t) for t in wanted.get(r.family, [])
                 if t not in have]
    return lost


def _run_plans(ft, src_doc, plans, overwrite_values):
    """[(target doc, [Result], error text or None)], one undo step per target.
    family_transfer.execute_plans when it exists; T3Lab Lite's family_transfer
    has only execute_plan, which raises when a whole target was rolled back."""
    many = getattr(ft, 'execute_plans', None)
    if many is not None:
        return many(src_doc, plans, overwrite_values=overwrite_values)
    out = []
    for dst, actions in plans:
        try:
            out.append((dst, ft.execute_plan(src_doc, dst, actions,
                                             overwrite_values=overwrite_values), None))
        except Exception as ex:
            text = '{}'.format(getattr(ex, 'Message', None) or ex).strip()
            out.append((dst, [], text.splitlines()[0] if text else ex.__class__.__name__))
    return out


def transfer_families(doc, args):
    from Snippets import family_transfer as ft
    try:
        from Snippets import doc_transfer as dt
    except ImportError:         # T3Lab Lite: sources and targets live in family_transfer
        dt = ft

    mode = '{}'.format(args.get('mode') or ft.SKIP).lower()
    if mode not in ft.CONFLICT_MODES:
        return {'error': 'mode must be skip, add or overwrite.'}
    family_names = _names(args.get('family_names'))
    if not family_names:
        return {'error': 'Name the families to copy in family_names.'}
    projects = dt.open_projects(doc.Application)
    sources = dt.list_sources(projects)
    src_title = args.get('source_document') or doc.Title
    src, _missing = match_documents([src_title], sources, lambda e: e.doc.Title)
    if not src:
        return {'error': 'Source "{}" is not an open project or a loaded link.'.format(src_title),
                'sources': [e.label for e in sources]}
    src_doc = src[0].doc
    targets, missing = match_documents(_names(args.get('target_documents')), projects)
    targets = [t for t in targets if not dt.same_doc(t, src_doc)]
    if missing or not targets:
        return {'error': 'Target project(s) not open: {}.'.format(', '.join(missing) or
                                                                  'none other than the source'),
                'open_projects': [p.Title for p in projects if not dt.same_doc(p, src_doc)]}
    selections, no_family, no_type = pick_families(
        ft.collect_families(src_doc), family_names, _names(args.get('type_names')))
    if no_family or no_type:
        return {'error': 'Not in "{}": {}.'.format(src_doc.Title, ', '.join(
                    ['family ' + n for n in no_family] + ['type ' + n for n in no_type])),
                'hint': 'Family and type names are exact (case-insensitive).'}
    plans = [(t, ft.plan_transfer(selections, ft.target_index(t), mode)) for t in targets]
    out = {'source': src_doc.Title, 'mode': mode, 'families': len(selections)}
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False,
                   plans=[dict(plan_rows(actions), document=t.Title) for t, actions in plans])
        return out

    rows, missing_after = [], []
    actions_of = dict((id(t), actions) for t, actions in plans)
    for dst, results, error in _run_plans(ft, src_doc, plans,
                                          bool(args.get('overwrite_values'))):
        row = {'document': dst.Title, 'summary': ft.summarize(results or [])}
        if error:
            row['error'] = '{} All changes to this project were undone.'.format(error)
        failed = [{'family': r.family, 'detail': r.detail} for r in results or []
                  if r.status == 'failed']
        if failed:
            row['failed'] = failed[:MAX_ROWS]
        if results:
            missing_after += lost_after(dst.Title, actions_of.get(id(dst), []), results,
                                        ft.target_index(dst))
        rows.append(row)
    ok = sum(r['summary'].get('ok', 0) for r in rows)
    out.update(success=ok > 0 or all(not r.get('error') and not r.get('failed') for r in rows),
               applied=True, transferred=ok, targets=rows)
    if missing_after:
        out['not_in_model_after'] = missing_after[:MAX_ROWS]
    return out


SPECS = [
    ToolSpec(
        'transfer_families',
        ('Copy loadable families, all types or only type_names, from an open project or a '
         'loaded link into other open projects (the Family Transfer window, without the '
         'window). One undo step per target. Families already in a target: skip, add the '
         'missing types, or overwrite (reloads the source definition). dry_run returns the '
         'plan only.'),
        transfer_families,
        properties={
            'source_document': {'type': 'string', 'description': (
                'Title of an open project or loaded link. Default: active document.')},
            'target_documents': {'type': 'array', 'items': {'type': 'string'},
                                 'description': 'Titles of open projects to copy into.'},
            'family_names': {'type': 'array', 'items': {'type': 'string'}},
            'type_names': {'type': 'array', 'items': {'type': 'string'},
                           'description': 'Only these types. Omit for every type.'},
            'mode': {'type': 'string', 'enum': ['skip', 'add', 'overwrite'], 'default': 'skip'},
            'overwrite_values': {'type': 'boolean',
                                 'description': 'overwrite: also replace parameter values.'},
            'dry_run': {'type': 'boolean', 'description': 'Return the plan, change nothing.'},
        },
        required=['target_documents', 'family_names'], effect='write',
        placement=('modify', 'multidoc'), verify=('self',),
        destructive=('mode', ('overwrite',)),       # the target's family definition is replaced
        title_vi='Chuyển family sang dự án khác',
        keywords=['chuyển family', 'copy family', 'family transfer', 'nạp family'],
        ribbon='FamiTransfer.pushbutton'),
]
