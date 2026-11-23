# -*- coding: utf-8 -*-
"""View tools as specs: transfer_view_templates (View Template Transfer window, headless).
Doc: docs/agents/headless-tools.md. Test: dev/test_headless_tools.py."""
from __future__ import unicode_literals

from core.tool_specs import ToolSpec
from core.tools._common import MAX_ROWS, names as _names, norm as _norm


def match_documents(wanted, docs, title=lambda d: d.Title):
    """(found docs in order, names not found); case- and space-insensitive,
    '.rvt' optional."""
    def key(text):
        k = _norm(text)
        return k[:-4] if k.endswith('.rvt') else k
    index = {}
    for d in docs:
        index.setdefault(key(title(d)), d)
    found, missing = [], []
    for name in wanted:
        d = index.get(key(name))
        if d is None:
            missing.append(name)
        elif d not in found:
            found.append(d)
    return found, missing


def pick_templates(templates, names, view_types):
    """(selected TemplateInfo, names with no template). No names = all."""
    types = set(_norm(v) for v in view_types)
    pool = [t for t in templates if not types or _norm(t.view_type) in types]
    if not names:
        return pool, []
    by_name = {}
    for t in pool:
        by_name.setdefault(_norm(t.name), []).append(t)
    picked, missing = [], []
    for n in names:
        hits = by_name.get(_norm(n))
        if hits:
            picked += [t for t in hits if t not in picked]
        else:
            missing.append(n)
    return picked, missing


def plan_rows(actions):
    """{'copy': [...], 'overwrite': [...], 'skip': [...]} as 'View type: name'."""
    out = {'copy': [], 'overwrite': [], 'skip': []}
    for a in actions:
        out.setdefault(a.kind, []).append('{}: {}'.format(a.template.view_type, a.template.name))
    return out


def transfer_view_templates(doc, args):
    from Snippets import view_template_transfer as vtt
    from Snippets import doc_transfer as dt

    mode = '{}'.format(args.get('mode') or vtt.SKIP).lower()
    if mode not in vtt.CONFLICT_MODES:
        return {'error': 'mode must be skip or overwrite.'}
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
    templates, not_found = pick_templates(vtt.collect_templates(src_doc),
                                          _names(args.get('template_names')),
                                          _names(args.get('view_types')))
    if not_found:
        return {'error': 'No view template named: {}.'.format(', '.join(not_found)),
                'hint': 'manage_view_template(operation="list") on the source lists them.'}
    if not templates:
        return {'error': 'The source has no view template to copy.'}
    plans = [(t, vtt.plan_transfer(templates, vtt.target_index(t), mode)) for t in targets]
    out = {'source': src_doc.Title, 'mode': mode, 'templates': len(templates)}
    if args.get('dry_run'):
        out.update(dry_run=True, applied=False,
                   plans=[dict(plan_rows(actions), document=t.Title) for t, actions in plans])
        return out

    rows, missing_after = [], []
    for dst, results, error in vtt.execute_plans(src_doc, plans):
        row = {'document': dst.Title, 'summary': vtt.summarize(results or [])}
        if error:
            row['error'] = error
        failed = [{'template': r.template, 'detail': r.detail} for r in results or []
                  if r.status == 'failed']
        if failed:
            row['failed'] = failed[:MAX_ROWS]
        # Read-back: every template reported ok must now be in the target.
        names_after = set(_norm(k[1]) for k in vtt.target_index(dst))
        lost = [r.template for r in results or []
                if r.status == 'ok' and _norm(r.template) not in names_after]
        if lost:
            missing_after += ['{}: {}'.format(dst.Title, n) for n in lost]
        rows.append(row)
    ok = sum(r['summary'].get('ok', 0) for r in rows)
    out.update(success=ok > 0 or all(not r.get('error') and not r.get('failed') for r in rows),
               applied=True, copied=ok, targets=rows)
    if missing_after:
        out['not_in_model_after'] = missing_after
    return out


SPECS = [
    ToolSpec(
        'transfer_view_templates',
        ('Copy view templates from an open project or a loaded link into other open projects '
         '(the View Template Transfer window, without the window). Revit brings the filters '
         'and patterns they use. mode=overwrite replaces same-name templates and moves their '
         'views to the new copy. dry_run returns the plan only.'),
        transfer_view_templates,
        properties={
            'source_document': {'type': 'string', 'description': (
                'Title of an open project or loaded link. Default: active document.')},
            'target_documents': {'type': 'array', 'items': {'type': 'string'},
                                 'description': 'Titles of open projects to copy into.'},
            'template_names': {'type': 'array', 'items': {'type': 'string'},
                               'description': 'Exact names. Omit to copy every template.'},
            'view_types': {'type': 'array', 'items': {'type': 'string'},
                           'description': 'Only these view types, e.g. "Floor Plan", "Section".'},
            'mode': {'type': 'string', 'enum': ['skip', 'overwrite'], 'default': 'skip',
                     'description': 'For templates already in a target.'},
            'dry_run': {'type': 'boolean', 'description': 'Return the plan, change nothing.'},
        },
        required=['target_documents'], effect='write', placement='modify', verify=('self',),
        destructive=('mode', ('overwrite',)),       # the target's own template is deleted
        title_vi='Chuyển view template sang dự án khác',
        keywords=['view template', 'chuyển template', 'copy template'],
        ribbon='ViewTemplateTransfer.pushbutton',
        requires=('Snippets/view_template_transfer.py', 'Snippets/doc_transfer.py')),
]
