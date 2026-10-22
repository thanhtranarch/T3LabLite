# -*- coding: utf-8 -*-
"""
t3lab_wiki -- the T3Lab construction wiki and drawing rules for EXTERNAL MCP
clients (Claude Desktop, Claude Code, Codex ... via bridge.py).

The pages are the T3Lab Assistant's skill library
(Intelligence/skills_engine.py): the built-in pages in
lib/Intelligence/skills/ -- construction wiki, drawing-completion rules,
naming and BEP playbooks -- plus the user's own skills
(%APPDATA%\\T3LabAI\\skills) and the active project's (its skills/ folder),
so a project BEP sits next to the generic examples and wins over them
(PROJECT_STANDARD_NOTE). One library, two readers: the Assistant
injects a page when its triggers match the request; an MCP client has no such
hook, so it asks for pages through this tool, and `initialize` tells it to
(core.mcp_policy.SERVER_INSTRUCTIONS).

Answered on the HTTP thread with no Revit API call (core.server routes it
before the ExternalEvent): it works with no document open and never waits
for Revit to go idle.

Plain Python, IronPython 2.7 + CPython 3. lite_guard/test_mcp_server.py runs
it without Revit.
"""

from __future__ import unicode_literals

import difflib
import io

TOOL_NAME = 't3lab_wiki'

# The page every drawing task starts from. Named in the tool description, the
# list hint and the initialize instructions; lite_guard checks it ships.
START_PAGE = 'drawing-completion-rules'

ACTIONS = ('list', 'search', 'read')
_DEFAULT_RESULTS = 5
_MAX_RESULTS = 10
# The Assistant caps a page at 6000 characters for small local models. An MCP
# client has room for a whole BEP playbook, so a read returns the file up to
# this size and says when it cut.
_MAX_PAGE_CHARS = 40000

TOOL_SCHEMA = {
    'name': TOOL_NAME,
    'description': (
        'T3Lab construction wiki and drawing rules (pages mostly in Vietnamese). '
        'Read the rules BEFORE you create, annotate, complete or issue drawings: '
        'start with page "' + START_PAGE + '", then the page for the drawing type '
        '(architectural-drawing-rules / structural-drawing-rules / '
        'mep-drawing-rules) and, before an issue, "drawing-issue-checklist". Also '
        'holds construction knowledge (paper sizes, scales, line weights, text '
        'heights, symbols, title block, rebar notation, design stages, '
        'Vietnamese-English terms mapped to Revit categories) and the naming / BEP '
        'standards, plus the user\'s and the active project\'s own pages (a '
        'project BEP wins over the generic examples). action "search" + query '
        'finds the pages for a task; "read" + page returns one page in full; '
        '"list" lists them all. Reads files only -- never touches the model.'),
    'inputSchema': {
        'type': 'object',
        'properties': {
            'action': {
                'type': 'string',
                'enum': list(ACTIONS),
                'description': ('What to do. Default: "read" when page is given, '
                                '"search" when query is given, else "list".'),
            },
            'query': {
                'type': 'string',
                'description': ('For search: the task or topic in the user\'s words, '
                                'Vietnamese with or without diacritics, or English -- '
                                'e.g. "hoàn thiện mặt bằng kết cấu", "rebar notation".'),
            },
            'page': {
                'type': 'string',
                'description': ('For read: a page id from list or search, e.g. "'
                                + START_PAGE + '".'),
            },
            'limit': {
                'type': 'integer',
                'description': ('For search: how many pages to return (default {}, '
                                'max {}).'.format(_DEFAULT_RESULTS, _MAX_RESULTS)),
            },
        },
        'required': [],
    },
}

# Sent with every page that declares `standard: project`. The Assistant's own
# wording (skills_engine._PROJECT_STANDARD_CONTRACT) points at its Settings
# screens; an MCP client gets the same rule without them.
PROJECT_STANDARD_NOTE = (
    'Company-specific page: the codes, prefixes, numbering, names and thresholds '
    'written here are illustrative industry examples, NOT this project\'s '
    'standard. The project\'s BEP / EIR / in-house standard wins -- use it when '
    'the user gives it, or a page with source "project" or "user". Without one, '
    'follow the convention the model already shows (revit_list_sheets, '
    'revit_list_views, list_levels ...) and say you inferred it from the model. '
    'Never invent a project code, sheet number, discipline prefix or revision code.')


def _engine():
    from Intelligence.skills_engine import get_skills_engine
    engine = get_skills_engine()
    # The server lives for the whole Revit session while the user adds skills
    # and switches project, so read the folders again on every call (about 35
    # small files).
    engine.scan()
    return engine


def _pages(engine):
    """{id: meta} of the pages the user has not switched off."""
    return dict((meta['id'], meta) for meta in engine.all_skills()
                if meta.get('enabled', True))


def _row(meta):
    row = {'page': meta['id'],
           'description': meta.get('description') or u'',
           'source': meta.get('source') or u'builtin'}
    if (meta.get('standard') or u'') == 'project':
        row['project_standard'] = True
    return row


def _tokens(text):
    from Intelligence.knowledge import vi_text
    return set(vi_text.tokenize(text or u'', bigrams=True))


def _clamp_limit(value):
    try:
        value = int(value)
    except (TypeError, ValueError):
        return _DEFAULT_RESULTS
    return max(1, min(_MAX_RESULTS, value))


def list_pages(engine):
    pages = _pages(engine)
    return {
        'count': len(pages),
        'pages': [_row(pages[pid]) for pid in sorted(pages)],
        'hint': ('Read a page with action="read", page="<page>". For drawing work '
                 'start with "{}".'.format(START_PAGE)),
    }


def search(engine, query, limit=_DEFAULT_RESULTS):
    """Best pages for `query`: the Assistant's trigger match (specific phrases,
    project pages first) plus word overlap with each page's title,
    description, triggers and text, so an English or loosely worded query
    still finds the Vietnamese page."""
    pages = _pages(engine)
    wanted = _tokens(query)
    trigger_scores = dict(engine.match_scored(query))
    scored = []
    for pid, meta in pages.items():
        head = _tokens(u' '.join([pid.replace(u'-', u' '),
                                  meta.get('name') or u'',
                                  meta.get('description') or u'',
                                  u' '.join(meta.get('triggers') or [])]))
        body = _tokens(engine.get_body(pid))
        score = (trigger_scores.get(pid, 0.0)
                 + 3.0 * len(wanted & head) + 1.0 * len(wanted & body))
        if score > 0:
            scored.append((score, pid))
    scored.sort(key=lambda pair: (-pair[0], pair[1]))
    results = []
    for score, pid in scored[:_clamp_limit(limit)]:
        row = _row(pages[pid])
        row['score'] = round(score, 1)
        results.append(row)
    out = {'query': query, 'count': len(results), 'results': results}
    if results:
        out['hint'] = 'Read the best page with action="read", page="{}".'.format(
            results[0]['page'])
    else:
        out['hint'] = ('No page matches. Try other words, or action="list" to see '
                       'every page.')
    return out


def _page_text(meta):
    """(body, truncated) of one page, frontmatter removed."""
    from Intelligence.skills_engine import parse_frontmatter
    with io.open(meta['path'], 'r', encoding='utf-8', errors='replace') as f:
        _, body = parse_frontmatter(f.read())
    body = (body or u'').strip()
    if len(body) > _MAX_PAGE_CHARS:
        return body[:_MAX_PAGE_CHARS], True
    return body, False


def _page_id(raw):
    """'/Drawing-Completion-Rules.md' -> 'drawing-completion-rules'."""
    pid = (raw or u'').strip().strip(u'/').strip()
    if pid.lower().endswith(u'.md'):
        pid = pid[:-3]
    return pid


def read(engine, page):
    wanted = _page_id(page)
    everything = dict((meta['id'], meta) for meta in engine.all_skills())
    meta = everything.get(wanted)
    if meta is None:
        lowered = dict((pid.lower(), pid) for pid in everything)
        if wanted.lower() in lowered:
            meta = everything[lowered[wanted.lower()]]
    if meta is None:
        close = difflib.get_close_matches(wanted.lower(), sorted(_pages(engine)),
                                          n=5, cutoff=0.5)
        return {'error': 'No wiki page "{}".'.format(page),
                'did_you_mean': close,
                'hint': 'Use action="search" with the topic, or action="list".',
                'tool': TOOL_NAME}
    if not meta.get('enabled', True):
        return {'error': ('Page "{}" is turned off in the T3Lab Assistant '
                          '(LLMs Setting > Skills). Ask the user before relying '
                          'on it.').format(meta['id']),
                'tool': TOOL_NAME}
    try:
        body, truncated = _page_text(meta)
    except Exception as e:
        return {'error': 'Could not read page "{}": {}'.format(meta['id'], e),
                'tool': TOOL_NAME}
    out = _row(meta)
    out['title'] = meta.get('name') or meta['id']
    out['content'] = body
    if truncated:
        out['truncated'] = True
        out['note'] = 'Page cut at {} characters.'.format(_MAX_PAGE_CHARS)
    tools = list(meta.get('tools') or [])
    if tools:
        out['tools'] = tools
    if out.get('project_standard'):
        out['project_standard'] = PROJECT_STANDARD_NOTE
    return out


def call(arguments):
    """One t3lab_wiki tools/call -> result dict (an `error` key on failure)."""
    args = arguments if isinstance(arguments, dict) else {}
    page = u'{}'.format(args.get('page') or u'').strip()
    query = u'{}'.format(args.get('query') or u'').strip()
    action = u'{}'.format(args.get('action') or u'').strip().lower()
    if not action:
        action = 'read' if page else ('search' if query else 'list')
    if action not in ACTIONS:
        return {'error': 'Unknown action "{}".'.format(action),
                'accepted_actions': list(ACTIONS), 'tool': TOOL_NAME}
    if action == 'search' and not query:
        return {'error': 'search needs a query: the task or topic to look up.',
                'tool': TOOL_NAME}
    if action == 'read' and not page:
        return {'error': 'read needs a page id (see action="list").',
                'tool': TOOL_NAME}
    try:
        engine = _engine()
        if action == 'list':
            return list_pages(engine)
        if action == 'search':
            return search(engine, query, args.get('limit'))
        return read(engine, page)
    except Exception as e:
        return {'error': 'The T3Lab wiki could not be read: {}'.format(e),
                'tool': TOOL_NAME}
