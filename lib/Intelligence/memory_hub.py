# -*- coding: utf-8 -*-
"""
Memory Hub

One place to look up what is remembered about the active project, for the
T3Lab Assistant and for AI apps connected through MCP (Claude Desktop,
Claude Code ...). Two sources:

  t3lab   the Assistant's own facts (Intelligence/assistant_memory.py):
          global preferences + this project's conventions. Read AND write.
  claude  Claude Code's memory for the folders linked to the project
          (LLMs Setting > Projects > Link Folder): CLAUDE.md /
          .claude/CLAUDE.md / CLAUDE.local.md in each folder, and Claude
          Code's auto memory ~/.claude/projects/<folder>/memory/*.md.
          Read only -- Claude Code owns those files. Switch off with
          agents.claude_code_memory = false.

Why: every fact used to be pasted into every Assistant prompt. With
recall_tool on (the agent turn has the `recall_memory` tool), a project
with many facts -- or with Claude Code notes -- gets a short index instead
(the newest facts + counts), and the model searches the rest only when the
task needs it. Few facts and no Claude Code notes: the full block, as before.

MCP side (core/server.py): `t3lab_memory` (list / search / read, read-only)
and `t3lab_memory_save` (save / update / forget a T3Lab fact), plus the
resource t3lab://memory/project. They answer on the HTTP thread, no Revit.

Plain Python, IronPython 2.7 + CPython 3, no Revit import.
lite_guard/test_memory_hub.py runs it without Revit.

Author: Tran Tien Thanh
"""

from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Memory Hub"

import glob
import io
import os
import re
import time

RECALL_TOOL_NAME = 'recall_memory'
READ_TOOL_NAME = 't3lab_memory'
SAVE_TOOL_NAME = 't3lab_memory_save'
RESOURCE_URI = 't3lab://memory/project'
SETTING_KEY = 'claude_code_memory'

# Inline every fact while the block stays this small and there is nothing
# from Claude Code; past it, the prompt carries an index instead.
COMPACT_MIN_FACTS = 12
COMPACT_MIN_CHARS = 1500
# How much of the newest facts the index still shows inline.
INDEX_FACT_CHARS = 800

_MAX_SOURCES = 40              # Claude Code files per project
_MAX_FILE_CHARS = 20000        # read cap per file
_CHUNK_CHARS = 700             # search unit inside a Claude Code file
_DEFAULT_RESULTS = 5
_MAX_RESULTS = 10

_FOLDER_FILES = ('CLAUDE.md', os.path.join('.claude', 'CLAUDE.md'),
                 'CLAUDE.local.md')

# Linked folders are often network shares, and every Assistant prompt asks
# for the source list: probe the disk at most this often per project.
_SOURCES_TTL = 30.0
_sources_cache = {}            # (pid, enabled) -> (read_at, sources)


def clear_cache():
    _sources_cache.clear()


# ─── Active project ───────────────────────────────────────────────────────────

def active_project_id():
    try:
        from config.project_store import ProjectStore
        return ProjectStore().get_active_project_id()
    except Exception:
        return None


def _project_dirs(project_id):
    """Folders linked to the project (where Claude Code would be opened)."""
    if not project_id:
        return []
    try:
        from config.project_store import ProjectStore
        return [d for d in ProjectStore().get_knowledge_dirs(project_id) or []
                if d and os.path.isdir(d)]
    except Exception:
        return []


def claude_code_enabled():
    try:
        from config.settings import get_settings
        return get_settings().get_agent_option(SETTING_KEY, True) is not False
    except Exception:
        return True


# ─── Claude Code memory ──────────────────────────────────────────────────────

def claude_config_dir():
    """Claude Code's data folder: $CLAUDE_CONFIG_DIR, else ~/.claude."""
    env = os.environ.get('CLAUDE_CONFIG_DIR')
    if env:
        return env
    return os.path.join(os.path.expanduser('~'), '.claude')


def encode_project_dir(path):
    """Claude Code's folder name for a project directory: the absolute path
    with every character that is not a letter or digit turned into '-'
    (C:\\Jobs\\Tower A -> C--Jobs-Tower-A)."""
    return re.sub(r'[^A-Za-z0-9]', '-', os.path.abspath(path))


def _auto_memory_dirs(folder):
    """~/.claude/projects/<encoded folder>/memory, matched without regard to
    case (a drive letter typed as c: or C:)."""
    root = os.path.join(claude_config_dir(), 'projects')
    want = encode_project_dir(folder).lower()
    try:
        names = os.listdir(root)
    except Exception:
        return []
    return [os.path.join(root, n, 'memory') for n in names
            if n.lower() == want and os.path.isdir(os.path.join(root, n, 'memory'))]


def claude_sources(project_id):
    """Claude Code memory files for the project: [{'path', 'kind', 'folder'}],
    in a stable order. [] when switched off or nothing is linked. Cached for
    _SOURCES_TTL seconds."""
    if not claude_code_enabled():
        return []
    key = project_id
    hit = _sources_cache.get(key)
    if hit and time.time() - hit[0] < _SOURCES_TTL:
        return list(hit[1])
    out = _find_sources(project_id)
    _sources_cache[key] = (time.time(), out)
    return list(out)


def _find_sources(project_id):
    out, seen = [], set()
    for folder in _project_dirs(project_id):
        found = []
        for rel in _FOLDER_FILES:
            found.append((os.path.join(folder, rel), 'CLAUDE.md'))
        for mdir in _auto_memory_dirs(folder):
            for path in sorted(glob.glob(os.path.join(mdir, '*.md'))):
                found.append((path, 'auto memory'))
        for path, kind in found:
            key = os.path.normcase(os.path.abspath(path))
            if key in seen or not os.path.isfile(path):
                continue
            seen.add(key)
            out.append({'path': path, 'kind': kind, 'folder': folder})
            if len(out) >= _MAX_SOURCES:
                return out
    return out


def _read_text(path, limit=_MAX_FILE_CHARS):
    try:
        with io.open(path, 'r', encoding='utf-8', errors='replace') as f:
            text = f.read(limit + 1)
    except Exception:
        return u'', False
    if len(text) > limit:
        return text[:limit], True
    return text, False


def _chunks(text):
    """Heading / paragraph pieces of at most _CHUNK_CHARS each."""
    parts, cur = [], []
    for line in text.splitlines():
        starts = line.lstrip().startswith('#') or not line.strip()
        if starts and cur and sum(len(l) + 1 for l in cur) > 80:
            parts.append(u'\n'.join(cur).strip())
            cur = []
        cur.append(line)
        if sum(len(l) + 1 for l in cur) >= _CHUNK_CHARS:
            parts.append(u'\n'.join(cur).strip())
            cur = []
    if cur:
        parts.append(u'\n'.join(cur).strip())
    return [p for p in parts if p]


# ─── Entries & search ─────────────────────────────────────────────────────────

def _facts(project_id):
    try:
        from Intelligence import assistant_memory
        return assistant_memory.list_facts(project_id)
    except Exception:
        return []


def _fact_entry(number, scope, fact):
    return {'id': u'fact:{}'.format(number), 'source': 't3lab',
            'scope': u'all projects' if scope == 'global' else u'this project',
            'text': fact.get('text') or u'', 'saved': fact.get('created') or u''}


def _tokens(text):
    try:
        from Intelligence.knowledge import vi_text
        return set(vi_text.tokenize(text or u'', bigrams=True))
    except Exception:
        return set(w for w in re.split(r'\W+', (text or u'').lower()) if len(w) > 1)


def _clamp(value, default=_DEFAULT_RESULTS):
    try:
        value = int(value)
    except (TypeError, ValueError):
        return default
    return max(1, min(_MAX_RESULTS, value))


def stats(project_id=None):
    facts = _facts(project_id)
    return {'facts': len(facts),
            'fact_chars': sum(len(f.get('text') or u'') for _s, f in facts),
            'claude_files': len(claude_sources(project_id))}


def search(query, project_id=None, limit=_DEFAULT_RESULTS):
    """Best matches for `query` across both sources. T3Lab facts weigh a
    little more (each one is a curated statement); among equals the newer
    fact wins."""
    wanted = _tokens(query)
    scored = []
    facts = _facts(project_id)
    for i, (scope, fact) in enumerate(facts):
        hit = len(wanted & _tokens(fact.get('text')))
        if hit:
            scored.append((hit * 1.5, i, _fact_entry(i + 1, scope, fact)))
    for k, src in enumerate(claude_sources(project_id)):
        text, _cut = _read_text(src['path'])
        for j, chunk in enumerate(_chunks(text)):
            hit = len(wanted & _tokens(chunk))
            if hit:
                scored.append((float(hit), -1, {
                    'id': u'claude:{}'.format(k + 1), 'source': 'claude',
                    'kind': src['kind'], 'file': src['path'],
                    'text': chunk}))
    scored.sort(key=lambda row: (-row[0], -row[1]))
    results = [row[2] for row in scored[:_clamp(limit)]]
    out = {'query': query, 'count': len(results), 'results': results}
    if not results:
        out['hint'] = ('Nothing matches. Try other words, or action="list" '
                       'to see what is saved.')
    return out


def read(entry_id, project_id=None):
    """One entry by id: 'fact:<n>' or 'claude:<n>' (a whole file)."""
    entry_id = u'{}'.format(entry_id or u'').strip().lower()
    kind, _sep, num = entry_id.partition(':')
    try:
        n = int(num)
    except ValueError:
        n = 0
    if kind == 'fact':
        facts = _facts(project_id)
        if 1 <= n <= len(facts):
            scope, fact = facts[n - 1]
            return _fact_entry(n, scope, fact)
    elif kind == 'claude':
        sources = claude_sources(project_id)
        if 1 <= n <= len(sources):
            src = sources[n - 1]
            text, cut = _read_text(src['path'])
            out = {'id': u'claude:{}'.format(n), 'source': 'claude',
                   'kind': src['kind'], 'file': src['path'], 'text': text}
            if cut:
                out['truncated'] = True
            return out
    return {'error': u'No memory entry "{}".'.format(entry_id),
            'hint': 'Use an id from action="list" or "search".'}


def list_all(project_id=None):
    facts = _facts(project_id)
    return {
        'project_id': project_id,
        'facts': [_fact_entry(i + 1, s, f) for i, (s, f) in enumerate(facts)],
        'claude_files': [{'id': u'claude:{}'.format(k + 1), 'kind': s['kind'],
                          'file': s['path']}
                         for k, s in enumerate(claude_sources(project_id))],
        'hint': ('Read a Claude Code file with action="read", id="claude:<n>". '
                 'Save a new convention with t3lab_memory_save.'),
    }


# ─── Assistant prompt block ──────────────────────────────────────────────────

def prompt_block(project_id=None, max_chars=None, recall_tool=False):
    """The memory section of an Assistant prompt.

    recall_tool=False (no tools on this turn): every fact, as before.
    recall_tool=True: every fact while the set is small and nothing comes
    from Claude Code; otherwise the newest facts plus a line naming what
    `recall_memory` can still find.
    """
    from Intelligence import assistant_memory
    if not recall_tool:
        return assistant_memory.build_memory_block(
            project_id, include_tool_hint=False, max_chars=max_chars)
    facts = _facts(project_id)
    fact_chars = sum(len(f.get('text') or u'') for _s, f in facts)
    n_claude = len(claude_sources(project_id))
    if not n_claude and len(facts) <= COMPACT_MIN_FACTS \
            and fact_chars <= COMPACT_MIN_CHARS:
        return assistant_memory.build_memory_block(
            project_id, include_tool_hint=True, max_chars=max_chars)
    budget = min(INDEX_FACT_CHARS, max_chars) if max_chars else INDEX_FACT_CHARS
    shown, used = [], 0
    for i in range(len(facts) - 1, -1, -1):          # newest first
        scope, fact = facts[i]
        tag = u'all projects' if scope == 'global' else u'this project'
        line = u'{}. [{}] {}'.format(i + 1, tag, fact.get('text') or u'')
        if shown and used + len(line) + 1 > budget:
            break
        shown.insert(0, line)
        used += len(line) + 1
    hidden = len(facts) - len(shown)
    lines = [u'## Persistent memory (index — newest facts saved in previous '
             u'chats; apply them without re-asking)']
    lines.extend(shown)
    more = []
    if hidden:
        more.append(u'{} more saved fact(s)'.format(hidden))
    if n_claude:
        more.append(u'{} Claude Code memory file(s) of this project'.format(
            n_claude))
    if more:
        lines.append(
            u'Not shown: {}. Before you assume a project convention (naming, '
            u'levels, sheet numbers, standards), call `{}` with the topic in '
            u'a few words; it searches all of them.'.format(
                u' and '.join(more), RECALL_TOOL_NAME))
    lines.append(assistant_memory.TOOL_HINT)
    return u'\n'.join(lines)


# ─── Tool schemas ─────────────────────────────────────────────────────────────

RECALL_TOOL_SCHEMA = {
    'name': RECALL_TOOL_NAME,
    'description': (
        "Search what is remembered about this project and this user: facts "
        "saved in earlier chats and the project's Claude Code memory "
        "(CLAUDE.md, auto memory). Use it before assuming a convention "
        "(naming, levels, sheet numbers, standards) that the memory index in "
        "your instructions says is not shown. Vietnamese or English. Read "
        "only; save facts with remember_fact."),
    'inputSchema': {
        'type': 'object',
        'properties': {
            'query': {'type': 'string',
                      'description': 'The topic in a few words, e.g. "sheet numbering".'},
            'id': {'type': 'string',
                   'description': 'Read one entry in full (fact:<n> or claude:<n>).'},
            'limit': {'type': 'integer', 'minimum': 1, 'maximum': _MAX_RESULTS},
        },
        'required': [],
    },
}

READ_TOOL_SCHEMA = {
    'name': READ_TOOL_NAME,
    'description': (
        "What the T3Lab Assistant and Claude Code remember about the Revit "
        "project open in T3Lab: saved conventions and preferences (sheet "
        "numbering, level names, standards, how the user wants answers) and "
        "the project's Claude Code memory files. Search it before asking the "
        "user or assuming a convention. action \"search\" + query, \"read\" + "
        "id (fact:<n> / claude:<n>), \"list\" for everything. Read only; "
        "save with t3lab_memory_save."),
    'inputSchema': {
        'type': 'object',
        'properties': {
            'action': {'type': 'string', 'enum': ['list', 'search', 'read'],
                       'description': ('Default: "read" when id is given, '
                                       '"search" when query is given, else "list".')},
            'query': {'type': 'string', 'description': 'For search: the topic.'},
            'id': {'type': 'string', 'description': 'For read: fact:<n> or claude:<n>.'},
            'limit': {'type': 'integer', 'minimum': 1, 'maximum': _MAX_RESULTS},
        },
        'required': [],
    },
}

SAVE_TOOL_SCHEMA = {
    'name': SAVE_TOOL_NAME,
    'description': (
        "Save, update or forget one durable fact in the T3Lab project memory "
        "the T3Lab Assistant also uses, so the next chat (in Revit or here) "
        "starts knowing it. Only lasting conventions or preferences stated "
        "by the user -- never one-off requests, secrets or element ids. "
        "Claude Code's own files are not changed."),
    'inputSchema': {
        'type': 'object',
        'properties': {
            'action': {'type': 'string', 'enum': ['save', 'update', 'forget'],
                       'description': 'save (default), update a changed fact, forget a cancelled one.'},
            'fact': {'type': 'string',
                     'description': 'One short English sentence (for forget: its gist).'},
            'scope': {'type': 'string', 'enum': ['project', 'global'],
                      'description': 'project (default) = this Revit project; global = every project.'},
            'replaces': {'type': 'string',
                         'description': 'For update / forget: the old fact\'s gist or its fact:<n> number.'},
        },
        'required': ['fact'],
    },
}


def call_read(arguments, project_id=None):
    """t3lab_memory / recall_memory -> result dict (an `error` key on failure)."""
    args = arguments if isinstance(arguments, dict) else {}
    pid = project_id if project_id is not None else active_project_id()
    entry = u'{}'.format(args.get('id') or u'').strip()
    query = u'{}'.format(args.get('query') or u'').strip()
    action = u'{}'.format(args.get('action') or u'').strip().lower()
    if not action:
        action = 'read' if entry else ('search' if query else 'list')
    try:
        if action == 'read':
            if not entry:
                return {'error': 'read needs an id (see action="list").'}
            return read(entry, pid)
        if action == 'search':
            if not query:
                return {'error': 'search needs a query: the topic to look up.'}
            return search(query, pid, args.get('limit'))
        if action == 'list':
            return list_all(pid)
    except Exception as e:
        return {'error': u'The project memory could not be read: {}'.format(e)}
    return {'error': u'Unknown action "{}".'.format(action),
            'accepted_actions': ['list', 'search', 'read']}


def _fact_number(text):
    m = re.match(r'^\s*fact:(\d+)\s*$', u'{}'.format(text or u''), re.I)
    return int(m.group(1)) if m else None


def call_save(arguments, project_id=None):
    """t3lab_memory_save -> result dict."""
    from Intelligence import assistant_memory
    args = arguments if isinstance(arguments, dict) else {}
    pid = project_id if project_id is not None else active_project_id()
    action = u'{}'.format(args.get('action') or u'save').strip().lower()
    fact = u'{}'.format(args.get('fact') or u'').strip()
    scope = u'{}'.format(args.get('scope') or u'project').strip().lower()
    if scope not in ('project', 'global'):
        scope = 'project'
    replaces = args.get('replaces')
    number = _fact_number(replaces)
    try:
        if action == 'save':
            ok, note = assistant_memory.add_fact(fact, scope=scope, project_id=pid)
        elif action == 'update':
            ok, note = assistant_memory.update_fact(
                number if number else (replaces or fact), fact,
                scope=scope, project_id=pid)
        elif action == 'forget':
            if number:
                ok, removed = assistant_memory.remove_fact(number, project_id=pid)
            else:
                ok, removed = assistant_memory.forget_fact(replaces or fact,
                                                           project_id=pid)
            note = (u'Forgot: {}'.format(removed) if ok
                    else u'No matching fact to forget.')
        else:
            return {'error': u'Unknown action "{}".'.format(action),
                    'accepted_actions': ['save', 'update', 'forget']}
    except Exception as e:
        return {'error': u'The project memory could not be written: {}'.format(e)}
    if ok:
        return {'success': True, 'note': note}
    return {'error': note or u'Nothing was saved.'}


# ─── MCP resource ─────────────────────────────────────────────────────────────

def resource_entry():
    return {'uri': RESOURCE_URI, 'name': 'project-memory',
            'title': 'T3Lab project memory',
            'description': ('Conventions and preferences saved for the '
                            'active T3Lab project, and its Claude Code '
                            'memory files.'),
            'mimeType': 'text/markdown',
            'annotations': {'audience': ['assistant', 'user'], 'priority': 0.8}}


def read_resource(uri, project_id=None):
    if u'{}'.format(uri or u'') != RESOURCE_URI:
        return None
    pid = project_id if project_id is not None else active_project_id()
    data = list_all(pid)
    lines = [u'# T3Lab project memory', u'']
    if data['facts']:
        for f in data['facts']:
            lines.append(u'- ({}) [{}] {}'.format(f['id'], f['scope'], f['text']))
    else:
        lines.append(u'No saved facts yet.')
    if data['claude_files']:
        lines += [u'', u'## Claude Code memory files', u'']
        for c in data['claude_files']:
            lines.append(u'- ({}) {} — {}'.format(c['id'], c['kind'], c['file']))
    return {'contents': [{'uri': RESOURCE_URI, 'mimeType': 'text/markdown',
                          'text': u'\n'.join(lines)}]}
