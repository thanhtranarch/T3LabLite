# -*- coding: utf-8 -*-
"""T3Lab BIM Harness, store side: one shared task board for several AI agents
working in the same Revit session (docs/agents/harness.md).

A lead opens a session and plans tasks; workers take the next ready task,
report, and the lead (or a reviewer) accepts or sends it back. The board only
coordinates: what an agent may do in Revit is still decided by its gateway
profile (docs/agents/safety.md).

Runs in the bridge process (CPython 3, outside Revit); deployed next to
bridge.py as a companion. Stdlib only. State is one SQLite file under
harness_dir(), so stdio bridges, the HTTP router and the CLI all see it.
Never put it on a synced or network folder (SQLite locks are unreliable there).

board.json next to it is a snapshot of the open sessions, rewritten after every
change: the readers that cannot use sqlite3 (IronPython in Revit: Agent
Cursors, the toast watcher; MCP Control) read that file, never the database.

Test: dev/test_harness.py.
"""
import contextlib
import functools
import json
import os
import re
import sqlite3
import threading
import time
import uuid

SCHEMA_VERSION = 3              # 2: tasks.started / finished; 3: sessions.exclusive / write_mode
CONFIG_FILE = 'config.json'     # the owner's setup: master agent and session defaults
# scoped: write tasks run together when their scopes differ. serial: one write
# task per document at a time, whatever the scopes say.
WRITE_MODES = ('scoped', 'serial')
BOARD_FILE = 'board.json'
BOARD_FORMAT = 'harness-board/1'
ACTIVE_S = 12 * 3600            # a session idle for longer no longer limits writes
OWNER = 'owner'                 # the person at the machine, in notes and reviews
ROLES = ('lead', 'worker', 'reviewer')
# todo -> claimed -> review -> done; failed and redo go back through todo.
STATUSES = ('todo', 'claimed', 'review', 'done', 'failed')
REPORT_STATUSES = ('progress', 'done', 'failed')
VERDICTS = ('accept', 'redo')
LEASE_S = 30 * 60               # a claim without a report returns to todo
MAX_TASKS = 200
MAX_NOTES_SHOWN = 10
RETRY_AFTER_S = 20
_LIMITS = {'goal': 1000, 'title': 200, 'detail': 4000, 'done_when': 1000,
           'summary': 4000, 'note': 2000, 'key': 40, 'member': 80}
# Scope dimensions; a missing one means "everything" in that dimension.
SCOPE_DIMS = ('categories', 'levels', 'views', 'worksets', 'element_ids')

_SCHEMA = """
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS sessions (
    id TEXT PRIMARY KEY, goal TEXT, document TEXT, lead TEXT, status TEXT,
    review INTEGER, created REAL, closed REAL, exclusive INTEGER, write_mode TEXT);
CREATE TABLE IF NOT EXISTS members (
    session TEXT, member TEXT, role TEXT, platform TEXT, joined REAL, last_seen REAL,
    PRIMARY KEY (session, member));
CREATE TABLE IF NOT EXISTS tasks (
    session TEXT, key TEXT, seq INTEGER, title TEXT, detail TEXT, scope TEXT,
    depends TEXT, assignee TEXT, done_when TEXT, status TEXT, owner TEXT,
    lease_until REAL, attempts INTEGER, summary TEXT, result TEXT, updated REAL,
    started REAL, finished REAL,
    PRIMARY KEY (session, key));
CREATE TABLE IF NOT EXISTS notes (
    id INTEGER PRIMARY KEY AUTOINCREMENT, session TEXT, task TEXT, member TEXT,
    text TEXT, at REAL);
"""


class HarnessError(Exception):
    """A refusal the agent can act on: message + stable code (+ hint)."""

    def __init__(self, message, code, retryable=False, **extra):
        Exception.__init__(self, message)
        self.payload = dict({'error': message, 'error_code': code,
                             'retryable': retryable}, **extra)


def harness_dir():
    """%APPDATA%\\T3LabAI\\harness, or T3LAB_HARNESS_DIR (tests)."""
    override = os.environ.get('T3LAB_HARNESS_DIR')
    if override:
        return override
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI', 'harness')


def _mutates(method):
    """Public method that changes the board: one writer at a time in this
    process, then a fresh board.json."""
    @functools.wraps(method)
    def wrapper(self, *args, **kwargs):
        with self._lock:
            try:
                return method(self, *args, **kwargs)
            finally:
                self._publish()
    return wrapper


def _text(value, field, required=False):
    text = u'{}'.format(value if value is not None else '').strip()
    if required and not text:
        raise HarnessError('"{}" is required.'.format(field), 'bad_arguments')
    return text[:_LIMITS[field]]


def _str_list(value):
    if value in (None, ''):
        return []
    if not isinstance(value, (list, tuple)):
        value = [value]
    return [u'{}'.format(v).strip() for v in value if u'{}'.format(v).strip()]


def clean_scope(scope):
    """{'write': bool, 'document': str, <dim>: [lower-case strings]} from loose input."""
    scope = scope if isinstance(scope, dict) else {}
    out = {'write': scope.get('write', True) is not False}
    if scope.get('document'):
        out['document'] = u'{}'.format(scope['document']).strip()
    for dim in SCOPE_DIMS:
        values = sorted(set(v.lower() for v in _str_list(scope.get(dim))))
        if values:
            out[dim] = values
    return out


def load_config(root=None):
    """The owner's setup: {'master', 'exclusive', 'write_mode'}.

    master: the only agent that may open sessions ('' = any person-facing
    client). exclusive: while a session is open, every other agent needs a
    task to change its documents, joined or not; on by default once a master
    is named. A missing or broken file is the open default."""
    out = {'master': '', 'exclusive': False, 'write_mode': WRITE_MODES[0]}
    try:
        with open(os.path.join(root or harness_dir(), CONFIG_FILE), encoding='utf-8-sig') as f:
            data = json.load(f)
    except (OSError, ValueError):
        return out
    if not isinstance(data, dict):
        return out
    out['master'] = u'{}'.format(data.get('master') or '').strip().lower()
    out['exclusive'] = bool(data.get('exclusive', bool(out['master'])))
    if data.get('write_mode') in WRITE_MODES:
        out['write_mode'] = data['write_mode']
    return out


def save_config(root=None, master=None, exclusive=None, write_mode=None):
    """Change the given keys of the setup; returns the new setup."""
    root = root or harness_dir()
    config = load_config(root)
    if master is not None:
        master = u'{}'.format(master).strip().lower()
        if master and not re.match(r'^[a-z0-9][a-z0-9._-]{1,39}$', master):
            raise HarnessError('master must be an agent id (a-z 0-9 . _ -), or "" for none.',
                               'bad_arguments')
        config['master'] = master
        if exclusive is None:               # naming a master turns exclusive on
            config['exclusive'] = bool(master)
    if exclusive is not None:
        config['exclusive'] = bool(exclusive)
    if write_mode is not None:
        if write_mode not in WRITE_MODES:
            raise HarnessError('write_mode must be one of: {}.'.format(', '.join(WRITE_MODES)),
                               'bad_arguments')
        config['write_mode'] = write_mode
    os.makedirs(root, exist_ok=True)
    path = os.path.join(root, CONFIG_FILE)
    tmp = '{}.{}.tmp'.format(path, uuid.uuid4().hex[:8])
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2)
    os.replace(tmp, path)
    return config


def scopes_conflict(a, b, default_document='', serial=False):
    """Two tasks may not run together: both write to the same document and no
    dimension both of them name keeps them apart (serial: nothing does)."""
    if not (a.get('write', True) and b.get('write', True)):
        return False
    doc_a = (a.get('document') or default_document or '').lower()
    doc_b = (b.get('document') or default_document or '').lower()
    if doc_a != doc_b:
        return False
    if serial:
        return True
    for dim in SCOPE_DIMS:
        if a.get(dim) and b.get(dim) and not set(a[dim]) & set(b[dim]):
            return False
    return True


def call_element_ids(arguments, limit=500):
    """Element ids a tool call targets, as strings: element_id / element_ids /
    ids at any of three levels, and `id` inside an `elements` list. Other *_id
    arguments (view_id, level_id, type_id ...) are references, not targets:
    counting them would refuse honest calls. "@saved" sets and anything not an
    integer are left out: unknown, not checked."""
    out = []

    def add(value):
        text = u'{}'.format(value).strip()
        if re.match(r'^\d+$', text) and text not in out and not isinstance(value, bool):
            out.append(text)

    def walk(node, depth, in_elements):
        if depth > 3 or len(out) >= limit:
            return
        if isinstance(node, dict):
            for key, value in node.items():
                k = u'{}'.format(key).lower()
                if k in ('element_id', 'element_ids', 'ids') or (k == 'id' and in_elements):
                    for v in (value if isinstance(value, (list, tuple)) else [value]):
                        add(v)
                elif isinstance(value, (dict, list, tuple)):
                    walk(value, depth + 1, k == 'elements')
        elif isinstance(node, (list, tuple)):
            for item in node:
                walk(item, depth + 1, in_elements)

    walk(arguments, 0, False)
    return out[:limit]


_BOARD_CACHE = {}


def read_board(root=None):
    """board.json as a dict ({} when missing or mid-write). Cached by mtime:
    cheap enough for every tool call."""
    path = os.path.join(root or harness_dir(), BOARD_FILE)
    try:
        stamp = os.stat(path)
        key = (stamp.st_mtime_ns, stamp.st_size)
    except OSError:
        _BOARD_CACHE.pop(path, None)
        return {}
    cached = _BOARD_CACHE.get(path)
    if cached and cached[0] == key:
        return cached[1]
    try:
        with open(path, encoding='utf-8') as f:
            board = json.load(f)
    except (OSError, ValueError):
        return cached[1] if cached else {}
    board = board if isinstance(board, dict) else {}
    _BOARD_CACHE[path] = (key, board)
    return board


def is_member(member, agent):
    """Board member of this agent: the id itself or one of its chats (agent/<as>)."""
    return member == agent or (member or '').startswith(agent + '/')


def task_of(agent, board):
    """The one task `agent` holds, from a board snapshot, as
    {session, task, title, member, document}; None when it holds none, or
    several (two chats of one client: which one calls is unknown)."""
    found = []
    for session in (board or {}).get('sessions') or []:
        for task in session.get('tasks') or []:
            if task.get('status') == 'claimed' and is_member(task.get('owner') or '', agent):
                found.append({'session': session.get('session'), 'task': task.get('task'),
                              'title': task.get('title'), 'member': task.get('owner'),
                              'document': task.get('document') or session.get('document') or ''})
    return found[0] if len(found) == 1 else None


class Harness(object):
    """The task board. Every method returns a JSON-safe dict or raises HarnessError."""

    _lock = threading.RLock()
    _ready = set()              # databases this process already created / upgraded

    def __init__(self, root=None, now=None):
        self.root = root or harness_dir()
        self._now = now or time.time
        self.path = os.path.join(self.root, 'harness.db')
        if self.path in self._ready and os.path.exists(self.path):
            return              # built for every gateway call: no schema work twice
        os.makedirs(self.root, exist_ok=True)
        with self._db() as db:
            db.executescript(_SCHEMA)
            columns = [r['name'] for r in db.execute('PRAGMA table_info(tasks)')]
            for column in ('started', 'finished'):      # schema 1 -> 2
                if column not in columns:
                    db.execute('ALTER TABLE tasks ADD COLUMN {} REAL'.format(column))
            columns = [r['name'] for r in db.execute('PRAGMA table_info(sessions)')]
            for column, kind in (('exclusive', 'INTEGER'), ('write_mode', 'TEXT')):   # 2 -> 3
                if column not in columns:
                    db.execute('ALTER TABLE sessions ADD COLUMN {} {}'.format(column, kind))
            db.execute('INSERT OR REPLACE INTO meta VALUES (?, ?)',
                       ('schema', str(SCHEMA_VERSION)))
        self._ready.add(self.path)

    def publish(self):
        """board.json again, after something outside the board changed (the setup)."""
        with self._lock:
            self._publish()

    def _publish(self):
        """Rewrite board.json from the database when the board changed. Best
        effort: a reader holding the file open makes the replace fail; the
        next change writes it again."""
        tmp = None
        try:
            with self._db() as db:
                sessions = []
                for srow in db.execute("SELECT * FROM sessions WHERE status = 'open' "
                                       'ORDER BY created').fetchall():
                    sid = srow['id']
                    members = db.execute('SELECT * FROM members WHERE session = ? '
                                         'ORDER BY joined', (sid,)).fetchall()
                    tasks, counts = [], {}
                    for row in db.execute('SELECT * FROM tasks WHERE session = ? ORDER BY seq',
                                          (sid,)).fetchall():
                        scope = json.loads(row['scope'] or '{}')
                        counts[row['status']] = counts.get(row['status'], 0) + 1
                        tasks.append({'task': row['key'], 'title': row['title'],
                                      'status': row['status'], 'owner': row['owner'],
                                      'assignee': row['assignee'],
                                      'document': scope.get('document') or '',
                                      'write': scope.get('write', True),
                                      'started': row['started'], 'finished': row['finished']})
                    sessions.append({
                        'session': sid, 'goal': srow['goal'], 'document': srow['document'] or '',
                        'lead': srow['lead'], 'review_required': bool(srow['review']),
                        'exclusive': bool(srow['exclusive']),
                        'write_mode': srow['write_mode'] or WRITE_MODES[0],
                        'members': [{'member': m['member'], 'role': m['role'],
                                     'platform': m['platform'] or ''} for m in members],
                        'counts': counts, 'tasks': tasks})
            path = os.path.join(self.root, BOARD_FILE)
            master = load_config(self.root)['master']
            old = read_board(self.root)
            if (old.get('sessions') == sessions and old.get('master') == master
                    and os.path.exists(path)):
                return          # e.g. a worker polling next: no rewrite, no watcher event
            tmp = '{}.{}.tmp'.format(path, uuid.uuid4().hex[:8])
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump({'format': BOARD_FORMAT, 'updated': self._now(), 'master': master,
                           'sessions': sessions}, f, ensure_ascii=False)
            os.replace(tmp, path)
        except Exception:
            if tmp:
                try:
                    os.remove(tmp)
                except OSError:
                    pass

    @contextlib.contextmanager
    def _db(self):
        """One transaction: commit on success, roll back on error, always close
        (sqlite3's own `with` leaves the file open, which locks it on Windows)."""
        db = sqlite3.connect(self.path, timeout=10)
        db.row_factory = sqlite3.Row
        try:
            with db:
                yield db
        finally:
            db.close()

    # ── lookups ──────────────────────────────────────────────────────────
    def _session(self, db, session, member=None, must_be_open=True):
        """Row of `session`; without an id, the one open session (the member's
        own first) so agents need not carry the id."""
        if session:
            row = db.execute('SELECT * FROM sessions WHERE id = ?',
                             (u'{}'.format(session).strip().lower(),)).fetchone()
            if row is None:
                raise HarnessError('No harness session "{}".'.format(session),
                                   'session_not_found', hint='action=status lists sessions.')
        else:
            rows = db.execute("SELECT * FROM sessions WHERE status = 'open' "
                              'ORDER BY created').fetchall()
            mine = [r for r in rows if member and db.execute(
                'SELECT 1 FROM members WHERE session = ? AND member = ?',
                (r['id'], member)).fetchone()]
            rows = mine or rows
            if len(rows) != 1:
                raise HarnessError(
                    'No open harness session.' if not rows
                    else 'Several sessions are open: pass session.',
                    'session_required', sessions=[r['id'] for r in rows],
                    hint='Lead: action=open with a goal.' if not rows else None)
            row = rows[0]
        if must_be_open and row['status'] != 'open':
            raise HarnessError('Session {} is closed.'.format(row['id']), 'session_closed')
        return row

    def _member(self, db, session_id, member):
        row = db.execute('SELECT * FROM members WHERE session = ? AND member = ?',
                         (session_id, member)).fetchone()
        if row is None:
            raise HarnessError('{} has not joined session {}.'.format(member, session_id),
                               'not_a_member', hint='action=join first.')
        db.execute('UPDATE members SET last_seen = ? WHERE session = ? AND member = ?',
                   (self._now(), session_id, member))
        return row

    def _seat(self, db, session, member, roles, what, as_owner=False):
        """(session row, member name) for an action limited to `roles`. as_owner:
        the person at the machine (MCP Control, Studio, CLI) needs no seat."""
        srow = self._session(db, session, None if as_owner else member)
        if as_owner:
            return srow, OWNER
        self._need_role(self._member(db, srow['id'], member), roles, what)
        return srow, member

    def _need_role(self, row, roles, what):
        if row['role'] not in roles:
            raise HarnessError('Only {} may {}.'.format(' / '.join(roles), what),
                               'role_denied', role=row['role'])

    def _task(self, db, session_id, key):
        row = db.execute('SELECT * FROM tasks WHERE session = ? AND key = ?',
                         (session_id, _text(key, 'key', True))).fetchone()
        if row is None:
            raise HarnessError('No task "{}".'.format(key), 'task_not_found')
        return row

    def _expire(self, db, session_id):
        """Claims whose lease ran out go back to todo."""
        db.execute("UPDATE tasks SET status = 'todo', owner = NULL, lease_until = NULL, "
                   "updated = ? WHERE session = ? AND status = 'claimed' AND lease_until < ?",
                   (self._now(), session_id, self._now()))

    @staticmethod
    def _task_view(row, full=False):
        out = {'task': row['key'], 'title': row['title'], 'status': row['status']}
        for field in ('owner', 'assignee'):
            if row[field]:
                out[field] = row[field]
        depends = json.loads(row['depends'] or '[]')
        if depends:
            out['depends_on'] = depends
        if row['summary']:
            out['summary'] = row['summary'] if full else row['summary'][:200]
        if full:
            out.update(detail=row['detail'], scope=json.loads(row['scope'] or '{}'),
                       done_when=row['done_when'], attempts=row['attempts'])
            if row['result']:
                out['result'] = json.loads(row['result'])
            for field in ('started', 'finished'):
                if row[field]:
                    out[field] = row[field]
        return out

    # ── lead ─────────────────────────────────────────────────────────────
    @_mutates
    def open(self, member, goal, document='', review=True, platform='', exclusive=None,
             write_mode=None, agent=None, as_owner=False):
        """New session; the caller becomes its lead. With a master set up
        (config.json), only that agent opens sessions. exclusive / write_mode
        default to the setup. as_owner: the person at the machine leads
        (Harness Studio), master or not."""
        member = OWNER if as_owner else _text(member, 'member', True)
        config = load_config(self.root)
        if not as_owner and config['master'] and not is_member(member, config['master']) \
                and (agent or member) != config['master']:
            raise HarnessError(
                'Only the master agent ({}) opens harness sessions on this machine.'.format(
                    config['master']), 'role_denied', master=config['master'],
                hint='Join its session: action=join, then action=next.')
        if write_mode is None:
            write_mode = config['write_mode']
        if write_mode not in WRITE_MODES:
            raise HarnessError('write_mode must be one of: {}.'.format(', '.join(WRITE_MODES)),
                               'bad_arguments')
        exclusive = config['exclusive'] if exclusive is None else bool(exclusive)
        with self._db() as db:
            sid = uuid.uuid4().hex[:8]
            now = self._now()
            db.execute('INSERT INTO sessions VALUES (?, ?, ?, ?, ?, ?, ?, NULL, ?, ?)',
                       (sid, _text(goal, 'goal', True), u'{}'.format(document or '').strip(),
                        member, 'open', 1 if review else 0, now, 1 if exclusive else 0,
                        write_mode))
            db.execute('INSERT INTO members VALUES (?, ?, ?, ?, ?, ?)',
                       (sid, member, 'lead', platform or '', now, now))
        return {'session': sid, 'role': 'lead', 'opened': True, 'exclusive': exclusive,
                'write_mode': write_mode}

    @_mutates
    def plan(self, member, tasks, session=None, as_owner=False):
        """Lead adds tasks, or rewrites ones nobody has started."""
        if not isinstance(tasks, list) or not tasks:
            raise HarnessError('"tasks" must be a non-empty list.', 'bad_arguments')
        with self._db() as db:
            srow, member = self._seat(db, session, member, ('lead',), 'plan tasks', as_owner)
            sid = srow['id']
            existing = dict((r['key'], r) for r in db.execute(
                'SELECT * FROM tasks WHERE session = ?', (sid,)))
            seq = len(existing)
            cleaned = []
            for i, task in enumerate(tasks):
                if not isinstance(task, dict):
                    raise HarnessError('tasks[{}] is not an object.'.format(i), 'bad_arguments')
                key = _text(task.get('key') or 't{}'.format(seq + len(cleaned) + 1), 'key')
                old = existing.get(key)
                if old is not None and old['status'] != 'todo':
                    raise HarnessError('Task "{}" is {}; only todo tasks can be re-planned.'
                                       .format(key, old['status']), 'task_started')
                cleaned.append((key, task))
            keys = set(existing) | set(k for k, _t in cleaned)
            if len(keys) > MAX_TASKS:
                raise HarnessError('At most {} tasks per session.'.format(MAX_TASKS),
                                   'too_many_tasks')
            for key, task in cleaned:
                depends = _str_list(task.get('depends_on'))
                unknown = [d for d in depends if d not in keys or d == key]
                if unknown:
                    raise HarnessError('Task "{}" depends on unknown task(s): {}.'.format(
                        key, ', '.join(unknown)), 'bad_dependency')
                if key not in existing:
                    seq += 1
                db.execute(
                    'INSERT OR REPLACE INTO tasks VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, '
                    'NULL, NULL, ?, NULL, NULL, ?, NULL, NULL)',
                    (sid, key, existing[key]['seq'] if key in existing else seq,
                     _text(task.get('title'), 'title', True), _text(task.get('detail'), 'detail'),
                     json.dumps(clean_scope(task.get('scope'))), json.dumps(depends),
                     _text(task.get('assignee'), 'member') or None,
                     _text(task.get('done_when'), 'done_when'), 'todo',
                     existing[key]['attempts'] if key in existing else 0, self._now()))
            self._check_cycles(db, sid)
        return dict(self.status(session=sid), planned=[k for k, _t in cleaned])

    @_mutates
    def remove(self, member, task, session=None, as_owner=False):
        """Lead deletes a task nobody has started and nothing depends on."""
        with self._db() as db:
            srow, member = self._seat(db, session, member, ('lead',), 'remove a task', as_owner)
            sid = srow['id']
            row = self._task(db, sid, task)
            if row['status'] != 'todo':
                raise HarnessError('Task "{}" is {}; only todo tasks can be removed.'.format(
                    row['key'], row['status']), 'task_started')
            users = [r['key'] for r in db.execute(
                'SELECT key, depends FROM tasks WHERE session = ?', (sid,))
                if row['key'] in json.loads(r['depends'] or '[]')]
            if users:
                raise HarnessError('Task(s) {} depend on "{}": unlink them first.'.format(
                    ', '.join(users), row['key']), 'bad_dependency', dependents=users)
            db.execute('DELETE FROM tasks WHERE session = ? AND key = ?', (sid, row['key']))
        return {'session': sid, 'task': row['key'], 'removed': True}

    @staticmethod
    def _check_cycles(db, sid):
        deps = dict((r['key'], json.loads(r['depends'] or '[]')) for r in db.execute(
            'SELECT key, depends FROM tasks WHERE session = ?', (sid,)))
        state = {}

        def visit(key):
            if state.get(key) == 1:
                raise HarnessError('Tasks depend on each other in a loop ({}).'.format(key),
                                   'bad_dependency')
            if state.get(key) == 2:
                return
            state[key] = 1
            for dep in deps.get(key, []):
                visit(dep)
            state[key] = 2

        for key in deps:
            visit(key)

    @_mutates
    def review(self, member, task, verdict, note='', session=None, as_owner=False):
        """Lead / reviewer accepts a reported task or sends it back. as_owner:
        the person at the machine (MCP Control, CLI), who needs no seat on the
        board."""
        if verdict not in VERDICTS:
            raise HarnessError('verdict must be accept or redo.', 'bad_arguments')
        with self._db() as db:
            sid = self._session(db, session, None if as_owner else member)['id']
            if as_owner:
                member = OWNER
            else:
                mrow = self._member(db, sid, member)
                self._need_role(mrow, ('lead', 'reviewer'), 'review a task')
            row = self._task(db, sid, task)
            if row['status'] != 'review':
                raise HarnessError('Task "{}" is {}, not waiting for review.'.format(
                    row['key'], row['status']), 'not_in_review')
            if not as_owner and row['owner'] == member and mrow['role'] != 'lead':
                raise HarnessError('A reviewer cannot review their own task.', 'role_denied')
            if verdict == 'accept':
                db.execute("UPDATE tasks SET status = 'done', updated = ? "
                           'WHERE session = ? AND key = ?', (self._now(), sid, row['key']))
            else:
                db.execute("UPDATE tasks SET status = 'todo', owner = NULL, lease_until = NULL, "
                           'updated = ? WHERE session = ? AND key = ?',
                           (self._now(), sid, row['key']))
            if note:
                self._add_note(db, sid, member, note, row['key'])
        return {'session': sid, 'task': row['key'],
                'status': 'done' if verdict == 'accept' else 'todo'}

    @_mutates
    def close(self, member, session=None, as_owner=False):
        with self._db() as db:
            sid = self._seat(db, session, member, ('lead',), 'close the session',
                             as_owner)[0]['id']
            db.execute("UPDATE sessions SET status = 'closed', closed = ? WHERE id = ?",
                       (self._now(), sid))
        return dict(self.status(session=sid), closed=True)

    # ── everyone ─────────────────────────────────────────────────────────
    @_mutates
    def join(self, member, role='worker', session=None, platform=''):
        member = _text(member, 'member', True)
        role = (role or 'worker').lower()
        if role not in ROLES:
            raise HarnessError('role must be one of: {}.'.format(', '.join(ROLES)),
                               'bad_arguments')
        with self._db() as db:
            srow = self._session(db, session, member)
            sid = srow['id']
            if role == 'lead' and srow['lead'] != member:
                raise HarnessError('Session {} already has a lead ({}).'.format(
                    sid, srow['lead']), 'role_denied', hint='Join as worker or reviewer.')
            now = self._now()
            db.execute('INSERT OR REPLACE INTO members VALUES (?, ?, ?, ?, '
                       'COALESCE((SELECT joined FROM members WHERE session = ? AND member = ?), ?), ?)',
                       (sid, member, role, platform or '', sid, member, now, now))
        return {'session': sid, 'role': role, 'goal': srow['goal'],
                'document': srow['document'] or None}

    @_mutates
    def next(self, member, session=None, lease_s=LEASE_S):
        """Claim the first task this member can start now, or say why none."""
        with self._db() as db:
            srow = self._session(db, session, member)
            sid = srow['id']
            mrow = self._member(db, sid, member)
            self._expire(db, sid)
            rows = db.execute('SELECT * FROM tasks WHERE session = ? ORDER BY seq',
                              (sid,)).fetchall()
            mine = [r for r in rows if r['status'] == 'claimed' and r['owner'] == member]
            if mine:                    # one task per member at a time
                return {'session': sid, 'task': self._task_view(mine[0], True),
                        'already_claimed': True}
            done = set(r['key'] for r in rows if r['status'] == 'done')
            busy = [r for r in rows if r['status'] in ('claimed', 'review')]
            waiting = []
            for row in rows:
                if row['status'] != 'todo':
                    continue
                why = self._blocked(row, mrow, done, busy, srow['document'],
                                    srow['write_mode'] == 'serial')
                if why:
                    waiting.append({'task': row['key'], 'why': why})
                    continue
                self._claim(db, sid, row['key'], member, lease_s)
                return {'session': sid, 'goal': srow['goal'],
                        'document': srow['document'] or None,
                        'task': self._task_view(self._task(db, sid, row['key']), True)}
            if rows and all(r['status'] == 'done' for r in rows):
                return {'session': sid, 'task': None, 'all_done': True}
            return {'session': sid, 'task': None, 'waiting': waiting[:10],
                    'retry_after_s': RETRY_AFTER_S}

    @staticmethod
    def _blocked(row, mrow, done, busy, document, serial=False):
        """Why `row` cannot start for this member right now ('' = it can)."""
        if row['assignee'] and row['assignee'] != mrow['member']:
            return 'assigned to {}'.format(row['assignee'])
        if not row['assignee'] and mrow['role'] != 'worker':
            return 'unassigned tasks go to workers'
        missing = [d for d in json.loads(row['depends'] or '[]') if d not in done]
        if missing:
            return 'waits for {}'.format(', '.join(missing))
        scope = json.loads(row['scope'] or '{}')
        for other in busy:
            if scopes_conflict(scope, json.loads(other['scope'] or '{}'), document or '',
                               serial):
                return '{} {} ({})'.format('one writer at a time: after' if serial
                                           else 'same scope as', other['key'], other['owner'])
        return ''

    def _claim(self, db, sid, key, member, lease_s):
        db.execute("UPDATE tasks SET status = 'claimed', owner = ?, lease_until = ?, "
                   'attempts = attempts + 1, updated = ?, started = ?, finished = NULL '
                   'WHERE session = ? AND key = ?',
                   (member, self._now() + max(60, int(lease_s or LEASE_S)), self._now(),
                    self._now(), sid, key))

    @_mutates
    def report(self, member, task, status='done', summary='', result=None, session=None):
        """Owner reports progress (renews the lease), done or failed."""
        if status not in REPORT_STATUSES:
            raise HarnessError('status must be one of: {}.'.format(', '.join(REPORT_STATUSES)),
                               'bad_arguments')
        with self._db() as db:
            srow = self._session(db, session, member)
            sid = srow['id']
            self._member(db, sid, member)
            row = self._task(db, sid, task)
            if row['status'] != 'claimed' or row['owner'] != member:
                raise HarnessError('Task "{}" is not claimed by {}.'.format(row['key'], member),
                                   'not_owner', status=row['status'], owner=row['owner'],
                                   hint='The claim may have expired: action=next.')
            summary = _text(summary, 'summary')
            if status == 'progress':
                db.execute('UPDATE tasks SET lease_until = ?, summary = ?, updated = ? '
                           'WHERE session = ? AND key = ?',
                           (self._now() + LEASE_S, summary or row['summary'], self._now(),
                            sid, row['key']))
                new = 'claimed'
            else:
                new = 'failed' if status == 'failed' else ('review' if srow['review'] else 'done')
                db.execute('UPDATE tasks SET status = ?, summary = ?, result = ?, '
                           'lease_until = NULL, updated = ?, finished = ? '
                           'WHERE session = ? AND key = ?',
                           (new, summary, json.dumps(result) if result is not None else None,
                            self._now(), self._now(), sid, row['key']))
        return {'session': sid, 'task': row['key'], 'status': new}

    @_mutates
    def retry(self, member, task, session=None, as_owner=False):
        """Lead puts a failed task back on the board."""
        with self._db() as db:
            sid = self._seat(db, session, member, ('lead',), 'retry a task', as_owner)[0]['id']
            row = self._task(db, sid, task)
            if row['status'] != 'failed':
                raise HarnessError('Task "{}" is {}, not failed.'.format(
                    row['key'], row['status']), 'bad_arguments')
            db.execute("UPDATE tasks SET status = 'todo', owner = NULL, updated = ? "
                       'WHERE session = ? AND key = ?', (self._now(), sid, row['key']))
        return {'session': sid, 'task': row['key'], 'status': 'todo'}

    def _add_note(self, db, sid, member, text, task=None):
        db.execute('INSERT INTO notes (session, task, member, text, at) VALUES (?, ?, ?, ?, ?)',
                   (sid, task, member, _text(text, 'note', True), self._now()))

    @_mutates
    def note(self, member, text, task=None, session=None, as_owner=False):
        """Shared note: findings and decisions the other agents should see."""
        with self._db() as db:
            srow, member = self._seat(db, session, member, ROLES, 'add a note', as_owner)
            sid = srow['id']
            key = self._task(db, sid, task)['key'] if task else None
            self._add_note(db, sid, member, text, key)
        return {'session': sid, 'noted': True}

    # ── gateway: writes stay inside the task ─────────────────────────────
    def write_refusal(self, agent, document='', element_ids=()):
        """Why a model-changing call by `agent` must not run, or None.

        Only what the gateway can know without Revit is checked: the document
        pin, scope.write and scope.element_ids. Categories / levels order the
        tasks on the board but are not verified per call. The session lead (a
        person's client) is not limited; a session idle for ACTIVE_S is ignored.
        """
        pin = (document or '').strip().lower()
        ids = set(u'{}'.format(i) for i in element_ids or ())
        with self._db() as db:
            for srow in db.execute("SELECT * FROM sessions WHERE status = 'open' "
                                   'ORDER BY created').fetchall():
                sid = srow['id']
                members = db.execute('SELECT * FROM members WHERE session = ?',
                                     (sid,)).fetchall()
                mine = [m for m in members if is_member(m['member'], agent)]
                if any(m['role'] == 'lead' for m in mine):
                    continue
                if not mine and not srow['exclusive']:
                    continue                    # open session: only its members are bound
                if self._now() - max(m['last_seen'] for m in members) > ACTIVE_S:
                    continue
                names = set(m['member'] for m in mine)
                sdoc = (srow['document'] or '').lower()
                tasks = [dict(r, scope=json.loads(r['scope'] or '{}')) for r in db.execute(
                    "SELECT * FROM tasks WHERE session = ? AND status IN ('claimed', 'review')",
                    (sid,)).fetchall()]
                for t in tasks:
                    t['doc'] = (t['scope'].get('document') or '').lower() or sdoc

                def same_doc(doc):
                    return not pin or not doc or doc == pin

                own = [t for t in tasks if t['status'] == 'claimed' and t['owner'] in names]
                if not own:
                    # The session's files: its own document and every task's.
                    docs = set([sdoc] + [
                        (json.loads(r['scope'] or '{}').get('document') or '').lower() or sdoc
                        for r in db.execute('SELECT scope FROM tasks WHERE session = ?',
                                            (sid,)).fetchall()])
                    if pin and '' not in docs and pin not in docs:
                        continue                # another file: not this session's business
                    if not mine:
                        return self._refusal(
                            sid, '{} is managed by harness session {} (lead: {}): changes go '
                            'through its task board.'.format(
                                '"{}"'.format(document) if document else 'This model',
                                sid, srow['lead']),
                            code='harness_managed', lead=srow['lead'], goal=srow['goal'],
                            hint='t3lab_harness action=join, then action=next for your task. '
                                 'Reading is not limited; for another file pass '
                                 'revit_document.')
                    return self._refusal(
                        sid, 'You are a {} in harness session {} and hold no task: changes to '
                        'the model go through a task.'.format(mine[0]['role'], sid),
                        hint='t3lab_harness action=next takes your next task. Work outside '
                             'the session: pass revit_document for another file, or ask the '
                             'lead to close the session.')
                here = [t for t in own if same_doc(t['doc'])]
                if not here:
                    return self._refusal(
                        sid, 'Your task {} is on "{}"; this call targets "{}".'.format(
                            own[0]['key'], own[0]['doc'], pin), task=own[0]['key'])
                writable = [t for t in here if t['scope'].get('write', True)]
                if not writable:
                    return self._refusal(
                        sid, 'Task {} is read-only (scope.write=false).'.format(here[0]['key']),
                        task=here[0]['key'],
                        hint='Report what should change; the lead plans a write task.')
                if ids:
                    fits = [t for t in writable if not t['scope'].get('element_ids')
                            or ids <= set(t['scope']['element_ids'])]
                    if not fits:
                        outside = sorted(ids - set(writable[0]['scope']['element_ids']))
                        return self._refusal(
                            sid, 'Task {} covers a fixed set of elements; {} of this call are '
                            'outside it.'.format(writable[0]['key'], len(outside)),
                            task=writable[0]['key'], outside_ids=outside[:10])
                    for other in tasks:
                        theirs = set(other['scope'].get('element_ids') or ())
                        if (other['owner'] not in names and other['scope'].get('write', True)
                                and same_doc(other['doc']) and ids & theirs):
                            return self._refusal(
                                sid, '{} holds task {} on {} of these elements.'.format(
                                    other['owner'], other['key'], len(ids & theirs)),
                                code='scope_claimed', task=other['key'],
                                claimed_ids=sorted(ids & theirs)[:10], retryable=True)
        return None

    @staticmethod
    def _refusal(sid, message, code='scope_denied', **extra):
        extra = dict((k, v) for k, v in extra.items() if v is not None)
        return dict({'error': message, 'error_code': code, 'retryable': False,
                     'nothing_ran': True, 'session': sid}, **extra)

    def stats(self, session=None, member=None):
        """Numbers of one session for comparing runs (docs/agents/harness-vs-single.md):
        wall_s = first claim to last report; work_s = sum of task times;
        parallelism = work_s / wall_s."""
        with self._db() as db:
            srow = self._session(db, session, member, must_be_open=False)
            rows = db.execute('SELECT * FROM tasks WHERE session = ? ORDER BY seq',
                              (srow['id'],)).fetchall()
        counts, by_member, work = {}, {}, 0.0
        for row in rows:
            counts[row['status']] = counts.get(row['status'], 0) + 1
            if row['owner'] and row['started'] and row['finished']:
                seconds = max(0.0, row['finished'] - row['started'])
                work += seconds
                who = by_member.setdefault(row['owner'], {'tasks': 0, 'work_s': 0.0})
                who['tasks'] += 1
                who['work_s'] = round(who['work_s'] + seconds, 1)
        started = [r['started'] for r in rows if r['started']]
        finished = [r['finished'] for r in rows if r['finished']]
        wall = max(finished) - min(started) if started and finished else 0.0
        return {'session': srow['id'], 'status': srow['status'], 'goal': srow['goal'],
                'tasks': len(rows), 'counts': counts,
                'attempts': sum(r['attempts'] or 0 for r in rows),
                'extra_attempts': sum(max(0, (r['attempts'] or 0) - 1) for r in rows),
                'started_at': min(started) if started else None,
                'finished_at': max(finished) if finished else None,
                'wall_s': round(wall, 1), 'work_s': round(work, 1),
                'parallelism': round(work / wall, 2) if wall > 0 else None,
                'by_member': by_member}

    def status(self, session=None, member=None, task=None, full=False):
        """The board. Without any open session: the list of sessions. full:
        every task with detail, scope and times (Harness Studio)."""
        with self._db() as db:
            try:
                srow = self._session(db, session, member, must_be_open=False)
            except HarnessError as ex:
                if session or ex.payload['error_code'] != 'session_required':
                    raise
                rows = db.execute('SELECT * FROM sessions ORDER BY created DESC LIMIT 20')
                return {'sessions': [{'session': r['id'], 'goal': r['goal'][:120],
                                      'status': r['status'], 'lead': r['lead']} for r in rows]}
            sid = srow['id']
            self._expire(db, sid)
            if task:
                return {'session': sid,
                        'task': self._task_view(self._task(db, sid, task), True),
                        'notes': self._notes(db, sid, task)}
            tasks = db.execute('SELECT * FROM tasks WHERE session = ? ORDER BY seq',
                               (sid,)).fetchall()
            counts = {}
            for row in tasks:
                counts[row['status']] = counts.get(row['status'], 0) + 1
            return {'session': sid, 'status': srow['status'], 'goal': srow['goal'],
                    'document': srow['document'] or None, 'lead': srow['lead'],
                    'review_required': bool(srow['review']),
                    'exclusive': bool(srow['exclusive']),
                    'write_mode': srow['write_mode'] or WRITE_MODES[0],
                    'members': [{'member': m['member'], 'role': m['role'],
                                 'platform': m['platform'] or None,
                                 'idle_s': int(self._now() - m['last_seen'])}
                                for m in db.execute(
                                    'SELECT * FROM members WHERE session = ? ORDER BY joined',
                                    (sid,))],
                    'counts': counts, 'tasks': [self._task_view(r, full) for r in tasks],
                    'notes': self._notes(db, sid)}

    def sessions(self, limit=20):
        """Newest sessions first, open and closed."""
        with self._db() as db:
            return [{'session': r['id'], 'goal': r['goal'][:120], 'status': r['status'],
                     'lead': r['lead'], 'document': r['document'] or ''}
                    for r in db.execute('SELECT * FROM sessions ORDER BY created DESC LIMIT ?',
                                        (int(limit),))]

    @staticmethod
    def _notes(db, sid, task=None):
        sql = 'SELECT * FROM notes WHERE session = ?'
        args = [sid]
        if task:
            sql += ' AND task = ?'
            args.append(task)
        rows = db.execute(sql + ' ORDER BY id DESC LIMIT ?', args + [MAX_NOTES_SHOWN]).fetchall()
        return [dict({'by': r['member'], 'text': r['text']},
                     **({'task': r['task']} if r['task'] else {})) for r in reversed(rows)]
