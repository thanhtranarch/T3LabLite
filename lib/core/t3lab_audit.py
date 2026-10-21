# -*- coding: utf-8 -*-
"""
MCP audit log — one JSONL line per tool call, on this machine only.

Design: "T3Lab MCP — Approval Gate & Audit Log" (2026-10-14). Every tool
call, external (bridge.py / Claude Desktop) or in-app assistant, passes
T3LabAIServer._execute_tool; that chokepoint calls begin() / running() /
end() here, so the log answers "which AI changed what, when, on which model".

    %LOCALAPPDATA%\\T3Lab\\audit\\audit-YYYY-MM-DD.jsonl    (one record per call)
    %LOCALAPPDATA%\\T3Lab\\audit\\code\\<sha>.py             (send_code_to_revit)
    %LOCALAPPDATA%\\T3Lab\\audit\\running-<pid>-<id>.json    (crash marker)

LOCALAPPDATA on purpose: never synced by OneDrive. Nothing leaves the machine.

Rules this module keeps:
  * Never fails a tool call — every public function swallows its own errors.
  * No log I/O on Revit's UI thread: records go through a bounded queue
    (5,000, newest dropped when full) to one writer thread. The crash marker
    is the one exception — a ~300-byte file written right before the tool
    runs, because a crash must find it on disk, not in a queue.
  * Runs on IronPython 2.7 and CPython 3 (jsonsafe for the IronPython
    escaper bug, io.open for UTF-8).

Change capture: subscribe() hooks Application.DocumentChanged once per server
object; while a call is executing on the main thread (running()), the handler
counts added / modified / deleted element ids and keeps 20 samples.

Crash marker: running() writes running-<pid>-<call_id>.json before the tool
and deletes it after. recover_crashed_runs() (server start) turns markers of
dead processes into `crash_recovered` records — the call that took Revit down.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__ = "MCP Audit Log"

import hashlib
import io
import json
import os
import sys
import threading
import time
import uuid

try:
    import queue as _queue_mod            # CPython 3
except ImportError:                       # pragma: no cover - IronPython 2.7
    import Queue as _queue_mod

try:
    from core import jsonsafe as _jsonsafe
except Exception:                         # pragma: no cover - loaded standalone
    _jsonsafe = None

QUEUE_MAX = 5000
ROTATE_BYTES = 20 * 1024 * 1024
RETENTION_DAYS = 30
ARGS_PREVIEW_CHARS = 2000
ERROR_CHARS = 1000
IDS_SAMPLE = 20
CODE_TOOL = 'send_code_to_revit'

_lock = threading.Lock()
_queue = _queue_mod.Queue(QUEUE_MAX)
_writer = None
_dropped = [0]
# The call executing on Revit's main thread right now (set by running()).
_current = [None]
# Identity of THIS load of the module: the server re-subscribes
# DocumentChanged when a pyRevit reload brought a new copy (see subscribe).
MODULE_TOKEN = object()
_housekeeping_started = [False]


# ── paths ──────────────────────────────────────────────────────────────────

def audit_dir():
    """%LOCALAPPDATA%\\T3Lab\\audit, created on first use.

    T3LAB_AUDIT_DIR overrides it (tests, or a machine that wants it elsewhere).
    """
    override = os.environ.get('T3LAB_AUDIT_DIR')
    if override:
        d = override
    else:
        base = os.environ.get('LOCALAPPDATA') or os.path.join(
            os.path.expanduser('~'), '.t3lab')
        d = os.path.join(base, 'T3Lab', 'audit')
    try:
        if not os.path.isdir(d):
            os.makedirs(d)
    except OSError:
        pass
    return d


def _day_file(day, folder=None):
    """Today's file, or the next .N file once it passed ROTATE_BYTES."""
    folder = folder or audit_dir()
    path = os.path.join(folder, 'audit-{}.jsonl'.format(day))
    n = 0
    while True:
        try:
            if not os.path.isfile(path) or os.path.getsize(path) < ROTATE_BYTES:
                return path
        except OSError:
            return path
        n += 1
        path = os.path.join(folder, 'audit-{}.{}.jsonl'.format(day, n))


# ── small helpers ──────────────────────────────────────────────────────────

def _dumps(obj, **kw):
    if _jsonsafe is not None:
        return _jsonsafe.dumps(obj, **kw)
    kw['ensure_ascii'] = True
    return json.dumps(obj, **kw)


def _text(value):
    try:
        return u'{}'.format(value)
    except Exception:
        try:
            return repr(value)
        except Exception:
            return u'<unprintable>'


def _sha(text, n=None):
    if not isinstance(text, bytes):
        text = _text(text).encode('utf-8')
    h = hashlib.sha256(text).hexdigest()
    return h[:n] if n else h


def _now_iso():
    t = time.time()
    return time.strftime('%Y-%m-%dT%H:%M:%S', time.gmtime(t)) + \
        '.{:03d}Z'.format(int((t % 1) * 1000))


def _args_json(arguments):
    """Readable JSON of the arguments (accents kept; the record itself is
    escaped by jsonsafe when written). ensure_ascii=False is the path that
    is sound on IronPython 2.7."""
    try:
        obj = arguments if arguments is not None else {}
        if _jsonsafe is not None:
            obj = _jsonsafe.coerce(obj)
        return json.dumps(obj, sort_keys=True, ensure_ascii=False, default=_text)
    except Exception:
        return _text(arguments)


def _pid():
    try:
        return os.getpid()
    except Exception:
        return 0


# ── the call record ────────────────────────────────────────────────────────

class AuditCall(object):
    """One tool call. Built by begin(), filled in by the server, sent by end()."""

    def __init__(self, tool, arguments, source, port=None, engine=None):
        self.call_id = uuid.uuid4().hex[:12]
        self.tool = tool
        self.arguments = arguments
        self.source = source
        self.port = port
        self.engine = engine
        self.ts = _now_iso()
        self.decision = 'allow'
        self.reason = None
        self.approved_by = 'n/a'
        self.wait_ms = 0
        self.exec_ms = None
        self.doc_title = None
        self.doc_path_sha = None
        self.added = 0
        self.modified = 0
        self.deleted = 0
        self.ids_sample = []
        self.agent = None
        self.session = None
        self.ended = False
        self._exec_t0 = None

    def set_decision(self, decision, reason=None, approved_by='n/a', wait_ms=0):
        self.decision = decision
        self.reason = reason
        self.approved_by = approved_by
        try:
            self.wait_ms = int(wait_ms or 0)
        except Exception:
            self.wait_ms = 0

    def _note_ids(self, kind, ids):
        """DocumentChanged handler payload: count + keep a few samples."""
        n = len(ids)
        if kind == 'added':
            self.added += n
        elif kind == 'modified':
            self.modified += n
        else:
            self.deleted += n
        room = IDS_SAMPLE - len(self.ids_sample)
        if room > 0:
            self.ids_sample.extend(ids[:room])


def begin(tool, arguments, source, port=None, engine=None):
    """Open the record of one call. Never raises (returns None on failure)."""
    try:
        return AuditCall(tool, arguments, source, port, engine)
    except Exception:
        return None


def _status_of(result):
    if isinstance(result, dict):
        if result.get('denied'):
            return 'denied'
        if 'accepted_arguments' in result and result.get('error'):
            return 'rejected'          # _reject_unknown_arguments
        if result.get('error'):
            return 'error'
    return 'ok'


def _result_chars(result):
    try:
        if isinstance(result, (dict, list, tuple)):
            return len(_dumps(result, default=_text))
        return len(_text(result)) if result is not None else 0
    except Exception:
        return None


def _record(call, result, status):
    args_json = _args_json(call.arguments)
    error = None
    if isinstance(result, dict) and result.get('error'):
        error = _text(result.get('error'))[:ERROR_CHARS]
    op = None
    if isinstance(call.arguments, dict):
        op = call.arguments.get('operation') or call.arguments.get('action')
    rec = {
        'ts': call.ts,
        'call_id': call.call_id,
        'event': 'call',
        'source': call.source,
        'agent': getattr(call, 'agent', None),
        'session': getattr(call, 'session', None),
        'revit_pid': _pid(),
        'revit_port': call.port,
        'engine': call.engine,
        'doc_title': call.doc_title,
        'doc_path_sha': call.doc_path_sha,
        'tool': call.tool,
        'op': _text(op) if op is not None else None,
        'args_sha256': _sha(args_json),
        'args_preview': args_json[:ARGS_PREVIEW_CHARS],
        'decision': call.decision,
        'reason': call.reason,
        'approved_by': call.approved_by,
        'wait_ms': call.wait_ms,
        'exec_ms': call.exec_ms,
        'status': status,
        'error': error,
        # send_code_to_revit failures carry core.code_errors' category.
        'error_category': (result.get('category')
                           if error and isinstance(result, dict) else None),
        'changes': {'added': call.added, 'modified': call.modified,
                    'deleted': call.deleted},
        'ids_sample': list(call.ids_sample),
        'result_chars': _result_chars(result),
    }
    # Tools with a before/after diff (core.server._DIFF_TOOLS).
    if isinstance(result, dict) and 'diff' in result:
        rec['dry_run'] = bool(result.get('dry_run'))
        rec['applied'] = bool(result.get('applied'))
        rec['diff_total'] = result.get('diff_total')
        rec['diff_sample'] = list(result.get('diff') or [])[:IDS_SAMPLE]
    code = None
    if call.tool == CODE_TOOL and isinstance(call.arguments, dict):
        code = call.arguments.get('code')
        if code:
            rec['code_sha256'] = _sha(code)
    return rec, code


def end(call, result, status=None):
    """Close the record and queue it for the writer. Idempotent, never raises."""
    try:
        if call is None or call.ended:
            return
        call.ended = True
        rec, code = _record(call, result, status or _status_of(result))
        _enqueue(rec, code)
    except Exception:
        pass


# ── execution on the main thread ───────────────────────────────────────────

class running(object):
    """`with audit.running(call, uiapp):` around the tool's real execution.

    Stamps the document, times the call, routes DocumentChanged into it and
    keeps a crash marker on disk while the tool runs. A None call is a no-op.
    """

    def __init__(self, call, uiapp=None):
        self.call = call
        self.uiapp = uiapp
        self.marker = None

    def __enter__(self):
        call = self.call
        if call is None:
            return call
        try:
            _stamp_document(call, self.uiapp)
        except Exception:
            pass
        self.marker = _write_marker(call)
        call._exec_t0 = time.time()
        _current[0] = call
        return call

    def __exit__(self, exc_type, exc, tb):
        call = self.call
        if call is not None:
            if _current[0] is call:
                _current[0] = None
            try:
                call.exec_ms = int((time.time() - (call._exec_t0 or time.time())) * 1000)
            except Exception:
                pass
            _remove(self.marker)
        return False


def _stamp_document(call, uiapp):
    if uiapp is None:
        return
    uidoc = getattr(uiapp, 'ActiveUIDocument', None)
    doc = getattr(uidoc, 'Document', None) if uidoc is not None else None
    if doc is None:
        return
    call.doc_title = _text(getattr(doc, 'Title', u'') or u'')
    path = _text(getattr(doc, 'PathName', u'') or u'')
    call.doc_path_sha = _sha(path.lower(), 12) if path else None


def _write_marker(call):
    try:
        path = os.path.join(audit_dir(), 'running-{}-{}.json'.format(
            _pid(), call.call_id))
        body = _dumps({'call_id': call.call_id, 'tool': call.tool,
                       'source': call.source, 'doc_title': call.doc_title,
                       'doc_path_sha': call.doc_path_sha,
                       'started': _now_iso(), 'revit_pid': _pid(),
                       'engine': call.engine, 'revit_port': call.port,
                       'args_preview': _args_json(call.arguments)[:ARGS_PREVIEW_CHARS]})
        with io.open(path, 'w', encoding='utf-8') as f:
            f.write(body)
        return path
    except Exception:
        return None


def _remove(path):
    if not path:
        return
    try:
        os.remove(path)
    except Exception:
        pass


# ── change capture (Application.DocumentChanged) ──────────────────────────

def _ids(collection):
    try:
        from Snippets._compat import eid_value
    except Exception:
        eid_value = None
    out = []
    for eid in collection:
        try:
            out.append(int(eid_value(eid)) if eid_value else int(eid.Value))
        except Exception:
            try:
                out.append(int(eid.IntegerValue))   # Revit < 2024
            except Exception:
                pass
    return out


def _on_document_changed(sender, args):
    """Count ids into the call running right now. Cheap; never raises."""
    call = _current[0]
    if call is None:
        return
    try:
        call._note_ids('added', _ids(args.GetAddedElementIds()))
        call._note_ids('modified', _ids(args.GetModifiedElementIds()))
        call._note_ids('deleted', _ids(args.GetDeletedElementIds()))
    except Exception:
        pass


_HANDLER_ATTR = '_t3lab_audit_doc_changed'


def subscribe(application, owner):
    """Hook DocumentChanged once per `owner` (the server object).

    The delegate lives on `owner`, which is anchored on the AppDomain, so a
    pyRevit reload finds the previous one and detaches it before attaching
    the new one — no double counting. Must run where the Revit API is valid
    (the main thread). Returns True when subscribed.
    """
    if application is None or owner is None:
        return False
    try:
        from System import EventHandler
        from Autodesk.Revit.DB.Events import DocumentChangedEventArgs
    except Exception:
        return False
    old = getattr(owner, _HANDLER_ATTR, None)
    if old is not None:
        try:
            application.DocumentChanged -= old
        except Exception:
            pass
    try:
        delegate = EventHandler[DocumentChangedEventArgs](_on_document_changed)
        application.DocumentChanged += delegate
        setattr(owner, _HANDLER_ATTR, delegate)
        return True
    except Exception:
        return False


# ── writer thread ──────────────────────────────────────────────────────────

def _enqueue(rec, code=None):
    _ensure_writer()
    try:
        _queue.put_nowait((rec, code))
    except _queue_mod.Full:
        _dropped[0] += 1


def _ensure_writer():
    global _writer
    if _writer is not None and _writer.is_alive():
        return
    with _lock:
        if _writer is not None and _writer.is_alive():
            return
        t = threading.Thread(target=_writer_loop, name='T3LabAuditWriter')
        t.daemon = True
        t.start()
        _writer = t


def _writer_loop():
    while True:
        try:
            item = _queue.get()
        except Exception:
            continue
        batch = [item]
        while len(batch) < 200:
            try:
                batch.append(_queue.get_nowait())
            except _queue_mod.Empty:
                break
        try:
            _write_batch(batch)
        except Exception:
            pass
        finally:
            for _ in batch:
                try:
                    _queue.task_done()
                except Exception:
                    pass


def _write_batch(batch):
    folder = audit_dir()
    lines = []
    for rec, code in batch:
        if code:
            _save_code(folder, rec.get('code_sha256') or _sha(code), code)
        try:
            lines.append(_dumps(rec, default=_text) + u'\n')
        except Exception:
            pass
    if _dropped[0]:
        lines.append(_dumps({'ts': _now_iso(), 'event': 'dropped',
                             'count': _dropped[0], 'revit_pid': _pid()}) + u'\n')
        _dropped[0] = 0
    if not lines:
        return
    path = _day_file(time.strftime('%Y-%m-%d', time.gmtime()), folder)
    # One write per batch, append mode: two engines (IronPython server,
    # CPython assistant) may append to the same day file.
    with io.open(path, 'a', encoding='utf-8') as f:
        f.write(u''.join(lines))


def _save_code(folder, sha, code):
    try:
        code_dir = os.path.join(folder, 'code')
        if not os.path.isdir(code_dir):
            os.makedirs(code_dir)
        path = os.path.join(code_dir, '{}.py'.format(sha))
        if os.path.isfile(path):
            return
        with io.open(path, 'w', encoding='utf-8') as f:
            f.write(_text(code))
    except Exception:
        pass


def flush(timeout=5.0):
    """Wait (bounded) until the queue is written. For tests and shutdown."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _queue.unfinished_tasks == 0:
            return True
        time.sleep(0.02)
    return False


# ── crash recovery & housekeeping ─────────────────────────────────────────

def _process_alive_as_revit(pid):
    """False only when we KNOW `pid` is gone or is not Revit.

    Without System.Diagnostics (dev runs off Windows) the answer is True —
    a marker is never declared crashed on a guess.
    """
    try:
        from System.Diagnostics import Process
    except Exception:
        return True
    try:
        proc = Process.GetProcessById(int(pid))
    except Exception:
        return False                       # ArgumentException: not running
    try:
        if proc.HasExited:
            return False
        return 'revit' in _text(proc.ProcessName).lower()
    except Exception:
        return True


def recover_crashed_runs():
    """Turn markers of dead Revit processes into `crash_recovered` records.

    Returns how many were recovered. Markers of this process, or of a Revit
    that is still alive, are left alone.
    """
    folder = audit_dir()
    me = _pid()
    count = 0
    try:
        names = os.listdir(folder)
    except Exception:
        return 0
    for name in names:
        if not (name.startswith('running-') and name.endswith('.json')):
            continue
        parts = name[:-5].split('-')
        if len(parts) < 3:
            continue
        try:
            pid = int(parts[1])
        except ValueError:
            continue
        if pid == me or _process_alive_as_revit(pid):
            continue
        path = os.path.join(folder, name)
        try:
            with io.open(path, 'r', encoding='utf-8') as f:
                marker = json.loads(f.read() or u'{}')
        except Exception:
            marker = {}
        rec = {'ts': _now_iso(), 'event': 'crash_recovered',
               'call_id': marker.get('call_id') or '-'.join(parts[2:]),
               'crashed_pid': pid, 'revit_pid': me,
               'tool': marker.get('tool'), 'source': marker.get('source'),
               'doc_title': marker.get('doc_title'),
               'doc_path_sha': marker.get('doc_path_sha'),
               'started': marker.get('started'),
               'engine': marker.get('engine'),
               'args_preview': marker.get('args_preview'),
               'status': 'crashed'}
        _enqueue(rec)
        _remove(path)
        count += 1
    return count


def cleanup_old_logs(days=RETENTION_DAYS):
    """Delete day files and saved code older than `days`. Returns the count."""
    folder = audit_dir()
    cutoff = time.time() - days * 86400
    removed = 0
    targets = []
    try:
        targets = [os.path.join(folder, n) for n in os.listdir(folder)
                   if n.startswith('audit-') and n.endswith('.jsonl')]
    except Exception:
        pass
    code_dir = os.path.join(folder, 'code')
    try:
        targets += [os.path.join(code_dir, n) for n in os.listdir(code_dir)
                    if n.endswith('.py')]
    except Exception:
        pass
    for path in targets:
        try:
            if os.path.getmtime(path) < cutoff:
                os.remove(path)
                removed += 1
        except Exception:
            pass
    return removed


def start_housekeeping():
    """recover_crashed_runs() + cleanup_old_logs() once per module load, on a
    background thread (both touch the disk; the caller is the main thread)."""
    with _lock:
        if _housekeeping_started[0]:
            return
        _housekeeping_started[0] = True

    def work():
        try:
            recover_crashed_runs()
        except Exception:
            pass
        try:
            cleanup_old_logs()
        except Exception:
            pass

    try:
        t = threading.Thread(target=work, name='T3LabAuditHousekeeping')
        t.daemon = True
        t.start()
    except Exception:
        pass


def query(date=None, tool=None, status=None, limit=100):
    """Read records back (newest last). For an audit_query tool and tests."""
    day = date or time.strftime('%Y-%m-%d', time.gmtime())
    folder = audit_dir()
    out = []
    try:
        names = sorted(n for n in os.listdir(folder)
                       if n.startswith('audit-{}'.format(day)) and n.endswith('.jsonl'))
    except Exception:
        return out
    for name in names:
        try:
            with io.open(os.path.join(folder, name), 'r', encoding='utf-8') as f:
                for line in f:
                    try:
                        rec = json.loads(line)
                    except Exception:
                        continue
                    if tool and rec.get('tool') != tool:
                        continue
                    if status and rec.get('status') != status:
                        continue
                    out.append(rec)
        except Exception:
            continue
    return out[-limit:] if limit else out
