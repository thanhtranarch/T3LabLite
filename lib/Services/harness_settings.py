# -*- coding: utf-8 -*-
"""
Owner-side BIM Harness helpers for MCP Control, the Harness ribbon button and
the Assistant: read the board.json snapshot the bridge writes, review a task as
the person at this machine, open Harness Studio (docs/harness.md). The board's
database is only ever touched by the bridge process (core/harness_store.py).

Pure Python, CPython 3 + IronPython 2.7, no Revit imports.
Test: python3 lite_guard/test_harness_settings.py
"""
from __future__ import unicode_literals

import io
import json
import os
import time

HARNESS_DIR = 'harness'                 # = core/harness_store.harness_dir()
BOARD_FILE = 'board.json'               # = core/harness_store.BOARD_FILE
# What the bridge would otherwise read from Revit's environment: a person
# clicking in Revit is always the 'cli' client.
_CHILD_ENV_DROP = ('PYTHONPATH', 'PYTHONHOME', 'T3LAB_AGENT')
_CREATE_NO_WINDOW = 0x08000000
STUDIO_IDLE_EXIT_MIN = 10       # a Studio started from Revit has no terminal to stop it

_HARNESS_ORDER = ('review', 'failed', 'claimed', 'todo', 'done')
HARNESS_STATUS_TEXT = {'review': 'Waiting for review', 'failed': 'Failed',
                       'claimed': 'Working', 'todo': 'To do', 'done': 'Done'}


def data_dir():
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI')


def harness_dir():
    return os.environ.get('T3LAB_HARNESS_DIR') or os.path.join(data_dir(), HARNESS_DIR)


def _read_json(path):
    try:
        with io.open(path, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return None


def _plural(n, word):
    return '{} {}{}'.format(n, word, '' if n == 1 else 's')


def read_board():
    board = _read_json(os.path.join(harness_dir(), BOARD_FILE))
    return board if isinstance(board, dict) else {}


def list_harness_sessions(board=None):
    """Open harness sessions from board.json; [] when there is no board."""
    board = read_board() if board is None else board
    return [s for s in board.get('sessions') or [] if isinstance(s, dict)]


def harness_rows(sessions=None):
    """One dict per task of every open session, the ones that need a person
    first: {session, task, title, status, status_text, owner, document}."""
    rows = []
    for session in (list_harness_sessions() if sessions is None else sessions):
        for task in session.get('tasks') or []:
            status = task.get('status') or ''
            rows.append({
                'session': session.get('session') or '', 'task': task.get('task') or '',
                'title': task.get('title') or '', 'status': status,
                'status_text': HARNESS_STATUS_TEXT.get(status, status),
                'owner': task.get('owner') or task.get('assignee') or '',
                'document': task.get('document') or session.get('document') or ''})
    rows.sort(key=lambda r: _HARNESS_ORDER.index(r['status'])
              if r['status'] in _HARNESS_ORDER else len(_HARNESS_ORDER))
    return rows


def harness_master(board=None):
    """Agent id the owner named master (bridge.py harness setup), or ''."""
    board = read_board() if board is None else board
    return board.get('master') or ''


def harness_summary(sessions=None, master=None):
    """One sentence about the open sessions, with counts; names the master
    agent when one is set up."""
    sessions = list_harness_sessions() if sessions is None else sessions
    master = harness_master() if master is None else master
    lead = 'Master: {}. '.format(master) if master else ''
    if not sessions:
        return lead + 'No harness session is open.'
    parts = []
    for session in sessions:
        counts = session.get('counts') or {}
        total = len(session.get('tasks') or [])
        bits = ['{} working'.format(counts['claimed'])] if counts.get('claimed') else []
        if counts.get('review'):
            bits.append('{} waiting for review'.format(counts['review']))
        if counts.get('failed'):
            bits.append('{} failed'.format(counts['failed']))
        bits.append('{} of {} done'.format(counts.get('done', 0), total))
        if session.get('write_mode') == 'serial':
            bits.append('one writer at a time')
        parts.append('{} ({}): {}'.format(
            session.get('goal') or session.get('session'),
            _plural(len(session.get('members') or []), 'agent'), ', '.join(bits)))
    return lead + ' · '.join(parts)


def describe_harness_result(verdict, code, payload, row=None):
    """Status-line sentence for a harness review answer."""
    title = (row or {}).get('title') or (row or {}).get('task') or 'the task'
    if isinstance(payload, dict) and payload.get('error'):
        if payload.get('error_code') == 'not_in_review':
            return '"{}" is no longer waiting for review. Click Refresh.'.format(title)
        return 'Could not review "{}": {}'.format(title, payload['error'])
    if code != 0:
        first = ('{}'.format(payload).strip().splitlines() or [''])[0][:160]
        return 'Review failed (bridge {}): {}'.format(
            'did not answer' if code is None else 'exit {}'.format(code), first or 'no output')
    if verdict == 'accept':
        return 'Accepted "{}". Tasks that waited for it can start.'.format(title)
    return 'Sent "{}" back. A worker takes it again.'.format(title)


# ─── Bridge CLI ──────────────────────────────────────────────────────────────

def _parse_output(text):
    text = (text or '').strip()
    try:
        return json.loads(text)
    except ValueError:
        return text


def _read_text(path):
    try:
        with io.open(path, encoding='utf-8', errors='replace') as f:
            return f.read()
    except Exception:
        return ''


def _bridge_command(python, bridge):
    if python is None or bridge is None:
        from Services import mcp_service
        python = python or mcp_service._find_python_executable()
        bridge = bridge or mcp_service._get_bridge_path()
    return python, bridge


def _child_env():
    env = dict((k, v) for k, v in os.environ.items() if k not in _CHILD_ENV_DROP)
    env['PYTHONIOENCODING'] = 'utf-8'
    return env


class BridgeJob(object):
    """One `python bridge.py <args>` run. Output goes to temp files, not pipes, so the
    child never blocks on a full pipe when nobody polls (window closed mid-run)."""

    def __init__(self, args, python=None, bridge=None):
        import subprocess
        import tempfile
        python, bridge = _bridge_command(python, bridge)
        fd_out, self._out = tempfile.mkstemp(prefix='t3lab_bridge_', suffix='.out')
        fd_err, self._err = tempfile.mkstemp(prefix='t3lab_bridge_', suffix='.err')
        kwargs = {}
        if os.name == 'nt':
            kwargs['creationflags'] = _CREATE_NO_WINDOW
        try:
            with os.fdopen(fd_out, 'wb') as out, os.fdopen(fd_err, 'wb') as err:
                self._proc = subprocess.Popen([python, bridge] + [str(a) for a in args],
                                              stdout=out, stderr=err, env=_child_env(),
                                              **kwargs)
        except Exception:
            self._cleanup()
            raise
        self._result = None

    def _cleanup(self):
        for path in (self._out, self._err):
            try:
                os.remove(path)
            except OSError:
                pass

    def poll(self):
        """None while running, else (exit_code, parsed JSON or text)."""
        if self._result is None:
            code = self._proc.poll()
            if code is None:
                return None
            out = _read_text(self._out)
            self._result = (code, _parse_output(out if out.strip() else _read_text(self._err)))
            self._cleanup()
        return self._result

    def wait(self, timeout=None):
        """poll() until done; None if still running after `timeout` seconds (child keeps running)."""
        deadline = None if timeout is None else time.time() + timeout
        while True:
            result = self.poll()
            if result is not None or (deadline is not None and time.time() >= deadline):
                return result
            time.sleep(0.05)


def start_bridge(args, python=None, bridge=None):
    return BridgeJob(args, python=python, bridge=bridge)


def run_bridge(args, python=None, bridge=None, timeout=None):
    """(exit_code, parsed JSON or text); exit_code None if `timeout` ran out."""
    result = start_bridge(args, python, bridge).wait(timeout)
    return result if result is not None else (None, 'The bridge is still running.')


def review_args(session, task, verdict, note=''):
    """`bridge.py harness review` arguments: the person at this machine
    (owner=true, accepted from the CLI only), no seat on the board needed."""
    args = {'task': task, 'verdict': verdict, 'owner': True}
    if note:
        args['note'] = note
    return ['harness', 'review', '--session={}'.format(session), '--json', json.dumps(args)]


def harness_review(session, task, verdict, note='', python=None, bridge=None, timeout=20):
    return run_bridge(review_args(session, task, verdict, note), python, bridge, timeout)


def open_studio(python=None, bridge=None, popen=None):
    """Start Harness Studio detached; it opens the browser itself, or the page
    of the Studio already running (harness/studio.json)."""
    import subprocess
    python, bridge = _bridge_command(python, bridge)
    kwargs = {'creationflags': _CREATE_NO_WINDOW} if os.name == 'nt' else {}
    with io.open(os.devnull, 'wb') as sink:
        return (popen or subprocess.Popen)(
            [python, bridge, 'harness', 'studio',
             '--idle-exit={}'.format(STUDIO_IDLE_EXIT_MIN)],
            stdout=sink, stderr=sink, env=_child_env(), **kwargs)
