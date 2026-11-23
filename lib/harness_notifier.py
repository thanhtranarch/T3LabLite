# -*- coding: utf-8 -*-
"""Windows toast when the BIM Harness board needs the person (docs/harness.md).

The bridge rewrites %APPDATA%\\T3LabAI\\harness\\board.json after every board
change. startup.py -> install() watches it with a FileSystemWatcher for the
whole Revit session: a task that starts waiting for review, a task that failed,
and a session whose tasks are all done raise a pyRevit toast.

Board part of t3lab_dev's lib/approval_notifier.py (Lite has no approval queue);
same AppDomain key, so one watcher per Revit even with both extensions loaded.
IronPython (startup.py, watcher threads are fine there: no GIL) and CPython
(tests). Never raises into startup. Opt-out: env T3LAB_NO_HARNESS_TOAST=1 or file
%APPDATA%\\T3LabAI\\harness_toast.disabled (t3lab_dev's approval-toast switch too).
Test: python lite_guard/test_harness_notifier.py
"""
from __future__ import unicode_literals

import io
import json
import os
import threading

HARNESS_DIR = 'harness'                 # = core/harness_store.harness_dir()
BOARD_FILE = 'board.json'               # = core/harness_store.BOARD_FILE
# AppDomain lives across pyRevit reloads: one watcher per Revit session.
HARNESS_WATCHER_KEY = 'T3Lab.HarnessNotifier.Watcher'
OPT_OUT_FILES = ('harness_toast.disabled', 'approval_toast.disabled')
OPT_OUT_ENV = ('T3LAB_NO_HARNESS_TOAST', 'T3LAB_NO_APPROVAL_TOAST')
TITLE = 'T3Lab Harness'
DEBOUNCE_S = 0.4            # the bridge writes tmp + replace: several events per save

INSTALLED, SKIPPED, FAILED = 'installed', 'skipped', 'failed'


def _data_dir():
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI')


def harness_dir():
    return os.environ.get('T3LAB_HARNESS_DIR') or os.path.join(_data_dir(), HARNESS_DIR)


def opted_out():
    return (any(os.environ.get(k, '').strip() in ('1', 'true') for k in OPT_OUT_ENV)
            or any(os.path.exists(os.path.join(_data_dir(), f)) for f in OPT_OUT_FILES))


def _plural(n, word):
    return '{} {}{}'.format(n, word, '' if n == 1 else 's')


def describe_board(board, seen):
    """(toasts, statuses now) for a harness board snapshot.

    seen = {'<session>/<task>': status} from the previous snapshot, None on
    the first look (nothing is announced then: no replay after a restart)."""
    now, messages = {}, []
    for session in (board or {}).get('sessions') or []:
        sid = session.get('session') or ''
        tasks = session.get('tasks') or []
        changed = False
        for task in tasks:
            key = '{}/{}'.format(sid, task.get('task'))
            status = task.get('status') or ''
            now[key] = status
            if seen is None or seen.get(key) == status:
                continue
            changed = True
            who = task.get('owner') or 'An agent'
            title = task.get('title') or task.get('task') or 'a task'
            if status == 'review':
                messages.append('{} finished "{}". It waits for review: MCP Control > '
                                'Harness, or ask the lead.'.format(who, title))
            elif status == 'failed':
                messages.append('{} could not finish "{}". The lead can retry it.'.format(
                    who, title))
        if changed and tasks and all(t.get('status') == 'done' for t in tasks):
            messages.append('Harness: all {} of "{}" are done.'.format(
                _plural(len(tasks), 'task'), session.get('goal') or sid))
    return messages, now


class BoardTracker(object):
    """Task statuses last seen on the harness board; scan() returns the toasts."""

    def __init__(self, folder):
        self.path = os.path.join(folder, BOARD_FILE)
        self.seen = None
        self._lock = threading.Lock()

    def _board(self):
        try:
            with io.open(self.path, encoding='utf-8') as f:
                board = json.load(f)
        except Exception:
            return None                    # missing or mid-write
        return board if isinstance(board, dict) else None

    def prime(self):
        with self._lock:
            self.seen = describe_board(self._board() or {}, None)[1]

    def scan(self):
        with self._lock:
            board = self._board()
            if board is None:
                return []
            messages, self.seen = describe_board(board, self.seen)
            return messages


def toast(message):
    """pyRevit toast; False when it could not be shown."""
    try:
        from pyrevit.forms import toaster
        toaster.send_toast(message, title=TITLE, appid='T3Lab')
        return True
    except Exception:
        return False


class _Debounced(object):
    """Collapse a burst of watcher events into one scan."""

    def __init__(self, tracker, show):
        self.tracker, self.show = tracker, show
        self._timer = None
        self._lock = threading.Lock()

    def poke(self, sender=None, args=None):
        with self._lock:
            if self._timer is not None:
                self._timer.cancel()
            self._timer = threading.Timer(DEBOUNCE_S, self._fire)
            self._timer.daemon = True
            self._timer.start()

    def _fire(self):
        try:
            for text in self.tracker.scan():
                self.show(text)
        except Exception:
            pass


def install(uiapp=None):
    """Start the board watcher once per Revit session. Returns (state, detail)."""
    if opted_out():
        return SKIPPED, 'opted out'
    try:
        from System import AppDomain
        from System.IO import FileSystemWatcher, NotifyFilters
    except Exception as exc:
        return SKIPPED, 'no .NET here: {}'.format(exc)
    try:
        domain = AppDomain.CurrentDomain
        if domain.GetData(HARNESS_WATCHER_KEY) is not None:
            return SKIPPED, 'already watching'
        folder = harness_dir()
        if not os.path.isdir(folder):
            os.makedirs(folder)
        tracker = BoardTracker(folder)
        tracker.prime()
        debounced = _Debounced(tracker, toast)
        watcher = FileSystemWatcher(folder, BOARD_FILE)
        watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite
        watcher.Created += debounced.poke
        watcher.Changed += debounced.poke
        watcher.Renamed += debounced.poke          # tmp file replaced onto board.json
        watcher.EnableRaisingEvents = True
        domain.SetData(HARNESS_WATCHER_KEY, watcher)    # keeps it alive and marks it done
        return INSTALLED, folder
    except Exception as exc:
        return FAILED, '{}'.format(exc)
