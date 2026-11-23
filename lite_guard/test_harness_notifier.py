# -*- coding: utf-8 -*-
"""CPython test: lib/harness_notifier.py (board toast text, tracking, install
outside Revit) and its startup.py hook. From the board part of t3lab_dev's
dev/test_approval_notifier.py. Run: python3 lite_guard/test_harness_notifier.py"""
import json
import os
import sys
import tempfile
import time

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'lib'))
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_harness_toast_test_')
os.environ.pop('T3LAB_HARNESS_DIR', None)

import harness_notifier as N   # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    print('  {}  {}{}'.format('ok  ' if cond else 'FAIL', name, '' if cond else '  ' + str(detail)))
    if not cond:
        FAILURES.append(name)


print('[harness board toasts]')


def board(*tasks, **kw):
    return {'sessions': [dict({'session': 's1', 'goal': 'Audit doors',
                               'tasks': [{'task': k, 'title': t, 'status': st, 'owner': o}
                                         for k, t, st, o in tasks]}, **kw)]}


first = board(('a', 'List doors', 'claimed', 'console-qa'), ('b', 'Fix marks', 'todo', None))
msgs, seen = N.describe_board(first, None)
check('first look announces nothing', msgs == [] and seen == {'s1/a': 'claimed', 's1/b': 'todo'})
msgs, seen = N.describe_board(first, seen)
check('no change, no toast', msgs == [])
msgs, seen = N.describe_board(
    board(('a', 'List doors', 'review', 'console-qa'), ('b', 'Fix marks', 'claimed', 'codex')), seen)
check('a task waiting for review', len(msgs) == 1 and msgs[0].startswith(
    'console-qa finished "List doors". It waits for review: MCP Control > Harness'), msgs)
msgs, seen = N.describe_board(
    board(('a', 'List doors', 'done', 'console-qa'), ('b', 'Fix marks', 'failed', 'codex')), seen)
check('a failed task', msgs == ['codex could not finish "Fix marks". The lead can retry it.'], msgs)
msgs, seen = N.describe_board(
    board(('a', 'List doors', 'done', 'console-qa'), ('b', 'Fix marks', 'done', 'codex')), seen)
check('board finished', msgs == ['Harness: all 2 tasks of "Audit doors" are done.'], msgs)
check('finished board is not announced twice', N.describe_board(
    board(('a', 'List doors', 'done', 'console-qa'), ('b', 'Fix marks', 'done', 'codex')),
    seen)[0] == [])
check('no board', N.describe_board({}, seen) == ([], {}))

HARNESS = N.harness_dir()
os.makedirs(HARNESS)
bt = N.BoardTracker(HARNESS)
bt.prime()
check('no file yet: nothing', bt.scan() == [])
with open(bt.path, 'w') as f:
    json.dump(first, f)
check('tasks that appear after start are not toasts by themselves', bt.scan() == [])
with open(bt.path, 'w') as f:
    json.dump(board(('a', 'List doors', 'review', 'console-qa'),
                    ('b', 'Fix marks', 'todo', None)), f)
check('tracker toasts the review once', len(bt.scan()) == 1 and bt.scan() == [])
with open(bt.path, 'w') as f:
    f.write('{"sessions": [')
check('half-written board skipped', bt.scan() == [])

print('[debounce]')
shown = []


class _Once(object):
    calls = 0

    def scan(self):
        _Once.calls += 1
        return ['one']


d = N._Debounced(_Once(), shown.append)
for _ in range(5):
    d.poke(None, None)
time.sleep(N.DEBOUNCE_S + 0.4)
check('a burst of events -> one scan, one toast', (_Once.calls, shown) == (1, ['one']),
      (_Once.calls, shown))

print('[install outside Revit]')
state, detail = N.install(None)
check('no .NET: skipped, never raises', state == N.SKIPPED, (state, detail))
for env in N.OPT_OUT_ENV:
    os.environ[env] = '1'
    check('opt-out respected: ' + env, N.install(None)[0] == N.SKIPPED and N.opted_out())
    del os.environ[env]
open(os.path.join(N._data_dir(), 'harness_toast.disabled'), 'w').close()
check('opt-out file respected', N.opted_out())
check('toast() never raises without pyRevit', N.toast('x') is False)
check('same AppDomain key as t3lab_dev (one watcher per Revit)',
      N.HARNESS_WATCHER_KEY == 'T3Lab.HarnessNotifier.Watcher')

with open(os.path.join(REPO, 'startup.py'), encoding='utf-8') as f:
    startup = f.read()
check('startup.py installs it', 'state, detail = install_harness_notifier()' in startup
      and '_load_lib_module("harness_notifier")' in startup)

print('\n{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)) if FAILURES else '\nall passed')
sys.exit(1 if FAILURES else 0)
