# -*- coding: utf-8 -*-
"""Runs MCPExternalEventHandler.Execute for real, with Revit stubbed.

Every external tools/call goes through this method on Revit's main thread.
It broke twice from edits nothing executed (2026-11-17 Lite, 2026-11-18
dev: UnboundLocalError, the caller hung until its timeout), so here it runs:
whatever a tool or an Agent Cursors hook does, a claimed task must end with
task.done set and Execute must not raise into Revit.

Run: python lite_guard/test_execute_path.py
"""
import os
import sys
import types

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(REPO, 'lib'))


def _stub_revit():
    """Enough of clr / Autodesk.Revit.UI for server.py to define the handler."""
    clr = types.ModuleType('clr')
    clr.AddReference = lambda *a, **k: None
    sys.modules['clr'] = clr
    ui = types.ModuleType('Autodesk.Revit.UI')
    ui.IExternalEventHandler = object
    ui.ExternalEvent = type('ExternalEvent', (object,), {})
    revit = types.ModuleType('Autodesk.Revit')
    revit.UI = ui
    autodesk = types.ModuleType('Autodesk')
    autodesk.Revit = revit
    sys.modules.update({'Autodesk': autodesk, 'Autodesk.Revit': revit, 'Autodesk.Revit.UI': ui})


_stub_revit()
from core import server as S            # noqa: E402
from core import agent_cursors as AC    # noqa: E402

FAILED = []


def check(label, ok):
    print('  {} {}'.format('ok  ' if ok else 'FAIL', label))
    if not ok:
        FAILED.append(label)


class FakeServer(object):
    """Only what Execute needs; anything else answers None."""
    _WRITE_TOOLS = frozenset(['write_tool'])

    def __init__(self, tool=None):
        self.calls = []
        self._tool = tool or (lambda name, args: {'success': True, 'tool': name})
        self._write_in_progress = False
        self._write_seen = None

    def _execute_tool_in_context(self, name, args, **kwargs):
        self.calls.append(name)
        self._write_seen = self._write_in_progress
        return self._tool(name, args)

    def __getattr__(self, name):        # optional hooks (audit gate, turbo ...)
        return lambda *a, **k: None


def run(tasks, tool=None):
    """Queue `tasks`, run Execute once; (server, error raised by Execute or None)."""
    server = FakeServer(tool)
    handler = S.MCPExternalEventHandler(server)
    for task in tasks:
        handler.tasks.put(task)
    try:
        handler.Execute(object())
        return server, None
    except BaseException as ex:         # noqa: B902 - the point of the test
        return server, ex


def boom(*a, **k):
    raise RuntimeError('hook failed')


def main():
    check('server.py defines the handler with Revit stubbed', getattr(S, 'HAS_REVIT_UI', False))
    if FAILED:
        return 1

    print('a read and a write complete')
    read, write = S._ToolTask('read_tool', {}), S._ToolTask('write_tool', {})
    server, err = run([read, write])
    check('Execute returns', err is None)
    check('both tasks are done', read.done.is_set() and write.done.is_set())
    check('both ran, in order', server.calls == ['read_tool', 'write_tool'])
    check('results reach the caller', (read.result or {}).get('tool') == 'read_tool'
          and read.exception is None)
    check('the write flag is up during the write and down after',
          server._write_seen is True and server._write_in_progress is False)

    print('a tool that raises still answers')
    task = S._ToolTask('read_tool', {})
    server, err = run([task], tool=boom)
    check('Execute returns', err is None)
    check('the task is done and carries the error', task.done.is_set()
          and isinstance(task.exception, RuntimeError))

    print('AI code that calls sys.exit() ends the call, not Execute')

    def leave(*a, **k):
        raise SystemExit(0)

    task = S._ToolTask('read_tool', {})
    server, err = run([task], tool=leave)
    check('Execute returns', err is None)
    check('the task is done', task.done.is_set())

    if hasattr(AC, 'offer'):
        print('the "turn on Agent Cursors?" dialog never sits in front of a plain call')
        asked = []
        saved_offer = AC.offer
        AC.offer = lambda app, agent: asked.append(agent)
        try:
            task = S._ToolTask('write_tool', {})
            task.agent = 'codex'
            server, err = run([task])
            check('no dialog, the call ran', asked == [] and task.done.is_set()
                  and server.calls == ['write_tool'])
        finally:
            AC.offer = saved_offer

    print('a task the read fallback took is skipped')
    task = S._ToolTask('read_tool', {})
    task.claim('fallback')
    server, err = run([task])
    check('it does not run twice', err is None and server.calls == [])

    print('Agent Cursors hooks can never hang a call')
    saved = dict((n, getattr(AC, n)) for n in ('before', 'notify', 'paused_refusal', 'offer',
                                                'waiting') if hasattr(AC, n))
    try:
        for name in saved:
            setattr(AC, name, boom)
        named = S._ToolTask('write_tool', {})
        named.agent = 'codex'                 # an external AI app
        plain = S._ToolTask('read_tool', {})
        server, err = run([named, plain])
        check('Execute returns', err is None)
        check('both tasks are done', named.done.is_set() and plain.done.is_set())
        check('the tools ran', server.calls == ['write_tool', 'read_tool'])
    finally:
        for name, fn in saved.items():
            setattr(AC, name, fn)

    print('Pause holds a write, lets a read through')
    saved = AC.paused_refusal
    try:
        AC.paused_refusal = lambda tool, is_write: ({'success': False, 'error_code': 'paused'}
                                                    if is_write else None)
        write, read = S._ToolTask('write_tool', {}), S._ToolTask('read_tool', {})
        server, err = run([write, read])
        check('Execute returns', err is None)
        check('the write is answered "paused" and did not run', write.done.is_set()
              and (write.result or {}).get('error_code') == 'paused'
              and 'write_tool' not in server.calls)
        check('the read ran', read.done.is_set() and server.calls == ['read_tool'])
    finally:
        AC.paused_refusal = saved

    if FAILED:
        print('\n{} check(s) failed.'.format(len(FAILED)))
        return 1
    print('\nAll Execute path checks passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
