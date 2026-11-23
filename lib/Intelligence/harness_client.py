# -*- coding: utf-8 -*-
"""The T3Lab Assistant on the BIM Harness board (docs/harness.md).
Deterministic `/harness` commands, no LLM:

    /harness                      the board: sessions, who holds what, what waits
    /harness studio               the board as a canvas in the browser
    /harness accept <task>        as the person at this machine
    /harness redo <task> [note]

The pane only looks and reviews. Taking a task (plan / join / next / done,
t3lab_dev's Assistant) is not in T3Lab Lite: the pane's own edits reach Revit
in-process, past the bridge's write check, so it never becomes a board member.
The board's database lives in the bridge process: reached only through
`bridge.py harness ...` and read through board.json. Pure Python.
Test: lite_guard/test_harness_client.py.
"""
from __future__ import unicode_literals

import json

AGENT = 't3lab-assistant'
TIMEOUT_S = 30
MEMBER_ACTIONS = ('plan', 'join', 'next', 'done', 'failed')

HELP = ('`/harness` board · `/harness studio` (the board as a canvas in the browser) · '
        '`/harness accept <task>` · `/harness redo <task> [note]`')


def parse(text):
    """(action, arg) for a /harness command, else None."""
    t = (text or '').strip()
    if not t.lower().startswith('/harness'):
        return None
    rest = t[len('/harness'):]
    if rest and not rest[0].isspace():
        return None
    parts = rest.strip().split(None, 1)
    return (parts[0].lower() if parts else 'status'), (parts[1].strip() if len(parts) > 1 else '')


def run(action, args=None, session=None, runner=None, timeout=TIMEOUT_S):
    """`bridge.py harness <action>` as this pane: (exit code, payload)."""
    if runner is None:
        from Services.harness_settings import run_bridge as runner
    argv = ['harness', action, '--agent={}'.format(AGENT)]
    if session:
        argv.append('--session={}'.format(session))
    if args:
        argv += ['--json', json.dumps(args, ensure_ascii=False)]
    return runner(argv, timeout=timeout)


def _failed(code, payload):
    if code is None:
        return 'The bridge did not answer in {} s. Is Python 3 set up (MCP Control > ' \
               'Connect)?'.format(TIMEOUT_S)
    if isinstance(payload, dict) and payload.get('error'):
        hint = payload.get('hint')
        return '{}{}'.format(payload['error'], ' ' + hint if hint else '')
    if code != 0:
        return 'The bridge failed (exit {}): {}'.format(code, '{}'.format(payload)[:200])
    return None


def board_text(board):
    """Markdown for the board snapshot (core/agent_cursors.board_rows)."""
    from core.agent_cursors import board_rows
    sessions = board_rows(board)
    if not sessions:
        lead = (board or {}).get('master')
        return ('No harness session is open{}.\n{}'.format(
            ' (master: {})'.format(lead) if lead else '', HELP))
    lines = []
    marks = {'current': 'working', 'review': 'waiting for review', 'failed': 'failed'}
    for s in sessions:
        lines.append('**{}** — {}{}'.format(s['goal'] or s['session'], s['summary'],
                                             ' · ' + s['document'] if s['document'] else ''))
        for state, text, document in s['rows']:
            lines.append('- {} ({}){}'.format(text, marks[state],
                                               ' · ' + document if document else ''))
        if s['more']:
            lines.append('- … {} more'.format(s['more']))
    return '\n'.join(lines + ['', HELP])


def handle(text, runner=None, board=None, studio=None):
    """Markdown reply for one command, or None when `text` is not /harness."""
    parsed = parse(text)
    if parsed is None:
        return None
    action, arg = parsed
    if action in ('status', 'board', 'help'):
        if board is None:
            from core.agent_cursors import read_board
            board = read_board()
        return board_text(board)
    if action == 'studio':
        try:
            if studio is None:
                from Services.harness_settings import open_studio as studio
            studio()
        except Exception as ex:
            return ('Could not start Harness Studio: {}. Check that Python 3 is set up '
                    '(MCP Control > Connect).'.format(ex))
        return ('Harness Studio is opening in your browser. From a terminal: '
                '`python bridge.py harness studio`.')
    if action in MEMBER_ACTIONS:
        return ('`/harness {}` is not in T3Lab Lite: the Assistant does not take tasks on the '
                'board. Ask your AI app (Claude, Codex) to use t3lab_harness, or plan in '
                'Harness Studio (`/harness studio`).'.format(action))
    if action in ('accept', 'redo'):
        parts = arg.split(None, 1)
        if not parts:
            return 'Which task? `/harness {} <task>`.'.format(action)
        review = {'task': parts[0], 'verdict': action, 'owner': True}
        if len(parts) > 1:
            review['note'] = parts[1]
        code, out = run('review', review, runner=runner)
        problem = _failed(code, out)
        return problem or '`{}` is now {}.'.format(parts[0], out.get('status'))
    return 'Unknown /harness command. {}'.format(HELP)
