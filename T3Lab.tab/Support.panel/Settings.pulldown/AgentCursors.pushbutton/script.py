#! python3
# -*- coding: utf-8 -*-
"""Agent Cursors — turn the live AI-agent cursor overlay on or off."""
__title__ = 'Agent\nCursors'
__author__ = 'T3Lab'

import os
import sys

_cur = os.path.dirname(os.path.abspath(__file__))
while _cur and not os.path.exists(os.path.join(_cur, 'lib')):
    _parent = os.path.dirname(_cur)
    if _parent == _cur:
        break
    _cur = _parent
_lib_dir = os.path.join(_cur, 'lib')
if os.path.exists(_lib_dir) and _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

try:
    import _cpython_bootstrap
    _cpython_bootstrap.init_cpython_paths()
except Exception:
    pass

from core import agent_cursors
from GUI.T3Dialog import confirm_options, show_info, show_error


def _option_rows(on):
    """Checkboxes for the dialog: the on/off switch (when on) + every bar option."""
    current = agent_cursors.options()
    rows = []
    if on:
        rows.append(('enabled', 'Show agent cursors', 'Untick to turn the cursors off', True))
    for name, _key, _default in agent_cursors.OPTIONS:
        label, tip = agent_cursors.OPTION_LABELS[name][:2]
        rows.append((name, label, tip, current.get(name, True)))
    return rows


def main():
    on = agent_cursors.is_enabled()
    if on:
        message = 'Agent cursors are on. Pick the options, or untick "Show agent cursors" to turn them off.'
        ok_text = 'Save'
    else:
        message = ('Show a live cursor for each AI agent (Claude, GPT, Gemini…) '
                   'working in Revit through MCP?')
        ok_text = 'Turn On'
    values = confirm_options(message, _option_rows(on), title='Agent Cursors', ok_text=ok_text,
                             cancel_text='Cancel', details=agent_cursors.pros_cons_text())
    if values is None:
        return
    enable = values.pop('enabled', True)
    try:
        for name, value in values.items():
            agent_cursors.set_option(name, value)
        if enable != on:
            agent_cursors.set_enabled(enable)
    except Exception as ex:
        show_error(
            'Could not save the Agent Cursors setting.',
            title='Agent Cursors',
            details='{}\nCheck that %APPDATA%\\T3LabAI\\mcp_paths.json is writable, '
                    'then try again.'.format(ex))
        return
    if not enable:
        show_info('Agent cursors are off.', title='Agent Cursors')
    elif not on:
        show_info(
            'Agent cursors are on.', title='Agent Cursors',
            details='They appear on the next MCP tool call: one cursor per agent and '
                    'action, a plain caption of what it does and its goal, the elements '
                    'it visits and the values it changes. Long calls say "Working"; a '
                    'call waiting for your OK pulses amber. When the agents go idle the '
                    'bar shows a summary with Select touched and Replay. The same options, '
                    'plus Speed, switch on the bar; its X turns the cursors off again.\n\n'
                    'Options: {}'.format(agent_cursors.options_text()))


if __name__ == '__main__':
    main()
