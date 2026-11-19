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
    """Checkboxes for the fallback dialog: the on/off switch (when on) + every bar option."""
    current = agent_cursors.options()
    rows = []
    if on:
        rows.append(('enabled', 'Show agent cursors', 'Untick to turn the cursors off', True))
    for name, _key, _default in agent_cursors.OPTIONS:
        label, tip = agent_cursors.OPTION_LABELS[name][:2]
        rows.append((name, label, tip, current.get(name, True)))
    return rows


def _ask(on):
    """{'enabled': bool, option: bool, ...} or None on Cancel. The settings
    window (GUI/AgentCursorsDialog.py); if it cannot be built, the plain
    checkbox dialog, which says why, so the setting is never out of reach."""
    try:
        from GUI.AgentCursorsDialog import ask
        return ask(on, agent_cursors.options(), agent_cursors.pros_cons_sections())
    except Exception as ex:
        why = '(settings window unavailable: {})'.format(ex)
    if on:
        message = 'Agent cursors are on. Pick the options, or untick "Show agent cursors" to turn them off.'
    else:
        message = ('Show a live cursor for each AI agent (Claude, GPT, Gemini…) '
                   'working in Revit through MCP?')
    return confirm_options(message, _option_rows(on), title='Agent Cursors',
                           ok_text='Save' if on else 'Turn On', cancel_text='Cancel',
                           details=agent_cursors.pros_cons_text() + '\n\n' + why)


def main():
    on = agent_cursors.is_enabled()
    values = _ask(on)
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
            details='The bar appears on the next MCP tool call: each agent, its goal '
                    'and tasks, a plain caption of each step and the values it changes. '
                    'Tick "On model" to also draw a cursor per agent and the elements it '
                    'visits on the views. Long calls say "Working"; a '
                    'call waiting for your OK pulses amber. When the agents go idle the '
                    'bar shows a summary with Select touched and Replay. The same options '
                    'and what each icon means are on the bar; its X turns the cursors off '
                    'again.\n\n'
                    'Options: {}'.format(agent_cursors.options_text()))


if __name__ == '__main__':
    main()
