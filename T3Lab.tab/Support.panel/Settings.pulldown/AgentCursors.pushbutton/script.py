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
from GUI.T3Dialog import confirm, show_info, show_error


def main():
    on = agent_cursors.is_enabled()
    if on:
        message = 'Agent cursors are ON. Turn them off?'
        ok_text = 'Turn Off'
    else:
        message = ('Show a live cursor for each AI agent (Claude, GPT, Gemini…) '
                   'working in Revit through MCP?')
        ok_text = 'Turn On'
    details = 'Options: {}\n\n{}'.format(agent_cursors.options_text(),
                                         agent_cursors.pros_cons_text())
    if not confirm(message, title='Agent Cursors', ok_text=ok_text,
                            cancel_text='Cancel', details=details):
        return
    try:
        agent_cursors.set_enabled(not on)
    except Exception as ex:
        show_error(
            'Could not save the Agent Cursors setting.',
            title='Agent Cursors',
            details='{}\nCheck that %APPDATA%\\T3LabAI\\mcp_paths.json is writable, '
                    'then try again.'.format(ex))
        return
    if on:
        show_info('Agent cursors are off.', title='Agent Cursors')
    else:
        show_info(
            'Agent cursors are on.', title='Agent Cursors',
            details='They appear on the next MCP tool call, one cursor per agent and '
                    'action (Read, Create, Modify, Delete, View, Export, Code), and fade '
                    'out after a few idle seconds. Follow opens the views the agent '
                    'works in. The bar at the bottom of the view switches Follow, Zoom, '
                    'Select and Tile; its X turns the cursors off again.\n\n'
                    'Options: {}'.format(agent_cursors.options_text()))


if __name__ == '__main__':
    main()
