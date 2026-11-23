#! python3
# -*- coding: utf-8 -*-
"""Harness — open Harness Studio, the BIM Harness board in the browser.

docs/harness.md. The Studio runs in the bridge process, never in Revit; this
button only starts it (or reopens the one already running).
"""
__title__ = 'Harness'
__author__ = 'T3Lab'

import os
import sys
import time

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

START_CHECK_S = 2.5     # a Studio that cannot start exits at once; longer would freeze Revit


def _error(message):
    from GUI.T3Dialog import show_error
    show_error(message, title='Harness Studio')


def main():
    from Services import harness_settings
    try:
        proc = harness_settings.open_studio()   # deploys the bridge files, then starts the Studio
    except Exception as ex:
        _error('Harness Studio could not start: {}.\n\nCheck that Python 3 is installed '
               '(MCP Control > Connect), then click Harness again.'.format(ex))
        return
    deadline = time.time() + START_CHECK_S
    while time.time() < deadline:
        code = proc.poll()
        if code is not None:
            if code != 0:
                _error('Harness Studio stopped while starting (exit {}).\n\nRun this in a '
                       'terminal to see why:\npython "%APPDATA%\\T3LabAI\\bridge.py" harness '
                       'studio'.format(code))
            return                          # 0: a Studio was already running, its page reopened
        time.sleep(0.1)


if __name__ == '__main__':
    main()
