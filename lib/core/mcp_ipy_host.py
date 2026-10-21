# -*- coding: utf-8 -*-
"""Start or stop the T3Lab MCP HTTP server inside pyRevit's IronPython engine.

No shebang on purpose: pyRevit picks the engine from the first line, and this
file must run in IronPython. Started by Services/mcp_service.py
(_run_ipy_host) when a CPython tool — MCP Control, the Assistant — asks for
the server, so the server never lives in the CPython engine, where it stops
answering as soon as no T3Lab window is open (see core/server.py,
start_server).

sys.argv: [this file, 'start' | 'stop', optional port].
The outcome goes to AppDomain data under RESULT_KEY as text — 'ok' or
'error: <reason>' — because a Python object cannot cross engines but a
string can. Never raises and never prints: an exception or a print would open
the pyRevit output window.
"""

import os
import sys

RESULT_KEY = '_t3lab_mcp_ipy_host_result'


def _report(text):
    try:
        from System import AppDomain
        AppDomain.CurrentDomain.SetData(RESULT_KEY, text)
    except Exception:
        pass


def _lib_dir():
    try:
        here = __file__
    except NameError:
        here = sys.argv[0] if sys.argv else ''
    return os.path.dirname(os.path.dirname(os.path.abspath(here)))


def _run():
    lib_dir = _lib_dir()
    ext_dir = os.path.dirname(lib_dir)
    for path in (ext_dir, lib_dir):
        if path not in sys.path:
            sys.path.insert(0, path)

    args = list(sys.argv[1:])
    action = args[0] if args else 'start'
    port = None
    if len(args) > 1:
        try:
            port = int(args[1])
        except ValueError:
            port = None

    from Services.mcp_service import MCPService
    if action == 'stop':
        ok, err = MCPService.stop_server()
    elif action == 'start':
        # Runs on Revit's main thread inside an API context (the caller's
        # command, or pyRevit's own ExternalEvent), so the ExternalEvent for
        # model-editing tools can be created here, in the engine that owns
        # the server.
        MCPService.ensure_external_event()
        ok, err = MCPService.start_server(port=port)
    else:
        ok, err = False, 'unknown action: {}'.format(action)
    _report('ok' if ok else 'error: {}'.format(err or 'unknown error'))


try:
    _run()
except Exception as ex:
    _report('error: {}'.format(ex))
