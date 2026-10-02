# -*- coding: utf-8 -*-
"""
Usage tracking entry points -- the only functions the rest of T3Lab calls.

  track_session_start()  startup.py, once per Revit start / pyRevit reload
  track_ribbon_click()   _cpython_bootstrap.init_cpython_paths(), which every
                         T3Lab.tab script.py calls, so it sees every click
  track_mcp_call()       core/server.py _handle_tool_call(), success and error

None of them ever raises or blocks: TelemetryService sends on a daemon thread.

Lite only. The call sites live in files shared with t3lab_dev, so a copy from
dev can drop them; lite_guard/manifest.json fails the commit when that happens.
"""
from __future__ import unicode_literals

import os
import sys
import time


def _service():
    from tracking.service import TelemetryService
    return TelemetryService


def track_session_start():
    """One row per Revit start / pyRevit reload."""
    try:
        _service().record_tool_usage(
            tool_name='session_start',
            tool_type='startup',
            purpose='Khởi động phiên làm việc Revit / pyRevit'
        )
    except Exception:
        pass


# init_cpython_paths() runs twice per click: once when script.py imports
# _cpython_bootstrap (it calls itself at module level) and once when the
# script calls it. Both runs see the same script globals, so the first one
# stamps them and the second one, a moment later, is skipped. The time limit
# keeps a later click counted even if pyRevit reuses the globals.
_CLICK_STAMP = '__t3lab_tracked_at__'
_SAME_CLICK_SECONDS = 5.0


def find_caller_script():
    """(path, globals) of the T3Lab.tab script.py up the call stack, or ('', None)."""
    try:
        frame = sys._getframe(1)
    except Exception:
        return '', None

    while frame:
        # pyRevit may compile the script from a string ("<string>"), so also
        # look at the __file__ it sets in the script's globals.
        candidates = (getattr(frame.f_code, 'co_filename', '') or '',
                      frame.f_globals.get('__file__') or '')
        for f_name in candidates:
            if f_name.endswith('script.py') and 'T3Lab.tab' in f_name:
                return f_name, frame.f_globals
        frame = frame.f_back
    return '', None


def _already_tracked(script_globals):
    """True when this script run already sent its ribbon row."""
    now = time.time()
    try:
        last = script_globals.get(_CLICK_STAMP)
        script_globals[_CLICK_STAMP] = now
    except Exception:
        return False
    return last is not None and now - last < _SAME_CLICK_SECONDS


def track_ribbon_click():
    """One row for the T3Lab ribbon button whose script.py is running.

    Does nothing when the caller is not a T3Lab.tab script (tests, MCP code).
    """
    try:
        caller_file, script_globals = find_caller_script()
        if not caller_file or _already_tracked(script_globals):
            return

        parts = caller_file.replace('/', '\\').split('\\')
        tool_name = None
        panel_name = None
        for part in parts:
            if part.endswith('.panel'):
                panel_name = part[:-6]
            if part.endswith(('.pushbutton', '.smartbutton', '.linkbutton')):
                tool_name = part.rsplit('.', 1)[0]

        if not tool_name:
            tool_name = os.path.basename(os.path.dirname(caller_file))

        _service().record_tool_usage(
            tool_name=tool_name,
            tool_type='ribbon',
            panel=panel_name,
            purpose=u"Người dùng chạy công cụ {} trên Revit Ribbon".format(tool_name)
        )
    except Exception:
        pass


def track_mcp_call(tool_name, arguments, result, t_start):
    """One row per MCP tool call; `t_start` is time.time() before the call."""
    try:
        duration_ms = int((time.time() - t_start) * 1000)
        _service().record_mcp_call(tool_name, arguments, result, duration_ms)
    except Exception:
        pass
