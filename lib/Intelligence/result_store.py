# -*- coding: utf-8 -*-
"""
Full copies of tool results the agent loop had to cut.

agent_loop feeds the model at most `result_chars` of each tool result. When
a result is bigger, the loop shrinks it (see agent_loop._shrink_result) and
— through the `spill_result` seam — saves the complete result here, so the
JSON the model reads can name a `full_result_path` instead of silently
losing rows. Files live under %APPDATA%\\T3LabAI\\results\\<YYYYMMDD>\\ and
are pruned after RETENTION_DAYS; nothing here is ever read by the model
itself (Phase 1 of the Tool Hub plan adds a read-back tool).

Write discipline as in telemetry.py: ASCII-safe JSON through core.jsonsafe,
written via io.open(encoding='utf-8') so IronPython 2.7 never truncates a
file on a Vietnamese name. Never raises.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__ = "Result Store"

import io
import os
import re
import threading
import time

RETENTION_DAYS = 7
_counter = [0]
_lock = threading.Lock()
_pruned = [False]


def _appdata_dir():
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI')


def results_dir(when=None):
    """Today's folder, created on demand."""
    day = time.strftime('%Y%m%d', time.localtime(when or time.time()))
    d = os.path.join(_appdata_dir(), 'results', day)
    if not os.path.isdir(d):
        try:
            os.makedirs(d)
        except Exception:
            pass
    return d


def _dumps(obj):
    try:
        from core import jsonsafe
        return jsonsafe.dumps(obj, indent=2, default=u'{}'.format)
    except Exception:
        import json
        return json.dumps(obj, ensure_ascii=True, indent=2, default=u'{}'.format)


def save_result(tool, result, turn_id=None):
    """Write `result` in full; return the file path, or None on failure."""
    try:
        with _lock:
            _counter[0] += 1
            idx = _counter[0]
        safe_tool = re.sub(r'[^A-Za-z0-9_.-]+', '_', u'{}'.format(tool or 'tool'))[:40]
        stamp = time.strftime('%H%M%S')
        name = u'{}_{}_{:03d}_{}.json'.format(
            stamp, u'{}'.format(turn_id or 't')[:12], idx, safe_tool)
        path = os.path.join(results_dir(), name)
        with io.open(path, 'w', encoding='utf-8') as f:
            f.write(_dumps(result))
        return path
    except Exception:
        return None


def prune(days=RETENTION_DAYS, now=None):
    """Delete day folders older than `days`. Best-effort; returns the count."""
    removed = 0
    try:
        root = os.path.join(_appdata_dir(), 'results')
        cutoff = (now or time.time()) - days * 86400.0
        for name in os.listdir(root):
            folder = os.path.join(root, name)
            if not os.path.isdir(folder) or not re.match(r'^\d{8}$', name):
                continue
            try:
                if os.path.getmtime(folder) >= cutoff:
                    continue
                for fn in os.listdir(folder):
                    try:
                        os.remove(os.path.join(folder, fn))
                        removed += 1
                    except Exception:
                        pass
                os.rmdir(folder)
            except Exception:
                pass
    except Exception:
        pass
    return removed


def prune_once():
    """prune() once per session."""
    if _pruned[0]:
        return 0
    _pruned[0] = True
    return prune()
