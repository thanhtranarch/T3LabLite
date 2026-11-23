#! python3
# -*- coding: utf-8 -*-
"""MakePattern — Interactive Vector Hatch Studio & Large Surface Preview.

Author: T3Lab
"""
__title__ = "Make\nPattern"
__author__ = "T3Lab"

import os
import sys

# ─── CPython 3 & lib bootstrap ────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# 3 levels for non-stacked (Annotation & Select.panel/MakePattern.pushbutton)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# T3Lab shares this CPython engine and its module names: never run its bootstrap
# (_cpython_bootstrap.claim_lib, lite_guard/test_shared_engine.py).
_t3_lib = os.path.join(os.path.abspath(__file__).split('.tab' + os.sep)[0].rsplit(os.sep, 1)[0], 'lib')
_t3_bs = sys.modules.get('_cpython_bootstrap')
if _t3_bs is not None and not os.path.normcase(os.path.abspath(
        getattr(_t3_bs, '__file__', '') or '')).startswith(os.path.normcase(_t3_lib) + os.sep):
    del sys.modules['_cpython_bootstrap']
if not sys.path or os.path.normcase(sys.path[0]) != os.path.normcase(_t3_lib):
    sys.path.insert(0, _t3_lib)
try:
    import _cpython_bootstrap
    _cpython_bootstrap.init_cpython_paths()
except Exception:
    pass
# ──────────────────────────────────────────────────────────────────────────────

def _main():
    from GUI.MakePatternDialog import show_dialog
    show_dialog()

if __name__ == '__main__':
    try:
        from GUI.ErrorGuard import run_tool
    except Exception:
        _main()
    else:
        run_tool('MakePattern', _main)
