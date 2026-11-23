#! python3
# -*- coding: utf-8 -*-
"""Workset Manager
Manage Revit worksets using a rule-based assignment interface.

Author: Tran Tien Thanh & T3Lab
Mail: trantienthanh909@gmail.com
"""

__author__ = "Tran Tien Thanh"
__title__  = "Workset\nManager"
__version__ = "1.0.0"

import os
import sys

# ─── CPython 3 & lib bootstrap ────────────────────────────────────────────────
# CPython engine paths come from lib/_cpython_bootstrap.py below - it finds
# the engine whatever the pyRevit clone is named or wherever it is installed.

_cur = os.path.dirname(os.path.abspath(__file__))
while _cur and not os.path.exists(os.path.join(_cur, 'lib')):
    _parent = os.path.dirname(_cur)
    if _parent == _cur:
        break
    _cur = _parent
_lib_dir = os.path.join(_cur, 'lib')
if os.path.exists(_lib_dir) and _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

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

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../lib'))
if _lib not in sys.path:
    sys.path.insert(0, _lib)

try:
    reload
except NameError:
    try:
        from importlib import reload
    except Exception:
        reload = None

if reload:
    if 'GUI.ManaWorksetDialog' in sys.modules:
        reload(sys.modules['GUI.ManaWorksetDialog'])
    elif 'ManaWorksetDialog' in sys.modules:
        reload(sys.modules['ManaWorksetDialog'])

from pyrevit import revit
from GUI.ManaWorksetDialog import show_workset_manager, quick_remove_unused

if __name__ == '__main__':
    doc = None
    try:
        doc = revit.doc
    except Exception:
        pass

    try:
        is_shift = bool(__shiftclick__)
    except NameError:
        is_shift = False

    if is_shift:
        quick_remove_unused(doc)
    else:
        show_workset_manager(doc)
