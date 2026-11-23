#! python3
# -*- coding: utf-8 -*-
"""Auto Join
Automatically join intersecting Revit elements by category rules.
Define priority categories (which cut) and join-with categories (which get cut).
T3Lab tool.

- Click       : Open Auto Join Manager (WPF rule-based dialog)
- Shift+Click : Quick join with default rules (Walls ↔ Floors, Columns)

Author: Tran Tien Thanh
"""

__author__  = "Tran Tien Thanh"
__title__   = "Auto\nJoin"
__version__ = "1.1.0"

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

_lib = os.path.normpath(os.path.join(os.path.dirname(__file__), '../../../../../lib'))
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
    if 'GUI.AutoJoinDialog' in sys.modules:
        reload(sys.modules['GUI.AutoJoinDialog'])
    elif 'AutoJoinDialog' in sys.modules:
        reload(sys.modules['AutoJoinDialog'])

from pyrevit import revit
from GUI.AutoJoinDialog import show_auto_join_dialog, quick_join

if __name__ == '__main__':
    doc = None
    uidoc = None
    try:
        doc = revit.doc
        uidoc = revit.uidoc
    except Exception:
        pass

    try:
        is_shift = bool(__shiftclick__)
    except NameError:
        is_shift = False

    if is_shift:
        quick_join(doc, uidoc)
    else:
        show_auto_join_dialog(doc, uidoc)
