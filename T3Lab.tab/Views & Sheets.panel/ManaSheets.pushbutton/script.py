#! python3
# -*- coding: utf-8 -*-
"""
Sheet Manager
Unified tool to manage sheets, sets, views on sheets, parameters, and re-number sheets.

Copyright (c) 2026 T3Lab
"""
__title__ = "Sheet\nManager"
__author__ = "T3Lab"

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

try:
    import _cpython_bootstrap
    _cpython_bootstrap.init_cpython_paths()
except Exception:
    pass
# ──────────────────────────────────────────────────────────────────────────────

# Ensure lib directory is in sys.path
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

try:
    reload
except NameError:
    try:
        from importlib import reload
    except Exception:
        reload = None


def _main():
    # Everything that loads the dialog runs inside the guard: when it cannot
    # load (e.g. "cannot import name 'revit' from 'pyrevit'" after an update
    # without a Revit restart) the user gets a readable dialog, not a raw
    # traceback.
    if reload:
        # Nạp lại bộ lọc cột TRƯỚC dialog: dialog `from ... import` class này nên
        # nếu module cũ còn trong sys.modules thì click lần 2 vẫn chạy code cũ.
        # An toàn để reload vì nó chỉ chứa class Python thuần (không kế thừa CLR,
        # không `__namespace__`) — xem CLAUDE.md · S15/S17.
        if 'GUI.DataGridColumnFilter' in sys.modules:
            reload(sys.modules['GUI.DataGridColumnFilter'])
        if 'GUI.ManaSheetsDialog' in sys.modules:
            reload(sys.modules['GUI.ManaSheetsDialog'])
        elif 'ManaSheetsDialog' in sys.modules:
            reload(sys.modules['ManaSheetsDialog'])

    from GUI.ManaSheetsDialog import show_sheet_manager
    show_sheet_manager()


if __name__ == '__main__':
    try:
        from GUI.ErrorGuard import run_tool
    except Exception:
        _main()
    else:
        run_tool("Mana Sheets", _main)
