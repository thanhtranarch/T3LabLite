#! python3
# -*- coding: utf-8 -*-
"""ManaSelect — Element explorer: counted tree of every element in scope.

Tick any Category / Family / Type, then Select, Add to Selection, Zoom,
Isolate, Hide, Export CSV or Delete. (Quick Select / Select Similar /
On Sheets / Warnings were removed 2026-10-02 — they duplicated Explore,
Revit's own Select All Instances, ManaDWG and ModelAuditor.)

Cửa sổ chạy MODELESS để người dùng bấm chọn trong model mà tool vẫn mở. Điều
đó cần engine thường trú, nên `__persistentengine__ = True` ở dưới; nếu engine
chưa được kích hoạt (chưa reload pyRevit sau khi thêm cờ này), `show_dialog()`
tự rơi về modal thay vì mở ra một cửa sổ chết.

Author: T3Lab
"""
__title__ = "Mana\nSelect"
__author__ = "T3Lab"
__persistentengine__ = True

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

# Add lib directory to system path
# __file__ is T3Lab.extension/<Tab>.tab/Annotation & Select.panel/Mana.stack/ManaSelect.pushbutton/script.py
extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

# Import and show the dialog
import GUI.ManaSelectDialog as ManaSelectDialog

if __name__ == '__main__':
    # modal=None: show_dialog() tự dò engine thường trú và chọn modeless/modal.
    ManaSelectDialog.show_dialog()
