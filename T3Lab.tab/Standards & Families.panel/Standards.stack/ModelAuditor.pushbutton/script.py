#! python3
# -*- coding: utf-8 -*-
__title__ = "Model\nAuditor"
__author__ = "T3Lab"
__doc__ = "Model Auditor — Model check, warnings, in-place models, and material lists."

import os, sys
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

# script -> pushbutton -> Standards.stack -> panel -> tab -> extension
_ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))))
_lib = os.path.join(_ext_dir, 'lib')
if _lib not in sys.path:
    sys.path.insert(0, _lib)


def _main():
    # Imported inside the guard: when the dialog module cannot load (e.g.
    # "cannot import name 'revit' from 'pyrevit'" after an update without a
    # Revit restart) the user gets a readable dialog, not a raw traceback.
    from GUI.ModelAuditorDialog import show_model_auditor
    show_model_auditor(os.path.dirname(__file__), __revit__)


if __name__ == '__main__':
    try:
        from GUI.ErrorGuard import run_tool
    except Exception:
        _main()
    else:
        run_tool("Model Auditor", _main)
