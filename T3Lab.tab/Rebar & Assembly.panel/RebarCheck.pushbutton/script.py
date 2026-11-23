#! python3
# -*- coding: utf-8 -*-
"""Rebar Check — find rebar without a valid host, rebar missing from its assembly, assemblies without drawings, duplicate numbers and bars outside their host."""

__title__   = "Rebar\nCheck"
__author__  = "Tran Tien Thanh"
__version__ = "1.0.0"

# ── IMPORTS & BOOTSTRAP ──────────────────────────────────────────────────────
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

SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import revit  # noqa: F401  (keeps pyRevit happy; not used at module level)

# ── ENTRY POINT ──────────────────────────────────────────────────────────────
if __name__ == '__main__':
    # Snippets._host.resolve_doc walks the real fallback chain (revit.doc ->
    # ActiveUIDocument -> the single open model) and returns a message worth
    # showing when there genuinely is nothing to act on.
    from Snippets._host import resolve_doc
    doc, doc_error = resolve_doc()

    if not doc:
        from GUI.T3Dialog import show_warning
        show_warning(
            doc_error or "Open a Revit project before running Rebar Check.",
            title="Rebar Check",
            details="Rebar Check needs an active document to read its rebar and assemblies from.")
    elif doc.IsFamilyDocument:
        from GUI.T3Dialog import show_warning
        show_warning(
            "Rebar Check works on projects, not on family documents.",
            title="Rebar Check",
            details="Open the project that contains the rebar and assemblies, then run it again.")
    else:
        from GUI.RebarCheckDialog import show_rebar_check
        show_rebar_check(doc)
