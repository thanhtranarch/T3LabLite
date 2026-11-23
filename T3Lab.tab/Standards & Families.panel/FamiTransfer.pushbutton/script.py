#! python3
# -*- coding: utf-8 -*-
"""Family Transfer — copy selected families (all types or some) between open projects."""
__title__ = "Family\nTransfer"
__author__ = "T3Lab"

# ── IMPORTS ──────────────────────────────────────────────────────────────
import os
import sys

# ── PATH SETUP ───────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Non-stacked: <tab>/<panel>/<tool>.pushbutton -> 3 levels up is the extension.
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
import _cpython_bootstrap
_cpython_bootstrap.init_cpython_paths()

# Not reloaded on purpose: Snippets.family_transfer defines .NET interface
# classes with a static __namespace__; a second definition in the same
# persistent engine raises "Duplicate type name within an assembly".
from GUI.FamilyTransferDialog import show_family_transfer

# ── ENTRY POINT ──────────────────────────────────────────────────────────
if __name__ == '__main__':
    show_family_transfer()
