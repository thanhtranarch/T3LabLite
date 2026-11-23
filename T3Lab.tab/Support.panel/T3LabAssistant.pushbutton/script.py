#! python3
# -*- coding: utf-8 -*-
"""
T3Lab Assistant

Open the T3Lab AI assistant for natural language Revit commands.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__  = "Tran Tien Thanh"
__title__   = "T3Lab Assistant"
__version__ = "1.0.0"
__persistentengine__ = True

import os
import sys

_script_dir = os.path.dirname(__file__)
_ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(_script_dir)))
_lib_dir = os.path.join(_ext_dir, 'lib')
if _lib_dir not in sys.path:
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

from GUI.T3LabAssistantDialog import show_assistant_dialog

if __name__ == '__main__':
    show_assistant_dialog()
