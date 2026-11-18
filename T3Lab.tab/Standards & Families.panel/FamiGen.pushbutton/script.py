#! python3
# -*- coding: utf-8 -*-
__title__  = "FamiGen"
__author__ = "Tran Tien Thanh"
__doc__    = ("FamiGen Studio — chat with Claude to create Revit families, review them in 3D "
              "(pyRevit 7.0+). Shift+Click: classic window (From CAD, JSON).")
# FamiGen Studio stays open next to Revit (modeless) when the engine persists;
# without it the Studio runs modally (FamilyGen.studio_host.open_studio).
__persistentengine__ = True

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

try:
    import _cpython_bootstrap
    _cpython_bootstrap.init_cpython_paths()
except Exception:
    pass
# ──────────────────────────────────────────────────────────────────────────────

_ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
_lib = os.path.join(_ext_dir, 'lib')
if _lib not in sys.path:
    sys.path.insert(0, _lib)

# Force reload GUI modules to prevent caching issues in pyRevit
try:
    from importlib import reload as _reload
    for _mod in ('GUI.WPF_Base', 'GUI.FamiGenDialog', 'FamiGenDialog'):
        if _mod in sys.modules:
            try:
                _reload(sys.modules[_mod])
            except Exception:
                pass
except Exception:
    pass

from GUI.FamiGenDialog import show_family_creator
from Snippets._host import resolve_doc, injected_uiapp


def _shift_click():
    try:
        from pyrevit import EXEC_PARAMS
        return bool(EXEC_PARAMS.config_mode)
    except Exception:
        return False


def _persistent_engine():
    try:
        from GUI.ManaSelectDialog import detect_persistent_engine
        return detect_persistent_engine()
    except Exception:
        return False


def _studio_unavailable(message):
    """Say why the Studio cannot open; offer the classic window. True = open classic."""
    from Autodesk.Revit.UI import TaskDialog, TaskDialogCommandLinkId, TaskDialogResult
    dlg = TaskDialog("FamiGen Studio")
    dlg.MainInstruction = "FamiGen Studio cannot open"
    dlg.MainContent = message
    dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Open classic FamiGen",
                       "From CAD blocks, or paste a family JSON")
    dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Close")
    return dlg.Show() == TaskDialogResult.CommandLink1


if __name__ == '__main__':
    doc, err = resolve_doc()
    if not doc:
        from pyrevit import forms
        forms.alert(err or "No active document found in Revit.", title="FamiGen")
    elif _shift_click():
        show_family_creator(doc, doc.Application)
    else:
        from FamilyGen import studio_host
        state, message = studio_host.open_studio(doc, injected_uiapp(), _persistent_engine())
        if state == 'unsupported' and _studio_unavailable(message):
            show_family_creator(doc, doc.Application)

