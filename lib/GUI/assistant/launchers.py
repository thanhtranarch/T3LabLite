# -*- coding: utf-8 -*-
"""
Assistant tool launchers — open a T3Lab tool from the chat, and the intent ->
launcher / title registry rebuilt from tool discovery.

TOOL_LAUNCHERS, TOOL_TITLES and DROPPED_TOOLS are mutated IN PLACE (never
rebound), so every module that imported them shares the same objects.

Split out of GUI/T3LabAssistantDialog.py, which re-imports every name
here: `from GUI.T3LabAssistantDialog import <name>` keeps working.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Launchers"

import os

from GUI.assistant._log import logger
from GUI.assistant.text import _exc_text
from GUI.assistant.deps import (HAS_EXECUTOR, HAS_NLP, get_registered_tools,
                                inject_discovered_tools, make_launcher_for)
try:
    from GUI.assistant.deps import configure_batchout_window
except ImportError:
    pass    # executor missing: HAS_EXECUTOR is False, never called


# ─── Tool launchers ───────────────────────────────────────────────────────────

def _get_tool_script_dir(*parts):
    """Path to a pushbutton's script.py, found by folder name (only the LAST part counts)."""
    from core.extension_paths import bundle_path
    return bundle_path(parts[-1], 'script.py')


def _load_script(name, script_path):
    """Load a tool script as a module. Works in both CPython and IronPython."""
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location(name, script_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod
    except Exception:
        pass
    try:
        import imp
        return imp.load_source(name, script_path)
    except Exception:
        pass
    return None


def _load_batchout_mod():
    """Load the BatchOut script module, raising RuntimeError on failure."""
    try:
        from GUI import BatchOutDialog
        return BatchOutDialog
    except Exception:
        pass
    script_path = _get_tool_script_dir('BatchOut.pushbutton')
    mod = _load_script('batchout_script', script_path)
    if mod is None:
        raise RuntimeError("Could not load BatchOut module from: {}".format(script_path))
    return mod


def launch_batchout():
    """Open the BatchOut export dialog. Returns (ok, error_text) like TOOL_LAUNCHERS."""
    try:
        mod = _load_batchout_mod()
        window = mod.ExportManagerWindow()
        window.show(modal=not window.modeless)
        return True, u""
    except Exception as ex:
        logger.error("Error launching BatchOut: {}".format(_exc_text(ex)))
        return False, _exc_text(ex)


def launch_batchout_configured(config, progress_cb=None):
    """Open BatchOut pre-configured (sheets, format, Create tab). Returns bool.

    config: dict with format, filter; progress_cb: optional callable(str).
    """
    try:
        mod = _load_batchout_mod()
        window = mod.ExportManagerWindow()

        if HAS_EXECUTOR:
            configure_batchout_window(window, config)
            fmt    = (config.get('format') or 'pdf').upper()
            filt   = config.get('filter') or ''
            filt_s = u" {} sheet".format(filt) if filt else u" all sheets"
            if progress_cb:
                progress_cb(u"BatchOut selected{}, format {} — press Export to run.".format(
                    filt_s, fmt))

        window.show(modal=not window.modeless)
        return True
    except Exception as ex:
        logger.error(u"Error launching configured BatchOut: {}".format(_exc_text(ex)))
        if progress_cb:
            progress_cb(u"Error: {}".format(_exc_text(ex)))
        return False


def launch_loadfamily():
    """Open the Family Manager on its Loader tab."""
    try:
        from GUI.ManaFamiDialog import show_family_manager, TAB_LOADER
        show_family_manager(mode=TAB_LOADER)
        return True, u""
    except Exception as ex:
        logger.error("Error launching LoadFamily: {}".format(_exc_text(ex)))
        return False, _exc_text(ex)


# Other launchers come from the tool_discovery registry, validated against disk
# (2026-07-28: hardcoded ones pointed at missing folders). See _build_tool_launchers.


# Only launchers that cannot be auto-derived; every other intent comes from the
# registry (_register_discovered_launchers) — a hardcoded map drifted stale.
_SPECIAL_LAUNCHERS = {
    "open_batchout":   launch_batchout,
    "open_loadfamily": launch_loadfamily,
}

TOOL_LAUNCHERS = dict(_SPECIAL_LAUNCHERS)

# intent → display title, for user-facing "Opening X..." messages.
TOOL_TITLES = {
    "open_batchout":   u"BatchOut",
    "open_loadfamily": u"Family Loader",
}

# Registry entries whose script.py is missing; shown in the activity log.
DROPPED_TOOLS = []


def _register_discovered_launchers(tools):
    """Rebuild TOOL_LAUNCHERS from the registry, validating each against disk.

    An intent whose script.py is missing never reaches the model (it would
    promise to open it).
    """
    del DROPPED_TOOLS[:]
    for tool in tools:
        intent = tool.get('intent')
        path   = tool.get('script_path')
        url    = tool.get('url')
        if not intent or intent in _SPECIAL_LAUNCHERS:
            continue
        # A urlbutton has no script.py; its hyperlink is the target.
        if not url and not (path and os.path.exists(path)):
            DROPPED_TOOLS.append((intent, path or u"<no path>"))
            TOOL_LAUNCHERS.pop(intent, None)
            TOOL_TITLES.pop(intent, None)
            continue
        title = tool.get('title') or intent
        TOOL_LAUNCHERS[intent] = make_launcher_for(tool)
        TOOL_TITLES[intent] = title

    if HAS_NLP:
        try:
            inject_discovered_tools(get_registered_tools())
        except Exception:
            pass


def get_tool_title(intent):
    """Display label for an intent; falls back to a readable form of the id."""
    if intent in TOOL_TITLES:
        return TOOL_TITLES[intent]
    return (intent or u"").replace('open_', '').replace('_', ' ').strip().title()

