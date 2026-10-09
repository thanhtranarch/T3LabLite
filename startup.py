# -*- coding: utf-8 -*-
"""
T3Lab Extension Startup Script
================================
Runs once per pyRevit load (OnStartup phase), BEFORE any T3Lab tool.

Không có shebang là CÓ Ý: file này chạy trên engine mặc định của pyRevit
(IronPython 2.7 / 3.4), không phải CPython. Mọi tool T3Lab là `#! python3`;
khi engine CPython của pyRevit không nạp được python3XX.dll, tool chết ngay
trong pythonnet trước khi chạy tới dòng code T3Lab nào, và Revit chỉ hiện
"The type initializer for 'Delegates' threw an exception". IronPython không
cần DLL đó, nên ở đây vẫn kiểm được, tìm ra lý do thật và nói bằng lời.

Responsibilities:
  0. Fix pyRevit's PYREVIT_CPYVERSION ("3.12.3" -> "3123") so '#! python3'
     tools can start at all (startup_fix_cpyversion.py).
  1. Register the T3Lab Assistant as a native Revit DockablePane (IronPython host
     in assistant_pane.py; CPython content mounts on button click).
  2. Auto-patch pyRevit `sessionmgr.py` via `lib/pyrevit_patches.py` so pyRevit
     Reload does not disable CPython engine ("This property must be set before
     runtime is initialized").
  3. Startup engine & runtime mismatch diagnostics.
  4. Retire ribbon folders an update left behind (lib/core/ribbon_retire.py).
  5. Register the right-click context-menu entry (Revit 2025+).
  6. Self-study idle loop (opt-in).
  7. Start the file-based task watcher (lib/core/file_watcher.py).
  8. Deploy the MCP bridge to %APPDATA%/T3LabAI/bridge.py and auto-start the
     MCP server (lib/Services/mcp_service.py).
  9. Run the once-a-week update check against GitHub (lib/core/updater.py).
 10. Record session start telemetry (lib/tracking/).

Runs under IronPython 2.7 / 3.4. No f-strings or Python-3-only syntax.
Never raises: a startup failure must not prevent Revit from opening.
"""
import os
import sys

# ─── Path bootstrap ────────────────────────────────────────────────────────────
_STARTUP_DIR = os.path.dirname(os.path.abspath(__file__))
_LIB_DIR     = os.path.join(_STARTUP_DIR, 'lib')
for _p in (_STARTUP_DIR, _LIB_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)

SUPPORTED_REVIT = (2022, 2027)
TITLE = "T3Lab"
LOG_NAME = "engine_check.log"
LOG_MAX_BYTES = 256 * 1024          # then rotated to engine_check.log.1
PATCH_OPT_OUT_ENV = "T3LAB_NO_PYREVIT_PATCH"
PATCH_OPT_OUT_FILE = "pyrevit_patch.disabled"


def _clone_root():
    """pyRevit clone folder, from pyrevitlib/pyrevit/__init__.py."""
    try:
        import pyrevit
        pkg = os.path.dirname(os.path.abspath(pyrevit.__file__))
        return os.path.dirname(os.path.dirname(pkg))
    except Exception:
        return None


def _engine_dlls(clone):
    """[(engine name, python3XX.dll or None)] pyRevit would run '#! python3' with."""
    try:
        from pyrevit.userconfig import user_config
        engine = user_config.get_active_cpython_engine()
        dll = str(getattr(engine, "AssemblyPath", "") or "") if engine is not None else ""
        if dll.lower().endswith(".dll") and os.path.isfile(dll):
            return [(os.path.basename(os.path.dirname(dll)), dll)]
    except Exception:
        pass
    root = os.path.join(clone, "bin", "cengines") if clone else ""
    try:
        names = sorted(os.listdir(root))
    except Exception:
        return []
    return [(n, _python_dll(os.path.join(root, n))) for n in names
            if n.upper().startswith("CPY") and os.path.isdir(os.path.join(root, n))]


def _python_dll(engine_dir):
    try:
        for fn in os.listdir(engine_dir):
            low = fn.lower()
            if low.startswith("python3") and low.endswith(".dll") and low != "python3.dll":
                return os.path.join(engine_dir, fn)
    except Exception:
        pass
    return None


def _load_error(path):
    """None when the DLL loads (or cannot be tested here), else the reason."""
    try:
        from System.Runtime.InteropServices import NativeLibrary   # .NET Core: Revit 2025+
    except Exception:
        NativeLibrary = None
    if NativeLibrary is not None:
        try:
            NativeLibrary.Load(path)
            return None
        except Exception as exc:
            return getattr(exc, "Message", None) or str(exc)
    try:
        import ctypes                                               # .NET Framework: Revit 2022-2024
        kernel32 = ctypes.windll.kernel32
        kernel32.LoadLibraryExW.restype = ctypes.c_void_p
        handle = kernel32.LoadLibraryExW(ctypes.c_wchar_p(path), None, 8)
        if not handle:
            return "Windows refused to load it (Win32 error %d)" % kernel32.GetLastError()
    except Exception:
        pass
    return None


def _revit_year():
    try:
        from pyrevit import HOST_APP
        return int(HOST_APP.version)
    except Exception:
        return None


def _runtime_assembly_version():
    """(major, minor, patch) of the loaded pyRevit runtime engine, or None."""
    try:
        from System import AppDomain
        best = None
        for asm in AppDomain.CurrentDomain.GetAssemblies():
            try:
                name = asm.GetName().Name
            except Exception:
                continue
            if name and name.startswith("pyRevitLabs.PyRevit.Runtime"):
                ver = asm.GetName().Version
                cand = (ver.Major, ver.Minor, ver.Build)
                if name[-1:].isdigit():      # ...Runtime.2024 -> the real engine
                    return cand
                if best is None:             # ...Runtime.Shared -> fallback
                    best = cand
        return best
    except Exception:
        return None


def find_runtime_mismatch():
    """(pylib_ver, dll_ver, clone) when pyRevit's Python and engine DLL differ."""
    try:
        import pyrevit
        pylib = (int(pyrevit.VERSION_MAJOR), int(pyrevit.VERSION_MINOR),
                 int(pyrevit.VERSION_PATCH))
    except Exception:
        return None
    dll = _runtime_assembly_version()
    if not dll:
        return None
    if (pylib[0], pylib[1]) == (dll[0], dll[1]):
        return None
    return (pylib, dll, _clone_root())


def find_problems():
    """[English sentence] — empty on a healthy machine."""
    problems = []
    year = _revit_year()
    lo, hi = SUPPORTED_REVIT
    if year is not None and not (lo <= year <= hi):
        problems.append("Revit %d is outside the supported range (Revit %d-%d). "
                        "Some T3Lab tools may not work." % (year, lo, hi))

    engines = _engine_dlls(_clone_root())
    if not engines:
        problems.append("This pyRevit installation has no CPython engine (bin\\cengines\\CPY*). "
                        "Every T3Lab tool needs it. Reinstall pyRevit 5 or newer.")
        return problems
    for name, dll in engines:
        if dll is None:
            problems.append("The CPython engine %s has no python3XX.dll - the pyRevit "
                            "installation is incomplete. Reinstall pyRevit." % name)
            continue
        reason = _load_error(dll)
        if reason:
            problems.append("The CPython engine %s cannot load %s: %s" %
                            (name, os.path.basename(dll), reason))
    return problems


def _same_as_last_entry(path, text):
    """True when the log already ends with this exact entry."""
    try:
        want = u"\n%s\n\n" % text
        size = os.path.getsize(path)
        span = min(size, 2 * len(want.encode("utf-8")) + 16)
        fh = open(path, "rb")
        try:
            fh.seek(size - span)
            tail = fh.read()
        finally:
            fh.close()
        tail = tail.decode("utf-8", "replace").replace(u"\r\n", u"\n")
        return tail.endswith(want)
    except Exception:
        return False


def _log(text):
    """Append one entry to %APPDATA%\\T3LabAI\\engine_check.log. Never raises."""
    try:
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        folder = os.path.join(base, "T3LabAI")
        if not os.path.isdir(folder):
            os.makedirs(folder)
        import io
        import time
        path = os.path.join(folder, LOG_NAME)
        if os.path.exists(path):
            if _same_as_last_entry(path, text):
                return
            if os.path.getsize(path) > LOG_MAX_BYTES:
                backup = path + ".1"
                try:
                    if os.path.exists(backup):
                        os.remove(backup)
                    os.rename(path, backup)
                except Exception:
                    os.remove(path)
        with io.open(path, "a", encoding="utf-8") as fh:
            fh.write(u"%s\n%s\n\n" % (time.strftime("%Y-%m-%d %H:%M:%S"), text))
    except Exception:
        pass


def _notify(problems):
    message = "T3Lab tools cannot start on this machine."
    details = "\n\n".join(problems) + (
        "\n\nUntil this is fixed, every T3Lab button fails with \"The type "
        "initializer for 'Delegates' threw an exception\".\n\n"
        "Next step: check your pyRevit CPython engines, fix them, then restart Revit.")
    _log(message + "\n" + details)
    try:
        from Autodesk.Revit.UI import TaskDialog, TaskDialogIcon
        dialog = TaskDialog(TITLE)
        dialog.MainInstruction = message
        dialog.MainContent = details
        dialog.MainIcon = TaskDialogIcon.TaskDialogIconWarning
        dialog.Show()
    except Exception:
        pass


def pythonnet_half_started():
    """True when an earlier failed CPython (re)start left pythonnet stuck."""
    try:
        from System import AppDomain
        from System.Reflection import BindingFlags
        asm = None
        for candidate in AppDomain.CurrentDomain.GetAssemblies():
            if candidate.GetName().Name == "pyRevitLabs.PythonNet":
                asm = candidate
                break
        if asm is None:
            return False
        flags = BindingFlags.NonPublic | BindingFlags.Static
        runtime_flag = asm.GetType("Python.Runtime.Runtime").GetField("_isInitialized", flags)
        engine_flag = asm.GetType("Python.Runtime.PythonEngine").GetField("initialized", flags)
        if runtime_flag is None or engine_flag is None:
            return False
        return bool(runtime_flag.GetValue(None)) and not bool(engine_flag.GetValue(None))
    except Exception:
        return False


def _notify_restart():
    message = "Restart Revit to use T3Lab tools again."
    details = ("A pyRevit reload earlier in this session (pyRevit > Reload, or "
               "enabling/disabling an extension) stopped pyRevit's CPython "
               "engine in a state it cannot recover from on Revit 2025 and "
               "newer. Until Revit restarts, every T3Lab button fails with "
               "\"This property must be set before runtime is initialized\".\n\n"
               "Save your work and restart Revit. T3Lab patches pyRevit so "
               "Reload no longer does this once Revit has restarted; also avoid "
               "Ctrl+Alt+Shift+Click on pyRevit buttons, which forces the same "
               "engine shutdown.")
    _log(message + "\n" + details)
    try:
        from Autodesk.Revit.UI import TaskDialog, TaskDialogIcon
        dialog = TaskDialog(TITLE)
        dialog.MainInstruction = message
        dialog.MainContent = details
        dialog.MainIcon = TaskDialogIcon.TaskDialogIconWarning
        dialog.Show()
    except Exception:
        pass


def _fmt_ver(ver):
    try:
        return "%d.%d.%d" % (ver[0], ver[1], ver[2])
    except Exception:
        return "unknown"


def _notify_mismatch(pylib, dll, clone):
    message = "T3Lab tools cannot start: this pyRevit install is inconsistent."
    details = (
        "pyRevit's Python library is version %s but its compiled engine is "
        "version %s, in the clone at:\n%s\n\n"
        "They are from different pyRevit releases, so pyRevit's own CPython "
        "launcher fails with \"Input string was not in a correct format\" "
        "before any T3Lab tool runs, and every '#! python3' button shows a "
        "blank \"Command Failure for External Command\".\n\n"
        "Fix: re-deploy pyRevit so both match. Close every open Revit, then "
        "update the clone (pyrevit update) or reinstall pyRevit, and start Revit again."
        % (_fmt_ver(pylib), _fmt_ver(dll), clone or "the running pyRevit clone"))
    _log(message + "\n" + details)
    try:
        from Autodesk.Revit.UI import TaskDialog, TaskDialogIcon
        dialog = TaskDialog(TITLE)
        dialog.MainInstruction = message
        dialog.MainContent = details
        dialog.MainIcon = TaskDialogIcon.TaskDialogIconError
        dialog.Show()
    except Exception:
        pass


def _data_dir():
    return os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "T3LabAI")


def patch_opted_out():
    """True when this machine's user switched the automatic pyRevit patch off."""
    if (os.environ.get(PATCH_OPT_OUT_ENV) or "").strip() not in ("", "0"):
        return True
    return os.path.isfile(os.path.join(_data_dir(), PATCH_OPT_OUT_FILE))


def _load_lib_module(name):
    """lib/<name>.py as a module, or None."""
    candidates = []
    try:
        candidates.append(_LIB_DIR)
    except Exception:
        pass
    candidates += list(sys.path)
    for lib in candidates:
        try:
            if os.path.isfile(os.path.join(lib, name + ".py")):
                if lib not in sys.path:
                    sys.path.insert(0, lib)
                break
        except Exception:
            continue
    try:
        return __import__(name)
    except Exception:
        return None


def _load_patches():
    """lib/pyrevit_patches.py, or None."""
    return _load_lib_module("pyrevit_patches")


def auto_patch_pyrevit(patches=None):
    """Scan the running pyRevit and apply the reload patch when it is missing."""
    if patch_opted_out():
        return None, "opted out"
    patches = patches or _load_patches()
    if patches is None:
        return None, "lib/pyrevit_patches.py not found"
    try:
        path = patches.sessionmgr_path()
        if not path:
            return None, "pyRevit's sessionmgr.py not found"
        return patches.ensure(path)
    except Exception as exc:
        return "error", str(exc)


def _notify_patched(detail):
    message = "T3Lab updated pyRevit so that Reload no longer breaks T3Lab tools."
    details = ("On Revit 2025 and newer, pyRevit > Reload (or enabling/disabling "
               "an extension) shut down pyRevit's CPython engine, which cannot "
               "restart there - every T3Lab button then failed with \"This "
               "property must be set before runtime is initialized\" until Revit "
               "restarted. Reload now keeps that engine running.\n\n"
               "Changed file: %s\nOriginal kept as: %s.t3lab-backup\n\n"
               "Takes effect after the next Revit restart. To undo: run "
               "lib\\pyrevit_patches.py restore, and create "
               "%%APPDATA%%\\T3LabAI\\%s to stop T3Lab from applying it again."
               % (detail, detail, PATCH_OPT_OUT_FILE))
    _log(message + "\n" + details)
    try:
        from Autodesk.Revit.UI import TaskDialog, TaskDialogIcon
        dialog = TaskDialog(TITLE)
        dialog.MainInstruction = message
        dialog.MainContent = details
        dialog.MainIcon = TaskDialogIcon.TaskDialogIconInformation
        dialog.Show()
    except Exception:
        pass


def _revit_handle():
    """pyRevit's __revit__ builtin for this startup run, or None."""
    try:
        return __revit__        # noqa: F821 - injected by pyRevit
    except NameError:
        return None


def register_assistant_pane():
    """Register the T3Lab Assistant dockable pane. Returns (state, detail)."""
    pane = _load_lib_module("assistant_pane")
    if pane is None:
        return "failed", "lib/assistant_pane.py not found"
    try:
        return pane.register(_revit_handle())
    except Exception as exc:
        return "failed", str(exc)


def main():
    # ─── 0. FIX: "The input string '3.12.3' was not in a correct format" ───────
    try:
        import startup_fix_cpyversion
        startup_fix_cpyversion.apply()
    except Exception:
        pass

    try:
        import _cpython_bootstrap
        _cpython_bootstrap.init_cpython_paths()
        _cpython_bootstrap.fix_std_streams()
    except Exception:
        pass

    # ─── 1. Retire ribbon folders an update left behind (Lite only) ────────────
    try:
        from core.ribbon_retire import retire_old_ribbon_folders, notify_retired
        notify_retired(retire_old_ribbon_folders())
    except Exception:
        pass

    # ─── 2. Dock pane registration (IronPython host in assistant_pane.py) ─────
    state, detail = register_assistant_pane()
    if state == "failed":
        _log("T3Lab Assistant dock pane not registered - the Assistant opens "
             "as a window instead: %s" % detail)

    # ─── 3. Keep pyRevit patched (Reload protection) ──────────────────────────
    state, detail = auto_patch_pyrevit()
    if state == "applied":
        _notify_patched(detail)
    elif state not in (None, "already"):
        _log("pyRevit reload patch not applied (%s): %s" % (state, detail))

    # ─── 4. Diagnostic checks on reload & startup ─────────────────────────────
    if pythonnet_half_started():
        _notify_restart()

    try:
        mismatch = find_runtime_mismatch()
    except Exception:
        mismatch = None
    if mismatch:
        _notify_mismatch(*mismatch)

    try:
        problems = find_problems()
    except Exception:
        problems = None

    if problems:
        blocking = [p for p in problems if not p.startswith("Revit ")]
        if blocking:
            _notify(problems)
        else:
            _log("\n".join(problems))



    # ─── 6. Self-study idle loop (opt-in: agents.self_study) ───────────────────
    try:
        _uictrld_idle = _revit_handle()
        if _uictrld_idle is None:
            try:
                from pyrevit import HOST_APP
                _uictrld_idle = getattr(HOST_APP, 'uicontrolledapp', None)
            except Exception:
                _uictrld_idle = None

        if _uictrld_idle is not None and hasattr(_uictrld_idle, 'add_Idling'):
            from Intelligence.learning import loop as _study_loop

            def _t3lab_on_idling(sender, args):
                try:
                    _study_loop.on_idling_tick()
                except Exception:
                    pass

            _uictrld_idle.Idling += _t3lab_on_idling
    except Exception:
        pass

    # ─── 7. Start file-based task watcher ──────────────────────────────────────
    try:
        from core.file_watcher import get_task_watcher
        get_task_watcher().start()
    except Exception:
        pass

    # ─── 8. Deploy MCP bridge + auto-start MCP server ──────────────────────────
    try:
        from Services.mcp_service import MCPService
        MCPService.deploy_bridge()

        from core import paths as _mcp_paths
        _auto = _mcp_paths.load_settings().get('auto_start_mcp')
        if _auto is None:
            _auto = True
            _mcp_paths.set_setting('auto_start_mcp', True)

        if _auto:
            MCPService.ensure_external_event()
            MCPService.start_server()
    except Exception:
        pass

    # ─── 9. Weekly auto-update from GitHub (Lite only) ─────────────────────────
    try:
        from core.updater import start_auto_update
        start_auto_update()
    except Exception:
        pass

    # ─── 10. Usage tracking: session start (Lite only) ─────────────────────────
    try:
        from tracking import track_session_start
        track_session_start()
    except Exception:
        pass


main()
