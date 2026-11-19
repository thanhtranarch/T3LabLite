# -*- coding: utf-8 -*-
"""
_cpython_bootstrap.py
=====================
Bootstraps Python 3 standard library paths and C-extension paths for CPython
running inside pyRevit.

pyRevit ships its CPython engine as ``<clone>/bin/cengines/CPY<ver>/``, holding
``python3XX.zip`` (the stdlib), an extracted ``Lib/`` and the C extensions
(``_sqlite3.pyd`` plus DLLs like ``sqlite3.dll``). Revit's embedded interpreter
does not reliably put those on ``sys.path``, which surfaces as 'No module named
configparser' / 'csv' / '_sqlite3' — and to the user as the generic
"Command Failure for External Command" dialog with no details.

Engine discovery is deliberately clone-agnostic, in order of reliability:

  1. from the running interpreter itself — ``sys.path`` / ``sys.prefix`` already
     point inside the clone, whatever it is named and wherever it was installed;
  2. from pyRevit's own environment variables;
  3. by scanning the conventional install roots for ANY folder whose name starts
     with "pyrevit".

Earlier versions hardcoded the clone names ``pyRevit-Master``/``pyRevit`` and
the engine folder ``CPY3123``. That matched exactly one machine layout — the
developer's — so a colleague with a differently named clone, another install
root, or a pyRevit build shipping a different engine (4.8.x ships CPY387) got
no stdlib injected at all and every tool died on its first import.

An engine is injected only when its Python version equals the running
interpreter's: mixing a 3.12 stdlib zip into a 3.8 engine (or a host 3.14 test
runner) raises "bad magic number", because .pyc magic differs per minor version.
"""

import os
import sys
import time


# Engine start. The persistent CPython engine keeps modules in sys.modules for
# the whole Revit session, so a lib/ file edited after this moment is stale in
# memory until it is reloaded (see reload_changed_modules).
_BOOT_TIME = time.time()

# Hot-reloaded by init_cpython_paths() when their file changes on disk.
# core.paths is plain Python (no .NET types), so a reload is safe; a stale copy
# broke MCP Control after an update without a Revit restart.
HOT_RELOAD_MODULES = ('GUI.WPF_Base', 'WPF_Base', 'GUI.forms',
                      'pyrevit.forms._cpy', 'Snippets._host', '_host',
                      'core.paths')
_RELOADED_AT = {}

# Probed by verify_stdlib(); these are the imports that actually broke tools.
_STDLIB_PROBES = ('configparser', 'csv', 'json', 'email')

# Filled in by init_cpython_paths(), read by describe_environment() and
# dev/doctor.py. A tool dying on ModuleNotFoundError can be diagnosed from this
# without digging through a Revit journal.
ENGINE_DIAGNOSIS = {
    'python': '%d.%d.%d' % sys.version_info[:3],
    'clone_roots': [],
    'engines_found': [],
    'engines_used': [],
    'stdlib_missing': [],
    'pyrevit_revit': 'not checked',     # set by heal_pyrevit_imports()
    'clr_types': 'not checked',         # set by free_clr_type_names()
    'status': 'not attempted',
}

_CLONE_ROOTS_CACHE = None
_ENGINES_CACHE = None


def _norm(p):
    """Windows-normalised path with no trailing separator. Never raises."""
    try:
        return (p or '').replace('/', '\\').rstrip('\\')
    except Exception:
        return ''


def _clone_root_from_path(entry):
    """``<clone>\\bin\\cengines\\CPY3123\\Lib`` -> ``<clone>``.

    Returns '' when the path is not inside a pyRevit clone. This is what makes
    discovery work for a clone with any name in any location: the running
    engine already carries such a path on sys.path.
    """
    norm = _norm(entry)
    low = norm.lower()
    for marker in ('\\bin\\cengines\\', '\\bin\\engines\\', '\\pyrevitlib'):
        i = low.find(marker)
        if i > 0:
            return norm[:i]
    return ''


def _candidate_clone_roots(refresh=False):
    """Every directory that might be a pyRevit clone, most reliable first."""
    global _CLONE_ROOTS_CACHE
    if _CLONE_ROOTS_CACHE is not None and not refresh:
        return _CLONE_ROOTS_CACHE

    roots = []
    seen = set()

    def add(path):
        path = _norm(path)
        if os.sep != '\\':
            # POSIX dev/test host: _norm() turned '/tmp/x' into '\\tmp\\x'.
            path = path.replace('\\', os.sep)
        if not path or path.lower() in seen:
            return
        try:
            if not os.path.isdir(path):
                return
        except Exception:
            return
        seen.add(path.lower())
        roots.append(path)

    # 1 — the running engine knows where it lives.
    probes = list(sys.path)
    for attr in ('prefix', 'base_prefix', 'exec_prefix', 'executable'):
        probes.append(getattr(sys, attr, ''))
    for entry in probes:
        add(_clone_root_from_path(entry))

    # 2 — pyRevit's own environment variables (set by the CLI / installer).
    for var in ('PYREVIT_CLONE', 'PYREVIT_HOME', 'PYREVIT_BIN'):
        val = os.environ.get(var, '')
        if val:
            add(_clone_root_from_path(val) or val)

    # 3 — conventional install roots, ANY clone folder name.
    bases = [os.environ.get(v, '') for v in
             ('APPDATA', 'PROGRAMDATA', 'LOCALAPPDATA',
              'ProgramFiles', 'ProgramFiles(x86)', 'ProgramW6432')]
    bases += ['C:\\Program Files', 'C:\\Program Files (x86)']
    for base in bases:
        if not base:
            continue
        try:
            names = os.listdir(base)
        except Exception:
            continue
        for name in names:
            if name.lower().startswith('pyrevit'):
                add(os.path.join(base, name))

    _CLONE_ROOTS_CACHE = roots
    return roots


def _engine_version(engine_dir, engine_name):
    """``(major, minor)`` for a cengines folder, or None.

    The stdlib zip is authoritative (``python312.zip`` -> 3.12); the folder name
    is the fallback (``CPY3123`` -> 3.12, ``CPY387`` -> 3.8).
    """
    try:
        for fn in os.listdir(engine_dir):
            low = fn.lower()
            if low.startswith('python') and low.endswith('.zip'):
                digits = ''.join(c for c in low[6:-4] if c.isdigit())
                if len(digits) >= 2:
                    return (int(digits[0]), int(digits[1:]))
    except Exception:
        pass
    digits = ''.join(c for c in (engine_name or '') if c.isdigit())
    if len(digits) >= 2:
        rest = digits[1:]
        # 'CPY3123' -> 3.12.3 (3 digits left); 'CPY387' -> 3.8.7 (2 digits left)
        minor = int(rest[:2]) if len(rest) >= 3 else int(rest[0])
        return (int(digits[0]), minor)
    return None


def find_cpython_engines(refresh=False):
    """``[(engine_dir, (major, minor))]`` for every CPython engine found."""
    global _ENGINES_CACHE
    if _ENGINES_CACHE is not None and not refresh:
        return _ENGINES_CACHE

    found = []
    seen = set()
    for root in _candidate_clone_roots(refresh=refresh):
        ceng = os.path.join(root, 'bin', 'cengines')
        try:
            names = sorted(os.listdir(ceng))
        except Exception:
            continue
        for name in names:
            if not name.upper().startswith('CPY'):
                continue
            d = os.path.join(ceng, name)
            key = d.lower()
            if key in seen:
                continue
            try:
                if not os.path.isdir(d):
                    continue
            except Exception:
                continue
            seen.add(key)
            ver = _engine_version(d, name)
            if ver:
                found.append((d, ver))

    _ENGINES_CACHE = found
    return found


def verify_stdlib():
    """Return the stdlib modules that still cannot be imported (usually [])."""
    missing = []
    for mod in _STDLIB_PROBES:
        if mod in sys.modules:
            continue
        try:
            __import__(mod)
        except Exception:
            missing.append(mod)
    ENGINE_DIAGNOSIS['stdlib_missing'] = missing
    return missing


def heal_pyrevit_imports():
    """Make `from pyrevit import revit` findable again when only the import
    caches lost it.

    "cannot import name 'revit' from 'pyrevit' (...pyrevit\\__init__.py)"
    means Python could not FIND pyrevit/revit/ — not that its code failed.
    The persistent CPython engine keeps sys.path_importer_cache and each
    folder's cached file list for the whole Revit session (and across a
    pyRevit Reload, see pyrevit_patches.py), so one bad lookup used to break
    every tool until Revit restarted. When the package is on disk but cannot
    be found, the caches are dropped; Python rebuilds them on the next import.

    Returns 'loaded', 'ok', 'healed', 'missing' (not on disk: pyRevit itself
    needs a repair), 'unresolved' (on disk, still not importable) or
    'skipped'.
    """
    status = 'skipped'
    if 'pyrevit.revit' in sys.modules:
        status = 'loaded'
    else:
        pkg_dirs = list(getattr(sys.modules.get('pyrevit'), '__path__', None) or [])
        if pkg_dirs:
            try:
                import importlib
                import importlib.util
                if importlib.util.find_spec('pyrevit.revit') is not None:
                    status = 'ok'
                elif not any(os.path.isfile(os.path.join(d, 'revit', '__init__.py'))
                             for d in pkg_dirs):
                    status = 'missing'
                else:
                    importlib.invalidate_caches()
                    for d in pkg_dirs:
                        sys.path_importer_cache.pop(d, None)
                    found = importlib.util.find_spec('pyrevit.revit') is not None
                    status = 'healed' if found else 'unresolved'
            except Exception:
                status = 'skipped'
    ENGINE_DIAGNOSIS['pyrevit_revit'] = status
    return status


def describe_environment():
    """Human-readable bootstrap report — English, safe to show in a dialog."""
    d = ENGINE_DIAGNOSIS
    lines = [
        'T3Lab CPython bootstrap',
        '  Interpreter      : Python %s' % d.get('python', '?'),
        '  Status           : %s' % d.get('status', '?'),
        '  pyRevit clones   : %s' % (', '.join(d.get('clone_roots') or []) or '(none found)'),
        '  Engines found    : %s' % (', '.join(d.get('engines_found') or []) or '(none found)'),
        '  Engines injected : %s' % (', '.join(d.get('engines_used') or []) or '(none)'),
    ]
    missing = d.get('stdlib_missing') or []
    if missing:
        lines.append('  MISSING stdlib   : %s' % ', '.join(missing))
        lines.append('  Fix: install or repair the pyRevit CPython engine, then '
                     'pyRevit > Reload. See INSTALL.md.')
    revit_pkg = d.get('pyrevit_revit')
    if revit_pkg == 'missing':
        lines.append('  BROKEN pyRevit   : pyrevitlib\\pyrevit\\revit is missing - '
                     'repair or reinstall pyRevit.')
    elif revit_pkg == 'unresolved':
        lines.append('  BROKEN pyRevit   : pyrevit.revit is on disk but cannot be '
                     'imported - restart Revit.')
    clr_types = d.get('clr_types') or ''
    if clr_types.startswith(('freed', 'failed')):
        lines.append('  CPython restarted: %s' % clr_types)
        lines.append('  Cause: pyRevit > Reload, an extension switched off and on, '
                     'or an update applied while Revit was open. Restart Revit '
                     'after an update instead of reloading.')
    return '\n'.join(lines)


_BOOTSTRAP_LOG_MAX_BYTES = 256 * 1024     # then rotated to bootstrap.log.1
_LOGGED_REPORTS = set()                   # reports already written this session


def _log_diagnosis():
    """Append the report to %APPDATA%/T3LabAI/bootstrap.log when something is
    wrong. Silent on a healthy machine, so the log stays a signal.

    Written once per state: init_cpython_paths() runs on EVERY click, and a
    machine with a broken engine used to append the same full report each
    time. A report identical to the one already at the end of the log (or
    already written in this session) is skipped; past the size cap the log
    is rotated to bootstrap.log.1.
    """
    if not (ENGINE_DIAGNOSIS.get('stdlib_missing')
            or ENGINE_DIAGNOSIS.get('pyrevit_revit') in ('missing', 'unresolved')
            or (ENGINE_DIAGNOSIS.get('clr_types') or '').startswith(('freed', 'failed'))):
        return
    report = describe_environment()
    if report in _LOGGED_REPORTS:
        return
    _LOGGED_REPORTS.add(report)
    try:
        base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
        d = os.path.join(base, 'T3LabAI')
        if not os.path.isdir(d):
            os.makedirs(d)
        path = os.path.join(d, 'bootstrap.log')
        entry = report + u'\n\n'
        if os.path.exists(path):
            size = os.path.getsize(path)
            raw = entry.encode('utf-8')
            # Only the tail is read, in binary; io.open wrote \r\n on Windows,
            # so line endings are normalised before comparing.
            with open(path, 'rb') as f:
                f.seek(max(0, size - 2 * len(raw) - 16))
                tail = f.read().replace(b'\r\n', b'\n')
            if tail.endswith(raw):
                return
            if size > _BOOTSTRAP_LOG_MAX_BYTES:
                try:
                    os.replace(path, path + '.1')
                except Exception:
                    os.remove(path)
        import io as _io
        with _io.open(path, 'a', encoding='utf-8') as f:
            f.write(entry)
    except Exception:
        pass


def _source_mtime(module):
    path = getattr(module, '__file__', None) or ''
    if path.endswith(('.pyc', '.pyo')):
        path = path[:-1]
    try:
        return os.path.getmtime(path) if path else None
    except Exception:
        return None


def reload_changed_modules(names=HOT_RELOAD_MODULES):
    """Reload only the cached modules whose source changed since they loaded.

    Every pushbutton calls init_cpython_paths(), and it used to reload all of
    HOT_RELOAD_MODULES unconditionally — recompiling WPF_Base (1400+ lines,
    ten clr.AddReference probes) on every click of every tool, three times on
    the first BatchOut open. An unchanged file now costs one stat() call; an
    edited one is still picked up without a pyRevit Reload.
    """
    import importlib
    reloaded = []
    for name in names:
        module = sys.modules.get(name)
        if module is None:
            continue
        mtime = _source_mtime(module)
        if mtime is None or mtime <= _RELOADED_AT.get(name, _BOOT_TIME):
            continue
        try:
            importlib.reload(module)
            reloaded.append(name)
        except Exception:
            pass
        _RELOADED_AT[name] = mtime
    return reloaded


def _timing(event):
    """open_timing.<event>(); never raises (the log is opt-in, off by default)."""
    try:
        import open_timing
        getattr(open_timing, event)()
    except Exception:
        pass


# ─── "Duplicate type name within an assembly" ────────────────────────────────
# Every Python class that inherits a .NET type (IExternalEventHandler,
# ISelectionFilter, IFailuresPreprocessor, ...) becomes a real .NET type, which
# pythonnet emits into a dynamic assembly called "Python.Runtime.Dynamic".
# pythonnet 3.0.3 keeps the builders for it in two static dictionaries filled by
# `GetModuleBuilder` and created in the STATIC CONSTRUCTOR of
# Python.Runtime.ClassDerivedObject (verified against pythonnet v3.0.3,
# src/runtime/Types/ClassDerived.cs: `assemblyBuilders`, `moduleBuilders`,
# emptied only by `Reset()`). A static constructor runs once per .NET AppDomain,
# i.e. once per Revit session, and nothing empties those dictionaries when the
# Python runtime shuts down.
#
# So when CPython restarts inside one Revit session - pyRevit > Reload, an
# extension switched off and on, an update applied while Revit was open -
# sys.modules starts empty while the dynamic assembly still holds every type
# from before. The first class re-created under a fixed name then dies, and
# pyRevit's own `pyrevit.revit.events.FuncAsEventHandler` is the first one in
# the way: `from pyrevit import revit / script / forms` fails for EVERY tool,
# T3Lab's and pyRevit's alike, with "TypeError: Duplicate type name within an
# assembly" (or "cannot import name 'revit' from 'pyrevit'"). Until now only a
# Revit restart cleared it (GUI/ErrorGuard.py says so).
#
# Emptying the two dictionaries hands pythonnet a fresh dynamic assembly: the
# names are free again, and the types emitted before stay valid - they are real
# .NET types and nothing is unloaded. Done once per runtime, before the first
# `import pyrevit...`, and only while the old types are provably unreachable
# (sys.modules no longer holds the module that defined them).
_DERIVED_CLASS = 'Python.Runtime.ClassDerivedObject'
_BUILDER_FIELDS = ('assemblyBuilders', 'moduleBuilders')
# Loaded, these hold the .NET types a restart would try to emit a second time.
_PYREVIT_CLR_MODULES = ('pyrevit.revit.events', 'pyrevit.revit', 'pyrevit.forms')
# How the leftovers read from Python. The second one is the same failure seen
# one import later: pyrevit.revit died on the first, so `pyrevit` never got the
# attribute. GUI/ErrorGuard.py shows the user the same two.
STALE_TYPE_MARKERS = ('Duplicate type name within an assembly',
                      "cannot import name 'revit' from 'pyrevit'")

CLR_TYPE_RESET_STATUS = 'not checked'
_CLR_TYPES_CHECKED = False


def stale_clr_types(exc):
    """True when this failure is a .NET type a past CPython runtime left here."""
    try:
        text = u'{}'.format(exc)
    except Exception:
        return False
    return any(marker in text for marker in STALE_TYPE_MARKERS)


def _clr_derived_class():
    """Python.Runtime.ClassDerivedObject - internal, so found by reflection."""
    from System import AppDomain
    for asm in AppDomain.CurrentDomain.GetAssemblies():
        try:
            found = asm.GetType(_DERIVED_CLASS)
        except Exception:
            continue
        if found is not None:
            return found
    return None


def _emitted_clr_types():
    """How many .NET types pythonnet emitted for Python subclasses this session.

    None means "do not touch": a pythonnet build whose fields are not the ones
    this patch was written against.
    """
    from System.Reflection import BindingFlags
    owner = _clr_derived_class()
    if owner is None:
        return None
    field = owner.GetField(
        'moduleBuilders',
        BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static)
    if field is None:
        return None
    builders = field.GetValue(None)
    if builders is None:
        return 0
    count = 0
    for entry in builders:
        try:
            count += entry.Value.GetTypes().Length
        except Exception:
            count += 1          # a type still under construction holds its name
    return count


def _reset_clr_type_builders():
    """Give pythonnet a new dynamic assembly. Returns how it was done, or ''."""
    from System import Activator, Array, Object
    from System.Reflection import BindingFlags
    flags = BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Static
    owner = _clr_derived_class()
    if owner is None:
        return ''
    reset = owner.GetMethod('Reset', flags)
    if reset is not None:
        try:
            reset.Invoke(None, Array[Object]([]))
        except Exception:
            reset.Invoke(None, None)
        return 'Reset()'
    done = []
    for name in _BUILDER_FIELDS:
        field = owner.GetField(name, flags)
        if field is None:
            continue
        field.SetValue(None, Activator.CreateInstance(field.FieldType))
        done.append(name)
    return ', '.join(done)


def free_clr_type_names(modules=None, emitted=None, reset=None, once=True):
    """Clear the .NET types a previous CPython runtime left in this session.

    Called at the top of init_cpython_paths(), so it runs before any tool
    imports pyrevit, and at most once per Python runtime (a restarted runtime
    has a fresh module, so the flag starts over with it - exactly when the work
    is needed). Never raises; returns the status, also in CLR_TYPE_RESET_STATUS.
    """
    global _CLR_TYPES_CHECKED

    def _done(status):
        global CLR_TYPE_RESET_STATUS
        CLR_TYPE_RESET_STATUS = status
        ENGINE_DIAGNOSIS['clr_types'] = status
        return status

    if sys.version_info[0] < 3:
        return _done('skipped: IronPython')
    if once:
        if _CLR_TYPES_CHECKED:
            return CLR_TYPE_RESET_STATUS
        _CLR_TYPES_CHECKED = True
    loaded = sys.modules if modules is None else modules
    still_loaded = [m for m in _PYREVIT_CLR_MODULES if m in loaded]
    if still_loaded:
        # Same runtime as the types: importing them again is a no-op, and the
        # live objects behind them must keep the .NET types they were built on.
        return _done('clean: %s loaded' % still_loaded[0])
    try:
        count = (emitted or _emitted_clr_types)()
    except Exception as exc:
        return _done('skipped: %s: %s' % (type(exc).__name__, exc))
    if count is None:
        return _done('skipped: pythonnet builders not found')
    if not count:
        return _done('clean: no .NET types emitted yet')
    try:
        how = (reset or _reset_clr_type_builders)()
    except Exception as exc:
        return _done('failed: %s: %s' % (type(exc).__name__, exc))
    if not how:
        return _done('failed: pythonnet builders not found')
    # The runtime restarted in this Revit session (pyRevit Reload, an extension
    # switched off and on, an update while Revit was open).
    return _done('freed %d stale .NET type name(s) via %s' % (count, how))


def _patch_pyrevit_forms():
    """Point pyrevit.forms at T3WPFWindow — pyRevit's own raises under CPython.

    Raises what the pyrevit import raises: init_cpython_paths() reads it to
    tell a restarted runtime from an ordinary failure.
    """
    reload_changed_modules()
    from GUI.WPF_Base import T3WPFWindow
    import pyrevit.forms as _pyrevit_forms
    _pyrevit_forms.WPFWindow = T3WPFWindow

    class _CPythonReactive(object):
        def OnPropertyChanged(self, *args, **kwargs):
            pass

    _pyrevit_forms.Reactive = _CPythonReactive
    try:
        import pyrevit.forms._cpy as _cpy
        _cpy.WPFWindow = T3WPFWindow
        _cpy.Reactive = _CPythonReactive
    except Exception:
        pass


def init_cpython_paths():
    """Ensure pyRevit CPython standard library and C extensions are in sys.path and DLL path."""
    sys.dont_write_bytecode = True
    _timing('run_started')      # opt-in open timing log (lib/open_timing.py)
    # Before ANY pyrevit import below (and before the tool's own): a restarted
    # CPython runtime would otherwise fail on pyrevit's FuncAsEventHandler.
    try:
        free_clr_type_names()
    except Exception:
        pass
    engine_dirs = []
    candidates = []
    running = sys.version_info[:2]

    all_engines = find_cpython_engines()
    ENGINE_DIAGNOSIS['clone_roots'] = list(_candidate_clone_roots())
    ENGINE_DIAGNOSIS['engines_found'] = ['%s (%d.%d)' % (p, v[0], v[1])
                                         for p, v in all_engines]

    for eng_dir, ver in all_engines:
        # Version gate: a 3.12 stdlib zip inside a 3.8 interpreter is worse than
        # nothing ("bad magic number in 'json'").
        if ver != running:
            continue
        engine_dirs.append(eng_dir)
        candidates.append(eng_dir)                                   # .pyd / .dll
        candidates.append(os.path.join(eng_dir, 'Lib'))              # extracted stdlib
        candidates.append(os.path.join(eng_dir, 'Lib', 'site-packages'))
        try:
            for fn in os.listdir(eng_dir):                           # python3XX.zip
                low = fn.lower()
                if low.startswith('python') and low.endswith('.zip'):
                    candidates.append(os.path.join(eng_dir, fn))
        except Exception:
            pass

    ENGINE_DIAGNOSIS['engines_used'] = list(engine_dirs)
    if engine_dirs:
        ENGINE_DIAGNOSIS['status'] = 'engine matched'
    elif all_engines:
        ENGINE_DIAGNOSIS['status'] = (
            'no engine matches Python %d.%d - pyRevit ships %s'
            % (running[0], running[1],
               ', '.join(sorted(set('%d.%d' % v for _, v in all_engines)))))
    else:
        ENGINE_DIAGNOSIS['status'] = 'no pyRevit CPython engine found'

    # Configure DLL search path for C extensions (e.g. sqlite3.dll, libffi-8.dll, libssl-3.dll)
    for ed in engine_dirs:
        for _dd in (ed, os.path.join(ed, 'Lib')):
            if hasattr(os, 'add_dll_directory') and os.path.isdir(_dd):
                try:
                    os.add_dll_directory(_dd)
                except Exception:
                    pass
        path_env = os.environ.get('PATH', '')
        _lib_d = os.path.join(ed, 'Lib')
        if ed not in path_env:
            os.environ['PATH'] = ed + os.pathsep + _lib_d + os.pathsep + path_env

    # Also make sure this extension's lib directory is in sys.path
    ext_lib = os.path.dirname(os.path.abspath(__file__))
    if ext_lib not in sys.path:
        sys.path.insert(0, ext_lib)

    # Prepend existing candidates to sys.path
    for c in reversed(candidates):
        if os.path.exists(c) and c not in sys.path:
            sys.path.insert(0, c)

    verify_stdlib()

    # Pick up edited base modules, and monkeypatch forms.WPFWindow.
    #
    # This is the first `import pyrevit...` of the click, so it is also where a
    # CPython runtime that restarted in this Revit session shows itself: free
    # the .NET type names it left behind and try once more. Without that, the
    # tool's own `from pyrevit import script` a few lines later dies with
    # "Duplicate type name within an assembly" and the tool is dead until Revit
    # restarts.
    for _attempt in (1, 2):
        try:
            _patch_pyrevit_forms()
            break
        except Exception as exc:
            if _attempt == 2 or not stale_clr_types(exc):
                break
            free_clr_type_names(modules={}, once=False)

    # pyrevit is imported by now (block above): recover a pyrevit.revit that
    # only the import caches lost, then log whatever is still broken.
    heal_pyrevit_imports()
    _log_diagnosis()

    # Scripts call init_cpython_paths() directly, so install the CPython API
    # shims here too rather than only at module import.
    for _installer in (fix_std_streams, install_forms_shim, install_script_shim,
                       install_imp_shim, install_wpf_shim, enable_safe_engine_shutdown):
        try:
            _installer()
        except Exception:
            pass

    # Usage tracking (Lite only, code in lib/tracking/). Every T3Lab.tab
    # script.py calls init_cpython_paths(), so this is the one place that
    # sees every ribbon click. 5751bd2 dropped it once when copying from
    # t3lab_dev -- lite_guard/manifest.json now fails the commit if it goes.
    try:
        from tracking import track_ribbon_click
        track_ribbon_click()
    except Exception:
        pass
    _timing('bootstrap_done')


class _NullStream(object):
    """Stand-in for a stdout/stderr that has no write() (pyRevit ScriptIO)."""

    def write(self, *_a, **_k):
        return 0

    def flush(self, *_a, **_k):
        pass

    def isatty(self):
        return False


class _STASafeStream(object):
    """Wraps sys.stdout / sys.stderr so background threads (MTA) never touch
    pyRevit's ScriptIO / ScriptOutput window.

    In pyRevit, ScriptIO.Write attempts to get or create a WPF MetroWindow
    output window. Calling Window..ctor() from a non-STA (background) thread
    causes WPF to throw System.InvalidOperationException:
        'The calling thread must be STA, because many UI components require this.'
    which immediately crashes Revit.exe.

    This wrapper detects whether the current thread is an STA thread. If not,
    it absorbs the write safely without calling ScriptIO, protecting Revit
    from fatal background thread crashes.
    """

    def __init__(self, target):
        self._target = target

    def write(self, s):
        try:
            import System.Threading
            if System.Threading.Thread.CurrentThread.GetApartmentState() == System.Threading.ApartmentState.STA:
                if hasattr(self._target, 'write'):
                    return self._target.write(s)
            return len(s) if s else 0
        except Exception:
            # Outside .NET (e.g. pure Python dev test runners), forward normally
            try:
                if hasattr(self._target, 'write'):
                    return self._target.write(s)
            except Exception:
                pass
            return len(s) if s else 0

    def flush(self):
        try:
            import System.Threading
            if System.Threading.Thread.CurrentThread.GetApartmentState() == System.Threading.ApartmentState.STA:
                if hasattr(self._target, 'flush'):
                    return self._target.flush()
        except Exception:
            try:
                if hasattr(self._target, 'flush'):
                    return self._target.flush()
            except Exception:
                pass

    def isatty(self):
        return False

    def writelines(self, lines):
        for line in lines:
            self.write(line)

    def __getattr__(self, name):
        return getattr(self._target, name)


def fix_std_streams():
    """Make `print()` and standard streams STA-safe and crash-proof.

    Under CPython, pyRevit ScriptIO can have no `write` method.
    Under IronPython (and CPython), pyRevit ScriptIO tries to instantiate
    a WPF MetroWindow on write(). Calling this from any background thread
    (MTA) throws `InvalidOperationException: The calling thread must be STA`,
    instantly crashing Revit.exe.
    This installs a stream wrapper that only forwards writes on STA threads.
    """
    for name in ('stdout', 'stderr'):
        try:
            stream = getattr(sys, name, None)
            if isinstance(stream, _STASafeStream):
                continue
            if stream is None:
                setattr(sys, name, _STASafeStream(_NullStream()))
            else:
                setattr(sys, name, _STASafeStream(stream))
        except Exception:
            pass


def _t3_alert(msg, title=None, sub_msg=None, expanded=None, footer='',
              ok=True, cancel=False, yes=False, no=False, retry=False,
              warn_icon=True, options=None, exitscript=False, **kwargs):
    """CPython-safe stand-in for `pyrevit.forms.alert`, backed by GUI.T3Dialog."""
    caption = title or 'T3Lab'
    details = expanded or sub_msg or footer or None
    labels = [u'{}'.format(o) for o in (options or ())]
    # With `options`, pyRevit returns the label of the button picked (None
    # when the dialog is closed). Returning True there made every
    # `if res == "Update now":` false, so Check Update never updated.
    answer = None if labels else True
    try:
        from GUI.T3Dialog import show_info, show_warning, confirm
        if len(labels) == 1:
            (show_warning if warn_icon else show_info)(msg, title=caption, details=details)
            answer = labels[0]
        elif labels:
            # T3Dialog has two buttons: the first two options. No caller passes more.
            answer = labels[0] if confirm(msg, title=caption, details=details,
                                          ok_text=labels[0],
                                          cancel_text=labels[1]) else labels[1]
        elif yes or no or cancel:
            # Label the buttons the way the caller asked the question: a
            # yes/no prompt answered by "Proceed"/"Cancel" reads as a different
            # question than the one on screen.
            ok_text = 'Yes' if (yes or no) else 'OK'
            cancel_text = 'No' if (yes or no) else 'Cancel'
            answer = bool(confirm(msg, title=caption, details=details,
                                  ok_text=ok_text, cancel_text=cancel_text))
        elif warn_icon:
            show_warning(msg, title=caption, details=details)
        else:
            show_info(msg, title=caption, details=details)
    except Exception:
        # Last resort: Revit's own dialog, so a message still reaches the user.
        try:
            from Autodesk.Revit.UI import TaskDialog
            TaskDialog.Show(caption, msg if not details
                            else u"{}\n\n{}".format(msg, details))
        except Exception:
            pass
    if exitscript:
        sys.exit(0)
    return answer


def _t3_pick_file(file_ext='', files_filter='', init_dir='', title=None,
                  multi_file=False, unc_paths=False, **kwargs):
    from Microsoft.Win32 import OpenFileDialog
    dlg = OpenFileDialog()
    dlg.Filter = files_filter or (
        "{0} files|*.{0}|All files|*.*".format(file_ext) if file_ext else "All files|*.*")
    dlg.Multiselect = bool(multi_file)
    if init_dir:
        dlg.InitialDirectory = init_dir
    if title:
        dlg.Title = title
    if dlg.ShowDialog() != True:  # noqa: E712 - Nullable<bool> from WPF
        return None
    return list(dlg.FileNames) if multi_file else dlg.FileName


def _t3_save_file(file_ext='', files_filter='', init_dir='', default_name='',
                  title=None, **kwargs):
    from Microsoft.Win32 import SaveFileDialog
    dlg = SaveFileDialog()
    dlg.Filter = files_filter or (
        "{0} files|*.{0}|All files|*.*".format(file_ext) if file_ext else "All files|*.*")
    if file_ext:
        dlg.DefaultExt = file_ext
    if default_name:
        dlg.FileName = default_name
    if init_dir:
        dlg.InitialDirectory = init_dir
    if title:
        dlg.Title = title
    return dlg.FileName if dlg.ShowDialog() == True else None  # noqa: E712


def _t3_pick_folder(title=None, owner=None, **kwargs):
    try:
        from System.Windows.Forms import FolderBrowserDialog, DialogResult
        dlg = FolderBrowserDialog()
        if title:
            dlg.Description = title
        return dlg.SelectedPath if dlg.ShowDialog() == DialogResult.OK else None
    except Exception:
        return None


def _wpf_list_dialog(items, title='Select', multiselect=False,
                     button_name='Select', name_attr=None, prompt=None):
    """Minimal WPF list picker built in code (no XAML dependency).

    Returns the ORIGINAL objects, not their labels, so callers that pass
    dict keys or Revit elements get their objects back.
    """
    from System.Windows import (Window, WindowStartupLocation, Thickness,
                                GridLength, GridUnitType, HorizontalAlignment)
    from System.Windows.Controls import (Grid, RowDefinition, ListBox, Button,
                                         StackPanel, Orientation, SelectionMode,
                                         TextBlock)

    entries = list(items or [])

    def label(obj):
        if name_attr:
            return u'{}'.format(getattr(obj, name_attr, obj))
        return u'{}'.format(obj)

    win = Window()
    win.Title = title or 'Select'
    win.Width = 420
    win.Height = 480
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen

    grid = Grid()
    grid.Margin = Thickness(12)
    for height in (GridLength(1, GridUnitType.Auto),
                   GridLength(1, GridUnitType.Star),
                   GridLength(1, GridUnitType.Auto)):
        rd = RowDefinition()
        rd.Height = height
        grid.RowDefinitions.Add(rd)

    caption = TextBlock()
    caption.Text = prompt or title or 'Select'
    caption.Margin = Thickness(0, 0, 0, 8)
    Grid.SetRow(caption, 0)
    grid.Children.Add(caption)

    box = ListBox()
    if multiselect:
        box.SelectionMode = SelectionMode.Extended
    for obj in entries:
        box.Items.Add(label(obj))
    Grid.SetRow(box, 1)
    grid.Children.Add(box)

    bar = StackPanel()
    bar.Orientation = Orientation.Horizontal
    bar.Margin = Thickness(0, 8, 0, 0)
    bar.HorizontalAlignment = HorizontalAlignment.Right

    state = {'ok': False}

    def _accept(_s, _e):
        state['ok'] = True
        win.DialogResult = True

    ok_btn = Button()
    ok_btn.Content = button_name or 'Select'
    ok_btn.Width = 96
    ok_btn.IsDefault = True
    ok_btn.Margin = Thickness(0, 0, 8, 0)
    ok_btn.Click += _accept
    bar.Children.Add(ok_btn)

    cancel_btn = Button()
    cancel_btn.Content = 'Cancel'
    cancel_btn.Width = 96
    cancel_btn.IsCancel = True
    bar.Children.Add(cancel_btn)

    Grid.SetRow(bar, 2)
    grid.Children.Add(bar)
    win.Content = grid
    win.ShowDialog()

    if not state['ok']:
        return [] if multiselect else None
    picked = [entries[i] for i in range(len(entries))
              if box.SelectedItems.Contains(label(entries[i]))]
    if multiselect:
        return picked
    return picked[0] if picked else None


class _T3SelectFromList(object):
    """CPython stand-in for `pyrevit.forms.SelectFromList`."""

    @classmethod
    def show(cls, context, title=None, button_name='Select', multiselect=False,
             name_attr=None, **kwargs):
        try:
            return _wpf_list_dialog(context, title=title or 'Select',
                                    multiselect=multiselect,
                                    button_name=button_name, name_attr=name_attr)
        except Exception:
            return [] if multiselect else None


class _T3CommandSwitchWindow(object):
    """CPython stand-in for `pyrevit.forms.CommandSwitchWindow` (option picker)."""

    @classmethod
    def show(cls, context, message=None, title=None, **kwargs):
        try:
            return _wpf_list_dialog(context, title=title or message or 'Select',
                                    prompt=message, button_name='OK')
        except Exception:
            return None


def _t3_ask_for_string(default='', prompt=None, title=None, **kwargs):
    """CPython stand-in for `pyrevit.forms.ask_for_string`."""
    from System.Windows import (Window, WindowStartupLocation, Thickness,
                                GridLength, GridUnitType, HorizontalAlignment,
                                SizeToContent)
    from System.Windows.Controls import (Grid, RowDefinition, TextBox, Button,
                                         StackPanel, Orientation, TextBlock)
    try:
        win = Window()
        win.Title = title or 'Input'
        win.Width = 420
        win.SizeToContent = SizeToContent.Height
        win.WindowStartupLocation = WindowStartupLocation.CenterScreen

        grid = Grid()
        grid.Margin = Thickness(12)
        for _ in range(3):
            rd = RowDefinition()
            rd.Height = GridLength(1, GridUnitType.Auto)
            grid.RowDefinitions.Add(rd)

        caption = TextBlock()
        caption.Text = prompt or 'Enter a value:'
        caption.Margin = Thickness(0, 0, 0, 8)
        Grid.SetRow(caption, 0)
        grid.Children.Add(caption)

        entry = TextBox()
        entry.Text = u'{}'.format(default or '')
        entry.SelectAll()
        Grid.SetRow(entry, 1)
        grid.Children.Add(entry)

        bar = StackPanel()
        bar.Orientation = Orientation.Horizontal
        bar.HorizontalAlignment = HorizontalAlignment.Right
        bar.Margin = Thickness(0, 12, 0, 0)
        state = {'ok': False}

        def _accept(_s, _e):
            state['ok'] = True
            win.DialogResult = True

        ok_btn = Button()
        ok_btn.Content = 'OK'
        ok_btn.Width = 96
        ok_btn.IsDefault = True
        ok_btn.Margin = Thickness(0, 0, 8, 0)
        ok_btn.Click += _accept
        bar.Children.Add(ok_btn)

        cancel_btn = Button()
        cancel_btn.Content = 'Cancel'
        cancel_btn.Width = 96
        cancel_btn.IsCancel = True
        bar.Children.Add(cancel_btn)

        Grid.SetRow(bar, 2)
        grid.Children.Add(bar)
        win.Content = grid
        entry.Focus()
        win.ShowDialog()
        return entry.Text if state['ok'] else None
    except Exception:
        return None


class _T3ProgressBar(object):
    """No-op context manager standing in for `pyrevit.forms.ProgressBar`."""

    def __init__(self, *_args, **_kwargs):
        self.cancelled = False
        self.title = _kwargs.get('title', '')

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        return False

    def update_progress(self, *_a, **_k):
        pass

    def reset(self, *_a, **_k):
        pass


class _T3WarningBar(object):
    """No-op context manager standing in for `pyrevit.forms.WarningBar`."""

    def __init__(self, *_args, **_kwargs):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *_exc):
        return False

    def Close(self):
        pass


def install_forms_shim():
    """Fill in the `pyrevit.forms` API that the CPython engine refuses to serve.

    `pyrevit/forms/__init__.py` defines a module `__getattr__` that raises
    PyRevitCPythonNotSupported for EVERY attribute, so `forms.alert` and friends
    are fatal under CPython — ~390 call sites in this extension. Binding real
    attributes on the module means normal lookup succeeds and `__getattr__` is
    never consulted, so existing call sites work unchanged and even
    already-imported (cached) modules are fixed.

    Only fills genuine gaps: anything the engine already provides is left alone.
    Newer pyRevit builds also ship `pyrevit/forms/_cpy.py`, whose
    `ask_for_string`, `pick_file`, `pick_folder`, `SelectFromList`,
    `CommandSwitchWindow` and `ProgressBar` EXIST but only raise
    PyRevitCPythonNotSupported when called — those count as gaps too.
    """
    try:
        import pyrevit.forms as _forms
    except Exception:
        return
    if getattr(_forms, 'IRONPY', False):
        return          # IronPython has the real implementations

    def _missing(name):
        try:
            attr = getattr(_forms, name)
        except Exception:
            return True
        # A stub from pyRevit's CPython backend is not an implementation.
        return getattr(attr, '__module__', '') == 'pyrevit.forms._cpy'

    for name, impl in (('alert', _t3_alert),
                       ('pick_file', _t3_pick_file),
                       ('save_file', _t3_save_file),
                       ('pick_folder', _t3_pick_folder),
                       ('SelectFromList', _T3SelectFromList),
                       ('CommandSwitchWindow', _T3CommandSwitchWindow),
                       ('ask_for_string', _t3_ask_for_string),
                       ('ProgressBar', _T3ProgressBar),
                       ('WarningBar', _T3WarningBar),
                       ('MessageBox', _t3_alert),
                       ('toast', lambda *a, **k: None),
                       ('to_items_source', lambda items: _forms_to_items_source(items)),
                       ('set_items_source', lambda ctrl, items: _forms_set_items_source(ctrl, items))):
        if _missing(name):
            try:
                setattr(_forms, name, impl)
            except Exception:
                pass


def _forms_to_items_source(items):
    try:
        from GUI.WPF_Base import to_items_source
        return to_items_source(items)
    except Exception:
        return items


def _forms_set_items_source(ctrl, items):
    try:
        from GUI.WPF_Base import set_items_source
        set_items_source(ctrl, items)
    except Exception:
        if ctrl is not None:
            ctrl.ItemsSource = items


class _NullOutput(object):
    """No-op stand-in for pyRevit's output window.

    The CPython engine has no ScriptOutput (`ScriptOutput.GetDefault` is
    missing), so `script.get_output()` raises and takes the tool with it. Every
    method here is a silent no-op: `print_md`, `print_table`, `close_others`,
    `set_width`, `self_destruct` and friends (71 `output.print_md` calls in this
    extension) become harmless instead of fatal.
    """

    def __getattr__(self, _name):
        def _noop(*_args, **_kwargs):
            return None
        return _noop


SAFE_SHUTDOWN_STATUS = 'not attempted'


def _formatter_interface_type():
    """The exact IFormatter type pythonnet itself uses.

    Taken from RuntimeData.CreateFormatter's return type rather than from a
    plain import: under .NET 8 an assembly can be loaded into more than one
    AssemblyLoadContext, and two IFormatter types with the same full name but
    different contexts are NOT assignable to each other — the FormatterType
    setter would reject our type with a message that says nothing about why.
    Falls back to the plain import when the internal method cannot be found.
    """
    import clr
    from System.Reflection import BindingFlags
    try:
        from Python.Runtime import RuntimeData
        method = clr.GetClrType(RuntimeData).GetMethod(
            'CreateFormatter',
            BindingFlags.NonPublic | BindingFlags.Static | BindingFlags.Public)
        if method is not None and method.ReturnType.IsInterface:
            return method.ReturnType
    except Exception:
        pass
    from System.Runtime.Serialization import IFormatter
    return clr.GetClrType(IFormatter)


def _build_noop_formatter_type():
    """Emit a .NET type implementing IFormatter whose Serialize does nothing.

    Reflection.Emit rather than Roslyn on purpose: no compiler assemblies to
    resolve, no MetadataReference wrangling, and the result is a plain CLR type,
    so nothing calls back into Python while the interpreter is shutting down.

    The members are derived from IFormatter by reflection instead of being
    hardcoded, so the shape stays correct whatever the interface declares:
      void  -> ret
      struct-> default value  (StreamingContext)
      class -> null           (SerializationBinder, ISurrogateSelector)
    """
    import clr
    from System import Array, Type
    from System.Reflection import AssemblyName, MethodAttributes, TypeAttributes
    from System.Reflection.Emit import (AssemblyBuilder, AssemblyBuilderAccess,
                                        OpCodes)

    iface = _formatter_interface_type()
    asm = AssemblyBuilder.DefineDynamicAssembly(
        AssemblyName('T3Lab.Interop.Serialization'), AssemblyBuilderAccess.Run)
    module = asm.DefineDynamicModule('T3Lab.Interop.Serialization')
    tb = module.DefineType(
        'T3Lab.Interop.NoopFormatter',
        TypeAttributes.Public | TypeAttributes.Class | TypeAttributes.Sealed)
    tb.AddInterfaceImplementation(iface)
    tb.DefineDefaultConstructor(MethodAttributes.Public)

    base_attrs = (MethodAttributes.Public | MethodAttributes.Virtual |
                  MethodAttributes.NewSlot | MethodAttributes.HideBySig |
                  MethodAttributes.Final)

    for iface_method in iface.GetMethods():
        attrs = base_attrs
        if iface_method.Name[:4] in ('get_', 'set_'):
            attrs = attrs | MethodAttributes.SpecialName
        arg_types = Array[Type](
            [p.ParameterType for p in iface_method.GetParameters()])
        mb = tb.DefineMethod(iface_method.Name, attrs,
                             iface_method.ReturnType, arg_types)
        il = mb.GetILGenerator()
        ret_type = iface_method.ReturnType
        if ret_type.FullName == 'System.Void':
            pass                                  # nothing to push
        elif ret_type.IsValueType:
            local = il.DeclareLocal(ret_type)     # default(T)
            il.Emit(OpCodes.Ldloca, local)
            il.Emit(OpCodes.Initobj, ret_type)
            il.Emit(OpCodes.Ldloc, local)
        else:
            il.Emit(OpCodes.Ldnull)
        il.Emit(OpCodes.Ret)
        tb.DefineMethodOverride(mb, iface_method)

    return tb.CreateType()


def enable_safe_engine_shutdown():
    """Make `Reload pyRevit` survive on Revit 2025+ / .NET 8+.

    Reload calls ScriptEngineManager.ClearEngines() -> PythonEngine.Shutdown()
    -> RuntimeData.Stash(), and Stash serializes with BinaryFormatter, which the
    runtime refuses:

        PlatformNotSupportedException: BinaryFormatter serialization and
        deserialization have been removed.

    sessionmgr._clear_running_engines() only catches AttributeError, so this
    escapes, aborts the reload, and leaves pythonnet half-torn-down — after
    which every script reports absent members on healthy .NET types and only a
    Revit restart clears it.

    `RuntimeData.FormatterType` exists precisely to replace BinaryFormatter:
    Stash does `Activator.CreateInstance(FormatterType)` and serializes through
    it. Pointing it at a no-op formatter lets Shutdown - and so the reload -
    complete.

    CORRECTION (2026-09-29, Revit 2026.5 journal): the stash is NOT discarded.
    pythonnet still stores an empty `sys.clr_data` capsule (the interpreter is
    never finalized), so the NEXT Initialize runs RestoreRuntimeData, gets None
    from Deserialize and dies with NullReferenceException in
    RestoreRuntimeDataImpl. The runtime is left half-initialized, and every
    later click fails in CPythonEngine.Start -> Runtime.set_PythonDLL with
    "This property must be set before runtime is initialized" until Revit
    restarts. This function only moves the failure from the reload to the next
    click; the real fix is not shutting CPython down on reload at all
    (lib/pyrevit_patches.py, opt-in because it edits pyRevit's install).
    startup.py detects the stuck state and asks for a restart.

    Note the earlier version of this function imported `Python.Runtime.NoopFormatter`,
    which does NOT exist in pyRevit's pythonnet build (verified 2026-09-04
    against pyRevit 6.5.5.26237): the ImportError was swallowed and the patch
    never applied. Hence emitting our own type, and hence SAFE_SHUTDOWN_STATUS -
    a silent failure here is indistinguishable from success until Reload dies.

    RuntimeData is static on the loaded assembly, so one call fixes Reload for
    the rest of the Revit session. Never raises; returns the status string.
    """
    global SAFE_SHUTDOWN_STATUS
    if sys.version_info[0] < 3:
        SAFE_SHUTDOWN_STATUS = 'skipped: IronPython 2'
        return SAFE_SHUTDOWN_STATUS
    try:
        try:
            from Python.Runtime import RuntimeData
        except ImportError:
            import clr
            clr.AddReference('pyRevitLabs.PythonNet')
            from Python.Runtime import RuntimeData
    except Exception as exc:
        SAFE_SHUTDOWN_STATUS = 'failed: Python.Runtime unavailable (%s)' % exc
        return SAFE_SHUTDOWN_STATUS

    try:
        if RuntimeData.FormatterType is not None:
            SAFE_SHUTDOWN_STATUS = 'already set: %s' % RuntimeData.FormatterType
            return SAFE_SHUTDOWN_STATUS
        RuntimeData.FormatterType = _build_noop_formatter_type()
        SAFE_SHUTDOWN_STATUS = 'ok: %s' % RuntimeData.FormatterType
    except Exception as exc:
        SAFE_SHUTDOWN_STATUS = 'failed: %s: %s' % (type(exc).__name__, exc)
    return SAFE_SHUTDOWN_STATUS


def safe_output():
    """The pyRevit output window, or a no-op stand-in — never raises.

    Call sites should use this instead of `script.get_output()` directly: the
    bootstrap runs long before a script imports pyrevit, so patching
    `pyrevit.script` from here is a timing gamble. This is not.
    """
    if sys.version_info[0] >= 3:
        return _NullOutput()        # CPython has no ScriptOutput.GetDefault
    try:
        from pyrevit import script
        return script.get_output()
    except Exception:
        return _NullOutput()


def install_script_shim():
    """Patch `pyrevit.script.get_output` when pyrevit is already loaded.

    Belt-and-braces for library modules; the reliable path is `safe_output()`.
    Deliberately does NOT import pyrevit itself — the bootstrap runs at the very
    top of a script, before pyrevit is ready, and forcing the import there is
    both fragile and a side effect we do not want.
    """
    if sys.version_info[0] < 3:
        return                      # IronPython has the real output window
    for mod_name in ('pyrevit.script', 'pyrevit.output'):
        mod = sys.modules.get(mod_name)
        if mod is None:
            continue                # not imported yet - safe_output() covers it
        try:
            mod.get_output = lambda *_a, **_k: _NullOutput()
        except Exception:
            pass


def install_imp_shim():
    """Shim the removed Python standard library module `imp` on Python 3.12+.

    Python 3.12 removed `imp` completely. Legacy pyRevit modules or wrappers
    (such as bundled `rpw.utils.sphinx_compat`) still have `import imp` at
    module top-level. Registering a compatibility shim in `sys.modules['imp']`
    prevents fatal `ModuleNotFoundError: No module named 'imp'` crashes across
    the entire Revit session.
    """
    if 'imp' in sys.modules:
        return
    try:
        import importlib
        import types

        imp = types.ModuleType('imp')
        imp.new_module = lambda name: types.ModuleType(name)
        imp.reload = importlib.reload
        imp.find_module = lambda *args, **kwargs: (None, None, ('', '', 0))
        imp.load_module = lambda *args, **kwargs: None
        imp.acquire_lock = lambda: None
        imp.release_lock = lambda: None
        imp.lock_held = lambda: False
        sys.modules['imp'] = imp
    except Exception:
        pass


def install_wpf_shim():
    """Shim the IronPython-only `wpf` module on CPython 3.

    Under IronPython, `import wpf; wpf.LoadComponent(self, xaml_path)` was standard.
    Under CPython (PythonNet), `wpf` does not exist. Providing this shim ensures
    legacy scripts and dialogs that import `wpf` don't throw ModuleNotFoundError.
    """
    if 'wpf' in sys.modules:
        return
    try:
        import types
        wpf = types.ModuleType('wpf')

        def LoadComponent(window, xaml_path):
            try:
                from GUI.WPF_Base import T3WPFWindow
                if isinstance(window, T3WPFWindow):
                    return
            except Exception:
                pass

        wpf.LoadComponent = LoadComponent
        sys.modules['wpf'] = wpf
    except Exception:
        pass


# Run immediately upon import.
# NOTE: this module lives in lib/, so the persistent CPython engine caches it in
# sys.modules — editing it requires one pyRevit Reload before the new code runs.
init_cpython_paths()
fix_std_streams()
install_forms_shim()
install_script_shim()
install_imp_shim()
install_wpf_shim()
enable_safe_engine_shutdown()
