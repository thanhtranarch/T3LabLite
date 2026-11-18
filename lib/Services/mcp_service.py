# -*- coding: utf-8 -*-
"""
MCPService — Shared MCP control backend

Single source of truth for all MCP-related operations:
  • MCP HTTP server (start / stop / status)
  • File-based task watcher (start / stop / status)
  • MCP client config snippet generation (JSON + Codex TOML)
  • Auto-configure Claude Desktop / ChatGPT (Codex) / Antigravity, for
    t3lab-revit and for the AutoCAD add-in's t3lab-autocad bridge
  • Data directory management

Import this from any dialog or tool that needs MCP control:
    from Services.mcp_service import MCPService
"""

from __future__ import unicode_literals

import os
import re
import sys
import time

# ─── Path helper ───────────────────────────────────────────────────────────────
def _ensure_lib_in_path():
    here    = os.path.dirname(os.path.abspath(__file__))   # lib/Services
    lib_dir = os.path.dirname(here)                        # lib
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)


# ─── Lazy imports (avoid crashing if running outside Revit) ────────────────────
def _get_server():
    _ensure_lib_in_path()
    from core.server import get_t3labai_server
    return get_t3labai_server()


def _get_watcher():
    _ensure_lib_in_path()
    from core.file_watcher import get_task_watcher
    return get_task_watcher()


def _get_paths_module():
    _ensure_lib_in_path()
    from core import paths
    return paths


def _get_data_dir():
    _ensure_lib_in_path()
    try:
        from core.file_watcher import T3LAB_DATA_DIR
        return T3LAB_DATA_DIR
    except Exception:
        try:
            return _get_paths_module().get_setting(
                'data_dir',
                lambda: os.path.join(os.path.expanduser('~'), 'T3Lab_AI_Data'))
        except Exception:
            return os.path.join(os.path.expanduser('~'), 'T3Lab_AI_Data')


def _source_bridge_path():
    """Absolute path to the bridge.py shipped inside this extension."""
    here    = os.path.dirname(os.path.abspath(__file__))   # lib/Services
    lib_dir = os.path.dirname(here)                        # lib
    return os.path.join(lib_dir, 'core', 'bridge.py')


def _get_bridge_path():
    """
    Bridge path to embed in Claude Desktop / Cursor configs.

    Always the machine-stable deployed copy (%APPDATA%/T3LabAI/bridge.py,
    see MCPService.deploy_bridge) — never a path inside the extension
    folder, so the config keeps working when the extension is moved,
    re-downloaded or updated. Falls back to the in-extension copy only if
    deployment fails (e.g. APPDATA not writable).
    """
    deployed, _err = MCPService.deploy_bridge()
    return (deployed or _source_bridge_path()).replace('\\', '/')


def _shown_bridge_path():
    """The path the snippet shows, without deploying: only Connect may copy the
    bridge and take its ownership (T3Lab and T3Lab Lite share the folder;
    opening MCP Control used to switch the bridge to this extension)."""
    try:
        target = os.path.join(_get_paths_module().settings_dir(), 'bridge.py')
        if os.path.isfile(target):
            return target.replace('\\', '/')
    except Exception:
        pass
    return _source_bridge_path().replace('\\', '/')


def _is_store_stub(path):
    """The WindowsApps app-execution alias: opens the Microsoft Store unless
    Python was installed from there. Tried last, never cached."""
    return '\\windowsapps\\' in (path or '').replace('/', '\\').lower()


def _store_python_installed():
    """Python from the Microsoft Store is really there (the alias then works)."""
    try:
        import glob
        root = os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'WindowsApps')
        return bool(glob.glob(os.path.join(root, 'PythonSoftwareFoundation.Python.3*')))
    except Exception:
        return False


def _python_usable(python):
    """False when the command Connect would write cannot start: no Python
    found (bare name), or only the Store alias that opens the Store."""
    if not python or not os.path.isfile(python):
        return False
    return not _is_store_stub(python) or _store_python_installed()


NO_PYTHON_TEXT = ('Python 3 was not found on this computer. The AI app needs it to run the '
                  'T3Lab connection: install Python 3 from python.org (tick "Add python.exe '
                  'to PATH"), then click Connect again.')


def _python_version_key(path):
    m = re.search(r'python(\d)(\d+)', os.path.basename(os.path.dirname(path)).lower())
    return (int(m.group(1)), int(m.group(2))) if m else (0, 0)


def _find_python_executable():
    """
    Locate a CPython 3 interpreter to run core/bridge.py.

    bridge.py needs f-strings and urllib.request (Python 3.6+) — the
    IronPython interpreter running this Revit process can't run it, and a
    bare "python" command isn't guaranteed to resolve on every machine
    (e.g. python.org installs that only register "python3", or a PATH
    that hasn't picked up a fresh install yet). Search PATH first, then
    fall back to common per-user/system install locations, and only use
    the bare command name as a last resort so Claude Desktop still gets
    something to try.

    The resolved path is read from / written to mcp_paths.json (see
    core/paths.py) so the scan only runs once per machine and the result
    is hand-editable afterwards; the cached value is revalidated on every
    call in case Python was reinstalled elsewhere.
    """
    paths = _get_paths_module()
    cached_path = paths.load_settings().get('python_executable')
    if cached_path and os.path.isfile(cached_path) and not _is_store_stub(cached_path):
        return cached_path

    is_windows = os.name == 'nt'
    exe_names = ['python.exe', 'python3.exe'] if is_windows else ['python3', 'python']
    found = None

    # 1) Search PATH directories for a real interpreter.
    path_dirs = os.environ.get('PATH', '').split(os.pathsep)
    stub = None
    for d in path_dirs:
        for name in exe_names:
            candidate = os.path.join(d, name)
            if os.path.isfile(candidate):
                if _is_store_stub(candidate):
                    stub = stub or candidate
                    continue
                found = candidate
                break
        if found:
            break

    # 2) Common install locations not always present on PATH.
    if not found:
        home = os.path.expanduser('~')
        fallback_globs = []
        if is_windows:
            fallback_globs.append(os.path.join(home, 'AppData', 'Local', 'Programs', 'Python', 'Python*', 'python.exe'))
            fallback_globs.append('C:/Python*/python.exe')
            fallback_globs.append('C:/Program Files/Python*/python.exe')
        else:
            fallback_globs.append('/usr/local/bin/python3')
            fallback_globs.append('/usr/bin/python3')
            fallback_globs.append(os.path.join(home, '.pyenv', 'shims', 'python3'))

        try:
            import glob
            for pattern in fallback_globs:
                # Newest first by number: a string sort puts Python39 above Python312.
                matches = sorted(glob.glob(pattern), key=_python_version_key, reverse=True)
                if matches:
                    found = matches[0]
                    break
        except Exception:
            pass

    if found:
        paths.set_setting('python_executable', found)
        return found
    if stub:
        return stub     # may work (Store Python installed); never cached

    # 3) Give up — return the bare command name and let the OS PATH try.
    # Not cached: it isn't a real resolved path, so nothing to reuse.
    return 'python' if is_windows else 'python3'


# ─── Cross-engine server detection ─────────────────────────────────────────────
# pyRevit runs startup.py and `#! python3` pushbuttons in DIFFERENT Python
# engines inside ONE Revit process (IronPython vs CPython). core/server.py
# anchors its singleton on the AppDomain as a Python object, which cannot cross
# that boundary: the other engine receives an opaque CLR object it cannot call
# ("'Object_1$1' object has no attribute 'get_server_stats'"), then happily
# binds a SECOND port in the same process — one Revit was seen LISTENING on
# both 48884 and 48885.
#
# /health is engine-independent and reports the owning pid, so probing it is
# the only reliable way to answer "does THIS Revit process already host a
# server?" regardless of which engine started it.

SERVER_PORT_MIN, SERVER_PORT_MAX = 48884, 48894


def _probe_health(port, timeout=0.35):
    """Return the /health payload for `port`, or None if nothing answers."""
    try:
        import json as _json
        try:
            from urllib.request import build_opener, ProxyHandler
        except ImportError:                       # IronPython 2
            from urllib2 import build_opener, ProxyHandler
        # No proxy for 127.0.0.1: a system / HTTP_PROXY proxy without a
        # localhost bypass made a running server look "Stopped" and let
        # Start bind a second one.
        opener = build_opener(ProxyHandler({}))
        raw = opener.open('http://127.0.0.1:{}/health'.format(port), timeout=timeout).read()
        if isinstance(raw, bytes):
            raw = raw.decode('utf-8', 'replace')
        return _json.loads(raw)
    except Exception:
        return None


def _recorded_port():
    """Port recorded by the server in this Revit: 0 = stopped, None = no
    record (outside Revit, or a server from before the record existed)."""
    try:
        from System import AppDomain
        value = AppDomain.CurrentDomain.GetData('_t3lab_mcp_server_singleton.port')
        return None if value is None else int(value)
    except Exception:
        return None


def _local_servers():
    """
    [(port, /health payload)] for every T3Lab server inside THIS Revit process.

    Matched by pid, so a server in another Revit window is correctly ignored —
    those are legitimate and the bridge switches between them on purpose.
    """
    try:
        mypid = os.getpid()
    except Exception:
        return []
    # The server records its port for the whole process (core.server
    # _record_port): one probe instead of the whole range. The scan cost
    # seconds on Revit's UI thread at every start and document open (a closed
    # loopback port is slow to refuse; IronPython ignores the connect timeout).
    recorded = _recorded_port()
    if recorded is not None:
        info = _probe_health(recorded, 1.0) if recorded else None
        if info and info.get('status') == 'ok' and info.get('pid') == mypid:
            return [(recorded, info)]
        return []
    if sys.platform == 'cli':
        return []       # IronPython hosts the server: no record = not started in this Revit
    found = []
    for port in range(SERVER_PORT_MIN, SERVER_PORT_MAX + 1):
        info = _probe_health(port)
        if info and info.get('status') == 'ok' and info.get('pid') == mypid:
            found.append((port, info))
    return found


def _local_server_ports():
    """Ports already hosting a T3Lab server inside THIS Revit process."""
    return [port for port, _info in _local_servers()]


# ─── Engine routing: the HTTP server lives in IronPython ──────────────────────
# pyRevit's CPython engine keeps the GIL on Revit's main thread whenever no
# CPython code is running (see core/server.py, start_server). A server thread
# started from a `#! python3` tool therefore froze as soon as the tool's window
# closed: MCP Control said "Running on port 48884" while the bridge found no
# server. So CPython never hosts it. It asks pyRevit's script executor to run
# core/mcp_ipy_host.py in the IronPython engine — the same engine and the same
# call pyRevit uses for the DocumentOpened hook that auto-starts the server.

IPY_HOST_ID     = 't3lab-mcp-ipy-host'
IPY_HOST_RESULT = '_t3lab_mcp_ipy_host_result'     # = mcp_ipy_host.RESULT_KEY
_EXEC_SUCCEEDED, _EXEC_DELAYED = 0, 11             # ScriptExecutorResultCodes


def _cpython_in_revit():
    """True in pyRevit's CPython engine inside Revit — the one engine that
    must not host the HTTP server. False under IronPython and in dev runs."""
    if sys.platform == 'cli':
        return False
    try:
        _ensure_lib_in_path()
        from core import server as _srv
        return bool(_srv.HAS_REVIT_UI)
    except Exception:
        return False


def _lib_dir():
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _pyrevit_runtime():
    """pyRevit's PyRevitLabs.PyRevit.Runtime namespace (ScriptExecutor & co)."""
    import clr
    try:
        import PyRevitLabs.PyRevit.Runtime as runtime
        runtime.ScriptExecutor
        return runtime
    except Exception:
        pass
    # Not registered with pythonnet yet: find the loaded assembly that holds
    # the executor and add it by path, as pyrevit.runtime.types does.
    from System import AppDomain
    for asm in AppDomain.CurrentDomain.GetAssemblies():
        try:
            if asm.GetType('PyRevitLabs.PyRevit.Runtime.ScriptExecutor') is not None:
                clr.AddReference(asm.Location)
                break
        except Exception:
            continue
    import PyRevitLabs.PyRevit.Runtime as runtime
    return runtime


def _ipy_search_paths(runtime):
    """sys.path for the IronPython host: the search paths pyRevit gave this
    extension's own hooks when it can tell, else the extension and pyRevit's
    library folders. Never CPython's sys.path — its stdlib is Python 3."""
    lib_dir = _lib_dir()
    ext_dir = os.path.dirname(lib_dir)
    hooks_dir = os.path.normcase(os.path.join(ext_dir, 'hooks'))
    try:
        for hook in runtime.EventHooks.GetAllEventHooks():
            script = os.path.normcase(os.path.dirname(str(hook.Script)))
            if script == hooks_dir:
                paths = [str(p) for p in hook.SearchPaths if p]
                if paths:
                    return paths
    except Exception:
        pass
    paths = [lib_dir, ext_dir]
    try:
        import pyrevit
        pyrevitlib = os.path.dirname(os.path.dirname(os.path.abspath(pyrevit.__file__)))
        paths.append(pyrevitlib)
        site = os.path.join(os.path.dirname(pyrevitlib), 'site-packages')
        if os.path.isdir(site):
            paths.append(site)
    except Exception:
        pass
    return paths


def _host_uiapp():
    """The UIApplication of this pyRevit session, or None."""
    try:
        import builtins
        app = getattr(builtins, '__revit__', None)
        if app is not None and hasattr(app, 'ActiveUIDocument'):
            return app
    except Exception:
        pass
    try:
        from pyrevit import HOST_APP
        return HOST_APP.uiapp
    except Exception:
        return None


def _run_ipy_host(action, port=None):
    """
    Run core/mcp_ipy_host.py in pyRevit's IronPython engine.

    On Revit's main thread pyRevit runs it right away and the host's answer is
    returned. From a worker thread pyRevit queues it on its own ExternalEvent
    and it runs when Revit is idle — that request returns (True, None) at once.

    Returns:
        (success: bool, error_message: str|None)
    """
    try:
        from System import AppDomain, String
        from Autodesk.Revit.DB import ElementSet
        _ensure_lib_in_path()
        from Snippets._compat import net_list
        runtime = _pyrevit_runtime()

        script = os.path.join(_lib_dir(), 'core', 'mcp_ipy_host.py')
        data = runtime.ScriptData()
        data.ScriptPath = script
        data.ConfigScriptPath = script
        data.CommandUniqueId = IPY_HOST_ID
        data.CommandName = 'T3Lab MCP server host'
        data.CommandBundle = 'T3Lab.hooks'
        data.CommandExtension = 'T3Lab'
        data.HelpSource = ''

        args = [action] + ([str(port)] if port else [])
        cfg = runtime.ScriptRuntimeConfigs()
        cfg.CommandData = None
        cfg.SelectedElements = ElementSet()
        cfg.SearchPaths = net_list(String, _ipy_search_paths(runtime))
        cfg.Arguments = net_list(String, args)
        # Same frame settings as pyRevit's hooks; the type keys pin IronPython
        # on pyRevit 6 and are ignored by older builds, where the missing
        # shebang already selects it.
        cfg.EngineConfigs = ('{"full_frame": true, "type": "IronPython", '
                             '"type_explicit": true}')
        cfg.RefreshEngine = False
        cfg.ConfigMode = False
        cfg.DebugMode = False
        cfg.ExecutedFromUI = False
        try:
            cfg.SuppressOutput = True       # pyRevit 6+: no output window
        except Exception:
            pass
        uiapp = _host_uiapp()
        if uiapp is not None:
            cfg.UIApp = uiapp

        exec_cfg = runtime.ScriptExecutorConfigs()
        exec_cfg.WaitForResult = False      # never spin a worker thread
        exec_cfg.SendTelemetry = False

        AppDomain.CurrentDomain.SetData(IPY_HOST_RESULT, None)
        code = int(runtime.ScriptExecutor.ExecuteScript(data, cfg, exec_cfg))
        result = AppDomain.CurrentDomain.GetData(IPY_HOST_RESULT)
    except Exception as ex:
        return False, 'Could not run the IronPython server host: {}'.format(ex)
    return _host_outcome(code, result)


def _host_outcome(code, result):
    """(success, error) from pyRevit's result code and the host's report."""
    if code == _EXEC_DELAYED:
        return True, None
    result = str(result) if result is not None else ''
    if code == _EXEC_SUCCEEDED and result == 'ok':
        return True, None
    if result.startswith('error: '):
        return False, result[len('error: '):]
    return False, 'The IronPython server host ended with code {}'.format(code)


# ─── AI client registry ────────────────────────────────────────────────────────
# Every MCP client that can be auto-configured launches the SAME stdio bridge
# (%APPDATA%/T3LabAI/bridge.py). Only the config file location and its syntax
# differ:
#
#   Claude Desktop   claude_desktop_config.json          JSON  "mcpServers"
#   ChatGPT (Codex)  ~/.codex/config.toml                TOML  [mcp_servers.X]
#   Antigravity      ~/.gemini/antigravity-ide/
#                        mcp_config.json                 JSON  "mcpServers"
#
# Cursor, Windsurf and Claude Code use the same JSON shape as Claude Desktop,
# so config_snippet(fmt='json') can be pasted into those by hand as well.
#
# ChatGPT note: the ChatGPT app's own "Connectors" (developer mode) only accept
# a PUBLIC https MCP endpoint — a local stdio bridge cannot be registered there.
# The supported local path is Codex (the CLI and the Codex panel inside the
# ChatGPT desktop app), which reads ~/.codex/config.toml. That is what we write.

MCP_SERVER_NAME = 't3lab-revit'

# Other Autodesk apps: each add-in ships its own stdio bridge, deployed to
# %APPDATA%/<folder>/bridge.py (AutoCAD: t3lab_CADmcp scripts/install.ps1).
# A bridge takes no port (it finds its app itself) and runs with `python -I`.
# mcp_paths.json '<key>_bridge' overrides the path. Only apps whose add-in CI
# builds and Connect installs in one click are listed: AutoCAD and Civil 3D
# (owner, 2026-11-17: Navisworks, 3ds Max, Inventor dropped). Order = the list.
AUTODESK_APPS = (
    # key,         label,        server name,        %APPDATA% folder,   add-in hint
    # A folder named in Services.addin_installer.PACKAGES is installed by Connect.
    # Civil 3D runs the AutoCAD add-in (T3LabCAD loads in every AutoCAD-based
    # product), so it shares AutoCAD's bridge and server entry.
    ('autocad',    'AutoCAD',    't3lab-autocad',    'T3LabCAD',
     'Connect installs it (AutoCAD and Civil 3D 2022-2027)'),
    ('civil3d',    'Civil 3D',   't3lab-autocad',    'T3LabCAD',
     'Connect installs it (the AutoCAD add-in also runs in Civil 3D 2022-2027)'),
)
CAD_SERVER_NAME = 't3lab-autocad'
SERVERS = dict([('revit', MCP_SERVER_NAME)] + [(a[0], a[2]) for a in AUTODESK_APPS])
_APP_BY_KEY = dict((a[0], a) for a in AUTODESK_APPS)
_BRIDGE_SERVERS = frozenset(a[2] for a in AUTODESK_APPS)
# Connect also pulls the latest add-in release when one is installed, at most
# once per package per _ADDIN_REFRESH_S (Connect All writes three clients).
_ADDIN_REFRESH_S = 600
_addin_refreshed = {}


def _server_name(server):
    try:
        return SERVERS[server]
    except KeyError:
        raise KeyError('Unknown MCP server: {}'.format(server))


def _app_bridge_path(app='autocad'):
    """%APPDATA%/<folder>/bridge.py of an Autodesk app, or mcp_paths.json
    '<key>_bridge' if set."""
    spec = _APP_BY_KEY[app]
    try:
        override = _get_paths_module().load_settings().get(app + '_bridge')
    except Exception:
        override = None
    path = override or os.path.join(_appdata(), spec[3], 'bridge.py')
    return path.replace('\\', '/')


def _cad_bridge_path():
    return _app_bridge_path('autocad')


def _claude_desktop_running():
    """
    True while the Claude Desktop app runs: it rewrites its config when it
    quits, so an entry written meanwhile is lost (t3lab_CADmcp 4b04809).
    Claude Code also ships a claude.exe, so match by install folder.
    """
    def _is_desktop(path):
        p = (path or '').replace('/', '\\').lower()
        return '\\anthropicclaude\\' in p or '\\windowsapps\\' in p

    try:
        from System.Diagnostics import Process
        for proc in Process.GetProcessesByName('claude'):
            try:
                if _is_desktop(proc.MainModule.FileName):
                    return True
            except Exception:
                continue
        return False
    except Exception:
        pass
    if os.name != 'nt':
        return False
    try:
        import subprocess
        out = subprocess.check_output(
            ['powershell', '-NoProfile', '-Command',
             '(Get-Process claude -ErrorAction SilentlyContinue).Path'],
            stderr=subprocess.STDOUT, creationflags=0x08000000)
        return any(_is_desktop(line.decode('utf-8', 'ignore').strip())
                   for line in out.splitlines())
    except Exception:
        return False


def _home():
    return os.path.expanduser('~')


def _appdata():
    return os.environ.get('APPDATA', os.path.join(_home(), 'AppData', 'Roaming'))


def _claude_candidates():
    import platform
    system = platform.system()
    if system == 'Windows':
        return [os.path.join(_appdata(), 'Claude', 'claude_desktop_config.json')]
    if system == 'Darwin':
        return [os.path.join(_home(), 'Library', 'Application Support', 'Claude',
                             'claude_desktop_config.json')]
    return [os.path.join(_home(), '.config', 'Claude', 'claude_desktop_config.json')]


def _codex_candidates():
    """Codex reads $CODEX_HOME/config.toml, defaulting to ~/.codex/config.toml."""
    out = []
    codex_home = os.environ.get('CODEX_HOME')
    if codex_home:
        out.append(os.path.join(codex_home, 'config.toml'))
    out.append(os.path.join(_home(), '.codex', 'config.toml'))
    return out


def _antigravity_candidates():
    """
    Antigravity keeps its MCP registry outside the VS Code-style user folder —
    it inherits Cascade's layout, so the live file is
    ~/.gemini/antigravity-ide/mcp_config.json. The other entries cover
    alternative layouts; whichever exists first wins.
    """
    return [
        # The live file first: when none exists yet, Connect creates this one.
        os.path.join(_home(), '.gemini', 'antigravity-ide', 'mcp_config.json'),
        os.path.join(_home(), '.gemini', 'config', 'mcp_config.json'),
        os.path.join(_home(), '.gemini', 'antigravity', 'mcp_config.json'),
        os.path.join(_home(), '.antigravity', 'mcp_config.json'),
        os.path.join(_appdata(), 'Antigravity', 'User', 'globalStorage', 'mcp_config.json'),
    ]


# key -> descriptor. `setting` is the mcp_paths.json key that overrides the
# auto-detected path, so a non-standard install stays hand-editable.
AI_CLIENTS = [
    {'key': 'claude',
     'label': 'Claude Desktop',
     'fmt': 'json',
     'setting': 'claude_desktop_config',
     'candidates': _claude_candidates,
     'agent': 'claude-desktop'},
    {'key': 'chatgpt',
     'label': 'ChatGPT (Codex)',
     'fmt': 'toml',
     'setting': 'codex_config',
     'candidates': _codex_candidates,
     'agent': 'codex'},
    {'key': 'antigravity',
     'label': 'Antigravity',
     'fmt': 'json',
     'setting': 'antigravity_config',
     'candidates': _antigravity_candidates,
     'agent': 'antigravity'},
]


def _client_spec(key):
    for spec in AI_CLIENTS:
        if spec['key'] == key:
            return spec
    raise KeyError('Unknown MCP client: {}'.format(key))


def _resolve_client_path(spec):
    """
    Config path for one client.

    mcp_paths.json wins if it names a file that exists (or nothing better
    exists); otherwise the first candidate that already exists on disk is
    chosen and persisted, so it becomes editable. When no candidate exists
    yet, the primary one is returned but NOT persisted: opening MCP Control
    before the AI app is installed (or before CODEX_HOME is set) must not pin
    a path the app will never read.
    """
    candidates = spec['candidates']()
    found = next((c for c in candidates if os.path.isfile(c)), None)
    try:
        paths = _get_paths_module()
        stored = paths.load_settings().get(spec['setting'])
    except Exception:
        paths, stored = None, None
    if stored and (os.path.isfile(stored) or not found):
        return stored
    if found:
        if paths is not None:
            try:
                paths.set_setting(spec['setting'], found)
            except Exception:
                pass
        return found
    return candidates[0]


def _backup_once(path):
    """
    Keep a pristine copy of a user config the first time we touch it.

    These files hold the user's other MCP servers, model providers and plugins;
    `<path>.t3lab-backup` is written only when it does not already exist, so the
    very first pre-T3Lab state stays recoverable. Best-effort — a failed backup
    never blocks the write.
    """
    try:
        backup = path + '.t3lab-backup'
        if os.path.isfile(path) and not os.path.isfile(backup):
            with open(path, 'rb') as f:
                data = f.read()
            with open(backup, 'wb') as f:
                f.write(data)
    except Exception:
        pass


def _agent_args(agent):
    """['--agent=<id>']: the bridge names the AI app with it (Agent Cursors,
    audit) instead of the app's self-declared clientInfo, which Antigravity
    sends as 'local-agent-mode-...' (owner, 2026-11-17: shown as "Claude")."""
    return ['--agent={}'.format(agent)] if agent else []


def _server_entry(python, bridge, port, name=MCP_SERVER_NAME, agent=None):
    """The JSON object describing our stdio server (Claude / Antigravity)."""
    if name in _BRIDGE_SERVERS:
        # -I = isolated: pyRevit's machine-wide PYTHONPATH cannot break it.
        return {'type': 'stdio', 'command': python, 'args': ['-I', bridge]}
    return {
        'type': 'stdio',
        'command': python,
        'args': ['-E', bridge, str(port)] + _agent_args(agent),
        'env': {'PYTHONPATH': ''},
    }


# ── JSON clients (Claude Desktop, Antigravity) ────────────────────────────────

def _read_json_config(path):
    import json as _json
    import codecs
    if not os.path.isfile(path):
        return {}
    # utf-8-sig: Notepad / PowerShell 5.1 save UTF-8 with a BOM, which
    # json.loads rejects ("Unexpected UTF-8 BOM").
    with codecs.open(path, 'r', encoding='utf-8-sig') as f:
        raw = f.read().strip()
    if not raw:
        return {}
    return _json.loads(raw)


def _write_text(path, text):
    """Write a client config through a temp file: never half-written."""
    data = text.encode('utf-8') if not isinstance(text, bytes) else text
    _get_paths_module().atomic_write(path, data)


def _json_is_configured(path, name=MCP_SERVER_NAME):
    config = _read_json_config(path)
    servers = config.get('mcpServers') if isinstance(config, dict) else None
    entry = (servers or {}).get(name)
    if entry is None:
        return False
    if name == MCP_SERVER_NAME:
        # An entry from before --agent (2026-11-17) runs under the app's
        # self-declared name: show Connect again so one click updates it.
        args = entry.get('args') if isinstance(entry, dict) else None
        return any('{}'.format(a).startswith('--agent=') for a in (args or []))
    return True


def _write_json_config(path, python, bridge, port, name=MCP_SERVER_NAME, agent=None):
    import json as _json
    import codecs

    # A file we cannot parse (comments, trailing comma, hand-edit typo) holds
    # the user's other MCP servers and settings: never replace it with only
    # our entry -- the one-time backup may be from long ago.
    try:
        config = _read_json_config(path)
    except Exception as ex:
        raise ValueError('{} is not valid JSON ({}). Fix it by hand (or add the '
                         'entry from Manual Config), then Connect again.'.format(path, ex))
    if not isinstance(config, dict):
        raise ValueError('{} does not hold a JSON object. Fix it by hand, then '
                         'Connect again.'.format(path))

    servers = config.get('mcpServers')
    if not isinstance(servers, dict):
        servers = {}
        config['mcpServers'] = servers

    entry = _server_entry(python, bridge, port, name, agent)
    # Antigravity stamps its own entries with a protobuf type marker; keep it
    # if a previous entry had one so the client still recognises the record.
    previous = servers.get(name)
    if isinstance(previous, dict) and '$typeName' in previous:
        entry['$typeName'] = previous['$typeName']
    servers[name] = entry

    cfg_dir = os.path.dirname(path)
    if cfg_dir and not os.path.isdir(cfg_dir):
        os.makedirs(cfg_dir)

    _backup_once(path)
    _write_text(path, _json.dumps(config, indent=2, ensure_ascii=False))


# ── TOML client (Codex / ChatGPT) ─────────────────────────────────────────────

# A TOML table header alone on its line: [x], [x.y], [[x]] — deliberately
# stricter than "starts with [" so a line of an inline array is never mistaken
# for the start of the next table.
_TOML_HEADER_RE = re.compile(r'^\s*\[\[?[^\[\]]+\]\]?\s*(?:#.*)?$')


def _toml_section_re(name):
    """Matches `[mcp_servers.<name>]` and any of its sub-tables (.env etc.)."""
    return re.compile(
        r'^\s*\[\s*mcp_servers\s*\.\s*["\']?' + re.escape(name) +
        r'["\']?\s*(?:\.[^\]]*)?\]\s*(?:#.*)?$')


def _toml_quote(value):
    """Basic TOML string — backslashes and quotes escaped (Windows paths)."""
    return value.replace('\\', '\\\\').replace('"', '\\"')


def _codex_block(python, bridge, port, name=MCP_SERVER_NAME, agent=None):
    if name in _BRIDGE_SERVERS:
        args = 'args = ["-I", "{}"]\n'.format(_toml_quote(bridge))
    else:
        # -E: pyRevit's machine-wide PYTHONPATH must not reach the bridge
        # (same as the JSON entry).
        args = 'args = [{}]\n'.format(', '.join('"{}"'.format(_toml_quote(a)) for a in
                                               ['-E', bridge, str(port)] + _agent_args(agent)))
    return (
        '[mcp_servers.{}]\n'.format(name) +
        'type = "stdio"\n' +
        'command = "{}"\n'.format(_toml_quote(python)) +
        args +
        # Revit can be slow to answer the first probe on a cold start; Codex
        # kills a server that has not handshaken within 10s by default.
        'startup_timeout_sec = 60\n'
    )


def _toml_upsert(text, block, name=MCP_SERVER_NAME):
    """
    Replace the `[mcp_servers.<name>]` table (and its sub-tables) in
    `text` with `block`, or append it when absent.

    Everything outside our own table is preserved line-for-line — the file
    holds the user's model providers, other MCP servers and plugins.
    """
    section_re = _toml_section_re(name)
    lines = text.splitlines()
    out = []
    i, n = 0, len(lines)
    replaced = False

    while i < n:
        if section_re.match(lines[i]):
            if not replaced:
                out.extend(block.rstrip('\n').split('\n'))
                replaced = True
            i += 1
            while i < n and not _TOML_HEADER_RE.match(lines[i]):
                i += 1
            continue
        out.append(lines[i])
        i += 1

    if not replaced:
        if out and out[-1].strip():
            out.append('')
        out.extend(block.rstrip('\n').split('\n'))

    return '\n'.join(out) + '\n'


def _read_toml_text(path):
    import codecs
    if not os.path.isfile(path):
        return ''
    with codecs.open(path, 'r', encoding='utf-8-sig') as f:
        return f.read()


def _toml_other_form(text, name=MCP_SERVER_NAME):
    """True when `name` is declared other than as a [mcp_servers.<name>]
    header: a key under [mcp_servers], a dotted key, or an inline
    `mcp_servers = {...}`. _toml_upsert cannot rewrite those, and appending a
    table would declare it twice: Codex then rejects the whole file."""
    key = r'["\']?' + re.escape(name) + r'["\']?'
    in_servers = False
    for line in text.splitlines():
        if _TOML_HEADER_RE.match(line):
            in_servers = re.match(r'^\s*\[\s*mcp_servers\s*\]', line) is not None
            continue
        if in_servers and re.match(r'^\s*' + key + r'\s*[=.]', line):
            return True
        if re.match(r'^\s*mcp_servers\s*\.\s*' + key + r'\s*[=.]', line):
            return True
        if re.match(r'^\s*mcp_servers\s*=', line) and name in text:
            return True
    return False


def _toml_is_configured(path, name=MCP_SERVER_NAME):
    section_re = _toml_section_re(name)
    inside = found = False
    for line in _read_toml_text(path).splitlines():
        if section_re.match(line):
            inside = found = True
            continue
        if inside and _TOML_HEADER_RE.match(line):
            inside = False
        if inside and name == MCP_SERVER_NAME and '--agent=' in line:
            return True
    return found and name != MCP_SERVER_NAME     # revit entry without --agent: Connect again


_TOML_ARGS_RE = re.compile(r'^(\s*args\s*=\s*\[.*?)\s*\](\s*(?:#.*)?)$')


def _json_add_agent(path, agent):
    """Append --agent=<id> to our own t3lab-revit entry when it has none.
    True when the file was changed. Raises on a file that cannot be parsed."""
    import json as _json
    config = _read_json_config(path)
    servers = config.get('mcpServers') if isinstance(config, dict) else None
    entry = (servers or {}).get(MCP_SERVER_NAME)
    args = entry.get('args') if isinstance(entry, dict) else None
    if not isinstance(args, list):
        return False
    texts = ['{}'.format(a) for a in args]
    if any(a.startswith('--agent=') for a in texts) or not any(
            a.replace('\\', '/').endswith('/bridge.py') for a in texts):
        return False                # already named, or not the entry we wrote
    entry['args'] = args + _agent_args(agent)
    _backup_once(path)
    _write_text(path, _json.dumps(config, indent=2, ensure_ascii=False))
    return True


def _toml_add_agent(path, agent):
    """Same for Codex: add "--agent=<id>" to the args of [mcp_servers.t3lab-revit]."""
    section_re = _toml_section_re(MCP_SERVER_NAME)
    lines = _read_toml_text(path).split('\n')
    inside = False
    for i, line in enumerate(lines):
        if section_re.match(line):
            inside = True
            continue
        if inside and _TOML_HEADER_RE.match(line):
            break
        m = _TOML_ARGS_RE.match(line.rstrip('\r')) if inside else None
        if m is None:
            continue
        if '--agent=' in line or 'bridge.py' not in line:
            return False
        end = '\r' if line.endswith('\r') else ''
        head = m.group(1).rstrip()
        if head.endswith(','):          # ["...", "48884",] : a second comma breaks the file
            head = head[:-1].rstrip()
        lines[i] = '{}, "{}"]{}{}'.format(head, _agent_args(agent)[0], m.group(2), end)
        _backup_once(path)
        _write_text(path, '\n'.join(lines))
        return True
    return False


def _write_toml_config(path, python, bridge, port, name=MCP_SERVER_NAME, agent=None):
    # A read error (non-UTF-8 byte, file locked) must not turn into '' and
    # overwrite the user's providers and servers: let configure_client
    # report it.
    try:
        text = _read_toml_text(path)
    except Exception as ex:
        raise ValueError('Could not read {} ({}). Save it as UTF-8, then '
                         'Connect again.'.format(path, ex))

    if _toml_other_form(text, name):
        raise ValueError('{} already declares {} as a key under [mcp_servers] (not '
                         'as its own [mcp_servers.{}] table). Remove that entry by '
                         'hand, then Connect again.'.format(path, name, name))

    cfg_dir = os.path.dirname(path)
    if cfg_dir and not os.path.isdir(cfg_dir):
        os.makedirs(cfg_dir)

    _backup_once(path)
    _write_text(path, _toml_upsert(text, _codex_block(python, bridge, port, name, agent), name))


# ─── MCPService ────────────────────────────────────────────────────────────────

class MCPService(object):
    """
    Stateless helper that wraps MCP server + file watcher operations.

    All methods return simple dicts so callers don't need to handle exceptions
    — failures come back as {'error': '<message>'}.

    Typical usage in a dialog:
        status = MCPService.server_status()
        if not status.get('error'):
            print(status['running'], status['port'])

        ok, err = MCPService.start_server()
        ok, err = MCPService.stop_server()
        on, err = MCPService.set_auto_start(False)

        ok, err = MCPService.start_watcher()
        ok, err = MCPService.stop_watcher()

        snippet = MCPService.config_snippet()
    """

    # ── MCP HTTP server ────────────────────────────────────────────────────────

    @staticmethod
    def server_status():
        """
        Return status of the MCP HTTP server.

        Returns:
            dict with keys: running (bool), port (int), tools_count (int),
                            commands_processed (int), foreign (bool),
                            error (str|None)

        `foreign` is True when the live server belongs to ANOTHER pyRevit engine
        in this same Revit process (see the cross-engine note above) and
        cannot be stopped from here. From CPython it is always False: the
        server lives in IronPython by design and Stop is routed there.
        `engine` names the Python that serves the port ('IronPython 2.7' …).
        """
        try:
            server = _get_server()
            stats  = server.get_server_stats()
            if stats.get('running'):
                return {
                    'running':              True,
                    'port':                 stats.get('port', 48884),
                    'tools_count':          stats.get('tools_count', 0),
                    'commands_processed':   stats.get('commands_processed', 0),
                    'external_event_ready': stats.get('external_event_ready', False),
                    'foreign':              False,
                    'engine':               stats.get('engine'),
                    'error':                None,
                }
            own_error = None
        except Exception as ex:
            # An uncallable cross-engine anchor lands here. Never surface the
            # raw AttributeError: the server may well be up and serving.
            own_error = str(ex)

        servers = _local_servers()
        if servers:
            port, info = servers[0]
            return {'running': True, 'port': port,
                    'tools_count': info.get('tools') or 0,
                    'commands_processed': 0, 'external_event_ready': False,
                    'foreign': not _cpython_in_revit(),
                    'engine': info.get('engine'), 'error': None}

        return {'running': False, 'port': 48884, 'tools_count': 0,
                'commands_processed': 0, 'external_event_ready': False,
                'foreign': False, 'engine': None, 'error': own_error}

    @staticmethod
    def ensure_external_event():
        """
        Create the Revit ExternalEvent that marshals model-editing tools onto
        Revit's main thread.

        MUST be called from Revit's UI thread (e.g. a pushbutton's main body),
        NOT from a background worker — ExternalEvent.Create throws outside a
        Revit API context. Without this, every create/modify/rename MCP tool
        fails with "transaction outside API context". Safe and idempotent.

        Returns:
            (success: bool, error_message: str|None)
        """
        try:
            server = _get_server()
            return server.ensure_external_event()
        except Exception as ex:
            return False, str(ex)

    @staticmethod
    def start_server(port=None):
        """
        Start the MCP HTTP server.

        Args:
            port (int|None): Override port. Uses server default if None.

        Returns:
            (success: bool, error_message: str)
        """
        # One Revit process must expose exactly one port. A server started by
        # another engine in this process is invisible to the object above, so
        # without this check start_server() binds a second port — which is how
        # one Revit ended up LISTENING on 48884 AND 48885.
        existing = _local_server_ports()
        if existing:
            return True, None

        # A `#! python3` tool must not host the server itself — it would stop
        # answering once the tool's window closes. IronPython hosts it.
        if _cpython_in_revit():
            return _run_ipy_host('start', port)

        try:
            server = _get_server()
            if port:
                server.port = int(port)
            ok = server.start_server()
            if ok:
                return True, None
            reason = getattr(server, '_start_error', None)
            return False, ('start_server() failed: {}'.format(reason)
                           if reason else 'start_server() returned False')
        except Exception as ex:
            return False, str(ex)

    @staticmethod
    def stop_server():
        """
        Stop the MCP HTTP server.

        Returns:
            (success: bool, error_message: str)
        """
        try:
            server = _get_server()
            ok = server.stop_server()
            if ok:
                return True, None
            err = 'stop_server() returned False'
        except Exception as ex:
            err = str(ex)

        # The object we hold could not stop it. If a server in this process is
        # still answering, it belongs to another pyRevit engine. From CPython
        # that is the IronPython host by design, so ask it to stop. Anywhere
        # else only a Revit restart releases it — say so instead of repeating
        # a raw AttributeError.
        if _local_server_ports():
            if _cpython_in_revit():
                return _run_ipy_host('stop')
            return False, ('The running server was started by another pyRevit '
                           'engine in this Revit session and cannot be stopped '
                           'from here — restart Revit to release it.')
        return False, err

    @staticmethod
    def toggle_server(current_port=None):
        """
        Start server if stopped, stop it if running.

        Returns:
            (new_state: 'running'|'stopped', error_message: str|None)
        """
        status = MCPService.server_status()
        if status.get('error'):
            return 'unknown', status['error']
        if status['running']:
            ok, err = MCPService.stop_server()
            return ('stopped' if ok else 'running'), err
        else:
            ok, err = MCPService.start_server(port=current_port)
            return ('running' if ok else 'stopped'), err

    # ── Auto-start (hooks/doc-opened.py · hooks/doc-created.py) ───────────────

    @staticmethod
    def auto_start_enabled():
        """
        True unless "auto_start_mcp" is explicitly false in mcp_paths.json —
        the same rule the document hooks apply (absent = on).
        """
        try:
            value = _get_paths_module().load_settings().get('auto_start_mcp')
        except Exception:
            return True
        return value is None or bool(value)

    @staticmethod
    def set_auto_start(on):
        """
        Persist the auto-start choice read by the document hooks.

        Returns:
            (state: bool, error_message: str|None) — state is what the file
            holds afterwards, so a failed write reports the old value.
        """
        want = bool(on)
        try:
            _get_paths_module().set_setting('auto_start_mcp', want)
        except Exception as ex:
            return MCPService.auto_start_enabled(), str(ex)
        # save_settings() is best-effort and swallows write errors — read the
        # file back so a locked/read-only settings file is reported, not hidden.
        state = MCPService.auto_start_enabled()
        if state != want:
            return state, 'Could not save the setting to mcp_paths.json'
        return state, None

    # ── Security policy for external MCP clients ──────────────────────────────

    @staticmethod
    def mcp_policy():
        """Current external-client policy: read_only, confirm_destructive,
        code_execution ('off' | 'ask' | 'allow'), toolset ('full' |
        'standard' | 'documentation'). See core/mcp_policy.py."""
        _ensure_lib_in_path()
        from core import mcp_policy
        return mcp_policy.load_policy()

    @staticmethod
    def set_mcp_policy(read_only=None, confirm_destructive=None,
                       code_execution=None, toolset=None):
        """
        Persist one or more policy fields. Applies to the next tool call; no
        restart needed.

        Returns:
            (policy: dict, error_message: str|None) — the policy as stored
            afterwards, so a failed write reports the old values.
        """
        _ensure_lib_in_path()
        from core import mcp_policy
        try:
            mcp_policy.save_policy(read_only=read_only,
                                   confirm_destructive=confirm_destructive,
                                   code_execution=code_execution,
                                   toolset=toolset)
        except Exception as ex:
            return mcp_policy.load_policy(), str(ex)
        stored = mcp_policy.load_policy()
        wanted = {'read_only': read_only,
                  'confirm_destructive': confirm_destructive,
                  'code_execution': code_execution,
                  'toolset': toolset}
        for key, value in wanted.items():
            if value is None:
                continue
            expected = (u'{}'.format(value).strip().lower()
                        if key in ('code_execution', 'toolset') else bool(value))
            if stored.get(key) != expected:
                return stored, 'Could not save the setting to mcp_paths.json'
        return stored, None

    # ── Open documents ─────────────────────────────────────────────────────────

    @staticmethod
    def list_open_documents():
        """
        List documents open in this Revit instance.

        Tool calls always target the ACTIVE document (what the user sees) —
        switching is done through the switch_active_document / open_document
        MCP tools, which activate the target's window for real.

        Returns:
            (documents: list[dict] with keys title/path/is_active, error: str|None)
        """
        try:
            server = _get_server()
            return server.get_open_documents(), None
        except Exception as ex:
            return [], str(ex)

    # ── Teaching capture (Opus distils via MCP) ────────────────────────────────

    @staticmethod
    def teaching_status():
        """Return the server's teaching-capture status dict.

        Keys: enabled (bool), sandbox (str|None), session_open (bool),
        sessions_recorded (int), error (str|None).
        """
        try:
            server = _get_server()
            info = server.get_teaching_status()
            info['error'] = None
            return info
        except Exception as ex:
            return {'enabled': False, 'sandbox': None, 'session_open': False,
                    'sessions_recorded': 0, 'error': str(ex)}

    @staticmethod
    def set_teaching_mode(on):
        """Enable/disable teaching capture. Returns (new_state: bool, err|None)."""
        try:
            server = _get_server()
            return bool(server.set_teaching_mode(bool(on))), None
        except Exception as ex:
            return False, str(ex)


    @staticmethod
    def mark_active_document_as_sandbox():
        """Designate the ACTIVE document as the teaching sandbox — the only doc
        model-modifying tools may touch while teaching mode is on.

        Resolves the active document via the server's open-documents list (safe,
        already used by the dialog). Returns (info: dict|None, error: str|None).
        """
        try:
            server = _get_server()
            docs = server.get_open_documents() or []
            active = None
            for d in docs:
                if d.get('is_active'):
                    active = d
                    break
            if active is None and docs:
                active = docs[0]
            if active is None:
                return None, 'No open document to mark as sandbox.'
            info = {'title': active.get('title') or '',
                    'path':  active.get('path') or ''}
            server.set_sandbox_document(info)
            return info, None
        except Exception as ex:
            return None, str(ex)


    # ── File watcher ───────────────────────────────────────────────────────────

    @staticmethod
    def watcher_status():
        """
        Return status of the file-based task watcher.

        Returns:
            dict with keys: running (bool), data_dir (str),
                            has_ext_event (bool), error (str|None)
        """
        try:
            watcher = _get_watcher()
            info    = watcher.get_status()
            info['error'] = None
            return info
        except Exception as ex:
            return {'running': False, 'data_dir': _get_data_dir(),
                    'has_ext_event': False, 'error': str(ex)}

    @staticmethod
    def start_watcher():
        """
        Start the file task watcher.

        Returns:
            (success: bool, error_message: str)
        """
        try:
            watcher = _get_watcher()
            ok = watcher.start()
            return (True, None) if ok else (False, 'start() returned False')
        except Exception as ex:
            return False, str(ex)

    @staticmethod
    def stop_watcher():
        """
        Stop the file task watcher.

        Returns:
            (success: bool, error_message: str)
        """
        try:
            watcher = _get_watcher()
            watcher.stop()
            return True, None
        except Exception as ex:
            return False, str(ex)

    @staticmethod
    def auto_start_watcher():
        """
        Start the file task watcher only if the user opted in
        ("auto_start_file_watcher" in mcp_paths.json). For automatic starts.

        Returns:
            (started: bool, error_message: str|None)
        """
        try:
            return bool(_get_watcher().start_if_enabled()), None
        except Exception as ex:
            return False, str(ex)

    @staticmethod
    def toggle_watcher():
        """
        Start watcher if stopped, stop it if running.

        Returns:
            (new_state: 'running'|'stopped', error_message: str|None)
        """
        status = MCPService.watcher_status()
        if status.get('error'):
            return 'unknown', status['error']
        # The button is the user's opt-in (or opt-out) for later Revit
        # sessions too: the watcher never starts on its own otherwise.
        try:
            from core.file_watcher import set_auto_start
        except Exception:
            set_auto_start = None
        if status['running']:
            ok, err = MCPService.stop_watcher()
            if ok and set_auto_start:
                set_auto_start(False)
            return ('stopped' if ok else 'running'), err
        else:
            ok, err = MCPService.start_watcher()
            if ok and set_auto_start:
                set_auto_start(True)
            return ('running' if ok else 'stopped'), err

    # ── Config & paths ─────────────────────────────────────────────────────────

    @staticmethod
    def deploy_bridge():
        """
        Copy the extension's core/bridge.py to the machine-stable location
        %APPDATA%/T3LabAI/bridge.py (next to mcp_token.txt / mcp_paths.json).

        The Claude Desktop config points at the deployed copy, never inside
        the extension folder — anyone can download the extension anywhere on
        their machine, and updating/moving/re-cloning it never breaks the MCP
        connection. Content-compared before writing. Called from Connect
        only: it always takes ownership, so startup.py refreshes through
        bridge_deploy.refresh() instead (T3Lab and Lite share the folder).

        hub_store.py (t3lab_hub) goes with it, and this extension becomes the
        bridge's owner, so startup.py keeps it current (lib/bridge_deploy.py).

        Returns:
            (deployed_path: str|None, error: str|None)
        """
        _ensure_lib_in_path()
        import bridge_deploy
        deployed, error, _written = bridge_deploy.deploy(
            _get_paths_module().settings_dir(), os.path.dirname(_source_bridge_path()))
        return deployed, error

    @staticmethod
    def config_snippet(port=None, fmt='json', server='revit'):
        """
        Return a ready-to-paste MCP server entry for this machine.

        Args:
            port (int|None): Port number to embed. Reads from running server if None.
                Ignored for an Autodesk app bridge (it finds its app itself).
            fmt (str): 'json' for the "mcpServers" block used by Claude Desktop,
                Antigravity, Cursor, Windsurf and Claude Code; 'toml' for the
                `[mcp_servers.<name>]` table used by ChatGPT's Codex;
                'claude_code' for a `claude mcp add` command line.
            server (str): 'revit' (t3lab-revit) or an AUTODESK_APPS key.

        Returns:
            str: Formatted snippet.
        """
        name = _server_name(server)
        fmt = str(fmt).lower()
        if name in _BRIDGE_SERVERS:
            python = _find_python_executable()
            bridge = _app_bridge_path(server)
            if fmt == 'toml':
                return _codex_block(python, bridge, 0, name).rstrip('\n')
            if fmt == 'claude_code':
                return 'claude mcp add --scope user {} -- "{}" -I "{}"'.format(
                    name, python, bridge)
            import json as _json
            return _json.dumps(
                {'mcpServers': {name: _server_entry(python.replace('\\', '/'),
                                                    bridge, 0, name)}},
                indent=2)

        if port is None:
            try:
                server = _get_server()
                port   = server.port
            except Exception:
                port = 48884

        bridge = _shown_bridge_path()

        if fmt == 'claude_code':
            return 'claude mcp add --scope user {} -- "{}" -E "{}" {}'.format(
                name, _find_python_executable(), bridge, port)
        if fmt == 'toml':
            # Codex reads real TOML — Windows backslashes must stay escaped.
            return _codex_block(_find_python_executable(), bridge, port).rstrip('\n')

        python = _find_python_executable().replace('\\', '/')
        # The same entry configure_client() writes (-E and an empty
        # PYTHONPATH keep pyRevit's paths out of the bridge's Python), built
        # with a JSON encoder: pasted by hand, a raw Windows path such as
        # C:\Users\... was invalid JSON.
        snippet = {'mcpServers': {
            MCP_SERVER_NAME: _server_entry(python, bridge, port)}}
        try:
            from core import jsonsafe
            return jsonsafe.dumps(snippet, indent=2)
        except Exception:
            import json as _json
            return _json.dumps(snippet, indent=2)

    @staticmethod
    def data_dir():
        """Return the T3Lab_AI_Data directory path (created if absent)."""
        d = _get_data_dir()
        try:
            if not os.path.isdir(d):
                os.makedirs(d)
        except OSError:
            pass
        return d

    @staticmethod
    def open_data_dir():
        """Open the T3Lab_AI_Data directory in the system file explorer."""
        d = MCPService.data_dir()
        try:
            import subprocess
            import platform
            if platform.system() == 'Windows':
                subprocess.Popen(['explorer', d])
            elif platform.system() == 'Darwin':
                subprocess.Popen(['open', d])
            else:
                subprocess.Popen(['xdg-open', d])
            return True, None
        except Exception as ex:
            return False, str(ex)

    # ── AI client auto-configure (Claude Desktop · ChatGPT/Codex · Antigravity) ─

    @staticmethod
    def clients():
        """Return [(key, label, fmt)] for every auto-configurable MCP client."""
        return [(c['key'], c['label'], c['fmt']) for c in AI_CLIENTS]


    @staticmethod
    def autodesk_apps():
        """[(key, label, server name)] of the Autodesk apps the tab lists."""
        return [(a[0], a[1], a[2]) for a in AUTODESK_APPS]

    @staticmethod
    def app_bridge_status(app='autocad'):
        """
        Is this Autodesk app's add-in bridge deployed on this machine?

        Returns:
            dict: installed (bool), path (str), label, server, hint,
            installable (Connect can install it), years (versions found on this PC)
        """
        spec = _APP_BY_KEY[app]
        path = _app_bridge_path(app)
        try:
            from Services import addin_installer
            installable = spec[3] in addin_installer.PACKAGES
            years = addin_installer.installed_years(app)
        except Exception:
            installable, years = False, []
        return {'installed': os.path.isfile(path), 'path': path, 'label': spec[1],
                'server': spec[2], 'hint': spec[4], 'installable': installable,
                'years': years}

    @staticmethod
    def install_addin(app='autocad'):
        """
        Download and install the app's T3Lab add-in (bundle + bridge) from its
        GitHub release (Services/addin_installer.py).

        Returns:
            (success: bool, message: str)
        """
        spec = _APP_BY_KEY[app]
        try:
            from Services import addin_installer
        except Exception as ex:
            return False, 'Add-in installer unavailable: {}'.format(ex)
        if spec[3] not in addin_installer.PACKAGES:
            return False, ('There is no T3Lab add-in for {} — only AutoCAD and Civil 3D '
                           'are supported.'.format(spec[1]))
        try:
            source = _get_paths_module().load_settings().get(spec[3] + '_addin_url')
        except Exception:
            source = None
        try:
            result = addin_installer.install(spec[3], source=source)
        except addin_installer.InstallError as ex:
            return False, '{} add-in not installed: {}'.format(spec[1], ex)
        except Exception as ex:
            return False, '{} add-in install failed: {}'.format(spec[1], ex)
        _addin_refreshed[spec[3]] = (time.time(), result)
        return True, addin_installer.summary(result)

    # Set by configure_client when Connect updated the add-in; MCP Control shows it.
    last_addin_note = ''

    @staticmethod
    def update_addin(app='autocad'):
        """
        Bring an installed add-in up to its latest release (Connect). Never
        fatal: the installed copy keeps working offline.

        Returns:
            str: what changed ('' when nothing did, or the check was skipped).
        """
        spec = _APP_BY_KEY[app]
        last = _addin_refreshed.get(spec[3])
        if last and time.time() - last[0] < _ADDIN_REFRESH_S:
            return ''
        ok, msg = MCPService.install_addin(app)
        if not ok:
            _addin_refreshed[spec[3]] = (time.time(), None)
            return ''
        result = _addin_refreshed[spec[3]][1]
        return msg if (result.get('written') or result.get('locked')) else ''

    @staticmethod
    def cad_bridge_status():
        """AutoCAD's app_bridge_status() (kept for older callers)."""
        return MCPService.app_bridge_status('autocad')

    @staticmethod
    def client_status(key, server='revit'):
        """
        Check one client's configuration status for `server` ('revit' or an AUTODESK_APPS key).

        Returns:
            dict: key, label, fmt, path, file_exists, configured, error
        """
        label, fmt = key, 'json'
        try:
            name  = _server_name(server)
            spec  = _client_spec(key)
            label = spec['label']
            fmt   = spec['fmt']
            path  = _resolve_client_path(spec)
        except Exception as ex:
            return {'key': key, 'label': label, 'fmt': fmt, 'path': '',
                    'file_exists': False, 'configured': False, 'error': str(ex)}

        base = {'key': key, 'label': label, 'fmt': fmt, 'path': path}
        if not os.path.isfile(path):
            base.update({'file_exists': False, 'configured': False, 'error': None})
            return base
        try:
            configured = (_toml_is_configured(path, name) if fmt == 'toml'
                          else _json_is_configured(path, name))
            base.update({'file_exists': True, 'configured': configured, 'error': None})
        except Exception as ex:
            base.update({'file_exists': True, 'configured': False,
                         'error': 'Parse error: {}'.format(ex)})
        return base


    @staticmethod
    def configure_client(key, port=None, server='revit'):
        """
        Write the `server` entry (t3lab-revit / t3lab-autocad) into one
        client's config file.

        Creates the file and its directory if absent, merges with existing
        entries, and keeps a one-time `<path>.t3lab-backup` of the original.
        AutoCAD: refuses when its bridge is not deployed, and refuses Claude
        Desktop while it runs (it would drop the entry on quit).

        Returns:
            (success: bool, message: str) — config path on success, error text
            on failure.
        """
        try:
            name = _server_name(server)
            spec = _client_spec(key)
        except Exception as ex:
            return False, str(ex)

        try:
            python = _find_python_executable()
            if not _python_usable(python):
                return False, NO_PYTHON_TEXT     # else the AI app shows "server disconnected"
            path   = _resolve_client_path(spec)
            if name in _BRIDGE_SERVERS:
                port = 0
                app = _APP_BY_KEY[server]
                bridge = _app_bridge_path(server)
                MCPService.last_addin_note = ''
                if not os.path.isfile(bridge):
                    ok, msg = MCPService.install_addin(server)
                    if not ok:
                        return False, msg
                    MCPService.last_addin_note = msg
                else:
                    MCPService.last_addin_note = MCPService.update_addin(server)
                if not os.path.isfile(bridge):
                    return False, ('{} add-in not installed ({} is missing) — {}'.format(
                        app[1], bridge, app[4]))
                if key == 'claude' and _claude_desktop_running():
                    return False, ('Claude Desktop is running and would drop the entry '
                                   'when it quits — quit it from the tray icon, then Connect again')
            else:
                if port is None:
                    try:
                        port = _get_server().port
                    except Exception:
                        port = 48884
                bridge = _get_bridge_path()

            agent = None if name in _BRIDGE_SERVERS else spec.get('agent')
            if spec['fmt'] == 'toml':
                _write_toml_config(path, python, bridge, port, name, agent)
            else:
                _write_json_config(path, python, bridge, port, name, agent)
            return True, path
        except Exception as ex:
            return False, str(ex)

    @staticmethod
    def upgrade_client_entries():
        """
        Revit start: name the AI app in t3lab-revit entries written before
        Connect added --agent (2.4.5). The config file a client reads says
        which app it is, so nothing is guessed (2026-11-17: Antigravity sends
        no name and showed as "AI app"). Only our own entry is touched, with
        the one-time backup; Claude Desktop is skipped while it runs (it
        rewrites its config when it quits). Never raises.

        Returns:
            list[str]: labels of the apps whose entry was updated.
        """
        done = []
        for spec in AI_CLIENTS:
            try:
                agent = spec.get('agent')
                path = _resolve_client_path(spec)
                if not agent or not os.path.isfile(path):
                    continue
                if spec['key'] == 'claude' and _claude_desktop_running():
                    continue
                add = _toml_add_agent if spec['fmt'] == 'toml' else _json_add_agent
                if add(path, agent):
                    done.append(spec['label'])
            except Exception:
                continue
        return done

    @staticmethod
    def configure_all_clients(port=None, server='revit'):
        """
        Configure every registered client in one go.

        Returns:
            list[dict]: key, label, ok, message — one entry per client.
        """
        results = []
        for spec in AI_CLIENTS:
            ok, msg = MCPService.configure_client(spec['key'], port=port, server=server)
            results.append({'key': spec['key'], 'label': spec['label'],
                            'ok': ok, 'message': msg})
        return results

    # ── Backwards-compatible Claude Desktop wrappers ──────────────────────────


    # ── Combined snapshot (for dashboard widgets) ──────────────────────────────

