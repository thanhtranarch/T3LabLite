# -*- coding: utf-8 -*-
"""
CPython 3 test: MCP Control never destroys a config file it cannot read.
No Revit needed; every file lives in a temp HOME / APPDATA.

  - AI-app configs (Claude Desktop / Antigravity JSON, Codex TOML): a BOM is
    read, other servers survive; a file that does not parse is refused and
    left byte-for-byte unchanged (it used to be replaced by our entry alone).
  - Client paths: opening MCP Control before an AI app exists does not pin a
    path; a pinned path that does not exist yields to the real file.
  - mcp_paths.json: a BOM is read; a file that does not parse is kept as
    .unreadable before it is overwritten.

Run:  python3 lite_guard/test_mcp_config_safety.py   (exit 0 = all pass)
"""
import codecs
import json
import os
import shutil
import sys
import tempfile
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'lib'))

import Services.mcp_service as M  # noqa: E402
from core import paths  # noqa: E402

PY = 'C:\\Py\\python.exe'


def _read(path):
    with open(path, 'rb') as f:
        return f.read()


def _write(path, data):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'wb') as f:
        f.write(data)


class _TempHome(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix='t3lab_cfg_')
        self.env = {k: os.environ.get(k) for k in ('HOME', 'USERPROFILE', 'APPDATA', 'CODEX_HOME')}
        os.environ['HOME'] = os.environ['USERPROFILE'] = self.dir
        os.environ['APPDATA'] = os.path.join(self.dir, 'AppData')
        os.environ.pop('CODEX_HOME', None)
        self.saved = {n: getattr(M, n) for n in ('_find_python_executable', '_python_usable', '_get_bridge_path',
                                                  '_claude_desktop_running')}
        M._find_python_executable = lambda: PY
        M._python_usable = lambda python: True     # PY is not a file here
        M._get_bridge_path = lambda: 'C:/T3LabAI/bridge.py'
        M._claude_desktop_running = lambda: False

    def tearDown(self):
        for n, v in self.saved.items():
            setattr(M, n, v)
        for k, v in self.env.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        shutil.rmtree(self.dir, ignore_errors=True)

    def ag_live(self):
        return os.path.join(self.dir, '.gemini', 'antigravity-ide', 'mcp_config.json')


class ClientConfigTests(_TempHome):
    def test_bom_json_keeps_other_servers(self):
        path = self.ag_live()
        _write(path, codecs.BOM_UTF8 + json.dumps(
            {'mcpServers': {'other': {'command': 'x'}}, 'globalShortcut': 'Ctrl+K'}).encode('utf-8'))
        ok, msg = M.MCPService.configure_client('antigravity')
        self.assertTrue(ok, msg)
        with open(path, encoding='utf-8') as f:
            data = json.load(f)
        self.assertIn('other', data['mcpServers'])
        self.assertIn('t3lab-revit', data['mcpServers'])
        self.assertEqual(data['globalShortcut'], 'Ctrl+K')

    def test_unparseable_json_is_refused_and_untouched(self):
        path = self.ag_live()
        original = b'{\n  // my servers\n  "mcpServers": {"other": {"command": "x"}},\n}\n'
        _write(path, original)
        ok, msg = M.MCPService.configure_client('antigravity')
        self.assertFalse(ok)
        self.assertIn('not valid JSON', msg)
        self.assertEqual(_read(path), original)

    def test_unreadable_toml_is_refused_and_untouched(self):
        path = os.path.join(self.dir, '.codex', 'config.toml')
        original = b'# caf\xe9 (saved as ANSI)\n[model_providers.x]\nname = "x"\n'
        _write(path, original)
        ok, msg = M.MCPService.configure_client('chatgpt')
        self.assertFalse(ok)
        self.assertEqual(_read(path), original)


class ClientPathTests(_TempHome):
    def spec(self):
        return M._client_spec('antigravity')

    def test_nothing_installed_pins_nothing(self):
        path = M._resolve_client_path(self.spec())
        self.assertEqual(path, self.ag_live())
        self.assertNotIn('antigravity_config', paths.load_settings())

    def test_stale_pin_yields_to_the_real_file(self):
        paths.set_setting('antigravity_config',
                          os.path.join(self.dir, '.gemini', 'config', 'mcp_config.json'))
        _write(self.ag_live(), b'{}')
        self.assertEqual(M._resolve_client_path(self.spec()), self.ag_live())
        self.assertEqual(paths.load_settings()['antigravity_config'], self.ag_live())

    def test_existing_pin_wins(self):
        mine = os.path.join(self.dir, 'custom', 'mcp_config.json')
        _write(mine, b'{}')
        _write(self.ag_live(), b'{}')
        paths.set_setting('antigravity_config', mine)
        self.assertEqual(M._resolve_client_path(self.spec()), mine)


class SettingsFileTests(_TempHome):
    def test_bom_settings_are_read(self):
        _write(paths.settings_file(), codecs.BOM_UTF8 + b'{"mcp_read_only": true}')
        self.assertTrue(paths.load_settings().get('mcp_read_only'))

    def test_unreadable_settings_are_kept_before_overwrite(self):
        bad = b'{"mcp_read_only": true,,}'
        _write(paths.settings_file(), bad)
        paths.set_setting('x', 1)
        self.assertEqual(_read(paths.settings_file() + '.unreadable'), bad)
        self.assertEqual(paths.load_settings(), {'x': 1})

    def test_non_object_root_is_ignored(self):
        _write(paths.settings_file(), b'"x"')
        self.assertEqual(paths.load_settings(), {})
        paths.set_setting('x', 1)
        self.assertEqual(paths.load_settings(), {'x': 1})


class NoPythonTests(_TempHome):
    """Connect must not write a command that cannot start (2026-11-18: the
    Store alias python.exe was written; the AI app showed "server disconnected")."""

    def test_connect_is_refused_without_python(self):
        M._python_usable = self.saved['_python_usable']
        M._find_python_executable = lambda: 'python'            # nothing found: the bare name
        ok, msg = M.MCPService.configure_client('antigravity')
        self.assertFalse(ok)
        self.assertIn('Python 3 was not found', msg)
        self.assertFalse(os.path.exists(self.ag_live()))

    def test_store_alias_alone_is_not_python(self):
        usable, is_stub = self.saved['_python_usable'], M._is_store_stub
        self.assertTrue(is_stub('C:\\Users\\a\\AppData\\Local\\Microsoft\\WindowsApps\\python.exe'))
        self.assertFalse(is_stub('C:\\Python312\\python.exe'))
        local = os.path.join(self.dir, 'Local')
        apps = os.path.join(local, 'Microsoft', 'WindowsApps')
        alias, real = os.path.join(apps, 'python.exe'), os.path.join(self.dir, 'Py312', 'python.exe')
        _write(alias, b'')
        _write(real, b'MZ')
        saved = os.environ.get('LOCALAPPDATA')
        os.environ['LOCALAPPDATA'] = local
        M._is_store_stub = lambda p: p == alias         # Windows separators are tested above
        try:
            self.assertTrue(usable(real))
            self.assertFalse(usable('python'))          # the bare name: nothing was found
            self.assertFalse(usable(alias))             # the alias only opens the Store
            os.makedirs(os.path.join(apps, 'PythonSoftwareFoundation.Python.3.12_qbz5n2kfra8p0'))
            self.assertTrue(usable(alias))              # Store Python is installed: it works
        finally:
            M._is_store_stub = is_stub
            if saved is None:
                os.environ.pop('LOCALAPPDATA', None)
            else:
                os.environ['LOCALAPPDATA'] = saved

class CodexTomlTests(_TempHome):
    def toml(self):
        return os.path.join(self.dir, '.codex', 'config.toml')

    def test_agent_is_added_after_a_trailing_comma(self):
        _write(self.toml(), b'[mcp_servers.t3lab-revit]\ncommand = "py"\n'
                            b'args = ["C:/T3LabAI/bridge.py", "48884",]\n')
        self.assertTrue(M._toml_add_agent(self.toml(), 'codex'))
        text = _read(self.toml()).decode('utf-8')
        self.assertIn('args = ["C:/T3LabAI/bridge.py", "48884", "--agent=codex"]', text)
        self.assertNotIn(',,', text)
        self.assertFalse(M._toml_add_agent(self.toml(), 'codex'))      # second start: no rewrite

    def test_entry_in_another_form_is_not_declared_twice(self):
        original = b'[mcp_servers]\nt3lab-revit = { command = "py", args = ["x/bridge.py"] }\n'
        _write(self.toml(), original)
        ok, msg = M.MCPService.configure_client('chatgpt')
        self.assertFalse(ok)
        self.assertIn('Remove that entry by hand', msg)
        self.assertEqual(_read(self.toml()), original)

    def test_codex_runs_the_bridge_without_pyrevits_pythonpath(self):
        ok, msg = M.MCPService.configure_client('chatgpt')
        self.assertTrue(ok, msg)
        self.assertIn('args = ["-E", ', _read(self.toml()).decode('utf-8'))


class StalePathsModuleTests(unittest.TestCase):
    """An update applied without a Revit restart left the old core.paths in
    sys.modules: MCP Control failed with "module 'core.paths' has no attribute
    'atomic_write'". _get_paths_module() reloads it from disk instead."""

    def test_stale_module_is_reloaded(self):
        saved = paths.atomic_write
        del paths.atomic_write
        try:
            fresh = M._get_paths_module()
            self.assertTrue(callable(getattr(fresh, 'atomic_write', None)))
        finally:
            if not hasattr(paths, 'atomic_write'):
                paths.atomic_write = saved

    def test_current_module_is_not_reloaded(self):
        import importlib
        calls = []
        real = importlib.reload
        importlib.reload = lambda m: calls.append(m) or m
        try:
            M._get_paths_module()
        finally:
            importlib.reload = real
        self.assertEqual(calls, [])

    def test_bootstrap_hot_reloads_paths(self):
        import _cpython_bootstrap as B
        self.assertIn('core.paths', B.HOT_RELOAD_MODULES)


if __name__ == '__main__':
    unittest.main()
