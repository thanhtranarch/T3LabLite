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
        self.saved = {n: getattr(M, n) for n in ('_find_python_executable', '_get_bridge_path',
                                                  '_claude_desktop_running')}
        M._find_python_executable = lambda: PY
        M._get_bridge_path = lambda take_ownership=True: 'C:/T3LabAI/bridge.py'
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


class ClaudeDesktopTests(_TempHome):
    """Store / MSIX Claude Desktop reads its own copy of the config."""

    def setUp(self):
        _TempHome.setUp(self)
        import platform
        self.system = platform.system
        platform.system = lambda: 'Windows'
        self.local = os.environ.get('LOCALAPPDATA')
        os.environ['LOCALAPPDATA'] = os.path.join(self.dir, 'Local')

    def tearDown(self):
        import platform
        platform.system = self.system
        if self.local is None:
            os.environ.pop('LOCALAPPDATA', None)
        else:
            os.environ['LOCALAPPDATA'] = self.local
        _TempHome.tearDown(self)

    def classic(self):
        return os.path.join(self.dir, 'AppData', 'Claude', 'claude_desktop_config.json')

    def msix(self):
        return os.path.join(self.dir, 'Local', 'Packages', 'Claude_pzs8sxrjxfjjc', 'LocalCache',
                            'Roaming', 'Claude', 'claude_desktop_config.json')

    def test_classic_install_uses_appdata(self):
        ok, msg = M.MCPService.configure_client('claude')
        self.assertTrue(ok, msg)
        self.assertEqual(os.path.normpath(msg), os.path.normpath(self.classic()))

    def test_msix_copy_is_the_one_written(self):
        _write(self.msix(), b'{"mcpServers": {"other": {"command": "x"}}}')
        ok, msg = M.MCPService.configure_client('claude')
        self.assertTrue(ok, msg)
        self.assertEqual(os.path.normpath(msg), os.path.normpath(self.msix()))
        servers = json.loads(_read(self.msix()).decode('utf-8'))['mcpServers']
        self.assertIn('other', servers)
        self.assertIn(M.MCP_SERVER_NAME, servers)
        self.assertFalse(os.path.exists(self.classic()))

    def test_msix_folder_without_file_is_used(self):
        os.makedirs(os.path.dirname(self.msix()))
        ok, msg = M.MCPService.configure_client('claude')
        self.assertTrue(ok, msg)
        self.assertTrue(os.path.isfile(self.msix()))

    def test_msix_copy_replaces_a_remembered_classic_path(self):
        ok, _ = M.MCPService.configure_client('claude')          # classic, remembered
        self.assertTrue(ok)
        self.assertTrue(M.MCPService.client_status('claude')['configured'])
        _write(self.msix(), b'{}')                               # Store build installed
        status = M.MCPService.client_status('claude')
        self.assertEqual(os.path.normpath(status['path']), os.path.normpath(self.msix()))
        self.assertFalse(status['configured'])

    def test_hand_set_path_still_wins(self):
        custom = os.path.join(self.dir, 'my', 'claude.json')
        _write(custom, b'{}')
        _write(self.msix(), b'{}')
        paths.set_setting('claude_desktop_config', custom)
        self.assertEqual(M.MCPService.client_status('claude')['path'], custom)

    def test_revit_refused_while_claude_desktop_runs(self):
        M._claude_desktop_running = lambda: True
        ok, msg = M.MCPService.configure_client('claude')
        self.assertFalse(ok)
        self.assertIn('quit', msg.lower())
        self.assertFalse(os.path.exists(self.classic()))
        ok, msg = M.MCPService.configure_client('antigravity')
        self.assertTrue(ok, msg)


class AutoConnectTests(ClaudeDesktopTests):
    """Revit startup connects the installed AI apps by itself."""

    def setUp(self):
        ClaudeDesktopTests.setUp(self)
        self.py = os.path.join(self.dir, 'python.exe')
        _write(self.py, b'')
        M._find_python_executable = lambda: self.py

    def servers(self, path):
        return json.loads(_read(path).decode('utf-8')).get('mcpServers', {})

    def test_installed_app_is_connected_once(self):
        _write(self.classic(), b'{"mcpServers": {"other": {"command": "x"}}}')
        done = M.MCPService.auto_connect_clients()
        self.assertEqual([(d['label'], d['ok']) for d in done], [('Claude Desktop', True)])
        self.assertEqual(sorted(self.servers(self.classic())), ['other', M.MCP_SERVER_NAME])
        self.assertEqual(M.MCPService.auto_connect_clients(), [])        # nothing left to do

    def test_app_folder_without_a_config_is_connected(self):
        os.makedirs(os.path.dirname(self.msix()))
        done = M.MCPService.auto_connect_clients()
        self.assertEqual(len(done), 1)
        self.assertIn(M.MCP_SERVER_NAME, self.servers(self.msix()))

    def test_app_that_is_not_installed_gets_no_config(self):
        self.assertEqual(M.MCPService.auto_connect_clients(), [])
        self.assertFalse(os.path.exists(os.path.dirname(self.classic())))
        self.assertFalse(os.path.exists(os.path.join(self.dir, '.codex')))
        self.assertFalse(os.path.exists(os.path.join(self.dir, '.gemini')))

    def test_existing_entry_on_another_port_is_kept(self):
        _write(self.classic(), b'{}')
        ok, _ = M.MCPService.configure_client('claude', port=48885)
        self.assertTrue(ok)
        before = _read(self.classic())
        self.assertEqual(M.MCPService.auto_connect_clients(port=48884), [])
        self.assertEqual(_read(self.classic()), before)

    def test_running_claude_desktop_is_skipped_not_written(self):
        _write(self.classic(), b'{}')
        M._claude_desktop_running = lambda: True
        done = M.MCPService.auto_connect_clients()
        self.assertEqual([d['ok'] for d in done], [False])
        self.assertEqual(_read(self.classic()), b'{}')

    def test_unreadable_config_is_left_alone(self):
        _write(self.classic(), b'{ not json')
        self.assertEqual(M.MCPService.auto_connect_clients(), [])
        self.assertEqual(_read(self.classic()), b'{ not json')

    def test_no_python_writes_nothing(self):
        _write(self.classic(), b'{}')
        M._find_python_executable = lambda: 'python'
        done = M.MCPService.auto_connect_clients()
        self.assertEqual([d['ok'] for d in done], [False])
        self.assertEqual(_read(self.classic()), b'{}')

    def test_switched_off_in_settings(self):
        _write(self.classic(), b'{}')
        paths.set_setting('auto_connect_clients', False)
        self.assertEqual(M.MCPService.auto_connect_clients(), [])
        self.assertEqual(_read(self.classic()), b'{}')

    def test_autocad_entry_only_with_its_addin_and_no_download(self):
        _write(self.classic(), b'{}')
        calls = []
        real = M.MCPService.update_addin
        M.MCPService.update_addin = staticmethod(lambda *a, **k: calls.append(a) or '')
        self.addCleanup(setattr, M.MCPService, 'update_addin', real)
        M.MCPService.auto_connect_clients()
        self.assertNotIn(M.CAD_SERVER_NAME, self.servers(self.classic()))
        _write(os.path.join(self.dir, 'AppData', 'T3LabCAD', 'bridge.py'), b'# bridge')
        done = M.MCPService.auto_connect_clients()
        self.assertEqual([(d['server'], d['ok']) for d in done], [(M.CAD_SERVER_NAME, True)])
        self.assertIn(M.CAD_SERVER_NAME, self.servers(self.classic()))
        self.assertEqual(calls, [])


class UpdateButtonTests(ClaudeDesktopTests):
    def test_unchanged_entry_is_fine_while_claude_desktop_runs(self):
        ok, _ = M.MCPService.configure_client('claude')
        self.assertTrue(ok)
        before = _read(self.classic())
        M._claude_desktop_running = lambda: True
        ok, msg = M.MCPService.configure_client('claude')
        self.assertTrue(ok, msg)
        self.assertEqual(_read(self.classic()), before)
        ok, msg = M.MCPService.configure_client('claude', port=48890)   # a change: refused
        self.assertFalse(ok)


if __name__ == '__main__':
    unittest.main()
