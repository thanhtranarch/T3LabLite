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


if __name__ == '__main__':
    unittest.main()
