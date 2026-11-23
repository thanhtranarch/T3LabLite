# -*- coding: utf-8 -*-
"""
CPython 3 test: a PC without Python 3 gets a private one for the MCP bridge
(lib/Services/python_installer.py), and only then. No network: packages are
built here.

Run:  python3 lite_guard/test_python_installer.py   (exit 0 = all pass)
"""
import json
import os
import shutil
import sys
import tempfile
import unittest
import zipfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'lib'))

from Services import python_installer as PI  # noqa: E402
import Services.mcp_service as M  # noqa: E402
from core import paths  # noqa: E402


def make_pkg(path, version='3.15.2', exe=True, extra=None):
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('python.nuspec', '<package><metadata><id>python</id>'
                                    '<version>{}</version></metadata></package>'.format(version))
        z.writestr('[Content_Types].xml', '<Types/>')
        if exe:
            z.writestr('tools/python.exe', 'MZ')
        z.writestr('tools/Lib/os.py', '# os')
        for name, data in (extra or {}).items():
            z.writestr(name, data)
    return path


class _Temp(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix='t3lab_py_test_')
        self.env = {k: os.environ.get(k) for k in ('HOME', 'USERPROFILE', 'APPDATA',
                                                   'CLAUDE_CONFIG_DIR', 'CODEX_HOME')}
        os.environ['HOME'] = os.environ['USERPROFILE'] = self.dir
        os.environ['APPDATA'] = os.path.join(self.dir, 'AppData')
        os.environ.pop('CLAUDE_CONFIG_DIR', None)
        os.environ.pop('CODEX_HOME', None)
        self.root = os.path.join(self.dir, 'AppData', 'T3LabAI', 'python')
        self.pkg = os.path.join(self.dir, 'python.nupkg')

    def tearDown(self):
        for k, v in self.env.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        shutil.rmtree(self.dir, ignore_errors=True)


class InstallerTests(_Temp):
    def test_only_final_releases_newest_last(self):
        self.assertEqual(PI.stable_versions(['3.9.13', '3.15.2', '3.16.0-a2', '3.15.0-rc1',
                                             '3.10.1', 'junk']),
                         ['3.9.13', '3.10.1', '3.15.2'])

    def test_installs_only_tools_into_the_private_folder(self):
        exe = PI.install(source=make_pkg(self.pkg))
        self.assertEqual(exe, os.path.join(self.root, '3.15.2', 'tools', 'python.exe'))
        self.assertTrue(os.path.isfile(os.path.join(self.root, '3.15.2', 'tools', 'Lib', 'os.py')))
        self.assertEqual(os.listdir(os.path.join(self.root, '3.15.2')), ['tools'])
        self.assertEqual(os.listdir(self.root), ['3.15.2'])          # no staging left
        self.assertEqual(PI.find(), exe)

    def test_find_takes_the_newest_complete_copy(self):
        self.assertEqual(PI.find(), '')
        PI.install(source=make_pkg(self.pkg, '3.9.13'))
        newest = PI.install(source=make_pkg(self.pkg, '3.15.2'))
        os.makedirs(os.path.join(self.root, '3.16.0', 'tools'))      # no python.exe in it
        self.assertEqual(PI.find(), newest)

    def test_bad_packages_are_refused_and_leave_nothing(self):
        cases = [make_pkg(os.path.join(self.dir, 'a.nupkg'), exe=False),
                 make_pkg(os.path.join(self.dir, 'b.nupkg'), version='3.16.0-a2'),
                 make_pkg(os.path.join(self.dir, 'c.nupkg'), extra={'tools/../../evil.py': 'x'})]
        notzip = os.path.join(self.dir, 'd.nupkg')
        with open(notzip, 'w') as f:
            f.write('hello')
        for src in cases + [notzip]:
            with self.assertRaises(PI.InstallError):
                PI.install(source=src)
        self.assertEqual(PI.find(), '')
        self.assertFalse(os.path.exists(os.path.join(self.dir, 'evil.py')))

    def test_source_must_be_https_or_a_file(self):
        with self.assertRaises(PI.InstallError):
            PI.install(source='http://example.invalid/python.nupkg')


class ServiceTests(_Temp):
    def setUp(self):
        _Temp.setUp(self)
        self.saved = {n: getattr(M, n) for n in ('_find_python_executable', '_get_bridge_path',
                                                  '_claude_desktop_running')}
        self.supported = PI.supported
        PI.supported = lambda: True
        M._get_bridge_path = lambda take_ownership=True: 'C:/T3LabAI/bridge.py'
        M._claude_desktop_running = lambda: False
        os.makedirs(os.path.join(self.dir, '.claude'))               # Claude Code is installed
        make_pkg(self.pkg)
        paths.set_setting('python_package_url', self.pkg)

    def tearDown(self):
        PI.supported = self.supported
        for n, v in self.saved.items():
            setattr(M, n, v)
        _Temp.tearDown(self)

    def claude_json(self):
        return os.path.join(self.dir, '.claude.json')

    def test_startup_asks_for_python_then_connects(self):
        M._find_python_executable = lambda: 'python'                  # none on this PC
        first = M.MCPService.auto_connect_clients()
        self.assertEqual([c.get('need_python') for c in first], [True])
        self.assertFalse(os.path.exists(self.claude_json()))
        exe, _msg = M.MCPService.install_python()
        self.assertTrue(os.path.isfile(exe))
        M._find_python_executable = self.saved['_find_python_executable']
        self.assertEqual(M._find_python_executable(), exe)
        done = M.MCPService.auto_connect_clients()
        self.assertEqual([(c['label'], c['ok']) for c in done], [('Claude Code', True)])
        with open(self.claude_json(), encoding='utf-8') as f:
            entry = json.load(f)['mcpServers'][M.MCP_SERVER_NAME]
        self.assertEqual(entry['command'], exe)
        self.assertIn('--agent=claude-code', entry['args'])

    def test_no_ai_app_means_no_python_download(self):
        shutil.rmtree(os.path.join(self.dir, '.claude'))
        M._find_python_executable = lambda: 'python'
        self.assertEqual(M.MCPService.auto_connect_clients(), [])

    def test_connect_click_installs_python(self):
        M._find_python_executable = lambda: 'python'
        ok, msg = M.MCPService.configure_client('claudecode')
        self.assertTrue(ok, msg)
        with open(self.claude_json(), encoding='utf-8') as f:
            command = json.load(f)['mcpServers'][M.MCP_SERVER_NAME]['command']
        self.assertEqual(command, PI.find())

    def test_switched_off_in_settings(self):
        paths.set_setting('auto_install_python', False)
        exe, msg = M.MCPService.install_python()
        self.assertEqual(exe, '')
        self.assertIn('turned off', msg)
        self.assertEqual(PI.find(), '')

    def test_failure_is_a_message_not_an_exception(self):
        paths.set_setting('python_package_url', 'http://example.invalid/x.nupkg')
        exe, msg = M.MCPService.install_python()
        self.assertEqual(exe, '')
        self.assertIn('not installed', msg)

    def test_claude_code_keeps_its_other_settings_in_order(self):
        with open(self.claude_json(), 'w', encoding='utf-8') as f:
            f.write('{"zeta": 1, "numStartups": 7, "mcpServers": {"other": {"command": "x"}}, '
                    '"alpha": {"b": 1, "a": 2}}')
        ok, msg = M.MCPService.configure_client('claudecode')
        self.assertTrue(ok, msg)
        with open(self.claude_json(), encoding='utf-8') as f:
            data = json.load(f)
        self.assertEqual(list(data), ['zeta', 'numStartups', 'mcpServers', 'alpha'])
        self.assertEqual(list(data['alpha']), ['b', 'a'])
        self.assertEqual(list(data['mcpServers']), ['other', M.MCP_SERVER_NAME])
        self.assertTrue(os.path.isfile(self.claude_json() + '.t3lab-backup'))

    def test_claude_code_not_installed_gets_no_file(self):
        shutil.rmtree(os.path.join(self.dir, '.claude'))
        M._find_python_executable = lambda: self.pkg                  # any real file
        self.assertEqual(M.MCPService.auto_connect_clients(), [])
        self.assertFalse(os.path.exists(self.claude_json()))

    def test_claude_config_dir_is_followed(self):
        os.environ['CLAUDE_CONFIG_DIR'] = os.path.join(self.dir, 'cc')
        os.makedirs(os.environ['CLAUDE_CONFIG_DIR'])
        M._find_python_executable = lambda: self.pkg
        done = M.MCPService.auto_connect_clients()
        self.assertEqual([c['ok'] for c in done], [True])
        self.assertTrue(os.path.isfile(os.path.join(self.dir, 'cc', '.claude.json')))
        self.assertFalse(os.path.exists(self.claude_json()))


if __name__ == '__main__':
    unittest.main()
