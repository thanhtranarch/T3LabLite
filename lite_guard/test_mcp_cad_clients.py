# -*- coding: utf-8 -*-
"""
CPython 3 test: MCP Control registers the AutoCAD add-in's t3lab-autocad
server next to t3lab-revit (Services/mcp_service.py, AutoCAD tab).

Run:  python3 lite_guard/test_mcp_cad_clients.py
Exit 0 = all pass. Every file lives in a temp dir; real configs are never read.
"""
import json
import os
import shutil
import sys
import tempfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)

import Services.mcp_service as M

FAILURES = []
PY = 'C:\\Py\\python.exe'


def check(name, cond, detail=''):
    if cond:
        print('  ok   ' + name)
    else:
        print('  FAIL ' + name + ('  -- ' + str(detail) if detail else ''))
        FAILURES.append(name)


class Env(object):
    """Temp APPDATA + client config paths, patched into the module."""

    def __enter__(self):
        self.dir = tempfile.mkdtemp(prefix='t3lab_cad_')
        self.paths = {
            'claude': os.path.join(self.dir, 'Claude', 'claude_desktop_config.json'),
            'chatgpt': os.path.join(self.dir, 'codex', 'config.toml'),
            'claudecode': os.path.join(self.dir, 'cc', '.claude.json'),
            'antigravity': os.path.join(self.dir, 'ag', 'mcp_config.json'),
        }
        self.bridge = os.path.join(self.dir, 'T3LabCAD', 'bridge.py')
        self.saved = {n: getattr(M, n) for n in (
            '_resolve_client_path', '_find_python_executable', '_app_bridge_path',
            '_claude_desktop_running', '_get_bridge_path')}
        self.desktop_running = False
        M._resolve_client_path = lambda spec: self.paths[spec['key']]
        M._find_python_executable = lambda: PY
        self.app_bridges = {}      # other apps: key -> path (default: none installed)
        M._app_bridge_path = lambda app='autocad': (
            self.bridge if app == 'autocad'
            else self.app_bridges.get(app, os.path.join(self.dir, app, 'bridge.py'))
        ).replace('\\', '/')
        M._claude_desktop_running = lambda: self.desktop_running
        M._get_bridge_path = lambda take_ownership=True: 'C:/T3LabAI/bridge.py'
        self.real_update = M.MCPService.update_addin       # no download in these tests
        M.MCPService.update_addin = staticmethod(lambda app='autocad', force=False: '')
        return self

    def install_bridge(self):
        os.makedirs(os.path.dirname(self.bridge), exist_ok=True)
        with open(self.bridge, 'w') as f:
            f.write('# bridge')

    def __exit__(self, *exc):
        for n, v in self.saved.items():
            setattr(M, n, v)
        M.MCPService.update_addin = self.real_update
        shutil.rmtree(self.dir, ignore_errors=True)


def _json(path):
    with open(path, encoding='utf-8') as f:
        return json.load(f)


def test_needs_bridge():
    print('[autocad: refuses when the add-in bridge is missing]')
    with Env() as env:
        real_install = M.MCPService.install_addin       # no download in this test
        M.MCPService.install_addin = staticmethod(
            lambda app='autocad': (False, 'AutoCAD add-in not installed: offline'))
        try:
            ok, msg = M.MCPService.configure_client('claude', server='autocad')
        finally:
            M.MCPService.install_addin = real_install
        check('refused', not ok)
        check('says the add-in is not installed', 'add-in not installed' in msg, msg)
        check('no file written', not os.path.exists(env.paths['claude']))
        check('bridge status not installed',
              M.MCPService.cad_bridge_status()['installed'] is False)


def test_json_keeps_revit_and_others():
    print('[autocad: JSON upsert keeps t3lab-revit and other servers]')
    with Env() as env:
        env.install_bridge()
        os.makedirs(os.path.dirname(env.paths['claude']))
        revit = {'command': PY, 'args': ['C:/T3LabAI/bridge.py', '48884']}
        with open(env.paths['claude'], 'w', encoding='utf-8') as f:
            json.dump({'mcpServers': {'t3lab-revit': revit, 'other': {'command': 'npx'}},
                       'theme': 'dark'}, f)
        check('revit not autocad yet',
              not M.MCPService.client_status('claude', server='autocad')['configured'])
        ok, msg = M.MCPService.configure_client('claude', server='autocad')
        check('configured', ok, msg)
        data = _json(env.paths['claude'])
        servers = data['mcpServers']
        check('revit entry untouched', servers.get('t3lab-revit') == revit)
        check('other server kept', servers.get('other') == {'command': 'npx'})
        check('top-level key kept', data.get('theme') == 'dark')
        cad = servers.get('t3lab-autocad') or {}
        check('cad args -I bridge', cad.get('args') == ['-I', env.bridge.replace('\\', '/')],
              str(cad))
        check('cad command', cad.get('command') == PY)
        check('status configured',
              M.MCPService.client_status('claude', server='autocad')['configured'])
        check('an old revit entry (no --agent) asks for Connect again',
              M.MCPService.client_status('claude')['configured'] is False)
        servers['t3lab-revit'] = dict(revit, args=revit['args'] + ['--agent=claude-desktop'])
        with open(env.paths['claude'], 'w') as f:
            json.dump({'mcpServers': servers}, f)
        check('revit entry with --agent is configured',
              M.MCPService.client_status('claude')['configured'] is True)
        check('backup written', os.path.isfile(env.paths['claude'] + '.t3lab-backup'))


def test_desktop_running_guard():
    print('[autocad: no Claude Desktop write while it runs]')
    with Env() as env:
        env.install_bridge()
        env.desktop_running = True
        ok, msg = M.MCPService.configure_client('claude', server='autocad')
        check('refused', not ok)
        check('says quit', 'quit' in msg.lower(), msg)
        check('no file written', not os.path.exists(env.paths['claude']))
        ok, msg = M.MCPService.configure_client('antigravity', server='autocad')
        check('other JSON client still written', ok, msg)


def test_toml():
    print('[autocad: Codex TOML table next to t3lab-revit]')
    with Env() as env:
        env.install_bridge()
        os.makedirs(os.path.dirname(env.paths['chatgpt']))
        with open(env.paths['chatgpt'], 'w', encoding='utf-8') as f:
            f.write('model = "gpt-5"\n\n[mcp_servers.t3lab-revit]\ntype = "stdio"\n'
                    'command = "x"\nargs = ["b", "48884"]\n')
        ok, msg = M.MCPService.configure_client('chatgpt', server='autocad')
        check('configured', ok, msg)
        with open(env.paths['chatgpt'], encoding='utf-8') as f:
            text = f.read()
        check('model kept', text.startswith('model = "gpt-5"'))
        check('revit table kept', '[mcp_servers.t3lab-revit]' in text and '"48884"' in text)
        check('cad table', '[mcp_servers.t3lab-autocad]' in text)
        check('cad args', 'args = ["-I", "{}"]'.format(env.bridge.replace('\\', '/')) in text, text)
        ok, _ = M.MCPService.configure_client('chatgpt', server='autocad')
        with open(env.paths['chatgpt'], encoding='utf-8') as f:
            again = f.read()
        check('idempotent', again.count('[mcp_servers.t3lab-autocad]') == 1)
        check('status configured',
              M.MCPService.client_status('chatgpt', server='autocad')['configured'])


def test_snippets_and_revit_default():
    print('[snippets + Revit behaviour unchanged]')
    with Env() as env:
        snippet = json.loads(M.MCPService.config_snippet(fmt='json', server='autocad'))
        cad = snippet['mcpServers']['t3lab-autocad']
        check('json snippet args', cad['args'][0] == '-I')
        check('toml snippet', '[mcp_servers.t3lab-autocad]' in
              M.MCPService.config_snippet(fmt='toml', server='autocad'))
        cmd = M.MCPService.config_snippet(fmt='claude_code', server='autocad')
        check('claude code command', cmd.startswith('claude mcp add --scope user t3lab-autocad -- '))
        revit = M.MCPService.config_snippet(port=48884, fmt='json')
        check('revit snippet unchanged', '"t3lab-revit"' in revit and '"48884"' in revit)
        ok, msg = M.MCPService.configure_client('antigravity', port=48885)
        data = _json(env.paths['antigravity'])
        check('revit configure default server',
              ok and data['mcpServers']['t3lab-revit']['args'] == [
                  '-E', 'C:/T3LabAI/bridge.py', '48885', '--agent=antigravity'],
              str(data))
        ok, msg = M.MCPService.configure_client('chatgpt', port=48885)
        toml = open(env.paths['chatgpt'], encoding='utf-8').read()
        check('Codex entry names its agent', ok and '"48885", "--agent=codex"]' in toml, toml)
        check('Codex entry with --agent is configured',
              M.MCPService.client_status('chatgpt')['configured'] is True)
        try:
            M.MCPService.config_snippet(server='nope')
            check('unknown server raises', False)
        except KeyError:
            check('unknown server raises', True)


def test_other_autodesk_apps():
    print('[other Autodesk apps: only AutoCAD and Civil 3D, one add-in, one server entry]')
    apps = M.MCPService.autodesk_apps()
    keys = [a[0] for a in apps]
    check('only apps Connect installs in one click', keys == ['autocad', 'civil3d'], keys)
    check('Civil 3D shares AutoCAD\'s server',
          dict((a[0], a[2]) for a in apps)['civil3d'] == 't3lab-autocad', apps)
    for dropped in ('navisworks', '3dsmax', 'inventor'):
        try:
            M.MCPService.config_snippet(server=dropped)
            check('dropped app raises: ' + dropped, False)
        except KeyError:
            check('dropped app raises: ' + dropped, True)
    with Env() as env:
        env.install_bridge()
        env.app_bridges['civil3d'] = env.bridge
        ok, msg = M.MCPService.configure_client('antigravity', server='civil3d')
        entry = _json(env.paths['antigravity'])['mcpServers'].get('t3lab-autocad', {})
        check('Civil 3D writes the AutoCAD entry', ok, msg)
        check('stdio entry (-I bridge, no port)',
              entry.get('args') == ['-I', env.bridge.replace('\\', '/')], entry)
        check('status configured', M.MCPService.client_status(
            'antigravity', server='civil3d').get('configured'))
        snip = M.MCPService.config_snippet(fmt='claude_code', server='civil3d')
        check('snippet per app', 't3lab-autocad' in snip and '-I' in snip, snip)
    check('cad_bridge_status kept for older callers',
          M.MCPService.cad_bridge_status()['server'] == 't3lab-autocad')


def test_dialog_attrs_assigned():
    """App switch crashed with "'MCPControlWindow' object has no attribute
    '_configure_all_btn'": every self.<attr> the dialog reads must be set
    somewhere (or come from the WPF base). Source scan: no WPF needed."""
    import re
    print('[MCP Control dialog: every self.<attr> read is assigned]')
    with open(os.path.join(LIB, 'GUI', 'MCPControlDialog.py')) as f:
        src = f.read()
    reads = set(re.findall(r'self\.(_?\w+)', src))
    assigned = set(re.findall(r'self\.(_?\w+)\s*(?:=[^=]|,)', src))
    assigned |= set(re.findall(r'def (\w+)\(self', src))
    assigned |= set(re.findall(r'^\s{4}(_?\w+)\s*=', src, re.M))
    base = {'Close', 'FindName'}
    missing = sorted(reads - assigned - base)
    check('no unassigned attributes', not missing, missing)
    check('Connect All button kept on self',
          "self._configure_all_btn   = self.FindName('configure_all_btn')" in src)


def main():
    for t in (test_needs_bridge, test_json_keeps_revit_and_others,
              test_desktop_running_guard, test_toml, test_snippets_and_revit_default,
              test_other_autodesk_apps, test_dialog_attrs_assigned):
        t()
    if FAILURES:
        print('\n{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
        return 1
    print('\nALL PASS')
    return 0


if __name__ == '__main__':
    sys.exit(main())
