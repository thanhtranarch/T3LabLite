# -*- coding: utf-8 -*-
"""MCP Control's Connect installs the T3Lab add-in for AutoCAD / Civil 3D / 3ds Max from its
release zip (lib/Services/addin_installer.py, MCPService.install_addin). Plain CPython 3."""
import io
import json
import os
import sys
import tempfile
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
sys.path.insert(0, LIB)
_TMP = tempfile.mkdtemp(prefix='t3lab_addin_test_')
os.environ['APPDATA'] = _TMP

from Services import addin_installer as AI   # noqa: E402


def make_max_zip(path):
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('manifest.json', json.dumps({'name': 'T3LabMax', 'version': '1.1.0'}))
        z.writestr('T3LabMax.bundle/PackageContents.xml', '<ApplicationPackage/>')
        z.writestr('T3LabMax.bundle/Contents/scripts/t3lab_max_mcp.py', '# server')
        z.writestr('T3LabMax.bundle/Contents/scripts/t3lab_max_mcp_start.ms', '-- start')
        z.writestr('T3LabMax/bridge.py', '# bridge')
    return path


def make_zip(path, name='T3LabCAD', extra=None, manifest=True, bridge=True):
    with zipfile.ZipFile(path, 'w') as z:
        if manifest:
            z.writestr('manifest.json', json.dumps({'name': name, 'version': '1.1.0'}))
        z.writestr('T3LabCAD.bundle/PackageContents.xml', '<ApplicationPackage/>')
        for s in ('R24', 'R25', 'R26'):
            z.writestr('T3LabCAD.bundle/Contents/{}/T3LabCAD.Loader.dll'.format(s), 'dll ' + s)
        if bridge:
            z.writestr('T3LabCAD/bridge.py', '# bridge')
        for arc, data in (extra or {}).items():
            z.writestr(arc, data)
    return path


class InstallTests(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix='t3lab_addin_src_')
        os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_addin_appdata_')

    def test_installs_bundle_and_bridge(self):
        src = make_zip(os.path.join(self.dir, 'T3LabCAD.zip'))
        res = AI.install('T3LabCAD', source=src)
        bundle = os.path.join(os.environ['APPDATA'], 'Autodesk', 'ApplicationPlugins', 'T3LabCAD.bundle')
        self.assertTrue(os.path.isfile(os.path.join(bundle, 'PackageContents.xml')))
        self.assertTrue(os.path.isfile(os.path.join(bundle, 'Contents', 'R26', 'T3LabCAD.Loader.dll')))
        self.assertTrue(os.path.isfile(os.path.join(os.environ['APPDATA'], 'T3LabCAD', 'bridge.py')))
        self.assertEqual(res['version'], '1.1.0')
        self.assertEqual(res['locked'], [])
        self.assertIn('1.1.0', AI.summary(res))
        again = AI.install('T3LabCAD', source=src)
        self.assertEqual(again['written'], [])          # unchanged files are left alone

    def test_installs_3dsmax(self):
        res = AI.install('T3LabMax', source=make_max_zip(os.path.join(self.dir, 'T3LabMax.zip')))
        bundle = os.path.join(os.environ['APPDATA'], 'Autodesk', 'ApplicationPlugins', 'T3LabMax.bundle')
        self.assertTrue(os.path.isfile(os.path.join(bundle, 'Contents', 'scripts', 't3lab_max_mcp_start.ms')))
        self.assertTrue(os.path.isfile(os.path.join(os.environ['APPDATA'], 'T3LabMax', 'bridge.py')))
        self.assertIn('Restart 3ds Max', AI.summary(res))
        with self.assertRaises(AI.InstallError):            # a CAD zip is not the Max package
            AI.install('T3LabMax', source=make_zip(os.path.join(self.dir, 'c.zip')))

    def test_old_layout_removed(self):
        old = os.path.join(os.environ['APPDATA'], 'Autodesk', 'ApplicationPlugins', 'T3LabCAD.bundle',
                           'Contents', 'T3LabCAD.Loader.dll')
        os.makedirs(os.path.dirname(old))
        open(old, 'w').close()
        AI.install('T3LabCAD', source=make_zip(os.path.join(self.dir, 'a.zip')))
        self.assertFalse(os.path.exists(old))

    def test_rejects_bad_zips(self):
        cases = {
            'unsafe': make_zip(os.path.join(self.dir, 'u.zip'), extra={'T3LabCAD/../../evil.py': 'x'}),
            'unexpected': make_zip(os.path.join(self.dir, 'x.zip'), extra={'other/file.txt': 'x'}),
            'wrong name': make_zip(os.path.join(self.dir, 'n.zip'), name='Other'),
            'no manifest': make_zip(os.path.join(self.dir, 'm.zip'), manifest=False),
            'no bridge': make_zip(os.path.join(self.dir, 'b.zip'), bridge=False),
        }
        for label, path in cases.items():
            with self.assertRaises(AI.InstallError, msg=label):
                AI.install('T3LabCAD', source=path)
        notzip = os.path.join(self.dir, 'n.txt')
        open(notzip, 'w').write('hello')
        with self.assertRaises(AI.InstallError):
            AI.install('T3LabCAD', source=notzip)
        with self.assertRaises(AI.InstallError):
            AI.install('T3LabCAD', source='http://example.com/T3LabCAD.zip')    # not https
        with self.assertRaises(AI.InstallError):
            AI.install('T3LabNavisworks')

    def test_installed_years(self):
        root = tempfile.mkdtemp(prefix='t3lab_pf_')
        for name in ('AutoCAD 2021', 'AutoCAD 2024', 'AutoCAD 2027', 'Navisworks Manage 2025',
                     '3ds Max 2026'):
            os.makedirs(os.path.join(root, name))
        os.makedirs(os.path.join(root, 'AutoCAD 2024', 'C3D'))
        self.assertEqual(AI.installed_years('autocad', [root]), [2024, 2027])
        self.assertEqual(AI.installed_years('civil3d', [root]), [2024])
        self.assertEqual(AI.installed_years('navisworks', [root]), [2025])
        self.assertEqual(AI.installed_years('3dsmax', [root]), [2026])
        self.assertEqual(AI.installed_years('inventor', [root]), [])       # dropped


class ServiceTests(unittest.TestCase):
    def setUp(self):
        os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_addin_svc_')
        from Services import mcp_service as M
        self.M = M

    def test_connect_installs_then_writes_the_entry(self):
        M = self.M
        src = make_zip(os.path.join(tempfile.mkdtemp(), 'T3LabCAD.zip'))
        paths = M._get_paths_module()
        real_load = paths.load_settings
        paths.load_settings = lambda *a, **k: dict(real_load(*a, **k), T3LabCAD_addin_url=src)
        self.addCleanup(setattr, paths, 'load_settings', real_load)
        cfg = os.path.join(os.environ['APPDATA'], 'agy.json')
        real_resolve = M._resolve_client_path
        M._resolve_client_path = lambda spec: cfg
        self.addCleanup(setattr, M, '_resolve_client_path', real_resolve)
        ok, msg = M.MCPService.configure_client('antigravity', server='civil3d')
        self.assertTrue(ok, msg)
        with io.open(cfg, encoding='utf-8') as f:
            entry = json.load(f)['mcpServers']['t3lab-autocad']
        self.assertTrue(entry['args'][-1].endswith('T3LabCAD/bridge.py'), entry)
        st = M.MCPService.app_bridge_status('autocad')
        self.assertTrue(st['installed'] and st['installable'])

    def test_apps_without_an_addin_say_so(self):
        ok, msg = self.M.MCPService.install_addin('navisworks')
        self.assertFalse(ok)
        self.assertIn('no T3Lab add-in for Navisworks yet', msg)
        self.assertFalse(self.M.MCPService.app_bridge_status('navisworks')['installable'])
        self.assertTrue(self.M.MCPService.app_bridge_status('3dsmax')['installable'])

    def test_inventor_is_gone(self):
        keys = [a[0] for a in self.M.MCPService.autodesk_apps()]
        self.assertNotIn('inventor', keys)
        self.assertNotIn('inventor', self.M.SERVERS)


if __name__ == '__main__':
    unittest.main()
