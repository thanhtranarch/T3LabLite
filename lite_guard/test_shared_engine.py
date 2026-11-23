# -*- coding: utf-8 -*-
"""T3Lab and T3Lab Lite share ONE CPython engine and the package names GUI,
Services, core: a Lite click must import Lite's modules even after a T3Lab click
loaded T3Lab's (2026-11-20: "No module named 'GUI.LLMSettingDialog'").
_cpython_bootstrap.claim_lib + the prelude of every button script. Linux, no Revit.

Run:  python3 lite_guard/test_shared_engine.py      Exit 0 = all pass.
"""
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)

import _cpython_bootstrap as B   # noqa: E402


def scripts():
    return sorted(glob.glob(os.path.join(REPO, 'T3Lab.tab', '**', 'script.py'), recursive=True))


def write(path, text=''):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        f.write(textwrap.dedent(text))


class ClaimLib(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix='t3_shared_')
        self.addCleanup(shutil.rmtree, self.root, True)
        self.ours = os.path.join(self.root, 'T3LabLite.extension', 'lib')
        self.dev = os.path.join(self.root, 'T3Lab.extension', 'lib')

    def test_foreign_modules_with_our_names_are_dropped(self):
        write(os.path.join(self.ours, 'GUI', '__init__.py'))
        write(os.path.join(self.ours, 'tool.py'))
        dev_gui = type(sys)('GUI')
        dev_gui.__file__ = os.path.join(self.dev, 'GUI', '__init__.py')
        ours_tool = type(sys)('tool')
        ours_tool.__file__ = os.path.join(self.ours, 'tool.py')
        other = type(sys)('json')
        other.__file__ = '/usr/lib/json/__init__.py'
        modules = {'GUI': dev_gui, 'GUI.Dialog': dev_gui, 'tool': ours_tool, 'json': other}
        path = ['/a', self.dev, self.ours]
        dropped = B.claim_lib(self.ours, modules, path)
        self.assertEqual(sorted(dropped), ['GUI', 'GUI.Dialog'])
        self.assertEqual(sorted(modules), ['json', 'tool'])
        self.assertEqual(path[0], os.path.normcase(os.path.abspath(self.ours)))
        self.assertEqual(len(path), 3)                  # moved, not duplicated

    def test_bootstrap_claims_on_every_click(self):
        with open(os.path.join(LIB, '_cpython_bootstrap.py'), encoding='utf-8') as f:
            src = f.read()
        body = src[src.index('def init_cpython_paths():'):]
        body = body[:body.index('\ndef ', 1)]
        self.assertLess(body.index('claim_lib()'), body.index('verify_stdlib()'))
        self.assertIn('track_ribbon_click()', body)     # Lite's tracking stays

    def test_end_to_end_after_a_t3lab_click(self):
        """T3Lab's bootstrap and GUI are loaded first; a Lite script still gets Lite's."""
        write(os.path.join(self.dev, '_cpython_bootstrap.py'), 'def init_cpython_paths(): pass\n')
        write(os.path.join(self.dev, 'GUI', '__init__.py'))
        os.makedirs(self.ours, exist_ok=True)
        shutil.copy(os.path.join(LIB, '_cpython_bootstrap.py'), self.ours)
        write(os.path.join(self.ours, 'GUI', '__init__.py'))
        write(os.path.join(self.ours, 'GUI', 'LLMSettingDialog.py'), 'OK = "lite"\n')
        with open(scripts()[0], encoding='utf-8') as f:
            real = f.read()
        prelude = re.search(r'# T3Lab shares.*?sys\.path\.insert\(0, _t3_lib\)\n', real, re.S)
        self.assertTrue(prelude, scripts()[0])
        script = os.path.join(self.root, 'T3LabLite.extension', 'T3Lab.tab', 'P.panel',
                              'B.pushbutton', 'script.py')
        write(script, 'import os\nimport sys\n' + prelude.group(0) + textwrap.dedent('''
            import _cpython_bootstrap
            _cpython_bootstrap.init_cpython_paths()
            from GUI.LLMSettingDialog import OK
            print(OK, _cpython_bootstrap.__file__)
            '''))
        driver = textwrap.dedent('''
            import runpy, sys
            sys.path.insert(0, {dev!r})
            import _cpython_bootstrap, GUI           # the T3Lab click
            runpy.run_path({script!r}, run_name="__main__")
            ''').format(dev=self.dev, script=script)
        out = subprocess.run([sys.executable, '-I', '-c', driver], capture_output=True,
                             text=True, timeout=60)
        self.assertEqual(out.returncode, 0, out.stderr)
        self.assertTrue(out.stdout.startswith('lite '), out.stdout)
        self.assertIn('T3LabLite.extension', out.stdout.split()[1])


class EveryScript(unittest.TestCase):
    def test_prelude_runs_before_the_bootstrap(self):
        found = scripts()
        self.assertGreaterEqual(len(found), 52)
        for path in found:
            with open(path, encoding='utf-8') as f:
                text = f.read()
            guard, boot = text.find('_t3_bs = sys.modules.get'), text.find('import _cpython_bootstrap')
            self.assertTrue(0 <= guard < boot, os.path.relpath(path, REPO))

    def test_prelude_finds_the_lib_of_its_own_extension(self):
        """The lib next to the tab folder, whatever the extension folder is called."""
        for ext in ('T3LabLite.extension', 't3lablite'):
            fake = os.sep.join(['', 'x', ext, 'T3Lab.tab', 'Support.panel', 'A.stack',
                                'B.pushbutton', 'script.py'])
            lib = os.path.join(fake.split('.tab' + os.sep)[0].rsplit(os.sep, 1)[0], 'lib')
            self.assertEqual(lib, os.sep.join(['', 'x', ext, 'lib']))
        with open(scripts()[0], encoding='utf-8') as f:
            self.assertIn("split('.tab' + os.sep)[0].rsplit(os.sep, 1)[0], 'lib')", f.read())


class SharedClrTypes(unittest.TestCase):
    """lib/ modules T3Lab also ships: a fixed __namespace__ would raise
    "Duplicate type name within an assembly" after the other extension's click."""
    SHARED = ('Snippets/finish_floors.py', 'Snippets/wall_tools.py', 'Snippets/_links.py',
              'Snippets/_auto_dimension.py', 'Services/point_cloud_build.py')

    def test_namespace_is_unique_per_load(self):
        from Snippets._compat import clr_namespace
        a, b = clr_namespace('X'), clr_namespace('X')
        self.assertTrue(a.startswith('T3Lab.X_') and a != b, (a, b))
        for rel in self.SHARED:
            with open(os.path.join(LIB, *rel.split('/')), encoding='utf-8') as f:
                lines = [l.strip() for l in f if '__namespace__' in l and '=' in l]
            self.assertTrue(lines, rel)
            for line in lines:
                self.assertIn('clr_namespace(', line, rel)

    def test_no_fixed_namespace_left(self):
        # core/server.py runs only under IronPython, where __namespace__ is ignored.
        allowed = {os.path.join('core', 'server.py')}
        pattern = re.compile(r"""__namespace__ *= *["']T3Lab\.[A-Za-z0-9_.]+["'] *($|#)""")
        fixed = []
        for root, _dirs, files in os.walk(LIB):
            for name in files:
                if not name.endswith('.py'):
                    continue
                path = os.path.join(root, name)
                rel = os.path.relpath(path, LIB)
                with open(path, encoding='utf-8') as f:
                    for n, line in enumerate(f, 1):
                        if pattern.search(line) and rel not in allowed:
                            fixed.append('{}:{}'.format(rel, n))
        self.assertEqual(fixed, [])


if __name__ == '__main__':
    unittest.main(verbosity=1)
