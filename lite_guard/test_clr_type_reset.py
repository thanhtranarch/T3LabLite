# -*- coding: utf-8 -*-
"""
CPython 3 test: "Duplicate type name within an assembly" self-heal
(lib/_cpython_bootstrap.py: free_clr_type_names).

Run:  python3 lite_guard/test_clr_type_reset.py      Exit 0 = all pass.

When the CPython runtime restarts inside one Revit session (pyRevit > Reload,
an extension switched off and on, an update applied while Revit was open),
sys.modules starts empty but pythonnet's dynamic assembly still holds every
.NET type emitted for a Python subclass. The first fixed-name class re-created
then dies - pyRevit's own pyrevit.revit.events.FuncAsEventHandler is the first
one in the way, so every tool fails on `from pyrevit import revit/script/forms`
until Revit restarts.

The bootstrap empties pythonnet's builder dictionaries before the first pyrevit
import, which hands it a fresh dynamic assembly. The .NET side needs Revit;
what is tested here is the decision logic (when it fires, when it must NOT) and
that the call really sits ahead of every pyrevit import.
"""
import ast
import os
import sys
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
if not os.path.isdir(LIB):
    LIB = os.path.join(REPO, 'lib')         # T3LabLite layout
sys.path.insert(0, LIB)
TAB = os.path.join(os.path.dirname(LIB), 'T3Lab.tab')

import _cpython_bootstrap as boot    # noqa: E402  (needs LIB on sys.path)

def read(path):
    with open(path, encoding='utf-8', errors='replace') as f:
        return f.read()


BOOT_SRC = read(os.path.join(LIB, '_cpython_bootstrap.py'))


class _Spy(object):
    """A stand-in for one of the two .NET steps."""

    def __init__(self, result=None, raises=None):
        self.result = result
        self.raises = raises
        self.calls = 0

    def __call__(self):
        self.calls += 1
        if self.raises is not None:
            raise self.raises
        return self.result


def heal(modules, emitted, reset):
    return boot.free_clr_type_names(modules=modules, emitted=emitted,
                                    reset=reset, once=False)


class DecisionTests(unittest.TestCase):
    """What the bootstrap does with the pythonnet builders, and when."""

    def test_resets_when_the_runtime_restarted(self):
        # No pyrevit module loaded, yet .NET types are already emitted: the
        # runtime restarted in this Revit session.
        emitted, reset = _Spy(7), _Spy('Reset()')
        status = heal({}, emitted, reset)
        self.assertEqual(reset.calls, 1)
        self.assertTrue(status.startswith('freed 7 stale .NET type'), status)
        self.assertIn('Reset()', status)

    def test_never_resets_while_pyrevit_is_loaded(self):
        # Same runtime as the types: live windows and event handlers still
        # depend on them, and importing those modules again is a no-op.
        for name in ('pyrevit.revit.events', 'pyrevit.revit', 'pyrevit.forms'):
            emitted, reset = _Spy(7), _Spy('Reset()')
            status = heal({name: object()}, emitted, reset)
            self.assertEqual(reset.calls, 0, name)
            self.assertEqual(emitted.calls, 0, name)
            self.assertTrue(status.startswith('clean:'), status)
            self.assertIn(name, status)

    def test_first_run_of_a_fresh_revit_does_nothing(self):
        emitted, reset = _Spy(0), _Spy('Reset()')
        status = heal({}, emitted, reset)
        self.assertEqual(reset.calls, 0)
        self.assertEqual(status, 'clean: no .NET types emitted yet')

    def test_unknown_pythonnet_build_is_left_alone(self):
        # _emitted_clr_types() returns None when the fields this patch was
        # written against are not there: touch nothing.
        emitted, reset = _Spy(None), _Spy('Reset()')
        status = heal({}, emitted, reset)
        self.assertEqual(reset.calls, 0)
        self.assertTrue(status.startswith('skipped:'), status)

    def test_a_broken_probe_never_breaks_the_tool(self):
        emitted, reset = _Spy(raises=AttributeError('no such field')), _Spy()
        status = heal({}, emitted, reset)
        self.assertEqual(reset.calls, 0)
        self.assertTrue(status.startswith('skipped: AttributeError'), status)

    def test_a_failing_reset_is_reported_not_raised(self):
        emitted, reset = _Spy(3), _Spy(raises=TypeError('bad invoke'))
        status = heal({}, emitted, reset)
        self.assertTrue(status.startswith('failed: TypeError'), status)

    def test_a_reset_that_found_nothing_is_a_failure(self):
        emitted, reset = _Spy(3), _Spy('')
        status = heal({}, emitted, reset)
        self.assertTrue(status.startswith('failed:'), status)

    def test_runs_once_per_runtime(self):
        # The flag lives in the module, so a restarted runtime (fresh module)
        # starts over - exactly when the work is needed again.
        boot._CLR_TYPES_CHECKED = False
        try:
            emitted, reset = _Spy(5), _Spy('Reset()')
            first = boot.free_clr_type_names(modules={}, emitted=emitted,
                                             reset=reset)
            second = boot.free_clr_type_names(modules={}, emitted=emitted,
                                              reset=reset)
            self.assertEqual(reset.calls, 1)
            self.assertEqual(first, second)
        finally:
            boot._CLR_TYPES_CHECKED = True

    def test_the_status_is_reported(self):
        emitted, reset = _Spy(2), _Spy('Reset()')
        status = heal({}, emitted, reset)
        self.assertEqual(boot.CLR_TYPE_RESET_STATUS, status)
        self.assertEqual(boot.ENGINE_DIAGNOSIS['clr_types'], status)
        report = boot.describe_environment()
        self.assertIn('CPython restarted', report)
        self.assertIn('Reload', report)


class RecogniseTests(unittest.TestCase):
    """Telling a restarted runtime from an ordinary failure."""

    def test_both_faces_of_a_stale_type_are_recognised(self):
        self.assertTrue(boot.stale_clr_types(
            TypeError('Duplicate type name within an assembly.')))
        self.assertTrue(boot.stale_clr_types(ImportError(
            "cannot import name 'revit' from 'pyrevit' (pyrevit/__init__.py)")))

    def test_an_ordinary_failure_is_not_one(self):
        self.assertFalse(boot.stale_clr_types(ImportError('No module named x')))
        self.assertFalse(boot.stale_clr_types(ValueError('bad sheet')))
        self.assertFalse(boot.stale_clr_types(None))


class WiringTests(unittest.TestCase):
    """The heal must happen before anything imports pyrevit."""

    def test_init_calls_it_before_its_own_pyrevit_import(self):
        tree = ast.parse(BOOT_SRC)
        init = next(n for n in tree.body
                    if isinstance(n, ast.FunctionDef)
                    and n.name == 'init_cpython_paths')
        call = min(n.lineno for n in ast.walk(init)
                   if isinstance(n, ast.Call)
                   and getattr(n.func, 'id', '') == 'free_clr_type_names')
        imports = [n.lineno for n in ast.walk(init)
                   if isinstance(n, (ast.Import, ast.ImportFrom))
                   and 'pyrevit' in (getattr(n, 'module', '') or
                                     ''.join(a.name for a in n.names))]
        self.assertEqual(imports, [], 'pyrevit is imported by '
                                      '_patch_pyrevit_forms(), nowhere else')
        patch = [n.lineno for n in ast.walk(init)
                 if isinstance(n, ast.Call)
                 and getattr(n.func, 'id', '') == '_patch_pyrevit_forms']
        self.assertTrue(patch, 'init_cpython_paths no longer patches pyrevit')
        self.assertLess(call, min(patch))

    def test_the_bootstrap_runs_before_pyrevit_in_every_cpython_tool(self):
        late = []
        for folder, _dirs, files in os.walk(TAB):
            if 'script.py' not in files:
                continue
            path = os.path.join(folder, 'script.py')
            src = read(path)
            if not src.startswith('#! python3'):
                continue
            tree = ast.parse(src)
            boots = [n.lineno for n in ast.walk(tree)
                     if isinstance(n, ast.Call)
                     and getattr(n.func, 'attr', '') == 'init_cpython_paths']
            pyrevits = [n.lineno for n in ast.walk(tree)
                        if isinstance(n, (ast.Import, ast.ImportFrom))
                        and (getattr(n, 'module', '') or '').startswith('pyrevit')]
            if not boots or (pyrevits and min(pyrevits) < min(boots)):
                late.append(os.path.relpath(path, REPO))
        self.assertEqual(late, [], 'pyrevit imported before the bootstrap')

    def test_the_first_pyrevit_import_retries_once_after_a_reset(self):
        # The pyrevit import inside init_cpython_paths() is the click's first;
        # a restarted runtime fails there, and the tool's own
        # `from pyrevit import script` would be next.
        tree = ast.parse(BOOT_SRC)
        init = next(n for n in tree.body
                    if isinstance(n, ast.FunctionDef)
                    and n.name == 'init_cpython_paths')
        body = ast.dump(init)
        self.assertIn("attr='_patch_pyrevit_forms'", body.replace('id=', 'attr='))
        self.assertIn("'stale_clr_types'", body)
        retry = [n for n in ast.walk(init)
                 if isinstance(n, ast.Call)
                 and getattr(n.func, 'id', '') == 'free_clr_type_names'
                 and any(kw.arg == 'once' for kw in n.keywords)]
        self.assertEqual(len(retry), 1, 'the retry no longer forces a reset')

    def test_the_forms_patch_reports_its_failure(self):
        # It must raise, not swallow: init_cpython_paths() reads the exception.
        with self.assertRaises(Exception):
            boot._patch_pyrevit_forms()

    def test_the_why_is_written_down(self):
        for text in ('Duplicate type name within an assembly',
                     'ClassDerivedObject', 'moduleBuilders',
                     'FuncAsEventHandler'):
            self.assertIn(text, BOOT_SRC)


if __name__ == '__main__':
    unittest.main(verbosity=1)
