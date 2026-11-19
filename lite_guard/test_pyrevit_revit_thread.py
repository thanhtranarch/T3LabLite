# -*- coding: utf-8 -*-
"""
CPython 3 test: pyrevit.revit is first imported on Revit's own thread.

Run:  python3 lite_guard/test_pyrevit_revit_thread.py      Exit 0 = all pass.

"Duplicate type name within an assembly" and "cannot import name 'revit' from
'pyrevit'" on every tool, across Revit restarts, came from one line: the usage
tracking worker thread did `from pyrevit import revit` before the tool did.
On pyRevit 6.5 / 7.0 that import creates the .NET type "<exec id>.
FuncAsEventHandler" and then calls UI.ExternalEvent.Create(), which Revit
refuses off its own thread - so the import dies with the type already emitted,
and the tool's own import of pyrevit.revit, in the same execution, cannot
succeed any more (lib/_cpython_bootstrap.py has the full account).

Covered here:
  - no thread started anywhere in lib/ can import pyrevit.revit / script;
  - tracking reads the Revit side of a row on the caller's thread and its
    worker imports nothing from pyrevit;
  - the bootstrap imports pyrevit.revit on Revit's thread, inside an API
    context, before the tracking thread starts - once, and keeps the real
    error when that fails;
  - the error dialog no longer blames pyRevit > Reload or asks for a restart.
The .NET side (pythonnet, Revit) needs Revit.
"""
import ast
import os
import re
import sys
import threading
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
if not os.path.isdir(LIB):
    LIB = os.path.join(REPO, 'lib')         # T3LabLite layout
sys.path.insert(0, LIB)

import _cpython_bootstrap as boot    # noqa: E402  (needs LIB on sys.path)

# Importing any of these runs Revit API code at module level.
REVIT_SIDE = re.compile(r'^pyrevit\.(revit|script|output|forms)(\.|$)')
THREAD_STARTERS = ('Thread', 'Timer', 'submit', 'start_new_thread',
                   'QueueUserWorkItem')


def read(path):
    with open(path, encoding='utf-8', errors='replace') as f:
        return f.read()


def pyrevit_imports(node, pattern=REVIT_SIDE):
    """(line, module) of every import under `node` that matches `pattern`."""
    found = []
    for n in ast.walk(node):
        names = []
        if isinstance(n, ast.Import):
            names = [a.name for a in n.names]
        elif isinstance(n, ast.ImportFrom) and n.module:
            names = (['pyrevit.' + a.name for a in n.names]
                     if n.module == 'pyrevit' else [n.module])
        found += [(n.lineno, name) for name in names if pattern.match(name)]
    return found


def thread_targets_importing_pyrevit(path):
    """Thread targets in one file that import the Revit side of pyrevit,
    directly or through a function of the same file they call."""
    try:
        tree = ast.parse(read(path))
    except SyntaxError:
        return []
    funcs = {}
    for n in ast.walk(tree):
        if isinstance(n, (ast.FunctionDef, ast.AsyncFunctionDef)):
            funcs.setdefault(n.name, []).append(n)
    hits = []
    for n in ast.walk(tree):
        if not isinstance(n, ast.Call):
            continue
        starter = getattr(n.func, 'attr', None) or getattr(n.func, 'id', None)
        if starter not in THREAD_STARTERS:
            continue
        target = next((kw.value for kw in n.keywords if kw.arg == 'target'), None)
        if target is None and n.args:
            target = n.args[-1] if starter == 'Timer' else n.args[0]
        name = getattr(target, 'id', None) or getattr(target, 'attr', None)
        for fn in funcs.get(name, []):
            seen = pyrevit_imports(fn)
            for call in ast.walk(fn):
                if isinstance(call, ast.Call):
                    callee = (getattr(call.func, 'attr', None)
                              or getattr(call.func, 'id', None))
                    for inner in funcs.get(callee, []):
                        seen += pyrevit_imports(inner)
            if seen:
                hits.append('%s:%d %s() -> %s' % (
                    os.path.relpath(path, REPO), n.lineno, name,
                    sorted(set(m for _, m in seen))))
    return hits


class NoThreadImportsPyrevitTests(unittest.TestCase):
    """The rule that was broken: Revit API modules load on Revit's thread."""

    def test_no_thread_in_lib_can_import_the_revit_side_of_pyrevit(self):
        hits = []
        for folder, _dirs, files in os.walk(LIB):
            for name in files:
                if name.endswith('.py'):
                    hits += thread_targets_importing_pyrevit(os.path.join(folder, name))
        self.assertEqual(hits, [], 'a background thread imports pyrevit.revit '
                                   'or pyrevit.script')

    def test_the_scan_catches_what_tracking_used_to_do(self):
        # The exact shape of the bug, so the scan above cannot go blind.
        import tempfile
        src = (
            "import threading\n"
            "class Service(object):\n"
            "    @classmethod\n"
            "    def get_revit_context(cls):\n"
            "        from pyrevit import revit, HOST_APP\n"
            "        return revit.doc.Title\n"
            "    @classmethod\n"
            "    def _dispatch_async(cls, payload):\n"
            "        def _worker():\n"
            "            payload['doc'] = cls.get_revit_context()\n"
            "        threading.Thread(target=_worker).start()\n")
        with tempfile.TemporaryDirectory() as folder:
            path = os.path.join(folder, 'old_service.py')
            with open(path, 'w', encoding='utf-8') as f:
                f.write(src)
            hits = thread_targets_importing_pyrevit(path)
        self.assertEqual(len(hits), 1)
        self.assertIn('pyrevit.revit', hits[0])

    def test_tracking_never_imports_the_revit_side_of_pyrevit(self):
        folder = os.path.join(LIB, 'tracking')
        for name in sorted(os.listdir(folder)):
            if name.endswith('.py'):
                found = pyrevit_imports(ast.parse(read(os.path.join(folder, name))))
                self.assertEqual(found, [], 'lib/tracking/%s' % name)

    def test_the_tracking_worker_imports_nothing_from_pyrevit(self):
        tree = ast.parse(read(os.path.join(LIB, 'tracking', 'service.py')))
        workers = [n for n in ast.walk(tree)
                   if isinstance(n, ast.FunctionDef) and n.name == '_worker']
        self.assertEqual(len(workers), 1)
        anything = re.compile(r'^pyrevit(\.|$)')
        self.assertEqual(pyrevit_imports(workers[0], anything), [])
        called = set(getattr(c.func, 'attr', '') for c in ast.walk(workers[0])
                     if isinstance(c, ast.Call))
        self.assertNotIn('get_revit_context', called)
        self.assertNotIn('note_revit_context', called)


class _Doc(object):
    Title = u'Tower A_tran'


class _UIDoc(object):
    Document = _Doc()


class _App(object):
    VersionNumber = u'2026'


class _UIApp(object):
    Application = _App()
    ActiveUIDocument = _UIDoc()


class _StartPage(object):           # Revit open, no document
    Application = _App()
    ActiveUIDocument = None


class _Controlled(object):          # __revit__ while Revit starts
    class ControlledApplication(object):
        VersionNumber = u'2025'


class TrackingContextTests(unittest.TestCase):
    """The Revit side of a row: read on Revit's thread, reused everywhere."""

    def setUp(self):
        from tracking.service import TelemetryService
        self.service = TelemetryService
        self.service._cached_context = None

    def test_noted_once_then_read_without_revit(self):
        self.assertEqual(self.service.get_revit_context(), ('unknown', ''))
        self.assertEqual(self.service.note_revit_context(_UIApp()),
                         (u'2026', u'Tower A_tran'))
        self.assertEqual(self.service.get_revit_context(),
                         (u'2026', u'Tower A_tran'))

    def test_a_closed_document_clears_the_title(self):
        self.service.note_revit_context(_UIApp())
        self.assertEqual(self.service.note_revit_context(_StartPage()),
                         (u'2026', u''))

    def test_startup_application_objects_are_understood(self):
        version, title = self.service.note_revit_context(_Controlled())
        self.assertEqual((version, title), (u'2025', u''))

    def test_no_application_keeps_what_was_known(self):
        self.service.note_revit_context(_UIApp())
        import tracking.service as module
        original = module._injected_revit
        module._injected_revit = lambda: None
        try:
            self.assertEqual(self.service.note_revit_context(),
                             (u'2026', u'Tower A_tran'))
        finally:
            module._injected_revit = original

    def test_the_row_is_filled_on_the_callers_thread(self):
        import tracking.service as module
        seen, sent = [], []
        done = threading.Event()

        class _Finder(object):          # any pyrevit import is a failure
            def find_spec(self, name, path=None, target=None):
                if name == 'pyrevit' or name.startswith('pyrevit.'):
                    seen.append(('import', name, threading.current_thread().name))
                return None

        def fake_urlopen(request, timeout=None):
            sent.append(request.data)
            done.set()

        original_get = self.service.__dict__['get_revit_context']
        original_urlopen = module.urlopen

        def spy(cls):
            seen.append(('context', threading.current_thread().name))
            return original_get.__func__(cls)

        self.service.note_revit_context(_UIApp())
        self.service.get_revit_context = classmethod(spy)
        module.urlopen = fake_urlopen
        finder = _Finder()
        sys.meta_path.insert(0, finder)
        try:
            self.service._dispatch_async({'tool_name': 'BatchOut'})
            self.assertTrue(done.wait(10), 'the worker never sent the row')
        finally:
            sys.meta_path.remove(finder)
            module.urlopen = original_urlopen
            self.service.get_revit_context = original_get

        here = threading.current_thread().name
        self.assertEqual(seen, [('context', here)],
                         'the context must be read once, on the calling thread, '
                         'and nothing may import pyrevit')
        import json
        row = json.loads(sent[0].decode('utf-8'))
        self.assertEqual(row['revit_version'], u'2026')
        self.assertTrue(row['doc_name'].startswith('doc_'))
        self.assertNotIn('Tower', sent[0].decode('utf-8'))   # the title stays local

    def test_a_ribbon_click_notes_the_context_before_the_row(self):
        tree = ast.parse(read(os.path.join(LIB, 'tracking', 'hooks.py')))
        click = next(n for n in tree.body if isinstance(n, ast.FunctionDef)
                     and n.name == 'track_ribbon_click')
        lines = dict((getattr(c.func, 'attr', ''), c.lineno) for c in ast.walk(click)
                     if isinstance(c, ast.Call))
        self.assertIn('note_revit_context', lines)
        self.assertLess(lines['note_revit_context'], lines['record_tool_usage'])


class _Spy(object):
    def __init__(self, raises=None):
        self.raises = raises
        self.calls = 0

    def __call__(self):
        self.calls += 1
        if self.raises is not None:
            raise self.raises


class LoadOnRevitThreadTests(unittest.TestCase):
    """load_pyrevit_revit(): when it imports, and when it must stay out."""

    def setUp(self):
        boot._PYREVIT_REVIT_TRIED = False
        boot.ENGINE_DIAGNOSIS['pyrevit_revit'] = 'not checked'
        boot.ENGINE_DIAGNOSIS['pyrevit_revit_error'] = ''

    def load(self, importer, thread=True, context=True, modules=None):
        return boot.load_pyrevit_revit(
            importer=importer, on_revit_thread=lambda: thread,
            in_api_context=lambda: context,
            modules={} if modules is None else modules)

    def test_imported_at_the_start_of_a_tool(self):
        importer = _Spy()
        self.assertEqual(self.load(importer), 'ok')
        self.assertEqual(importer.calls, 1)
        self.assertFalse(boot.pyrevit_revit_failed())

    def test_already_loaded_runs_nothing(self):
        importer = _Spy()
        self.assertEqual(self.load(importer, modules={'pyrevit.revit': object()}),
                         'loaded')
        self.assertEqual(importer.calls, 0)

    def test_never_on_a_background_thread(self):
        importer = _Spy()
        status = self.load(importer, thread=False)
        self.assertTrue(status.startswith('skipped'), status)
        self.assertEqual(importer.calls, 0)

    def test_never_outside_an_api_context(self):
        # A button of a modeless window: ExternalEvent.Create would fail there.
        importer = _Spy()
        status = self.load(importer, context=False)
        self.assertTrue(status.startswith('skipped'), status)
        self.assertEqual(importer.calls, 0)
        # ...and the next tool start, in context, still gets its one attempt.
        self.assertEqual(self.load(importer), 'ok')

    def test_a_failure_is_kept_and_not_repeated(self):
        importer = _Spy(raises=RuntimeError('Attempting to create an ExternalEvent '
                                            'outside of a standard API execution'))
        status = self.load(importer)
        self.assertTrue(status.startswith('failed: RuntimeError'), status)
        self.assertTrue(boot.pyrevit_revit_failed())
        self.assertIn('outside of a standard API execution',
                      boot.ENGINE_DIAGNOSIS['pyrevit_revit_error'])
        # A second try in the same execution could only fail with "Duplicate
        # type name" and bury the real error.
        self.assertEqual(self.load(importer), status)
        self.assertEqual(importer.calls, 1)
        self.assertIn('outside of a standard API execution',
                      boot.describe_environment())

    def test_ironpython_is_left_alone(self):
        importer, platform = _Spy(), sys.platform
        sys.platform = 'cli'
        try:
            self.assertEqual(self.load(importer), 'skipped: IronPython')
        finally:
            sys.platform = platform
        self.assertEqual(importer.calls, 0)

    def test_api_context_is_read_from_the_injected_application(self):
        import builtins

        class _InCommand(object):
            ActiveAddInId = object()

        class _Idle(object):
            ActiveAddInId = None

        had = getattr(builtins, '__revit__', None)
        try:
            builtins.__revit__ = _InCommand()
            self.assertTrue(boot._in_api_context())
            builtins.__revit__ = _Idle()
            self.assertFalse(boot._in_api_context())
        finally:
            if had is None:
                del builtins.__revit__
            else:
                builtins.__revit__ = had
        self.assertFalse(boot._in_api_context())     # no Revit at all

    def test_this_thread_counts_as_revits(self):
        self.assertTrue(boot._on_revit_thread())
        seen = []
        worker = threading.Thread(target=lambda: seen.append(boot._on_revit_thread()))
        worker.start()
        worker.join()
        self.assertEqual(seen, [False])

    def test_loaded_before_the_tracking_thread_starts(self):
        tree = ast.parse(read(os.path.join(LIB, '_cpython_bootstrap.py')))
        init = next(n for n in tree.body if isinstance(n, ast.FunctionDef)
                    and n.name == 'init_cpython_paths')
        lines = dict((getattr(c.func, 'id', ''), c.lineno) for c in ast.walk(init)
                     if isinstance(c, ast.Call))
        self.assertIn('load_pyrevit_revit', lines)
        self.assertLess(lines['load_pyrevit_revit'], lines['track_ribbon_click'])

    def test_the_wrong_explanations_are_gone(self):
        src = read(os.path.join(LIB, '_cpython_bootstrap.py'))
        for name in ('free_clr_type_names', 'heal_pyrevit_imports',
                     '_reset_clr_type_builders'):
            self.assertNotIn('def ' + name, src)
        for fact in ('EXEC_PARAMS.exec_id', 'ExternalEvent.Create',
                     'Duplicate type name within an assembly'):
            self.assertIn(fact, src)


class ErrorDialogTests(unittest.TestCase):
    """What the user is told when the follow-up error still shows."""

    BOTH = ('TypeError: Duplicate type name within an assembly.',
            "ImportError: cannot import name 'revit' from 'pyrevit' (x)")

    def setUp(self):
        from GUI import ErrorGuard
        self.guard = ErrorGuard
        boot.ENGINE_DIAGNOSIS['pyrevit_revit_error'] = ''

    def test_it_says_run_again_not_restart(self):
        for text in self.BOTH:
            step = self.guard.next_step(text)
            self.assertIn('Run the tool again', step)
            self.assertNotIn('estart Revit', step)
            self.assertNotIn('Reload', step)
            self.assertTrue(boot.failed_pyrevit_import(Exception(text)))
        self.assertTrue(self.guard.next_step('ValueError: bad sheet')
                        .startswith('Try again.'))
        self.assertFalse(boot.failed_pyrevit_import(ValueError('bad sheet')))

    def test_the_real_error_is_added_to_the_report(self):
        boot.ENGINE_DIAGNOSIS['pyrevit_revit_error'] = (
            'Traceback ...\nInvalidOperationException: Attempting to create an '
            'ExternalEvent outside of a standard API execution')
        for text in self.BOTH:
            cause = self.guard.first_failure(text)
            self.assertTrue(cause.startswith('First failure (the cause):'), cause)
            self.assertIn('outside of a standard API execution', cause)
        self.assertEqual(self.guard.first_failure('ValueError: bad sheet'), u'')

    def test_nothing_is_added_when_nothing_was_recorded(self):
        for text in self.BOTH:
            self.assertEqual(self.guard.first_failure(text), u'')


if __name__ == '__main__':
    unittest.main(verbosity=1)
