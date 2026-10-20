#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Split buttons remember their last-used button across Revit sessions.

    python3 lite_guard/test_ribbon_memory.py

Same tests as t3lab-revit-api dev/test_ribbon_memory.py, with Lite's layout
(the extension is the repository root, the tab is T3Lab.tab).

Revit is not available here, so the ribbon is faked with plain objects that
have the same members lib/ribbon_memory.py touches (GetRibbonPanels,
GetItems, Name, CurrentButton, Idling, ApplicationClosing). What the real
SplitButton does with CurrentButton must be verified in Revit.

Plain CPython 3, standard library only. CI runs it next to guard.py
(.github/workflows/lite-guard.yml).
"""
import importlib.util
import io
import json
import os
import re
import shutil
import sys
import tempfile
import tokenize
import types
import unittest

sys.dont_write_bytecode = True

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXT = ROOT
LIB = os.path.join(EXT, 'lib')
STARTUP = os.path.join(EXT, 'startup.py')
MEMORY = os.path.join(LIB, 'ribbon_memory.py')

_spec = importlib.util.spec_from_file_location(
    't3_extension_paths', os.path.join(LIB, 'core', 'extension_paths.py'))
_paths = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_paths)
if LIB not in sys.path:
    sys.path.insert(0, LIB)
import ribbon_memory as rm  # noqa: E402

STACK = _paths.bundle_path('CloudLinks.stack')
SPLIT = _paths.bundle_path('PDFMarkup.splitbutton')
# The Revit tab the split sits on = its tab folder without '.tab'.
TAB = [os.path.basename(t)[:-len('.tab')] for t in _paths.tab_dirs(EXT)
       if os.path.abspath(SPLIT).startswith(os.path.abspath(t) + os.sep)][0]


def _read(path):
    with io.open(path, encoding='utf-8') as fh:
        return fh.read()


def _function(src, name):
    m = re.search(r'^def %s\(.*?(?=^\S|\Z)' % re.escape(name), src, re.S | re.M)
    return m.group() if m else ''


def _code(text):
    return '\n'.join(line.split('#', 1)[0] for line in text.splitlines())


def _layout(folder):
    names, inside = [], False
    for line in _read(os.path.join(folder, 'bundle.yaml')).splitlines():
        if re.match(r'layout\s*:', line):
            inside = True
            continue
        item = re.match(r'\s+-\s*(.+?)\s*$', line)
        if inside and item:
            names.append(item.group(1).strip('"\''))
        elif inside and line.strip():
            break
    return names


# ── fake Revit ribbon ────────────────────────────────────────────────────


class Button(object):
    def __init__(self, name):
        self.Name = name


class Split(object):
    """Revit SplitButton: first child is current, as pyRevit builds it."""

    def __init__(self, name, children):
        self.Name = name
        self._children = [Button(c) for c in children]
        self.CurrentButton = self._children[0]

    def GetItems(self):
        return list(self._children)

    def click(self, child):
        # IsSynchronizedWithCurrentItem = True: the clicked button goes on top.
        self.CurrentButton = [b for b in self._children if b.Name == child][0]


class Pulldown(object):
    def __init__(self, name):
        self.Name = name


class Panel(object):
    def __init__(self, *items):
        self._items = list(items)

    def GetItems(self):
        return list(self._items)


class Event(object):
    def __init__(self):
        self.handlers = []

    def __iadd__(self, fn):
        self.handlers.append(fn)
        return self

    def __isub__(self, fn):
        self.handlers.remove(fn)
        return self

    def fire(self):
        for fn in list(self.handlers):
            fn(None, None)


class UIApp(object):
    def __init__(self, tabs=None):
        self.tabs = tabs or {}
        self.Idling = Event()
        self.ApplicationClosing = Event()

    def GetRibbonPanels(self, tab):
        if tab not in self.tabs:
            raise ValueError('The tab %s does not exist.' % tab)
        return self.tabs[tab]


def _ribbon(current='Bluebeam Status'):
    split = Split('PDFMarkup', ['Bluebeam Status', 'T3Lab PDF Markup'])
    split.click(current)
    stack = Panel(Button('Autodesk Forma'), Button('Autodesk Health'), split,
                  Pulldown('Settings'))
    return split, UIApp({TAB: [stack]})


class _FakeAppDomain(object):
    def __init__(self):
        self.data = {}
        self.CurrentDomain = self

    def GetData(self, key):
        return self.data.get(key)

    def SetData(self, key, value):
        self.data[key] = value


# ── tests ────────────────────────────────────────────────────────────────


class IronPythonSafety(unittest.TestCase):
    """startup.py runs ribbon_memory.py on IronPython 2.7 / 3.4."""

    def test_no_f_strings(self):
        with io.open(MEMORY, 'rb') as fh:
            tokens = list(tokenize.tokenize(fh.readline))
        fstrings = [t for t in tokens
                    if t.type == tokenize.STRING and re.match(r'(?i)[rb]*f', t.string)]
        fstrings += [t for t in tokens if t.type == getattr(tokenize, 'FSTRING_START', -1)]
        self.assertEqual(fstrings, [])

    def test_no_python3_only_syntax(self):
        src = _read(MEMORY)
        for pattern in (r'\bnonlocal\b', r':=', r'^\s*async\s', r'\bawait\b',
                        r'def \w+\([^)]*\*,'):
            self.assertIsNone(re.search(pattern, src, re.M), pattern)

    def test_no_dotnet_import_at_module_level(self):
        # Importing from CPython (tests, tools) must not need Revit or .NET.
        top = [l for l in _read(MEMORY).splitlines() if re.match(r'(from|import) ', l)]
        self.assertFalse([l for l in top if 'System' in l or 'Autodesk' in l], top)


class RibbonLayout(unittest.TestCase):

    def test_split_button_follows_the_last_click(self):
        # .splitbutton = IsSynchronizedWithCurrentItem True (pyRevit 6.5.5
        # SplitButtonBuilder); .splitpushbutton would pin the first button.
        self.assertTrue(os.path.isdir(SPLIT))
        self.assertIn('PDFMarkup', _layout(STACK))
        self.assertNotIn('Bluebeam Status', _layout(STACK))

    def test_bluebeam_first_then_t3lab_pdf(self):
        self.assertEqual(_layout(SPLIT), ['Bluebeam Status', 'T3Lab PDF Markup'])
        for child in _layout(SPLIT):
            self.assertTrue(os.path.isdir(os.path.join(SPLIT, child + '.urlbutton')), child)

    def test_t3lab_pdf_opens_the_app(self):
        text = _read(os.path.join(SPLIT, 'T3Lab PDF Markup.urlbutton', 'bundle.yaml'))
        self.assertRegex(text, r'hyperlink:\s*"https://pdf\.t3lab\.space/?"')
        self.assertIn('context: "zero-doc"', text)

    def test_bluebeam_still_opens_its_status_page(self):
        text = _read(os.path.join(SPLIT, 'Bluebeam Status.urlbutton', 'bundle.yaml'))
        self.assertIn('https://status.bluebeam.com/', text)

    def test_discovery_finds_the_split_and_the_tab(self):
        self.assertIn('PDFMarkup', rm.split_names())
        self.assertIn(TAB, rm.tab_names())


class StateFile(unittest.TestCase):

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.path = os.path.join(self.tmp, 'sub', 'ribbon_state.json')

    def tearDown(self):
        shutil.rmtree(self.tmp)

    def test_missing_or_broken_file_is_empty(self):
        self.assertEqual(rm.load_state(self.path), {})
        os.makedirs(os.path.dirname(self.path))
        with io.open(self.path, 'w', encoding='utf-8') as fh:
            fh.write(u'{not json')
        self.assertEqual(rm.load_state(self.path), {})

    def test_round_trip_and_merge(self):
        self.assertTrue(rm.save_state({'PDFMarkup': 'T3Lab PDF Markup'}, self.path))
        self.assertTrue(rm.save_state({'Other': 'B'}, self.path))
        self.assertEqual(rm.load_state(self.path),
                         {'PDFMarkup': 'T3Lab PDF Markup', 'Other': 'B'})
        with io.open(self.path, encoding='utf-8') as fh:
            self.assertIn('split_current', json.load(fh))

    def test_nothing_to_save_writes_nothing(self):
        self.assertFalse(rm.save_state({}, self.path))
        self.assertFalse(os.path.exists(self.path))

    def test_default_path_is_per_user_not_in_the_extension(self):
        self.assertTrue(rm.state_path().endswith(os.path.join('T3LabAI', 'ribbon_state.json')))
        self.assertNotIn(os.path.abspath(EXT), os.path.abspath(rm.state_path()))


class RememberAndRestore(unittest.TestCase):

    names = ['PDFMarkup']
    tabs = [TAB]

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.path = os.path.join(self.tmp, 'ribbon_state.json')
        rm._LIVE.clear()

    def tearDown(self):
        shutil.rmtree(self.tmp)
        rm._LIVE.clear()

    def test_finds_a_split_inside_a_stack_and_skips_missing_tabs(self):
        split, app = _ribbon()
        found = rm.find_splits(app, self.names, ['No Such Tab', TAB])
        self.assertIs(found.get('PDFMarkup'), split)

    def test_a_pulldown_with_the_same_name_is_not_a_split(self):
        app = UIApp({TAB: [Panel(Pulldown('PDFMarkup'))]})
        self.assertEqual(rm.find_splits(app, self.names, self.tabs), {})

    def test_last_click_survives_a_restart(self):
        # Session 1: user clicks T3Lab PDF Markup, then closes Revit.
        split, app = _ribbon()
        split.click('T3Lab PDF Markup')
        self.assertEqual(rm.remember(app, self.path, self.names, self.tabs), 1)
        # Session 2: pyRevit rebuilds the ribbon with Bluebeam on top.
        split2, app2 = _ribbon()
        self.assertEqual(split2.CurrentButton.Name, 'Bluebeam Status')
        self.assertEqual(rm.restore(app2, self.path, self.names, self.tabs), (1, 1))
        self.assertEqual(split2.CurrentButton.Name, 'T3Lab PDF Markup')

    def test_and_back_to_bluebeam(self):
        split, app = _ribbon('T3Lab PDF Markup')
        split.click('Bluebeam Status')
        rm.remember(app, self.path, self.names, self.tabs)
        split2, app2 = _ribbon('T3Lab PDF Markup')
        rm.restore(app2, self.path, self.names, self.tabs)
        self.assertEqual(split2.CurrentButton.Name, 'Bluebeam Status')

    def test_a_removed_button_falls_back_to_the_first(self):
        rm.save_state({'PDFMarkup': 'Gone Tool'}, self.path)
        split, app = _ribbon()
        self.assertEqual(rm.restore(app, self.path, self.names, self.tabs), (1, 0))
        self.assertEqual(split.CurrentButton.Name, 'Bluebeam Status')

    def test_no_ribbon_yet_saves_nothing(self):
        # First start: startup.py runs before pyRevit builds the ribbon.
        rm.save_state({'PDFMarkup': 'T3Lab PDF Markup'}, self.path)
        app = UIApp()
        self.assertEqual(rm.remember(app, self.path, self.names, self.tabs), 0)
        self.assertEqual(rm.restore(app, self.path, self.names, self.tabs), (0, 0))
        self.assertEqual(rm.load_state(self.path), {'PDFMarkup': 'T3Lab PDF Markup'})

    def test_a_renamed_tab_is_still_saved_on_close(self):
        # Ribbon Names shortens tab titles mid-session; the split found at
        # restore time is still read when Revit closes.
        split, app = _ribbon()
        rm.restore(app, self.path, self.names, self.tabs)
        app.tabs = {'Short': app.tabs[TAB]}
        split.click('T3Lab PDF Markup')
        self.assertEqual(rm.remember(app, self.path, self.names, self.tabs), 1)
        self.assertEqual(rm.load_state(self.path), {'PDFMarkup': 'T3Lab PDF Markup'})

    def test_nothing_raises_without_revit(self):
        self.assertEqual(rm.remember(None, self.path), 0)
        self.assertEqual(rm.restore(None, self.path), (0, 0))
        state, detail = rm.install(None)
        self.assertEqual(state, rm.FAILED)
        self.assertTrue(detail)


class InstallInRevit(unittest.TestCase):
    """install() with a fake System.AppDomain, as startup.py calls it."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.path = os.path.join(self.tmp, 'ribbon_state.json')
        self._state_path = rm.state_path
        rm.state_path = lambda: self.path
        self.domain = _FakeAppDomain()
        rm._LIVE.clear()
        self._system = sys.modules.get('System')
        sys.modules['System'] = types.SimpleNamespace(AppDomain=self.domain)

    def tearDown(self):
        rm.state_path = self._state_path
        rm._LIVE.clear()
        if self._system is None:
            sys.modules.pop('System', None)
        else:
            sys.modules['System'] = self._system
        shutil.rmtree(self.tmp)

    def test_first_start_restores_on_idling_and_saves_on_close(self):
        rm.save_state({'PDFMarkup': 'T3Lab PDF Markup'}, self.path)
        app = UIApp()                                   # ribbon not built yet
        state, detail = rm.install(app)
        self.assertEqual(state, rm.INSTALLED, detail)
        self.assertEqual(len(app.Idling.handlers), 1)
        app.Idling.fire()                               # still no ribbon: keep waiting
        self.assertEqual(len(app.Idling.handlers), 1)
        split, built = _ribbon()                        # pyRevit builds the ribbon
        app.tabs = built.tabs
        app.Idling.fire()
        self.assertEqual(split.CurrentButton.Name, 'T3Lab PDF Markup')
        self.assertEqual(app.Idling.handlers, [])       # one-shot
        split.click('Bluebeam Status')
        app.ApplicationClosing.fire()
        self.assertEqual(rm.load_state(self.path), {'PDFMarkup': 'Bluebeam Status'})

    def test_idling_gives_up_when_the_split_never_appears(self):
        app = UIApp()
        rm.install(app)
        for _ in range(rm.RESTORE_MAX_TRIES):
            app.Idling.fire()
        self.assertEqual(app.Idling.handlers, [])

    def test_reload_keeps_this_sessions_choice(self):
        rm.save_state({'PDFMarkup': 'Bluebeam Status'}, self.path)   # last close
        split, app = _ribbon()
        split.click('T3Lab PDF Markup')                 # clicked in this session
        rm.install(app)                                 # pyRevit Reload
        split.click('Bluebeam Status')                  # reload rebuilt it
        app.Idling.fire()
        self.assertEqual(split.CurrentButton.Name, 'T3Lab PDF Markup')

    def test_closing_is_hooked_once_per_revit_session(self):
        app = UIApp()
        rm.install(app)
        rm.install(app)                                 # pyRevit Reload
        self.assertEqual(len(app.ApplicationClosing.handlers), 1)


class StartupInstallsIt(unittest.TestCase):

    def test_installed_before_any_early_return(self):
        main = _code(_function(_read(STARTUP), 'main'))
        call = main.find('install_ribbon_memory()')
        ret = main.find('return')
        self.assertGreater(call, 0)
        self.assertTrue(ret == -1 or call < ret)

    def test_failure_is_logged(self):
        main = _function(_read(STARTUP), 'main')
        self.assertRegex(main, r'install_ribbon_memory\(\)\s*\n\s*if state == "failed":\s*\n\s*_log\(')

    def test_never_raises_outside_revit(self):
        body = _read(STARTUP).rsplit('main()', 1)[0]
        ns = {'__name__': 'startup_probe', '__file__': STARTUP}
        exec(compile(body, STARTUP, 'exec'), ns)
        state, detail = ns['install_ribbon_memory']()
        self.assertEqual(state, 'failed')
        self.assertTrue(detail)


if __name__ == '__main__':
    unittest.main(verbosity=2)
