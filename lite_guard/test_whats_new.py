#!/usr/bin/env python3
# -*- coding: utf-8 -*-
""""What's new" popup once per version (lib/core/whats_new.py).

    python3 lite_guard/test_whats_new.py

Revit is not available here: Idling is a fake event and the dialog is a
recorder. What the real TaskDialog looks like must be checked in Revit.
Plain CPython 3, standard library only. CI runs it next to guard.py.
"""
import io
import json
import os
import re
import shutil
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
if LIB not in sys.path:
    sys.path.insert(0, LIB)

from core import whats_new as wn  # noqa: E402

CHANGELOG = """# Changelog

## [Unreleased]

### Added
- Not released yet.

## [2.2.0] - 2026-11-01

### Added
- **Big** feature, see [T3Lab PDF](https://pdf.t3lab.space/).
  - Nested `detail`.

## [2.1.0] - 2026-10-19

### Fixed
- Fix one
  wrapped onto two lines.

## [2.0.0] - 2026-10-09

### Added
- Old thing.
"""


class FakeEvent(object):
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


class FakeApp(object):
    def __init__(self):
        self.Idling = FakeEvent()


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.ext = os.path.join(self.tmp, 'ext')
        os.makedirs(self.ext)
        self._appdata = os.environ.get('APPDATA')
        os.environ['APPDATA'] = os.path.join(self.tmp, 'appdata')
        self.write('CHANGELOG.md', CHANGELOG)
        self.shown = []
        self._orig_show = wn.show_dialog
        wn.show_dialog = lambda version, lines: self.shown.append((version, lines))

    def tearDown(self):
        wn.show_dialog = self._orig_show
        if self._appdata is None:
            os.environ.pop('APPDATA', None)
        else:
            os.environ['APPDATA'] = self._appdata
        shutil.rmtree(self.tmp, ignore_errors=True)

    def write(self, name, text):
        with io.open(os.path.join(self.ext, name), 'w', encoding='utf-8') as fh:
            fh.write(text)

    def seen(self):
        return wn._settings().load_settings().get(wn.SEEN_KEY)

    def start(self, version, app=None):
        self.write('version.txt', version + '\n')
        app = app or FakeApp()
        return app, wn.install(app, ext_dir=self.ext)


class TestNotes(unittest.TestCase):
    def test_first_popup_shows_only_the_current_release(self):
        lines = wn.notes_for(CHANGELOG, '2.1.0', None)
        self.assertEqual(lines[0], 'Version 2.1.0 - 2026-10-19')
        self.assertNotIn('Version 2.0.0 - 2026-10-09', lines)

    def test_skipped_updates_show_every_release_since_the_last_seen(self):
        lines = wn.notes_for(CHANGELOG, '2.2.0', '2.0.0')
        self.assertIn('Version 2.2.0 - 2026-11-01', lines)
        self.assertIn('Version 2.1.0 - 2026-10-19', lines)
        self.assertNotIn('Version 2.0.0 - 2026-10-09', lines)
        self.assertFalse(any('Not released yet' in l for l in lines))

    def test_upto_stops_at_the_installed_version(self):
        lines = wn.notes_for(CHANGELOG, '2.1.0', '2.0.0')
        self.assertNotIn('Version 2.2.0 - 2026-11-01', lines)

    def test_markdown_becomes_plain_text(self):
        lines = wn.notes_for(CHANGELOG, '2.2.0', None)
        self.assertIn('  • Big feature, see T3Lab PDF (https://pdf.t3lab.space/).', lines)
        self.assertIn('      - Nested detail.', lines)

    def test_wrapped_bullet_is_joined(self):
        lines = wn.notes_for(CHANGELOG, '2.1.0', None)
        self.assertIn('  • Fix one wrapped onto two lines.', lines)

    def test_pending(self):
        self.assertTrue(wn.is_pending('2.1.0', None))
        self.assertTrue(wn.is_pending('2.1.0', '2.0.0'))
        self.assertFalse(wn.is_pending('2.1.0', '2.1.0'))
        self.assertFalse(wn.is_pending('2.0.0', '2.1.0'))     # downgrade

    def test_repo_changelog_has_the_current_version(self):
        with io.open(os.path.join(ROOT, 'version.txt'), encoding='utf-8') as fh:
            version = fh.read().strip()
        with io.open(os.path.join(ROOT, 'CHANGELOG.md'), encoding='utf-8') as fh:
            text = fh.read()
        self.assertTrue(wn.notes_for(text, version, None),
                        'CHANGELOG.md has no "## [%s]" section: the update popup '
                        'would show nothing' % version)


class TestInstall(Base):
    def test_shows_once_on_first_idling(self):
        app, (state, _) = self.start('2.1.0')
        self.assertEqual(state, wn.PENDING)
        self.assertEqual(self.shown, [])            # not before Idling
        app.Idling.fire()
        self.assertEqual(len(self.shown), 1)
        self.assertEqual(self.shown[0][0], '2.1.0')
        self.assertEqual(self.seen(), '2.1.0')
        self.assertEqual(app.Idling.handlers, [])   # unhooked
        app.Idling.fire()
        self.assertEqual(len(self.shown), 1)

    def test_next_start_same_version_shows_nothing(self):
        app, _ = self.start('2.1.0')
        app.Idling.fire()
        app2, (state, _) = self.start('2.1.0')
        self.assertEqual(state, wn.SKIPPED)
        self.assertEqual(app2.Idling.handlers, [])

    def test_next_update_shows_again(self):
        app, _ = self.start('2.1.0')
        app.Idling.fire()
        app2, (state, _) = self.start('2.2.0')
        self.assertEqual(state, wn.PENDING)
        app2.Idling.fire()
        self.assertEqual([v for v, _ in self.shown], ['2.1.0', '2.2.0'])
        self.assertIn('Version 2.2.0 - 2026-11-01', self.shown[1][1])
        self.assertNotIn('Version 2.1.0 - 2026-10-19', self.shown[1][1])

    def test_reload_before_idling_shows_once(self):
        app = FakeApp()
        self.start('2.1.0', app)
        self.start('2.1.0', app)                    # pyRevit Reload
        app.Idling.fire()
        self.assertEqual(len(self.shown), 1)

    def test_closed_before_idling_shows_on_next_start(self):
        self.start('2.1.0')                         # never idles
        self.assertIsNone(self.seen())
        app, (state, _) = self.start('2.1.0')
        self.assertEqual(state, wn.PENDING)

    def test_no_changelog_entry_marks_seen_without_popup(self):
        app, (state, _) = self.start('9.9.9')
        self.assertEqual(state, wn.SKIPPED)
        self.assertEqual(self.seen(), '9.9.9')
        self.assertEqual(app.Idling.handlers, [])

    def test_dialog_error_does_not_raise(self):
        def boom(version, lines):
            raise RuntimeError('no Revit here')
        wn.show_dialog = boom
        app, _ = self.start('2.1.0')
        app.Idling.fire()
        self.assertEqual(self.seen(), '2.1.0')

    def test_no_idling_app(self):
        self.write('version.txt', '2.1.0\n')
        state, _ = wn.install(object(), ext_dir=self.ext)
        self.assertEqual(state, wn.FAILED)


class TestWiring(unittest.TestCase):
    def test_startup_installs_the_popup(self):
        with io.open(os.path.join(ROOT, 'startup.py'), encoding='utf-8') as fh:
            src = fh.read()
        self.assertTrue(re.search(r'^\s+state, detail = install_whats_new\(_revit_handle\(\)\)',
                                  src, re.M))

    def test_updater_reuses_the_parser(self):
        with io.open(os.path.join(LIB, 'core', 'updater.py'), encoding='utf-8') as fh:
            src = fh.read()
        self.assertIn('from core.whats_new import extract_whats_new', src)
        self.assertNotIn('def extract_whats_new(', src)

    def test_ironpython_safe_syntax(self):
        with io.open(os.path.join(LIB, 'core', 'whats_new.py'), encoding='utf-8') as fh:
            src = fh.read()
        self.assertNotRegex(src, r'\bf["\']')       # no f-strings
        self.assertNotIn('nonlocal', src)


if __name__ == '__main__':
    unittest.main(verbosity=1)
