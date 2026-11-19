# -*- coding: utf-8 -*-
"""
CPython 3 test: the pyRevit reload patch is for pyRevit before 7.0 only
(lib/pyrevit_patches.py).

Run:  python3 lite_guard/test_pyrevit_patch_versions.py      Exit 0 = all pass.

pyRevit 6.5 shuts CPython down on a Reload and cannot start it again, so
T3Lab rewrites `_clear_running_engines` in pyRevit's sessionmgr.py. pyRevit
7.0 fixed that itself, but the function is still letter for letter the 6.5.5
one in 7.0.0, 7.0.1 and 7.1.0 (read in those sources, 2026-11-19) - so
matching the text alone patched pyRevit 7 too. The version decides now.
"""
import io
import os
import sys
import tempfile
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
if not os.path.isdir(LIB):
    LIB = os.path.join(REPO, 'lib')         # T3LabLite layout
sys.path.insert(0, LIB)

import pyrevit_patches as patches    # noqa: E402  (needs LIB on sys.path)

HEAD = u'"""pyRevit session manager (test copy)."""\nimport os\n\n\n'
TAIL = u'\n\ndef load_session():\n    _clear_running_engines()\n'


class Clone(object):
    """A pyRevit clone on disk: just the two files the patch looks at."""

    def __init__(self, folder, version, newline=u'\n'):
        package = os.path.join(folder, 'pyRevit-Master', 'pyrevitlib', 'pyrevit')
        os.makedirs(os.path.join(package, 'loader'))
        self.path = os.path.join(package, 'loader', 'sessionmgr.py')
        self.version_file = os.path.join(package, 'version')
        self.original = (HEAD + patches.ORIGINAL + TAIL).replace(u'\n', newline)
        with io.open(self.path, 'w', encoding='utf-8', newline='') as f:
            f.write(self.original)
        self.set_version(version)

    def set_version(self, version):
        if version is None:
            if os.path.exists(self.version_file):
                os.remove(self.version_file)
            return
        with io.open(self.version_file, 'w', encoding='utf-8') as f:
            f.write(version)

    def text(self):
        with io.open(self.path, 'r', encoding='utf-8', newline='') as f:
            return f.read()

    def backup_exists(self):
        return os.path.exists(self.path + patches.BACKUP_SUFFIX)


class VersionGateTests(unittest.TestCase):

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self._count = 0

    def clone(self, version, newline=u'\n'):
        self._count += 1
        return Clone(os.path.join(self._tmp.name, str(self._count)), version, newline)

    def test_version_is_read_from_the_clone(self):
        self.assertEqual(patches.pyrevit_version(self.clone(u'6.5.5.26237+2044').path), (6, 5))
        self.assertEqual(patches.pyrevit_version(self.clone(u'7.1.0.26319+1907\n').path), (7, 1))
        self.assertEqual(patches.pyrevit_version(self.clone(u'﻿7.0.1.26312+1155').path), (7, 0))
        self.assertIsNone(patches.pyrevit_version(self.clone(None).path))
        self.assertIsNone(patches.pyrevit_version(self.clone(u'develop').path))

    def test_pyrevit_6_5_is_patched(self):
        clone = self.clone(u'6.5.5.26237+2044')
        self.assertEqual(patches.status(clone.path), patches.MISSING)
        self.assertEqual(patches.ensure(clone.path), (patches.APPLIED, clone.path))
        self.assertIn(patches.MARKER, clone.text())
        self.assertTrue(clone.backup_exists())
        self.assertEqual(patches.ensure(clone.path), (patches.ALREADY, clone.path))

    def test_pyrevit_7_is_left_exactly_as_it_shipped(self):
        for version in (u'7.0.0.26278+2233', u'7.0.1.26312+1155',
                        u'7.1.0.26319+1907', u'8.0.0'):
            clone = self.clone(version)
            self.assertEqual(patches.status(clone.path), patches.NOT_NEEDED, version)
            state, detail = patches.ensure(clone.path)
            self.assertEqual(state, patches.NOT_NEEDED, version)
            self.assertIn('by itself', detail)
            self.assertEqual(clone.text(), clone.original, version)
            self.assertFalse(clone.backup_exists(), version)

    def test_a_patch_an_older_t3lab_put_on_pyrevit_7_is_taken_out(self):
        for newline in (u'\n', u'\r\n'):
            clone = self.clone(u'6.5.5.26237+2044', newline)
            self.assertEqual(patches.ensure(clone.path)[0], patches.APPLIED)
            clone.set_version(u'7.1.0.26319+1907')      # same function text in 7.x
            self.assertEqual(patches.ensure(clone.path), (patches.REMOVED, clone.path))
            self.assertEqual(clone.text(), clone.original)
            self.assertEqual(patches.ensure(clone.path)[0], patches.NOT_NEEDED)

    def test_an_unreadable_version_keeps_the_old_behaviour(self):
        # Missing the patch on 6.5 breaks every tool after a Reload.
        clone = self.clone(None)
        self.assertTrue(patches.needs_patch(clone.path))
        self.assertEqual(patches.ensure(clone.path)[0], patches.APPLIED)

    def test_a_changed_function_is_never_touched(self):
        clone = self.clone(u'6.5.5.26237+2044')
        changed = clone.original.replace(u'# clear the cached engines',
                                         u'# clear the engines')
        with io.open(clone.path, 'w', encoding='utf-8', newline='') as f:
            f.write(changed)
        self.assertEqual(patches.ensure(clone.path)[0], patches.UNSUPPORTED)
        self.assertEqual(clone.text(), changed)

    def test_startup_stays_quiet_on_pyrevit_7(self):
        with io.open(os.path.join(os.path.dirname(LIB), 'startup.py'),
                     encoding='utf-8') as f:
            src = f.read()
        self.assertIn('elif state == "removed":', src)
        self.assertIn('elif state not in (None, "already", "not needed"):', src)


if __name__ == '__main__':
    unittest.main(verbosity=1)
