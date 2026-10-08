#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests for lib/core/ribbon_retire.py (old ribbon folders an update left behind).

    python3 lite_guard/test_ribbon_retire.py

  * the folders it retires are not shipped any more, and the files it waits for
    are -- otherwise it would move live buttons out of the ribbon on every
    Revit start
  * a fake extension with the old and the new layout: the old folders move to
    _retired_ribbon, per-user settings are kept, nothing else is touched
  * it does nothing while the new layout is incomplete

Plain CPython 3, standard library only -- no Revit, no pyRevit. CI runs it next
to guard.py (.github/workflows/lite-guard.yml).
"""
import importlib.util
import io
import os
import shutil
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_spec = importlib.util.spec_from_file_location(
    't3_ribbon_retire', os.path.join(ROOT, 'lib', 'core', 'ribbon_retire.py'))
rr = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(rr)


def _path(root, rel):
    return os.path.join(root, *rel.split('/'))


def _write(root, rel, text=u'x'):
    path = _path(root, rel)
    if not os.path.isdir(os.path.dirname(path)):
        os.makedirs(os.path.dirname(path))
    with io.open(path, 'w', encoding='utf-8') as fh:
        fh.write(text)


class AgainstThisRepo(unittest.TestCase):
    def test_retired_folders_are_not_shipped(self):
        for rel in rr.RETIRED:
            self.assertFalse(os.path.exists(_path(ROOT, rel)),
                             'ribbon_retire.RETIRED lists a folder this repo ships: ' + rel)

    def test_new_layout_markers_are_shipped(self):
        for rel in rr.NEW_LAYOUT:
            self.assertTrue(os.path.isfile(_path(ROOT, rel)),
                            'ribbon_retire.NEW_LAYOUT names a missing file: ' + rel)

    def test_user_files_sit_in_retired_folders(self):
        for rel, _parts in rr.USER_FILES:
            self.assertTrue(any(rel.startswith(r + '/') for r in rr.RETIRED), rel)

    def test_backup_folder_is_never_a_ribbon_bundle(self):
        self.assertNotIn('.', rr.BACKUP_DIR)

    def test_repo_runs_as_a_no_op(self):
        backup = tempfile.mkdtemp()
        try:
            self.assertEqual(rr.retire_old_ribbon_folders(ROOT, backup_root=backup,
                                                          data_dir=backup), [])
            self.assertEqual(os.listdir(backup), [])
        finally:
            shutil.rmtree(backup)


class FakeExtension(unittest.TestCase):
    def setUp(self):
        self.ext = tempfile.mkdtemp()
        self.data = tempfile.mkdtemp()
        for rel in rr.NEW_LAYOUT:
            _write(self.ext, rel, u'layout:\n')
        # a live button inside a panel that also holds retired leftovers
        _write(self.ext, 'T3Lab.tab/Views & Sheets.panel/ManaViews.pushbutton/script.py')
        _write(self.ext, 'T3Lab.tab/Support.panel/T3LabAssistant.pushbutton/script.py')
        # what a download-and-copy update leaves behind
        _write(self.ext, 'T3Lab.tab/Data.panel/IFC-SG.pushbutton/script.py')
        _write(self.ext, 'T3Lab.tab/Modeling & Datum.panel/DatumSync.pushbutton/script.py')
        _write(self.ext, 'T3Lab.tab/Views & Sheets.panel/SheetGen.pushbutton/script.py')
        _write(self.ext, 'T3Lab.tab/Support.panel/UI.stack/Ribbon Names.pushbutton/script.py')
        _write(self.ext, 'T3Lab.tab/Support.panel/UI.stack/Ribbon Names.pushbutton/'
                         'dqt_ribbon_map.json', u'{"Architecture": "A"}')
        _write(self.ext, 'T3Lab.tab/Support.panel/UI.stack/BG Theme.pushbutton/'
                         'dqt_bg_config.json', u'{"preset": "Black"}')

    def tearDown(self):
        shutil.rmtree(self.ext)
        shutil.rmtree(self.data)

    def run_retire(self):
        return rr.retire_old_ribbon_folders(self.ext, data_dir=self.data)

    def test_moves_only_the_retired_folders(self):
        moved = self.run_retire()
        self.assertEqual(sorted(moved), sorted([
            'T3Lab.tab/Data.panel',
            'T3Lab.tab/Modeling & Datum.panel',
            'T3Lab.tab/Views & Sheets.panel/SheetGen.pushbutton',
            'T3Lab.tab/Support.panel/UI.stack',
        ]))
        for rel in moved:
            self.assertFalse(os.path.exists(_path(self.ext, rel)), rel)
        self.assertTrue(os.path.isfile(_path(
            self.ext, 'T3Lab.tab/Views & Sheets.panel/ManaViews.pushbutton/script.py')))
        self.assertTrue(os.path.isfile(_path(
            self.ext, 'T3Lab.tab/Support.panel/T3LabAssistant.pushbutton/script.py')))
        for rel in rr.NEW_LAYOUT:
            self.assertTrue(os.path.isfile(_path(self.ext, rel)), rel)

    def test_moved_folders_are_kept_in_the_backup(self):
        self.run_retire()
        backup = os.path.join(self.ext, rr.BACKUP_DIR)
        stamps = [d for d in os.listdir(backup) if os.path.isdir(os.path.join(backup, d))]
        self.assertEqual(len(stamps), 1)
        self.assertTrue(os.path.isfile(_path(
            os.path.join(backup, stamps[0]), 'T3Lab.tab/Data.panel/IFC-SG.pushbutton/script.py')))
        self.assertTrue(os.path.isfile(os.path.join(backup, 'retired.log')))

    def test_keeps_per_user_settings(self):
        self.run_retire()
        with io.open(os.path.join(self.data, 'ribbon_names', 'ribbon_map.json'),
                     encoding='utf-8') as fh:
            self.assertEqual(fh.read(), u'{"Architecture": "A"}')
        self.assertTrue(os.path.isfile(
            os.path.join(self.data, 'bg_theme', 'bg_theme_config.json')))

    def test_never_overwrites_newer_settings(self):
        _write(self.data, 'ribbon_names/ribbon_map.json', u'{"Architecture": "Arch"}')
        self.run_retire()
        with io.open(os.path.join(self.data, 'ribbon_names', 'ribbon_map.json'),
                     encoding='utf-8') as fh:
            self.assertEqual(fh.read(), u'{"Architecture": "Arch"}')

    def test_waits_for_the_complete_new_layout(self):
        os.remove(_path(self.ext, rr.NEW_LAYOUT[0]))
        self.assertEqual(self.run_retire(), [])
        self.assertTrue(os.path.isdir(_path(self.ext, 'T3Lab.tab/Data.panel')))
        self.assertFalse(os.path.exists(os.path.join(self.ext, rr.BACKUP_DIR)))

    def test_second_run_is_a_no_op(self):
        self.assertTrue(self.run_retire())
        self.assertEqual(self.run_retire(), [])

    def test_a_locked_folder_is_left_for_the_next_start(self):
        real = rr.os.rename

        def rename(src, dst):
            if src.endswith('Data.panel'):
                raise OSError('in use')
            return real(src, dst)

        rr.os.rename = rename
        try:
            moved = self.run_retire()
        finally:
            rr.os.rename = real
        self.assertNotIn('T3Lab.tab/Data.panel', moved)
        self.assertTrue(os.path.isfile(_path(
            self.ext, 'T3Lab.tab/Data.panel/IFC-SG.pushbutton/script.py')))
        self.assertIn('T3Lab.tab/Modeling & Datum.panel', moved)


if __name__ == '__main__':
    unittest.main()
