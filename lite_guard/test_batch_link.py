# -*- coding: utf-8 -*-
"""
Tests for BatchLink logic: file scanning, backup detection, item formatting,
and the Revit-free helpers in Snippets/_links.py (workset name matching,
display-mode mapping).
Run: python dev/test_batch_link.py
"""
import os
import sys
import re
import unittest
import tempfile
import shutil

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB_DIR = os.path.join(REPO, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

BACKUP_REGEX = re.compile(r'\.\d{3,4}\.rvt$', re.IGNORECASE)


def _load_links_module():
    """Exec Snippets/_links.py with the Revit and .NET imports stubbed out.

    The module is only importable inside Revit, so the Revit-free helpers are
    exercised against the real shipped source rather than a copy.
    """
    import types

    class _AnyMeta(type):
        """Any attribute lookup yields another permissive placeholder type."""
        def __getattr__(cls, item):
            return _make_any(item)

    def _make_any(name):
        return _AnyMeta(str(name), (), {})

    def _stub(name):
        mod = types.ModuleType(name)
        mod.__getattr__ = _make_any
        return mod

    saved = {}
    stubs = {
        'Autodesk': _stub('Autodesk'),
        'Autodesk.Revit': _stub('Autodesk.Revit'),
        'Autodesk.Revit.DB': _stub('Autodesk.Revit.DB'),
        'System': _stub('System'),
        'System.Collections': _stub('System.Collections'),
        'System.Collections.Generic': _stub('System.Collections.Generic'),
    }
    stubs['System.Collections.Generic'].List = {}
    # wire the dotted attribute chain so `import a.b.c as x` resolves
    stubs['Autodesk'].Revit = stubs['Autodesk.Revit']
    stubs['Autodesk.Revit'].DB = stubs['Autodesk.Revit.DB']
    stubs['System'].Collections = stubs['System.Collections']
    stubs['System.Collections'].Generic = stubs['System.Collections.Generic']
    for name, mod in stubs.items():
        saved[name] = sys.modules.get(name)
        sys.modules[name] = mod

    try:
        source_path = os.path.join(LIB_DIR, 'Snippets', '_links.py')
        with open(source_path, 'r', encoding='utf-8') as handle:
            source = handle.read()
        module = types.ModuleType('t3_links_under_test')
        module.__file__ = source_path
        exec(compile(source, source_path, 'exec'), module.__dict__)
        return module
    finally:
        for name, previous in saved.items():
            if previous is None:
                sys.modules.pop(name, None)
            else:
                sys.modules[name] = previous


class _FakeWorkset(object):
    def __init__(self, workset_id, name, is_open=True):
        self.workset_id = workset_id
        self.name = name
        self.is_open = is_open


class _FakeLinkRecord(object):
    """Stands in for Snippets._links.LinkRecord in the instance-fan-out tests."""

    def __init__(self, instance_ids, instance_id=None):
        self.instance_ids = list(instance_ids)
        self.instance_id = (instance_id if instance_id is not None
                            else (instance_ids[0] if instance_ids else None))


class TestLinkHelpers(unittest.TestCase):
    """Revit-free helpers from Snippets/_links.py."""

    @classmethod
    def setUpClass(cls):
        cls.links = _load_links_module()

    def test_split_workset_ids_matches_by_name(self):
        worksets = [_FakeWorkset(1, 'Shared Levels'),
                    _FakeWorkset(2, 'Link Architecture'),
                    _FakeWorkset(3, 'Furniture')]
        open_ids, close_ids = self.links.split_workset_ids(
            worksets, {'Shared Levels', 'Link Architecture'}, {'Furniture'})
        self.assertEqual(open_ids, [1, 2])
        self.assertEqual(close_ids, [3])

    def test_split_workset_ids_leaves_unlisted_untouched(self):
        """A workset named in neither set must not be opened or closed."""
        worksets = [_FakeWorkset(7, 'Only In This Link')]
        open_ids, close_ids = self.links.split_workset_ids(
            worksets, {'Shared Levels'}, {'Furniture'})
        self.assertEqual(open_ids, [])
        self.assertEqual(close_ids, [])

    def test_split_workset_ids_handles_empty(self):
        self.assertEqual(self.links.split_workset_ids(None, set(), set()), ([], []))
        self.assertEqual(self.links.split_workset_ids([], {'x'}, {'y'}), ([], []))

    def test_workset_config_ids_covers_every_workset(self):
        """Worksets the grid never listed keep their own state, not "open"."""
        worksets = [_FakeWorkset(1, 'Shared Levels', is_open=True),
                    _FakeWorkset(2, 'Furniture', is_open=True),
                    _FakeWorkset(3, 'Only In This Link', is_open=False),
                    _FakeWorkset(4, 'Also Only Here', is_open=True)]
        open_ids, close_ids = self.links.workset_config_ids(
            worksets, {'Shared Levels'}, {'Furniture'})
        # 1 ticked open, 4 unlisted but currently open -> open.
        self.assertEqual(sorted(open_ids), [1, 4])
        # 2 ticked closed, 3 unlisted and currently closed -> closed.
        self.assertEqual(sorted(close_ids), [2, 3])

    def test_workset_config_ids_handles_empty(self):
        self.assertEqual(self.links.workset_config_ids(None, set(), set()), ([], []))
        self.assertEqual(self.links.workset_config_ids([], {'x'}, {'y'}), ([], []))

    def test_placed_instance_ids_returns_every_instance(self):
        record = _FakeLinkRecord([11, 12, 13])
        self.assertEqual(self.links._placed_instance_ids(record), [11, 12, 13])

    def test_placed_instance_ids_falls_back_to_single(self):
        record = _FakeLinkRecord([], instance_id=7)
        self.assertEqual(self.links._placed_instance_ids(record), [7])

    def test_placed_instance_ids_of_unplaced_link_is_empty(self):
        self.assertEqual(self.links._placed_instance_ids(None), [])
        self.assertEqual(self.links._placed_instance_ids(_FakeLinkRecord([])), [])

    def test_display_modes_order(self):
        self.assertEqual(self.links.DISPLAY_MODES,
                         ("By Host View", "By Linked View", "Custom"))

    def test_mode_index_round_trip(self):
        self.assertEqual(self.links._mode_index('ByHostView'), 0)
        self.assertEqual(self.links._mode_index('ByLinkView'), 1)
        self.assertEqual(self.links._mode_index('Custom'), 2)
        self.assertEqual(self.links._mode_index('SomethingElse'), 0)

    def test_only_view_filters_accepts_custom(self):
        """Revit rejects LinkVisibility.Custom on every aspect but View Filters."""
        allowing = [key for key, _label, _kind, custom in self.links.CUSTOM_ASPECTS if custom]
        self.assertEqual(allowing, ['ViewFilterType'])
        self.assertEqual(len(self.links.aspect_modes(True)), 3)
        self.assertEqual(len(self.links.aspect_modes(False)), 2)

    def test_every_aspect_has_a_known_kind(self):
        for _key, _label, kind, _custom in self.links.CUSTOM_ASPECTS:
            self.assertIn(kind, ('prop', 'discipline', 'detail'))


# ── VIEW DISPLAY: category tabs copied from a view template ──────────────────
# RevitLinkGraphicsSettings has no member for the Model / Annotation tabs, so
# "Show annotation categories" can only reach a view through a template's
# "V/G Overrides RVT Links" block.

SCALE, VG_MODEL = 901, 902
RVT_LINKS = 903          # stands in for BuiltInParameter.VIS_GRAPHICS_RVT_LINKS


class _Id(object):
    def __init__(self, value):
        self.Value = value

    def __repr__(self):
        return "Id(%d)" % self.Value


class _NetListOf(list):
    def Add(self, item):
        self.append(item)


class _NetListFactory(object):
    def __getitem__(self, _item_type):
        return _NetListOf


class _FakeTemplate(object):
    def __init__(self, tid, name, params, non_controlled, log):
        self.Id = _Id(tid)
        self.Name = name
        self.IsTemplate = True
        self._params = [_Id(p) for p in params]
        self._non_controlled = [_Id(p) for p in non_controlled]
        self._log = log

    def GetTemplateParameterIds(self):
        return list(self._params)

    def GetNonControlledTemplateParameterIds(self):
        return list(self._non_controlled)

    def SetNonControlledTemplateParameterIds(self, ids):
        self._non_controlled = list(ids)
        self._log.append(('include', sorted(i.Value for i in ids)))

    def controlled(self):
        nc = {i.Value for i in self._non_controlled}
        return sorted(p.Value for p in self._params if p.Value not in nc)


class _FakeDoc(object):
    def __init__(self, log):
        self.elements = {}
        self._log = log

    def GetElement(self, element_id):
        return self.elements.get(element_id.Value)

    def Regenerate(self):
        self._log.append(('regenerate',))


class _FakeView(object):
    def __init__(self, doc, template_id=-1, valid=True):
        self.Document = doc
        self._template_id = _Id(template_id)
        self._valid = valid
        self._log = doc._log

    def IsValidViewTemplate(self, _tid):
        return self._valid

    @property
    def ViewTemplateId(self):
        return self._template_id

    @ViewTemplateId.setter
    def ViewTemplateId(self, value):
        self._template_id = value
        tpl = self.Document.GetElement(value)
        # what Revit writes onto the view while the template is assigned
        self._log.append(('assign', value.Value, tpl.controlled() if tpl else []))


class TestTemplateLinkDisplay(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.links = _load_links_module()
        cls.links.rvt_links_param_id = lambda: _Id(RVT_LINKS)
        cls.links.NetList = _NetListFactory()

    def setUp(self):
        self.log = []
        self.doc = _FakeDoc(self.log)

    def _template(self, tid, name="LINK - Model only", non_controlled=()):
        tpl = _FakeTemplate(tid, name, [SCALE, VG_MODEL, RVT_LINKS],
                            non_controlled, self.log)
        self.doc.elements[tid] = tpl
        return tpl

    def test_template_controls_links_reads_the_include_box(self):
        self.assertTrue(self.links.template_controls_links(self._template(1)))
        self.assertFalse(self.links.template_controls_links(
            self._template(2, non_controlled=(RVT_LINKS,))))
        self.assertFalse(self.links.template_controls_links(None))
        no_links = _FakeTemplate(3, "Sheet", [SCALE], (), self.log)
        self.assertFalse(self.links.template_controls_links(no_links))

    def test_view_template_locking_links(self):
        self._template(5, "Plan - locked")
        self._template(6, "Plan - free", non_controlled=(RVT_LINKS,))
        self.assertIsNone(self.links.view_template_locking_links(_FakeView(self.doc, -1)))
        self.assertIsNone(self.links.view_template_locking_links(_FakeView(self.doc, 6)))
        locked = self.links.view_template_locking_links(_FakeView(self.doc, 5))
        self.assertEqual(locked.Name, "Plan - locked")

    def test_copy_moves_only_the_rvt_links_block_and_restores(self):
        own = self._template(6, "Plan - own", non_controlled=(RVT_LINKS,))
        src = self._template(9, non_controlled=(SCALE,))
        view = _FakeView(self.doc, template_id=6)
        ok, message = self.links.copy_link_display_from_template(view, src)
        self.assertTrue(ok, message)
        self.assertEqual(self.log, [
            ('include', [SCALE, VG_MODEL]),          # only RVT Links stays controlled
            ('assign', 9, [RVT_LINKS]),              # the view takes that block alone
            ('regenerate',),
            ('assign', 6, [SCALE, VG_MODEL]),        # its own template comes back
            ('include', [SCALE]),                    # source Include boxes as before
        ])
        self.assertEqual(view.ViewTemplateId.Value, 6)
        self.assertEqual(src.controlled(), [VG_MODEL, RVT_LINKS])
        self.assertEqual(own.controlled(), [SCALE, VG_MODEL])

    def test_copy_onto_a_view_without_template_leaves_none(self):
        src = self._template(9)
        view = _FakeView(self.doc, template_id=-1)
        ok, _message = self.links.copy_link_display_from_template(view, src)
        self.assertTrue(ok)
        self.assertEqual(view.ViewTemplateId.Value, -1)
        self.assertEqual(src.controlled(), [SCALE, VG_MODEL, RVT_LINKS])

    def test_copy_refuses_when_own_template_controls_links(self):
        self._template(5, "Plan - locked")
        src = self._template(9)
        ok, message = self.links.copy_link_display_from_template(
            _FakeView(self.doc, template_id=5), src)
        self.assertFalse(ok)
        self.assertIn("Plan - locked", message)
        self.assertEqual(self.log, [])                # nothing touched

    def test_copy_refuses_template_without_rvt_links_ticked(self):
        """Ticking it for a moment would push its links onto every view using it."""
        src = self._template(9, non_controlled=(RVT_LINKS,))
        ok, message = self.links.copy_link_display_from_template(_FakeView(self.doc), src)
        self.assertFalse(ok)
        self.assertIn("V/G Overrides RVT Links", message)
        self.assertEqual(self.log, [])

    def test_copy_refuses_template_of_another_view_type(self):
        src = self._template(9)
        ok, message = self.links.copy_link_display_from_template(
            _FakeView(self.doc, valid=False), src)
        self.assertFalse(ok)
        self.assertIn("view type", message)
        self.assertEqual(self.log, [])

    def test_no_api_parameter_is_a_clean_refusal(self):
        saved = self.links.rvt_links_param_id
        self.links.rvt_links_param_id = lambda: None
        try:
            ok, _message = self.links.copy_link_display_from_template(
                _FakeView(self.doc), self._template(9))
            self.assertFalse(ok)
            self.assertFalse(self.links.template_controls_links(self._template(10)))
        finally:
            self.links.rvt_links_param_id = saved


class TestDisplayTabContract(unittest.TestCase):
    """View Display tab: the template row is wired and the presets stay honest."""

    @classmethod
    def setUpClass(cls):
        xaml = _read(XAML_PATH)
        cls.tab = xaml[xaml.index('x:Name="tab_item_display"'):xaml.index('</TabControl>')]
        cls.dialog = _read(DIALOG_PATH)

    def test_template_row_is_wired(self):
        for name in ('cb_disp_template', 'btn_disp_copy_template',
                     'banner_view_template', 'lbl_view_template'):
            self.assertIn('x:Name="%s"' % name, self.tab)
            self.assertIn('self.%s' % name, self.dialog)
        self.assertIn('Click="disp_copy_template_clicked"', self.tab)
        self.assertIn('def disp_copy_template_clicked', self.dialog)

    def test_copy_is_confirmed_and_one_transaction(self):
        body = self.dialog[self.dialog.index('def disp_copy_template_clicked'):]
        body = body[:body.index('\n    def ')]
        self.assertIn('t3_confirm(', body)
        self.assertEqual(body.count('disposing(Transaction('), 1)
        self.assertIn('transaction.RollBack()', body)
        self.assertIn('TransactionStatus.Committed', body)

    def test_one_primary_button_only(self):
        self.assertNotIn('T3.Button.Primary', self.tab)

    def test_model_only_preset_does_not_promise_hidden_grids(self):
        start = self.dialog.index('def preset_model_clicked')
        body = self.dialog[start:self.dialog.index('\n    def ', start + 10)]
        self.assertIn('grids', body)
        self.assertNotIn("annotations stay hidden", self.tab)

    def test_apply_skips_display_write_on_template_locked_view(self):
        body = self.dialog[self.dialog.index('def _apply_display'):
                           self.dialog.index('def _refresh_disp_rows')]
        self.assertIn('_sync_view_template_lock(view)', body)
        self.assertIn('Template-locked', body)

class TestBatchLinkLogic(unittest.TestCase):
    def setUp(self):
        self.test_dir = tempfile.mkdtemp(prefix="t3_batch_link_test_")

    def tearDown(self):
        shutil.rmtree(self.test_dir, ignore_errors=True)

    def test_backup_regex(self):
        """Test detection of Revit numeric backup files."""
        self.assertTrue(BACKUP_REGEX.search("Project1.0001.rvt"))
        self.assertTrue(BACKUP_REGEX.search("Tower_A.0123.rvt"))
        self.assertTrue(BACKUP_REGEX.search("Hospital.001.rvt"))
        self.assertFalse(BACKUP_REGEX.search("Project1.rvt"))
        self.assertFalse(BACKUP_REGEX.search("Tower_A_0001.rvt"))
        self.assertFalse(BACKUP_REGEX.search("Project.rvt.txt"))

    def test_item_formatting(self):
        """Test RevitModelItem properties and size formatting."""
        # Mock item
        class MockItem(object):
            def __init__(self, size):
                self.file_size = size
            @property
            def FileSize(self):
                if self.file_size < 1024 * 1024:
                    return "{:.1f} KB".format(self.file_size / 1024.0)
                return "{:.1f} MB".format(self.file_size / (1024.0 * 1024.0))

        item_small = MockItem(512 * 1024)
        self.assertEqual(item_small.FileSize, "512.0 KB")

        item_large = MockItem(15 * 1024 * 1024)
        self.assertEqual(item_large.FileSize, "15.0 MB")

    def test_folder_scanning(self):
        """Test scanning folder for rvt files excluding temp and backup files."""
        # Create test files
        open(os.path.join(self.test_dir, "Model_A.rvt"), "w").close()
        open(os.path.join(self.test_dir, "Model_A.0001.rvt"), "w").close()
        open(os.path.join(self.test_dir, "Model_B.rvt"), "w").close()
        open(os.path.join(self.test_dir, "something.dwg"), "w").close()
        open(os.path.join(self.test_dir, "~$Model_C.rvt"), "w").close()

        # Scan
        found = []
        for f in os.listdir(self.test_dir):
            if not f.lower().endswith(".rvt"):
                continue
            if f.startswith("~$"):
                continue
            if BACKUP_REGEX.search(f):
                continue
            found.append(f)

        found.sort()
        self.assertEqual(found, ["Model_A.rvt", "Model_B.rvt"])


# ── LINK WORKSET TAB: pending / applied / failed state ───────────────────────

from GUI import BatchLinkWorksets as WS   # pure Python, no clr / Revit


class _Rec(object):
    def __init__(self, name, instance_count=1):
        self.name = name
        self.instance_count = instance_count


def _row(name="A.rvt", current=("ARC_Walls", 10), can_move=True,
         note="Ready", severity="Success", lock_reason=""):
    return WS.LinkWorksetRow(_Rec(name), current[0], current[1], can_move,
                             note, severity, lock_reason=lock_reason)


class TestWorksetRowState(unittest.TestCase):

    def test_fresh_row_is_clean(self):
        row = _row()
        self.assertEqual(row.row_state, "")
        self.assertEqual(row.WorksetName, "ARC_Walls")
        self.assertEqual(row.WorksetEditable, "yes")
        self.assertEqual((row.StatusText, row.Severity), ("Ready", "Success"))

    def test_stage_marks_pending_and_shows_target(self):
        row = _row()
        self.assertTrue(row.stage(20, "Links"))
        self.assertEqual(row.row_state, "pending")
        self.assertEqual(row.WorksetName, "Links")
        self.assertEqual((row.StatusText, row.Severity), ("Pending", "Warning"))
        self.assertIn("ARC_Walls", row.WorksetTip)
        self.assertIn("Links", row.WorksetTip)

    def test_same_pick_twice_is_no_change(self):
        row = _row()
        row.stage(20, "Links")
        self.assertFalse(row.stage(20, "Links"))

    def test_picking_current_workset_unstages(self):
        row = _row()
        row.stage(20, "Links")
        self.assertTrue(row.stage(10, "ARC_Walls"))
        self.assertEqual(row.row_state, "")
        self.assertIsNone(row.pending_id)
        self.assertEqual(row.StatusText, "Ready")

    def test_regeneration_does_not_wipe_green(self):
        """Grid refresh re-selects the shown value: an applied row stays green."""
        row = _row()
        row.stage(20, "Links")
        row.mark_applied("Moved")
        self.assertFalse(row.stage(20, "Links"))
        self.assertEqual(row.row_state, "applied")

    def test_regeneration_does_not_wipe_failed(self):
        row = _row()
        row.stage(20, "Links")
        row.mark_failed("Owned by bob")
        self.assertFalse(row.stage(20, "Links"))
        self.assertEqual(row.row_state, "failed")

    def test_locked_row_ignores_stage_and_explains(self):
        row = _row(can_move=False, note="Not placed", severity="Warning",
                   lock_reason="This link is not placed in the model.")
        self.assertFalse(row.stage(20, "Links"))
        self.assertEqual(row.WorksetEditable, "no")
        self.assertEqual(row.WorksetTip, "This link is not placed in the model.")
        self.assertEqual(WS.pending_count([row]), 0)

    def test_mark_applied_moves_current(self):
        row = _row()
        row.stage(20, "Links")
        row.mark_applied("Moved")
        self.assertEqual((row.current_id, row.current_label), (20, "Links"))
        self.assertIsNone(row.pending_id)
        self.assertEqual((row.row_state, row.StatusText, row.Severity),
                         ("applied", "Moved", "Success"))

    def test_mark_failed_keeps_pending_for_retry(self):
        row = _row()
        row.stage(20, "Links")
        row.mark_failed("Owned by bob\nmore detail")
        self.assertEqual(row.row_state, "failed")
        self.assertEqual(row.pending_id, 20)
        self.assertEqual(row.current_id, 10)
        self.assertEqual(row.Severity, "Danger")
        self.assertTrue(row.StatusText.startswith("Failed: Owned by bob"))
        self.assertNotIn("\n", row.StatusText)
        self.assertLessEqual(len(row.StatusText), WS.STATUS_MAX)
        self.assertIn("Owned by bob", row.WorksetTip)
        self.assertEqual(WS.pending_count([row]), 1)

    def test_long_failure_is_shortened_in_cell_not_tooltip(self):
        row = _row()
        row.stage(20, "Links")
        long = "The element is borrowed by another user on a different machine"
        row.mark_failed(long)
        self.assertLessEqual(len(row.StatusText), WS.STATUS_MAX)
        self.assertIn(long, row.WorksetTip)

    def test_clear_result_only_drops_green(self):
        a, b, c = _row("a"), _row("b"), _row("c")
        a.stage(20, "Links"); a.mark_applied()
        b.stage(20, "Links"); b.mark_failed("x")
        c.stage(20, "Links")
        self.assertEqual(WS.clear_results([a, b, c]), 1)
        self.assertEqual([r.row_state for r in (a, b, c)], ["", "failed", "pending"])

    def test_set_current_drops_pending_already_there(self):
        row = _row()
        row.stage(20, "Links")
        row.set_current("Links", 20)
        self.assertIsNone(row.pending_id)
        self.assertEqual(row.row_state, "")


class TestWorksetBulkAndApply(unittest.TestCase):

    def _rows(self):
        rows = [_row("a"), _row("b"), _row("c", current=("Links", 20)),
                _row("d", can_move=False, note="Not placed", severity="Warning")]
        for r in rows:
            r.IsSelected = True
        rows[1].IsSelected = False
        return rows

    def test_stage_checked_only_touches_checked_movable_rows(self):
        rows = self._rows()
        staged, unchanged = WS.stage_checked(rows, 20, "Links")
        self.assertEqual((staged, unchanged), (1, 1))     # a staged, c already there
        self.assertEqual([r.row_state for r in rows], ["pending", "", "", ""])
        self.assertEqual(WS.pending_count(rows), 1)

    def test_primary_label(self):
        self.assertEqual(WS.primary_label(0), "Apply")
        self.assertEqual(WS.primary_label(3), "Apply (3)")

    def test_tally_text_counts_pending(self):
        rows = self._rows()
        self.assertEqual(WS.tally_text(rows),
                         u"4 links · 3 can move · 2 checked")
        rows[0].stage(30, "MEP")
        rows[1].stage(30, "MEP")
        self.assertEqual(WS.tally_text(rows),
                         u"4 links · 3 can move · 2 checked · 2 pending")

    def test_apply_pending_records_each_row(self):
        rows = [_row("ok"), _row("same"), _row("bad"), _row("boom"), _row("idle")]
        for r in rows[:4]:
            r.stage(20, "Links")
        seen = []

        def move_one(row):
            seen.append(row.LinkName)
            if row.LinkName == "bad":
                return False, "Owned by bob"
            if row.LinkName == "boom":
                raise RuntimeError("API exploded\nstack")
            if row.LinkName == "same":
                return True, "Already there"
            return True, "Moved"

        steps = []
        result = WS.apply_pending(rows, move_one,
                                  step=lambda r, i, n: steps.append((i, n)))
        self.assertEqual(result, (1, 1, 2))
        self.assertEqual(seen, ["ok", "same", "bad", "boom"])   # idle untouched
        self.assertEqual(steps, [(1, 4), (2, 4), (3, 4), (4, 4)])
        self.assertEqual([r.row_state for r in rows],
                         ["applied", "applied", "failed", "failed", ""])
        self.assertEqual(rows[1].StatusText, "Already there")
        self.assertEqual(rows[3].result_message, "API exploded")
        # failed rows keep their target, so Apply (2) retries them
        self.assertEqual(WS.pending_count(rows), 2)

    def test_rollback_all_reverts_applied_rows_only(self):
        old_green = _row("old")
        old_green.stage(30, "MEP"); old_green.mark_applied()
        a, b = _row("a"), _row("b")
        a.stage(20, "Links"); b.stage(20, "Links")
        rows = [old_green, a, b]
        WS.snapshot_before(rows)
        WS.apply_pending(rows, lambda r: (True, "Moved") if r is a else (False, "Owned"))
        WS.rollback_all(rows, "Rolled back: commit failed")
        self.assertEqual((a.row_state, a.current_id, a.pending_id), ("failed", 10, 20))
        self.assertEqual(a.result_message, "Rolled back: commit failed")
        self.assertEqual(b.result_message, "Owned")                   # own reason kept
        self.assertEqual((old_green.row_state, old_green.current_id), ("applied", 30))

    def test_summary_text(self):
        self.assertEqual(WS.summary_text(1, 0, 0), u"1 link moved · 0 failed")
        self.assertEqual(WS.summary_text(3, 2, 1),
                         u"3 links moved · 2 already there · 1 failed")


# ── LINK WORKSET TAB: XAML / dialog contract ─────────────────────────────────

XAML_PATH = os.path.join(LIB_DIR, 'GUI', 'Tools', 'BatchLink.xaml')
DIALOG_PATH = os.path.join(LIB_DIR, 'GUI', 'BatchLinkDialog.py')


def _read(path):
    with open(path, 'r', encoding='utf-8') as handle:
        return handle.read()


class TestWorksetTabContract(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        xaml = _read(XAML_PATH)
        start = xaml.index('x:Name="grid_ws_links"')
        cls.grid = xaml[start:xaml.index('</DataGrid>', start)]
        cls.tab = xaml[xaml.index('x:Name="tab_item_worksets"'):
                       xaml.index('x:Name="tab_item_display"')]
        cls.dialog = _read(DIALOG_PATH)

    def test_single_editable_workset_column(self):
        self.assertIn('Header="WORKSET"', self.grid)
        self.assertNotIn('CURRENT WORKSET', self.grid)
        self.assertNotIn('NEW WORKSET', self.grid)
        self.assertIn('<ComboBox ItemsSource="{Binding Tag, RelativeSource='
                      '{RelativeSource AncestorType=DataGrid}}"', self.grid)

    def test_combo_reads_through_string_bridge(self):
        self.assertIn('x:Name="row_ws_name_text" Text="{Binding WorksetName}"', self.grid)
        self.assertIn('SelectedValue="{Binding Text, ElementName=row_ws_name_text, '
                      'Mode=OneWay}"', self.grid)

    def test_locked_rows_show_text_not_combo(self):
        """WorksetEditable reaches WPF as a PyObject: the triggers must read it
        through the hidden TextBlock, never {Binding WorksetEditable} itself."""
        self.assertIn('x:Name="row_ws_editable_text" Text="{Binding WorksetEditable}" '
                      'Visibility="Collapsed"', self.grid)
        self.assertEqual(self.grid.count(
            '<DataTrigger Binding="{Binding Text, ElementName=row_ws_editable_text}" '
            'Value="no">'), 2)
        self.assertNotIn('<DataTrigger Binding="{Binding WorksetEditable}"', self.grid)

    def test_states_painted_from_t3_tokens(self):
        # Cell: row_state copied into the cell's own string property (no template
        # to hold a TextBlock bridge). Combo inside the template: hidden TextBlock.
        self.assertIn('<Setter Property="AutomationProperties.ItemStatus" '
                      'Value="{Binding row_state}"/>', self.grid)
        self.assertIn('x:Name="row_ws_state_text" Text="{Binding row_state}" '
                      'Visibility="Collapsed"', self.grid)
        for state, token in (("pending", "Warning"), ("applied", "Success"),
                             ("failed", "Danger")):
            self.assertIn('<Trigger Property="AutomationProperties.ItemStatus" Value="%s">'
                          % state, self.grid)
            self.assertIn('<DataTrigger Binding="{Binding Text, ElementName=row_ws_state_text}" '
                          'Value="%s">' % state, self.grid)
            self.assertIn('{StaticResource T3.%s.Fill}' % token, self.grid)
        self.assertNotIn('<DataTrigger Binding="{Binding row_state}"', self.grid)
        self.assertNotRegex(self.grid, r'"#[0-9A-Fa-f]{3,8}"')

    def test_no_templated_selection_changed(self):
        """SelectionChanged= in a DataTemplate is never wired; Python adds it."""
        self.assertNotIn('SelectionChanged=', self.grid)
        self.assertIn('Selector.SelectionChangedEvent', self.dialog)

    def test_bulk_editor_stages_only(self):
        self.assertIn('x:Name="btn_ws_stage_checked"', self.tab)
        self.assertIn('Click="ws_stage_checked_clicked"', self.tab)
        self.assertIn('def ws_stage_checked_clicked', self.dialog)
        self.assertNotIn('ws_target_changed', self.tab + self.dialog)
        self.assertIn('nothing is written until you press Apply', self.tab)

    def test_apply_is_one_transaction_with_rollback(self):
        body = self.dialog[self.dialog.index('def _apply_worksets'):
                           self.dialog.index('def _confirm_discard_ws')]
        self.assertEqual(body.count('disposing(Transaction('), 1)
        self.assertIn('disposing(SubTransaction(', body)
        self.assertIn('TransactionStatus.Committed', body)
        self.assertIn('rollback_all', body)
        self.assertNotIn('Transaction(doc', body.replace('disposing(Transaction(doc', '')
                         .replace('disposing(SubTransaction(doc', ''))

    def test_primary_label_is_apply(self):
        self.assertIn('TAB_WORKSETS: "Apply"', self.dialog)
        self.assertNotIn('Move to Workset', self.dialog)

    def test_refresh_and_close_confirm_discard(self):
        reload_body = self.dialog[self.dialog.index('def reload_links_clicked'):]
        reload_body = reload_body[:reload_body.index('\n    def ')]
        self.assertIn('_confirm_discard_ws', reload_body)
        self.assertIn('self.Closing += self._on_closing', self.dialog)
        closing = self.dialog[self.dialog.index('def _on_closing'):]
        closing = closing[:closing.index('\n    def ')]
        self.assertIn('_confirm_discard_ws', closing)



class TestManageLinks(unittest.TestCase):
    """Manage Links tab: status/pin labels, workset naming, tally, tab contract."""

    @classmethod
    def setUpClass(cls):
        cls.links = _load_links_module()
        cls.xaml = _read(os.path.join(LIB_DIR, 'GUI', 'Tools', 'BatchLink.xaml'))
        cls.dialog = _read(os.path.join(LIB_DIR, 'GUI', 'BatchLinkDialog.py'))

    def test_status_labels_and_severity(self):
        self.assertEqual(self.links.status_label("Loaded"), ("Loaded", "Success"))
        self.assertEqual(self.links.status_label("NotFound"), ("Not found", "Danger"))
        self.assertEqual(self.links.status_label("LocallyUnloaded"),
                         ("Unloaded for me", "Warning"))
        self.assertEqual(self.links.status_label("Weird"), ("Weird", "Warning"))

    def test_pinned_label(self):
        self.assertEqual(self.links.pinned_label(0, 0), u"\u2014")
        self.assertEqual(self.links.pinned_label(2, 2), "Yes")
        self.assertEqual(self.links.pinned_label(0, 3), "No")
        self.assertEqual(self.links.pinned_label(1, 3), "1 of 3")

    def test_link_workset_name(self):
        self.assertEqual(self.links.link_workset_name("Link_", r"C:\x\ARC Model.rvt"),
                         "Link_ARC Model")
        self.assertEqual(self.links.link_workset_name("", "STR.rvt"), "STR")

    def test_manage_tally(self):
        import ast
        tree = ast.parse(self.dialog)
        tree.body = [n for n in tree.body
                     if isinstance(n, ast.FunctionDef) and n.name == 'manage_tally']
        scope = {}
        exec(compile(tree, 'BatchLinkDialog.py', 'exec'), scope)
        row = lambda loaded, checked=False: type('R', (), {
            'record': type('Rec', (), {'is_loaded': loaded})(), 'IsSelected': checked})()
        self.assertEqual(scope['manage_tally']([row(True), row(False, True)]),
                         "2 links · 1 loaded · 1 not loaded · 1 checked")
        self.assertEqual(scope['manage_tally']([row(True)]), "1 link · 1 loaded")

    def test_tabs_follow_the_daily_flow_and_match_the_constants(self):
        order = [self.xaml.index('x:Name="tab_item_{}"'.format(n))
                 for n in ('manage', 'link', 'worksets', 'display')]
        self.assertEqual(order, sorted(order))
        for name, tag in (('manage', 0), ('link', 1), ('worksets', 2), ('display', 3)):
            chip = self.xaml[self.xaml.index('x:Name="chip_tab_{}"'.format(name)):]
            self.assertIn('Tag="{}"'.format(tag), chip[:300])
        for const, value in (('TAB_MANAGE', 0), ('TAB_LINK', 1),
                             ('TAB_WORKSETS', 2), ('TAB_DISPLAY', 3)):
            self.assertIn('{} = {}'.format(const, value), self.dialog)

    def test_no_close_or_select_all_duplicates(self):
        # The title-bar X closes; each table's header checkbox selects all/none.
        for name in ('btn_cancel', 'btn_select_all', 'btn_select_none',
                     'btn_ws_select_all', 'btn_ws_select_none'):
            self.assertNotIn('x:Name="{}"'.format(name), self.xaml)
        for handler in ('cancel_button_clicked', 'def select_all_clicked',
                        'def ws_select_all_clicked'):
            self.assertNotIn(handler, self.dialog)

    def test_every_tab_has_the_same_toolbar_and_count_strip(self):
        for invert, refresh, count in (
                ('btn_mng_invert', 'btn_mng_refresh', 'lbl_mng_count'),
                ('btn_invert', 'btn_rescan', 'lbl_file_count'),
                ('btn_ws_invert', 'btn_ws_reload_links', 'lbl_ws_link_count'),
                ('btn_disp_invert', 'btn_disp_refresh', 'lbl_disp_count')):
            for name in (invert, refresh, count):
                self.assertIn('x:Name="{}"'.format(name), self.xaml)
        self.assertEqual(self.xaml.count('Style="{StaticResource T3.Search}"'), 4)
        self.assertEqual(self.xaml.count('Tag="Search'), 4)

    def test_reload_and_unload_never_run_inside_a_transaction(self):
        body = self.dialog[self.dialog.index('def _mng_run'):
                           self.dialog.index('def _mng_reload')]
        self.assertIn('if transaction_name:', body)
        for handler in ('def _mng_reload', 'def mng_unload_clicked',
                        'def mng_reload_from_clicked', 'def mng_to_cloud_clicked'):
            start = self.dialog.index(handler)
            end = self.dialog.index('\n    def ', start + 10)
            self.assertNotIn('transaction_name=', self.dialog[start:end])


# ── AUTODESK DOCS (CLOUD) LINKS ──────────────────────────────────────────────

from Snippets import _cloud_links as CL   # pure Python above its REVIT PART

P_GUID = "1b2c3d4e-0000-4000-8000-00000000aaaa"
P_OTHER = "1b2c3d4e-0000-4000-8000-00000000bbbb"
DC_ROOT = r"E:\DC\ACCDocs\WH_BIM360DOC\IR2\Project Files"
POD = ("BIM HUB", "01_WIP", "04_FAC", "02_JH", "01_Podium")


def _guid(n):
    return "00000000-0000-4000-8000-%012d" % n


def _model(n, path, project_guid=P_GUID, region="", source=CL.SOURCE_LINKED):
    return CL.CloudModel.from_path(path, project_guid, _guid(n), region, source)


def _dc(*segments):
    return "\\".join((DC_ROOT,) + segments)


class TestCloudPaths(unittest.TestCase):
    """Autodesk Docs / Desktop Connector path parsing (screenshot paths)."""

    def test_split_cloud_path(self):
        self.assertEqual(
            CL.split_cloud_path("Autodesk Docs://IR2/BIM HUB/01_WIP/04_FAC/02_JH/"
                                "01_Podium/MBSIR2_JCW_FC_POD_CM_L08-L09.rvt"),
            ("IR2", POD, "MBSIR2_JCW_FC_POD_CM_L08-L09.rvt"))
        self.assertEqual(CL.split_cloud_path("BIM 360://IR2/Model.rvt"), ("IR2", (), "Model.rvt"))
        self.assertIsNone(CL.split_cloud_path(r"E:\models\a.rvt"))
        self.assertIsNone(CL.split_cloud_path("Autodesk Docs://IR2"))

    def test_desktop_connector_parts(self):
        path = _dc(*(POD[:4] + ("04_Arena", "MBSIR2_JCW_FC_ARN_CM.rvt")))
        self.assertEqual(CL.desktop_connector_parts(path),
                         ("IR2", POD[:4] + ("04_Arena", "MBSIR2_JCW_FC_ARN_CM.rvt")))
        # A project-level folder narrows to the whole project.
        self.assertEqual(CL.desktop_connector_parts(r"E:\DC\ACCDocs\WH_BIM360DOC\IR2"),
                         ("IR2", ()))
        self.assertEqual(CL.desktop_connector_parts(
            r"C:\Users\me\BIM 360\Hub\P1\Project Files\A\m.rvt"), ("P1", ("A", "m.rvt")))
        # Plans / account level / plain folders are not model folders.
        self.assertIsNone(CL.desktop_connector_parts(
            r"E:\DC\ACCDocs\WH_BIM360DOC\IR2\Plans\a.rvt"))
        self.assertIsNone(CL.desktop_connector_parts(r"E:\DC\ACCDocs\WH_BIM360DOC"))
        self.assertIsNone(CL.desktop_connector_parts(r"D:\Projects\IR2\a.rvt"))

    def test_location_and_display_path_round_trip(self):
        m = _model(1, "Autodesk Docs://IR2/BIM HUB/01_WIP/a.rvt")
        self.assertEqual(m.location, "Autodesk Docs://IR2/BIM HUB/01_WIP")
        self.assertEqual(m.display_path, "Autodesk Docs://IR2/BIM HUB/01_WIP/a.rvt")

    def test_guid_text(self):
        self.assertEqual(CL.guid_text("{%s}" % P_GUID.upper()), P_GUID)
        self.assertEqual(CL.guid_text("b." + P_GUID), P_GUID)
        self.assertEqual(CL.guid_text("not-a-guid"), "")
        self.assertEqual(CL.guid_text(None), "")

    def test_parse_cache_path(self):
        root = r"C:\Users\me\AppData\Local\Autodesk\Revit\Autodesk Revit 2025\CollaborationCache"
        self.assertEqual(CL.parse_cache_path(r"%s\ABC123XYZ\%s\%s.rvt" % (root, P_GUID, _guid(7))),
                         (P_GUID, _guid(7)))
        self.assertEqual(CL.parse_cache_path(
            r"%s\ABC123XYZ\%s\CentralCache\%s.rvt" % (root, P_GUID, _guid(7))),
            (P_GUID, _guid(7)))
        # Backups and files outside the cache are not cloud models.
        self.assertIsNone(CL.parse_cache_path(r"%s\ABC\%s\%s.0001.rvt" % (root, P_GUID, _guid(7))))
        self.assertIsNone(CL.parse_cache_path(r"D:\x\%s\%s.rvt" % (P_GUID, _guid(7))))

    def test_cache_files_walks_only_the_wanted_projects(self):
        root = tempfile.mkdtemp(prefix="CollaborationCache_")
        try:
            cache = os.path.join(root, "CollaborationCache")
            for project, rel in ((P_GUID, _guid(1) + ".rvt"),
                                 (P_GUID, os.path.join("CentralCache", _guid(2) + ".rvt")),
                                 (P_GUID, os.path.join(_guid(1) + "_backup", "x.0001.rvt")),
                                 (P_OTHER, _guid(3) + ".rvt")):
                path = os.path.join(cache, "ABC123XYZ", project, rel)
                os.makedirs(os.path.dirname(path), exist_ok=True)
                open(path, "w").close()
            found = sorted((p, m) for _path, p, m in CL.cache_files(cache, [P_GUID]))
            self.assertEqual(found, [(P_GUID, _guid(1)), (P_GUID, _guid(2))])
            self.assertEqual(len(CL.cache_files(cache, [])), 3)
            self.assertEqual(CL.cache_files(os.path.join(root, "missing"), [P_GUID]), [])
        finally:
            shutil.rmtree(root, ignore_errors=True)


class TestCloudCatalog(unittest.TestCase):
    """Matching a Desktop Connector file to the cloud model behind it."""

    def setUp(self):
        self.cat = CL.CloudCatalog()
        self.cat.host_guid = _guid(99)
        self.cat.host_project_guid = P_GUID
        self.cat.host_region = "EMEA"
        self.cat.add(_model(99, "Autodesk Docs://IR2/BIM HUB/01_WIP/HOST.rvt", source=CL.SOURCE_HOST))
        self.pod = _model(1, "Autodesk Docs://IR2/" + "/".join(POD) + "/MBSIR2_WHP_ST_POD_CM_L06.rvt")
        self.cat.add(self.pod)
        self.cat.linked_guids.add(_guid(1))
        # Same file name in SHARED: a different model the matcher must not pick.
        self.cat.add(_model(2, "Autodesk Docs://IR2/BIM HUB/02_SHARED/MBSIR2_WHP_ST_POD_CM_L06.rvt"))
        self.cat.add(_model(3, "BIM 360://IR2/OLD_NO_FOLDER.rvt", source=CL.SOURCE_CACHE))

    def test_resolve_needs_the_same_folder(self):
        model, reason = self.cat.resolve("IR2", POD, "MBSIR2_WHP_ST_POD_CM_L06.rvt")
        self.assertIs(model, self.pod)
        self.assertEqual(reason, "")
        model, reason = self.cat.resolve("ir2", ("BIM HUB", "02_SHARED"),
                                         "mbsir2_whp_st_pod_cm_l06.RVT")
        self.assertEqual(model.model_guid, _guid(2))
        model, reason = self.cat.resolve("IR2", ("BIM HUB", "03_PUBLISHED"),
                                         "MBSIR2_WHP_ST_POD_CM_L06.rvt")
        self.assertIsNone(model)
        self.assertIn("another folder", reason)

    def test_resolve_folderless_cache_path_and_unknown(self):
        model, _ = self.cat.resolve("IR2", ("Any", "Folder"), "OLD_NO_FOLDER.rvt")
        self.assertEqual(model.model_guid, _guid(3))
        model, reason = self.cat.resolve("IR2", POD, "NEVER_SEEN.rvt")
        self.assertIsNone(model)
        self.assertEqual(reason, CL.REASON_UNKNOWN)

    def test_first_source_wins_and_later_ones_fill_gaps(self):
        cat = CL.CloudCatalog()
        cat.add(_model(5, "BIM 360://IR2/A.rvt", region="", source=CL.SOURCE_LINKED))
        merged = cat.add(_model(5, "Autodesk Docs://IR2/X/A.rvt", region="US",
                                source=CL.SOURCE_CACHE))
        self.assertEqual(len(cat), 1)
        self.assertEqual(merged.source, CL.SOURCE_LINKED)
        self.assertEqual(merged.folders, ("X",))
        self.assertEqual(merged.region, "US")
        # Rows without both GUIDs are never kept.
        self.assertIsNone(cat.add(CL.CloudModel("a.rvt", "IR2", (), "", _guid(6))))

    def test_region_for(self):
        self.assertEqual(self.cat.region_for(self.pod), "EMEA")       # the host's
        self.cat.add(_model(8, "Autodesk Docs://P2/a.rvt", project_guid=P_OTHER, region="US"))
        other = _model(9, "Autodesk Docs://P2/b.rvt", project_guid=P_OTHER)
        self.cat.add(other)
        self.assertEqual(self.cat.region_for(other), "US")             # its project's

    def test_in_folder(self):
        names = lambda models: sorted(m.model_guid for m in models)
        self.assertEqual(names(self.cat.in_folder("IR2", ("BIM HUB",), recursive=True)),
                         [_guid(1), _guid(2), _guid(99)])
        self.assertEqual(names(self.cat.in_folder("IR2", ("BIM HUB",), recursive=False)), [])
        # A model with no known folder only belongs to the project level.
        self.assertIn(_guid(3), names(self.cat.in_folder("IR2", ())))

    def test_candidates_statuses_and_dedupe(self):
        self.cat.local_link_names.add("mbsir2_jcw_fc_arn_cm.rvt")
        arena = _model(4, "Autodesk Docs://IR2/BIM HUB/01_WIP/MBSIR2_JCW_FC_ARN_CM.rvt")
        self.cat.add(arena)
        dc_files = [_dc(*(POD + ("MBSIR2_WHP_ST_POD_CM_L06.rvt",))),
                    _dc(*(POD + ("UNKNOWN.rvt",))),
                    r"D:\elsewhere\x.rvt"]
        rows = CL.cloud_candidates(self.cat, dc_files)
        by_name = dict((r.name, r) for r in rows)
        pod = [r for r in rows if r.model is self.pod]
        self.assertEqual(len(pod), 1)                                   # file + link = one row
        self.assertEqual(pod[0].status, CL.STATUS_LINKED)
        self.assertTrue(pod[0].local_path.endswith("MBSIR2_WHP_ST_POD_CM_L06.rvt"))
        self.assertEqual(by_name["HOST.rvt"].status, CL.STATUS_HOST)
        self.assertFalse(by_name["HOST.rvt"].selectable)
        self.assertEqual(by_name["MBSIR2_JCW_FC_ARN_CM.rvt"].status, CL.STATUS_AS_FILE)
        self.assertFalse(by_name["MBSIR2_JCW_FC_ARN_CM.rvt"].checked)
        self.assertEqual(by_name["UNKNOWN.rvt"].status, CL.STATUS_UNKNOWN)
        self.assertFalse(by_name["UNKNOWN.rvt"].selectable)
        self.assertEqual(by_name["x.rvt"].tip, CL.REASON_NOT_DC)
        ready = [r for r in rows if r.status == CL.STATUS_READY]
        self.assertTrue(ready and all(r.checked and r.selectable for r in ready))
        self.assertEqual(rows, sorted(rows, key=lambda r: (r.location.lower(), r.name.lower())))

    def test_candidates_narrowed_to_a_desktop_connector_folder(self):
        folder = CL.desktop_connector_parts(_dc("BIM HUB", "02_SHARED"))
        rows = CL.cloud_candidates(self.cat, [], folder, recursive=True)
        self.assertEqual([r.model.model_guid for r in rows], [_guid(2)])

    def test_match_link(self):
        model, _ = CL.match_link(self.cat, _dc(*(POD + ("MBSIR2_WHP_ST_POD_CM_L06.rvt",))))
        self.assertIs(model, self.pod)
        model, reason = CL.match_link(self.cat, "Autodesk Docs://IR2/a.rvt")
        self.assertIsNone(model)
        self.assertIn("Already", reason)
        # A plain disk copy: only a name that is unique in this project.
        model, _ = CL.match_link(self.cat, r"\\server\models\OLD_NO_FOLDER.rvt")
        self.assertEqual(model.model_guid, _guid(3))
        model, reason = CL.match_link(self.cat, r"D:\copy\MBSIR2_WHP_ST_POD_CM_L06.rvt")
        self.assertIsNone(model)
        self.assertIn("2 cloud models", reason)


class TestCloudDialogContract(unittest.TestCase):
    """Batch Link wiring of the Autodesk Docs source."""

    @classmethod
    def setUpClass(cls):
        cls.xaml = _read(os.path.join(LIB_DIR, 'GUI', 'Tools', 'BatchLink.xaml'))
        cls.dialog = _read(os.path.join(LIB_DIR, 'GUI', 'BatchLinkDialog.py'))

    def test_column_constants_match_the_grid(self):
        grid = self.xaml[self.xaml.index('x:Name="grid_files"'):]
        grid = grid[:grid.index('</DataGrid.Columns>')]
        headers = re.findall(r'<DataGrid\w*Column\s+Header="([^"]+)"', grid)
        self.assertEqual(headers, ["FILE NAME", "FOLDER", "SIZE", "DATE MODIFIED", "STATUS"])
        # +1: the checkbox column has its header as a child element.
        for const, header in (('COL_FOLDER', 'FOLDER'), ('COL_SIZE', 'SIZE'),
                              ('COL_DATE', 'DATE MODIFIED')):
            self.assertIn('{} = {}'.format(const, headers.index(header) + 1), self.dialog)

    def test_source_radios(self):
        for name, tag in (('rb_source_folder', 'folder'), ('rb_source_cloud', 'cloud')):
            radio = self.xaml[self.xaml.index('x:Name="{}"'.format(name)):]
            radio = radio[:radio.index('/>')]
            self.assertIn('Tag="{}"'.format(tag), radio)
            self.assertIn('Checked="source_mode_checked"', radio)
        self.assertIn('str(sender.Tag) == "cloud"', self.dialog)

    def test_cloud_links_use_a_cloud_model_path(self):
        body = self.dialog[self.dialog.index('def _link_models'):]
        body = body[:body.index('\n    def ')]
        self.assertIn('_cloud.cloud_model_path(', body)
        # cloud links are never relative (RevitLinkOptions(False) in _links.link_one)
        self.assertIn('relative=is_relative and not cloud', body)
        self.assertIn('_links.link_one(', body)
        self.assertIn('_cloud.is_cloud_doc(self.doc)', body)

    def test_list_rvt_files(self):
        import ast
        tree = ast.parse(self.dialog)
        tree.body = [n for n in tree.body
                     if (isinstance(n, ast.FunctionDef) and n.name == 'list_rvt_files')
                     or (isinstance(n, ast.Assign)
                         and getattr(n.targets[0], 'id', '') == 'BACKUP_PATTERN')]
        scope = {'os': os, 're': re}
        exec(compile(tree, 'BatchLinkDialog.py', 'exec'), scope)
        root = tempfile.mkdtemp(prefix="t3_list_rvt_")
        try:
            os.makedirs(os.path.join(root, "sub"))
            for rel in ("A.rvt", "A.0001.rvt", "~$B.rvt", "c.dwg", os.path.join("sub", "D.RVT")):
                open(os.path.join(root, rel), "w").close()
            names = lambda *a: sorted(os.path.basename(p) for p in scope['list_rvt_files'](*a))
            self.assertEqual(names(root, False, True), ["A.rvt"])
            self.assertEqual(names(root, True, True), ["A.rvt", "D.RVT"])
            self.assertEqual(names(root, False, False), ["A.0001.rvt", "A.rvt"])
        finally:
            shutil.rmtree(root, ignore_errors=True)



class _FakeGuid(object):
    def __init__(self, text):
        self.text = text

    def __str__(self):
        return self.text


class _FakeModelPath(object):
    def __init__(self, project_guid, model_guid, region="EMEA", cloud=True):
        self.CloudPath = cloud
        self.Region = region
        self._ids = (project_guid, model_guid)

    def GetProjectGUID(self):
        return _FakeGuid(self._ids[0])

    def GetModelGUID(self):
        return _FakeGuid(self._ids[1])


class _FakeInfo(dict):
    def ContainsKey(self, key):
        return key in self


class _FakeLinkType(object):
    """visible = InSessionPath; path = ExternalFileReference.GetPath() (None raises)."""

    def __init__(self, name, visible, path, info=None):
        self.Name = name
        self._ref = type('Ref', (), {'InSessionPath': visible,
                                     'GetReferenceInformation': lambda s: _FakeInfo(info or {})})()
        self._path = path

    def GetExternalResourceReference(self, _kind):
        return self._ref

    def GetExternalFileReference(self):
        if self._path is None:
            raise RuntimeError("no external file reference")
        return type('Ext', (), {'GetPath': lambda s, p=self._path: p})()


class _FakeRevitDoc(object):
    def __init__(self, title, cloud_path=None, visible="", links=(), link_docs=()):
        self.Title = title
        self.IsModelInCloud = cloud_path is not None
        self.IsFamilyDocument = False
        self.PathName = visible
        self._cloud_path = cloud_path
        self.elements = {'RevitLinkType': list(links),
                         'RevitLinkInstance': [type('Inst', (), {'GetLinkDocument': lambda s, d=d: d})()
                                               for d in link_docs]}

    def GetCloudModelPath(self):
        return self._cloud_path

    def Equals(self, other):
        return self is other


class TestCloudCollectors(unittest.TestCase):
    """collect_known_models against fake Revit objects: every source, merged once."""

    def setUp(self):
        import types
        db = types.ModuleType('Autodesk.Revit.DB')
        db.RevitLinkType = type('RevitLinkType', (), {})
        db.RevitLinkInstance = type('RevitLinkInstance', (), {})
        db.FilteredElementCollector = lambda doc: type('C', (), {
            'OfClass': lambda s, cls: doc.elements.get(cls.__name__, [])})()
        db.ExternalResourceTypes = type('ERT', (), {'BuiltInExternalResourceTypes':
                                                    type('B', (), {'RevitLink': 'RevitLink'})})
        db.ModelPathUtils = type('MPU', (), {
            'CloudRegionUS': 'US',
            'ConvertModelPathToUserVisiblePath': staticmethod(lambda mp: ""),
            'ConvertCloudGUIDsToCloudPath': staticmethod(lambda r, p, m: ('cloud', r, p, m))})
        self.central = {}
        central = self.central
        db.BasicFileInfo = type('BFI', (), {'Extract': staticmethod(
            lambda path: type('I', (), {'CentralPath': central.get(os.path.basename(path), ""),
                                        'Dispose': lambda s: None})())})
        links = types.ModuleType('Snippets._links')
        links.element_name = lambda e: e.Name
        system = types.ModuleType('System')
        system.Guid = type('Guid', (), {'Parse': staticmethod(lambda t: ('guid', t))})
        autodesk = types.ModuleType('Autodesk')
        revit = types.ModuleType('Autodesk.Revit')
        autodesk.Revit, revit.DB = revit, db
        self._saved = {}
        for name, mod in (('Autodesk', autodesk), ('Autodesk.Revit', revit),
                          ('Autodesk.Revit.DB', db), ('Snippets._links', links),
                          ('System', system)):
            self._saved[name] = sys.modules.get(name)
            sys.modules[name] = mod
        self._saved_env = os.environ.get('LOCALAPPDATA')
        self.local = tempfile.mkdtemp(prefix="t3_localappdata_")
        os.environ['LOCALAPPDATA'] = self.local

    def tearDown(self):
        for name, mod in self._saved.items():
            if mod is None:
                sys.modules.pop(name, None)
            else:
                sys.modules[name] = mod
        if self._saved_env is None:
            os.environ.pop('LOCALAPPDATA', None)
        else:
            os.environ['LOCALAPPDATA'] = self._saved_env
        shutil.rmtree(self.local, ignore_errors=True)

    def _cache_file(self, project_guid, model_guid, central):
        folder = os.path.join(self.local, "Autodesk", "Revit", "Autodesk Revit 2025",
                              "CollaborationCache", "ABC123XYZ", project_guid)
        os.makedirs(folder, exist_ok=True)
        open(os.path.join(folder, model_guid + ".rvt"), "w").close()
        self.central[model_guid + ".rvt"] = central

    def test_every_source_lands_in_one_catalog(self):
        wip = "Autodesk Docs://IR2/BIM HUB/01_WIP/"
        link_doc = _FakeRevitDoc("STR", links=[_FakeLinkType(
            "B.rvt", wip + "B.rvt", _FakeModelPath(P_GUID, _guid(2)))])
        host = _FakeRevitDoc(
            "HOST", _FakeModelPath(P_GUID, _guid(99), "EMEA"), wip + "HOST.rvt",
            links=[_FakeLinkType("A.rvt", wip + "A.rvt", _FakeModelPath(P_GUID, _guid(1))),
                   _FakeLinkType("MBSIR2_JCW_FC_ARN_CM.rvt", "",
                                 _FakeModelPath("", "", cloud=False)),
                   # Unloaded: no cloud ModelPath, the reference still has the IDs.
                   _FakeLinkType("U.rvt", "", None, {"LinkedModelProjectId": "b." + P_GUID,
                                                     "LinkedModelModelId": _guid(5)})],
            link_docs=[link_doc])
        other = _FakeRevitDoc("OPEN", _FakeModelPath(P_OTHER, _guid(3), "US"),
                              "Autodesk Docs://P2/OPEN.rvt",
                              links=[_FakeLinkType("C.rvt", "Autodesk Docs://P2/C.rvt",
                                                   _FakeModelPath(P_OTHER, _guid(4), "US"))])
        app = type('App', (), {'VersionNumber': '2025', 'Documents': [host, other, link_doc]})()
        for doc in (host, other, link_doc):
            doc.Application = app
        self._cache_file(P_GUID, _guid(6), "Autodesk Docs://IR2/BIM HUB/02_SHARED/S.rvt")
        self._cache_file(P_GUID, _guid(1), "Autodesk Docs://IR2/BIM HUB/01_WIP/A.rvt")
        self._cache_file(P_OTHER, _guid(7), "Autodesk Docs://P2/NOT_READ.rvt")

        steps = []
        cat = CL.collect_known_models(host, progress=lambda i, n: steps.append((i, n)))
        self.assertEqual(sorted(m.model_guid for m in cat.models()),
                         sorted(_guid(n) for n in (1, 2, 3, 4, 5, 6, 99)))
        self.assertEqual((cat.host_guid, cat.host_project_guid, cat.host_region),
                         (_guid(99), P_GUID, "EMEA"))
        self.assertEqual(cat.linked_guids, {_guid(1), _guid(5)})
        self.assertEqual(cat.local_link_names, {"mbsir2_jcw_fc_arn_cm.rvt"})
        self.assertEqual(cat.get(_guid(1)).source, CL.SOURCE_LINKED)     # link beats cache
        self.assertEqual(cat.get(_guid(2)).source, CL.SOURCE_IN_LINK)
        self.assertEqual(cat.get(_guid(4)).source, CL.SOURCE_OPEN)
        self.assertEqual(cat.get(_guid(6)).display_path,
                         "Autodesk Docs://IR2/BIM HUB/02_SHARED/S.rvt")
        self.assertEqual(len(steps), 2)          # only this model's project is read
        unloaded = cat.get(_guid(5))             # named by its type, project by its GUID
        self.assertEqual((unloaded.name, unloaded.project, unloaded.folders), ("U.rvt", "IR2", ()))

        # A model read from a link keeps its own ModelPath; a cached one is built.
        self.assertIs(CL.cloud_model_path(cat.get(_guid(1))), cat.get(_guid(1)).model_path)
        cached = cat.get(_guid(6))
        self.assertEqual(CL.cloud_model_path(cached, cat.region_for(cached)),
                         ('cloud', 'EMEA', ('guid', P_GUID), ('guid', _guid(6))))

    def test_disk_host_has_no_cloud_catalog(self):
        host = _FakeRevitDoc("LOCAL", None, r"D:\a.rvt")
        host.Application = type('App', (), {'VersionNumber': '2025', 'Documents': [host]})()
        cat = CL.collect_known_models(host)
        self.assertFalse(CL.is_cloud_doc(host))
        self.assertEqual((len(cat), cat.host_guid), (0, ""))

if __name__ == '__main__':
    unittest.main()