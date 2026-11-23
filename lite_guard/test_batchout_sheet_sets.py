# -*- coding: utf-8 -*-
"""Save V/S Set in BatchOut (2026-11-20).

Two Revit rules the old code broke:
  - `PrintManager.ViewSheetSetting` only exists while `PrintRange` is `Select`;
    reading it otherwise throws "This property is only available when user
    choose Select of Print Range".
  - `SaveAs` persists what the current set holds at that moment, so the views
    must be assigned first.

Runs ViewSheetSetsService against fakes of the Revit API that enforce both.
Run: python lite_guard/test_batchout_sheet_sets.py
"""
import ast
import io
import os
import sys
import types
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'T3Lab.extension', 'lib')
if not os.path.isdir(LIB):
    LIB = os.path.join(ROOT, 'lib')          # T3LabLite ships lib/ at the root
DIALOG = os.path.join(LIB, 'GUI', 'BatchOutDialog.py')


# ── Revit API fakes ─────────────────────────────────────────────────────────
class PrintRange(object):
    Select = 'Select'
    Current = 'Current'
    Visible = 'Visible'


class ElementId(object):
    def __init__(self, value):
        self.Value = value


class View(object):
    def __init__(self, id_value, name='View', printable=True, template=False):
        self.Id = ElementId(id_value)
        self.Name = name
        self.CanBePrinted = printable
        self.IsTemplate = template
        self.IsValidObject = True


class ViewSheet(View):
    def __init__(self, id_value, name='Sheet', printable=True, placeholder=False):
        View.__init__(self, id_value, name, printable)
        self.IsPlaceholder = placeholder


class ViewSet(object):
    def __init__(self):
        self._views = []

    def Insert(self, view):
        self._views.append(view)
        return True

    def __iter__(self):
        return iter(self._views)

    def __len__(self):
        return len(self._views)


class ViewSheetSet(object):
    def __init__(self, name, id_value, doc=None):
        self.Name = name
        self.Id = ElementId(id_value)
        self.Views = ViewSet()       # what the setting holds right now
        self.persisted = []          # names Revit has actually stored
        self.IsValidObject = True
        self.Document = doc


class FilteredElementCollector(object):
    def __init__(self, doc):
        self.doc = doc

    def OfClass(self, _cls):
        return self

    def ToElements(self):
        return list(self.doc.sets)

    def __iter__(self):
        return iter(self.doc.sets)


class FakeViewSheetSetting(object):
    """SaveAs stores the current set as it stands; a later edit needs Save()."""

    def __init__(self, doc):
        self.doc = doc
        self.InSession = ViewSheetSet('In-Session View/Sheet Set', 0, doc)
        self.CurrentViewSheetSet = self.InSession

    def SaveAs(self, name):
        saved = ViewSheetSet(name, len(self.doc.sets) + 1, self.doc)
        saved.Views = self.CurrentViewSheetSet.Views
        saved.persisted = [view.Name for view in saved.Views]
        self.doc.sets.append(saved)
        self.CurrentViewSheetSet = saved
        return True

    def Save(self):
        current = self.CurrentViewSheetSet
        current.persisted = [view.Name for view in current.Views]
        return True


class FakePrintManager(object):
    def __init__(self, doc):
        self.PrintRange = PrintRange.Current       # what Revit opens a document on
        self._setting = FakeViewSheetSetting(doc)
        self.ranges_seen = []

    def __setattr__(self, name, value):
        if name == 'PrintRange' and 'ranges_seen' in self.__dict__:
            self.ranges_seen.append(value)
        object.__setattr__(self, name, value)

    @property
    def ViewSheetSetting(self):
        if self.PrintRange != PrintRange.Select:
            raise RuntimeError('This property is only available when user '
                               'choose Select of Print Range.')
        return self._setting


class FakeDoc(object):
    def __init__(self):
        self.sets = []
        self.IsModifiable = True
        self.PrintManager = FakePrintManager(self)

    def GetElement(self, element_id):
        return next((s for s in self.sets if s.Id.Value == element_id.Value), None)


def _install_fakes():
    """Put the Revit API and Snippets._compat in sys.modules before the import."""
    db = types.ModuleType('Autodesk.Revit.DB')
    for name, value in (('PrintRange', PrintRange), ('ElementId', ElementId),
                        ('View', View), ('ViewSheet', ViewSheet),
                        ('ViewSet', ViewSet), ('ViewSheetSet', ViewSheetSet),
                        ('FilteredElementCollector', FilteredElementCollector)):
        setattr(db, name, value)
    revit = types.ModuleType('Autodesk.Revit')
    revit.DB = db
    autodesk = types.ModuleType('Autodesk')
    autodesk.Revit = revit

    class _disposing(object):
        def __init__(self, obj):
            self.obj = obj

        def __enter__(self):
            return self.obj

        def __exit__(self, *_exc):
            return False

    compat = types.ModuleType('Snippets._compat')
    compat.disposing = _disposing
    compat.fold = lambda text: text.casefold()
    snippets = types.ModuleType('Snippets')
    snippets._compat = compat
    snippets.__path__ = []

    sys.modules.update({'Autodesk': autodesk, 'Autodesk.Revit': revit,
                        'Autodesk.Revit.DB': db,
                        'Snippets': snippets, 'Snippets._compat': compat})


_install_fakes()
if LIB not in sys.path:
    sys.path.insert(0, LIB)
from Services.SheetManager.viewsheet_sets_service import ViewSheetSetsService  # noqa: E402


class SaveSetTests(unittest.TestCase):
    def setUp(self):
        self.doc = FakeDoc()
        self.service = ViewSheetSetsService(self.doc)
        self.sheets = [ViewSheet(11, 'A101'), ViewSheet(12, 'A102')]
        self.views = [View(21, 'Level 1'), View(22, '3D')]

    def saved(self, name):
        return next((s for s in self.doc.sets if s.Name == name), None)

    def test_print_range_is_switched_to_select_and_put_back(self):
        """The old code read ViewSheetSetting off the manager as it stood."""
        self.service.create_set_from_views('Set A', self.sheets)
        self.assertEqual(self.doc.PrintManager.ranges_seen,
                         [PrintRange.Select, PrintRange.Current])
        self.assertEqual(self.doc.PrintManager.PrintRange, PrintRange.Current)

    def test_views_are_assigned_before_saveas(self):
        """SaveAs first and Views after stored an empty set."""
        self.service.create_set_from_views('Set A', self.sheets)
        saved = self.saved('Set A')
        self.assertIsNotNone(saved)
        self.assertEqual(saved.persisted, ['A101', 'A102'])

    def test_views_as_well_as_sheets_can_be_saved(self):
        self.service.create_set_from_views('Mixed', self.sheets + self.views)
        self.assertEqual(self.saved('Mixed').persisted,
                         ['A101', 'A102', 'Level 1', '3D'])

    def test_duplicates_are_dropped(self):
        self.service.create_set_from_views('Set A', self.sheets + [self.sheets[0]])
        self.assertEqual(len(self.saved('Set A').persisted), 2)

    def test_existing_name_is_refused_before_revit_sees_it(self):
        self.service.create_set_from_views('Set A', self.sheets)
        with self.assertRaises(ValueError) as caught:
            self.service.create_set_from_views('set a', self.views)
        self.assertIn('already exists', str(caught.exception))
        self.assertEqual(self.doc.PrintManager.PrintRange, PrintRange.Current)

    def test_empty_name_is_refused(self):
        with self.assertRaises(ValueError):
            self.service.create_set_from_views('   ', self.sheets)

    def test_empty_selection_is_refused(self):
        with self.assertRaises(ValueError):
            self.service.create_set_from_views('Set A', [])

    def test_unprintable_items_are_refused(self):
        for bad in (View(31, 'Template', template=True),
                    View(32, 'Hidden', printable=False),
                    ViewSheet(33, 'Placeholder', placeholder=True),
                    object()):
            with self.assertRaises(ValueError):
                self.service.create_set_from_views('Set A', [bad])

    def test_a_transaction_is_required(self):
        self.doc.IsModifiable = False
        with self.assertRaises(RuntimeError):
            self.service.create_set_from_views('Set A', self.sheets)

    def test_the_sheet_only_entry_point_still_works(self):
        self.doc.sets.append(ViewSheetSet('placeholder', 99, self.doc))
        self.doc.sets.pop()
        for sheet in self.sheets:
            setattr(self.doc, 'sheet_%s' % sheet.Id.Value, sheet)
        self.doc.GetElement = lambda eid: next(
            (s for s in self.sheets if s.Id.Value == eid.Value),
            next((s for s in self.doc.sets if s.Id.Value == eid.Value), None))
        self.service.create_sheet_set('Set B', [s.Id for s in self.sheets])
        self.assertEqual(self.saved('Set B').persisted, ['A101', 'A102'])


class DialogKeepsOffPrintManagerTests(unittest.TestCase):
    """The dialog must go through the service, not PrintManager."""

    def setUp(self):
        with io.open(DIALOG, encoding='utf-8-sig') as fh:
            self.source = fh.read()
        tree = ast.parse(self.source)
        cls = next(n for n in tree.body
                   if isinstance(n, ast.ClassDef) and n.name == 'ExportManagerWindow')
        method = next(n for n in cls.body if isinstance(n, ast.FunctionDef)
                      and n.name == 'save_current_selection_as_vs_set')
        self.method_source = ast.unparse(method)   # code only; comments explain the bug

    def test_the_save_handler_never_touches_viewsheetsetting(self):
        self.assertNotIn('ViewSheetSetting', self.method_source)
        self.assertNotIn('PrintManager', self.method_source)

    def test_the_save_handler_calls_the_service(self):
        self.assertIn('create_set_from_views', self.method_source)

    def test_the_transaction_is_disposed(self):
        self.assertIn('disposing(Transaction(', self.method_source)


if __name__ == '__main__':
    unittest.main(verbosity=2)
