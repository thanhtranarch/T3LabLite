"""BatchOut selection: the tally describes the entire batch, including filtered
rows, the row checkboxes show that selection, the Views-mode type filter is
built from the loaded views, and a PDF24 row ends in a queue state.
Ported from t3lab_dev dev/test_batchout_selection.py.

Run:  python3 lite_guard/test_batchout_selection.py      Exit 0 = all pass.
"""
import ast
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock


class SelectionTallyTests(unittest.TestCase):
    def setUp(self):
        source = Path(__file__).resolve().parents[1] / 'lib/GUI/BatchOutDialog.py'
        tree = ast.parse(source.read_text(encoding='utf-8-sig'))
        cls = next(node for node in tree.body if isinstance(node, ast.ClassDef)
                   and node.name == 'ExportManagerWindow')
        method = next(node for node in cls.body if isinstance(node, ast.FunctionDef)
                      and node.name == 'update_selection_count')
        scope = {'logger': Mock()}
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(source), 'exec'), scope)
        self.update = scope['update_selection_count']

    def test_hidden_selection_is_reported_for_sheets_and_views(self):
        for mode in ('sheets', 'views'):
            with self.subTest(mode=mode):
                rows = [SimpleNamespace(IsSelected=True), SimpleNamespace(IsSelected=True)]
                window = SimpleNamespace(selection_mode=mode,
                                         selection_count_text=SimpleNamespace(Text=''),
                                         header_checkbox=SimpleNamespace(IsChecked=False))
                setattr(window, 'all_' + mode, rows)
                setattr(window, 'filtered_' + mode, rows[:1])
                self.update(window)
                self.assertIn('2 {} selected'.format(mode), window.selection_count_text.Text)
                self.assertIn('1 selected hidden', window.selection_count_text.Text)
                self.assertTrue(window.header_checkbox.IsChecked)

    def test_empty_filter_does_not_hide_selected_batch_count(self):
        window = SimpleNamespace(selection_mode='sheets',
                                 all_sheets=[SimpleNamespace(IsSelected=True)], filtered_sheets=[],
                                 selection_count_text=SimpleNamespace(Text=''),
                                 header_checkbox=SimpleNamespace(IsChecked=True))
        self.update(window)
        self.assertIn('1 sheets selected | 0 shown', window.selection_count_text.Text)
        self.assertFalse(window.header_checkbox.IsChecked)


class RowCheckboxShowsModelTests(unittest.TestCase):
    """pythonnet hands WPF a PyObject for a Python attribute: WPF converts it
    to string but not to bool?, so IsChecked bound straight to IsSelected read
    unchecked and every tick vanished on Items.Refresh() (2026-09-26)."""

    X = '{http://schemas.microsoft.com/winfx/2006/xaml}'
    P = '{http://schemas.microsoft.com/winfx/2006/xaml/presentation}'

    def setUp(self):
        xaml = Path(__file__).resolve().parents[1] / 'lib/GUI/Tools/ExportManager.xaml'
        root = ET.parse(xaml).getroot()
        listview = next(n for n in root.iter() if n.get(self.X + 'Name') == 'sheets_listview')
        self.row_boxes = [n for n in listview.iter(self.P + 'CheckBox')
                          if n.get('Click') == 'row_checkbox_clicked']
        self.bridges = {n.get(self.X + 'Name'): n for n in listview.iter(self.P + 'TextBlock')
                        if n.get(self.X + 'Name')}

    def test_row_checkbox_reads_selection_through_a_string(self):
        self.assertEqual(len(self.row_boxes), 1)
        binding = self.row_boxes[0].get('IsChecked')
        self.assertNotIn('Binding IsSelected', binding)
        self.assertIn('ElementName=row_selected_text', binding)
        self.assertIn('Mode=OneWay', binding)
        self.assertEqual(self.bridges['row_selected_text'].get('Text'), '{Binding IsSelected}')
        self.assertEqual(self.bridges['row_selected_text'].get('Visibility'), 'Collapsed')

    def test_click_handler_writes_the_model(self):
        source = Path(__file__).resolve().parents[1] / 'lib/GUI/BatchOutDialog.py'
        tree = ast.parse(source.read_text(encoding='utf-8-sig'))
        cls = next(node for node in tree.body if isinstance(node, ast.ClassDef)
                   and node.name == 'ExportManagerWindow')
        method = next(node for node in cls.body if isinstance(node, ast.FunctionDef)
                      and node.name == 'row_checkbox_clicked')
        scope = {'logger': Mock()}
        exec(compile(ast.Module(body=[method], type_ignores=[]), str(source), 'exec'), scope)
        row = SimpleNamespace(IsSelected=False)
        window = SimpleNamespace(sheets_listview=SimpleNamespace(Items=Mock()),
                                 update_selection_count=Mock(),
                                 update_export_preview_if_needed=Mock())
        scope['row_checkbox_clicked'](window, SimpleNamespace(DataContext=row, IsChecked=True), None)
        self.assertTrue(row.IsSelected)
        window.sheets_listview.Items.Refresh.assert_called_once()
        window.update_selection_count.assert_called_once()


def _window_methods(*names):
    """ExportManagerWindow methods, compiled alone (no Revit, no WPF)."""
    source = Path(__file__).resolve().parents[1] / 'lib/GUI/BatchOutDialog.py'
    tree = ast.parse(source.read_text(encoding='utf-8-sig'))
    cls = next(node for node in tree.body if isinstance(node, ast.ClassDef)
               and node.name == 'ExportManagerWindow')
    methods = [node for node in cls.body if isinstance(node, ast.FunctionDef)
               and node.name in names]
    scope = {'logger': Mock()}
    exec(compile(ast.Module(body=methods, type_ignores=[]), str(source), 'exec'), scope)
    return scope


class _Combo(object):
    """Stand-in for the view_type_filter ComboBox."""

    def __init__(self, labels, index=0):
        self.Items = []
        for label in labels:
            self.Items.append(SimpleNamespace(Content=label))
        self.SelectedIndex = index

    @property
    def SelectedItem(self):
        if 0 <= self.SelectedIndex < len(self.Items):
            return self.Items[self.SelectedIndex]
        return None


class ViewTypeFilterTests(unittest.TestCase):
    """Views mode showed 0 views (2026-10-10): the dropdown's first entry is
    "All Types" but apply_filters skipped the filter only for "All Views", and
    the other entries were enum names ("FloorPlan") while rows carry labels
    ("Floor Plan") — any use of the dropdown hid every view."""

    def setUp(self):
        self.views = [SimpleNamespace(ViewName='L1', ViewType='Floor Plan'),
                      SimpleNamespace(ViewName='S1', ViewType='Section'),
                      SimpleNamespace(ViewName='3D', ViewType='3D View')]

    def _filter(self, combo, search=''):
        apply_filters = _window_methods('apply_filters')['apply_filters']
        window = SimpleNamespace(selection_mode='views', all_views=self.views,
                                 search_textbox=SimpleNamespace(Text=search),
                                 view_type_filter=combo,
                                 update_items_list=Mock(),
                                 status_text=SimpleNamespace(Text=''))
        apply_filters(window)
        return window

    def test_all_types_shows_every_view(self):
        window = self._filter(_Combo(['All Types', 'Floor Plan', 'Section'], 0))
        self.assertEqual(len(window.filtered_views), 3)
        self.assertEqual(window.status_text.Text, 'Found 3 views')

    def test_a_type_shows_only_its_views(self):
        window = self._filter(_Combo(['All Types', 'Floor Plan', 'Section'], 2))
        self.assertEqual([v.ViewName for v in window.filtered_views], ['S1'])

    def test_dropdown_is_built_from_the_row_labels(self):
        controls = type(sys)('System.Windows.Controls')
        controls.ComboBoxItem = lambda: SimpleNamespace(Content=None)
        saved = {k: sys.modules.get(k) for k in ('System', 'System.Windows',
                                                 'System.Windows.Controls')}
        sys.modules['System'] = type(sys)('System')
        sys.modules['System.Windows'] = type(sys)('System.Windows')
        sys.modules['System.Windows.Controls'] = controls
        try:
            fill = _window_methods('_fill_view_type_filter')['_fill_view_type_filter']
            combo = _Combo(['All Types', 'FloorPlan', 'Section'], 2)     # "Section" picked
            combo.Items = _Items(combo.Items)
            window = SimpleNamespace(view_type_filter=combo, all_views=self.views)
            fill(window)
        finally:
            for k, v in saved.items():
                if v is None:
                    sys.modules.pop(k, None)
                else:
                    sys.modules[k] = v
        self.assertEqual([i.Content for i in combo.Items],
                         ['All Types', '3D View', 'Floor Plan', 'Section'])
        self.assertEqual(combo.SelectedItem.Content, 'Section')     # pick kept
        self.assertFalse(window._filling_view_types)
        # Every entry but "All Types" now matches at least one row.
        self.assertEqual({i.Content for i in list(combo.Items)[1:]},
                         {v.ViewType for v in self.views})

    def test_xaml_ships_only_the_all_types_entry(self):
        X = '{http://schemas.microsoft.com/winfx/2006/xaml}'
        P = '{http://schemas.microsoft.com/winfx/2006/xaml/presentation}'
        xaml = Path(__file__).resolve().parents[1] / 'lib/GUI/Tools/ExportManager.xaml'
        combo = next(n for n in ET.parse(xaml).getroot().iter()
                     if n.get(X + 'Name') == 'view_type_filter')
        self.assertEqual([i.get('Content') for i in combo.iter(P + 'ComboBoxItem')], ['All Types'])


class Pdf24QueueStateTests(unittest.TestCase):
    """The PDF24 row passed "Printing to PDF24...", which is no QUEUE_* state,
    so the row never closed out when a run stopped (t3lab_dev 16626d5)."""

    def test_pdf24_progress_uses_a_queue_state(self):
        source = Path(__file__).resolve().parents[1] / 'lib/GUI/BatchOutDialog.py'
        text = source.read_text(encoding='utf-8-sig')
        self.assertNotIn("'PDF', 50, 'Printing to PDF24...')", text)
        self.assertNotIn('"PDF", 50, "Printing to PDF24...")', text)
        self.assertIn('QUEUE_EXPORTING)', text)


class _Items(list):
    """ComboBox.Items: a list with Clear() / Add()."""

    def Clear(self):
        del self[:]

    def Add(self, item):
        self.append(item)


if __name__ == '__main__':
    unittest.main()
