# -*- coding: utf-8 -*-
"""
Schedule Manager Dialog — WPF host for ManaSched.xaml.

Three pages, all things Revit cannot do on its own:

  New Table      (TableGen-style) the workbook — a local .xlsx or a Google
                 Sheet (pasted link / .gsheet shortcut, Snippets/gsheet_source.py)
                 — opens in a sheet viewer that
                 looks like Excel (fills, bold, colours, column widths, sheet
                 tabs); drag over the cells — or pick a named range / Excel
                 table — and create a schedule, drafting view or legend with
                 the Excel formatting, linked or as a static copy.
  Linked Tables  which tables are out of date; update / reload / unlink.
                                             → Snippets/excel_table_revit.py
  Schedule Data  export schedule values to .xlsx (Element ID + one column per
                 field), edit in Excel, import, review every change, write the
                 ticked ones back in one undo step.   → Snippets/schedule_data.py

Removed on purpose (Revit already does it): schedule duplication / view
template assignment (Duplicate View, View Templates) and the "as displayed"
text export (File > Export > Reports > Schedule).

This module is UI only; the Revit logic lives in the two Snippets above and
the Excel reader in Snippets/xlsx_table.py.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__ = "Tran Tien Thanh"
__title__ = "Schedule Manager Dialog"

import os

import clr
clr.AddReference('System')
clr.AddReference('System.Data')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Data import DataTable
from System.Windows import Visibility
from System.Windows.Controls import (ComboBoxItem, DataGridCellInfo, DataGridLength,
                                     DataGridTemplateColumn, RadioButton)
from System.Windows.Markup import XamlReader

from GUI.WPF_Base import T3WPFWindow
from GUI import T3Dialog
from Snippets._compat import eid_value, make_eid, short_error
from Snippets import xlsx_table as xt
from Snippets import excel_table_revit as etr
from Snippets import gsheet_source as gs
from Snippets import schedule_data as sd

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'ManaSched.xaml')

RANGE_SELECTION = "Selected cells"
VIEW_MAX_ROWS = 400                 # sheet viewer limits (the import itself has none)
VIEW_MAX_COLS = 50
HIGHLIGHT_MAX_CELLS = 5000          # beyond this a named range is not drawn as a selection
KIND_ORDER = [etr.KIND_SCHEDULE, etr.KIND_DRAFTING, etr.KIND_LEGEND]
KIND_HINT = {
    etr.KIND_SCHEDULE: "Header-only schedule: place it on as many sheets as you like. "
                       "Fonts, fills, borders, merges and sizes are set cell by cell.",
    etr.KIND_DRAFTING: "Closest copy of the Excel look (text notes, lines, solid fills). "
                       "A drafting view goes on one sheet.",
    etr.KIND_LEGEND: "Same drawing as a drafting view, placeable on many sheets. "
                     "Needs one legend view in the project to copy.",
}
TABLES_HINT = ("Tick tables and press Update Tables. With nothing ticked, every table "
               "whose Excel file or Google Sheet changed is updated.")
EXCEL_FILTER = ("Excel workbook or Google Drive shortcut (*.xlsx;*.xlsm;*.gsheet)|"
                "*.xlsx;*.xlsm;*.gsheet|Excel workbook (*.xlsx;*.xlsm)|*.xlsx;*.xlsm")


# ============================================================
# ROW MODELS (plain Python, read through string bridges)
# ============================================================

class TableRow(object):
    def __init__(self, view, info):
        self.view_id = eid_value(view.Id)
        self.info = info
        self.name = view.Name
        self.kind_text = etr.KIND_LABEL.get(info.get('kind'), info.get('kind', ''))
        rng = info.get('range') or xt.USED_RANGE
        self.source = u"{} · {} · {}".format(etr.source_label(info),
                                            info.get('sheet') or '?', rng)
        self.updated = info.get('updated') or ''
        self.StatusText = "Checking"
        self.Severity = "Warning"
        self.detail = ''
        self.is_checked = False


class ScheduleRow(object):
    def __init__(self, name, schedule):
        self.name = name
        self.sched_id = eid_value(schedule.Id)
        try:
            self.field_count = str(schedule.Definition.GetFieldCount())
        except Exception:
            self.field_count = "—"
        self.is_checked = False


class ChangeRow(object):
    def __init__(self, change):
        self.change = change
        self.schedule = change['schedule']
        self.element_id = str(change['element_id'])
        self.field = change['field']
        self.old = change['old']
        self.new = change['new']
        self.is_checked = True


def _visible(flag):
    return Visibility.Visible if flag else Visibility.Collapsed


def _combo_fill(combo, labels, select=0):
    combo.Items.Clear()
    for label in labels:
        item = ComboBoxItem()
        item.Content = label
        combo.Items.Add(item)
    if labels:
        combo.SelectedIndex = max(0, min(select, len(labels) - 1))


def _combo_text(combo):
    item = combo.SelectedItem
    if item is None:
        return ''
    return str(getattr(item, 'Content', item) or '')


# ============================================================
# MAIN WINDOW
# ============================================================

class ManaSchedWindow(T3WPFWindow):
    """Excel Tables page + Schedule Data page. Chrome is wired by T3WPFWindow."""

    def __init__(self, script_dir, revit_obj):
        T3WPFWindow.__init__(self, _XAML)
        self._script_dir = script_dir
        self._uidoc = revit_obj.ActiveUIDocument
        self._doc = self._uidoc.Document

        self._table_rows = []
        self._sched_rows = []
        self._schedules = []            # [(name, ViewSchedule)]
        self._change_rows = []
        self._xl_path = None            # workbook open on the New Table page
        self._gsheet = None             # its Google Sheets link (path = cache copy)
        self._source_title = ''         # file stem, or "Google Sheet"
        self._sheet = None              # its active sheet
        self._view = None               # xt.Table of the sheet viewer
        self._choices = []              # range names of the active sheet
        self._syncing = False           # True while code changes the selection
        self._form_cache = None         # (key, Table) of the last read range
        self._range_bounds = None       # (r1, c1, r2, c2) shown in the summary
        self._keep_range_once = False   # Re-read keeps the range of the same sheet

        # navigation + primary action
        self.nav_new.Checked += self._on_nav
        self.nav_tables.Checked += self._on_nav
        self.nav_data.Checked += self._on_nav
        self.btn_primary.Click += self._on_primary

        # New Table page
        self.btn_browse.Click += self._on_browse
        self.btn_open_link.Click += lambda s, e: self._open_source(self.txt_file.Text)
        self.txt_file.KeyDown += self._on_file_key
        self.btn_reread.Click += self._on_reread
        self.dg_sheet.SelectedCellsChanged += self._on_cells_selected
        self.cmb_range.SelectionChanged += self._on_range_changed
        self.txt_range.LostFocus += self._on_range_typed
        self.txt_range.KeyDown += self._on_range_key
        self.chk_hidden.Click += lambda s, e: self._show_sheet(keep_range=True)
        self.cmb_kind.SelectionChanged += self._on_kind_changed
        self.txt_name.TextChanged += lambda s, e: self._update_summary()
        self.chk_keep_link.Click += lambda s, e: self._update_summary()

        # Linked Tables page
        self.tbl_search_box.TextChanged += lambda s, e: self._fill_tables()
        self.btn_tbl_reload.Click += self._on_tbl_reload
        self.btn_tbl_unlink.Click += self._on_tbl_unlink

        # Schedule Data page
        self.data_search_box.TextChanged += lambda s, e: self._fill_schedules()
        self.btn_import.Click += self._on_import
        self.btn_apply.Click += self._on_apply

        self.enable_bulk_tick(self.dg_tables, "is_checked", on_change=self._count_tables)
        self.enable_bulk_tick(self.dg_schedules, "is_checked", on_change=self._count_schedules)
        self.enable_bulk_tick(self.dg_changes, "is_checked", on_change=self._count_changes)

        _combo_fill(self.cmb_kind, [etr.KIND_LABEL[k] for k in KIND_ORDER], 1)
        self._load_tables()
        self._load_schedules()
        self._show_changes([], ["No file imported."])
        self.tab_main.SelectedItem = self.tab_new
        self._sync_primary()

    # ------------------------------------------------------------------
    # SHARED
    # ------------------------------------------------------------------

    def _status(self, text):
        self.txt_statusbar.Text = text

    def _page(self):
        if self.nav_tables.IsChecked:
            return 'tables'
        if self.nav_data.IsChecked:
            return 'data'
        return 'new'

    def _on_nav(self, sender, e):
        self.tab_main.SelectedItem = {'new': self.tab_new, 'tables': self.tab_tables,
                                      'data': self.tab_data}[self._page()]
        self._sync_primary()

    def _sync_primary(self):
        text, glyph = {'new': ("Create Table", u"\uE710"),
                       'tables': ("Update Tables", u"\uE72C"),
                       'data': ("Export to Excel", u"\uE896")}[self._page()]
        self.txt_primary.Text = text
        self.ico_primary.Text = glyph

    def _on_primary(self, sender, e):
        page = self._page()
        if page == 'new':
            self._create_table()
        elif page == 'tables':
            self._update_tables()
        else:
            self._export()

    def _pick_file(self, title):
        from Microsoft.Win32 import OpenFileDialog
        dlg = OpenFileDialog()
        dlg.Filter = EXCEL_FILTER
        dlg.Title = title
        folder = os.path.dirname(self._doc.PathName or '')
        if folder and os.path.isdir(folder):
            dlg.InitialDirectory = folder
        return dlg.FileName if dlg.ShowDialog(self) == True else None   # noqa: E712

    def _save_file(self, default_name):
        from Microsoft.Win32 import SaveFileDialog
        dlg = SaveFileDialog()
        dlg.Filter = "Excel workbook (*.xlsx)|*.xlsx"
        dlg.DefaultExt = "xlsx"
        dlg.FileName = default_name
        folder = os.path.dirname(self._doc.PathName or '')
        if folder and os.path.isdir(folder):
            dlg.InitialDirectory = folder
        return dlg.FileName if dlg.ShowDialog(self) == True else None   # noqa: E712

    # ── header select-all (rule 23) ───────────────────────────────────
    def select_all_dg_tables_clicked(self, sender, e):
        self.toggle_all_rows(self.dg_tables, "is_checked", sender.IsChecked)
        self._count_tables()

    def select_all_dg_schedules_clicked(self, sender, e):
        self.toggle_all_rows(self.dg_schedules, "is_checked", sender.IsChecked)
        self._count_schedules()

    def select_all_dg_changes_clicked(self, sender, e):
        self.toggle_all_rows(self.dg_changes, "is_checked", sender.IsChecked)
        self._count_changes()

    # ------------------------------------------------------------------
    # EXCEL TABLES — linked list
    # ------------------------------------------------------------------

    def _load_tables(self):
        self._table_rows = []
        for view, info in etr.find_linked_views(self._doc):
            row = TableRow(view, info)
            row.StatusText, row.Severity, row.detail = etr.check_status(self._doc, info)
            self._table_rows.append(row)
        self.tbl_note.Text = etr.storage_note() or TABLES_HINT
        self._fill_tables()

    def _fill_tables(self):
        q = (self.tbl_search_box.Text or '').strip().lower()
        rows = [r for r in self._table_rows
                if not q or q in r.name.lower() or q in r.source.lower()]
        self.set_items_source(self.dg_tables, rows)
        self.tbl_empty.Visibility = _visible(not rows)
        if self._table_rows and not rows:
            self.tbl_empty.Text = "No table matches the search."
        else:
            self.tbl_empty.Text = ("No Excel tables in this project yet. "
                                   "Create one on the New Table page.")
        self._count_tables()

    def _count_tables(self):
        total = len(self._table_rows)
        changed = sum(1 for r in self._table_rows if r.Severity != "Success")
        ticked = sum(1 for r in self._table_rows if r.is_checked)
        parts = ["{} table(s)".format(total)]
        if changed:
            parts.append("{} need attention".format(changed))
        if ticked:
            parts.append("{} ticked".format(ticked))
        self.tbl_count_text.Text = " · ".join(parts)
        self.sync_header_checkbox(self.chk_all_dg_tables, self.dg_tables, "is_checked")

    def _ticked_tables(self):
        return [r for r in self._table_rows if r.is_checked]

    def _view_of(self, row):
        return self._doc.GetElement(make_eid(row.view_id))

    def _update_tables(self):
        rows = self._ticked_tables()
        if not rows:
            rows = [r for r in self._table_rows
                    if r.StatusText in ("Excel changed", "Sheet changed")]
            if not rows:
                T3Dialog.show_info("Tick the tables to update. Every linked table is "
                                   "up to date with its Excel file.", "Nothing to Update",
                                   owner=self)
                return
            question = "No table is ticked. Update all {} table(s) whose Excel file changed?".format(len(rows))
        else:
            question = "Update {} table(s) from Excel?".format(len(rows))
        details = ("Drafting views and legends are redrawn in place. Schedules are rebuilt "
                   "and put back on the same sheets at the same points.\n\n"
                   + "\n".join(r.name for r in rows[:20]))
        if not T3Dialog.confirm(question, "Update Excel Tables", ok_text="Update",
                                details=details, owner=self):
            return
        self._run_updates(rows)

    def _run_updates(self, rows, path_override=None):
        done, problems = 0, []
        for row in rows:
            view = self._view_of(row)
            if view is None:
                problems.append("'{}': the view no longer exists.".format(row.name))
                continue
            try:
                _, warnings = etr.update_table(self._doc, view, row.info, path_override)
                problems.extend(warnings)
                done += 1
            except etr.TableError as ex:
                problems.append("'{}': {}".format(row.name, ex))
        self._load_tables()
        self._status("Updated {} of {} table(s).".format(done, len(rows)))
        if problems:
            T3Dialog.show_warning("Updated {} of {} table(s).".format(done, len(rows)),
                                  "Update Excel Tables", details="\n".join(problems[:30]),
                                  owner=self)

    def _on_tbl_reload(self, sender, e):
        rows = self._ticked_tables()
        if len(rows) != 1:
            T3Dialog.show_info("Tick exactly one table, then pick the workbook it should "
                               "read from.", "Reload From", owner=self)
            return
        path = self._pick_file("Reload '{}' from workbook".format(rows[0].name))
        if not path:
            return
        info = rows[0].info
        if not T3Dialog.confirm(
                "Read '{}' from {} (sheet '{}', {}) and update it?".format(
                    rows[0].name, os.path.basename(path), info.get('sheet'),
                    info.get('range') or xt.USED_RANGE),
                "Reload From", ok_text="Reload", owner=self):
            return
        self._run_updates(rows, path_override=path)

    def _on_tbl_unlink(self, sender, e):
        rows = self._ticked_tables()
        if not rows:
            T3Dialog.show_info("Tick the tables to unlink.", "Unlink", owner=self)
            return
        if not T3Dialog.confirm(
                "Stop syncing {} table(s)? The drawings stay in the model as static "
                "copies; only the link to Excel is removed.".format(len(rows)),
                "Unlink Excel Tables", ok_text="Unlink", owner=self):
            return
        views = [v for v in (self._view_of(r) for r in rows) if v is not None]
        try:
            n = etr.unlink_tables(self._doc, views)
        except Exception as ex:
            T3Dialog.show_error("Unlink failed — nothing was changed.", "Unlink",
                                details=short_error(ex), owner=self)
            return
        self._load_tables()
        self._status("Unlinked {} table(s).".format(n))

    # ------------------------------------------------------------------
    # EXCEL TABLES — new table form
    # ------------------------------------------------------------------

    def _on_browse(self, sender, e):
        path = self._pick_file("Pick the Excel workbook")
        if path:
            self._open_source(path)

    def _on_file_key(self, sender, e):
        from System.Windows.Input import Key
        if e.Key == Key.Return:
            self._open_source(self.txt_file.Text)
            e.Handled = True

    def _on_reread(self, sender, e):
        if self._gsheet:
            self._open_source(self._gsheet, keep_sheet=self._sheet, keep_range=True)
        elif self._xl_path:
            self._open_workbook(self._xl_path, keep_sheet=self._sheet, keep_range=True)

    def _open_source(self, text, keep_sheet=None, keep_range=False):
        """A local .xlsx path, a Google Sheets link or a .gsheet shortcut."""
        text = (text or '').strip().strip('"')
        if not text:
            T3Dialog.show_info("Browse for an .xlsx workbook, or paste a Google Sheets link "
                               "in the Workbook box.", "Open Workbook", owner=self)
            return
        if not gs.is_google(text):
            if not os.path.isfile(text):
                T3Dialog.show_warning("'{}' is neither an Excel file on this computer nor a "
                                      "Google Sheets link.".format(text), "Open Workbook",
                                      owner=self)
                return
            self._gsheet = None
            self._source_title = os.path.splitext(os.path.basename(text))[0]
            self._open_workbook(text, keep_sheet, keep_range)
            return
        self._status("Downloading the Google Sheet...")
        try:
            path, src = gs.download(text)
        except gs.GSheetError as ex:
            self._status("Google Sheet not opened.")
            T3Dialog.show_error(str(ex), "Cannot Open Google Sheet", owner=self)
            return
        same = self._gsheet == src['url']
        self._gsheet = src['url']
        self._source_title = "Google Sheet"
        self._open_workbook(path, keep_sheet, keep_range and same)
        self.txt_file.Text = src['url']

    def _open_workbook(self, path, keep_sheet=None, keep_range=False):
        try:
            with xt.Workbook(path) as wb:
                sheets = wb.sheet_names()
        except xt.XlsxError as ex:
            T3Dialog.show_error(str(ex), "Cannot Read Workbook", owner=self)
            return
        first_open = path != self._xl_path
        self._xl_path = path
        self._form_cache = None
        self.txt_file.Text = path
        self.btn_reread.IsEnabled = True
        self._keep_range_once = keep_range and not first_open
        self.pnl_sheet_tabs.Children.Clear()
        try:
            chip = self.FindResource("T3.Chip")
        except Exception:
            chip = None
        target = keep_sheet if keep_sheet in sheets else sheets[0]
        to_check = None
        for name in sheets:
            rb = RadioButton()
            rb.Content = name
            rb.GroupName = "xl_sheet_tabs"
            rb.ToolTip = "Show sheet '{}'".format(name)
            if chip is not None:
                rb.Style = chip
            rb.Checked += self._on_sheet_tab
            self.pnl_sheet_tabs.Children.Add(rb)
            if name == target:
                to_check = rb
        if to_check is not None:
            if to_check.IsChecked:
                self._sheet = target
                self._show_sheet(keep_range=self._keep_range_once)
            else:
                to_check.IsChecked = True             # fires _on_sheet_tab
        self._status("Workbook: {} ({} sheet(s)).".format(
            self._source_title or os.path.basename(path), len(sheets)))

    def _on_sheet_tab(self, sender, e):
        name = str(sender.Content)
        keep = self._keep_range_once and name == self._sheet
        self._keep_range_once = False
        self._sheet = name
        self._show_sheet(keep_range=keep)

    # ── the sheet viewer ─────────────────────────────────────────────
    def _cell_template(self, j):
        xaml = (u'<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">'
                u'<Border Background="{{Binding b{0}}}">'
                u'<TextBlock Text="{{Binding t{0}}}" Foreground="{{Binding f{0}}}" '
                u'FontWeight="{{Binding w{0}}}" FontStyle="{{Binding i{0}}}" '
                u'TextAlignment="{{Binding a{0}}}" TextTrimming="CharacterEllipsis" '
                u'Padding="4,0" VerticalAlignment="Center"/></Border></DataTemplate>').format(j)
        return XamlReader.Parse(xaml)

    def _row_number_template(self):
        return XamlReader.Parse(
            u'<DataTemplate xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation">'
            u'<TextBlock Text="{Binding rn}" Foreground="{DynamicResource T3.TextMuted}" '
            u'FontFamily="Consolas" TextAlignment="Right" Padding="0,0,4,0" '
            u'VerticalAlignment="Center"/></DataTemplate>')

    def _theme_text_color(self):
        try:
            return self.FindResource("T3.Text").Color.ToString()
        except Exception:
            return "Black"

    def _show_sheet(self, keep_range=False):
        """Fill the viewer with the active sheet; keep or reset the range."""
        if not self._xl_path or not self._sheet:
            return
        old_choice = _combo_text(self.cmb_range)
        old_text = (self.txt_range.Text or '').strip()
        try:
            with xt.Workbook(self._xl_path) as wb:
                view = wb.read_view(self._sheet, bool(self.chk_hidden.IsChecked),
                                    VIEW_MAX_ROWS, VIEW_MAX_COLS)
                self._choices = wb.range_choices(self._sheet)
        except xt.XlsxError as ex:
            self._view = None
            self.dg_sheet.ItemsSource = None
            self.dg_sheet.Columns.Clear()
            self.sheet_empty.Text = str(ex)
            self.sheet_empty.Visibility = Visibility.Visible
            self.txt_selection.Text = u"—"
            self._update_summary(error=str(ex))
            return
        self._view = view
        self._build_viewer(view)
        self.sheet_empty.Visibility = Visibility.Collapsed

        stem = self._source_title or os.path.splitext(os.path.basename(self._xl_path))[0]
        if not keep_range or not (self.txt_name.Text or '').strip():
            self.txt_name.Text = u"{} - {}".format(stem, self._sheet)

        self._syncing = True
        try:
            labels = [RANGE_SELECTION] + self._choices
            if keep_range and old_choice in labels:
                _combo_fill(self.cmb_range, labels, labels.index(old_choice))
            else:
                _combo_fill(self.cmb_range, labels, 1)        # Used range
                old_choice = xt.USED_RANGE
        finally:
            self._syncing = False
        if keep_range and old_choice == RANGE_SELECTION and old_text:
            self._apply_typed_range(old_text)
        else:
            self._apply_choice(old_choice)
        note = ""
        if view.clipped:
            note = " The viewer shows the first {} rows × {} columns; type a range " \
                   "to import beyond them.".format(VIEW_MAX_ROWS, VIEW_MAX_COLS)
        self._status(u"Sheet '{}' — used range {}.{}".format(
            self._sheet, xt.range_text(*view.used_bounds), note))

    def _build_viewer(self, view):
        grid = self.dg_sheet
        grid.ItemsSource = None
        grid.Columns.Clear()
        dt = DataTable()
        dt.Columns.Add("rn")
        dt.Columns.Add("ri")
        for j in range(view.cols):
            for key in ('t', 'b', 'f', 'w', 'i', 'a'):
                dt.Columns.Add("{}{}".format(key, j))
        fg_default = self._theme_text_color()
        align = {'left': 'Left', 'center': 'Center', 'right': 'Right'}
        for i, row in enumerate(view.cells):
            dr = dt.NewRow()
            dr["rn"] = str(view.source_rows[i])
            dr["ri"] = str(i)
            for j, cell in enumerate(row):
                dr["t{}".format(j)] = cell.text.replace('\n', ' ')
                dr["b{}".format(j)] = ("#{:02X}{:02X}{:02X}".format(*cell.fill)
                                       if cell.fill else "Transparent")
                col = cell.font.get('color')
                dr["f{}".format(j)] = ("#{:02X}{:02X}{:02X}".format(*col)
                                       if col and tuple(col) != (0, 0, 0) else fg_default)
                dr["w{}".format(j)] = "Bold" if cell.font.get('bold') else "Normal"
                dr["i{}".format(j)] = "Italic" if cell.font.get('italic') else "Normal"
                dr["a{}".format(j)] = align.get(cell.halign, 'Left')
            dt.Rows.Add(dr)

        rn_col = DataGridTemplateColumn()
        rn_col.Header = "#"
        rn_col.Width = DataGridLength(44)
        rn_col.CellTemplate = self._row_number_template()
        grid.Columns.Add(rn_col)
        for j, c in enumerate(view.source_cols):
            col = DataGridTemplateColumn()
            col.Header = xt.col_letter(c)
            col.Width = DataGridLength(max(24.0, view.col_widths_in[j] * 96.0))
            col.CellTemplate = self._cell_template(j)
            grid.Columns.Add(col)
        grid.ItemsSource = dt.DefaultView

    # ── selection <-> range ──────────────────────────────────────────
    def _on_cells_selected(self, sender, e):
        if self._syncing or self._view is None:
            return
        cells = []
        for info in self.dg_sheet.SelectedCells:
            try:
                header = str(info.Column.Header)
                if header == "#":
                    continue
                ri = int(str(info.Item["ri"]))
                cells.append((self._view.source_rows[ri], xt.col_number(header)))
            except Exception:
                continue
        rng = xt.bounds_of_cells(cells)
        if not rng:
            return
        self._syncing = True
        try:
            self.cmb_range.SelectedIndex = 0              # Selected cells
            self.txt_range.Text = rng
        finally:
            self._syncing = False
        self._show_range(xt.parse_range(rng))

    def _on_range_changed(self, sender, e):
        if self._syncing:
            return
        self._apply_choice(_combo_text(self.cmb_range))

    def _apply_choice(self, choice):
        if not self._view or not choice:
            return
        if choice == RANGE_SELECTION:
            self._apply_typed_range(self.txt_range.Text)
            return
        try:
            if choice == xt.USED_RANGE:
                bounds = self._view.used_bounds
            else:
                with xt.Workbook(self._xl_path) as wb:
                    bounds = wb.resolve_range(self._sheet, choice)
        except xt.XlsxError as ex:
            self._update_summary(error=str(ex))
            return
        self._syncing = True
        try:
            self.txt_range.Text = xt.range_text(*bounds)
        finally:
            self._syncing = False
        self._highlight(bounds)
        self._show_range(bounds)

    def _on_range_key(self, sender, e):
        from System.Windows.Input import Key
        if e.Key == Key.Return:
            self._on_range_typed(sender, e)
            e.Handled = True

    def _on_range_typed(self, sender, e):
        if self._syncing or not self._view:
            return
        self._syncing = True
        try:
            self.cmb_range.SelectedIndex = 0
        finally:
            self._syncing = False
        self._apply_typed_range(self.txt_range.Text)

    def _apply_typed_range(self, text):
        try:
            bounds = xt.parse_range(text)
        except xt.XlsxError as ex:
            self.txt_selection.Text = u"—"
            self._update_summary(error=str(ex))
            return
        self._syncing = True
        try:
            self.txt_range.Text = xt.range_text(*bounds)
        finally:
            self._syncing = False
        self._highlight(bounds)
        self._show_range(bounds)

    def _highlight(self, bounds):
        """Select `bounds` in the viewer so the user sees what will be read."""
        grid, view = self.dg_sheet, self._view
        r1, c1, r2, c2 = bounds
        rows = [i for i, r in enumerate(view.source_rows) if r1 <= r <= r2]
        cols = [col for col in list(grid.Columns)[1:]
                if c1 <= xt.col_number(str(col.Header)) <= c2]
        self._syncing = True
        try:
            grid.SelectedCells.Clear()
            if rows and cols and len(rows) * len(cols) <= HIGHLIGHT_MAX_CELLS:
                items = grid.Items
                for i in rows:
                    item = items[i]
                    for col in cols:
                        grid.SelectedCells.Add(DataGridCellInfo(item, col))
            if rows:
                grid.ScrollIntoView(grid.Items[rows[0]])
        except Exception:
            pass
        finally:
            self._syncing = False

    def _show_range(self, bounds):
        r1, c1, r2, c2 = bounds
        self._range_bounds = bounds
        self.txt_selection.Text = u"{}  ·  {} × {}".format(
            xt.range_text(*bounds), r2 - r1 + 1, c2 - c1 + 1)
        self._update_summary()

    def _on_kind_changed(self, sender, e):
        self.txt_kind_hint.Text = KIND_HINT.get(self._kind(), '')
        self._update_summary()

    def _kind(self):
        i = self.cmb_kind.SelectedIndex
        return KIND_ORDER[i] if 0 <= i < len(KIND_ORDER) else etr.KIND_DRAFTING

    def _range_for_read(self):
        """What the link stores: the name of a named range / Excel table (it
        follows the table when it grows in Excel) or the fixed A1 range."""
        choice = _combo_text(self.cmb_range)
        if choice and choice != RANGE_SELECTION:
            return choice
        return (self.txt_range.Text or '').strip()

    def _update_summary(self, error=None):
        if error:
            self.txt_summary.Text = error
            return
        if not self._xl_path or not self._range_bounds:
            self.txt_summary.Text = "Pick a workbook and the cells to import."
            return
        r1, c1, r2, c2 = self._range_bounds
        choice = _combo_text(self.cmb_range)
        source = u"{}!{}".format(self._sheet, xt.range_text(r1, c1, r2, c2))
        if choice and choice not in (RANGE_SELECTION, xt.USED_RANGE):
            source += u" ('{}')".format(choice)
        self.txt_summary.Text = u"{} · {} rows × {} columns → {} '{}'{}.".format(
            source, r2 - r1 + 1, c2 - c1 + 1, etr.KIND_LABEL[self._kind()].lower(),
            (self.txt_name.Text or '').strip() or "Excel Table",
            ", linked" if self.chk_keep_link.IsChecked else ", static copy")

    def _read_form_table(self):
        """(Table, None) or (None, error text) for the current form."""
        if not self._xl_path or not self._sheet:
            return None, "Pick an Excel workbook first."
        rng = self._range_for_read()
        if not rng:
            return None, "Drag over the cells to import, or pick a range."
        hidden = bool(self.chk_hidden.IsChecked)
        mtime = os.path.getmtime(self._xl_path) if os.path.isfile(self._xl_path) else 0
        key = (self._xl_path, self._sheet, rng, hidden, mtime)
        if self._form_cache and self._form_cache[0] == key:
            return self._form_cache[1], None
        try:
            table = xt.read_table(self._xl_path, self._sheet, rng, hidden)
        except xt.XlsxError as ex:
            return None, str(ex)
        self._form_cache = (key, table)
        return table, None

    def _create_table(self):
        table, err = self._read_form_table()
        if err:
            T3Dialog.show_warning(err, "Create Table", owner=self)
            return
        try:
            scale = float((self.txt_scale.Text or '100').strip()) / 100.0
            if not 0.1 <= scale <= 10:
                raise ValueError
        except ValueError:
            T3Dialog.show_warning("Size % must be a number from 10 to 1000.", "Create Table",
                                  owner=self)
            return
        kind = self._kind()
        name = (self.txt_name.Text or '').strip() or "Excel Table"
        link = bool(self.chk_keep_link.IsChecked)
        options = {'scale': scale, 'formatting': bool(self.chk_formatting.IsChecked),
                   'hidden': bool(self.chk_hidden.IsChecked)}
        size = "{} rows × {} columns".format(table.rows, table.cols)
        question = "Create {} '{}' from {}!{} ({})?".format(
            etr.KIND_LABEL[kind].lower(), name, table.sheet,
            xt.range_text(*table.bounds), size)
        if not link:
            details = "Static copy: it will not follow later edits."
        elif self._gsheet:
            details = ("It stays linked to the Google Sheet: Update Tables downloads the sheet "
                       "again and redraws the table.")
        else:
            details = "It stays linked: Update Tables redraws it when the Excel file changes."
        if table.rows * table.cols > 3000:
            details += "\n\nLarge range — building it can take a while."
        if not T3Dialog.confirm(question, "Create Table", ok_text="Create",
                                details=details, owner=self):
            return
        self._status("Building {}...".format(name))
        try:
            view = etr.create_table(self._doc, self._xl_path, table, kind, name, options, link,
                                    gsheet=self._gsheet)
        except etr.TableError as ex:
            self._status("Create failed.")
            T3Dialog.show_error("Could not create the table.", "Create Table",
                                details=str(ex), owner=self)
            return
        self._load_tables()
        try:
            self._uidoc.RequestViewChange(view)     # opens it when the dialog closes
        except Exception:
            pass
        self._status("Created {} '{}' — {}{}.".format(
            etr.KIND_LABEL[kind].lower(), view.Name, size, ", linked" if link else ""))

    # ------------------------------------------------------------------
    # SCHEDULE DATA — schedule list
    # ------------------------------------------------------------------

    def _load_schedules(self):
        # Excel-table schedules (New Table page) hold no element rows: keep them out.
        table_ids = set(r.view_id for r in self._table_rows
                        if r.info.get('kind') == etr.KIND_SCHEDULE)
        self._schedules = [(n, s) for n, s in sd.list_schedules(self._doc)
                           if eid_value(s.Id) not in table_ids]
        self._sched_rows = [ScheduleRow(n, s) for n, s in self._schedules]
        self._fill_schedules()

    def _fill_schedules(self):
        q = (self.data_search_box.Text or '').strip().lower()
        rows = [r for r in self._sched_rows if not q or q in r.name.lower()]
        self.set_items_source(self.dg_schedules, rows)
        self.data_empty.Visibility = _visible(not rows)
        self._count_schedules()

    def _count_schedules(self):
        ticked = sum(1 for r in self._sched_rows if r.is_checked)
        total = len(self._sched_rows)
        self.data_count_text.Text = ("{} / {} ticked".format(ticked, total) if ticked
                                     else "{} schedule(s)".format(total))
        self.sync_header_checkbox(self.chk_all_dg_schedules, self.dg_schedules, "is_checked")

    def _ticked_schedules(self):
        ids = set(r.sched_id for r in self._sched_rows if r.is_checked)
        return [(n, s) for n, s in self._schedules if eid_value(s.Id) in ids]

    def _export(self):
        chosen = self._ticked_schedules()
        if not chosen:
            T3Dialog.show_info("Tick the schedules to export.", "Export to Excel", owner=self)
            return
        default = (chosen[0][0] if len(chosen) == 1 else "Schedules") + ".xlsx"
        for ch in '/\\*?[]:"<>|':
            default = default.replace(ch, '_')
        path = self._save_file(default)
        if not path:
            return
        self._status("Exporting {} schedule(s)...".format(len(chosen)))
        try:
            rows = sd.export_xlsx(self._doc, chosen, path)
        except (IOError, OSError) as ex:
            self._status("Export failed.")
            T3Dialog.show_error("Could not write {}. If it is open in Excel, close it and "
                                "try again.".format(os.path.basename(path)), "Export to Excel",
                                details=short_error(ex), owner=self)
            return
        except Exception as ex:
            self._status("Export failed.")
            T3Dialog.show_error("Export failed.", "Export to Excel", details=short_error(ex),
                                owner=self)
            return
        self._status("Exported {} schedule(s), {} row(s) to {}.".format(
            len(chosen), rows, os.path.basename(path)))
        if T3Dialog.confirm("Exported {} schedule(s), {} row(s). Open the file in Excel now?".format(
                len(chosen), rows), "Export Complete", ok_text="Open in Excel",
                cancel_text="Not now", details=path, owner=self):
            try:
                os.startfile(path)
            except OSError as ex:
                T3Dialog.show_warning("Could not open the file.", "Export Complete",
                                      details=short_error(ex), owner=self)

    # ------------------------------------------------------------------
    # SCHEDULE DATA — import, review, apply
    # ------------------------------------------------------------------

    def _on_import(self, sender, e):
        path = self._pick_file("Pick the edited workbook")
        if not path:
            return
        ticked = self._ticked_schedules()
        fallback = ticked[0] if len(ticked) == 1 else None
        self._status("Reading {}...".format(os.path.basename(path)))
        try:
            changes, report = sd.read_changes(self._doc, path, self._schedules, fallback)
        except xt.XlsxError as ex:
            self._status("Import failed.")
            T3Dialog.show_error(str(ex), "Import from Excel", owner=self)
            return
        self._show_changes(changes, report)
        self._status("{}: {} change(s) found.".format(os.path.basename(path), len(changes)))

    def _show_changes(self, changes, report):
        self._change_rows = [ChangeRow(c) for c in changes]
        self.set_items_source(self.dg_changes, self._change_rows)
        self.changes_empty.Visibility = _visible(not self._change_rows)
        if report and not changes and report != ["No file imported."]:
            self.changes_empty.Text = "No value differs from the model."
        self.txt_changes_info.Text = "  ".join(report) if report else ""
        self._count_changes()

    def _count_changes(self):
        n = sum(1 for r in self._change_rows if r.is_checked)
        self.btn_apply.IsEnabled = n > 0
        self.txt_apply.Text = "Apply {} change(s)".format(n) if n else "Apply changes"
        self.sync_header_checkbox(self.chk_all_dg_changes, self.dg_changes, "is_checked")

    def _on_apply(self, sender, e):
        rows = [r for r in self._change_rows if r.is_checked]
        if not rows:
            return
        elements = len(set(r.element_id for r in rows))
        if not T3Dialog.confirm(
                "Write {} change(s) to {} element(s)? Ctrl+Z undoes all of them in one step.".format(
                    len(rows), elements),
                "Apply Excel Changes", ok_text="Apply", owner=self):
            return
        done, errors, failed = sd.apply_changes(self._doc, [r.change for r in rows])
        written = set(id(r.change) for r in rows) - set(id(c) for c in failed)
        remaining = [r.change for r in self._change_rows if id(r.change) not in written]
        self._show_changes(remaining, ["Applied {} of {} change(s).".format(done, len(rows))])
        self._status("Applied {} of {} change(s) to {} element(s).".format(done, len(rows), elements))
        if errors:
            T3Dialog.show_warning("{} of {} change(s) could not be written; they stay in the "
                                  "list.".format(len(rows) - done, len(rows)),
                                  "Apply Excel Changes", details="\n".join(errors[:30]),
                                  owner=self)


# ============================================================
# PUBLIC ENTRY POINT
# ============================================================

def show_schedule_manager(script_dir, revit_obj):
    """Open the Schedule Manager window modally."""
    ManaSchedWindow(script_dir, revit_obj).ShowDialog()
