# -*- coding: utf-8 -*-
"""ManaDWG — Dialog and data collection for CAD Imports & Links Management."""

import os
import sys

from pyrevit import revit, forms, script
from GUI.WPF_Base import T3WPFWindow, to_items_source

try:
    from GUI import RevitTheme as _theme
except Exception:
    _theme = None

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import WindowState, LogicalTreeHelper
from System.Windows.Media import VisualTreeHelper
from Autodesk.Revit.DB import Transaction

from Snippets import _cad_imports
from Snippets._cad_imports import (DWGItem, collect_dwg_items, delete_cad_items,
                                   resolve_dwg_view)

XAML_FILE = os.path.join(os.path.dirname(__file__), 'Tools', 'DWGManagement.xaml')
logger = script.get_logger()
_cad_imports.logger = logger     # warnings to the pyRevit output, as before

FILTER_ALL     = "All"
FILTER_IMPORTS = "Imports Only"
FILTER_LINKS   = "Links Only"


def request_dwg_view(document, view_id):
    """Request navigation after the modal DWG window has closed."""
    try:
        uidoc = revit.uidoc
        if uidoc is None or uidoc.Document != document:
            raise ValueError("The active Revit document has changed. Reopen DWG Management.")
        view = resolve_dwg_view(document, view_id)
        uidoc.RequestViewChange(view)
    except Exception as ex:
        logger.warning("Could not open CAD owner view: {}".format(ex))
        forms.alert("Could not open view:\n{}".format(ex), title="ManaDWG")


class DWGManagementWindow(T3WPFWindow):
    """WPF dialog for listing, renaming, and deleting CAD imports/links."""

    def __init__(self, doc=None):
        T3WPFWindow.__init__(self, XAML_FILE)
        self._doc = doc or revit.doc
        self._requested_view_id = None
        self._all_items      = []
        self._filtered_items = []
        self._active_filter  = FILTER_ALL

        self._adopt_host_font()
        self._apply_theme()
        self._load_data()

    def _adopt_host_font(self):
        if _theme is None:
            return
        family, size = _theme.host_font()
        if family:
            try:
                self.FontFamily = family
                if size and size > 0:
                    self.FontSize = size
            except Exception:
                pass

    def _apply_theme(self, theme=None):
        if _theme is None:
            return
        try:
            _theme.apply(self, theme)
        except Exception:
            pass

    def _load_data(self):
        try:
            self._all_items = collect_dwg_items(self._doc)
        except Exception as ex:
            forms.alert("Error collecting CAD data:\n{}".format(ex), title="ManaDWG")
            self._all_items = []

        self._apply_filter()
        self._update_status()

    def _apply_filter(self):
        search_text = ""
        try:
            search_text = (self.SearchBox.Text or "").strip().lower()
        except Exception:
            pass

        result = []
        for item in self._all_items:
            if self._active_filter == FILTER_IMPORTS and item.IsLink:
                continue
            if self._active_filter == FILTER_LINKS and not item.IsLink:
                continue
            if search_text:
                haystack = (item.Name + " " + item.ViewNames + " " + item.FilePath).lower()
                if search_text not in haystack:
                    continue
            result.append(item)

        self._filtered_items = result

        try:
            self.DWGDataGrid.ItemsSource = None
            self.DWGDataGrid.ItemsSource = to_items_source(self._filtered_items)
        except Exception as ex:
            logger.warning("Could not bind grid: {}".format(ex))

    def _update_status(self):
        total   = len(self._all_items)
        visible = len(self._filtered_items)
        imports = sum(1 for x in self._all_items if not x.IsLink)
        links   = sum(1 for x in self._all_items if x.IsLink)

        msg = "Showing {} of {} item(s)   |   {} Import(s)   |   {} Link(s)".format(
            visible, total, imports, links)
        try:
            self.StatusText.Text = msg
            self.StatusBarText.Text = msg
        except Exception:
            pass

    def filter_button_clicked(self, sender, e):
        sender.IsChecked = True

        try:
            self.FilterAll.IsChecked     = (sender == self.FilterAll)
            self.FilterImports.IsChecked = (sender == self.FilterImports)
            self.FilterLinks.IsChecked   = (sender == self.FilterLinks)
        except Exception:
            pass

        if sender == self.FilterImports:
            self._active_filter = FILTER_IMPORTS
        elif sender == self.FilterLinks:
            self._active_filter = FILTER_LINKS
        else:
            self._active_filter = FILTER_ALL

        self._apply_filter()
        self._update_status()

    def search_text_changed(self, sender, e):
        self._apply_filter()
        self._update_status()

    def header_checkbox_clicked(self, sender, e):
        checked = sender.IsChecked
        for item in self._filtered_items:
            item.IsSelected = checked if checked is not None else False
        try:
            self.DWGDataGrid.ItemsSource = None
            self.DWGDataGrid.ItemsSource = to_items_source(self._filtered_items)
        except Exception:
            pass

    def view_name_mouse_down(self, sender, e):
        if e.ClickCount != 2:
            return
        source = e.OriginalSource
        while source is not None and getattr(source, "Name", None) != "DWGViewName":
            if source == sender:
                return
            try:
                parent = VisualTreeHelper.GetParent(source)
            except Exception:
                parent = None
            if parent is None:
                try:
                    parent = LogicalTreeHelper.GetParent(source)
                except Exception:
                    parent = None
            source = parent
        item = getattr(source, "DataContext", None)
        if not isinstance(item, DWGItem):
            return
        e.Handled = True
        try:
            resolve_dwg_view(self._doc, item.OwnerViewId)
        except Exception as ex:
            forms.alert("Could not open view:\n{}".format(ex), title="ManaDWG")
            return
        self._requested_view_id = item.OwnerViewId
        self.Close()

    def refresh_button_clicked(self, sender, e):
        self._load_data()

    def rename_button_clicked(self, sender, e):
        selected = self.DWGDataGrid.SelectedItem
        if selected is None:
            forms.alert("Please select a CAD item in the list.", title="Rename")
            return

        if not selected.IsLink:
            forms.alert(
                "Rename is only supported for CAD Links, not CAD Imports.\n\n"
                "Selected item '{}' is a CAD Import.".format(selected.Name),
                title="Rename"
            )
            return

        if selected.LinkType is None:
            forms.alert("Cannot rename this item - its type element could not be found.", title="Rename")
            return

        new_name = forms.ask_for_string(
            prompt="Enter a new name for the CAD link:",
            default=selected.Name,
            title="Rename CAD Link"
        )

        if new_name is None:
            return

        new_name = new_name.strip()
        if not new_name:
            forms.alert("Name cannot be empty.", title="Rename")
            return

        if new_name == selected.Name:
            return

        t = Transaction(self._doc, "T3Lab: Rename CAD Link")
        t.Start()
        try:
            selected.LinkType.Name = new_name
            t.Commit()
        except Exception as ex:
            t.RollBack()
            forms.alert("Failed to rename CAD link:\n{}".format(ex), title="Rename Error")
            return

        self._load_data()

    def delete_button_clicked(self, sender, e):
        selected_items = [item for item in self._filtered_items if item.IsSelected]

        if not selected_items:
            grid_sel = self.DWGDataGrid.SelectedItem
            if grid_sel is not None:
                selected_items = [grid_sel]

        if not selected_items:
            forms.alert(
                "Please select one or more CAD items to delete.\n\n"
                "Use the checkboxes or click a row to select.",
                title="Delete"
            )
            return

        count = len(selected_items)
        names = "\n".join("  - {}  [{}]".format(i.Name, i.DWGType) for i in selected_items[:10])
        if count > 10:
            names += "\n  ... and {} more".format(count - 10)

        pinned = sum(1 for i in selected_items if i.Instance is not None and i.Instance.Pinned)
        pinned_note = "{} pinned item(s) will be unpinned and deleted.\n\n".format(pinned) if pinned else ""

        confirmed = forms.alert(
            "Delete {} CAD item(s)?\n\n{}\n\n{}"
            "For CAD Links, the link type will also be removed if no other views reference it.\n\n"
            "This action cannot be undone.".format(count, names, pinned_note),
            title="Confirm Delete",
            yes=True,
            cancel=True
        )

        if not confirmed:
            return

        try:
            deleted, failures = delete_cad_items(self._doc, selected_items)
        except Exception as ex:
            forms.alert(
                "Nothing was deleted: {}\n\n"
                "Check that no other user has these items borrowed, then try again.".format(ex),
                title="Delete Error"
            )
            return

        self._load_data()
        self.StatusText.Text = "Deleted {} of {} CAD item(s)   |   {}".format(deleted, count, self.StatusText.Text)

        if failures:
            lines = "\n".join("  - {}: {}".format(name, reason) for name, reason in failures[:10])
            if len(failures) > 10:
                lines += "\n  ... and {} more".format(len(failures) - 10)
            forms.alert(
                "Deleted {} of {} CAD item(s). {} could not be deleted and stay in the model:\n\n{}\n\n"
                "Check that no other user has them borrowed, then try again.".format(
                    deleted, count, len(failures), lines),
                title="Delete"
            )

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
        else:
            self.WindowState = WindowState.Maximized

    def close_button_clicked(self, sender, e):
        self.Close()


def show_dwg_manager(doc=None):
    d = doc or revit.doc
    if not d:
        forms.alert("No active Revit document found.", title="ManaDWG")
        return
    win = DWGManagementWindow(d)
    win.ShowDialog()
    if win._requested_view_id is not None:
        request_dwg_view(d, win._requested_view_id)
