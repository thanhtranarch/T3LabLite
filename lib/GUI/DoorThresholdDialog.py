# -*- coding: utf-8 -*-
"""DoorThreshold — Dialog and logic for generating threshold floors under doors."""

import os
import sys

from pyrevit import revit, forms, script
from GUI.WPF_Base import T3WPFWindow, to_items_source
from Snippets._compat import eid_value, net_list
from Snippets._units import project_length_unit


import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import WindowState
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    FloorType,
    BuiltInParameter,
    ElementId,
)
from Autodesk.Revit.UI import TaskDialog
# Revit side shared with the create_door_thresholds tool (core/tools/model.py).
from Snippets.finish_floors import ThresholdGenerator

XAML_FILE = os.path.join(os.path.dirname(__file__), 'Tools', 'DoorThreshold.xaml')
logger = script.get_logger()


class DoorItem(object):
    def __init__(self, door_element, document=None):
        self.Element = door_element
        self.IsSelected = False

        mark_param = door_element.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
        if not mark_param:
            mark_param = door_element.LookupParameter("Mark")

        self.Mark = mark_param.AsString() if (mark_param and mark_param.AsString()) else ""

        try:
            self.Name = door_element.Name or door_element.Symbol.Name or ""
        except Exception:
            self.Name = ""

        try:
            doc = document or door_element.Document
            level = doc.GetElement(door_element.LevelId)
            self.Level = level.Name if (level and level.Name) else ""
        except Exception:
            self.Level = ""


class DoorThresholdWindow(T3WPFWindow):
    def __init__(self, doc=None, uidoc=None):
        T3WPFWindow.__init__(self, XAML_FILE)
        self._doc = doc or revit.doc
        self._uidoc = uidoc
        if self._uidoc is None:
            try:
                self._uidoc = revit.uidoc
            except Exception:
                self._uidoc = None

        # Read per window: the module stays loaded while the user switches
        # between a metric and an imperial project.
        self._unit = project_length_unit(self._doc)
        self.generator = ThresholdGenerator(self._doc, self._unit)
        self._all_doors = []
        self._floor_type_map = {}

        self._apply_units()
        self._load_doors()
        self._load_floor_types()
        self._update_status()

    def _apply_units(self):
        """Offset label and default in the length unit the project displays."""
        self.lbl_offset.Text = self._unit.label("Height Offset") + ":"
        self.txt_offset.Text = self._unit.text(0.0)

    def _load_doors(self):
        pre_selected = set()
        if self._uidoc:
            try:
                selected_ids = self._uidoc.Selection.GetElementIds()
                for eid in selected_ids:
                    pre_selected.add(eid_value(eid))
            except Exception:
                pass

        try:
            door_elements = FilteredElementCollector(self._doc, self._doc.ActiveView.Id) \
                .OfCategory(BuiltInCategory.OST_Doors) \
                .WhereElementIsNotElementType() \
                .ToElements()
        except Exception:
            door_elements = []

        if not door_elements:
            try:
                door_elements = FilteredElementCollector(self._doc) \
                    .OfCategory(BuiltInCategory.OST_Doors) \
                    .WhereElementIsNotElementType() \
                    .ToElements()
            except Exception:
                door_elements = []

        self._all_doors = []
        for d in door_elements:
            item = DoorItem(d, self._doc)
            if eid_value(d.Id) in pre_selected:
                item.IsSelected = True
            self._all_doors.append(item)

        self._all_doors.sort(key=lambda x: (x.Level, x.Mark))
        self.door_datagrid.ItemsSource = to_items_source(self._all_doors)

    def _load_floor_types(self):
        try:
            floor_types = FilteredElementCollector(self._doc) \
                .OfClass(FloorType) \
                .ToElements()
        except Exception:
            floor_types = []

        self._floor_type_map = {}
        for t in floor_types:
            try:
                type_name = t.LookupParameter("Type Name").AsString() if t.LookupParameter("Type Name") else t.Name
            except Exception:
                type_name = t.Name

            if not type_name:
                type_name = ""

            fam_name = t.FamilyName if t.FamilyName else ""
            key = "{}: {}".format(fam_name, type_name)
            self._floor_type_map[key] = t

        sorted_names = sorted(self._floor_type_map.keys())
        for name in sorted_names:
            self.cmb_floor_type.Items.Add(name)

        if self.cmb_floor_type.Items.Count > 0:
            self.cmb_floor_type.SelectedIndex = 0

    def _get_selected_doors(self):
        return [d for d in self._all_doors if d.IsSelected]

    def _update_status(self):
        selected = len(self._get_selected_doors())
        total = len(self._all_doors)
        self.status_count.Text = "{} doors".format(total)
        self.status_text.Text = "{} door(s) selected".format(selected) if selected > 0 else "Ready"

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Normal if self.WindowState == WindowState.Maximized else WindowState.Maximized

    def close_button_clicked(self, sender, e):
        self.Close()

    def select_all_clicked(self, sender, e):
        for d in self._all_doors:
            d.IsSelected = True
        self.door_datagrid.Items.Refresh()
        self._update_status()
        self.sync_header_checkbox(
            self.FindName("chk_all_door_datagrid"), self.door_datagrid, "IsSelected")

    def select_none_clicked(self, sender, e):
        for d in self._all_doors:
            d.IsSelected = False
        self.door_datagrid.Items.Refresh()
        self._update_status()
        self.sync_header_checkbox(
            self.FindName("chk_all_door_datagrid"), self.door_datagrid, "IsSelected")

    def search_changed(self, sender, e):
        query = self.txt_search.Text.strip().upper()
        if not query:
            self.door_datagrid.ItemsSource = to_items_source(self._all_doors)
        else:
            self.door_datagrid.ItemsSource = to_items_source([
                d for d in self._all_doors
                if query in d.Name.upper() or query in d.Mark.upper() or query in d.Level.upper()
            ])
        self._update_status()

    def create_thresholds_clicked(self, sender, e):
        selected_doors = self._get_selected_doors()
        if not selected_doors:
            TaskDialog.Show("Door Threshold", "Please select at least one door.")
            return

        floor_type_name = self.cmb_floor_type.SelectedItem
        if not floor_type_name:
            if not self._floor_type_map:
                TaskDialog.Show(
                    "Door Threshold",
                    "This model has no floor types to use as a threshold.\n\n"
                    "Load or create at least one floor type, then reopen this tool.")
            else:
                TaskDialog.Show("Door Threshold", "Please select a threshold type (floor).")
            return

        floor_type = self._floor_type_map.get(floor_type_name)
        if floor_type is None:
            TaskDialog.Show("Door Threshold",
                            "The selected threshold type is no longer available.")
            return

        # Bare number = project unit; "150 mm", "6\"" or "0'-6\"" work anywhere.
        try:
            offset_mm = self._unit.parse_mm(self.txt_offset.Text)
        except ValueError as ex:
            msg = "Height Offset: {}".format(ex)
            self.status_text.Text = msg
            TaskDialog.Show("Door Threshold", msg)
            return

        new_floors, created, errors, error_messages = self.generator.generate_thresholds(
            [d.Element for d in selected_doors],
            floor_type,
            offset_mm
        )

        if new_floors and self._uidoc:
            try:
                self._uidoc.Selection.SetElementIds(net_list(ElementId, [f.Id for f in new_floors if f.IsValidObject]))
            except Exception as ex:
                logger.debug("Failed to select created floors: {}".format(ex))

        msg = "Successfully created {} thresholds.".format(created)
        if errors > 0:
            msg += "\n{} errors occurred:".format(errors)
            for err in error_messages[:10]:
                msg += "\n- {}".format(err)
            if len(error_messages) > 10:
                msg += "\n... and {} more".format(len(error_messages) - 10)

        TaskDialog.Show("Door Threshold", msg)
        self.Close()

    def select_all_door_datagrid_clicked(self, sender, e):
        """Header checkbox: select/deselect all rows currently shown in door_datagrid."""
        self.toggle_all_rows(self.door_datagrid, "IsSelected", sender.IsChecked)


def show_door_threshold_dialog(doc=None, uidoc=None):
    d = doc or revit.doc
    if not d:
        forms.alert("No active Revit document found.", title="Door Threshold")
        return
    u = uidoc
    if u is None:
        try:
            u = revit.uidoc
        except Exception:
            u = None
    win = DoorThresholdWindow(d, u)
    win.ShowDialog()
