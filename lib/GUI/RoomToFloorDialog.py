# -*- coding: utf-8 -*-
"""RoomToFloor — Dialog and logic for generating floors from room boundaries."""

import os
import sys
import json

from pyrevit import revit, forms, script
from GUI.WPF_Base import T3WPFWindow, to_items_source
from Snippets._compat import make_eid, eid_value, net_list
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
    ElementId,
)
from Autodesk.Revit.UI import TaskDialog
# Revit side shared with the room_to_floor tool (core/tools/model.py).
from Snippets.finish_floors import FloorGenerator

XAML_FILE = os.path.join(os.path.dirname(__file__), 'Tools', 'RoomToFloor.xaml')
logger = script.get_logger()


class RoomItem(object):
    """Represents a room item in the DataGrid."""
    def __init__(self, room_element, document=None):
        self.Element = room_element
        self.IsSelected = False
        p_num = room_element.LookupParameter("Number")
        self.Number = p_num.AsString() if p_num else ""
        p_name = room_element.LookupParameter("Name")
        self.Name = p_name.AsString() if p_name else ""
        try:
            doc = document or room_element.Document
            level = doc.GetElement(room_element.LevelId)
            self.Level = level.Name if level else ""
        except Exception:
            self.Level = ""


class RoomToFloorWindow(T3WPFWindow):
    """WPF window for creating floors from rooms."""

    PP_PANEL      = "r2f_progress_panel"
    PP_BAR        = "r2f_pb"
    PP_PAUSE      = "r2f_btn_pause"
    PP_STOP       = "r2f_btn_stop"
    PP_PAUSE_ICON = "r2f_btn_pause_icon"
    PP_PAUSE_TEXT = "r2f_btn_pause_label"
    PP_STATUS     = "status_text"
    PP_STOP_MSG   = u"Stopping… finishing current room"

    def __init__(self, doc=None, uidoc=None):
        T3WPFWindow.__init__(self, XAML_FILE)
        self._doc = doc or revit.doc
        self._uidoc = uidoc
        if self._uidoc is None:
            try:
                self._uidoc = revit.uidoc
            except Exception:
                self._uidoc = None

        self.generator = FloorGenerator(self._doc)
        self._all_rooms = []
        self._floor_type_map = {}
        # Read per window: the module stays loaded while the user switches
        # between a metric and an imperial project.
        self._unit = project_length_unit(self._doc)

        self._apply_units()
        self._load_rooms()
        self._load_floor_types()
        self._update_status()

    def _apply_units(self):
        """Offset label and default in the length unit the project displays."""
        self.lbl_offset.Text = self._unit.label("Height Offset") + ":"
        self.txt_offset.Text = self._unit.text(0.0)

    def _load_rooms(self):
        try:
            room_elements = FilteredElementCollector(self._doc) \
                .OfCategory(BuiltInCategory.OST_Rooms) \
                .ToElements()
        except Exception:
            room_elements = []

        self._all_rooms = []
        for r in room_elements:
            if r.Location is None:
                continue
            self._all_rooms.append(RoomItem(r, self._doc))

        self._all_rooms.sort(key=lambda x: (x.Level, x.Number))
        self.room_datagrid.ItemsSource = to_items_source(self._all_rooms)

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
            self._floor_type_map["{}: {}".format(fam_name, type_name)] = t

        sorted_names = sorted(self._floor_type_map.keys())
        for name in sorted_names:
            self.cmb_floor_type.Items.Add(name)

        if self.cmb_floor_type.Items.Count > 0:
            self.cmb_floor_type.SelectedIndex = 0

    def _get_selected_rooms(self):
        return [r for r in self._all_rooms if r.IsSelected]

    def _update_status(self):
        selected = len(self._get_selected_rooms())
        total = len(self._all_rooms)
        self.status_count.Text = "{} rooms".format(total)
        self.status_text.Text = "{} room(s) selected".format(selected) if selected > 0 else "Ready"

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Normal if self.WindowState == WindowState.Maximized else WindowState.Maximized

    def close_button_clicked(self, sender, e):
        self.Close()

    def select_all_clicked(self, sender, e):
        for r in self._all_rooms:
            r.IsSelected = True
        self.room_datagrid.Items.Refresh()
        self._update_status()
        self.sync_header_checkbox(
            self.FindName("chk_all_room_datagrid"), self.room_datagrid, "IsSelected")

    def select_none_clicked(self, sender, e):
        for r in self._all_rooms:
            r.IsSelected = False
        self.room_datagrid.Items.Refresh()
        self._update_status()
        self.sync_header_checkbox(
            self.FindName("chk_all_room_datagrid"), self.room_datagrid, "IsSelected")

    def search_changed(self, sender, e):
        query = self.txt_search.Text.strip().upper()
        if not query:
            self.room_datagrid.ItemsSource = to_items_source(self._all_rooms)
        else:
            self.room_datagrid.ItemsSource = to_items_source([
                r for r in self._all_rooms
                if query in r.Name.upper() or query in r.Number.upper() or query in r.Level.upper()
            ])
        self._update_status()

    def create_floors_clicked(self, sender, e):
        selected_rooms = self._get_selected_rooms()
        if not selected_rooms:
            TaskDialog.Show("Room to Floor", "Please select at least one room.")
            return

        floor_type_name = self.cmb_floor_type.SelectedItem
        if not floor_type_name:
            TaskDialog.Show("Room to Floor", "Please select a floor type.")
            return

        floor_type = self._floor_type_map[floor_type_name]

        # Bare number = project unit; "150 mm", "6\"" or "0'-6\"" work anywhere.
        # FloorGenerator (and run_headless for MCP) keep taking millimetres.
        try:
            offset_mm = self._unit.parse_mm(self.txt_offset.Text)
        except ValueError as ex:
            msg = "Height Offset: {}".format(ex)
            self.status_text.Text = msg
            TaskDialog.Show("Room to Floor", msg)
            return

        is_structural = self.chk_structural.IsChecked
        use_finish = self.chk_room_finish.IsChecked

        self.begin_progress(len(selected_rooms), disable=[sender])

        def progress_cb(current, total):
            self.step_progress(current, "Creating floor {}/{}...".format(current + 1, total))

        new_floors, created, errors = self.generator.generate_floors(
            [r.Element for r in selected_rooms],
            floor_type,
            offset_mm,
            is_structural,
            use_finish,
            progress_callback=progress_cb,
            cancel_check=lambda: self.is_cancelled
        )

        cancelled = self.is_cancelled
        self.end_progress()

        if new_floors and self._uidoc:
            try:
                self._uidoc.Selection.SetElementIds(net_list(ElementId, [f.Id for f in new_floors if f.IsValidObject]))
            except Exception as e:
                logger.debug("Failed to select created floors: {}".format(e))

        if cancelled:
            msg = "Cancelled — created {} floor(s).".format(created)
        else:
            msg = "Successfully created {} floors.".format(created)
        if errors > 0:
            msg += "\n{} errors occurred.".format(errors)

        TaskDialog.Show("Room to Floor", msg)
        self.Close()

    def select_all_room_datagrid_clicked(self, sender, e):
        """Header checkbox: select/deselect all rows currently shown in room_datagrid."""
        self.toggle_all_rows(self.room_datagrid, "IsSelected", sender.IsChecked)


def run_headless(args_json, doc=None):
    d = doc or revit.doc
    try:
        data = json.loads(args_json)
        gen = FloorGenerator(d)

        room_ids = data.get("room_ids", [])
        type_id = data.get("type_id")
        offset = data.get("offset_mm", 0)
        structural = data.get("structural", False)
        use_finish = data.get("use_finish", True)

        rooms = [d.GetElement(make_eid(int(rid))) for rid in room_ids if rid]
        floor_type = d.GetElement(make_eid(int(type_id))) if type_id else None

        new_floors, created, errors = gen.generate_floors(rooms, floor_type, offset, structural, use_finish)
        print(json.dumps({"status": "success", "created": created, "errors": errors}))
    except Exception as ex:
        print(json.dumps({"status": "error", "message": str(ex)}))


def show_room_to_floor_dialog(doc=None, uidoc=None):
    d = doc or revit.doc
    if not d:
        forms.alert("No active Revit document found.", title="Room to Floor")
        return
    u = uidoc
    if u is None:
        try:
            u = revit.uidoc
        except Exception:
            u = None
    win = RoomToFloorWindow(d, u)
    win.ShowDialog()
