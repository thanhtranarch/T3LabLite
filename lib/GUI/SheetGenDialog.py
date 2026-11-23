#! python3
# -*- coding: utf-8 -*-
"""
Create Room Plan

Create Plan Views from Room List with WPF UI.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__author__  = "Tran Tien Thanh"
__title__   = "Create Plan Views"

# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
import os
import sys
# ─── CPython 3 & lib bootstrap ────────────────────────────────────────────────
# CPython engine paths are injected by lib/_cpython_bootstrap.py below;
# it discovers the engine whatever the pyRevit clone is called.

_cur = os.path.dirname(os.path.abspath(__file__))
while _cur and not os.path.exists(os.path.join(_cur, 'lib')):
    _parent = os.path.dirname(_cur)
    if _parent == _cur:
        break
    _cur = _parent
_lib_dir = os.path.join(_cur, 'lib')
if os.path.exists(_lib_dir) and _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

try:
    import _cpython_bootstrap
    _cpython_bootstrap.init_cpython_paths()
except Exception:
    pass
# ──────────────────────────────────────────────────────────────────────────────
import clr
import re
import math
import json

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System.Windows import WindowState, Visibility, Thickness, CornerRadius, HorizontalAlignment, VerticalAlignment, FontWeights
from System.Windows.Controls import Border, Grid, TextBlock, Canvas, DataGridRow, CheckBox, TextBox
from System.Windows.Media import SolidColorBrush, Color, VisualTreeHelper
from System.Windows.Shapes import Line

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.DB import (
    Transaction,
    View,
    FilteredElementCollector,
    BuiltInCategory,
    ViewType,
    ViewPlan,
    ViewSheet,
    SpatialElementBoundaryOptions,
    SpatialElementBoundaryLocation,
)
from Autodesk.Revit.UI import TaskDialog
from GUI.WPF_Base import T3WPFWindow, to_items_source
# Crop offset = model length (project unit); header strip = paper size (mm
# in a metric project, inches in an imperial one). Read per window.
from Snippets._units import project_length_unit, paper_unit, MILLIMETERS
from Snippets import sheetgen as sg     # the Revit side, shared with generate_room_views

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
# ==================================================
logger        = script.get_logger()
# CPython has no ScriptOutput.GetDefault; safe_output()
# returns a no-op window instead of killing the tool.
try:
    from _cpython_bootstrap import safe_output
    output = safe_output()
except Exception:
    output = script.get_output()
# `revit.doc` / `revit.uidoc` RAISE AttributeError (not return None) when no
# UIDocument is active. At module scope that kills the import outright, so the
# tool dies before it can explain itself. Resolve defensively and let the entry
# point report the real problem.
try:
    uidoc = revit.uidoc
except Exception:
    uidoc = None
try:
    doc = revit.doc
except Exception:
    doc = None
# Read from the Application, not the document: `revit.doc` is None when this
# tool is launched from the Assistant pane or with no project open, and the
# old `int(revit.doc.Application.VersionNumber)` raised at IMPORT time
# ('NoneType' object has no attribute 'Application') so the window never opened.
from Snippets._host import get_revit_version
REVIT_VERSION = get_revit_version()

SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR    = os.path.dirname(os.path.dirname(SCRIPT_DIR))
lib_dir    = os.path.join(EXT_DIR, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

XAML_FILE  = os.path.join(os.path.dirname(__file__), 'Tools', 'SheetGen.xaml')

from GUI.ProgressPauseMixin import ProgressPauseMixin

# Defaults the tool has always used, in mm: 1 m of model space around the
# room, a 70 mm title-block header strip on paper.
DEFAULT_OFFSET_MM = 1000.0
DEFAULT_STRIP_MM  = 70.0

# ╔═╗╦  ╔═╗╔═╗╔═╗╔═╗╔═╗
# ║  ║  ╠═╣╚═╗╚═╗║╣ ╚═╗
# ╚═╝╩═╝╩ ╩╚═╝╚═╝╚═╝╚═╝ CLASSES
# ==================================================

class RoomItem(object):
    """Represents a room item in the DataGrid."""
    def __init__(self, room_element, existing_plan_names=None):
        self.Element = room_element
        self.IsSelected = False
        p_num = room_element.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER)
        self.Number = p_num.AsString() if p_num else ""
        p_name = room_element.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
        self.Name = p_name.AsString() if p_name else ""

        type_param = room_element.LookupParameter("Room Type")
        self.RoomType = type_param.AsString() if type_param else ""

        try:
            level = doc.GetElement(room_element.LevelId)
            self.Level = level.Name if level else ""
        except Exception:
            self.Level = ""

        # Count floor plans matching this room number suffix (#Number)
        search_suffix = "(#{})".format(self.Number)
        if existing_plan_names:
            self.FloorPlanCount = sum(1 for name in existing_plan_names if search_suffix in name)
        else:
            self.FloorPlanCount = 0

        # Quantity to generate
        self.GenQty = 1


class CreateRoomPlanWindow(T3WPFWindow):
    """WPF window for creating plan views from rooms."""

    # ProgressPauseMixin — SheetGen.xaml footer progress panel
    PP_PANEL      = "sg_progress_panel"
    PP_BAR        = "sg_pb"
    PP_PAUSE      = "sg_btn_pause"
    PP_STOP       = "sg_btn_stop"
    PP_PAUSE_ICON = "sg_btn_pause_icon"
    PP_PAUSE_TEXT = "sg_btn_pause_label"
    PP_STATUS     = "status_text"
    PP_STOP_MSG   = u"Stopping… finishing current room"

    def __init__(self, doc_param=None, uidoc_param=None):
        global doc, uidoc
        if doc_param:
            doc = doc_param
        if uidoc_param:
            uidoc = uidoc_param
        T3WPFWindow.__init__(self, XAML_FILE)
        self._all_rooms = []
        self._first_sheet = None
        self._generated_sheets = []
        self._tb_size_cache = {}  # titleblock type id -> (w, h) feet
        self._load_rooms()
        self._load_view_templates()
        self._load_plan_type_options()
        self._load_title_blocks()
        self.cmb_strip_side.SelectedIndex = 0  # default: right (vertical) strip
        self._unit = project_length_unit(doc)
        self._paper = paper_unit(self._unit)
        self._apply_units()
        self._update_status()
        self._update_mockup()

    def _apply_units(self):
        """Unit labels + defaults: crop offset in the project unit, header
        strip in the paper unit. Runs after the data loads because setting
        the text boxes fires their TextChanged (mockup redraw)."""
        self.lbl_offset.Text = self._unit.label("OFFSET")
        self.txt_offset.Text = self._unit.default_text(DEFAULT_OFFSET_MM)
        self.lbl_strip_size.Text = self._paper.label("HEADER SIZE")
        self.txt_strip_mm.Text = self._paper.default_text(DEFAULT_STRIP_MM)


    # ── Data loading ──────────────────────────────────
    def _load_rooms(self):
        """Collect all placed rooms from the document."""
        room_elements = FilteredElementCollector(doc) \
            .OfCategory(BuiltInCategory.OST_Rooms) \
            .ToElements()

        # Collect names of all existing floor plans in the document
        view_elements = FilteredElementCollector(doc) \
            .OfClass(ViewPlan) \
            .WhereElementIsNotElementType() \
            .ToElements()
        
        existing_plan_names = [
            v.Name for v in view_elements
            if not v.IsTemplate and v.ViewType == ViewType.FloorPlan
        ]

        self._all_rooms = []
        for r in room_elements:
            # Skip unplaced rooms (area == 0 or no location)
            try:
                if r.Location is None:
                    continue
            except Exception:
                continue
            self._all_rooms.append(RoomItem(r, existing_plan_names))

        # Sort by number
        self._all_rooms.sort(key=lambda x: x.Number)
        self.room_datagrid.ItemsSource = to_items_source(self._all_rooms)
        if self._all_rooms:
            self.room_datagrid.SelectedItem = self._all_rooms[0]

    def _load_view_templates(self):
        """Floor Plan / Ceiling Plan / Elevation templates into their combo boxes."""
        maps = sg.template_maps(doc)
        self._plan_template_map = maps['floor']
        self._rcp_template_map = maps['ceiling']
        self._elev_template_map = maps['elevation']
        for combo, mapping in ((self.cmb_plan_template, self._plan_template_map),
                               (self.cmb_rcp_template, self._rcp_template_map),
                               (self.cmb_elev_template, self._elev_template_map)):
            combo.Items.Add("<None>")
            for name in sorted(mapping):
                combo.Items.Add(name)
            combo.SelectedIndex = 0

    def _load_plan_type_options(self):
        """Floor Plan / Ceiling Plan / Elevation view family types (None = missing)."""
        (self._floor_plan_type_id, self._ceiling_plan_type_id,
         self._elevation_type_id) = sg.find_view_types(doc)

    def _load_title_blocks(self):
        """Load all title block types into cmb_titleblock."""
        self._titleblock_map = sg.titleblock_map(doc)  # display_name -> ElementId

        for name in sorted(self._titleblock_map.keys()):
            self.cmb_titleblock.Items.Add(name)
        if self._titleblock_map:
            self.cmb_titleblock.SelectedIndex = 0
        else:
            self.chk_layout_on_sheet.IsEnabled = False
            self.chk_layout_on_sheet.ToolTip = "No title blocks found in project"

    # ── Helpers ───────────────────────────────────────
    def _get_selected_rooms(self):
        """Return list of RoomItems that are checked."""
        return [r for r in self._all_rooms if r.IsSelected]

    def _update_status(self):
        """Update status bar text."""
        selected = len(self._get_selected_rooms())
        total = len(self._all_rooms)
        self.status_count.Text = "{} rooms".format(total)
        if selected > 0:
            self.status_text.Text = "{} room(s) selected".format(selected)
        else:
            self.status_text.Text = "Ready"

    def _get_offset(self):
        """Crop offset typed in the project unit (or with a unit: 1 m,
        3'-0\") → feet. Unreadable text falls back to the 1 m default
        (Create validates it first and says what is wrong)."""
        try:
            return self._unit.parse(self.txt_offset.Text)
        except (ValueError, TypeError, AttributeError):
            return MILLIMETERS.to_feet(DEFAULT_OFFSET_MM)

    # ── Window chrome handlers ────────────────────────
    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, e):
        self.Close()

    def nav_toggle_clicked(self, sender, e):
        """Handle sidebar toggle button clicks mutually exclusively, switch tabs, and update title subtitle."""
        try:
            # Set clicked button to checked, and all others to unchecked
            is_rooms = (sender.Name == "nav_rooms")
            self.nav_rooms.IsChecked = is_rooms
            self.nav_layout.IsChecked = not is_rooms

            # Switch TabControl index
            if is_rooms:
                self.main_tab_control.SelectedIndex = 0
                self.lbl_subtitle.Text = "Generate plan views from selected rooms"
            else:
                self.main_tab_control.SelectedIndex = 1
                self.lbl_subtitle.Text = "Configure sheet layout and preview viewport placement"
                self._update_mockup()
        except Exception as ex:
            logger.error("Navigation click error: {}".format(ex))

    # ── Layout / Preview handlers ─────────────────────
    def layout_toggle_changed(self, sender, e):
        """Enable/disable layout sub-options based on toggle."""
        enabled = bool(self.chk_layout_on_sheet.IsChecked)
        self.pnl_layout_options.IsEnabled = enabled
        self.pnl_layout_options.Opacity = 1.0 if enabled else 0.45

    def room_selection_changed(self, sender, e):
        """Handle selected room change to update the realtime preview."""
        self._update_status()
        self._update_mockup()

    def room_row_clicked(self, sender, e):
        """Handle click on a Room List row to toggle its checkbox.

        Lets users click anywhere on a row to toggle IsSelected, in addition
        to using the checkbox itself (same UX as BatchOut's sheet list).
        """
        try:
            row = None
            element = e.OriginalSource
            while element is not None:
                if isinstance(element, (TextBox, CheckBox)):
                    # Click was on the checkbox or an editable cell - let it handle itself
                    return
                if isinstance(element, DataGridRow):
                    row = element
                    break
                element = VisualTreeHelper.GetParent(element)

            if row is not None:
                data_item = row.Item
                if data_item is not None:
                    data_item.IsSelected = not data_item.IsSelected
                    self.room_datagrid.Items.Refresh()
                    self._update_status()
                    self._update_mockup()
        except Exception as ex:
            logger.debug("Error handling room row click: {}".format(ex))

    def mockup_setting_changed(self, sender, e):
        """Handle view checkbox or ComboBox changes to redraw mockup."""
        self._update_mockup()

    def offset_changed(self, sender, e):
        """Handle offset textbox text changes to redraw mockup."""
        self._update_mockup()

    def _get_preview_room(self):
        """Get the room item to use for the layout preview."""
        try:
            row = self.room_datagrid.SelectedItem
            if row:
                return row
        except Exception:
            pass
        
        checked_rooms = self._get_selected_rooms()
        if checked_rooms:
            return checked_rooms[0]
            
        if self._all_rooms:
            return self._all_rooms[0]
            
        return None

    def _get_room_dimensions(self, room_item):
        """Get room bounding box width and depth in feet, with a fallback."""
        if room_item is None or room_item.Element is None:
            return 15.0, 12.0
        try:
            active_view = doc.ActiveView
            bbox = room_item.Element.get_BoundingBox(active_view)
            if bbox:
                w = bbox.Max.X - bbox.Min.X
                h = bbox.Max.Y - bbox.Min.Y
                if w > 0 and h > 0:
                    return w, h
        except Exception:
            pass
        return 15.0, 12.0

    def _get_template_scale(self, template_name):
        """Get the scale factor of a view template by name, defaulting to active view scale."""
        if not template_name or template_name == "<None>":
            try:
                return doc.ActiveView.Scale
            except Exception:
                return 50
        
        t_id = None
        if hasattr(self, '_plan_template_map') and template_name in self._plan_template_map:
            t_id = self._plan_template_map[template_name]
        elif hasattr(self, '_rcp_template_map') and template_name in self._rcp_template_map:
            t_id = self._rcp_template_map[template_name]
        elif hasattr(self, '_elev_template_map') and template_name in self._elev_template_map:
            t_id = self._elev_template_map[template_name]
            
        if t_id:
            try:
                template_view = doc.GetElement(t_id)
                if template_view:
                    return template_view.Scale
            except Exception:
                pass
        
        try:
            return doc.ActiveView.Scale
        except Exception:
            return 50

    def _get_titleblock_size(self, tb_id):
        """Real paper size (w, h) in feet for a title block type.

        SHEET_WIDTH/SHEET_HEIGHT only exist on placed instances (verified:
        the FJX/WH custom families expose nothing on the type), so probe by
        creating a temp sheet inside a rolled-back transaction. Cached per
        type so each title block is probed at most once per window.
        """
        if tb_id in self._tb_size_cache:
            return self._tb_size_cache[tb_id]
        size = None
        t = Transaction(doc, "T3Lab: Probe Title Block Size")
        try:
            t.Start()
            sheet = ViewSheet.Create(doc, tb_id)
            doc.Regenerate()
            inst = FilteredElementCollector(doc, sheet.Id) \
                .OfCategory(BuiltInCategory.OST_TitleBlocks) \
                .FirstElement()
            if inst:
                pw = inst.get_Parameter(DB.BuiltInParameter.SHEET_WIDTH)
                ph = inst.get_Parameter(DB.BuiltInParameter.SHEET_HEIGHT)
                if pw and ph and pw.AsDouble() > 0 and ph.AsDouble() > 0:
                    size = (pw.AsDouble(), ph.AsDouble())
                else:
                    bb = inst.get_BoundingBox(sheet)
                    if bb:
                        w = bb.Max.X - bb.Min.X
                        h = bb.Max.Y - bb.Min.Y
                        if w > 0 and h > 0:
                            size = (w, h)
        except Exception as ex:
            logger.debug("Title block size probe failed: {}".format(ex))
        finally:
            try:
                if t.HasStarted():
                    t.RollBack()
            except Exception:
                pass
        if not size:
            size = (2.759, 1.949)  # A1 fallback
            logger.debug("Title block size probe empty - falling back to A1")
        self._tb_size_cache[tb_id] = size
        return size

    def _get_preview_sheet_size(self):
        """Sheet dimensions (w, h) in feet of the selected title block."""
        try:
            tb_name = self.cmb_titleblock.SelectedItem
            tb_id = self._titleblock_map.get(tb_name) if tb_name else None
        except Exception:
            tb_id = None
        if tb_id:
            return self._get_titleblock_size(tb_id)
        return 2.759, 1.949  # A1 default before title blocks are loaded

    def _get_strip_config(self):
        """Title-block header strip: ('right'|'bottom'|'none', size in feet)."""
        side = 'right'
        try:
            idx = self.cmb_strip_side.SelectedIndex
            if 0 <= idx <= 2:
                side = ('right', 'bottom', 'none')[idx]
        except Exception:
            pass
        try:
            strip = self._paper.parse(self.txt_strip_mm.Text)
        except (ValueError, TypeError, AttributeError):
            strip = MILLIMETERS.to_feet(DEFAULT_STRIP_MM)
        if strip < 0:
            strip = 0.0
        return side, strip

    # Geometry shared with the real placement (Snippets/sheetgen.py).
    @staticmethod
    def _usable_rect(*args):
        return sg.usable_rect(*args)

    @staticmethod
    def _clamp_center(*args):
        return sg.clamp_center(*args)

    @staticmethod
    def _slot_centers(*args):
        return sg.slot_centers(*args)

    @staticmethod
    def _paper_label(w_ft, h_ft, paper=None):
        """Human label for a paper size, e.g. 'A1 (841 x 594 mm)'; the
        size is written in the paper unit (inches in an imperial project)."""
        paper = paper or MILLIMETERS
        size = "{} x {}".format(paper.text(w_ft), paper.show(h_ft))
        w_mm = int(round(MILLIMETERS.from_feet(w_ft)))   # ISO match, always mm
        h_mm = int(round(MILLIMETERS.from_feet(h_ft)))
        iso = ((1189, 841, "A0"), (841, 594, "A1"), (594, 420, "A2"),
               (420, 297, "A3"), (297, 210, "A4"))
        for iw, ih, name in iso:
            if abs(w_mm - iw) <= 6 and abs(h_mm - ih) <= 6:
                return "{}  ({})".format(name, size)
        return size

    def _draw_viewport(self, canvas, title, detail_num, w_px, h_px, x_px, y_px, bg_color, border_color):
        """Draw a viewport rectangle and Revit-style title mark on the WPF canvas."""
        # 1. Viewport Box
        box = Border()
        box.Width = w_px
        box.Height = h_px
        box.Background = SolidColorBrush(bg_color)
        box.BorderBrush = SolidColorBrush(border_color)
        box.BorderThickness = Thickness(1.5)
        box.CornerRadius = CornerRadius(4)
        Canvas.SetLeft(box, x_px)
        Canvas.SetTop(box, y_px)
        canvas.Children.Add(box)

        # 2. Revit Title Line
        line_y = y_px + h_px + 6
        line = Line()
        line.X1 = x_px
        line.Y1 = line_y
        line.X2 = x_px + max(w_px * 0.7, 40.0)
        line.Y2 = line_y
        line.Stroke = SolidColorBrush(Color.FromRgb(0x71, 0x71, 0x7A))
        line.StrokeThickness = 1
        canvas.Children.Add(line)

        # 3. Detail Circle
        circle = Border()
        circle.Width = 14
        circle.Height = 14
        circle.CornerRadius = CornerRadius(7)
        circle.BorderBrush = SolidColorBrush(border_color)
        circle.BorderThickness = Thickness(1)
        circle.Background = SolidColorBrush(Color.FromRgb(0xFF, 0xFF, 0xFF))
        
        circle_text = TextBlock()
        circle_text.Text = str(detail_num)
        circle_text.FontSize = 8
        circle_text.FontWeight = FontWeights.Bold
        circle_text.Foreground = SolidColorBrush(border_color)
        circle_text.HorizontalAlignment = HorizontalAlignment.Center
        circle_text.VerticalAlignment = VerticalAlignment.Center
        circle.Child = circle_text
        
        Canvas.SetLeft(circle, x_px)
        Canvas.SetTop(circle, line_y + 3)
        canvas.Children.Add(circle)

        # 4. View Title Text
        title_text = TextBlock()
        title_text.Text = title
        title_text.FontSize = 8.5
        title_text.FontWeight = FontWeights.Bold
        title_text.Foreground = SolidColorBrush(Color.FromRgb(0x18, 0x18, 0x1B))
        Canvas.SetLeft(title_text, x_px + 18)
        Canvas.SetTop(title_text, line_y + 2)
        canvas.Children.Add(title_text)

    def _setup_sheet_canvas(self, border, canvas, w_sheet, h_sheet,
                            max_w, max_h, strip_side, strip_ft):
        """Size the mockup sheet to the true paper aspect ratio, draw the
        header strip band and the paper-size label. Returns px-per-foot."""
        scale_px = min(max_w / w_sheet, max_h / h_sheet)
        border.Width = w_sheet * scale_px
        border.Height = h_sheet * scale_px

        if strip_side != 'none' and strip_ft > 0:
            band = Border()
            band.Background = SolidColorBrush(Color.FromRgb(0xF8, 0xFA, 0xFC))
            band.BorderBrush = SolidColorBrush(Color.FromRgb(0xCB, 0xD5, 0xE1))
            if strip_side == 'right':
                band.Width = strip_ft * scale_px
                band.Height = h_sheet * scale_px
                band.BorderThickness = Thickness(1, 0, 0, 0)
                Canvas.SetLeft(band, (w_sheet - strip_ft) * scale_px)
                Canvas.SetTop(band, 0)
            else:  # bottom
                band.Width = w_sheet * scale_px
                band.Height = strip_ft * scale_px
                band.BorderThickness = Thickness(0, 1, 0, 0)
                Canvas.SetLeft(band, 0)
                Canvas.SetTop(band, (h_sheet - strip_ft) * scale_px)
            canvas.Children.Add(band)

        lbl = TextBlock()
        lbl.Text = self._paper_label(w_sheet, h_sheet,
                                     getattr(self, '_paper', None))
        lbl.FontSize = 9
        lbl.FontWeight = FontWeights.Bold
        lbl.Foreground = SolidColorBrush(Color.FromRgb(0x94, 0xA3, 0xB8))
        Canvas.SetLeft(lbl, 6)
        Canvas.SetTop(lbl, 4)
        canvas.Children.Add(lbl)
        return scale_px

    def _draw_viewport_to_canvas(self, canvas, title, detail_num, w_paper, h_paper,
                                 cx, cy, h_sheet, scale_px, fits,
                                 bg_color, border_color):
        """Map sheet feet (origin bottom-left) to canvas px (origin top-left)
        and draw the viewport. Boxes that overflow the usable area are
        outlined in red."""
        if not fits:
            border_color = Color.FromRgb(0xEF, 0x44, 0x44)
        x_px = (cx - w_paper / 2.0) * scale_px
        y_px = (h_sheet - (cy + h_paper / 2.0)) * scale_px
        w_px = w_paper * scale_px
        h_px = h_paper * scale_px
        self._draw_viewport(canvas, title, detail_num, w_px, h_px, x_px, y_px, bg_color, border_color)

    def _update_mockup(self):
        """Update the real-time layout mockup based on current settings and selected room."""
        # _all_rooms is only set after LoadComponent returns - events fired
        # while the XAML is still parsing must not reach the drawing code.
        if not hasattr(self, "_all_rooms"):
            return
        if not hasattr(self, "combined_canvas") or self.combined_canvas is None:
            return
        
        # 1. Clear previous drawings
        self.combined_canvas.Children.Clear()
        self.plans_canvas.Children.Clear()
        self.elevations_canvas.Children.Clear()

        # 2. Get the room to preview
        room_item = self._get_preview_room()
        if not room_item:
            return

        w_room, h_room = self._get_room_dimensions(room_item)
        offset = self._get_offset()

        # 3. Check enabled view types and lookup templates
        do_floor = bool(self.chk_floor_plan.IsChecked)
        do_ceiling = bool(self.chk_ceiling_plan.IsChecked)
        do_elevations = bool(self.chk_elevations.IsChecked)

        plan_template = str(self.cmb_plan_template.SelectedItem) if self.cmb_plan_template.SelectedItem else "<None>"
        rcp_template = str(self.cmb_rcp_template.SelectedItem) if self.cmb_rcp_template.SelectedItem else "<None>"
        elev_template = str(self.cmb_elev_template.SelectedItem) if self.cmb_elev_template.SelectedItem else "<None>"

        plan_scale = self._get_template_scale(plan_template)
        rcp_scale = self._get_template_scale(rcp_template)
        elev_scale = self._get_template_scale(elev_template)

        # 4. Viewport sizes on paper (feet)
        w_floor_paper = (w_room + 2.0 * offset) / plan_scale
        h_floor_paper = (h_room + 2.0 * offset) / plan_scale

        w_rcp_paper = (w_room + 2.0 * offset) / rcp_scale
        h_rcp_paper = (h_room + 2.0 * offset) / rcp_scale

        h_elev_paper = 10.0 / elev_scale
        w_elev_s_n_paper = (w_room + 2.0 * offset) / elev_scale
        w_elev_w_e_paper = (h_room + 2.0 * offset) / elev_scale

        # 5. Real sheet size from the selected title block + usable area
        #    (same helpers the actual placement uses, so preview == result)
        w_sheet, h_sheet = self._get_preview_sheet_size()
        strip_side, strip_ft = self._get_strip_config()
        usable = self._usable_rect(0.0, 0.0, w_sheet, h_sheet,
                                   sg.SHEET_MARGIN_FT, strip_side, strip_ft)

        # Colors
        color_plan_bg = Color.FromRgb(0xEF, 0xF6, 0xFF)
        color_plan_border = Color.FromRgb(0x25, 0x63, 0xEB)

        color_rcp_bg = Color.FromRgb(0xF5, 0xF3, 0xFF)
        color_rcp_border = Color.FromRgb(0x7C, 0x3A, 0xED)

        color_elev_bg = Color.FromRgb(0xFF, 0xFB, 0xEB)
        color_elev_border = Color.FromRgb(0xD9, 0x77, 0x06)

        plan_views = []
        if do_floor:
            plan_views.append(("Floor Plan", w_floor_paper, h_floor_paper,
                               color_plan_bg, color_plan_border))
        if do_ceiling:
            plan_views.append(("Ceiling Plan", w_rcp_paper, h_rcp_paper,
                               color_rcp_bg, color_rcp_border))
        elev_specs = [
            ("Elevation 1", w_elev_s_n_paper, h_elev_paper),
            ("Elevation 2", w_elev_w_e_paper, h_elev_paper),
            ("Elevation 3", w_elev_s_n_paper, h_elev_paper),
            ("Elevation 4", w_elev_w_e_paper, h_elev_paper),
        ]

        # 6. Draw Combined Canvas (one sheet per room)
        if bool(self.rdo_layout_combined.IsChecked):
            scale_px = self._setup_sheet_canvas(
                self.CombinedSheetMockup, self.combined_canvas,
                w_sheet, h_sheet, 580.0, 380.0, strip_side, strip_ft)
            plan_centers, elev_centers = self._slot_centers(
                usable, len(plan_views), True)
            num = 1
            for spec, center in zip(plan_views, plan_centers):
                title, wp, hp, bg, bcol = spec
                cx, cy, fits = self._clamp_center(center[0], center[1], wp, hp, usable)
                self._draw_viewport_to_canvas(
                    self.combined_canvas, title, str(num), wp, hp, cx, cy,
                    h_sheet, scale_px, fits, bg, bcol)
                num += 1
            if do_elevations:
                for spec, center in zip(elev_specs, elev_centers):
                    title, wp, hp = spec
                    cx, cy, fits = self._clamp_center(center[0], center[1], wp, hp, usable)
                    self._draw_viewport_to_canvas(
                        self.combined_canvas, title, str(num), wp, hp, cx, cy,
                        h_sheet, scale_px, fits, color_elev_bg, color_elev_border)
                    num += 1

        # 7. Draw Separate Canvases (plans sheet + elevations sheet)
        else:
            if plan_views:
                scale_px = self._setup_sheet_canvas(
                    self.PlansSheetMockup, self.plans_canvas,
                    w_sheet, h_sheet, 300.0, 200.0, strip_side, strip_ft)
                plan_centers, _unused = self._slot_centers(
                    usable, len(plan_views), False)
                num = 1
                for spec, center in zip(plan_views, plan_centers):
                    title, wp, hp, bg, bcol = spec
                    cx, cy, fits = self._clamp_center(center[0], center[1], wp, hp, usable)
                    self._draw_viewport_to_canvas(
                        self.plans_canvas, title, str(num), wp, hp, cx, cy,
                        h_sheet, scale_px, fits, bg, bcol)
                    num += 1
            if do_elevations:
                scale_px = self._setup_sheet_canvas(
                    self.ElevationsSheetMockup, self.elevations_canvas,
                    w_sheet, h_sheet, 300.0, 200.0, strip_side, strip_ft)
                _unused, elev_centers = self._slot_centers(usable, 0, False)
                num = 1
                for spec, center in zip(elev_specs, elev_centers):
                    title, wp, hp = spec
                    cx, cy, fits = self._clamp_center(center[0], center[1], wp, hp, usable)
                    self._draw_viewport_to_canvas(
                        self.elevations_canvas, title, str(num), wp, hp, cx, cy,
                        h_sheet, scale_px, fits, color_elev_bg, color_elev_border)
                    num += 1

    def open_sheet_clicked(self, sender, e):
        """Activate the selected generated sheet in Revit."""
        try:
            selected_idx = self.cmb_generated_sheets.SelectedIndex
            if selected_idx >= 0 and selected_idx < len(self._generated_sheets):
                selected_sheet = self._generated_sheets[selected_idx]
                uidoc.ActiveView = selected_sheet
            else:
                TaskDialog.Show("Create Room Plan", "Please select a valid generated sheet from the list.")
        except Exception as ex:
            logger.error("Open sheet error: {}".format(ex))

    # ── Toolbar handlers ──────────────────────────────
    def select_all_clicked(self, sender, e):
        for r in self._all_rooms:
            r.IsSelected = True
        self.room_datagrid.Items.Refresh()
        self._update_status()
        # Giu checkbox select-all o header khop voi nut nay.
        self.sync_header_checkbox(
            self.FindName("chk_all_room_datagrid"), self.room_datagrid, "IsSelected")

    def select_none_clicked(self, sender, e):
        for r in self._all_rooms:
            r.IsSelected = False
        self.room_datagrid.Items.Refresh()
        self._update_status()
        # Giu checkbox select-all o header khop voi nut nay.
        self.sync_header_checkbox(
            self.FindName("chk_all_room_datagrid"), self.room_datagrid, "IsSelected")

    def search_changed(self, sender, e):
        """Filter room list by search text."""
        query = self.txt_search.Text.strip().upper()
        if not query:
            self.room_datagrid.ItemsSource = to_items_source(self._all_rooms)
        else:
            filtered = [
                r for r in self._all_rooms
                if query in r.Name.upper()
                or query in r.Number.upper()
                or query in (r.RoomType or "").upper()
                or query in (r.Level or "").upper()
            ]
            self.room_datagrid.ItemsSource = to_items_source(filtered)
        self._update_status()

    # ── Main action ───────────────────────────────────
    def create_plans_clicked(self, sender, e):
        if getattr(self, '_sheetgen_running', False):
            return
        self._sheetgen_running = True
        try:
            self._create_plans(sender, e)
        except (Exception, sg.SheetGenPending) as ex:
            logger.exception("SheetGen creation interrupted: {}".format(ex))
            TaskDialog.Show("Create Room Plan",
                            "Creation interrupted. Previously committed results remain.\n{}".format(ex))
        finally:
            self._sheetgen_running = False
            self.end_progress()

    def _generator(self, options):
        """The Revit side of a run (Snippets/sheetgen.py, shared with the MCP tool)."""
        return sg.RoomSheetGen(doc, options, logger)

    def _create_plans(self, sender, e):
        """Create plan views (and optionally lay them out on sheets) for selected rooms."""
        selected_rooms = self._get_selected_rooms()
        if not selected_rooms:
            TaskDialog.Show("Create Room Plan", "Please select at least one room.")
            return

        do_floor      = bool(self.chk_floor_plan.IsChecked)
        do_ceiling    = bool(self.chk_ceiling_plan.IsChecked)
        do_elevations = bool(self.chk_elevations.IsChecked)
        do_layout     = bool(self.chk_layout_on_sheet.IsChecked)

        if not do_floor and not do_ceiling and not do_elevations:
            TaskDialog.Show("Create Room Plan", "Please select at least one view type.")
            return

        missing = []
        for enabled, type_id, label in (
                (do_floor, self._floor_plan_type_id, "Floor Plan"),
                (do_ceiling, self._ceiling_plan_type_id, "Ceiling Plan"),
                (do_elevations, self._elevation_type_id, "Elevation")):
            if enabled and type_id is None:
                missing.append(label)
        tb_name = self.cmb_titleblock.SelectedItem
        tb_id = self._titleblock_map.get(tb_name) if tb_name else None
        if missing or (do_layout and tb_id is None):
            TaskDialog.Show("Create Room Plan", "Missing required types: " +
                            ", ".join(missing + (["Title Block"] if do_layout and tb_id is None else [])))
            return
        try:
            try:
                offset_ft = self._unit.parse(self.txt_offset.Text)
            except ValueError as ex:
                raise ValueError("Crop offset: {}".format(ex))
            if not math.isfinite(offset_ft) or offset_ft < 0:
                raise ValueError("Crop offset must be zero or more.")
            if do_floor:
                for item in selected_rooms:
                    qty = int(item.GenQty)
                    if qty < 0 or float(item.GenQty) != qty:
                        raise ValueError("Floor plan quantity must be a non-negative integer.")
        except (ValueError, TypeError, OverflowError) as ex:
            TaskDialog.Show("Create Room Plan", "Invalid creation settings: {}".format(ex))
            return

        def template_id(combo, mapping):
            name = combo.SelectedItem
            return mapping.get(name) if name != "<None>" else None

        strip_side, strip_ft = self._get_strip_config()
        gen = self._gen = self._generator({
            'floor': do_floor, 'ceiling': do_ceiling, 'elevations': do_elevations,
            'floor_type_id': self._floor_plan_type_id,
            'ceiling_type_id': self._ceiling_plan_type_id,
            'elevation_type_id': self._elevation_type_id,
            'floor_template_id': template_id(self.cmb_plan_template, self._plan_template_map),
            'ceiling_template_id': template_id(self.cmb_rcp_template, self._rcp_template_map),
            'elevation_template_id': template_id(self.cmb_elev_template, self._elev_template_map),
            'offset_ft': self._get_offset(),
            'cropbox_visible': bool(self.chk_cropbox_visible.IsChecked),
            'strip_side': strip_side, 'strip_ft': strip_ft,
        })
        gen.seed_names()

        room_results = []
        self.begin_progress(len(selected_rooms), disable=[sender])
        for _sg_idx, room_item in enumerate(selected_rooms):
            if not self.step_progress(_sg_idx, "Processing room {}/{}...".format(_sg_idx + 1, len(selected_rooms))):
                break
            result = gen.create_room(room_item)
            if result is not None:
                room_results.append(result)
        created_count = gen.created
        error_count = gen.errors

        # ── Sheet Layout ────────────────────────────────
        sheets_created = 0
        self._generated_sheets = []
        self.cmb_generated_sheets.Items.Clear()
        self.cmb_generated_sheets.IsEnabled = False
        self.btn_open_sheet.IsEnabled = False
        if do_layout and room_results and not self.is_cancelled:
            combined = bool(self.rdo_layout_combined.IsChecked)
            for layout_index, result in enumerate(room_results):
                if not self.step_progress(layout_index, "Laying out room {}/{}...".format(layout_index + 1, len(room_results))):
                    break
                try:
                    sheets = self._layout_views_on_sheets(result, tb_id, combined)
                    sheets_created += len(sheets)
                    self._generated_sheets.extend(sheets)
                except Exception as ex:
                    error_count += 1
                    logger.error("Sheet layout error for {}: {}".format(
                        result['room_item'].Name, ex))

            if self._generated_sheets:
                self.cmb_generated_sheets.Items.Clear()
                for sh in self._generated_sheets:
                    sheet_num = sh.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER)
                    num_str = sheet_num.AsString() if sheet_num else "???"
                    self.cmb_generated_sheets.Items.Add("{} - {}".format(num_str, sh.Name))
                self.cmb_generated_sheets.SelectedIndex = 0
                self.cmb_generated_sheets.IsEnabled = True
                self.btn_open_sheet.IsEnabled = True

        # ── Result ─────────────────────────────────────
        if self.is_cancelled:
            msg = "Cancelled — {} view(s) created.".format(created_count)
        else:
            msg = "{} view(s) created.".format(created_count)
        if sheets_created:
            msg += "\n{} sheet(s) created.".format(sheets_created)
        if error_count > 0:
            msg += "\n{} error(s) — see output for details.".format(error_count)

        self.status_text.Text = msg
        TaskDialog.Show("Create Room Plan", msg)
        if not do_layout:
            self.Close()          # with a layout the window stays for Open Selected Sheet

    def _layout_views_on_sheets(self, result, titleblock_id, combined):
        return self._gen.layout(result, titleblock_id, combined)

    # ── Select-all o header cot checkbox ────────────────────────────────
    # toggle_all_rows() nam trong T3WPFWindow: no chay tren grid.Items nen chi
    # dong dang hien thi (sau filter/sort) bi doi, dung nhu nguoi dung thay.

    def select_all_room_datagrid_clicked(self, sender, e):
        """Header checkbox: chon/bo chon moi dong dang hien thi cua room_datagrid."""
        self.toggle_all_rows(self.room_datagrid, "IsSelected", sender.IsChecked)


# ╔╦╗╔═╗╦╔╗╔
# ║║║╠═╣║║║║
# ╩ ╩╩ ╩╩╝╚╝ MAIN
# ==================================================
def show_sheet_gen_dialog(doc=None, uidoc=None):
    d = doc
    if d is None:
        try:
            d = revit.doc
        except Exception:
            d = None
    if d is None:
        forms.alert("No active Revit document found.", title="SheetGen")
        return
    u = uidoc
    if u is None:
        try:
            u = revit.uidoc
        except Exception:
            u = None
    try:
        window = CreateRoomPlanWindow(d, u)
        window.ShowDialog()
    except Exception as ex:
        logger.error("SheetGen error: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())

if __name__ == '__main__':
    show_sheet_gen_dialog()
