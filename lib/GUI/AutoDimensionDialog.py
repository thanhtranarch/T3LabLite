# -*- coding: utf-8 -*-
"""
Auto Dimension

Auto-dimension walls, structural columns, architectural columns,
and grids in the current view.
T3Lab tool.

--------------------------------------------------------
Author: Tran Tien Thanh
--------------------------------------------------------
"""

__author__  = "Tran Tien Thanh"
__title__   = "Auto Dimension"
__version__ = "2.2.2"

# IMPORT LIBRARIES
# ==================================================
import os
import sys
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import WindowState, Visibility
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInParameter,
    DimensionType,
    DimensionStyleType,
    View,
    ViewType,
)
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI import TaskDialog
from pyrevit import revit, forms, script
from GUI.WPF_Base import T3WPFWindow
from Snippets._units import project_length_unit, paper_unit

# PATH SETUP
# ==================================================
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR    = os.path.dirname(os.path.dirname(SCRIPT_DIR))
LIB_DIR    = SCRIPT_DIR
XAML_FILE  = os.path.join(SCRIPT_DIR, 'Tools', 'AutoDimension.xaml')


if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

# DEFINE VARIABLES
# ==================================================
logger = script.get_logger()

# `revit.doc` / `revit.uidoc` RAISE AttributeError (not return None) when no
# UIDocument is active — pyrevit/revit/__init__.py __getattr__. At module scope
# that killed the whole import before show_dialog() could report anything, so
# resolve defensively here; show_dialog() does the real check and explains.
try:
    doc = revit.doc
except Exception:
    doc = None
try:
    uidoc = revit.uidoc
except Exception:
    uidoc = None

# CONSTANTS
# ==================================================
MM_TO_FEET  = 1.0 / 304.8


# Length fields: TextBox → (unit label beside it, default in mm, name in
# messages). Every one is a MODEL distance — how far the dimension line sits
# from the elements, a segment length, the facade search band — so they are
# typed and shown in the project's length unit. Only the "Offsets from Scale"
# sizes are paper distances (see offsets_from_scale).
LENGTH_FIELDS = (
    ('txt_offset',        'lbl_offset_unit',        1000, 'Offset'),
    ('txt_l1',            'lbl_l1_unit',             500, 'L1 offset'),
    ('txt_l2',            'lbl_l2_unit',            1000, 'L2 offset'),
    ('txt_l3',            'lbl_l3_unit',            1500, 'L3 offset'),
    ('txt_min_seg',       'lbl_min_seg_unit',        300, 'Minimum segment'),
    ('txt_facade_offset', 'lbl_facade_offset_unit', 1200, 'Facade offset'),
    ('txt_facade_tol',    'lbl_facade_tol_unit',    1800, 'Perimeter tolerance'),
)

# "Offsets from Scale": L1/L2/L3 = these sizes ON PAPER × the view scale, then
# rounded in the model and never below a minimum. In the paper unit of the
# project (paper_unit): mm in a metric project, inches in an imperial one.
#           paper sizes (L1, L2, L3)  rounding step  minimum (L1, L2, L3)
SCALE_OFFSETS = {
    True:  ((6.0, 12.0, 18.0),        50.0,          (400.0, 800.0, 1200.0)),
    False: ((0.25, 0.5, 0.75),        1.0,           (16.0, 32.0, 48.0)),
}


def offsets_from_scale(scale, unit):
    """(L1, L2, L3) in internal feet for a view at 1:scale. Metric: 6 / 12 /
    18 mm on paper rounded to 50 mm (1:100 → 600 / 1200 / 1800 mm). Imperial:
    1/4 / 1/2 / 3/4 in on paper rounded to 1 in (1/8" = 1'-0" → 2' / 4' / 6')."""
    paper = paper_unit(unit)
    sizes, step, minimum = SCALE_OFFSETS[paper.is_metric]
    scale = float(scale or 100)
    return tuple(paper.to_feet(max(low, round(size * scale / step) * step))
                 for size, low in zip(sizes, minimum))


# Revit side (helpers, per-view dimensioning): Snippets/_auto_dimension.py,
# shared with the auto_dimension AI tool.
from Snippets._auto_dimension import (    # noqa: E402
    _is_valid_view, element_id_int, dim_one_view, dim_section_view)


# WINDOW CLASS
# ==================================================

class AutoDimensionWindow(T3WPFWindow):
    """
    WPF window for the Auto Dimension tool.
    Inherits from pyrevit forms.WPFWindow which handles XAML loading.
    """

    def __init__(self, uidoc_ref, doc_ref):
        T3WPFWindow.__init__(self, XAML_FILE)
        self.uidoc = uidoc_ref
        self.doc   = doc_ref
        self._dim_types = []  # list of DimensionType elements
        self._views = []      # list of plan View elements
        # Read per window: the module stays loaded across projects.
        self._unit = project_length_unit(self.doc)

        self._apply_units()
        self._populate_dim_types()
        self._populate_views()
        self._set_status("Ready")


    # ── Project length unit ──────────────────────────────────────────────

    def _apply_units(self):
        """Unit labels, defaults and the Offsets-from-Scale tooltip in the
        length unit the project displays (LENGTH_FIELDS)."""
        for box, label, default_mm, _name in LENGTH_FIELDS:
            getattr(self, label).Text = self._unit.tag
            getattr(self, box).Text = self._unit.default_text(default_mm)
        paper = paper_unit(self._unit)
        sizes = SCALE_OFFSETS[paper.is_metric][0]
        self.btn_offsets_from_scale.ToolTip = (
            u"Set L1/L2/L3 to {} on paper at the active view scale".format(
                u" / ".join(paper.show(paper.to_feet(v)) for v in sizes)))

    def _read_length(self, box):
        """Text of one LENGTH_FIELDS box → internal feet. A bare number is in
        the project unit; "600 mm" or "2'-0\"" work anywhere. Raises
        ValueError naming the field."""
        name = next(n for b, _l, _d, n in LENGTH_FIELDS if b == box)
        try:
            return self._unit.parse(getattr(self, box).Text)
        except ValueError as ex:
            raise ValueError(u"{}: {}".format(name, ex))

    def on_offsets_from_scale_clicked(self, sender, args):
        """L1/L2/L3 theo tỉ lệ view: 6 / 12 / 18 mm giấy (dự án imperial:
        1/4 / 1/2 / 3/4 in), làm tròn 50 mm (1 in) thực — offsets_from_scale.

        Công thức tất định — cùng tỉ lệ luôn ra cùng khoảng cách (trước đây là
        nút "AI Auto-Offsets", AI chỉ đoán lại chính mấy con số này).
        """
        try:
            scale = 100
            try:
                if self.uidoc and self.uidoc.ActiveView:
                    scale = self.uidoc.ActiveView.Scale or 100
            except Exception:
                pass
            self._prev_offsets_backup = {
                'mode': getattr(self.cmb_offset_mode, 'SelectedIndex', 0),
                'l1': getattr(self.txt_l1, 'Text', ''),
                'l2': getattr(self.txt_l2, 'Text', ''),
                'l3': getattr(self.txt_l3, 'Text', ''),
                'offset': getattr(self.txt_offset, 'Text', ''),
            }
            l1, l2, l3 = offsets_from_scale(scale, self._unit)
            self.cmb_offset_mode.SelectedIndex = 1      # 3-Level auto
            self.txt_l1.Text = self._unit.text(l1)
            self.txt_l2.Text = self._unit.text(l2)
            self.txt_l3.Text = self._unit.text(l3)
            self.txt_offset.Text = self._unit.text(l2)
            self.btn_offsets_undo.Visibility = Visibility.Visible
            self._set_status(u"Offsets for 1:{}: L1={}, L2={}, L3={}.".format(
                scale, self._unit.show(l1), self._unit.show(l2),
                self._unit.show(l3)))
        except Exception as ex:
            logger.error("Offsets from scale failed: {}".format(ex))
            self._set_status("Could not compute offsets: {}".format(ex))

    def on_offsets_undo_clicked(self, sender, args):
        """Trả offset về giá trị trước lần bấm Offsets from Scale."""
        try:
            bak = getattr(self, '_prev_offsets_backup', None)
            if not bak:
                return
            if hasattr(self, 'cmb_offset_mode'):
                self.cmb_offset_mode.SelectedIndex = bak.get('mode', 0)
            if hasattr(self, 'txt_l1'):
                self.txt_l1.Text = bak.get('l1', '')
            if hasattr(self, 'txt_l2'):
                self.txt_l2.Text = bak.get('l2', '')
            if hasattr(self, 'txt_l3'):
                self.txt_l3.Text = bak.get('l3', '')
            if hasattr(self, 'txt_offset'):
                self.txt_offset.Text = bak.get('offset', '')

            self._prev_offsets_backup = None
            undo_btn = getattr(self, 'btn_offsets_undo', None)
            if undo_btn:
                undo_btn.Visibility = Visibility.Collapsed
            self._set_status("Reverted dimension offsets to previous values.")
        except Exception as ex:
            logger.error("Undo offsets failed: {}".format(ex))


    # ── Status ───────────────────────────────────────────────────────────

    def _set_status(self, msg):
        try:
            self.status_text.Text = msg
        except Exception:
            pass

    # ── Populate DimensionType combobox ──────────────────────────────────

    def _populate_dim_types(self):
        try:
            all_types = FilteredElementCollector(self.doc)\
                .OfClass(DimensionType)\
                .WhereElementIsElementType()\
                .ToElements()

            self._dim_types = []
            for dt in all_types:
                try:
                    if dt.StyleType == DimensionStyleType.Linear:
                        self._dim_types.append(dt)
                except Exception:
                    pass

            self.cmb_dim_type.Items.Clear()
            try:
                self.cmb_check_dim_type.Items.Clear()
            except Exception:
                pass
            for dt in self._dim_types:
                try:
                    param = dt.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
                    name = param.AsString() if param and param.AsString() else dt.Name
                except Exception:
                    name = "Type {}".format(element_id_int(dt.Id))
                self.cmb_dim_type.Items.Add(name)
                try:
                    self.cmb_check_dim_type.Items.Add(name)
                except Exception:
                    pass

            if self._dim_types:
                self.cmb_dim_type.SelectedIndex = 0
                try:
                    self.cmb_check_dim_type.SelectedIndex = 0
                except Exception:
                    pass
                self._set_status("Loaded {} linear dimension type(s).".format(
                    len(self._dim_types)))
            else:
                self._set_status("No linear dimension types found in document.")
        except Exception as ex:
            logger.warning("Could not load dimension types: {}".format(ex))

    def _populate_views(self):
        """Populate the plan view list box with all non-template plan views."""
        try:
            plan_types = (
                ViewType.FloorPlan,
                ViewType.EngineeringPlan,
                ViewType.AreaPlan,
                ViewType.CeilingPlan,
            )
            all_views = list(
                FilteredElementCollector(self.doc)
                .OfClass(View)
                .ToElements()
            )
            self._views = sorted(
                [v for v in all_views
                 if not v.IsTemplate and v.ViewType in plan_types],
                key=lambda v: v.Name
            )
            self.lst_views.Items.Clear()
            for v in self._views:
                self.lst_views.Items.Add(v.Name)
        except Exception as ex:
            logger.warning("Could not load plan views: {}".format(ex))

    def select_all_views_clicked(self, sender, args):
        self.lst_views.SelectAll()

    def clear_views_clicked(self, sender, args):
        self.lst_views.UnselectAll()

    def specific_views_changed(self, sender, args):
        """Toggle the plan-view picker. When off, the tool dimensions the
        active view; clear any lingering selection so the run falls back to it."""
        try:
            is_specific = self.chk_specific_views.IsChecked == True
            self.pnl_views_list.Visibility = (
                Visibility.Visible if is_specific else Visibility.Collapsed)
            if not is_specific:
                self.lst_views.UnselectAll()
        except Exception:
            pass

    # ── Offset mode toggle ────────────────────────────────────────────────

    def offset_mode_changed(self, sender, args):
        if not hasattr(self, 'pnl_single_offset'):
            return
        is_3level = self.cmb_offset_mode.SelectedIndex == 1
        self.pnl_single_offset.Visibility = (
            Visibility.Collapsed if is_3level else Visibility.Visible)
        self.pnl_3level_offsets.Visibility = (
            Visibility.Visible if is_3level else Visibility.Collapsed)

    def checking_mode_changed(self, sender, args):
        is_check = self.chk_checking_mode.IsChecked == True
        show = Visibility.Visible
        hide = Visibility.Collapsed
        try:
            self.pnl_check_dim_type.Visibility = show if is_check else hide
        except Exception:
            pass
        for name in ('grp_wall_layer', 'grp_dim_type', 'grp_dim_settings', 'brd_info'):
            try:
                getattr(self, name).Visibility = hide if is_check else show
            except Exception:
                pass

    # ── Window chrome event handlers ──────────────────────────────────────

    def minimize_button_clicked(self, sender, args):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, args):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, args):
        self.Close()

    # ── Run handler ───────────────────────────────────────────────────────

    def run_clicked(self, sender, args):
        try:
            self._run_auto_dimension()
        except Exception as ex:
            msg = "Unexpected error: {}".format(ex)
            logger.error(msg)
            TaskDialog.Show("Auto Dimension — Error", msg)
            self._set_status("Error.")

    # ── Core logic ───────────────────────────────────────────────────────

    def _run_auto_dimension(self):
        # ── Read UI ────────────────────────────────────────────────────────
        do_walls       = self.chk_walls.IsChecked == True
        do_struct_cols = self.chk_struct_columns.IsChecked == True
        do_arch_cols   = self.chk_arch_columns.IsChecked == True
        do_grids       = self.chk_grids.IsChecked == True   # grid-to-grid overall chains
        do_windows     = self.chk_windows.IsChecked == True
        do_doors       = self.chk_doors.IsChecked == True
        do_lifts       = self.chk_lifts.IsChecked == True

        # Wall reference layer: 'ext' | 'int' | 'both'
        wall_mode = 'ext'
        try:
            if self.rad_wall_int.IsChecked == True:
                wall_mode = 'int'
            elif self.rad_wall_both.IsChecked == True:
                wall_mode = 'both'
        except Exception:
            pass

        sel_idx = self.cmb_dim_type.SelectedIndex
        if sel_idx < 0 or sel_idx >= len(self._dim_types):
            TaskDialog.Show("Auto Dimension", "Please select a Dimension Type.")
            return
        dim_type = self._dim_types[sel_idx]

        # Checking mode: align to nearest grid only, use check dim type
        check_mode = False
        try:
            check_mode = self.chk_checking_mode.IsChecked == True
        except Exception:
            pass
        if check_mode:
            try:
                ck_idx = self.cmb_check_dim_type.SelectedIndex
                if 0 <= ck_idx < len(self._dim_types):
                    dim_type = self._dim_types[ck_idx]
                elif self._dim_types:
                    dim_type = self._dim_types[0]
            except Exception:
                pass
        ck_off = 300.0 * MM_TO_FEET  # fixed offset for checking mode dims

        # ── Offset levels ──────────────────────────────────────────────────
        is_3level = False
        try:
            is_3level = self.cmb_offset_mode.SelectedIndex == 1
        except Exception:
            pass

        # ── Min-segment warning & facade switches ─────────────────────────
        check_min = False
        try:
            check_min = self.chk_min_seg.IsChecked == True
        except Exception:
            pass
        do_facade = False
        try:
            do_facade = self.chk_facade.IsChecked == True
        except Exception:
            pass

        # ── Lengths typed by the user (project unit — LENGTH_FIELDS) ─────
        # A box that is not a length stops the run with the reason; only
        # boxes the run uses are read (hidden L1-L3 / facade boxes are not).
        min_seg_feet = 300.0 * MM_TO_FEET   # used only when check_min
        facade_off = 1200.0 * MM_TO_FEET    # used only when do_facade
        facade_tol = 1800.0 * MM_TO_FEET
        try:
            if is_3level:
                l1_feet = self._read_length('txt_l1')
                l2_feet = self._read_length('txt_l2')
                l3_feet = self._read_length('txt_l3')
            else:
                base = self._read_length('txt_offset')
                l1_feet = base * 0.5
                l2_feet = base
                l3_feet = base * 1.5
            if check_min:
                min_seg_feet = self._read_length('txt_min_seg')
            if do_facade:
                facade_off = self._read_length('txt_facade_offset')
                facade_tol = self._read_length('txt_facade_tol')
        except ValueError as ex:
            self._set_status(u"{}".format(ex))
            TaskDialog.Show("Auto Dimension", u"{}".format(ex))
            return

        # ── Direction & placement ─────────────────────────────────────────
        dir_idx = 0
        try:
            dir_idx = self.cmb_direction.SelectedIndex
        except Exception:
            pass
        run_x = dir_idx != 2   # False only when "Y only"
        run_y = dir_idx != 1   # False only when "X only"

        both_sides = False
        if not check_mode:
            try:
                both_sides = self.chk_both_sides.IsChecked == True
            except Exception:
                pass

        # ── Collect views to process ───────────────────────────────────────
        views_to_dim = []
        try:
            idx_list = list(self.lst_views.SelectedItems)
            for name in idx_list:
                for v in self._views:
                    if v.Name == name:
                        views_to_dim.append(v)
                        break
        except Exception:
            pass

        if not views_to_dim:
            active = self.uidoc.ActiveView
            if not _is_valid_view(active):
                TaskDialog.Show("Auto Dimension",
                                "Please activate a Plan, Section, Elevation, or Detail view.")
                return
            views_to_dim = [active]
        else:
            views_to_dim = [v for v in views_to_dim if _is_valid_view(v)]
            if not views_to_dim:
                TaskDialog.Show("Auto Dimension",
                                "None of the selected views are valid for dimensioning.\n"
                                "Please select Floor Plan or Engineering Plan views.")
                return

        # margin is recomputed inside _dim_one_view; pass a placeholder derived from l1_feet
        margin = l1_feet * 0.5

        # ── Run per-view and accumulate results ────────────────────────────
        total_dims_created = 0
        total_attempts = [0]
        total_failures = []
        all_created_dims = []

        for view in views_to_dim:
            n, att, fails, cdims = self._dim_one_view(
                view, do_walls, do_struct_cols, do_arch_cols, do_grids,
                do_windows, do_doors, do_lifts, do_facade,
                wall_mode, dim_type, check_mode, ck_off, is_3level,
                l1_feet, l2_feet, l3_feet, margin,
                run_x, run_y, both_sides, check_min, min_seg_feet,
                facade_off, facade_tol
            )
            total_dims_created += n
            total_attempts[0] += att
            total_failures.extend(fails)
            all_created_dims.extend(cdims)

        dims_created = total_dims_created
        dim_attempts = total_attempts
        dim_failures = total_failures[:5]
        created_dims = all_created_dims

        # ── Post-dim conflict detection ────────────────────────────────────
        small_count = 0
        if check_min and created_dims:
            for dim_elem in created_dims:
                try:
                    for seg in dim_elem.Segments:
                        try:
                            val = seg.Value
                            if val is not None and val < min_seg_feet:
                                small_count += 1
                        except Exception:
                            pass
                except Exception:
                    pass

        # Build element summary for diagnostic message
        elem_summary = (
            "Views processed: {}\n"
            "Dim attempts (>=2 refs): {}\n"
            "NewDimension failures: {}"
        ).format(
            len(views_to_dim),
            dim_attempts[0],
            len(dim_failures)
        )

        if dims_created == 0:
            failure_detail = ""
            if dim_failures:
                failure_detail = "\n\nFirst error:\n" + dim_failures[0]
            elif dim_attempts[0] == 0:
                failure_detail = (
                    "\n\nNo dimension was attempted — all ref_pos lists had < 2 refs.\n"
                    "Likely cause: grid references could not be obtained."
                )
            self._set_status("Done — 0 dimensions. See dialog for details.")
            TaskDialog.Show(
                "Auto Dimension — Diagnostic",
                "No dimensions were created across {} view(s).\n\n"
                "{}{}\n\n"
                "Tip: Check Revit journal for 'Chain dim failed' warnings.".format(
                    len(views_to_dim), elem_summary, failure_detail)
            )
        elif small_count > 0:
            warn = u"  {} segment(s) < {} — text may overlap.".format(
                small_count, self._unit.show(min_seg_feet))
            self._set_status("Done ({} dims). {}".format(dims_created, warn))
            TaskDialog.Show(
                "Auto Dimension — Conflict Warning",
                "Created {} dimension string(s) across {} view(s).\n\n"
                "{} segment(s) are shorter than {}.\n"
                "Consider increasing the offset or adjusting element positions "
                "to prevent text overlap.".format(
                    dims_created, len(views_to_dim),
                    small_count, self._unit.show(min_seg_feet))
            )
        else:
            self._set_status("Done — {} dimension string(s) created.".format(dims_created))
            TaskDialog.Show(
                "Auto Dimension",
                "Done.\nCreated {} dimension string(s) across {} view(s).".format(
                    dims_created, len(views_to_dim))
            )

    def _dim_one_view(self, view, do_walls, do_struct_cols, do_arch_cols, do_grids,
                      do_windows, do_doors, do_lifts, do_facade,
                      wall_mode, dim_type, check_mode, ck_off, is_3level,
                      l1_feet, l2_feet, l3_feet, margin,
                      run_x, run_y, both_sides, check_min, min_seg_feet,
                      facade_off, facade_tol):
        """All dimensioning phases on one view (Snippets/_auto_dimension.dim_one_view).
        Returns (dims_created, attempts, failures_list, created_dims)."""
        return dim_one_view(self.doc, view, do_walls, do_struct_cols, do_arch_cols, do_grids,
                            do_windows, do_doors, do_lifts, do_facade,
                            wall_mode, dim_type, check_mode, ck_off, is_3level,
                            l1_feet, l2_feet, l3_feet, margin,
                            run_x, run_y, both_sides, check_min, min_seg_feet,
                            facade_off, facade_tol, on_error=self._transaction_error)

    def _dim_section_view(self, view, do_grids, dim_type, l3_feet):
        return dim_section_view(self.doc, view, do_grids, dim_type, l3_feet,
                                on_error=self._transaction_error)

    @staticmethod
    def _transaction_error(msg):
        logger.error(msg)
        TaskDialog.Show("Auto Dimension — Error", msg)



# MAIN SCRIPT
# ==================================================
def show_dialog():
    # NOTE: `pyrevit.forms.<anything>` raises PyRevitCPythonNotSupported under the
    # CPython engine — its __getattr__ refuses every attribute. GUI.T3Dialog is the
    # T3 replacement and works on both engines.
    from GUI.T3Dialog import show_warning
    from Snippets._host import resolve_doc, host_uiapp

    doc_inst, doc_error = resolve_doc()
    if doc_inst is None:
        show_warning(doc_error or "Please open a Revit document first.",
                     title="Auto Dimension")
        return

    uiapp = host_uiapp()
    uidoc_inst = getattr(uiapp, 'ActiveUIDocument', None) if uiapp else None
    if uidoc_inst is None:
        show_warning("No active view to dimension in. Open a plan or section "
                     "view in your model, then run Auto Dimension again.",
                     title="Auto Dimension")
        return
    doc_inst = uidoc_inst.Document

    window = AutoDimensionWindow(uidoc_inst, doc_inst)
    # Modal (ShowDialog) is required: Run starts a Transaction directly in the
    # button handler, which is only legal while the command is still inside
    # Revit API context. A modeless window drops that context and
    # Transaction.Start/NewDimension throw "outside of API context"
    # (TagCheckerDialog reverted to modal for the same reason). View selection
    # happens in-dialog, so viewport interaction is not needed while open.
    window.show(modal=True)

if __name__ == '__main__':
    show_dialog()
