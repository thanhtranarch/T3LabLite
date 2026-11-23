#! python3
# -*- coding: utf-8 -*-
"""
Point Cloud to Model
--------------------
Scan-to-BIM wizard: extract point cloud data, detect the architectural
and structural shell — Walls, Floors, Ceilings, Doors, Windows, Columns,
Stairs, Roof — then create all elements in a single TransactionGroup.
MEP content (ducts, pipes, equipment) is deliberately OUT of scope:
walls require tall evidence below the ceiling (rejects shelving/racks/
low MEP runs) and columns must be near-square, >= 200 mm and rise above
furniture height (rejects pipes, flat ducts and floor-mounted units).

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__title__   = "Point Cloud\nto Model"
__author__  = "Tran Tien Thanh"
__version__  = "1.0.0"

# IMPORT LIBRARIES
# ==============================================================================
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
import math

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from System import Action
from System.Windows import WindowState, Visibility
from System.Windows.Threading import DispatcherPriority
from System.Collections.Generic import List

from pyrevit import revit, DB
from Autodesk.Revit.DB import (
    Transaction,
    TransactionGroup,
    FilteredElementCollector,
    BuiltInCategory,
    BuiltInParameter,
    WallType,
    Wall,
    Level,
    Grid,
    FloorType,
    Floor,
    XYZ,
    Plane,
    Line,
    CurveLoop,
    CurveArray,
    BoundingBoxXYZ,
    PointCloudInstance,
    IFailuresPreprocessor,
    FailureProcessingResult,
    FailureSeverity,
    ElementId,
    FamilySymbol,
    DetailLine,
    ElementTransformUtils
)
try:
    from Autodesk.Revit.DB import StructuralType
except ImportError:
    from Autodesk.Revit.DB.Structure import StructuralType

from Autodesk.Revit.DB.PointClouds import PointCloudFilterFactory
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter, PickBoxStyle
from Autodesk.Revit.Exceptions import OperationCanceledException
from pyrevit import forms, script
from GUI.WPF_Base import T3WPFWindow, to_items_source

# Path setup
# ==============================================================================
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR    = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))))
lib_dir    = os.path.join(EXT_DIR, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

XAML_FILE  = os.path.join(os.path.dirname(__file__), 'Tools', 'PointCloud.xaml')

# DEFINE VARIABLES
# ==============================================================================
logger        = script.get_logger()
# Read from the Application, not the document: `revit.doc` is None when this
# tool is launched from the Assistant pane or with no project open, and the
# old `int(revit.doc.Application.VersionNumber)` raised at IMPORT time
# ('NoneType' object has no attribute 'Application') so the window never opened.
from Snippets._host import get_revit_version, resolve_doc, resolve_uidoc, host_uiapp
REVIT_VERSION = get_revit_version()

# Same null-document trap, one layer deeper: `revit.doc` is None from the
# Assistant pane, and ElementBuilder runs a FilteredElementCollector in
# PointCloudModelWindow.__init__ — a bare `revit.doc` took the window down
# before it ever rendered. resolve_doc() is the repo-wide fallback chain and
# hands back an actionable message instead of a null-document crash.
doc, DOC_ERROR = resolve_doc()

_resolve_uidoc = resolve_uidoc
uidoc = resolve_uidoc()

# CLASS / FUNCTIONS
# ==============================================================================

# ── Sections 1-5 — unit helpers, extraction, geometry math, detected
#    elements and PointCloudAnalyzer — live in Services/point_cloud_analysis.py,
#    shared with the MCP server's point cloud tools (2026-10-02). ──────────────
from Services.point_cloud_analysis import (
    MM_PER_FOOT, ft_to_mm, mm_to_ft, build_cloud_filter,
    extract_full_cloud, extract_full_cloud_from_region,
    line_thickness_2d, cluster_2d, z_histogram, find_histogram_peaks,
    DetectedElement, DetectedWall, DetectedFloor, DetectedColumn,
    DetectedOpening, DetectedStair, DetectedRoof, PointCloudAnalyzer,
)

import uuid
from Snippets._compat import disposing
from Snippets._units import project_length_unit, LengthUnit


# ── Display in the project's length unit ──────────────────────────────────────
# The analysis service writes each element's DIMENSIONS text in mm (the MCP
# point cloud tools hand that text to the model); the window rewrites it in
# the unit the project displays. Only lengths switch — slope stays in degrees.

_METERS = LengthUnit("meters")


def _sizes_text(unit, sizes):
    """[('W', ft), ('H', ft)] → 'W=900 H=2100 mm' / 'W=3' - 0" H=7' - 0"'."""
    parts = u" ".join(u"{}={}".format(key, unit.text(ft)) for key, ft in sizes)
    return u"{} {}".format(parts, unit.tag) if unit.style == "decimal" else parts


def _area_text(unit, w_ft, d_ft):
    """Plan area: m² in a metric project, ft² in an imperial one."""
    if unit.is_metric:
        return u"~{:.1f} m²".format(_METERS.from_feet(w_ft) * _METERS.from_feet(d_ft))
    return u"~{:.0f} ft²".format(w_ft * d_ft)


def _dimensions_text(elem, unit):
    """The DIMENSIONS column for one detected element, written in `unit`.
    Falls back to the service's mm text when the element's data is missing."""
    d = getattr(elem, '_data', None) or {}
    kind = getattr(elem, 'Type', None)
    try:
        if kind == 'Wall':
            return u"L={}  T={}".format(unit.show(d['length_ft']),
                                       unit.show(d['thickness_ft']))
        if kind in ('Floor', 'Ceiling'):
            c = d['corners_xy']
            w, depth = abs(c[1][0] - c[0][0]), abs(c[2][1] - c[0][1])
            return u"{}  @Z={}".format(_area_text(unit, w, depth),
                                       unit.show(d['z_ft']))
        if kind == 'Column':
            return _sizes_text(unit, [('W', d['width_ft']), ('D', d['depth_ft']),
                                      ('H', d['z_top_ft'] - d['z_bot_ft'])])
        if kind in ('Door', 'Window'):
            return _sizes_text(unit, [('W', d['width_ft']), ('H', d['height_ft'])])
        if kind == 'Stair':
            lo, hi = unit.text(d['z_bot_ft']), unit.text(d['z_top_ft'])
            rise = (u"{}–{} {}".format(lo, hi, unit.tag) if unit.style == "decimal"
                    else u"{} to {}".format(lo, hi))
            return u"{} treads  Rise={}".format(d['tread_count'], rise)
        if kind == 'Roof':
            return u"Slope={:.1f}°  @Z={}".format(d['slope_deg'],
                                                 unit.show(d['z_ft']))
    except (KeyError, IndexError, TypeError, ValueError):
        pass
    return getattr(elem, 'Dimensions', u"")


# The pushbutton script importlib.reload()s this module on every click; a
# fixed __namespace__ would define the same .NET type twice and raise
# "Duplicate type name within an assembly" (rule S15).
_NS_SUFFIX = uuid.uuid4().hex[:8]

class PointCloudSelectionFilter(ISelectionFilter):
    __namespace__ = "T3Lab.PointCloud_" + _NS_SUFFIX
    def AllowElement(self, element):
        return isinstance(element, PointCloudInstance)

    def AllowReference(self, reference, position):
        return False


# ── Section 6: ElementBuilder — Services/point_cloud_build.py, shared with
#    the build_from_point_cloud AI tool. ─────────────────────────────────────
from Services.point_cloud_build import ElementBuilder


# ── Section 7: WPF Wizard Window ──────────────────────────────────────────────

class PointCloudModelWindow(T3WPFWindow):
    """Single-window UI for Point Cloud to Model analysis and element generation."""

    DENSITY_CAPS = [5000, 20000, 50000]

    def __init__(self, state=None):
        T3WPFWindow.__init__(self, XAML_FILE)
        state = state or {}
        self._pc_instance       = state.get('pc_instance')
        self._custom_min_pt     = state.get('min_pt')
        self._custom_max_pt     = state.get('max_pt')
        self._detected_elements = state.get('elements') or []
        self._wall_height_ft    = state.get('wall_height')
        self._builder           = ElementBuilder(doc)
        # Read per window: the module stays loaded across projects.
        self._unit              = project_length_unit(doc)
        self.result             = None
        # 'cloud' / 'region' — set by pick buttons. The window CLOSES for
        # every pick: hiding a ShowDialog window ends its modal loop, the
        # command's Execute returns, and every later click then runs Revit
        # API calls OUTSIDE the API context — the uncatchable 0xe0434352
        # crashes behind journal 1034-1037. The main loop below re-opens the
        # window after performing the pick inside the still-live command.
        self.pick_request       = None

        self.status_count.Text = u"Revit {}".format(REVIT_VERSION)
        if self._pc_instance is not None:
            self._set_cloud(self._pc_instance)
        else:
            self._try_preselect_cloud()
        if state.get('custom_checked') or self._custom_min_pt is not None:
            self.rb_custom_region.IsChecked = True
        if self._custom_min_pt is not None:
            self._show_region_state()
        if state.get('region_error'):
            self._show_region_state(state['region_error'])
        if self._detected_elements:
            self._populate_results()
            self.status_text.Text = (
                u"Detected {} elements. Review and click Generate.".format(
                    len(self._detected_elements)))
        if state.get('status'):
            self.status_text.Text = state['status']

    def _try_preselect_cloud(self):
        try:
            sel_ids = uidoc.Selection.GetElementIds()
            for eid in sel_ids:
                el = doc.GetElement(eid)
                if isinstance(el, PointCloudInstance):
                    self._set_cloud(el)
                    break
        except Exception:
            pass

    def _set_cloud(self, pc_inst):
        self._pc_instance = pc_inst
        try:
            name = pc_inst.Name or u"Point Cloud"
        except Exception:
            name = u"Point Cloud"
        self.lbl_cloud_name.Text       = name
        self.lbl_cloud_name.Foreground = self._brush('#0F172A')
        self.status_text.Text          = u"Cloud selected: {}".format(name)

    def _brush(self, hex_color):
        from System.Windows.Media import BrushConverter
        return BrushConverter().ConvertFromString(hex_color)

    def _set_status(self, msg):
        """Update the status bar and pump the dispatcher so the text repaints
        while long work runs on the UI thread."""
        self.status_text.Text = msg
        try:
            # Render priority repaints the label without pumping queued
            # input events (no re-entrant clicks mid-analysis).
            # NOTE: Action first, priority second — the (priority, delegate)
            # legacy overload fails to bind under IronPython.
            self.Dispatcher.Invoke(Action(lambda: None),
                                   DispatcherPriority.Render)
        except Exception:
            pass

    # ── Window chrome ──────────────────────────────────────────────────────────

    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        self.WindowState = (WindowState.Normal
                            if self.WindowState == WindowState.Maximized
                            else WindowState.Maximized)

    def close_button_clicked(self, sender, e):
        self.Close()

    # ── Step 0 handlers ────────────────────────────────────────────────────────

    def btn_pick_cloud_clicked(self, sender, e):
        # Close and let the main loop pick inside the live command context —
        # NEVER Hide() a ShowDialog window for picking (see __init__ note).
        self.pick_request = 'cloud'
        self.Close()

    def rb_custom_region_checked(self, sender, e):
        self.btn_pick_region.IsEnabled = True
        if self._custom_min_pt is None:
            self._show_region_state()

    def rb_full_extent_checked(self, sender, e):
        self.btn_pick_region.IsEnabled = False
        self._custom_min_pt = None
        self._custom_max_pt = None
        self._show_region_state()

    def _show_region_state(self, error_msg=None):
        """Reflect the custom-region state in the label under the pick button."""
        try:
            if error_msg:
                self.lbl_region_info.Text       = error_msg
                self.lbl_region_info.Foreground = self._brush('#D23B3B')
            elif self._custom_min_pt is not None:
                w_ft = self._custom_max_pt.X - self._custom_min_pt.X
                d_ft = self._custom_max_pt.Y - self._custom_min_pt.Y
                self.lbl_region_info.Text = u"Region set: {} × {}".format(
                    self._unit.show(w_ft), self._unit.show(d_ft))
                self.lbl_region_info.Foreground = self._brush('#0B8A5A')
                self.btn_pick_region.Content    = u"Re-pick Region"
            else:
                self.lbl_region_info.Text       = u"No region defined"
                self.lbl_region_info.Foreground = self._brush('#94A3B8')
                self.btn_pick_region.Content    = u"Drag Region Box"
        except Exception:
            pass

    def btn_use_crop_clicked(self, sender, e):
        """Region from the active view's crop box. Read-only API — never
        enters a selection mode, so it cannot trip the Revit 2023 point
        cloud engine crash that interactive picking triggers on some clouds."""
        try:
            view = None
            try:
                view = uidoc.ActiveGraphicalView
            except Exception:
                pass
            if view is None:
                view = doc.ActiveView
            if not view.CropBoxActive:
                TaskDialog.Show(
                    "Point Cloud to Model",
                    "The active view has no crop region enabled.\n"
                    "Turn on the crop region, size it around the scan area, "
                    "then click 'Use View Crop' again.")
                return
            cb = view.CropBox
            tf = cb.Transform
            corners = [
                XYZ(cb.Min.X, cb.Min.Y, cb.Min.Z),
                XYZ(cb.Max.X, cb.Min.Y, cb.Min.Z),
                XYZ(cb.Min.X, cb.Max.Y, cb.Min.Z),
                XYZ(cb.Max.X, cb.Max.Y, cb.Min.Z),
                XYZ(cb.Min.X, cb.Min.Y, cb.Max.Z),
                XYZ(cb.Max.X, cb.Max.Y, cb.Max.Z),
            ]
            mpts = [tf.OfPoint(p) for p in corners]
            xs = [p.X for p in mpts]
            ys = [p.Y for p in mpts]
            zs = [p.Z for p in mpts]
            self._custom_min_pt = XYZ(min(xs), min(ys), min(zs) - mm_to_ft(500.0))
            self._custom_max_pt = XYZ(max(xs), max(ys), max(zs) + mm_to_ft(500.0))
            self.rb_custom_region.IsChecked = True
            self._show_region_state()
            self.status_text.Text = (
                u"Region taken from the view crop — click Analyze.")
        except Exception as ex:
            import traceback
            logger.error("use_crop: {}\n{}".format(ex, traceback.format_exc()))
            self._show_region_state(u"Could not read the view crop")
            self.status_text.Text = u"Use View Crop failed: {}".format(ex)

    def btn_pick_region_clicked(self, sender, e):
        # Close and let the main loop run PickBox inside the live command
        # context — NEVER Hide() a ShowDialog window for picking.
        self.pick_request = 'region'
        self.Close()

    def btn_analyze_clicked(self, sender, e):
        if self._pc_instance is None:
            TaskDialog.Show("Point Cloud to Model",
                            "Please select a Point Cloud Instance first.")
            return
        if self.rb_custom_region.IsChecked and self._custom_min_pt is None:
            TaskDialog.Show("Point Cloud to Model",
                            "Custom Region selected but no region defined.\n"
                            "Use 'Drag Region Box' or 'Use View Crop' first, "
                            "or switch to Full Cloud Extent.")
            return
        if REVIT_VERSION <= 2023:
            # Reading big clouds through the 2023 API can crash Revit itself
            # (ReCap engine AV — uncatchable). Warn and let the user back out.
            from Autodesk.Revit.UI import TaskDialogCommonButtons, TaskDialogResult
            td = TaskDialog("Point Cloud to Model")
            td.MainInstruction = "Continue on Revit {}?".format(REVIT_VERSION)
            td.MainContent = (
                "Reading large point clouds through the Revit 2023 API can "
                "crash Revit itself (a point cloud engine bug outside this "
                "tool). Save your work before continuing.\n\n"
                "Revit 2026 reads the same clouds reliably.\n\n"
                "Run the analysis now?")
            td.CommonButtons = (TaskDialogCommonButtons.Yes
                                | TaskDialogCommonButtons.No)
            if td.Show() != TaskDialogResult.Yes:
                self.status_text.Text = u"Analysis cancelled."
                return
        self.btn_analyze.IsEnabled = False
        try:
            from System.Windows.Input import Cursors
            self.Cursor = Cursors.Wait
        except Exception:
            pass
        try:
            self._run_analysis()
        except Exception as ex:
            import traceback
            logger.error("Analysis error: {}\n{}".format(ex, traceback.format_exc()))
            self._set_status(u"Analysis failed — see pyRevit log.")
            TaskDialog.Show("Point Cloud to Model",
                            u"Analysis error: {}".format(ex))
        finally:
            self.btn_analyze.IsEnabled = True
            try:
                self.Cursor = None
            except Exception:
                pass

    def _default_settings(self):
        """
        Architectural + structural scope: Walls, Floors, Ceilings, Doors,
        Windows, Columns, Stairs, Roof. MEP (ducts, pipes, equipment) must
        NOT be modelled — the column detector carries anti-MEP filters
        (min 200 mm cross-section, aspect <= 2.5, must rise above furniture
        height), and walls require tall evidence below the ceiling.
        """
        return {
            'detect_wall':    True,
            'snap_tol':       1.0,
            'wall_min_len':   500.0,
            'detect_floor':   True,
            'floor_zbin':     50.0,
            'floor_min_area': 1.0,
            'detect_ceiling': True,
            'ceil_zbin':      50.0,
            'detect_door':    True,
            'door_min_w':     600.0,
            'door_min_h':     1800.0,
            'detect_window':  True,
            'win_min_w':      400.0,
            'win_min_h':      400.0,
            'detect_column':  True,
            'col_min':        200.0,   # pipes/small risers are thinner
            'col_max':        600.0,
            'detect_stair':   True,
            'stair_riser':    165.0,
            'detect_roof':    True,
            'roof_slope':     5.0,
        }

    def _run_analysis(self):
        idx = self.cmb_density.SelectedIndex
        if idx < 0:
            idx = 1
        density_cap = self.DENSITY_CAPS[idx]

        self._set_status(u"Extracting up to {} points…".format(density_cap))

        if self.rb_custom_region.IsChecked and self._custom_min_pt:
            min_pt = self._custom_min_pt
            max_pt = self._custom_max_pt
            # Corners picked in a plan view share one Z (span = the ±500 mm
            # padding only) — expand to the cloud's full height instead of
            # scanning a 1 m slab
            if (max_pt.Z - min_pt.Z) < mm_to_ft(1500.0):
                bbox = self._pc_instance.get_BoundingBox(None)
                if bbox is not None:
                    min_pt = XYZ(min_pt.X, min_pt.Y, bbox.Min.Z - 0.5)
                    max_pt = XYZ(max_pt.X, max_pt.Y, bbox.Max.Z + 0.5)
            half_x = (max_pt.X - min_pt.X) / 2.0
            half_y = (max_pt.Y - min_pt.Y) / 2.0
            half_z = (max_pt.Z - min_pt.Z) / 2.0
            center = XYZ(
                (min_pt.X + max_pt.X) / 2.0,
                (min_pt.Y + max_pt.Y) / 2.0,
                (min_pt.Z + max_pt.Z) / 2.0)
            pts = extract_full_cloud_from_region(
                self._pc_instance, center, half_x, half_y, half_z, density_cap,
                progress_cb=self._set_status)
        else:
            pts = extract_full_cloud(self._pc_instance, density_cap,
                                     progress_cb=self._set_status)

        if len(pts) < 50:
            self._set_status(u"Extraction returned {} points.".format(len(pts)))
            TaskDialog.Show(
                "Point Cloud to Model",
                "Only {} points extracted. "
                "The cloud may be out of range, unloaded or empty.\n"
                "Check that the point cloud file (.rcp/.rcs) is loaded and "
                "available offline (OneDrive cloud-only placeholders can't "
                "be read), then try again.".format(len(pts)))
            return

        self._set_status(u"Analyzing {} points…".format(len(pts)))
        settings = self._default_settings()
        analyzer = PointCloudAnalyzer(pts, doc=doc)
        self._detected_elements = analyzer.run(settings, progress_cb=self._set_status)

        # Wall height for generation: full scan height, clamped to a sane range
        z_range = analyzer._z_max - analyzer._z_min
        self._wall_height_ft = min(max(z_range, mm_to_ft(2200.0)), mm_to_ft(6000.0))

        self._populate_results()
        total = len(self._detected_elements)
        self._set_status(
            u"Detected {} elements. Review and click Generate.".format(total))

    def _populate_results(self):
        elems = self._detected_elements

        type_counts = {}
        for e in elems:
            type_counts[e.Type] = type_counts.get(e.Type, 0) + 1

        badge_map = {
            'Wall':    self.badge_wall,
            'Floor':   self.badge_floor,
            'Ceiling': self.badge_ceiling,
            'Door':    self.badge_door,
            'Window':  self.badge_window,
            'Column':  self.badge_column,
            'Stair':   self.badge_stair,
            'Roof':    self.badge_roof,
        }
        labels = {
            'Wall': 'Walls', 'Floor': 'Floors', 'Ceiling': 'Ceilings',
            'Door': 'Doors', 'Window': 'Windows', 'Column': 'Columns',
            'Stair': 'Stairs', 'Roof': 'Roofs',
        }
        for t, badge in badge_map.items():
            cnt = type_counts.get(t, 0)
            badge.Text = u"{} {}".format(cnt, labels[t])

        for e in elems:
            e.Dimensions = _dimensions_text(e, self._unit)

        total = len(elems)
        self.pnl_empty_state.Visibility = Visibility.Collapsed
        if total > 0:
            self.results_grid.ItemsSource     = to_items_source(elems)
            self.results_grid_card.Visibility = Visibility.Visible
            self.pnl_no_results.Visibility    = Visibility.Collapsed
            self.btn_generate.IsEnabled       = True
            self.btn_generate.Content         = u"Generate {} Selected".format(total)
        else:
            self.results_grid_card.Visibility = Visibility.Collapsed
            self.pnl_no_results.Visibility    = Visibility.Visible
            self.btn_generate.IsEnabled       = False
            self.btn_generate.Content         = u"Generate"

    # ── Step 2 handlers ────────────────────────────────────────────────────────

    def btn_generate_clicked(self, sender, e):
        selected = [el for el in self._detected_elements if el.Include]
        if not selected:
            TaskDialog.Show(
                "Point Cloud to Model",
                "No elements selected. Check the Include checkboxes.")
            return

        self.btn_generate.IsEnabled = False
        self._set_status(u"Creating {} elements…".format(len(selected)))
        try:
            summary = self._builder.build_all(selected, self._wall_height_ft)
            counts  = summary['counts']
            errors  = summary['errors']
            skipped = summary.get('skipped') or {}
            parts   = [u"{} {}".format(v, k)
                       for k, v in counts.items() if v > 0]
            msg = (u"Created: " + u", ".join(parts)) if parts else u"No elements created."
            n_stair = skipped.get('Stair', 0)
            if n_stair:
                msg += (u"\n{} stair(s) detected but NOT created — stair "
                        u"creation is not supported yet; model them "
                        u"manually.".format(n_stair))
            if errors:
                msg += u"\n{} element(s) failed — see pyRevit log.".format(errors)
            TaskDialog.Show("Point Cloud to Model", msg)
            self.result = summary
            self.Close()
        except Exception as ex:
            import traceback
            logger.error("Generate error: {}\n{}".format(ex, traceback.format_exc()))
            self._set_status(u"Generation failed — see pyRevit log.")
            TaskDialog.Show("Point Cloud to Model",
                            u"Generation error: {}".format(ex))
        finally:
            self.btn_generate.IsEnabled = True

    # ── Select-all o header cot checkbox ────────────────────────────────
    # toggle_all_rows() nam trong T3WPFWindow: no chay tren grid.Items nen chi
    # dong dang hien thi (sau filter/sort) bi doi, dung nhu nguoi dung thay.

    def select_all_results_grid_clicked(self, sender, e):
        """Header checkbox: chon/bo chon moi dong dang hien thi cua results_grid."""
        self.toggle_all_rows(self.results_grid, "Include", sender.IsChecked)


# MAIN
# ==============================================================================

def _pick_cloud_into(state):
    """PickObject inside the live command context; updates state in place."""
    try:
        ref = uidoc.Selection.PickObject(
            ObjectType.Element,
            PointCloudSelectionFilter(),
            "Select a Point Cloud Instance")
        state['pc_instance'] = doc.GetElement(ref.ElementId)
        state['status'] = None
    except OperationCanceledException:
        pass  # Esc — keep the previous cloud, if any
    except Exception as ex:
        logger.error("PickObject: {}".format(ex))
        state['status'] = u"Cloud selection failed: {}".format(ex)


def _pick_region_into(state):
    """PickBox inside the live command context; updates state in place."""
    state['custom_checked'] = True
    try:
        box = uidoc.Selection.PickBox(
            PickBoxStyle.Directional,
            "Drag a rectangle around the scan region")
        pt1, pt2 = box.Min, box.Max
        w = abs(pt1.X - pt2.X)
        d = abs(pt1.Y - pt2.Y)
        if w < mm_to_ft(100.0) or d < mm_to_ft(100.0):
            state['min_pt'] = None
            state['max_pt'] = None
            state['region_error'] = u"Region too small — drag a larger rectangle"
            state['status'] = u"Region not set: the dragged rectangle is nearly empty."
        else:
            state['min_pt'] = XYZ(
                min(pt1.X, pt2.X), min(pt1.Y, pt2.Y),
                min(pt1.Z, pt2.Z) - mm_to_ft(500.0))
            state['max_pt'] = XYZ(
                max(pt1.X, pt2.X), max(pt1.Y, pt2.Y),
                max(pt1.Z, pt2.Z) + mm_to_ft(500.0))
            state['status'] = u"Custom region captured — click Analyze."
    except OperationCanceledException:
        if state.get('min_pt') is not None:
            state['status'] = u"Pick cancelled — keeping the previous region."
        else:
            state['status'] = u"Pick cancelled — no region set yet."
    except Exception as ex:
        logger.error("PickBox: {}".format(ex))
        state['region_error'] = u"Pick failed — see pyRevit log"
        state['status'] = u"Pick failed: {}".format(ex)


def show_point_cloud_dialog(doc_param=None, uidoc_param=None):
    global doc, uidoc
    if doc_param:
        doc = doc_param
    if uidoc_param:
        uidoc = uidoc_param
    if not doc:
        try:
            doc = revit.doc
            uidoc = revit.uidoc
        except Exception:
            pass

    try:
        if not doc or not uidoc:
            TaskDialog.Show(
                "Point Cloud to Model",
                u"No active document found.\nOpen a project and try again.")
            return

        if not doc.ActiveView:
            TaskDialog.Show(
                "Point Cloud to Model",
                u"This tool needs an active Revit view to select a point "
                u"cloud in.\nOpen the model's plan or 3D view and run it "
                u"from the T3Lab ribbon button.")
            return

        state = {}
        while True:
            window = PointCloudModelWindow(state)
            window.ShowDialog()
            req = window.pick_request
            state = {
                'pc_instance':    window._pc_instance,
                'min_pt':         window._custom_min_pt,
                'max_pt':         window._custom_max_pt,
                'elements':       window._detected_elements,
                'wall_height':    window._wall_height_ft,
                'custom_checked': bool(window.rb_custom_region.IsChecked),
            }
            if req == 'cloud':
                _pick_cloud_into(state)
            elif req == 'region':
                _pick_region_into(state)
            else:
                break
    except SystemExit:
        pass
    except Exception as ex:
        logger.error("Point Cloud to Model error: {}".format(ex))
        import traceback
        logger.error(traceback.format_exc())
        TaskDialog.Show("Point Cloud to Model",
                        u"Unexpected error: {}".format(ex))


if __name__ == '__main__':
    show_point_cloud_dialog()
