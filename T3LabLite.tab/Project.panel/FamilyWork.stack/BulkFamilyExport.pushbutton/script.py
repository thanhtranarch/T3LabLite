# -*- coding: utf-8 -*-
"""
Bulk Family Export  (merged with DWG to Family)

Export CAD blocks from imported DWG/DXF files as individual Revit families,
with an option to load the exported families directly into the project.

Parametric door/window families use reference planes + dimensions linked to
Width/Height parameters, following the same pattern as the JSONtoFamily tool.

Author: T3Lab
"""

from __future__ import unicode_literals

__author__  = "T3Lab"
__title__   = "Bulk Family\nExport"
__version__ = "2.0.0"

# ─── Imports ──────────────────────────────────────────────────────────────────

import os
import re
import traceback

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')

from System.Windows import WindowState
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind

from pyrevit import revit, DB, forms, script
from Autodesk.Revit.DB import (
    ImportInstance, FilteredElementCollector,
    Options, GeometryInstance,
    Line, Arc, XYZ, Plane,
    CurveArray, CurveArrArray,
    SketchPlane, SaveAsOptions,
    Transaction, ElementId,
    View, ViewType, ReferencePlane, ReferenceArray,
    PlanarFace, Solid,
)

from Utils.DWGFamilyHelpers import get_xy_bounds

# ─── Revit context ────────────────────────────────────────────────────────────

uidoc = __revit__.ActiveUIDocument
doc   = uidoc.Document
app   = __revit__.Application

logger = script.get_logger()

# ─── Category-to-template mapping ────────────────────────────────────────────

CATEGORY_TEMPLATES = [
    ("Generic Model",        ["Generic Model.rft", "Metric Generic Model.rft"]),
    ("Door",                 ["Door.rft", "Metric Door.rft"]),
    ("Window",               ["Window.rft", "Metric Window.rft"]),
    ("Furniture",            ["Furniture.rft", "Metric Furniture.rft"]),
    ("Plumbing Fixture",     ["Plumbing Fixture.rft", "Metric Plumbing Fixture.rft"]),
    ("Electrical Equipment", ["Electrical Equipment.rft"]),
    ("Mechanical Equipment", ["Mechanical Equipment.rft"]),
    ("Specialty Equipment",  ["Specialty Equipment.rft", "Metric Specialty Equipment.rft"]),
    ("Casework",             ["Casework.rft", "Metric Casework.rft"]),
    ("Columns",              ["Column.rft", "Metric Column.rft"]),
    ("Lighting Fixture",     ["Lighting Fixture.rft", "Metric Lighting Fixture.rft"]),
    ("Site",                 ["Site.rft", "Metric Site.rft"]),
    ("Entourage",            ["Entourage.rft", "Metric Entourage.rft"]),
]

DISCIPLINES = [
    "Architecture",
    "Structure",
    "Mechanical",
    "Electrical",
    "Plumbing",
    "Fire Protection",
    "General"
]

# --- Door Preset Definitions -----------------------------------------------
# Source: T3Lab_Door_location_Not_Shared_2.rfa (W=810, H=2200, FW=65, PE=25, PI=25, LT=40)

DOOR_PRESETS = [
    # (label, width_mm, height_mm, frame_w_mm, proj_ext_mm, proj_int_mm, leaf_t_mm, door_count)
    ("Single_Swing_700x2100",   700, 2100, 65, 25, 25, 40, 1),
    ("Single_Swing_810x2200",   810, 2200, 65, 25, 25, 40, 1),
    ("Single_Swing_900x2200",   900, 2200, 65, 25, 25, 40, 1),
    ("Single_Swing_1000x2200", 1000, 2200, 65, 25, 25, 40, 1),
    ("Single_Swing_810x2400",   810, 2400, 65, 25, 25, 40, 1),
    ("Single_Swing_900x2400",   900, 2400, 65, 25, 25, 40, 1),
    ("Double_Swing_1600x2200", 1600, 2200, 65, 25, 25, 40, 2),
    ("Double_Swing_1800x2200", 1800, 2200, 65, 25, 25, 40, 2),
    ("Double_Swing_2000x2200", 2000, 2200, 65, 25, 25, 40, 2),
    ("Double_Swing_1600x2400", 1600, 2400, 65, 25, 25, 40, 2),
]

SCL = 1.0 / 304.8   # mm to Revit internal feet (confirmed working)




# ─── Data Model ───────────────────────────────────────────────────────────────

_CAT_HINTS = [
    ("Furniture",      ["chair", "table", "desk", "sofa", "bed", "cabinet"]),
    ("Plumbing Fixture", ["wc", "toilet", "sink", "basin", "shower", "bath", "urinal"]),
    ("Lighting Fixture", ["light", "lamp", "fixture", "luminaire"]),
    ("Casework",       ["casework", "counter", "kitchen", "shelv"]),
    ("Specialty Equipment", ["equip", "machine", "appliance"]),
]


def _suggest_category(name, arc_count, width_mm, depth_mm):
    """Auto-suggest a Revit family category from block geometry + name heuristics."""
    lname = name.lower()
    for cat, keywords in _CAT_HINTS:
        if any(k in lname for k in keywords):
            return cat
    if arc_count >= 1:
        return "Door"
    if arc_count == 0 and 0 < depth_mm < 350 and width_mm >= 400:
        return "Window"
    return "Generic Model"


class BlockItem(object):
    """Represents a discovered CAD block for DataGrid binding."""

    def __init__(self, name, curve_count, instance_count, curves, layer_level=""):
        self.IsSelected    = True
        self.BlockName     = name
        self.CurveCount    = curve_count
        self.InstanceCount = instance_count
        self.LayerLevel    = layer_level
        self._curves       = curves          # internal use only

        # ── Geometry detail properties (shown in DataGrid) ──
        arc_count = sum(1 for c in curves if isinstance(c, Arc))
        self.ArcCount = arc_count

        try:
            min_x, max_x, min_y, max_y = get_xy_bounds(curves)
            w = (max_x - min_x) * 304.8   # feet → mm
            d = (max_y - min_y) * 304.8
            self.WidthMM = "{:.0f}".format(w)
            self.DepthMM = "{:.0f}".format(d)
        except Exception:
            w, d = 0.0, 0.0
            self.WidthMM = "-"
            self.DepthMM = "-"

        self.SuggestedCat = _suggest_category(name, arc_count, w, d)


# ─── Main Window ──────────────────────────────────────────────────────────────

class BulkFamilyExportWindow(forms.WPFWindow):

    def __init__(self):
        # script.py → pushbutton → stack → panel → tab → extension (5 levels)
        ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(
            os.path.dirname(os.path.dirname(__file__)))))
        xaml_path = os.path.join(ext_dir, 'lib', 'GUI', 'BulkFamilyExport.xaml')
        forms.WPFWindow.__init__(self, xaml_path)

        self._ext_dir       = ext_dir
        self._cad_instances = []
        self._block_items   = []

        self._load_logo()
        self._init_cad_files()
        self._init_disciplines()
        self._init_categories()
        self._update_status("Ready")

    # ── Logo ──────────────────────────────────────────────────────────────

    def _load_logo(self):
        """Load T3Lab logo into title bar and window icon."""
        try:
            logo_path = os.path.join(self._ext_dir, 'lib', 'GUI', 'T3Lab_logo.png')
            if os.path.exists(logo_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                self.logo_image.Source = bitmap
                self.Icon = bitmap
        except Exception:
            pass

    # ── Initialisation ────────────────────────────────────────────────────

    def _init_cad_files(self):
        """Populate the CAD file ComboBox with all ImportInstances."""
        collector = FilteredElementCollector(doc).OfClass(ImportInstance)
        self.cad_file_combo.Items.Add("<All Imported CAD Files>")
        for inst in collector:
            name = self._get_cad_name(inst)
            self._cad_instances.append(inst)
            self.cad_file_combo.Items.Add(name)
        if self._cad_instances:
            self.cad_file_combo.SelectedIndex = 0

    def _get_cad_name(self, inst):
        """Return the Type Name for an ImportInstance."""
        try:
            type_id = inst.GetTypeId()
            if type_id and type_id != ElementId.InvalidElementId:
                elem_type = doc.GetElement(type_id)
                if elem_type and hasattr(elem_type, 'Name') and elem_type.Name:
                    return elem_type.Name
        except Exception:
            pass
        return inst.Name if hasattr(inst, 'Name') else "Unknown CAD Type"

    def _init_disciplines(self):
        """Populate the discipline ComboBox."""
        for name in DISCIPLINES:
            self.discipline_combo.Items.Add(name)
        self.discipline_combo.SelectedIndex = 6

    def _init_categories(self):
        """Populate the category ComboBox."""
        for name, _ in CATEGORY_TEMPLATES:
            self.category_combo.Items.Add(name)
        self.category_combo.SelectedIndex = 0

    def _update_status(self, text):
        try:
            self.status_text.Text = text
        except Exception:
            pass

    # ── Block Scanning ────────────────────────────────────────────────────

    def scan_blocks_clicked(self, sender, e):
        """Event handler: scan the selected CAD file for blocks."""
        if not self._cad_instances:
            forms.alert("No imported CAD files found in the document.")
            return

        idx = self.cad_file_combo.SelectedIndex - 1
        if idx < -1 or idx >= len(self._cad_instances):
            forms.alert("Please select a CAD file.")
            return

        self._update_status("Scanning blocks...")

        blocks = []
        try:
            if idx == -1:
                # All imported CAD files
                name_counts = {}
                for inst in self._cad_instances:
                    item = self._scan_entire_cad(inst)
                    if item:
                        base_name = item.BlockName
                        if base_name in name_counts:
                            name_counts[base_name] += 1
                            item.BlockName = "{}_{}".format(base_name, name_counts[base_name])
                        else:
                            name_counts[base_name] = 1
                        blocks.append(item)
            else:
                import_inst = self._cad_instances[idx]
                blocks = self._scan_blocks(import_inst)
                if not blocks:
                    # Fallback: treat the entire CAD as one block
                    item = self._scan_entire_cad(import_inst)
                    if item:
                        blocks.append(item)
        except Exception as ex:
            logger.error("Scan error:\n{}".format(traceback.format_exc()))
            forms.alert("Error scanning blocks:\n{}".format(str(ex)))
            self._update_status("Scan failed")
            return

        if not blocks:
            forms.alert(
                "No blocks or curves found in the selected CAD file(s).\n\n"
                "Make sure the CAD files contain geometry.")
            self._update_status("No geometry found")
            return

        self._block_items = blocks
        self.blocks_grid.ItemsSource = blocks
        self._update_status("Found {} unique item(s)".format(len(blocks)))
        self.block_count_text.Text = "{} items found".format(len(blocks))

    def _scan_entire_cad(self, import_inst):
        """Treat an ImportInstance as a single block and return a BlockItem."""
        opt = Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True

        geom = import_inst.get_Geometry(opt)
        if not geom:
            return None

        min_len = getattr(app, 'ShortCurveTolerance', 0.00256)

        def is_curve(item):
            try:
                from Autodesk.Revit.DB import Curve as _Curve
                return (isinstance(item, _Curve)
                        and item.IsBound
                        and item.Length >= min_len)
            except Exception:
                return False

        def collect_curves(geo_elem):
            curves = []
            for item in geo_elem:
                if is_curve(item):
                    curves.append(item)
                elif isinstance(item, GeometryInstance):
                    nested = item.GetInstanceGeometry()
                    if nested:
                        curves.extend(collect_curves(nested))
            return curves

        curves = collect_curves(geom)
        if curves:
            name = self._get_cad_name(import_inst)

            level_name = ""
            try:
                level_param = import_inst.get_Parameter(DB.BuiltInParameter.IMPORT_BASE_LEVEL)
                if level_param and level_param.HasValue:
                    level_name = level_param.AsValueString()
                else:
                    level_id = import_inst.LevelId
                    if level_id and str(level_id) != "-1":
                        level_elem = doc.GetElement(level_id)
                        if level_elem:
                            level_name = level_elem.Name
            except Exception:
                pass

            return BlockItem(name, len(curves), 1, curves, layer_level=level_name)
        return None

    def _scan_blocks(self, import_inst):
        """Walk the geometry tree of an ImportInstance and return unique blocks
        as a list of BlockItem objects."""
        opt = Options()
        opt.ComputeReferences = True
        opt.IncludeNonVisibleObjects = True

        geom = import_inst.get_Geometry(opt)
        if not geom:
            return []

        min_len = getattr(app, 'ShortCurveTolerance', 0.00256)
        found   = {}      # fingerprint -> {name, curves, count}
        counter = [0]     # mutable counter (IronPython 2 – no nonlocal)

        # ── helpers ──

        def is_curve(item):
            try:
                from Autodesk.Revit.DB import Curve as _Curve
                return (isinstance(item, _Curve)
                        and item.IsBound
                        and item.Length >= min_len)
            except Exception:
                return False

        def collect_curves(geo_elem):
            """Recursively collect all valid curves from a geometry element."""
            curves = []
            for item in geo_elem:
                if is_curve(item):
                    curves.append(item)
                elif isinstance(item, GeometryInstance):
                    nested = item.GetInstanceGeometry()
                    if nested:
                        curves.extend(collect_curves(nested))
            return curves

        def fingerprint(curves):
            """Rotation- and position-invariant fingerprint for a set of curves."""
            n     = len(curves)
            total = round(sum(c.Length for c in curves), 1)
            return (n, total)

        def style_name(geo_inst):
            """Try to retrieve the DWG layer name from the GraphicsStyle."""
            try:
                sid = geo_inst.GraphicsStyleId
                if sid and sid != ElementId.InvalidElementId:
                    style = doc.GetElement(sid)
                    if style and hasattr(style, 'Name') and style.Name:
                        return style.Name
            except Exception:
                pass
            return None

        def register(curves, geo_inst):
            """Register a block (or increment its instance count)."""
            fp = fingerprint(curves)
            if fp in found:
                found[fp]['count'] += 1
                return

            layer = style_name(geo_inst)
            counter[0] += 1

            block_name = ""
            try:
                if hasattr(geo_inst, 'Symbol') and geo_inst.Symbol:
                    block_name = geo_inst.Symbol.Name
            except Exception:
                pass

            if block_name:
                name = block_name
            elif layer:
                name = "{}_Block_{:03d}".format(layer, counter[0])
            else:
                name = "Block_{:03d}".format(counter[0])

            found[fp] = {'name': name, 'curves': curves, 'count': 1, 'layer': layer or ""}

        # ── walk geometry tree ──

        def walk(geo_elem, depth):
            for item in geo_elem:
                if not isinstance(item, GeometryInstance):
                    continue

                inst_geom = item.GetInstanceGeometry()
                if not inst_geom:
                    continue

                if depth == 0:
                    # Top-level container (DWG model space) – dive in
                    walk(inst_geom, depth + 1)
                else:
                    # Block reference at depth >= 1
                    curves = collect_curves(inst_geom)
                    if curves:
                        register(curves, item)

        walk(geom, 0)

        # Convert to sorted list of BlockItem
        items = []
        for data in sorted(found.values(), key=lambda d: d['name']):
            items.append(BlockItem(
                data['name'], len(data['curves']), data['count'], data['curves'],
                layer_level=data.get('layer', "")))
        return items

    # ── Folder picker ─────────────────────────────────────────────────────

    def browse_folder_clicked(self, sender, e):
        folder = forms.pick_folder()
        if folder:
            self.output_path.Text = folder

    # ── Select / Deselect ─────────────────────────────────────────────────

    def select_all_clicked(self, sender, e):
        for item in self._block_items:
            item.IsSelected = True
        self.blocks_grid.Items.Refresh()

    def deselect_all_clicked(self, sender, e):
        for item in self._block_items:
            item.IsSelected = False
        self.blocks_grid.Items.Refresh()

    # ── Export ────────────────────────────────────────────────────────────

    def export_clicked(self, sender, e):
        """Export each selected block as a .rfa family file."""
        output_folder = self.output_path.Text
        if not output_folder or not os.path.isdir(output_folder):
            forms.alert("Please select a valid output folder.")
            return

        selected = [b for b in self._block_items if b.IsSelected]
        if not selected:
            forms.alert("No blocks selected for export.")
            return

        cat_idx = self.category_combo.SelectedIndex
        if cat_idx < 0:
            forms.alert("Please select a family category.")
            return

        category_name = CATEGORY_TEMPLATES[cat_idx][0]

        disc_idx = self.discipline_combo.SelectedIndex
        if disc_idx < 0:
            forms.alert("Please select a discipline.")
            return
        discipline_name = DISCIPLINES[disc_idx]
        template_path = self._find_template(cat_idx)
        if not template_path:
            forms.alert(
                "Could not find family template for '{}'.\n\n"
                "Please ensure Revit family templates are installed.".format(
                    category_name))
            return

        load_to_project = (self.chk_load_to_project.IsChecked == True)

        self._update_status("Exporting {} block(s)...".format(len(selected)))
        success, failed = 0, 0

        for item in selected:
            try:
                self._update_status("Exporting: {}".format(item.BlockName))
                ok = self._export_block(
                    item, template_path, output_folder,
                    discipline_name, category_name, load_to_project)
                if ok:
                    success += 1
                else:
                    failed += 1
            except Exception:
                logger.error("Export '{}' failed:\n{}".format(
                    item.BlockName, traceback.format_exc()))
                failed += 1

        self._update_status(
            "Done: {} exported, {} failed".format(success, failed))

        extra = "\nFamilies loaded to project." if load_to_project and success else ""
        forms.alert(
            "Export complete!\n\n"
            "Succeeded: {}\nFailed: {}\n\n"
            "Output folder:\n{}{}".format(success, failed, output_folder, extra))

    # ── Template lookup ───────────────────────────────────────────────────

    def _find_template(self, cat_idx):
        """Locate the .rft family template for the chosen category."""
        _, template_names = CATEGORY_TEMPLATES[cat_idx]

        search_dirs = []

        # Revit-configured path (most reliable)
        try:
            tdir = app.FamilyTemplatePath
            if tdir and os.path.isdir(tdir):
                search_dirs.append(tdir)
        except Exception:
            pass

        # Standard fallback paths
        ver  = app.VersionNumber
        base = r"C:\ProgramData\Autodesk\RVT {}".format(ver)
        for sub in ("English", "", "English-Imperial", "English_I"):
            if sub:
                search_dirs.append(
                    os.path.join(base, "Family Templates", sub))
            else:
                search_dirs.append(os.path.join(base, "Family Templates"))

        for d in search_dirs:
            if not os.path.isdir(d):
                continue
            for tname in template_names:
                fp = os.path.join(d, tname)
                if os.path.isfile(fp):
                    return fp
        return None

    # ── Parametric reference planes (JSONtoFamily approach) ───────────────

    def _find_family_views(self, fam_doc):
        """Return (plan_view, elev_view) from a family document, or (None, None)."""
        plan_view = None
        elev_view = None
        for v in FilteredElementCollector(fam_doc).OfClass(View):
            try:
                if v.IsTemplate:
                    continue
                vt = v.ViewType
                if vt == ViewType.FloorPlan and plan_view is None:
                    plan_view = v
                elif vt == ViewType.Elevation and elev_view is None:
                    try:
                        if abs(v.ViewDirection.Y) > 0.99:
                            elev_view = v
                    except Exception:
                        pass
                if plan_view and elev_view:
                    break
            except Exception:
                continue
        return plan_view, elev_view

    def _create_parametric_refs(self, fam_doc, half_w, height,
                                 plan_view, elev_view,
                                 param_width_fp, param_height_fp):
        """Create reference planes + dimensions linked to Width/Height parameters.

        This mirrors the JSONtoFamily approach: reference planes are created at
        the geometry edges, then dimensioned and labelled so that changing
        Width/Height in the family drives the geometry.

        All failures are silenced so a bad template view never breaks export.
        """
        rp_left = rp_right = rp_top = None

        # ── Plan: Left / Right planes for Width ──
        if plan_view is not None:
            try:
                rp_left = fam_doc.FamilyCreate.NewReferencePlane(
                    XYZ(-half_w, -3, 0), XYZ(-half_w, 3, 0), XYZ.BasisZ, plan_view)
                rp_left.Name = "Edge_Left"

                rp_right = fam_doc.FamilyCreate.NewReferencePlane(
                    XYZ(half_w, -3, 0), XYZ(half_w, 3, 0), XYZ.BasisZ, plan_view)
                rp_right.Name = "Edge_Right"

                if param_width_fp is not None:
                    ref_arr = ReferenceArray()
                    ref_arr.Append(rp_left.GetReference())
                    ref_arr.Append(rp_right.GetReference())
                    dim_line = Line.CreateBound(
                        XYZ(-half_w * 1.5, 2, 0),
                        XYZ(half_w * 1.5, 2, 0))
                    dim = fam_doc.FamilyCreate.NewDimension(plan_view, dim_line, ref_arr)
                    if dim:
                        dim.FamilyLabel = param_width_fp
            except Exception:
                pass

        # ── Elevation: Top plane for Height ──
        if elev_view is not None:
            try:
                rp_top = fam_doc.FamilyCreate.NewReferencePlane(
                    XYZ(-3, 0, height), XYZ(3, 0, height), XYZ.BasisY, elev_view)
                rp_top.Name = "Top"

                if param_height_fp is not None:
                    # Find an existing Level / Ref-Level plane
                    rp_level = None
                    for rp in FilteredElementCollector(fam_doc).OfClass(ReferencePlane):
                        try:
                            n = rp.Name.lower()
                            if any(k in n for k in ("level", "floor", "bottom", "ref level")):
                                rp_level = rp
                                break
                        except Exception:
                            continue

                    if rp_level:
                        ref_arr = ReferenceArray()
                        ref_arr.Append(rp_level.GetReference())
                        ref_arr.Append(rp_top.GetReference())
                        dim_line = Line.CreateBound(
                            XYZ(0, 0, -0.1),
                            XYZ(0, 0, height + 0.1))
                        dim = fam_doc.FamilyCreate.NewDimension(elev_view, dim_line, ref_arr)
                        if dim:
                            dim.FamilyLabel = param_height_fp
            except Exception:
                pass

        return rp_left, rp_right, rp_top

    def _lock_faces_to_planes(self, fam_doc, solid_elem,
                               plan_view, elev_view,
                               rp_left, rp_right, rp_top):
        """Align + lock the extrusion side faces to the parametric reference planes."""
        try:
            geom_opt = Options()
            geom_opt.ComputeReferences = True
            geom_elem = solid_elem.get_Geometry(geom_opt)
            for geom_obj in geom_elem:
                if not isinstance(geom_obj, Solid):
                    continue
                for face in geom_obj.Faces:
                    if not isinstance(face, PlanarFace):
                        continue
                    n = face.FaceNormal
                    pairs = []
                    if rp_right and plan_view and n.X > 0.99:
                        pairs.append((rp_right, plan_view))
                    elif rp_left and plan_view and n.X < -0.99:
                        pairs.append((rp_left, plan_view))
                    elif rp_top and elev_view and n.Z > 0.99:
                        pairs.append((rp_top, elev_view))
                    for rp, view in pairs:
                        try:
                            align = fam_doc.FamilyCreate.NewAlignment(
                                view, rp.GetReference(), face.Reference)
                            if align:
                                align.IsLocked = True
                        except Exception:
                            pass
        except Exception:
            pass

    # ── Window geometry (JSONtoFamily hollow-extrusion approach) ──────────

    def _create_window_body(self, fam_doc, sketch_plane, half_w, half_depth, height,
                            param_height_fp, param_material):
        """Create proper window geometry using the JSONtoFamily inner-loop technique.

        Frame  : hollow extrusion in plan (outer rect + inner rect as second loop).
                 The inner loop punches a hole, exactly like JSONtoFamily 'inner_loops'.
        Glass  : thin slab centred at Y=0, hidden in plan/ceiling views.

        Returns (frame_ext, glass_ext); glass_ext may be None on failure.
        """
        from Autodesk.Revit.DB import (
            FamilyElementVisibility, FamilyElementVisibilityType, BuiltInParameter,
        )

        FRAME_W = max(min(half_w * 0.12, 0.1312), 0.0492)   # 12 % of half-width, ~15-40 mm
        half_d  = max(half_depth, 0.2461)                     # min ~75 mm (wall reveal)

        # ── helper: closed rectangle as CurveArray in the Z=0 plane ──
        def rect_loop(xmin, xmax, ymin, ymax):
            arr = CurveArray()
            pts = [XYZ(xmin, ymin, 0), XYZ(xmax, ymin, 0),
                   XYZ(xmax, ymax, 0), XYZ(xmin, ymax, 0)]
            for i in range(4):
                arr.Append(Line.CreateBound(pts[i], pts[(i + 1) % 4]))
            return arr

        # Outer boundary
        outer = rect_loop(-half_w,           half_w,           -half_d, half_d)
        # Inner boundary → punches hole through frame (JSONtoFamily inner_loops pattern)
        inner = rect_loop(-(half_w - FRAME_W), (half_w - FRAME_W),
                          -(half_d - FRAME_W), (half_d - FRAME_W))

        frame_profile = CurveArrArray()
        frame_profile.Append(outer)
        frame_profile.Append(inner)   # second loop = hollow centre

        frame_ext = fam_doc.FamilyCreate.NewExtrusion(True, frame_profile, sketch_plane, height)

        try:
            if param_height_fp:
                end_p = frame_ext.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                if end_p:
                    fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(
                        end_p, param_height_fp)
            if param_material:
                mat_p = frame_ext.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                if mat_p:
                    fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(
                        mat_p, param_material)
        except Exception:
            pass

        # ── Glass pane: thin slab centred at depth mid-point ──
        # Tiny Y extent (≈5 mm) in plan, extruded in Z from sill+FRAME_W to head-FRAME_W.
        # Hidden in plan/ceiling views – model-only visibility (JSONtoFamily visible_param).
        glass_ext = None
        try:
            iw    = half_w - FRAME_W
            GLASS = 0.0082   # ≈2.5 mm half-thickness → 5 mm total pane

            glass_rect = rect_loop(-iw, iw, -GLASS, GLASS)
            glass_profile = CurveArrArray()
            glass_profile.Append(glass_rect)

            glass_height = max(height - FRAME_W * 2, FRAME_W)
            glass_ext = fam_doc.FamilyCreate.NewExtrusion(
                True, glass_profile, sketch_plane, glass_height)
            glass_ext.StartOffset = FRAME_W

            vis = FamilyElementVisibility(FamilyElementVisibilityType.Model)
            vis.IsShownInTopBottom = False
            glass_ext.SetVisibility(vis)
        except Exception:
            glass_ext = None

        return frame_ext, glass_ext

    # ── Single-block export ───────────────────────────────────────────────

    def _export_block(self, block_item, template_path, output_folder,
                      discipline_name, category_name, load_to_project=False):
        """Create a .rfa family from a block's curves and save it.

        For Door and Window categories the family is fully parametric:
        reference planes are created at the geometry edges and dimensioned
        with the Width / Height family parameters (JSONtoFamily technique).
        """
        curves = block_item._curves
        if not curves:
            return False

        # Create a fresh family document
        fam_doc = app.NewFamilyDocument(template_path)

        # Bounding box via shared helper
        min_x, max_x, min_y, max_y = get_xy_bounds(curves)

        is_door   = "door"   in category_name.lower()
        is_window = "window" in category_name.lower()
        door_width = None

        if is_door:
            frame_xs, frame_ys = [], []
            for curve in curves:
                if isinstance(curve, Arc):
                    try:
                        C  = curve.Center
                        p0 = curve.GetEndPoint(0)
                        p1 = curve.GetEndPoint(1)
                        frame_xs.append(C.X)
                        frame_ys.append(C.Y)
                        if abs(p0.Y - C.Y) < abs(p1.Y - C.Y):
                            frame_xs.append(p0.X); frame_ys.append(p0.Y)
                        else:
                            frame_xs.append(p1.X); frame_ys.append(p1.Y)
                    except Exception:
                        pass
            if frame_xs:
                cx = (min(frame_xs) + max(frame_xs)) / 2.0
                cy = (min(frame_ys) + max(frame_ys)) / 2.0
                calc_w = max(frame_xs) - min(frame_xs)
                if calc_w > 0.01:
                    door_width = calc_w
            else:
                cx = (min_x + max_x) / 2.0
                cy = (min_y + max_y) / 2.0
                door_width = max_x - min_x
        else:
            cx = (min_x + max_x) / 2.0
            cy = (min_y + max_y) / 2.0

        half_w = max((max_x - min_x) / 2.0, 0.01)
        half_h = max((max_y - min_y) / 2.0, 0.01)

        mode_2d_only = False
        try:
            if hasattr(self, 'rb_2d_lines') and self.rb_2d_lines.IsChecked:
                mode_2d_only = True
        except Exception:
            pass

        t = Transaction(fam_doc, 'Create Block Geometry')
        t.Start()
        try:
            # Find or create a Z-up sketch plane
            sketch_plane = None
            for sp in FilteredElementCollector(fam_doc).OfClass(SketchPlane):
                try:
                    if abs(sp.GetPlane().Normal.Z - 1.0) < 0.001:
                        sketch_plane = sp
                        break
                except Exception:
                    pass
            if not sketch_plane:
                sketch_plane = SketchPlane.Create(
                    fam_doc,
                    Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero))

            from Autodesk.Revit.DB import Transform
            
            if mode_2d_only:
                translator = Transform.CreateTranslation(XYZ(-cx, -cy, 0.0))
                for curve in curves:
                    try:
                        new_c = curve.CreateTransformed(translator)
                        fam_doc.FamilyCreate.NewModelCurve(new_c, sketch_plane)
                    except Exception:
                        pass
            else:
                from Autodesk.Revit.DB import (
                    FamilyElementVisibility, FamilyElementVisibilityType,
                    GraphicsStyleType, BuiltInParameter,
                )

                THICKNESS      = 0.1312
                HEIGHT         = 7.2178   # ≈2195 mm – default door height
                WINDOW_HEIGHT  = 4.9213   # ≈1500 mm – default window height
                extrusion_depth = HEIGHT if is_door else (WINDOW_HEIGHT if is_window else 1.0)

                # ── Door subcategories ──
                swing_gs = frame_gs = None
                if is_door:
                    try:
                        fam_cat = fam_doc.OwnerFamily.FamilyCategory
                        def get_or_create_subcat(name):
                            if fam_cat.SubCategories.Contains(name):
                                return fam_cat.SubCategories.get_Item(name)
                            return fam_doc.Settings.Categories.NewSubcategory(fam_cat, name)
                        swing_subcat = get_or_create_subcat("Plan Swing")
                        frame_subcat = get_or_create_subcat("Frame/Mullion")
                        if swing_subcat:
                            swing_gs = swing_subcat.GetGraphicsStyle(GraphicsStyleType.Projection)
                        if frame_subcat:
                            frame_gs = frame_subcat.GetGraphicsStyle(GraphicsStyleType.Projection)
                    except Exception:
                        pass

                # ── Family parameters ──
                param_height_fp = param_width_fp = param_material = None
                try:
                    fam_mgr = fam_doc.FamilyManager
                    for param in fam_mgr.Parameters:
                        pname = param.Definition.Name.lower()
                        if pname in ("height", "chiều cao"):
                            fam_mgr.Set(param, extrusion_depth)
                            param_height_fp = param
                        elif pname in ("width", "chiều rộng"):
                            if door_width:
                                fam_mgr.Set(param, door_width)
                            param_width_fp = param
                        elif pname in ("depth", "chiều sâu", "length", "chiều dài"):
                            if not is_door and half_h * 2.0 > 0.01:
                                fam_mgr.Set(param, half_h * 2.0)
                        elif pname in ("material", "vật liệu"):
                            param_material = param
                except Exception:
                    pass

                # ── Geometry ──
                ext_box = None

                if is_window:
                    # Window: hollow frame + glass pane via JSONtoFamily inner-loop approach.
                    # _create_window_body() uses CurveArrArray with outer + inner loops so the
                    # centre is cut away (same as JSONtoFamily 'inner_loops' on an Extrusion).
                    window_frame_ext, _window_glass = self._create_window_body(
                        fam_doc, sketch_plane,
                        half_w, half_h, extrusion_depth,
                        param_height_fp, param_material)
                    ext_box = window_frame_ext   # used below for face-locking

                elif not is_door:
                    # Bounding-rectangle extrusion for other non-door categories
                    c1 = XYZ(-half_w, -half_h, 0.0)
                    c2 = XYZ( half_w, -half_h, 0.0)
                    c3 = XYZ( half_w,  half_h, 0.0)
                    c4 = XYZ(-half_w,  half_h, 0.0)

                    rect = CurveArray()
                    rect.Append(Line.CreateBound(c1, c2))
                    rect.Append(Line.CreateBound(c2, c3))
                    rect.Append(Line.CreateBound(c3, c4))
                    rect.Append(Line.CreateBound(c4, c1))

                    profile = CurveArrArray()
                    profile.Append(rect)

                    ext_box = fam_doc.FamilyCreate.NewExtrusion(
                        True, profile, sketch_plane, extrusion_depth)
                    try:
                        if param_height_fp:
                            end_p = ext_box.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                            if end_p:
                                fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(
                                    end_p, param_height_fp)
                        if param_material:
                            mat_p = ext_box.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
                            if mat_p:
                                fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(
                                    mat_p, param_material)
                    except Exception:
                        pass

                top_sp = SketchPlane.Create(
                    fam_doc,
                    Plane.CreateByNormalAndOrigin(
                        XYZ.BasisZ,
                        XYZ(0.0, 0.0, extrusion_depth) if not is_door else XYZ.Zero))

                panel_ext = None
                for curve in curves:
                    try:
                        if is_door:
                            translator = Transform.CreateTranslation(XYZ(-cx, -cy, 0.0))
                            new_c = curve.CreateTransformed(translator)

                            if isinstance(curve, Line):
                                sym_line = fam_doc.FamilyCreate.NewSymbolicCurve(new_c, sketch_plane)
                                if frame_gs:
                                    sym_line.Subcategory = frame_gs

                            elif isinstance(curve, Arc):
                                sym_arc = fam_doc.FamilyCreate.NewSymbolicCurve(new_c, sketch_plane)
                                if swing_gs:
                                    sym_arc.Subcategory = swing_gs

                                # 3D panel extrusion in the "closed" direction
                                ctr = curve.Center
                                nc  = ctr + XYZ(-cx, -cy, 0.0)

                                p0_orig = curve.GetEndPoint(0)
                                p1_orig = curve.GetEndPoint(1)
                                p_closed_orig = (
                                    p0_orig if abs(p0_orig.Y - ctr.Y) < abs(p1_orig.Y - ctr.Y)
                                    else p1_orig)
                                np_closed = p_closed_orig + XYZ(-cx, -cy, 0.0)

                                v_dir   = (np_closed - nc).Normalize()
                                v_ortho = XYZ(-v_dir.Y, v_dir.X, 0.0)
                                half_t  = THICKNESS / 2.0

                                pt1 = nc + v_ortho * half_t
                                pt2 = nc - v_ortho * half_t
                                pt3 = pt2 + v_dir * curve.Radius
                                pt4 = pt1 + v_dir * curve.Radius

                                p_rect = CurveArray()
                                p_rect.Append(Line.CreateBound(pt1, pt2))
                                p_rect.Append(Line.CreateBound(pt2, pt3))
                                p_rect.Append(Line.CreateBound(pt3, pt4))
                                p_rect.Append(Line.CreateBound(pt4, pt1))

                                p_profile = CurveArrArray()
                                p_profile.Append(p_rect)

                                panel_ext = fam_doc.FamilyCreate.NewExtrusion(
                                    True, p_profile, sketch_plane, HEIGHT)

                                try:
                                    vis = FamilyElementVisibility(
                                        FamilyElementVisibilityType.Model)
                                    vis.IsShownInTopBottom = False
                                    panel_ext.SetVisibility(vis)

                                    if param_height_fp:
                                        end_p = panel_ext.get_Parameter(
                                            BuiltInParameter.EXTRUSION_END_PARAM)
                                        if end_p:
                                            fam_doc.FamilyManager\
                                                .AssociateElementParameterToFamilyParameter(
                                                    end_p, param_height_fp)
                                    if param_material:
                                        mat_p = panel_ext.get_Parameter(
                                            BuiltInParameter.MATERIAL_ID_PARAM)
                                        if mat_p:
                                            fam_doc.FamilyManager\
                                                .AssociateElementParameterToFamilyParameter(
                                                    mat_p, param_material)
                                except Exception:
                                    pass

                        elif not is_window:
                            # Window geometry is handled by _create_window_body(); skip model curves.
                            translator = Transform.CreateTranslation(
                                XYZ(-cx, -cy, extrusion_depth))
                            new_c = curve.CreateTransformed(translator)
                            fam_doc.FamilyCreate.NewModelCurve(new_c, top_sp)
                    except Exception:
                        pass

                # ── Parametric reference planes + dimensions (JSONtoFamily approach) ──
                # Applied to Door and Window so Width/Height params drive the geometry.
                if is_door or is_window:
                    fam_doc.Regenerate()
                    plan_view, elev_view = self._find_family_views(fam_doc)
                    rp_left, rp_right, rp_top = self._create_parametric_refs(
                        fam_doc,
                        half_w if not is_door else (door_width / 2.0 if door_width else half_w),
                        HEIGHT if is_door else extrusion_depth,
                        plan_view, elev_view,
                        param_width_fp, param_height_fp)

                    # Lock geometry faces to the new reference planes
                    fam_doc.Regenerate()
                    target_solid = panel_ext if is_door else ext_box
                    if target_solid and (rp_left or rp_right or rp_top):
                        self._lock_faces_to_planes(
                            fam_doc, target_solid,
                            plan_view, elev_view,
                            rp_left, rp_right, rp_top)

            t.Commit()
        except Exception:
            try:
                t.RollBack()
            except Exception:
                pass
            fam_doc.Close(False)
            raise

        # ── Save .rfa to output folder ────────────────────────────────────

        safe_cad_name = block_item.BlockName.strip() or "Family"
        base_name = "T3Lab_{}_{}".format(
            category_name.replace(" ", "_"),
            safe_cad_name.replace(" ", "_"))
        base_name = re.sub(r'[\\/*?:"<>|]', "", base_name)

        save_path = os.path.join(output_folder, "{}.rfa".format(base_name))
        counter = 1
        while os.path.exists(save_path):
            save_path = os.path.join(output_folder, "{}_{}.rfa".format(base_name, counter))
            counter += 1

        try:
            opts = SaveAsOptions()
            opts.OverwriteExistingFile = True
            fam_doc.SaveAs(save_path, opts)
        finally:
            fam_doc.Close(False)

        logger.info("Exported: {}".format(save_path))

        # ── Load to Project (absorbed from DWGtoFamily) ───────────────────

        if load_to_project:
            try:
                t_load = Transaction(doc, 'Load Family - {}'.format(safe_cad_name))
                t_load.Start()
                try:
                    doc.LoadFamily(save_path)
                    t_load.Commit()
                    logger.info("Loaded to project: {}".format(safe_cad_name))
                except Exception:
                    try:
                        t_load.RollBack()
                    except Exception:
                        pass
                    logger.warning("Could not load family to project: {}".format(save_path))
            except Exception:
                logger.warning("Load-to-project transaction failed: {}".format(
                    traceback.format_exc()))

        return True



    # ── Door Preset Generator (no CAD) ───────────────────────────────────

    def door_presets_clicked(self, sender, e):
        """Generate parametric door families from preset sizes. No CAD file needed."""
        output_folder = self.output_path.Text
        if not output_folder or not os.path.isdir(output_folder):
            forms.alert("Please select a valid output folder (Browse...) first.")
            return

        labels = [p[0] for p in DOOR_PRESETS]
        selected_labels = forms.SelectFromList.show(
            labels,
            title="T3Lab - Select Door Types to Generate",
            multiselect=True,
            button_name="Generate")
        if not selected_labels:
            return

        selected = [p for p in DOOR_PRESETS if p[0] in selected_labels]
        load_to_project = (self.chk_load_to_project.IsChecked == True)

        door_template = self._find_template(1)   # index 1 = "Door" in CATEGORY_TEMPLATES
        if not door_template:
            forms.alert("Door template (.rft) not found. Install Revit family templates.")
            return

        ok_count = fail_count = 0
        for preset in selected:
            self._update_status("Generating: " + preset[0])
            try:
                if self._generate_door_from_preset(preset, door_template, output_folder, load_to_project):
                    ok_count += 1
                else:
                    fail_count += 1
            except Exception:
                logger.error("Door preset error: {}\n{}".format(preset[0], traceback.format_exc()))
                fail_count += 1

        self._update_status("Done: {} ok, {} failed".format(ok_count, fail_count))
        forms.alert(
            "Door Preset Export\n\n"
            "Generated: {}\nFailed: {}\n\n"
            "Folder: {}".format(ok_count, fail_count, output_folder))

    def _generate_door_from_preset(self, preset, template_path, output_folder, load_to_project):
        """Build a parametric door .rfa purely from geometry (no CAD).

        Frame geometry (tested & working):
          - Uses the exterior sketch plane (normal = -Y, at origin) found in Metric Door.rft.
          - 3 separate rectangular extrusions per side:
              left jamb  : x[-half_fw..-half_w], full height including head
              right jamb : x[half_w..half_fw],   full height including head
              top head   : x[-half_fw..half_fw],  z[door_h..total_h]
          - positive depth  => exterior projection (depth = proj_ext_mm)
          - negative depth  => interior projection (depth = -proj_int_mm)
        Leaf geometry:
          - Plan sketch plane (normal = +Z)
          - Rectangle profile in XY, extruded in Z to full door height
          - Single: one leaf; Double: two leaves with 1mm gap
        """
        label, width_mm, height_mm, frame_w_mm, proj_ext_mm, proj_int_mm, leaf_t_mm, door_count = preset

        half_w   = (width_mm  / 2.0) * SCL
        h        = height_mm         * SCL
        fw       = frame_w_mm        * SCL
        fpe      = proj_ext_mm       * SCL
        fpi      = proj_int_mm       * SCL
        dt       = leaf_t_mm         * SCL
        half_fw  = half_w + fw
        total_fh = h + fw

        fam_doc = app.NewFamilyDocument(template_path)
        t = Transaction(fam_doc, "T3Lab Door - " + label)
        t.Start()
        try:
            # ── Find sketch planes from template ──
            plan_sp = elev_sp = None
            for sp in FilteredElementCollector(fam_doc).OfClass(SketchPlane):
                n, org = sp.GetPlane().Normal, sp.GetPlane().Origin
                if abs(n.Z - 1.0) < 0.001 and abs(org.X) < 0.01 and abs(org.Y) < 0.01 and plan_sp is None:
                    plan_sp = sp
                if abs(n.Y + 1.0) < 0.001 and abs(org.X) < 0.01 and abs(org.Y) < 0.01 and elev_sp is None:
                    elev_sp = sp

            if plan_sp is None:
                plan_sp = SketchPlane.Create(fam_doc, Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero))
            if elev_sp is None:
                elev_sp = SketchPlane.Create(fam_doc, Plane.CreateByNormalAndOrigin(XYZ(0.0,-1.0,0.0), XYZ.Zero))

            # ── Set Width / Height parameters ──
            param_width_fp = param_height_fp = None
            try:
                fm = fam_doc.FamilyManager
                for p in fm.Parameters:
                    pn = p.Definition.Name.lower()
                    if pn == "height":
                        try: fm.Set(p, h)
                        except Exception: pass
                        param_height_fp = p
                    elif pn == "width":
                        try: fm.Set(p, half_w * 2.0)
                        except Exception: pass
                        param_width_fp = p
                    elif pn == "frame width":
                        try: fm.Set(p, fw)
                        except Exception: pass
                    elif pn in ("frame projection ext.", "frame projection ext"):
                        try: fm.Set(p, fpe)
                        except Exception: pass
                    elif pn in ("frame projection int.", "frame projection int"):
                        try: fm.Set(p, fpi)
                        except Exception: pass
            except Exception:
                pass

            # ── Subcategories ──
            frame_gs = leaf_gs = None
            try:
                from Autodesk.Revit.DB import GraphicsStyleType
                fam_cat = fam_doc.OwnerFamily.FamilyCategory
                def _sc(name):
                    return (fam_cat.SubCategories.get_Item(name)
                            if fam_cat.SubCategories.Contains(name)
                            else fam_doc.Settings.Categories.NewSubcategory(fam_cat, name))
                sc_f = _sc("Frame/Mullion")
                sc_p = _sc("Panel")
                if sc_f: frame_gs = sc_f.GetGraphicsStyle(GraphicsStyleType.Projection)
                if sc_p: leaf_gs  = sc_p.GetGraphicsStyle(GraphicsStyleType.Projection)
            except Exception:
                pass

            # ── Helpers ──
            def _extrude_xz(x0, x1, z0, z1, depth, gs=None):
                """Rect profile in XZ plane, extruded in Y (+depth=ext, -depth=int)."""
                arr = CurveArray()
                arr.Append(Line.CreateBound(XYZ(x0,0,z0), XYZ(x1,0,z0)))
                arr.Append(Line.CreateBound(XYZ(x1,0,z0), XYZ(x1,0,z1)))
                arr.Append(Line.CreateBound(XYZ(x1,0,z1), XYZ(x0,0,z1)))
                arr.Append(Line.CreateBound(XYZ(x0,0,z1), XYZ(x0,0,z0)))
                prof = CurveArrArray()
                prof.Append(arr)
                ext = fam_doc.FamilyCreate.NewExtrusion(True, prof, elev_sp, depth)
                if gs:
                    try: ext.Subcategory = gs
                    except Exception: pass
                return ext

            def _extrude_xy_leaf(x0, x1, y0, y1):
                """Leaf profile in XY plan, extruded in Z to full door height."""
                arr = CurveArray()
                pts = [XYZ(x0,y0,0), XYZ(x1,y0,0), XYZ(x1,y1,0), XYZ(x0,y1,0)]
                for i in range(4): arr.Append(Line.CreateBound(pts[i], pts[(i+1)%4]))
                prof = CurveArrArray()
                prof.Append(arr)
                ext = fam_doc.FamilyCreate.NewExtrusion(True, prof, plan_sp, h)
                if leaf_gs:
                    try: ext.Subcategory = leaf_gs
                    except Exception: pass
                try:
                    if param_height_fp:
                        from Autodesk.Revit.DB import BuiltInParameter
                        ep = ext.get_Parameter(BuiltInParameter.EXTRUSION_END_PARAM)
                        if ep: fam_doc.FamilyManager.AssociateElementParameterToFamilyParameter(ep, param_height_fp)
                except Exception:
                    pass
                return ext

            # Frame pieces — 3 per side (left jamb, right jamb, top head)
            frame_pieces = [
                (-half_fw, -half_w, 0.0, total_fh),   # left jamb (full height incl. head)
                ( half_w,  half_fw, 0.0, total_fh),   # right jamb
                (-half_fw,  half_fw, h,   total_fh),   # top head only
            ]
            for x0, x1, z0, z1 in frame_pieces:
                _extrude_xz(x0, x1, z0, z1,  fpe, frame_gs)   # exterior
                _extrude_xz(x0, x1, z0, z1, -fpi, frame_gs)   # interior

            # Door leaf(ves)
            GAP = 0.00328   # 1 mm in feet
            if door_count == 1:
                _extrude_xy_leaf(-half_w, half_w, 0.0, dt)
            else:
                _extrude_xy_leaf(-half_w, -GAP/2.0, 0.0, dt)
                _extrude_xy_leaf( GAP/2.0,  half_w, 0.0, dt)

            # Parametric reference planes
            fam_doc.Regenerate()
            plan_view, elev_view = self._find_family_views(fam_doc)
            self._create_parametric_refs(
                fam_doc, half_w, h, plan_view, elev_view,
                param_width_fp, param_height_fp)

            t.Commit()
        except Exception:
            try: t.RollBack()
            except Exception: pass
            fam_doc.Close(False)
            raise

        # ── Save ──
        safe = re.sub(r'[/*?:"<>|]', "_", label)
        save_path = os.path.join(output_folder, "T3Lab_Door_{}.rfa".format(safe))
        ctr = 1
        while os.path.exists(save_path):
            save_path = os.path.join(output_folder, "T3Lab_Door_{}_{}.rfa".format(safe, ctr))
            ctr += 1

        try:
            opts = SaveAsOptions()
            opts.OverwriteExistingFile = True
            fam_doc.SaveAs(save_path, opts)
        finally:
            fam_doc.Close(False)

        logger.info("Saved: " + save_path)

        if load_to_project:
            try:
                t2 = Transaction(doc, "Load " + label)
                t2.Start()
                try: doc.LoadFamily(save_path); t2.Commit()
                except Exception:
                    try: t2.RollBack()
                    except Exception: pass
            except Exception:
                pass

        return True


    # ── Window chrome handlers ────────────────────────────────────────────

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


# ─── Entry point ──────────────────────────────────────────────────────────────

if __name__ == '__main__':
    try:
        window = BulkFamilyExportWindow()
        window.ShowDialog()
    except Exception:
        logger.error(traceback.format_exc())
        forms.alert("Unexpected error. Check the pyRevit log for details.")
