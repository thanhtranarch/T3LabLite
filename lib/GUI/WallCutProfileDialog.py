# -*- coding: utf-8 -*-
"""Wall Cut Profile — Dialog and execution logic for Wall Profile Cut & Openings.

Cutting elements always come from the active model; ticking "Include linked
models" adds the ticked Revit links (several at once).
"""

import os
import sys

from pyrevit import forms


import System

from GUI.WPF_Base import T3WPFWindow, set_items_source
from Snippets._compat import eid_value
from Snippets._units import project_length_unit

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'WallCutProfile.xaml')
_TITLE = "Wall Cut Profile"

try:
    import Autodesk.Revit.DB as DB
    from Autodesk.Revit.DB import (
        FilteredElementCollector,
        BuiltInCategory,
        BuiltInParameter,
        Transaction,
        ElementId,
        Wall,
        RevitLinkInstance,
        FamilySymbol,
        XYZ,
        Line,
        IFailuresPreprocessor,
        FailureProcessingResult,
        FailureSeverity,
    )
    from Autodesk.Revit.UI.Selection import ObjectType, ISelectionFilter
except Exception:
    DB = None


import uuid
# The pushbutton script importlib.reload()s this module on every click; a
# fixed __namespace__ would define the same .NET type twice and raise
# "Duplicate type name within an assembly" (rule S15).
_NS_SUFFIX = uuid.uuid4().hex[:8]


def _plural(count, word):
    return u"{} {}{}".format(count, word, "" if count == 1 else "s")


class LinkFilter(ISelectionFilter if DB else object):
    __namespace__ = "T3Lab.WallCutProfile_LinkFilter_" + _NS_SUFFIX

    def AllowElement(self, elem):
        # Unloaded links have no elements to cut with.
        return isinstance(elem, RevitLinkInstance) and elem.GetLinkDocument() is not None

    def AllowReference(self, ref, point):
        return False


class WallFilter(ISelectionFilter if DB else object):
    __namespace__ = "T3Lab.WallCutProfile_WallFilter_" + _NS_SUFFIX

    def AllowElement(self, elem):
        if not isinstance(elem, Wall):
            return False
        if elem.WallType and elem.WallType.Kind in (DB.WallKind.Curtain, DB.WallKind.Stacked):
            return False
        loc = elem.Location
        return loc is not None and hasattr(loc, 'Curve') and loc.Curve is not None

    def AllowReference(self, ref, point):
        return False


class WarningSwallower(IFailuresPreprocessor if DB else object):
    __namespace__ = "T3Lab.WallCutProfile_Failures_" + _NS_SUFFIX

    def PreprocessFailures(self, fa):
        for f in fa.GetFailureMessages():
            sev = f.GetSeverity()
            if sev == FailureSeverity.Warning:
                fa.DeleteWarning(f)
        return FailureProcessingResult.Continue


class OpeningFamilyItem(object):
    def __init__(self, symbol):
        self.symbol = symbol
        self.symbol_id = symbol.Id
        try:
            fam_name = symbol.Family.Name
        except Exception:
            fam_name = "Unknown"
        try:
            p = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
            type_name = p.AsString() if p else symbol.Name
        except Exception:
            type_name = symbol.Name
        self.name = "{} : {}".format(fam_name, type_name)

    def __str__(self):
        return self.name


class LinkItem(object):
    """One loaded Revit link instance in the SOURCE MODELS list."""

    def __init__(self, instance, title, path=u""):
        self.instance = instance
        self.title = title
        self.path = path or title
        self.checked = True


class WallCutProfileWindow(T3WPFWindow):
    def __init__(self, doc, uidoc):
        T3WPFWindow.__init__(self, _XAML)
        self._doc = doc
        self._uidoc = uidoc
        self._links = []
        self._link_boxes = []
        self._picked_walls = []
        self._opening_families = []
        # Read per window: the module stays loaded across projects.
        self._unit = project_length_unit(doc)

        self._apply_units()
        self._init_controls()
        self._load_links()
        self._load_opening_families()
        self._update_wall_status()
        self._update_status()

    def _apply_units(self):
        """Clearance label and default (25 mm) in the project's length unit."""
        self.lbl_offset.Text = self._unit.label("CLEARANCE")
        self.txt_offset.Text = self._unit.default_text(25)

    def _init_controls(self):
        methods = ["Place Opening Family", "Edit Wall Profile", "Wall Opening"]
        set_items_source(self.cmb_method, methods)
        self.cmb_method.SelectedIndex = 0
        title = getattr(self._doc, 'Title', None)
        if title:
            self.txt_host_model.Text = u"This model ({}) is always scanned".format(title)

    # ── source models ───────────────────────────────────────────────────────
    def _load_links(self):
        """Loaded link instances, all ticked; copies of one link get their instance name."""
        items = []
        if self._doc:
            for inst in FilteredElementCollector(self._doc).OfClass(RevitLinkInstance):
                ldoc = inst.GetLinkDocument()
                if ldoc is None:
                    continue
                items.append(LinkItem(inst, ldoc.Title, getattr(ldoc, 'PathName', u"")))
        titles = [i.title for i in items]
        for item in items:
            if titles.count(item.title) > 1:
                item.title = item.instance.Name
        self._links = sorted(items, key=lambda i: i.title.lower())
        self._fill_link_boxes()
        self._refresh_link_count()

    def _fill_link_boxes(self):
        from System.Windows import TextTrimming
        from System.Windows.Controls import CheckBox, TextBlock

        try:
            style = self.FindResource("T3.Filter.Item")
        except Exception:
            style = None
        self.pnl_links.Children.Clear()
        self._link_boxes = []
        for item in self._links:
            box = CheckBox()
            if style is not None:
                box.Style = style
            # A TextBlock, not a string: CheckBox content hides the first "_" (access key).
            label = TextBlock()
            label.Text = item.title
            label.TextTrimming = TextTrimming.CharacterEllipsis
            box.Content = label
            box.ToolTip = item.path
            box.IsChecked = item.checked
            box.Click += self._link_box_clicked
            self.pnl_links.Children.Add(box)
            self._link_boxes.append((item, box))
        vis = System.Windows.Visibility
        self.txt_links_empty.Visibility = vis.Collapsed if self._links else vis.Visible

    def _ticked_links(self):
        return [item for item in self._links if item.checked]

    def _refresh_link_count(self):
        self.txt_link_count.Text = u"{} of {} ticked".format(
            len(self._ticked_links()), _plural(len(self._links), "linked model"))

    def _sources(self):
        """[(document, transform)] to take cutting elements from; transform None = this model."""
        sources = [(self._doc, None)]
        if self.chk_include_links.IsChecked:
            for item in self._ticked_links():
                ldoc = item.instance.GetLinkDocument()
                if ldoc is not None:
                    sources.append((ldoc, item.instance.GetTotalTransform()))
        return sources

    def include_links_changed(self, sender, e):
        vis = System.Windows.Visibility
        self.pnl_link_area.Visibility = vis.Visible if self.chk_include_links.IsChecked else vis.Collapsed
        self._update_status()

    def _link_box_clicked(self, sender, e):
        for item, box in self._link_boxes:
            item.checked = bool(box.IsChecked)
        self._refresh_link_count()
        self._update_status()

    def btn_pick_links_clicked(self, sender, e):
        if not self._uidoc or not self._links:
            return
        refs = None
        self.Hide()
        try:
            refs = self._uidoc.Selection.PickObjects(
                ObjectType.Element,
                LinkFilter(),
                "Select linked models, then click Finish"
            )
        except Exception:
            pass    # Esc cancels the pick
        finally:
            self.Show()
        if not refs:
            return
        picked = set(eid_value(r.ElementId) for r in refs)
        for item, box in self._link_boxes:
            item.checked = eid_value(item.instance.Id) in picked
            box.IsChecked = item.checked
        self._refresh_link_count()
        self._update_status()

    # ── target walls ────────────────────────────────────────────────────────
    def btn_pick_walls_clicked(self, sender, e):
        if not self._uidoc:
            return
        self.Hide()
        try:
            refs = self._uidoc.Selection.PickObjects(
                ObjectType.Element,
                WallFilter(),
                "Select target walls, then click Finish"
            )
            walls = []
            if refs:
                for r in refs:
                    w = self._doc.GetElement(r.ElementId)
                    if w:
                        walls.append(w)
            self._picked_walls = walls
            self.rb_walls_sel.IsChecked = True
        except Exception:
            pass
        finally:
            self.Show()
        self._update_wall_status()
        self._update_status()

    def wall_scope_changed(self, sender, e):
        self._update_wall_status()
        self._update_status()

    def _update_wall_status(self):
        if self.rb_walls_sel.IsChecked:
            n = len(self._picked_walls)
            self.txt_wall_status.Text = (_plural(n, "wall") + " picked" if n
                                         else "No walls picked yet - click Pick Walls")
            return
        try:
            view = self._doc.ActiveView.Name
        except Exception:
            view = None
        self.txt_wall_status.Text = u"Active view: {}".format(view) if view else "Walls in the active view"

    def _update_status(self):
        """Footer line: what Create Openings will scan."""
        if self.chk_include_links.IsChecked:
            n = len(self._ticked_links())
            src = (u"This model + " + _plural(n, "linked model")) if n else u"No linked model ticked"
        else:
            src = u"This model only"
        if self.rb_walls_sel.IsChecked:
            walls = _plural(len(self._picked_walls), "picked wall")
        else:
            walls = u"walls in current view"
        self.txt_status.Text = u"{} · {}".format(src, walls)

    # ── openings ────────────────────────────────────────────────────────────
    def _load_opening_families(self):
        if not self._doc:
            return
        results = []
        cats = [
            BuiltInCategory.OST_GenericModel,
            BuiltInCategory.OST_Windows,
            BuiltInCategory.OST_Doors,
        ]
        for bic in cats:
            try:
                coll = FilteredElementCollector(self._doc).OfCategory(bic).OfClass(FamilySymbol)
                for sym in coll:
                    fam = sym.Family
                    if not fam:
                        continue
                    p = fam.get_Parameter(BuiltInParameter.FAMILY_HOSTING_BEHAVIOR)
                    if p and p.AsInteger() == 1:
                        results.append(OpeningFamilyItem(sym))
                    elif any(kw in fam.Name.lower() for kw in ["opening", "void", "cutout", "penetration", "sleeve"]):
                        results.append(OpeningFamilyItem(sym))
            except Exception:
                pass

        seen = set()
        unique = []
        for item in results:
            sid = eid_value(item.symbol_id)
            if sid not in seen:
                seen.add(sid)
                unique.append(item)

        self._opening_families = sorted(unique, key=lambda x: x.name)
        set_items_source(self.cmb_families, [f.name for f in self._opening_families])
        if len(self._opening_families) > 0:
            self.cmb_families.SelectedIndex = 0

    def cmb_method_changed(self, sender, e):
        method = str(self.cmb_method.SelectedItem or "")
        if "Place Opening Family" in method:
            self.panel_family.Visibility = System.Windows.Visibility.Visible
        else:
            self.panel_family.Visibility = System.Windows.Visibility.Collapsed

    def btn_apply_clicked(self, sender, e):
        if self.chk_include_links.IsChecked and not self._ticked_links():
            forms.alert("Include linked models is on but no linked model is ticked. "
                        "Tick at least one link, or clear Include linked models to scan only this model.",
                        title=_TITLE)
            return
        sources = self._sources()

        # Get walls
        walls = []
        if self.rb_walls_sel.IsChecked:
            walls = self._picked_walls
            if not walls:
                forms.alert("No walls picked. Click Pick Walls, or choose 'All walls in current view'.",
                            title=_TITLE)
                return
        else:
            coll = FilteredElementCollector(self._doc, self._doc.ActiveView.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType()
            w_filter = WallFilter()
            for eid in coll.ToElementIds():
                w = self._doc.GetElement(eid)
                if w_filter.AllowElement(w):
                    walls.append(w)

        if not walls:
            forms.alert("No walls to cut in the current view (curtain and stacked walls are skipped). "
                        "Open a view that shows the walls, or pick them with Pick Walls.", title=_TITLE)
            return

        # Selected categories
        sel_cats = []
        if self.chk_ducts.IsChecked: sel_cats.append(("Ducts", BuiltInCategory.OST_DuctCurves))
        if self.chk_pipes.IsChecked: sel_cats.append(("Pipes", BuiltInCategory.OST_PipeCurves))
        if self.chk_trays.IsChecked: sel_cats.append(("Cable Trays", BuiltInCategory.OST_CableTray))
        if self.chk_framing.IsChecked: sel_cats.append(("Structural Framing", BuiltInCategory.OST_StructuralFraming))
        if self.chk_columns.IsChecked: sel_cats.append(("Columns", BuiltInCategory.OST_Columns))

        if not sel_cats:
            forms.alert("Tick at least one intersecting category.", title=_TITLE)
            return

        # Bare number = project unit; "25 mm" or "1\"" work in any project.
        try:
            offset_ft = self._unit.parse(self.txt_offset.Text)
        except ValueError as ex:
            msg = "Clearance: {}".format(ex)
            self.txt_status.Text = msg
            forms.alert(msg, title=_TITLE)
            return

        method = str(self.cmb_method.SelectedItem or "Place Opening Family")
        sel_symbol = None
        if "Place Opening Family" in method:
            fam_idx = self.cmb_families.SelectedIndex
            if fam_idx >= 0 and fam_idx < len(self._opening_families):
                sel_symbol = self._opening_families[fam_idx].symbol
            else:
                forms.alert("Choose an opening family. Load a wall-hosted void family "
                            "if the list is empty.", title=_TITLE)
                return

        self.txt_status.Text = "Scanning intersections..."
        created, errors = self._execute_openings(walls, sources, sel_cats, method, offset_ft, sel_symbol)

        links = len(sources) - 1
        scanned = "this model" + (" + " + _plural(links, "linked model") if links else "")
        msg = "Created {} in {} (scanned {}).".format(
            _plural(created, "opening"), _plural(len(walls), "wall"), scanned)
        if errors > 0:
            msg += "\nFailed: {}.".format(_plural(errors, "item"))
        self.Close()
        forms.alert(msg, title="Wall Profile Cut Results")

    def _execute_openings(self, walls, sources, categories, method, offset_ft, symbol):
        """sources: [(document, transform or None)] — see _sources()."""
        created = 0
        errors = 0

        # Bounding boxes of the cutting elements, in this model's coordinates
        cutters = []
        for src_doc, transform in sources:
            for cat_name, bic in categories:
                try:
                    coll = FilteredElementCollector(src_doc).OfCategory(bic).WhereElementIsNotElementType()
                    for el in coll:
                        bb = el.get_BoundingBox(None)
                        if not bb:
                            continue
                        pt_min, pt_max = bb.Min, bb.Max
                        center = (pt_min + pt_max) * 0.5
                        if transform is not None:
                            pt_min = transform.OfPoint(pt_min)
                            pt_max = transform.OfPoint(pt_max)
                            center = transform.OfPoint(center)
                        cutters.append({
                            'elem': el,
                            'cat': cat_name,
                            'center': center,
                            'min': pt_min,
                            'max': pt_max,
                            'dx': abs(pt_max.X - pt_min.X),
                            'dy': abs(pt_max.Y - pt_min.Y),
                            'dz': abs(pt_max.Z - pt_min.Z),
                        })
                except Exception:
                    pass

        t = Transaction(self._doc, "T3Lab: Wall Cut Profile")
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(WarningSwallower())
        t.SetFailureHandlingOptions(opts)

        t.Start()
        try:
            if symbol and not symbol.IsActive:
                symbol.Activate()
                self._doc.Regenerate()

            for w in walls:
                w_bb = w.get_BoundingBox(None)
                if not w_bb:
                    continue
                w_min = w_bb.Min
                w_max = w_bb.Max

                for item in cutters:
                    # Bounding box overlap test
                    i_c = item['center']
                    if not (w_min.X - offset_ft <= i_c.X <= w_max.X + offset_ft and
                            w_min.Y - offset_ft <= i_c.Y <= w_max.Y + offset_ft and
                            w_min.Z - offset_ft <= i_c.Z <= w_max.Z + offset_ft):
                        continue

                    # Overlapping: place opening
                    try:
                        if symbol and "Place Opening Family" in method:
                            inst = self._doc.Create.NewFamilyInstance(i_c, symbol, w, DB.Structure.StructuralType.NonStructural)
                            # Set parameters if available
                            for p_name, val in [("Width", max(item['dx'], item['dy']) + offset_ft * 2),
                                                ("Height", item['dz'] + offset_ft * 2)]:
                                p = inst.LookupParameter(p_name)
                                if p and not p.IsReadOnly:
                                    p.Set(val)
                            created += 1
                        elif "Wall Opening" in method:
                            # Create rectangular opening
                            p0 = XYZ(i_c.X - 0.5, i_c.Y - 0.5, i_c.Z - item['dz']*0.5 - offset_ft)
                            p1 = XYZ(i_c.X + 0.5, i_c.Y + 0.5, i_c.Z + item['dz']*0.5 + offset_ft)
                            self._doc.Create.NewOpening(w, p0, p1)
                            created += 1
                    except Exception:
                        errors += 1

            t.Commit()
        except Exception:
            try:
                t.RollBack()
            except Exception:
                pass
            # Rolled back: nothing from this run is left in the model.
            return 0, errors + 1

        return created, errors


def show_wall_cut_profile(doc, uidoc):
    if not doc or not uidoc:
        forms.alert("No active Revit document found.", title=_TITLE)
        return
    win = WallCutProfileWindow(doc, uidoc)
    win.ShowDialog()
