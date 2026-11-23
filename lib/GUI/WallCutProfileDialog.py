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
# Revit side shared with the cut_wall_openings tool (core/tools/model.py).
from Snippets.wall_tools import (execute_openings, is_cut_wall, opening_symbols,
                                 symbol_label)

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
        return is_cut_wall(elem)

    def AllowReference(self, ref, point):
        return False


class OpeningFamilyItem(object):
    def __init__(self, symbol):
        self.symbol = symbol
        self.symbol_id = symbol.Id
        self.name = symbol_label(symbol)

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
        unique = [OpeningFamilyItem(sym) for sym in opening_symbols(self._doc)]
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
        return execute_openings(self._doc, walls, sources, categories, method, offset_ft, symbol)


def show_wall_cut_profile(doc, uidoc):
    if not doc or not uidoc:
        forms.alert("No active Revit document found.", title=_TITLE)
        return
    win = WallCutProfileWindow(doc, uidoc)
    win.ShowDialog()
