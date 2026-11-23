# -*- coding: utf-8 -*-
"""
CAD to Elements — convert CAD linework into Revit elements.

One window, nine modes (rail tiles), ONE fixed layout:

    A  CAD SOURCE | LEVEL          shared by every mode, selection persists
    B  mode title + description
    C  options card                same row skeleton in every mode
    D  CAD layers                  same toolbar / list / tally in every mode

Switching mode only swaps which `opt_<mode>` grid is visible inside the fixed
options card and re-ticks that mode's own layer selection — nothing moves.

Modes: Walls · Floors · Ceilings · Rooms · Columns · Beams · Grids · Lines ·
MEP Runs (ducts, pipes, cable trays, conduits).

Rules deciding WHAT gets created live in `Snippets/_cad_geometry.py` (pure,
unit-tested); Revit calls in `Snippets/_cad_revit.py` (one transaction per
run). Every run shows the count first (P5 confirm), then a result with counts.

Units: every length the user types or reads is in the project's length unit
(Manage > Project Units > Length), read each time the window opens through
`Snippets/_units.py`. Labels get the unit ("HEIGHT (MM)", "HEIGHT (FT-IN)"),
defaults are written in it, and typed text is parsed by it (a bare number is
in the project unit; "1200 mm", "1.2 m" or 3'-6" always work). Column size
rounding and the per-size column type names use mm / inches (`paper_unit`).

Copyright (c) 2026 T3Lab
"""
import os

from System.Windows import Visibility, RoutedEventHandler
from System.Windows.Controls import CheckBox, ComboBoxItem
from System.Windows.Input import Key

import Autodesk.Revit.DB as DB
from pyrevit import revit

from GUI.WPF_Base import T3WPFWindow
from GUI.T3Dialog import show_info, show_warning, show_error, confirm
from Snippets import _cad_geometry as geo
from Snippets import _cad_revit as cr
from Snippets import _cad_plan as cp
from Snippets._compat import eid_value
from Snippets._units import project_length_unit, paper_unit

_XAML_HUB = os.path.join(os.path.dirname(__file__), "Tools", "CADToElements.xaml")
TITLE = "CAD to Elements"

ROUND_SKIP = {"element": None, "name": "(Skip circles)"}
UNCONNECTED = {"element": None, "name": "Unconnected (use height)"}

# Every model length of the options card except the MEP ones (those follow
# geo.MEP_CATEGORIES): TextBox, label without the unit, default kept in mm.
# When the window opens the label `lbl_<same suffix>` gets the project unit and
# the TextBox the default written in that unit.
LENGTH_FIELDS = (
    ("txt_wall_height", "HEIGHT", 3000),
    ("txt_wall_offset", "BASE OFFSET", 0),
    ("txt_wall_thickness", "UNPAIRED THICKNESS", 200),
    ("txt_floor_offset", "HEIGHT OFFSET", 0),
    ("txt_part_thickness", "PART THICKNESS", 200),
    ("txt_ceiling_offset", "HEIGHT OFFSET", 2700),
    ("txt_ceiling_part_thickness", "PART THICKNESS", 25),
    ("txt_column_height", "HEIGHT", 3000),
    ("txt_beam_offset", "TOP OFFSET", -50),
    ("txt_beam_min_width", "MIN WIDTH", 50),
    ("txt_beam_max_width", "MAX WIDTH", 1500),
    ("txt_grid_extend", "EXTEND ENDS", 0),
    ("txt_grid_min_length", "MIN LENGTH", 1000),
    ("txt_lines_offset", "OFFSET FROM LEVEL", 0),
)

# Column SIZE ROUNDING is a step for type sizes, which are written in mm or
# inches (`paper_unit`), never in m / ft: 10 mm in decimal feet reads "0.03"
# (9.1 mm) and would turn a 400 mm column into 402 mm. Default per unit tag.
COLUMN_ROUNDING_DEFAULT = {"mm": 10, "in": 0.5}


def label_name(field):
    """txt_wall_height -> lbl_wall_height (the label above the TextBox)."""
    return "lbl_" + field[len("txt_"):]


def with_unit(unit, feet):
    """A length for a sentence: '2800 mm', '2.8 m', 9' - 2 1/4" (ft-in and
    fractional inches already carry their ' and " marks)."""
    text = unit.text(feet)
    return text if unit.style != "decimal" else u"{} {}".format(text, unit.tag)


class _InputError(Exception):
    """A field the user has to fix; the message is shown as-is."""


class LayerRow(object):
    """One CAD layer in the shared list. `count` is text for the grid."""

    def __init__(self, geom):
        self.geom = geom
        self.name = geom.name
        self.is_selected = False
        self.count = "0"


def _counted(n, noun):
    """'1 wall' / '3 walls' — nouns are given in the plural."""
    if n == 1 and noun.endswith("s"):
        noun = noun[:-1]
    return u"{} {}".format(n, noun)


class CADToElementsWindow(T3WPFWindow):
    AI_TOOL = "CADToElements"

    # ------------------------------------------------------------------
    # Construction — everything is loaded here (rule S7), not on Loaded
    # ------------------------------------------------------------------

    def __init__(self):
        T3WPFWindow.__init__(self, _XAML_HUB)
        self._doc = revit.doc
        self._uidoc = revit.uidoc
        # Read on every open, never at import: the module outlives the project.
        self._unit = project_length_unit(self._doc)
        self._size_unit = paper_unit(self._unit)     # column type sizes: mm / inches
        self._mode = "wall"
        self._mep_key = "duct"
        self._mep_values = {}
        self._mep_types = {}
        self._mep_systems = {}
        self._column_structural = None
        self._cad_list = []
        self._levels = []
        self._rows = []
        self._selected = dict((k, set()) for k in geo.MODE_KEYS)
        self._lists = {}
        self._busy = True

        self._tiles = dict((k, getattr(self, "btn_mode_" + k)) for k in geo.MODE_KEYS)
        self._panels = dict((k, getattr(self, "opt_" + k)) for k in geo.MODE_KEYS)

        # Row checkbox clicks bubble to the grid; the bridge has already
        # written the row by then, so the tally reads the new state.
        self._row_click_handler = RoutedEventHandler(self._on_row_checkbox_click)
        self.grid_layers.AddHandler(CheckBox.ClickEvent, self._row_click_handler, True)
        self.PreviewKeyDown += self._on_key_down

        self._apply_units()
        self._load_all(scan=False)
        self._busy = False
        self._apply_mode("wall")
        self.init_ai_badge()
        # The DWG scan grows with the drawing (4.6 s before the window showed,
        # 2026-11-17 open_times.log): show the window first, then read layers.
        self._first_scan_handler = self._on_first_render
        self.ContentRendered += self._first_scan_handler

    def _on_first_render(self, sender, e):
        try:
            self.ContentRendered -= self._first_scan_handler
        except Exception:
            pass
        self._busy = True
        try:
            self._scan_current_cad()
        finally:
            self._busy = False
        self._apply_mode(self._mode)

    # ------------------------------------------------------------------
    # Units — labels, defaults and hints in the project's length unit
    # ------------------------------------------------------------------

    def _apply_units(self):
        unit, size_unit = self._unit, self._size_unit
        for field, label, default_mm in LENGTH_FIELDS:
            getattr(self, label_name(field)).Text = unit.label(label)
            getattr(self, field).Text = unit.default_text(default_mm)
        self.lbl_column_rounding.Text = size_unit.label("SIZE ROUNDING")
        self.txt_column_rounding.Text = size_unit.text(
            size_unit.to_feet(COLUMN_ROUNDING_DEFAULT.get(size_unit.tag, 10)))
        self.lbl_mep_height.Text = unit.label("HEIGHT")
        self.lbl_mep_offset.Text = unit.label("OFFSET FROM LEVEL")
        self.txt_wall_hint.Text = (u"Two parallel lines up to {} apart make one wall on their "
                                   u"centerline.".format(self._len(geo.MAX_WALL_THICKNESS)))
        self.txt_column_hint.Text = (u"Shapes from {} to {} across are read; a column drawn "
                                     u"twice is placed once.".format(
                                         self._len(geo.mm(geo.FOOTPRINT_MIN_MM)),
                                         self._len(geo.mm(geo.FOOTPRINT_MAX_MM))))

    def _len(self, feet, unit=None):
        """A length (feet) as the user reads it: '2800 mm' / 9' - 2 1/4"."""
        return with_unit(unit or self._unit, feet)

    # ------------------------------------------------------------------
    # Loading
    # ------------------------------------------------------------------

    def _load_all(self, scan=True):
        keep_cad = self._current_cad_id()
        keep_level = self._current_level_id()
        self._cad_list = cr.get_cad_instances(self._doc)
        self._fill("cmb_cad_files", self._cad_list, "No CAD file in this model",
                   default=self._index_of(self._cad_list, keep_cad))
        self._levels = cr.get_levels(self._doc)
        for lv in self._levels:
            lv["label"] = u"{} ({})".format(lv["name"], self._len(lv["elevation"]))
        self._fill("cmb_levels", self._levels, "No level in this model", label="label",
                   default=self._default_level_index(keep_level))
        self._fill_types()
        self._fill_level_dependent()
        if scan:
            self._scan_current_cad()

    def _fill_types(self):
        doc = self._doc
        self._fill("cmb_wall_type", cr.get_basic_wall_types(doc), "No basic wall type loaded",
                   prefer=("generic",))
        self._fill("cmb_floor_type", cr.get_floor_types(doc), "No floor type loaded")
        self._fill("cmb_ceiling_type", cr.get_ceiling_types(doc), "No ceiling type loaded")
        cats = cr.get_ds_categories()
        for name, prefer in (("cmb_wall_part_category", "walls"),
                             ("cmb_part_category", "floors"),
                             ("cmb_ceiling_part_category", "ceilings"),
                             ("cmb_beam_part_category", "structural framing")):
            self._fill(name, cats, "No category", prefer=(prefer,), exact=True)
        self._fill("cb_beam_types",
                   cr.get_symbols(doc, DB.BuiltInCategory.OST_StructuralFraming),
                   "No structural framing family loaded")
        self._fill("cmb_grid_type", cr.get_grid_types(doc), "No grid type loaded")
        self._fill("cmb_line_style", cr.get_line_styles(doc), "No line style found",
                   prefer=("thin",))
        for c in geo.MEP_CATEGORIES:
            self._mep_types[c["key"]] = cr.get_mep_types(doc, c["key"])
            self._mep_systems[c["key"]] = cr.get_mep_systems(doc, c["key"])
        self._column_structural = None
        self._fill_column_types()
        self._fill_mep(self._mep_key, restore=True)

    def _fill_column_types(self):
        structural = self._checked(self.rb_column_structural)
        if structural == self._column_structural:
            return
        self._column_structural = structural
        bic = (DB.BuiltInCategory.OST_StructuralColumns if structural
               else DB.BuiltInCategory.OST_Columns)
        syms = cr.get_symbols(self._doc, bic)
        self._fill("cmb_column_type", syms,
                   "No {} column family loaded".format("structural" if structural else "architectural"),
                   prefer=("rectangular", "concrete", "square"))
        round_default = 0
        for i, s in enumerate(syms):
            low = s["name"].lower()
            if "round" in low or "circular" in low or "pipe" in low:
                round_default = i + 1
                break
        self._fill("cmb_column_round_type", [ROUND_SKIP] + syms, "No column family loaded",
                   default=round_default)

    def _fill_level_dependent(self):
        level = self._pick("cmb_levels")
        views = cr.get_plan_views(self._doc, level["id"]) if level else []
        active = 0
        try:
            av_id = eid_value(self._doc.ActiveView.Id)
            for i, v in enumerate(views):
                if eid_value(v["id"]) == av_id:
                    active = i
        except Exception:
            pass
        self._fill("cmb_room_view", views, "No floor plan of this level", default=active)
        above = []
        if level is not None:
            above = [lv for lv in self._levels if lv["elevation"] > level["elevation"] + 1e-6]
        self._fill("cmb_column_top_level", [UNCONNECTED] + above, "No level",
                   default=1 if above else 0)
        self._update_enabling()

    def _scan_current_cad(self):
        cad = self._pick("cmb_cad_files")
        layers = {}
        if cad is not None:
            self._set_status(u"Reading layers of {}…".format(cad["name"]))
            self._do_events()
            try:
                layers = cr.scan_cad(self._doc, cad["element"])
            except Exception as ex:
                self._set_status(u"Could not read {}: {}".format(cad["name"], ex), "danger")
                layers = {}
        self._rows = [LayerRow(layers[n]) for n in sorted(layers, key=lambda s: s.lower())]
        self._restore_selection()
        self._update_counts()
        self._apply_filter()
        if cad is None:
            self._set_status("No CAD file in this model — import or link a DWG first.", "warning")
        elif cad is not None and self._rows:
            self._set_status(u"Read {} layers from {}.".format(len(self._rows), cad["name"]))

    # ------------------------------------------------------------------
    # Combo helpers
    # ------------------------------------------------------------------

    def _fill(self, name, items, empty_text, label="name", default=0, prefer=None, exact=False):
        combo = getattr(self, name)
        combo.Items.Clear()
        self._lists[name] = list(items)
        if not items:
            hint = ComboBoxItem()
            hint.Content = empty_text
            hint.IsEnabled = False
            combo.Items.Add(hint)
            combo.SelectedIndex = 0
            return
        for it in items:
            combo.Items.Add(it[label])
        if prefer:
            for i, it in enumerate(items):
                low = it[label].lower()
                if any((low == p) if exact else (p in low) for p in prefer):
                    default = i
                    break
        combo.SelectedIndex = default if 0 <= default < len(items) else 0

    def _pick(self, name):
        items = self._lists.get(name) or []
        idx = getattr(self, name).SelectedIndex
        return items[idx] if 0 <= idx < len(items) else None

    @staticmethod
    def _index_of(items, item_id):
        for i, it in enumerate(items):
            if item_id is not None and it.get("id") == item_id:
                return i
        return 0

    def _current_cad_id(self):
        cad = self._pick("cmb_cad_files") if "cmb_cad_files" in self._lists else None
        return cad["id"] if cad else None

    def _current_level_id(self):
        lv = self._pick("cmb_levels") if "cmb_levels" in self._lists else None
        return eid_value(lv["id"]) if lv else None

    def _default_level_index(self, keep_id):
        wanted = keep_id
        if wanted is None:
            try:
                gen = self._doc.ActiveView.GenLevel
                wanted = eid_value(gen.Id) if gen is not None else None
            except Exception:
                wanted = None
        for i, lv in enumerate(self._levels):
            if wanted is not None and eid_value(lv["id"]) == wanted:
                return i
        return 0

    @staticmethod
    def _checked(ctrl):
        try:
            return bool(ctrl.IsChecked)
        except Exception:
            return False

    # ------------------------------------------------------------------
    # Status line
    # ------------------------------------------------------------------

    def _set_status(self, msg, state="ok"):
        try:
            self.txt_status.Text = msg
        except Exception:
            pass
        key = {"ok": "T3.Success.Accent", "warning": "T3.Warning.Accent",
               "danger": "T3.Danger.Accent"}.get(state, "T3.Success.Accent")
        try:
            brush = self.TryFindResource(key)
            if brush is not None:
                self.dot_status.Fill = brush
        except Exception:
            pass

    def _update_status(self, text):
        """Progress messages from T3WPFWindow.step_progress land here."""
        self._set_status(text)

    # ------------------------------------------------------------------
    # Modes
    # ------------------------------------------------------------------

    def mode_tile_clicked(self, sender, e):
        self._apply_mode(str(sender.Tag))

    def _apply_mode(self, key):
        if key not in self._panels:
            return
        self._store_selection()
        self._mode = key
        for k, tile in self._tiles.items():
            tile.IsChecked = (k == key)
        for k, panel in self._panels.items():
            panel.Visibility = Visibility.Visible if k == key else Visibility.Collapsed
        m = geo.mode(key)
        self.txt_mode_title.Text = m["title"]
        self.txt_mode_desc.Text = m["desc"]
        try:
            self.grid_layers.Columns[2].Header = geo.COUNT_HEADERS[m["count_kind"]]
        except Exception:
            pass
        self._restore_selection()
        self._update_counts()
        self._apply_filter()
        self._update_run_label()
        self._update_enabling()

    def _update_run_label(self):
        verb = geo.mode(self._mode)["verb"]
        if self._mode == "mep":
            verb = geo.mep_category(self._mep_key)["verb"]
        self.btn_run.Content = verb

    def create_as_changed(self, sender, e):
        if self._busy:
            return
        if self._mode == "column":
            self._fill_column_types()
        self._update_enabling()

    def column_top_level_changed(self, sender, e):
        self._update_enabling()

    def _update_enabling(self):
        """Inside a mode, fields that do not apply are disabled — never hidden."""
        c = self._checked
        wall_part = c(self.rb_wall_part_mode)
        for ctrl in (self.cmb_wall_type, self.chk_wall_match_thickness, self.chk_structural):
            ctrl.IsEnabled = not wall_part
        self.cmb_wall_part_category.IsEnabled = wall_part

        floor_part = c(self.rb_part_mode)
        self.cmb_floor_type.IsEnabled = not floor_part
        self.chk_floor_structural.IsEnabled = not floor_part
        self.txt_part_thickness.IsEnabled = floor_part
        self.cmb_part_category.IsEnabled = floor_part

        ceil_part = c(self.rb_ceiling_part_mode)
        self.cmb_ceiling_type.IsEnabled = not ceil_part
        self.txt_ceiling_part_thickness.IsEnabled = ceil_part
        self.cmb_ceiling_part_category.IsEnabled = ceil_part

        lines_only = c(self.rb_room_lines)
        self.cmb_room_view.IsEnabled = not c(self.rb_room_only)
        self.txt_room_name.IsEnabled = not lines_only
        self.chk_room_skip_existing.IsEnabled = not lines_only

        top = self._pick("cmb_column_top_level")
        self.txt_column_height.IsEnabled = top is None or top.get("element") is None

        beam_part = c(self.rb_beam_part_mode)
        self.cb_beam_types.IsEnabled = not beam_part
        self.chk_beam_match_size.IsEnabled = not beam_part
        self.cmb_beam_part_category.IsEnabled = beam_part

        auto = c(self.rb_grid_autoname)
        self.txt_grid_start_number.IsEnabled = auto
        self.txt_grid_start_letter.IsEnabled = auto

        self.txt_lines_offset.IsEnabled = not c(self.rb_lines_detail)

        cat = geo.mep_category(self._mep_key)
        double = cat["double"] and self.cmb_mep_line_mode.SelectedIndex == 1
        self.cmb_mep_system.IsEnabled = cat["system"]
        self.cmb_mep_line_mode.IsEnabled = cat["double"]
        self.txt_mep_width.IsEnabled = not double
        self.txt_mep_height.IsEnabled = cat["height"]

    # ── MEP category (ducts / pipes / cable trays / conduits) ──

    def mep_category_changed(self, sender, e):
        if self._busy:
            return
        key = str(sender.Tag)
        if key != self._mep_key:
            self._save_mep(self._mep_key)
            self._fill_mep(key, restore=True)
            self._update_run_label()
            self._update_counts()
            self._update_tally()

    def mep_line_mode_changed(self, sender, e):
        if not self._busy:
            self._update_enabling()

    def _save_mep(self, key):
        self._mep_values[key] = dict(
            type=self.cmb_mep_type.SelectedIndex, system=self.cmb_mep_system.SelectedIndex,
            line=self.cmb_mep_line_mode.SelectedIndex, width=self.txt_mep_width.Text,
            height=self.txt_mep_height.Text, offset=self.txt_mep_offset.Text)

    def _fill_mep(self, key, restore):
        self._mep_key = key
        cat = geo.mep_category(key)
        busy, self._busy = self._busy, True
        try:
            self.lbl_mep_type.Text = cat["type_label"].upper()
            self.lbl_mep_width.Text = self._unit.label(cat["width_label"])
            self._fill("cmb_mep_type", self._mep_types.get(key, []),
                       "No {} loaded — load an MEP template".format(cat["type_label"]))
            if cat["system"]:
                self._fill("cmb_mep_system", self._mep_systems.get(key, []),
                           "No system type loaded")
            else:
                self._fill("cmb_mep_system", [], "Not used for {}".format(cat["noun"]))
            saved = self._mep_values.get(key) if restore else None
            if saved:
                for name, idx in (("cmb_mep_type", saved["type"]), ("cmb_mep_system", saved["system"])):
                    combo = getattr(self, name)
                    if 0 <= idx < combo.Items.Count:
                        combo.SelectedIndex = idx
                self.cmb_mep_line_mode.SelectedIndex = saved["line"] if cat["double"] else 0
                self.txt_mep_width.Text = saved["width"]
                self.txt_mep_height.Text = saved["height"]
                self.txt_mep_offset.Text = saved["offset"]
            else:
                self.cmb_mep_line_mode.SelectedIndex = 0
                unit = self._unit
                self.txt_mep_width.Text = unit.default_text(cat["width"])
                self.txt_mep_height.Text = unit.default_text(cat["height_mm"]) if cat["height_mm"] else ""
                self.txt_mep_offset.Text = unit.default_text(cat["offset"])
        finally:
            self._busy = busy
        self._update_enabling()

    # ------------------------------------------------------------------
    # Source / level
    # ------------------------------------------------------------------

    def cad_file_changed(self, sender, e):
        if not self._busy:
            self._scan_current_cad()

    def level_changed(self, sender, e):
        if not self._busy:
            self._fill_level_dependent()

    def refresh_clicked(self, sender, e):
        self._rescan()

    def _rescan(self):
        self._store_selection()
        self._busy = True
        try:
            self._load_all()
        finally:
            self._busy = False
        self._apply_mode(self._mode)

    def _on_key_down(self, sender, e):
        if e.Key == Key.F5:
            self._rescan()
            e.Handled = True

    # ------------------------------------------------------------------
    # Layer list (shared by every mode)
    # ------------------------------------------------------------------

    def _store_selection(self):
        if self._rows:
            self._selected[self._mode] = set(r.name for r in self._rows if r.is_selected)

    def _restore_selection(self):
        chosen = self._selected.get(self._mode, set())
        for r in self._rows:
            r.is_selected = r.name in chosen

    def _count_kind(self):
        return geo.mode(self._mode)["count_kind"]

    def _update_counts(self):
        kind = self._count_kind()
        for r in self._rows:
            try:
                r.count = str(r.geom.count(kind))
            except Exception:
                r.count = "0"

    def _visible_rows(self):
        try:
            text = (self.txt_layer_search.Text or "").strip().lower()
        except Exception:
            text = ""
        if not text:
            return list(self._rows)
        return [r for r in self._rows if text in r.name.lower()]

    def _apply_filter(self):
        shown = self._visible_rows()
        self.set_items_source(self.grid_layers, shown)
        self.txt_layers_empty.Visibility = Visibility.Collapsed if shown else Visibility.Visible
        self._sync_header()
        self._update_tally()

    def _refresh_rows(self):
        try:
            self.grid_layers.Items.Refresh()
        except Exception:
            pass
        self._sync_header()
        self._update_tally()

    def _sync_header(self):
        try:
            self.sync_header_checkbox(self.chk_all_grid_layers, self.grid_layers, "is_selected")
        except Exception:
            pass

    def _update_tally(self):
        kind = self._count_kind()
        unit = geo.COUNT_HEADERS[kind].lower()
        selected = [r for r in self._rows if r.is_selected]
        total = 0
        for r in selected:
            try:
                total += int(r.count)
            except ValueError:
                pass
        text = u"{} of {} layers selected · {} {}".format(len(selected), len(self._rows), total, unit)
        shown = len(self._visible_rows())
        if shown != len(self._rows):
            text += u" · showing {}".format(shown)
        self.txt_layer_tally.Text = text

    def layer_search_changed(self, sender, e):
        self._apply_filter()

    def select_all_grid_layers_clicked(self, sender, e):
        self.toggle_all_rows(self.grid_layers, "is_selected", sender.IsChecked)
        self._store_selection()
        self._refresh_rows()

    def _on_row_checkbox_click(self, sender, e):
        self._store_selection()
        self._sync_header()
        self._update_tally()

    # ------------------------------------------------------------------
    # AI Select — semantic match of layer names (T3 AI Mode)
    # ------------------------------------------------------------------

    def ai_select_clicked(self, sender, e):
        if not self.ai_require():
            return
        names = [r.name for r in self._rows]
        if not names:
            show_info("There are no CAD layers to choose from.", title="AI Select",
                      details="Pick a CAD file under CAD SOURCE first.", owner=self)
            return
        mode_key = self._mode
        label = geo.mode(mode_key)["ai_label"]
        if mode_key == "mep":
            label = geo.mep_category(self._mep_key)["ai_label"]
        system_prompt = (
            "You are an expert BIM manager and CAD/Revit specialist. You read AutoCAD/DWG "
            "layer names and decide which layers hold a given building element category. "
            "Return JSON: {\"matched_layers\": [\"LAYER1\"], \"confidence\": 0.0-1.0, "
            "\"reasoning\": \"<one short English sentence>\"}.")
        prompt = ("Element category: {}\nCAD layer names in the drawing:\n{}\n\n"
                  "List every layer name that holds {}. Return only names from the list."
                  ).format(label, "\n".join("- " + n for n in names[:200]), label)
        self.ai_busy(self.btn_ai_select, True)
        self._set_status(u"AI is reading {} layer names for {}…".format(len(names), label))

        def worker():
            return self.ai_bridge.ask_json(prompt, system_prompt=system_prompt)

        def on_success(result):
            try:
                matched = set()
                if isinstance(result, dict):
                    matched = set(str(m).strip().lower() for m in result.get("matched_layers") or [])
                hits = [n for n in names if n.strip().lower() in matched]
                if self._mode == mode_key:
                    for r in self._rows:
                        if r.name in hits:
                            r.is_selected = True
                    self._store_selection()
                    self._refresh_rows()
                else:
                    self._selected[mode_key].update(hits)
                reason = result.get("reasoning", "") if isinstance(result, dict) else ""
                if hits:
                    self._set_status(u"AI ticked {} layer(s) for {}. {}".format(len(hits), label, reason))
                else:
                    self._set_status(u"AI found no layer that looks like {}. Tick the layers by hand."
                                     .format(label), "warning")
            finally:
                self.ai_busy(self.btn_ai_select, False)

        def on_error(err):
            self.ai_busy(self.btn_ai_select, False)
            self._set_status(u"AI Select failed: {}. Tick the layers by hand.".format(err), "danger")

        self.run_ai_async(worker, on_success=on_success, on_error=on_error)

    # ------------------------------------------------------------------
    # Run — plan, confirm (P5), create, report
    # ------------------------------------------------------------------

    def run_clicked(self, sender, e):
        self._store_selection()
        try:
            ctx = self._context()
            planners = {"wall": self._plan_wall, "floor": self._plan_floor,
                        "ceiling": self._plan_ceiling, "room": self._plan_room,
                        "column": self._plan_column, "beam": self._plan_beam,
                        "grid": self._plan_grid, "lines": self._plan_lines,
                        "mep": self._plan_mep}
            plan = planners[self._mode](ctx)
        except (_InputError, cp.PlanError) as ex:
            show_warning(str(ex), title=TITLE, owner=self)
            self._set_status(str(ex).splitlines()[0], "warning")
            return
        if plan.count <= 0:
            show_warning(u"Nothing to create: no {} found in the ticked layers.".format(plan.noun),
                         title=TITLE, details=plan.empty_msg, owner=self)
            self._set_status(u"No {} found in the ticked layers.".format(plan.noun), "warning")
            return
        if not confirm(plan.question, title=TITLE, ok_text=plan.ok_text,
                       details=plan.details, owner=self):
            self._set_status("Cancelled — nothing was created.", "warning")
            return
        self.begin_progress(max(plan.count, 1), disable=[self.btn_run, self.btn_refresh])
        try:
            result = plan.execute(self._progress)
        except Exception as ex:
            self.end_progress()
            self._set_status(u"Failed — nothing was created: {}".format(ex), "danger")
            show_error(u"Creating {} failed and the change was rolled back.".format(plan.noun),
                       title=TITLE, owner=self,
                       details=u"{}\n\nCheck the chosen type and level, then run again. If it "
                               u"keeps failing, read the latest Revit journal for the stack trace."
                               .format(ex))
            return
        self.end_progress()
        self._report(result)

    def _progress(self, i, total):
        step = max(1, total // 100)
        if i % step == 0 or i == total - 1:
            self.step_progress(i, u"Creating {} of {}…".format(i + 1, total))

    def _report(self, result):
        lines = [u"Created: {}".format(_counted(result.created, result.noun))]
        if result.failed:
            lines.append(u"Failed: {}".format(result.failed))
        if result.skipped:
            lines.append(u"Skipped: {}".format(result.skipped))
        if result.warnings:
            lines.append(u"Revit warnings dismissed: {}".format(result.warnings))
        lines.extend(result.notes)
        if result.errors:
            lines.append(u"First error: {}".format(result.errors[0]))
        details = u"\n".join(lines)
        level = self._pick("cmb_levels")
        where = u" on {}".format(level["name"]) if level and self._mode not in ("grid", "lines") else u""
        if result.created <= 0:
            self._set_status(u"No {} were created{} — nothing changed.".format(result.noun, where), "danger")
            show_warning(u"No {} were created; the model is unchanged.".format(result.noun),
                         title=TITLE, owner=self,
                         details=details + u"\n\nCheck the chosen type and the ticked layers, then run again.")
        elif result.failed:
            self._set_status(u"Created {}{} · {} failed.".format(
                _counted(result.created, result.noun), where, result.failed), "warning")
            show_warning(u"Created {}{}; {} could not be created.".format(
                _counted(result.created, result.noun), where, result.failed),
                title=TITLE, details=details, owner=self)
        else:
            self._set_status(u"Created {}{}. Ctrl+Z undoes the whole run.".format(
                _counted(result.created, result.noun), where))
            show_info(u"Created {}{}.".format(_counted(result.created, result.noun), where),
                      title=TITLE, details=details, owner=self)

    # ── shared inputs ──

    def _context(self):
        cad = self._pick("cmb_cad_files")
        if cad is None:
            raise _InputError(
                "No CAD file is selected.\nImport or link a DWG into this model, then pick it "
                "under CAD SOURCE.")
        layers = [r.geom for r in self._rows if r.is_selected]
        title = geo.mode(self._mode)["title"]
        if not layers:
            raise _InputError(
                u"No CAD layer is ticked for {}.\nTick the layers that hold the {} in the list "
                u"on the right, or use AI Select.".format(title, title.lower()))
        level = self._pick("cmb_levels")
        if level is None:
            raise _InputError("The model has no level.\nCreate a level, then run again.")
        return dict(cad=cad, layers=layers, level=level["element"], level_name=level["name"])

    def _length(self, box, label, positive=False, minimum=None, maximum=None, unit=None):
        """A length typed in the options card -> feet.

        A bare number is in the project unit (or `unit`); "1200 mm", "1.2 m"
        or 3'-6" work in any project. `minimum` / `maximum` are in feet.
        """
        unit = unit or self._unit
        text = box.Text or ""
        try:
            feet = unit.parse(text)
        except ValueError as ex:
            raise _InputError(u"{}: {}\nFix it in the options card, then run again."
                              .format(label, ex))
        # _units.parse drops the sign when the feet part is zero (-0' - 2"
        # comes back as +2"); the sign the user typed wins.
        if text.strip().startswith("-") and feet > 0:
            feet = -feet
        rule = None
        if positive and feet <= 0:
            rule = u"more than zero"
        elif minimum is not None and feet < minimum - 1e-9:
            rule = u"at least {}".format(self._len(minimum, unit))
        elif maximum is not None and feet > maximum + 1e-9:
            rule = u"at most {}".format(self._len(maximum, unit))
        if rule:
            raise _InputError(u"{} is {}, but it must be {}.\nChange it in the options card, "
                              u"then run again.".format(label, self._len(feet, unit), rule))
        return feet

    def _num(self, box, label, default, minimum=None, maximum=None):
        """A plain number that is not a length (the first grid number)."""
        value, ok = geo.parse_number(box.Text, default, minimum, maximum)
        if not ok:
            rng = ""
            if minimum is not None and maximum is not None:
                rng = " between {:g} and {:g}".format(minimum, maximum)
            elif minimum is not None:
                rng = " of at least {:g}".format(minimum)
            raise _InputError(u"{} is \"{}\", which is not a usable number.\nEnter a number{} "
                              u"in the options card, then run again.".format(label, box.Text, rng))
        return value

    def _need(self, name, what):
        item = self._pick(name)
        if item is None or item.get("element") is None:
            raise _InputError(u"No {} is available.\nLoad one into the model (or pick one in the "
                              u"options card), then run again.".format(what))
        return item

    def _category(self, combo_name):
        cat = self._pick(combo_name)
        if cat is None:
            raise _InputError("Pick a part category in the options card, then run again.")
        return cat

    # ── Plans: the window reads its controls, Snippets/_cad_plan plans ──

    def _plan_wall(self, ctx):
        height = self._length(self.txt_wall_height, "Height", positive=True)
        offset = self._length(self.txt_wall_offset, "Base offset")
        default_thk = self._length(self.txt_wall_thickness, "Unpaired thickness", positive=True,
                                   maximum=geo.mm(3000))
        part = self._checked(self.rb_wall_part_mode)
        return cp.plan_wall(
            self._doc, ctx["layers"], ctx["level"], ctx["level_name"], height, offset,
            default_thk, self._checked(self.chk_include_unpaired),
            self._checked(self.chk_merge_collinear),
            part_category=self._category("cmb_wall_part_category") if part else None,
            wall_type=None if part else self._need("cmb_wall_type", "basic wall type"),
            match=self._checked(self.chk_wall_match_thickness),
            structural=self._checked(self.chk_structural), fmt=self._len)

    def _plan_floor(self, ctx):
        offset = self._length(self.txt_floor_offset, "Height offset")
        holes = self._checked(self.chk_floor_holes)
        if self._checked(self.rb_part_mode):
            thk = self._length(self.txt_part_thickness, "Part thickness", positive=True)
            return cp.plan_floor(self._doc, ctx["layers"], ctx["level"], ctx["level_name"],
                                 offset, holes, part_category=self._category("cmb_part_category"),
                                 part_thickness=thk, fmt=self._len)
        return cp.plan_floor(self._doc, ctx["layers"], ctx["level"], ctx["level_name"], offset,
                             holes, floor_type=self._need("cmb_floor_type", "floor type"),
                             structural=self._checked(self.chk_floor_structural), fmt=self._len)

    def _plan_ceiling(self, ctx):
        offset = self._length(self.txt_ceiling_offset, "Height offset")
        holes = self._checked(self.chk_ceiling_holes)
        if self._checked(self.rb_ceiling_part_mode):
            thk = self._length(self.txt_ceiling_part_thickness, "Part thickness", positive=True)
            return cp.plan_ceiling(self._doc, ctx["layers"], ctx["level"], ctx["level_name"],
                                   offset, holes,
                                   part_category=self._category("cmb_ceiling_part_category"),
                                   part_thickness=thk, fmt=self._len)
        return cp.plan_ceiling(self._doc, ctx["layers"], ctx["level"], ctx["level_name"], offset,
                               holes, ceiling_type=self._need("cmb_ceiling_type", "ceiling type"),
                               fmt=self._len)

    def _plan_room(self, ctx):
        make_lines = not self._checked(self.rb_room_only)
        return cp.plan_room(self._doc, ctx["layers"], ctx["level"], ctx["level_name"],
                            make_lines, not self._checked(self.rb_room_lines),
                            self._pick("cmb_room_view") if make_lines else None,
                            (self.txt_room_name.Text or "").strip(),
                            self._checked(self.chk_room_skip_existing))

    def _plan_column(self, ctx):
        size_unit = self._size_unit
        step = self._length(self.txt_column_rounding, "Size rounding", positive=True,
                            maximum=geo.mm(500), unit=size_unit)
        height = self._length(self.txt_column_height, "Height", positive=True)
        return cp.plan_column(self._doc, ctx["layers"], ctx["level"], ctx["level_name"], step,
                              height, self._pick("cmb_column_type"),
                              self._pick("cmb_column_round_type"),
                              self._pick("cmb_column_top_level"),
                              self._checked(self.rb_column_structural),
                              self._checked(self.chk_column_match_size),
                              self._checked(self.chk_column_rotate),
                              size_unit=size_unit, fmt=self._len)

    def _plan_beam(self, ctx):
        offset = self._length(self.txt_beam_offset, "Top offset")
        w_min = self._length(self.txt_beam_min_width, "Min width", positive=True)
        w_max = self._length(self.txt_beam_max_width, "Max width", minimum=w_min)
        part = self._checked(self.rb_beam_part_mode)
        return cp.plan_beam(
            self._doc, ctx["layers"], ctx["level"], ctx["level_name"], offset, w_min, w_max,
            self._checked(self.chk_beam_merge),
            part_category=self._category("cmb_beam_part_category") if part else None,
            beam_type=None if part else self._need("cb_beam_types", "structural framing type"),
            match=self._checked(self.chk_beam_match_size), fmt=self._len)

    def _plan_grid(self, ctx):
        extend = self._length(self.txt_grid_extend, "Extend ends", minimum=0.0,
                              maximum=geo.mm(100000))
        min_len = self._length(self.txt_grid_min_length, "Min length", minimum=0.0)
        auto = self._checked(self.rb_grid_autoname)
        start_number = int(self._num(self.txt_grid_start_number, "First number", 1, 0, 100000)) if auto else 1
        start_letter = (self.txt_grid_start_letter.Text or "A").strip() or "A"
        return cp.plan_grid(self._doc, ctx["layers"], extend, min_len, auto, start_number,
                            start_letter, self._checked(self.chk_grid_skip_existing),
                            self._pick("cmb_grid_type"), fmt=self._len)

    def _plan_lines(self, ctx):
        doc, level = self._doc, ctx["level"]
        detail = self._checked(self.rb_lines_detail)
        view, z = None, level.Elevation
        if detail:
            view = doc.ActiveView
            allowed = (DB.ViewType.FloorPlan, DB.ViewType.CeilingPlan,
                       DB.ViewType.EngineeringPlan, DB.ViewType.AreaPlan,
                       DB.ViewType.DraftingView)
            if view is None or view.ViewType not in allowed:
                raise _InputError(u"Detail lines need a plan or drafting view, but the active view "
                                  u"\"{}\" is not one.\nClose this window, open a floor plan or "
                                  u"drafting view, and run again — or choose Model lines."
                                  .format(view.Name if view is not None else "?"))
            z = 0.0
            try:
                if view.GenLevel is not None:
                    z = view.GenLevel.Elevation
            except Exception:
                pass
        else:
            z = level.Elevation + self._length(self.txt_lines_offset, "Offset from level")
        return cp.plan_lines(doc, ctx["layers"], ctx["level_name"], view, z,
                             self._checked(self.chk_lines_merge), self._pick("cmb_line_style"))

    def _plan_mep(self, ctx):
        key = self._mep_key
        cat = geo.mep_category(key)
        tp = self._need("cmb_mep_type", cat["type_label"])
        system_id = None
        if cat["system"]:
            system_id = self._need("cmb_mep_system", "{} system type".format(cat["label"].lower()))["id"]
        double = cat["double"] and self.cmb_mep_line_mode.SelectedIndex == 1
        width = None                    # double lines: the width comes from the pair spacing
        if not double:
            width = self._length(self.txt_mep_width, cat["width_label"].capitalize(), positive=True)
        height = None
        if cat["height"]:
            height = self._length(self.txt_mep_height, "Height", positive=True)
        offset = self._length(self.txt_mep_offset, "Offset from level")
        return cp.plan_mep(self._doc, ctx["layers"], ctx["level"], ctx["level_name"], key, tp,
                           system_id, double, width, height, offset,
                           self._checked(self.chk_mep_merge), self._checked(self.chk_mep_elbows),
                           fmt=self._len)


# ===========================================================================
# PUBLIC ENTRY POINT
# ===========================================================================

def show_cad_to_elements(script_dir=None, revit_app=None):
    """Called by the pushbutton script.py (arguments kept for compatibility)."""
    try:
        CADToElementsWindow().ShowDialog()
    except Exception as ex:
        import traceback
        show_error(u"CAD to Elements could not open.", title=TITLE,
                   details=u"{}\n\n{}".format(ex, traceback.format_exc()))
