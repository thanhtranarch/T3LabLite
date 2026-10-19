# -*- coding: utf-8 -*-
"""
excel_table_revit — draw an Excel range in Revit and keep it linked.

Three targets, all built from the same `xlsx_table.Layout`:

  schedule  a ViewSchedule whose HEADER section holds the whole table: rows and
            columns inserted, cells merged, text, fonts, colours, borders,
            column widths and row heights set. The body is kept empty (one Mark
            field filtered to a value no element has, headers hidden), so the
            schedule is placeable on any number of sheets like a normal one.
  drafting  a Drafting View drawn with text notes, detail lines and solid
            filled regions — the closest copy of the Excel look.
  legend    the same drawing in a Legend view (placeable on many sheets).
            Revit has no API to create a legend from nothing, so the project
            needs one legend view to duplicate.

Revit itself can do none of this (no Excel import, no live link): that is why
it lives here. Every element this module creates is named "T3Lab Table ...",
and an update only deletes those, so annotation a user adds by hand survives.

Link data (file, sheet, range, options, fingerprint) is stored on the view
with Extensible Storage, so it travels with the model and every team member
sees the same link. If Extensible Storage fails under this engine, a JSON file
in %APPDATA%\\T3LabAI is used instead (per machine) and the status says so.

Units: layout is paper inches; Revit paper space is feet (1/12); drafting and
legend coordinates are paper x view scale.

NEEDS VERIFICATION in Revit (not testable here): TableCellStyle.TextSize unit
(auto-detected from the default header style), header row heights, line
weights of the generated line styles.

Author: Tran Tien Thanh
"""

import json
import os
import time

from Autodesk.Revit import DB

from Snippets._compat import disposing, eid_value, elem_name, net_list
from Snippets import xlsx_table as xt
from Snippets import gsheet_source as gs

PREFIX = "T3Lab Table"
KIND_SCHEDULE = "schedule"
KIND_DRAFTING = "drafting"
KIND_LEGEND = "legend"
KIND_LABEL = {KIND_SCHEDULE: "Schedule", KIND_DRAFTING: "Drafting view",
              KIND_LEGEND: "Legend"}

# Revit "text size" is close to the cap height of a font; Excel sizes are the
# em height in points. 0.7 keeps the text inside the Excel row heights.
TEXT_HEIGHT_FACTOR = 0.7
CELL_PAD_IN = 0.03                    # ~3 px inner margin, like Excel
LINE_WEIGHT = {1: 1, 2: 3, 3: 5}      # Excel border weight -> Revit pen
NO_ROWS_TOKEN = "T3LAB-EXCEL-TABLE-NO-ROWS"

_SCHEMA_GUID = "6b8e9a3c-2f41-4d7a-9e05-71c3b0d4a8f2"
_SCHEMA_NAME = "T3LabExcelTableLink"
_FIELD = "Json"


class TableError(Exception):
    """User-facing failure: what went wrong and what to do."""


def _try(obj, name, value):
    try:
        setattr(obj, name, value)
        return True
    except Exception:
        return False


def _color(rgb):
    return DB.Color(int(rgb[0]), int(rgb[1]), int(rgb[2]))


def _hex(rgb):
    return "#{:02X}{:02X}{:02X}".format(*rgb)


# ════════════════════════════════════════════════════════════════════
# LINK STORAGE
# ════════════════════════════════════════════════════════════════════

_ES_BROKEN = False


def _schema(create=False):
    """The link schema; built only when asked (inside a transaction)."""
    import System
    from Autodesk.Revit.DB.ExtensibleStorage import Schema, SchemaBuilder, AccessLevel
    guid = System.Guid(_SCHEMA_GUID)
    schema = Schema.Lookup(guid)
    if schema is not None or not create:
        return schema
    builder = SchemaBuilder(guid)
    builder.SetSchemaName(_SCHEMA_NAME)
    builder.SetReadAccessLevel(AccessLevel.Public)
    builder.SetWriteAccessLevel(AccessLevel.Public)
    builder.AddSimpleField(_FIELD, System.Type.GetType("System.String"))
    return builder.Finish()


def _es_get(view):
    import System
    schema = _schema()
    if schema is None:                 # nothing was ever linked in this session
        return None
    ent = view.GetEntity(schema)
    if ent is None or not ent.IsValid():
        return None
    field = schema.GetField(_FIELD)
    try:
        return ent.Get[System.String](field)
    except TypeError:
        return ent.Get[str](_FIELD)


def _es_set(view, text):
    import System
    from Autodesk.Revit.DB.ExtensibleStorage import Entity
    schema = _schema(create=True)
    ent = Entity(schema)
    field = schema.GetField(_FIELD)
    try:
        ent.Set[System.String](field, text)
    except TypeError:
        ent.Set[str](_FIELD, text)
    view.SetEntity(ent)


def _fallback_path():
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    folder = os.path.join(base, 'T3LabAI')
    if not os.path.isdir(folder):
        os.makedirs(folder)
    return os.path.join(folder, 'excel_table_links.json')


def _doc_key(doc):
    return doc.PathName or doc.Title


def _fallback_load():
    try:
        with open(_fallback_path(), encoding='utf-8') as fh:
            return json.load(fh)
    except (IOError, OSError, ValueError):
        return {}


def _fallback_save(data):
    with open(_fallback_path(), 'w', encoding='utf-8') as fh:
        json.dump(data, fh, indent=1)


def storage_note():
    """'' normally; a sentence when links are only stored on this machine."""
    if _ES_BROKEN:
        return ("Links are saved on this computer only (Extensible Storage "
                "is not available in this session).")
    return ""


# Errors that mean "Extensible Storage cannot be called from this engine"
# (pythonnet generic binding, missing assembly) — only those switch the
# session to the JSON fallback. Anything else is a real per-view error.
_BINDING_ERRORS = (TypeError, AttributeError, ImportError, NotImplementedError)


def get_link(view, fallback=None):
    """Link dict stored on `view`, or None. `fallback` = preloaded JSON data."""
    global _ES_BROKEN
    text = None
    if not _ES_BROKEN:
        try:
            text = _es_get(view)
        except _BINDING_ERRORS:
            _ES_BROKEN = True
        except Exception:
            text = None
    if text is None:
        data = fallback if fallback is not None else _fallback_load()
        info = data.get(_doc_key(view.Document), {}).get(view.UniqueId)
        return dict(info) if info else None
    try:
        return json.loads(text) if text else None
    except ValueError:
        return None


def set_link(view, info):
    """Store `info` on `view` (call inside a transaction)."""
    global _ES_BROKEN
    text = json.dumps(info)
    if not _ES_BROKEN:
        try:
            _es_set(view, text)
            return
        except _BINDING_ERRORS:
            _ES_BROKEN = True
    data = _fallback_load()
    data.setdefault(_doc_key(view.Document), {})[view.UniqueId] = info
    _fallback_save(data)


def clear_link(view):
    """Remove the link (call inside a transaction). The drawing stays."""
    if not _ES_BROKEN:
        try:
            schema = _schema()
            if schema is not None:
                view.DeleteEntity(schema)
        except Exception:
            pass
    data = _fallback_load()
    per_doc = data.get(_doc_key(view.Document), {})
    if per_doc.pop(view.UniqueId, None) is not None:
        _fallback_save(data)


def find_linked_views(doc):
    """[(view, info)] for every view in `doc` that carries a link."""
    out = []
    fallback = _fallback_load()
    for view in DB.FilteredElementCollector(doc).OfClass(DB.View).ToElements():
        try:
            if view.IsTemplate:
                continue
            if view.ViewType not in (DB.ViewType.Schedule, DB.ViewType.DraftingView,
                                     DB.ViewType.Legend):
                continue
        except Exception:
            continue
        info = get_link(view, fallback)
        if info:
            out.append((view, info))
    out.sort(key=lambda p: p[0].Name.lower())
    return out


def resolve_path(doc, info):
    """Absolute path of the linked workbook, trying the stored relative path
    when the absolute one is gone (project folder moved / other PC)."""
    path = info.get('file') or ''
    if path and os.path.isfile(path):
        return path
    rel = info.get('rel') or ''
    folder = os.path.dirname(doc.PathName or '')
    if rel and folder:
        cand = os.path.normpath(os.path.join(folder, rel))
        if os.path.isfile(cand):
            return cand
    return None


def make_info(doc, path, table, kind, options, gsheet=None):
    """Link record. `gsheet` = the Google Sheets link when the workbook is a
    downloaded Google Sheet; `path` is then only the local cache copy."""
    rel = ''
    folder = os.path.dirname(doc.PathName or '')
    if folder and not gsheet:
        try:
            rel = os.path.relpath(path, folder)
        except ValueError:            # other drive
            rel = ''
    return {
        'v': 1, 'file': path, 'rel': rel, 'sheet': table.sheet,
        'range': table.range_name, 'kind': kind,
        'formatting': bool(options.get('formatting', True)),
        'hidden': bool(options.get('hidden', False)),
        'scale': float(options.get('scale', 1.0)),
        'fp': table.fingerprint(),
        'updated': time.strftime('%Y-%m-%d %H:%M'),
        'gsheet': gsheet or '',
    }


def source_label(info):
    """Short text for the SOURCE column."""
    if info.get('gsheet'):
        return "Google Sheet"
    return os.path.basename(info.get('file') or '?')


def read_linked_table(doc, info):
    """(table, path) for a link, raising TableError with a readable reason.
    A Google Sheet is downloaded again, so the table reads the live sheet."""
    if info.get('gsheet'):
        try:
            path, _ = gs.download(info['gsheet'])
        except gs.GSheetError as ex:
            raise TableError(str(ex))
    else:
        path = resolve_path(doc, info)
    if not path:
        raise TableError("Excel file not found: {}".format(info.get('file')))
    try:
        table = xt.read_table(path, info.get('sheet'), info.get('range'),
                              info.get('hidden', False))
    except xt.XlsxError as ex:
        raise TableError(str(ex))
    return table, path


def check_status(doc, info):
    """(status text, severity, detail) for the grid."""
    google = bool(info.get('gsheet'))
    if not google and not resolve_path(doc, info):
        return "File missing", "Danger", "Excel file not found: {}".format(info.get('file'))
    try:
        table, path = read_linked_table(doc, info)
    except TableError as ex:
        if google and str(ex).startswith("Cannot reach"):
            return "Offline", "Warning", str(ex)
        return "Read error", "Danger", str(ex)
    if table.fingerprint() == info.get('fp'):
        return "Up to date", "Success", info.get('gsheet') or path
    return ("Sheet changed" if google else "Excel changed"), "Warning", info.get('gsheet') or path


# ════════════════════════════════════════════════════════════════════
# SHARED REVIT STYLE CACHE
# ════════════════════════════════════════════════════════════════════

class _Styles(object):
    """Finds or creates the line styles, text types and fill types a table needs."""

    def __init__(self, doc):
        self.doc = doc
        self._lines = {}
        self._texts = {}
        self._fills = {}
        self._invisible = None
        self._solid = None

    # ── line styles (sub-categories of Lines) ────────────────────────
    def line_style(self, weight, rgb):
        """GraphicsStyle for an Excel border (weight 1..3, colour)."""
        key = (weight, tuple(rgb))
        if key in self._lines:
            return self._lines[key]
        name = "{} - W{} {}".format(PREFIX, weight, _hex(rgb))
        cats = self.doc.Settings.Categories
        lines = cats.get_Item(DB.BuiltInCategory.OST_Lines)
        sub = None
        for c in lines.SubCategories:
            if c.Name == name:
                sub = c
                break
        if sub is None:
            sub = cats.NewSubcategory(lines, name)
            _try(sub, 'LineColor', _color(rgb))
            try:
                sub.SetLineWeight(LINE_WEIGHT.get(weight, 1), DB.GraphicsStyleType.Projection)
            except Exception:
                pass
        gs = sub.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
        self._lines[key] = gs
        return gs

    def invisible(self):
        if self._invisible is None:
            self._invisible = False
            cat = None
            try:
                cat = DB.Category.GetCategory(self.doc, DB.BuiltInCategory.OST_InvisibleLines)
            except Exception:
                pass
            if cat is None:
                try:
                    cat = self.doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_InvisibleLines)
                except Exception:
                    cat = None
            if cat is None:
                try:
                    lines = self.doc.Settings.Categories.get_Item(DB.BuiltInCategory.OST_Lines)
                    cat = next((c for c in lines.SubCategories
                                if c.Name == '<Invisible lines>'), None)
                except Exception:
                    cat = None
            if cat is not None:
                try:
                    self._invisible = cat.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
                except Exception:
                    self._invisible = False
        return self._invisible or None

    # ── text note types ─────────────────────────────────────────────
    def text_type(self, font, size_in):
        """TextNoteType id for a font dict at a paper text height (inches)."""
        size_in = max(1.0, round(size_in * 64)) / 64.0  # 1/64" steps, few types, never 0
        rgb = tuple(font.get('color') or (0, 0, 0))
        flags = ''.join(f for f, on in (('B', font.get('bold')), ('I', font.get('italic')),
                                        ('U', font.get('underline'))) if on)
        name = "{} - {} {:.2f}mm{} {}".format(
            PREFIX, font.get('name') or 'Arial', size_in * 25.4,
            (' ' + flags) if flags else '', _hex(rgb))
        if name in self._texts:
            return self._texts[name]
        found = None
        for t in DB.FilteredElementCollector(self.doc).OfClass(DB.TextNoteType):
            if elem_name(t) == name:
                found = t
                break
        if found is None:
            base = self.doc.GetElement(
                self.doc.GetDefaultElementTypeId(DB.ElementTypeGroup.TextNoteType))
            if base is None:
                base = DB.FilteredElementCollector(self.doc).OfClass(DB.TextNoteType).FirstElement()
            found = base.Duplicate(name)
            p = DB.BuiltInParameter
            for bip, val in ((p.TEXT_FONT, font.get('name') or 'Arial'),
                             (p.TEXT_SIZE, size_in / 12.0),
                             (p.TEXT_STYLE_BOLD, 1 if font.get('bold') else 0),
                             (p.TEXT_STYLE_ITALIC, 1 if font.get('italic') else 0),
                             (p.TEXT_STYLE_UNDERLINE, 1 if font.get('underline') else 0),
                             (p.LINE_COLOR, rgb[0] + rgb[1] * 256 + rgb[2] * 65536),
                             (p.TEXT_BACKGROUND, 1),            # transparent
                             (p.TEXT_BOX_VISIBILITY, 0),
                             (p.TEXT_WIDTH_SCALE, 1.0)):
                try:
                    prm = found.get_Parameter(bip)
                    if prm is not None and not prm.IsReadOnly:
                        prm.Set(val)
                except Exception:
                    pass
        self._texts[name] = found.Id
        return found.Id

    # ── solid filled region types ───────────────────────────────────
    def solid_pattern(self):
        if self._solid is None:
            self._solid = DB.ElementId.InvalidElementId
            for fp in DB.FilteredElementCollector(self.doc).OfClass(DB.FillPatternElement):
                try:
                    pat = fp.GetFillPattern()
                    if pat.IsSolidFill and pat.Target == DB.FillPatternTarget.Drafting:
                        self._solid = fp.Id
                        break
                except Exception:
                    pass
        return self._solid

    def fill_type(self, rgb):
        key = tuple(rgb)
        if key in self._fills:
            return self._fills[key]
        name = "{} Fill {}".format(PREFIX, _hex(rgb))
        found = None
        types = list(DB.FilteredElementCollector(self.doc).OfClass(DB.FilledRegionType))
        for t in types:
            if elem_name(t) == name:
                found = t
                break
        if found is None:
            if not types:
                raise TableError("The project has no Filled Region type to copy.")
            found = types[0].Duplicate(name)
            _try(found, 'ForegroundPatternId', self.solid_pattern())
            _try(found, 'ForegroundPatternColor', _color(rgb))
            _try(found, 'BackgroundPatternId', DB.ElementId.InvalidElementId)
            _try(found, 'IsMasking', False)
        self._fills[key] = found.Id
        return found.Id


# ════════════════════════════════════════════════════════════════════
# NAMES
# ════════════════════════════════════════════════════════════════════

def unique_view_name(doc, wanted, ignore_id=None):
    taken = set()
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
        if ignore_id is not None and eid_value(v.Id) == ignore_id:
            continue
        try:
            taken.add(v.Name.lower())
        except Exception:
            pass
    name = (wanted or "Excel Table").strip()
    for ch in '{}[]|;<>?`~:\\':
        name = name.replace(ch, '-')
    if name.lower() not in taken:
        return name
    i = 2
    while "{} ({})".format(name, i).lower() in taken:
        i += 1
    return "{} ({})".format(name, i)


# ════════════════════════════════════════════════════════════════════
# DRAFTING / LEGEND RENDERER
# ════════════════════════════════════════════════════════════════════

def _clear_drawing(doc, view):
    """Delete the T3Lab Table elements of a view; keep everything else."""
    ids = []
    for el in DB.FilteredElementCollector(doc, view.Id).WhereElementIsNotElementType():
        try:
            if isinstance(el, DB.TextNote) or isinstance(el, DB.FilledRegion):
                tname = elem_name(doc.GetElement(el.GetTypeId()))
                if tname.startswith(PREFIX):
                    ids.append(el.Id)
            elif isinstance(el, DB.CurveElement):
                gs = el.LineStyle
                if gs is not None and gs.Name.startswith(PREFIX):
                    ids.append(el.Id)
        except Exception:
            pass
    if ids:
        doc.Delete(net_list(DB.ElementId, ids))
    return len(ids)


def _draw(doc, view, layout, styles):
    """Draw a layout in a drafting / legend view. Returns element count."""
    k = view.Scale / 12.0                       # paper inch -> model feet

    def pt(x, y):
        return DB.XYZ(x * k, -y * k, 0.0)

    count = 0
    short = doc.Application.ShortCurveTolerance
    invisible = styles.invisible()

    # 1 · fills (drawn first, under lines and text)
    for (x, y, w, h, rgb) in layout.fills:
        if w * k < short or h * k < short:              # zero-size (hidden) cell
            continue
        loop = DB.CurveLoop()
        corners = [pt(x, y), pt(x + w, y), pt(x + w, y + h), pt(x, y + h)]
        for i in range(4):
            loop.Append(DB.Line.CreateBound(corners[i], corners[(i + 1) % 4]))
        region = DB.FilledRegion.Create(doc, styles.fill_type(rgb), view.Id,
                                        net_list(DB.CurveLoop, [loop]))
        if invisible is not None:
            try:
                region.SetLineStyleId(invisible.Id)
            except Exception:
                pass
        count += 1

    # 2 · borders
    for (p1, p2, weight, rgb) in layout.segments:
        a, b = pt(*p1), pt(*p2)
        if a.DistanceTo(b) < short:
            continue
        curve = doc.Create.NewDetailCurve(view, DB.Line.CreateBound(a, b))
        curve.LineStyle = styles.line_style(weight, rgb)
        count += 1

    # 3 · text — created, then moved into its cell once Revit knows its size
    halign = {'left': DB.HorizontalTextAlignment.Left,
              'center': DB.HorizontalTextAlignment.Center,
              'right': DB.HorizontalTextAlignment.Right}
    notes = []
    for t in layout.texts:
        size_in = t['font']['size'] / 72.0 * TEXT_HEIGHT_FACTOR
        type_id = styles.text_type(t['font'], size_in)
        opts = DB.TextNoteOptions(type_id)
        opts.HorizontalAlignment = halign.get(t['halign'], DB.HorizontalTextAlignment.Left)
        _try(opts, 'VerticalAlignment', DB.VerticalTextAlignment.Top)
        indent = t['indent'] * 0.1 if t['halign'] != 'center' else 0.0
        width_ft = max(0.0, t['w'] - 2 * CELL_PAD_IN - indent) / 12.0
        try:
            width_ft = max(width_ft, DB.TextElement.GetMinimumAllowedWidth(doc, type_id))
            width_ft = min(width_ft, DB.TextElement.GetMaximumAllowedWidth(doc, type_id))
        except Exception:
            pass
        x0 = t['x'] + CELL_PAD_IN + (indent if t['halign'] == 'left' else 0.0)
        note = DB.TextNote.Create(doc, view.Id, pt(x0, t['y'] + CELL_PAD_IN),
                                  width_ft, t['text'].replace('\r\n', '\r').replace('\n', '\r'),
                                  opts)
        notes.append((note, t, indent))
        count += 1

    if notes:
        doc.Regenerate()
        for note, t, indent in notes:
            bb = note.get_BoundingBox(view)
            if bb is None:
                continue
            left, right = bb.Min.X / k, bb.Max.X / k
            top, height = -bb.Max.Y / k, (bb.Max.Y - bb.Min.Y) / k
            # horizontal: the text box edges / centre on the cell
            if t['halign'] == 'right':
                want_x = (t['x'] + t['w'] - CELL_PAD_IN - indent) - right
            elif t['halign'] == 'center':
                want_x = (t['x'] + t['w'] / 2.0) - (left + right) / 2.0
            else:
                want_x = (t['x'] + CELL_PAD_IN + indent) - left
            # vertical: top / centre / bottom of the (possibly merged) cell
            if t['valign'] == 'top':
                want_top = t['y'] + CELL_PAD_IN
            elif t['valign'] == 'center':
                want_top = t['y'] + (t['h'] - height) / 2.0
            else:
                want_top = t['y'] + t['h'] - CELL_PAD_IN - height
            dx, dy = want_x * k, -(want_top - top) * k
            if abs(dx) > 1e-9 or abs(dy) > 1e-9:
                DB.ElementTransformUtils.MoveElement(doc, note.Id, DB.XYZ(dx, dy, 0.0))
    return count


def _new_drafting(doc, name):
    vft = None
    for t in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType):
        if t.ViewFamily == DB.ViewFamily.Drafting:
            vft = t
            break
    if vft is None:
        raise TableError("The project has no Drafting View type.")
    view = DB.ViewDrafting.Create(doc, vft.Id)
    view.Name = unique_view_name(doc, name)
    return view


def _new_legend(doc, name):
    source = None
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View):
        try:
            if v.ViewType == DB.ViewType.Legend and not v.IsTemplate \
                    and v.CanViewBeDuplicated(DB.ViewDuplicateOption.Duplicate):
                source = v
                break
        except Exception:
            pass
    if source is None:
        raise TableError("Revit cannot create a legend from scratch. Create one "
                         "legend view first (View > Legends > Legend), then try again.")
    view = doc.GetElement(source.Duplicate(DB.ViewDuplicateOption.Duplicate))
    view.Name = unique_view_name(doc, name)
    return view


# ════════════════════════════════════════════════════════════════════
# SCHEDULE RENDERER (header section)
# ════════════════════════════════════════════════════════════════════

def _mark_field(definition):
    want = eid_value(DB.ElementId(DB.BuiltInParameter.ALL_MODEL_MARK))
    for sf in definition.GetSchedulableFields():
        if eid_value(sf.ParameterId) == want:
            return sf
    fields = list(definition.GetSchedulableFields())
    return fields[0] if fields else None


def _new_header_schedule(doc, name, total_width_ft):
    vs = DB.ViewSchedule.CreateSchedule(doc, DB.ElementId(DB.BuiltInCategory.OST_GenericModel))
    vs.Name = unique_view_name(doc, name)
    d = vs.Definition
    sf = _mark_field(d)
    if sf is None:
        raise TableError("Could not set up the schedule body (no schedulable field).")
    field = d.AddField(sf)
    _try(d, 'ShowHeaders', False)
    _try(d, 'ShowGrandTotal', False)
    _try(d, 'IsItemized', True)
    _try(d, 'ShowGridLines', False)
    try:
        d.AddFilter(DB.ScheduleFilter(field.FieldId, DB.ScheduleFilterType.Equal, NO_ROWS_TOKEN))
    except Exception:
        pass
    _try(field, 'GridColumnWidth', total_width_ft)
    return vs


def _text_size_value(points, default_value):
    """TableCellStyle.TextSize for an Excel size. The API does not state the
    unit; the default header size tells: < 0.2 means feet, else points."""
    inches = points / 72.0 * TEXT_HEIGHT_FACTOR
    if default_value and 0 < default_value < 0.2:
        return inches / 12.0
    return inches * 72.0


def _fill_header(doc, vs, table, layout, formatting, scale, styles):
    hdr = vs.GetTableData().GetSectionData(DB.SectionType.Header)
    r0, c0 = hdr.FirstRowNumber, hdr.FirstColumnNumber

    # Never loop on Revit: a call that does not grow the section stops here.
    while hdr.NumberOfRows < table.rows:
        before = hdr.NumberOfRows
        hdr.InsertRow(hdr.FirstRowNumber + 1)
        if hdr.NumberOfRows <= before:
            raise TableError("Revit did not add a header row ({} of {}).".format(
                before, table.rows))
    while hdr.NumberOfColumns < table.cols:
        before = hdr.NumberOfColumns
        hdr.InsertColumn(hdr.LastColumnNumber)
        if hdr.NumberOfColumns <= before:
            raise TableError("Revit did not add a header column ({} of {}).".format(
                before, table.cols))
    try:
        vs.RefreshData()
    except Exception:
        pass

    widths = [w * scale / 12.0 for w in table.col_widths_in]
    for j, w in enumerate(widths):
        try:
            hdr.SetColumnWidth(c0 + j, w)
        except Exception:
            pass
    for i, h in enumerate(table.row_heights_in):
        try:
            hdr.SetRowHeight(r0 + i, h * scale / 12.0)
        except Exception:
            pass

    for (a, b, c, d) in table.merges:
        try:
            hdr.MergeCells(DB.TableMergedCell(r0 + a, c0 + b, r0 + c, c0 + d))
        except Exception:
            pass

    try:
        default_size = hdr.GetTableCellStyle(r0, c0).TextSize
    except Exception:
        default_size = None

    merged_inner = set()
    for (a, b, c, d) in table.merges:
        for r in range(a, c + 1):
            for col in range(b, d + 1):
                if (r, col) != (a, b):
                    merged_inner.add((r, col))

    h_map = {'left': DB.HorizontalAlignmentStyle.Left,
             'center': DB.HorizontalAlignmentStyle.Center,
             'right': DB.HorizontalAlignmentStyle.Right}
    v_map = {'top': DB.VerticalAlignmentStyle.Top,
             'center': DB.VerticalAlignmentStyle.Middle,
             'bottom': DB.VerticalAlignmentStyle.Bottom}
    invisible = styles.invisible()
    styled = 0
    for r in range(table.rows):
        for c in range(table.cols):
            if (r, c) in merged_inner:
                continue
            cell = table.cells[r][c]
            row, col = r0 + r, c0 + c
            try:
                hdr.SetCellType(row, col, DB.CellType.Text)
            except Exception:
                pass
            try:
                hdr.SetCellText(row, col, cell.text)
            except Exception:
                pass
            try:
                if not hdr.AllowOverrideCellStyle(row, col):
                    continue
                st = hdr.GetTableCellStyle(row, col)
                opts = st.GetCellStyleOverrideOptions()
            except Exception:
                continue
            font = cell.font if formatting else xt.DEFAULT_FONT
            pairs = [
                ('Font', 'FontName', font.get('name') or 'Arial'),
                ('FontSize', 'TextSize', _text_size_value(font.get('size', 11) * scale, default_size)),
                ('Bold', 'IsFontBold', bool(font.get('bold'))),
                ('Italics', 'IsFontItalic', bool(font.get('italic'))),
                ('Underline', 'IsFontUnderline', bool(font.get('underline'))),
                ('FontColor', 'TextColor', _color(font.get('color') or (0, 0, 0))),
                ('HorizontalAlignment', 'FontHorizontalAlignment', h_map.get(cell.halign)),
                ('VerticalAlignment', 'FontVerticalAlignment', v_map.get(cell.valign)),
            ]
            if formatting:
                pairs.append(('BackgroundColor', 'BackgroundColor',
                              _color(cell.fill or (255, 255, 255))))
                for side, opt, prop in (('top', 'BorderTopLineStyle', 'BorderTopLineStyle'),
                                        ('bottom', 'BorderBottomLineStyle', 'BorderBottomLineStyle'),
                                        ('left', 'BorderLeftLineStyle', 'BorderLeftLineStyle'),
                                        ('right', 'BorderRightLineStyle', 'BorderRightLineStyle')):
                    b = cell.borders.get(side)
                    gs = styles.line_style(b[0], b[1]) if b else invisible
                    if gs is not None:
                        pairs.append((opt, prop, gs.Id))
            for opt_name, prop, value in pairs:
                if value is None:
                    continue
                if _try(st, prop, value):
                    _try(opts, opt_name, True)
            try:
                st.SetCellStyleOverrideOptions(opts)
                hdr.SetCellStyle(row, col, st)
                styled += 1
            except Exception:
                pass
    return styled


def _sheet_placements(doc, schedule_id):
    out = []
    want = eid_value(schedule_id)
    for inst in DB.FilteredElementCollector(doc).OfClass(DB.ScheduleSheetInstance):
        try:
            if eid_value(inst.ScheduleId) != want or inst.IsTitleblockRevisionSchedule:
                continue
            rot = getattr(inst, 'Rotation', None)
            out.append((inst.OwnerViewId, inst.Point, rot))
        except Exception:
            pass
    return out


# ════════════════════════════════════════════════════════════════════
# PUBLIC: CREATE / UPDATE / UNLINK
# ════════════════════════════════════════════════════════════════════

def _commit(t):
    """Commit and fail loudly when Revit rolled the transaction back."""
    status = t.Commit()
    if status is not None and status != DB.TransactionStatus.Committed:
        raise TableError("Revit rolled the change back ({}). Check the warnings "
                         "Revit showed.".format(status))


def _render(doc, kind, name, table, options, styles, view=None):
    """Build the table; returns (view, element count). `view` = redraw target
    (drafting / legend only)."""
    scale = float(options.get('scale', 1.0))
    formatting = bool(options.get('formatting', True))
    layout = xt.build_layout(table, scale, formatting)
    if kind == KIND_SCHEDULE:
        vs = _new_header_schedule(doc, name, layout.width / 12.0)
        doc.Regenerate()
        return vs, _fill_header(doc, vs, table, layout, formatting, scale, styles)
    if view is None:
        view = _new_drafting(doc, name) if kind == KIND_DRAFTING else _new_legend(doc, name)
    return view, _draw(doc, view, layout, styles)


def create_table(doc, path, table, kind, name, options, link=True, gsheet=None):
    """One transaction: new view with the table (+ link). Returns the view."""
    with disposing(DB.Transaction(doc, "T3Lab: Excel Table - Create")) as t:
        t.Start()
        try:
            view, _ = _render(doc, kind, name, table, options, _Styles(doc))
            if link:
                set_link(view, make_info(doc, path, table, kind, options, gsheet))
            _commit(t)
        except TableError:
            t.RollBack()
            raise
        except Exception as ex:
            t.RollBack()
            raise TableError("Revit refused to build the table: {}".format(ex))
    return view


def update_table(doc, view, info, path_override=None):
    """Re-read the Excel range and redraw. Returns (view, warnings).

    Drafting / legend: the T3Lab elements are redrawn in place, so the view
    stays on its sheets. Schedule: a new schedule is built, put on the same
    sheets at the same points, the old one deleted and the name kept.
    """
    if path_override:
        info = dict(info)
        if gs.is_google(path_override):
            src = gs.parse(path_override)
            info['gsheet'] = src['url'] if src else path_override
            info['file'] = ''
        else:
            info['gsheet'] = ''
            info['file'] = path_override
        info['rel'] = ''
    table, path = read_linked_table(doc, info)
    kind = info.get('kind', KIND_DRAFTING)
    label = view.Name
    warnings = []
    if kind == KIND_SCHEDULE:
        try:
            active = doc.ActiveView
        except Exception:
            active = None
        if active is not None and eid_value(active.Id) == eid_value(view.Id):
            raise TableError("'{}' is the active view — Revit cannot rebuild it while it "
                             "is open. Switch to another view, then update.".format(label))
    with disposing(DB.Transaction(doc, "T3Lab: Excel Table - Update")) as t:
        t.Start()
        try:
            styles = _Styles(doc)
            if kind == KIND_SCHEDULE:
                old_name, old_id = view.Name, view.Id
                places = _sheet_placements(doc, old_id)
                new_view, _ = _render(doc, kind, old_name + " (updating)", table, info, styles)
                doc.Delete(old_id)
                new_view.Name = unique_view_name(doc, old_name)
                for sheet_id, point, rot in places:
                    try:
                        inst = DB.ScheduleSheetInstance.Create(doc, sheet_id, new_view.Id, point)
                        if rot is not None:
                            _try(inst, 'Rotation', rot)
                    except Exception as ex:
                        warnings.append("Could not put '{}' back on a sheet: {}".format(
                            old_name, ex))
                view = new_view
            else:
                _clear_drawing(doc, view)
                _render(doc, kind, view.Name, table, info, styles, view=view)
            set_link(view, make_info(doc, path, table, kind, info, info.get('gsheet')))
            _commit(t)
        except TableError:
            t.RollBack()
            raise
        except Exception as ex:
            t.RollBack()
            raise TableError("Update failed for '{}': {}".format(label, ex))
    return view, warnings


def unlink_tables(doc, views):
    """Remove the link from views; the tables stay as static copies."""
    with disposing(DB.Transaction(doc, "T3Lab: Excel Table - Unlink")) as t:
        t.Start()
        for v in views:
            clear_link(v)
        _commit(t)
    return len(views)
