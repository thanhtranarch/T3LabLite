# -*- coding: utf-8 -*-
"""SheetGen without the window: per room floor / ceiling plans, four interior
elevations, sheets and viewport layout.

Used by GUI/SheetGenDialog.py (options read from its controls) and the MCP tool
generate_room_views (core/tools/room_views.py). Revit is reached through api()
so tests run it on fakes; the MCP server imports it, so IronPython 2.7-safe.
Tests: dev/test_sheetgen_logic.py, dev/test_tools_views_domain.py.
"""
from __future__ import unicode_literals

from contextlib import contextmanager
import math

# Printable margin inside the title block (~20 mm), shared by the dialog's
# layout preview and the real placement so they can never drift apart.
SHEET_MARGIN_FT = 0.066
A1_FT = (2.759, 1.949)                  # 841 x 594 mm
DIRECTIONS = ('North', 'East', 'South', 'West')
ENLARGED_PREFIX = 'ENLARGED PLAN'

_API = {}


def api():
    """Autodesk.Revit.DB (tests put a fake in _API['DB'])."""
    if 'DB' not in _API:
        from Autodesk.Revit import DB
        _API['DB'] = DB
    return _API['DB']


class SheetGenPending(BaseException):
    """Abort the batch without further API writes while Revit resolves failures."""


class _NoLog(object):
    def _skip(self, *_args):
        pass

    debug = info = warning = error = _skip


@contextmanager
def transaction(doc, name, subtransaction=False, group=False):
    """Only expose committed results; roll back failed SheetGen operations."""
    DB = api()
    if group:
        t = DB.TransactionGroup(doc, "T3Lab: " + name)
    elif subtransaction:
        t = DB.SubTransaction(doc)
    else:
        t = DB.Transaction(doc, "T3Lab: " + name)
    try:
        t.Start()
        if t.GetStatus() != DB.TransactionStatus.Started:
            raise RuntimeError("Could not start " + name)
        if not group and not subtransaction:
            options = t.GetFailureHandlingOptions()
            options.SetForcedModalHandling(True)
            options.SetClearAfterRollback(True)
            t.SetFailureHandlingOptions(options)
        yield t
        status = t.Assimilate() if group else t.Commit()
        if status == DB.TransactionStatus.Pending:
            raise SheetGenPending("Revit is awaiting failure resolution: " + name)
        if status != DB.TransactionStatus.Committed:
            raise RuntimeError("{} was not committed: {}".format(name, status))
    except SheetGenPending:
        # Pending child transactions also prevent rolling back their group.
        raise
    except BaseException:
        if t.GetStatus() == DB.TransactionStatus.Started:
            t.RollBack()
        raise
    finally:
        if t.GetStatus() not in (DB.TransactionStatus.Started,
                                 DB.TransactionStatus.Pending):
            t.Dispose()


# ── pure geometry / naming (preview and placement share these) ─────────────
def usable_rect(x0, y0, w, h, margin, strip_side, strip):
    """Printable rect (x0, y0, w, h) inside a sheet after the margin and
    the title-block header strip (vertical right / horizontal bottom)."""
    ux0 = x0 + margin
    uy0 = y0 + margin
    uw = w - 2.0 * margin
    uh = h - 2.0 * margin
    if strip_side == 'right':
        uw -= strip
    elif strip_side == 'bottom':
        uy0 += strip
        uh -= strip
    return ux0, uy0, max(uw, 0.1), max(uh, 0.1)


def clamp_center(cx, cy, w_paper, h_paper, usable):
    """Clamp a box center so the whole box stays inside the usable rect.
    Returns (cx, cy, fits) - fits is False when the box is larger than
    the rect (it then gets centered so the spill is symmetric)."""
    ux0, uy0, uw, uh = usable
    fits = True
    if w_paper >= uw:
        cx = ux0 + uw / 2.0
        fits = False
    else:
        cx = min(max(cx, ux0 + w_paper / 2.0), ux0 + uw - w_paper / 2.0)
    if h_paper >= uh:
        cy = uy0 + uh / 2.0
        fits = False
    else:
        cy = min(max(cy, uy0 + h_paper / 2.0), uy0 + uh - h_paper / 2.0)
    return cx, cy, fits


def slot_centers(usable, plan_count, combined):
    """Viewport slot centers inside a usable rect: (plan_centers, elev_centers)
    as (x, y) sheet-feet tuples."""
    ux0, uy0, uw, uh = usable
    plans, elevs = [], []
    if combined:
        px = ux0 + uw * 0.225  # middle of the left 45% plan zone
        if plan_count == 1:
            plans.append((px, uy0 + uh * 0.5))
        elif plan_count >= 2:
            plans.append((px, uy0 + uh * 0.72))
            plans.append((px, uy0 + uh * 0.28))
        for fx, fy in ((0.625, 0.75), (0.875, 0.75), (0.625, 0.25), (0.875, 0.25)):
            elevs.append((ux0 + uw * fx, uy0 + uh * fy))
    else:
        if plan_count == 1:
            plans.append((ux0 + uw * 0.5, uy0 + uh * 0.5))
        elif plan_count >= 2:
            plans.append((ux0 + uw * 0.5, uy0 + uh * 0.70))
            plans.append((ux0 + uw * 0.5, uy0 + uh * 0.28))
        for fx, fy in ((0.25, 0.75), (0.75, 0.75), (0.25, 0.25), (0.75, 0.25)):
            elevs.append((ux0 + uw * fx, uy0 + uh * fy))
    return plans, elevs


def build_view_name(room, copy_index=0):
    """Plan view (and sheet) name from room info."""
    if "UNIT" in room.Name.upper() and room.RoomType:
        base = "ENLARGED PLAN - TYPE {} ({})".format(room.RoomType, room.Name)
    else:
        base = "ENLARGED PLAN - {} - (#{})".format(room.Name, room.Number)
    if copy_index > 0:
        return "{} - Copy {}".format(base, copy_index)
    return base


def build_sheet_number(room, suffix=""):
    """EPL-101, EPL-101-P or EPL-101-E."""
    base = "EPL-{}".format(room.Number)
    return "{}-{}".format(base, suffix) if suffix else base


def elevation_name(room, label):
    return "INTERIOR ELEV - {} - {} ({})".format(room.Name, label, room.Number)


def direction_label(view):
    """Label from the elevation's real ViewDirection (marker index order is not
    guaranteed by the API). ViewDirection points toward the viewer: a view of
    the north wall reports -Y."""
    try:
        d = view.ViewDirection
        if abs(d.Y) >= abs(d.X):
            return "North" if d.Y < 0 else "South"
        return "East" if d.X < 0 else "West"
    except Exception:
        return "View"


def view_type_name(vt):
    """Type name of a ViewFamilyType, defensively."""
    try:
        p = vt.get_Parameter(api().BuiltInParameter.SYMBOL_NAME_PARAM)
        name = p.AsString() if p else None
        if name:
            return name
    except Exception:
        pass
    try:
        return vt.Name
    except Exception:
        return str(vt.Id)


# ── model lookups (the dialog's combo boxes and the tool read the same) ──
class RoomInfo(object):
    """What SheetGen needs from a room (the dialog's RoomItem has the same fields)."""

    def __init__(self, room, doc=None):
        DB = api()
        self.Element = room
        p = room.get_Parameter(DB.BuiltInParameter.ROOM_NUMBER)
        self.Number = (p.AsString() or "") if p else ""
        p = room.get_Parameter(DB.BuiltInParameter.ROOM_NAME)
        self.Name = (p.AsString() or "") if p else ""
        p = room.LookupParameter("Room Type")
        self.RoomType = (p.AsString() or "") if p else ""
        self.GenQty = 1


def find_view_types(doc):
    """(floor plan, ceiling plan, elevation) ViewFamilyType ids, None when missing."""
    DB = api()
    floor = ceiling = elevation = None
    for vt in DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType) \
            .WhereElementIsElementType().ToElements():
        if vt.FamilyName == 'Floor Plan' and floor is None:
            floor = vt.Id
        elif vt.FamilyName == 'Ceiling Plan' and ceiling is None:
            ceiling = vt.Id
        elif vt.ViewFamily == DB.ViewFamily.Elevation and elevation is None:
            elevation = vt.Id
    return floor, ceiling, elevation


def template_maps(doc):
    """{'floor' | 'ceiling' | 'elevation': {template name: id}}."""
    DB = api()
    kinds = (('floor', DB.ViewType.FloorPlan), ('ceiling', DB.ViewType.CeilingPlan),
             ('elevation', DB.ViewType.Elevation))
    out = {'floor': {}, 'ceiling': {}, 'elevation': {}}
    for v in DB.FilteredElementCollector(doc).OfClass(DB.View) \
            .WhereElementIsNotElementType().ToElements():
        if not v.IsTemplate:
            continue
        for kind, view_type in kinds:
            if v.ViewType == view_type:
                out[kind][v.Name] = v.Id
    return out


def titleblock_map(doc):
    """{'Family: Type': type id} of every title block type."""
    DB = api()
    out = {}
    for tb in DB.FilteredElementCollector(doc).OfCategory(DB.BuiltInCategory.OST_TitleBlocks) \
            .WhereElementIsElementType().ToElements():
        try:
            fam_name = tb.Family.Name
            p = tb.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM)
            type_name = p.AsString() if p else ""
            display = "{}: {}".format(fam_name, type_name) if type_name else fam_name
        except Exception:
            display = str(tb.Id)
        out[display] = tb.Id
    return out


DEFAULTS = {
    'floor': True, 'ceiling': False, 'elevations': False,
    'floor_type_id': None, 'ceiling_type_id': None, 'elevation_type_id': None,
    'floor_template_id': None, 'ceiling_template_id': None, 'elevation_template_id': None,
    'offset_ft': 1000.0 / 304.8, 'cropbox_visible': True,
    'strip_side': 'right', 'strip_ft': 70.0 / 304.8,
}


class RoomSheetGen(object):
    """One SheetGen run on `doc`; options: see DEFAULTS. Counts views in
    `created`, failures in `errors` (messages in `problems`)."""

    def __init__(self, doc, options=None, log=None):
        self.doc = doc
        self.opt = dict(DEFAULTS)
        self.opt.update(options or {})
        self.log = log or _NoLog()
        self.created = 0
        self.errors = 0
        self.problems = []
        self._used_view_names = {}
        self._used_sheet_numbers = set()
        self._resolved_elev_type = None

    def _fail(self, text):
        self.errors += 1
        self.problems.append(text)
        self.log.error(text)

    def seed_names(self):
        """Seed name/number trackers from the document, so a re-run on a room
        that has views/sheets auto-suffixes instead of "Name must be unique"."""
        DB = api()
        self._used_view_names = {}
        for v in DB.FilteredElementCollector(self.doc).OfClass(DB.View) \
                .WhereElementIsNotElementType().ToElements():
            self._used_view_names.setdefault(v.ViewType, set()).add(v.Name)
        self._used_sheet_numbers = set()
        for sh in DB.FilteredElementCollector(self.doc).OfClass(DB.ViewSheet).ToElements():
            p = sh.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER)
            if p:
                self._used_sheet_numbers.add(p.AsString())
        self._resolved_elev_type = None          # model may have changed

    def _unique_view_name(self, base_name, view_type):
        """base_name, else "base_name (n)"; unique per ViewType (a Floor Plan and
        a Ceiling Plan may share a name), including names handed out this run."""
        used = self._used_view_names.setdefault(view_type, set())
        name = base_name
        i = 2
        while name in used:
            name = "{} ({})".format(base_name, i)
            i += 1
        used.add(name)
        return name

    def _unique_sheet_number(self, base_number):
        number = base_number
        i = 2
        while number in self._used_sheet_numbers:
            number = "{}-{}".format(base_number, i)
            i += 1
        self._used_sheet_numbers.add(number)
        return number

    @staticmethod
    def _offset_bbox(bbox, offset=1):
        DB = api()
        new_bbox = DB.BoundingBoxXYZ()
        new_bbox.Min = DB.XYZ(bbox.Min.X - offset, bbox.Min.Y - offset, bbox.Min.Z - offset)
        new_bbox.Max = DB.XYZ(bbox.Max.X + offset, bbox.Max.Y + offset, bbox.Max.Z + offset)
        return new_bbox

    def _find_plan_view_for_level(self, level_id):
        """A floor plan of the level to host the elevation marker; skips this
        tool's own ENLARGED PLAN views, prefers an uncropped one."""
        DB = api()
        fallback = None
        for v in DB.FilteredElementCollector(self.doc).OfClass(DB.ViewPlan) \
                .WhereElementIsNotElementType().ToElements():
            if (v.IsTemplate or v.ViewType != DB.ViewType.FloorPlan
                    or v.GenLevel is None or v.GenLevel.Id != level_id):
                continue
            if v.Name.upper().startswith(ENLARGED_PREFIX):
                continue
            if not v.CropBoxActive:
                return v
            if fallback is None:
                fallback = v
        return fallback

    def _create_interior_elevation_view(self, marker, host_plan, idx,
                                        cropbox_visible, max_dim, offset, elev_template_id):
        """One elevation at the marker index; named later by _finalize_elevation_name
        (after the fallback rotation)."""
        DB = api()
        ev = marker.CreateElevation(self.doc, host_plan.Id, idx)
        ev.CropBoxActive = True
        ev.CropBoxVisible = cropbox_visible
        p_far = ev.get_Parameter(DB.BuiltInParameter.VIEWER_BOUND_FAR_CLIPPING)
        if p_far:
            p_far.Set(1)
        p_off = ev.get_Parameter(DB.BuiltInParameter.VIEWER_BOUND_OFFSET_FAR)
        if p_off:
            p_off.Set(max_dim / 2 + offset)
        if elev_template_id:
            ev.ViewTemplateId = elev_template_id
        return ev

    def _finalize_elevation_name(self, ev, room_item):
        """Name from ViewDirection; call after the marker's final rotation + Regenerate."""
        label = direction_label(ev)
        ev.Name = self._unique_view_name(elevation_name(room_item, label), ev.ViewType)
        return label

    def _resolve_elevation_type(self):
        """(type_id, type_name, capacity) of the elevation type whose marker hosts
        the most views. A single-slot marker ('Building Elevation') makes
        CreateElevation on index 1..3 throw, and capacity is an instance property,
        so each type is probed with a temp marker in a rolled-back transaction.
        Prefers capacity >= 4 with 'interior' in the name. Cached per run."""
        if self._resolved_elev_type:
            return self._resolved_elev_type
        DB = api()
        candidates = []
        elev_types = [vt for vt in DB.FilteredElementCollector(self.doc)
                      .OfClass(DB.ViewFamilyType).WhereElementIsElementType().ToElements()
                      if vt.ViewFamily == DB.ViewFamily.Elevation]
        probe = DB.Transaction(self.doc, "T3Lab: Probe Elevation Marker Capacity")
        try:
            probe.Start()
            for vt in elev_types:
                name = view_type_name(vt)
                try:
                    m = DB.ElevationMarker.CreateElevationMarker(
                        self.doc, vt.Id, DB.XYZ(0, 0, 0), 50)
                    candidates.append((vt.Id, name, m.MaximumViewCount))
                except Exception as ex:
                    self.log.debug("Elevation type '{}' probe failed: {}".format(name, ex))
        except Exception as ex:
            self.log.debug("Elevation capacity probe failed: {}".format(ex))
        finally:
            try:
                if probe.HasStarted():
                    probe.RollBack()
            except Exception:
                pass
        if candidates:
            self._resolved_elev_type = max(candidates, key=lambda c: (
                c[2] >= 4, "INTERIOR" in c[1].upper() if c[2] >= 4 else False, c[2]))
        else:
            # Nothing usable: first-found type, single-slot fallback path.
            self._resolved_elev_type = (self.opt['elevation_type_id'], "<first found>", 1)
        return self._resolved_elev_type

    def _log_marker_debug(self, marker, type_name):
        try:
            cap = marker.MaximumViewCount
        except Exception:
            cap = "?"
        states = []
        for idx in range(4):
            try:
                states.append("{}={}".format(idx, marker.IsAvailableIndex(idx)))
            except Exception as ex:
                states.append("{}=err({})".format(idx, ex))
        self.log.debug("Elevation marker (type '{}'): MaximumViewCount={}, "
                       "IsAvailableIndex[{}]".format(type_name, cap, ", ".join(states)))

    # ── views per room ──────────────────────────────────────────────────────
    def _create_plan(self, label, type_id, level_id, bbox, name, template_id):
        DB = api()
        with transaction(self.doc, "Create " + label):
            vp = DB.ViewPlan.Create(self.doc, type_id, level_id)
            vp.CropBoxActive = True
            vp.CropBoxVisible = self.opt['cropbox_visible']
            vp.CropBox = bbox
            vp.Name = name
            if template_id:
                vp.ViewTemplateId = template_id
        return vp

    def create_room(self, room_item):
        """Views for one room: {'room_item', 'floor_plan', 'ceiling_plan',
        'elevations'}, or None when the room has no bounding box. Each view is
        its own committed transaction; SheetGenPending propagates."""
        DB = api()
        o = self.opt
        room = room_item.Element
        room_level_id = room.LevelId
        room_bbox = room.get_BoundingBox(None)
        if room_bbox is None:
            self._fail("Room {} ('{}') has no bounding box - it is most likely unplaced, "
                       "unenclosed or redundant. Nothing was created for it.".format(
                           room_item.Number, room_item.Name))
            return None
        offset = o['offset_ft']
        new_bbox = self._offset_bbox(room_bbox, offset)
        try:
            qty = max(int(room_item.GenQty), 0)
        except Exception:
            qty = 1
        result = {'room_item': room_item, 'floor_plan': None, 'ceiling_plan': None,
                  'elevations': []}

        if o['floor'] and o['floor_type_id'] and qty > 0:
            for idx in range(qty):
                view_name = self._unique_view_name(
                    build_view_name(room_item, copy_index=idx), DB.ViewType.FloorPlan)
                try:
                    vp = self._create_plan("Floor Plan", o['floor_type_id'], room_level_id,
                                           new_bbox, view_name, o['floor_template_id'])
                    if result['floor_plan'] is None:   # first copy goes on the sheet
                        result['floor_plan'] = vp
                    self.created += 1
                except Exception as ex:
                    self._fail("Floor plan error for {}: {}".format(view_name, ex))

        if o['ceiling'] and o['ceiling_type_id']:
            name = self._unique_view_name(build_view_name(room_item), DB.ViewType.CeilingPlan)
            try:
                result['ceiling_plan'] = self._create_plan(
                    "Ceiling Plan", o['ceiling_type_id'], room_level_id, new_bbox, name,
                    o['ceiling_template_id'])
                self.created += 1
            except Exception as ex:
                self._fail("Ceiling plan error for {}: {}".format(name, ex))

        if o['elevations'] and o['elevation_type_id']:
            try:
                self._create_elevations(room_item, room, room_bbox, offset, result)
            except Exception as ex:
                result['elevations'] = []
                self._fail("Elevation error for room {}: {}".format(room_item.Number, ex))
        return result

    def _create_elevations(self, room_item, room, room_bbox, offset, result):
        DB = api()
        host_plan = self._find_plan_view_for_level(room.LevelId)
        if host_plan is None:
            self._fail("No floor plan found for level to host elevation marker")
            return
        center = room.Location.Point
        max_dim = max(room_bbox.Max.X - room_bbox.Min.X,
                      room_bbox.Max.Y - room_bbox.Min.Y) + offset * 2
        elev_template_id = self.opt['elevation_template_id']
        cropbox_visible = self.opt['cropbox_visible']
        # Most capable marker type; nothing hosts 4 -> 4 single markers rotated 90 deg.
        elev_type_id, elev_type_name, elev_capacity = self._resolve_elevation_type()
        use_fallback = elev_capacity < 4
        if use_fallback:
            self.log.warning(
                "No elevation type with a 4-view marker in this model (best: '{}' hosts {}). "
                "Using 4 single markers rotated 90 degrees apart.".format(
                    elev_type_name, elev_capacity))

        with transaction(self.doc, "Create Interior Elevations"):
            scale = host_plan.Scale
            if not use_fallback:
                marker = DB.ElevationMarker.CreateElevationMarker(
                    self.doc, elev_type_id, center, scale)
                self.doc.Regenerate()
                self._log_marker_debug(marker, elev_type_name)
                # IsAvailableIndex() differs across versions (2023/2025): try each
                # index and let the API reject a slot.
                for idx in range(4):
                    try:
                        with transaction(self.doc, "Create elevation direction",
                                         subtransaction=True):
                            ev = self._create_interior_elevation_view(
                                marker, host_plan, idx, cropbox_visible, max_dim,
                                offset, elev_template_id)
                            self.doc.Regenerate()
                            self._finalize_elevation_name(ev, room_item)
                        result['elevations'].append(ev)
                    except Exception as ex:
                        self._fail(
                            "Elevation index {} failed for room {} [host plan: '{}', type: "
                            "'{}', marker capacity: {}]: {}: {}".format(
                                idx, room_item.Number, host_plan.Name, elev_type_name,
                                elev_capacity, type(ex).__name__, ex))
            else:
                z_axis_top = DB.XYZ(center.X, center.Y, center.Z + 1.0)
                for i in range(4):
                    try:
                        with transaction(self.doc, "Create elevation direction",
                                         subtransaction=True):
                            marker_i = DB.ElevationMarker.CreateElevationMarker(
                                self.doc, elev_type_id, center, scale)
                            ev = self._create_interior_elevation_view(
                                marker_i, host_plan, 0, cropbox_visible, max_dim,
                                offset, elev_template_id)
                            if i > 0:
                                axis = DB.Line.CreateBound(center, z_axis_top)
                                DB.ElementTransformUtils.RotateElement(
                                    self.doc, marker_i.Id, axis, i * math.pi / 2.0)
                            self.doc.Regenerate()
                            self._finalize_elevation_name(ev, room_item)
                        result['elevations'].append(ev)
                    except Exception as ex:
                        self._fail(
                            "Fallback elevation {} of 4 failed for room {} [host plan: "
                            "'{}', type: '{}']: {}: {}".format(
                                i + 1, room_item.Number, host_plan.Name, elev_type_name,
                                type(ex).__name__, ex))
            if not result['elevations']:
                raise RuntimeError("No interior elevations could be created")
        self.created += len(result['elevations'])

    # ── sheets ──────────────────────────────────────────────────────────────
    def _get_sheet_rect(self, sheet):
        """(x0, y0, w, h) of the title block on a sheet in feet, from its real bbox:
        title blocks do not always start at (0,0). A1 at the origin otherwise."""
        DB = api()
        try:
            tb_elems = DB.FilteredElementCollector(self.doc, sheet.Id) \
                .OfCategory(DB.BuiltInCategory.OST_TitleBlocks).ToElements()
            if tb_elems:
                bb = tb_elems[0].get_BoundingBox(sheet)
                if bb:
                    w = bb.Max.X - bb.Min.X
                    h = bb.Max.Y - bb.Min.Y
                    if w > 0 and h > 0:
                        return (bb.Min.X, bb.Min.Y, w, h)
        except Exception:
            pass
        return (0.0, 0.0) + A1_FT

    def _sheet_name(self, room_item, name_suffix=""):
        name = build_view_name(room_item)
        if name_suffix:
            name = "{} — {}".format(name, name_suffix)
        return self._unique_view_name(name, api().ViewType.DrawingSheet)

    def _create_sheet(self, room_item, titleblock_id, name_suffix="", num_suffix=""):
        DB = api()
        sheet_name = self._sheet_name(room_item, name_suffix)
        sheet_num = self._unique_sheet_number(build_sheet_number(room_item, num_suffix))
        with transaction(self.doc, "Create Sheet"):
            sheet = DB.ViewSheet.Create(self.doc, titleblock_id)
            sheet.Name = sheet_name
            p_num = sheet.get_Parameter(DB.BuiltInParameter.SHEET_NUMBER)
            if p_num is None or p_num.IsReadOnly:
                raise RuntimeError("Sheet number is unavailable or read-only")
            if not p_num.Set(sheet_num):
                raise RuntimeError("Could not assign sheet number " + sheet_num)
        return sheet

    def _place_viewport_centered(self, sheet, view_id, cx, cy, usable=None):
        """Viewport at (cx, cy), nudged back inside the usable rect by its real
        paper footprint. Returns the committed Viewport; raises on failure."""
        DB = api()
        try:
            with transaction(self.doc, "Place Viewport"):
                vp = DB.Viewport.Create(self.doc, sheet.Id, view_id, DB.XYZ(cx, cy, 0))
                if vp is None:
                    raise RuntimeError("Revit did not create a viewport")
                if usable is not None:
                    self.doc.Regenerate()
                    try:
                        box = vp.GetBoxOutline()
                        w_p = box.MaximumPoint.X - box.MinimumPoint.X
                        h_p = box.MaximumPoint.Y - box.MinimumPoint.Y
                        bcx = (box.MaximumPoint.X + box.MinimumPoint.X) / 2.0
                        bcy = (box.MaximumPoint.Y + box.MinimumPoint.Y) / 2.0
                        nx, ny, fits = clamp_center(bcx, bcy, w_p, h_p, usable)
                        if abs(nx - bcx) > 1e-6 or abs(ny - bcy) > 1e-6:
                            vp.SetBoxCenter(DB.XYZ(nx, ny, 0))
                        if not fits:
                            self.log.debug("Viewport of view {} is larger than the usable "
                                           "area of sheet {}.".format(view_id, sheet.Id))
                    except Exception as ex:
                        self.log.debug("Viewport clamp skipped: {}".format(ex))
            return vp
        except Exception as ex:
            self.log.error("Viewport place error: {}".format(ex))
            raise

    def _usable(self, sheet):
        x0, y0, w, h = self._get_sheet_rect(sheet)
        return usable_rect(x0, y0, w, h, SHEET_MARGIN_FT, self.opt['strip_side'],
                           self.opt['strip_ft'])

    @staticmethod
    def _plan_views(result):
        return [v for v in (result.get('floor_plan'), result.get('ceiling_plan'))
                if v is not None]

    def _layout_combined(self, result, sheet):
        """Plans left, 4 elevations right (2x2)."""
        usable = self._usable(sheet)
        plan_views = self._plan_views(result)
        plan_centers, elev_centers = slot_centers(usable, len(plan_views), True)
        for v, c in zip(plan_views, plan_centers):
            self._place_viewport_centered(sheet, v.Id, c[0], c[1], usable)
        for ev, c in zip(result.get('elevations', [])[:4], elev_centers):
            self._place_viewport_centered(sheet, ev.Id, c[0], c[1], usable)

    def _layout_separate(self, result, plan_sheet, elev_sheet):
        if plan_sheet:
            usable = self._usable(plan_sheet)
            plan_views = self._plan_views(result)
            plan_centers, _unused = slot_centers(usable, len(plan_views), False)
            for v, c in zip(plan_views, plan_centers):
                self._place_viewport_centered(plan_sheet, v.Id, c[0], c[1], usable)
        if elev_sheet:
            usable = self._usable(elev_sheet)
            _unused, elev_centers = slot_centers(usable, 0, False)
            for ev, c in zip(result.get('elevations', [])[:4], elev_centers):
                self._place_viewport_centered(elev_sheet, ev.Id, c[0], c[1], usable)

    def layout(self, result, titleblock_id, combined):
        """Sheets of one room in one TransactionGroup: all or nothing."""
        with transaction(self.doc, "Layout room sheets", group=True):
            sheets = self._layout_room_sheets(result, titleblock_id, combined)
        return sheets

    def _layout_room_sheets(self, result, titleblock_id, combined):
        room_item = result['room_item']
        has_plans = result['floor_plan'] or result['ceiling_plan']
        has_elevs = bool(result['elevations'])
        sheets = []
        if combined:
            if has_plans or has_elevs:
                sheet = self._create_sheet(room_item, titleblock_id)
                self._layout_combined(result, sheet)
                sheets.append(sheet)
            return sheets
        p_sheet = e_sheet = None
        if has_plans:
            p_sheet = self._create_sheet(room_item, titleblock_id,
                                         name_suffix="Plans", num_suffix="P")
            sheets.append(p_sheet)
        if has_elevs:
            e_sheet = self._create_sheet(room_item, titleblock_id,
                                         name_suffix="Elevations", num_suffix="E")
            sheets.append(e_sheet)
        self._layout_separate(result, p_sheet, e_sheet)
        return sheets

    # ── dry run ─────────────────────────────────────────────────────────────
    def plan(self, room_item, layout=None):
        """What create_room + layout would make for one room, without writing:
        {'room', 'views': [names], 'sheets': ['number - name'], 'problems'?}.
        Elevation directions are only known once the marker exists."""
        DB = api()
        o = self.opt
        row = {'room': '{} {}'.format(room_item.Number, room_item.Name).strip(),
               'views': [], 'sheets': []}
        room = room_item.Element
        if room.get_BoundingBox(None) is None:
            row['problems'] = ['no bounding box (unplaced or unenclosed room)']
            return row
        has_plans = False
        if o['floor'] and o['floor_type_id']:
            for idx in range(max(int(room_item.GenQty or 0), 0)):
                row['views'].append(self._unique_view_name(
                    build_view_name(room_item, copy_index=idx), DB.ViewType.FloorPlan))
                has_plans = True
        if o['ceiling'] and o['ceiling_type_id']:
            row['views'].append(self._unique_view_name(build_view_name(room_item),
                                                       DB.ViewType.CeilingPlan))
            has_plans = True
        has_elevs = False
        if o['elevations'] and o['elevation_type_id']:
            if self._find_plan_view_for_level(room.LevelId) is None:
                row['problems'] = ['no floor plan on its level to host the elevation marker']
            else:
                has_elevs = True
                for label in DIRECTIONS:
                    row['views'].append(self._unique_view_name(
                        elevation_name(room_item, label), DB.ViewType.Elevation))
        if layout == 'combined' and (has_plans or has_elevs):
            parts = [("", "")]
        elif layout == 'separate':
            parts = ([("Plans", "P")] if has_plans else []) + \
                    ([("Elevations", "E")] if has_elevs else [])
        else:
            parts = []
        for name_suffix, num_suffix in parts:
            number = self._unique_sheet_number(build_sheet_number(room_item, num_suffix))
            row['sheets'].append('{} - {}'.format(number, self._sheet_name(room_item,
                                                                           name_suffix)))
        return row
