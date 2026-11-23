# -*- coding: utf-8 -*-
"""Floors under rooms (Room to Floor) and under doors (Door Threshold).

Revit side of GUI/RoomToFloorDialog.py, GUI/DoorThresholdDialog.py and the
room_to_floor / create_door_thresholds tools (core/tools/model.py). Revit is
imported inside functions: the MCP server (IronPython 2.7) and the tests load
this module without it. Test: dev/test_tools_model_domain.py.
"""
import logging

MM_TO_FT = 1.0 / 304.8
_log = logging.getLogger('T3Lab.finish_floors')
_SWALLOWER = None


def _warning_swallower():
    """IFailuresPreprocessor that dismisses warnings (class built once, lazily)."""
    global _SWALLOWER
    if _SWALLOWER is None:
        from Autodesk.Revit.DB import (IFailuresPreprocessor, FailureProcessingResult,
                                       FailureSeverity)
        from Snippets._compat import clr_namespace

        class FinishFloorWarnings(IFailuresPreprocessor):
            __namespace__ = clr_namespace('FinishFloors')

            def PreprocessFailures(self, accessor):
                for failure in accessor.GetFailureMessages():
                    if failure.GetSeverity() == FailureSeverity.Warning:
                        accessor.DeleteWarning(failure)
                return FailureProcessingResult.Continue

        _SWALLOWER = FinishFloorWarnings
    return _SWALLOWER()


def _silence(transaction):
    opts = transaction.GetFailureHandlingOptions()
    opts.SetFailuresPreprocessor(_warning_swallower())
    transaction.SetFailureHandlingOptions(opts)


def _revit_version(doc):
    from Snippets._compat import revit_year
    return revit_year(doc)


# ── floor types ─────────────────────────────────────────────────────────────

def floor_type_label(floor_type):
    """'Family: Type' as both windows list floor types."""
    try:
        p = floor_type.LookupParameter("Type Name")
        type_name = p.AsString() if p else floor_type.Name
    except Exception:
        type_name = floor_type.Name
    return u"{}: {}".format(floor_type.FamilyName or u"", type_name or u"")


def floor_types(doc):
    """{'Family: Type': FloorType} of the model."""
    from Autodesk.Revit.DB import FilteredElementCollector, FloorType
    try:
        types = FilteredElementCollector(doc).OfClass(FloorType).ToElements()
    except Exception:
        types = []
    return dict((floor_type_label(t), t) for t in types)


def pick_floor_type(types, wanted=None):
    """(type, label) from floor_types(); (None, None) when nothing matches.

    No name = the first label (the windows' default). Matching, case-insensitive:
    the full label, then the type name, then a part of the label.
    """
    from Snippets._compat import fold
    labels = sorted(types)
    if not labels:
        return None, None
    if not wanted:
        return types[labels[0]], labels[0]
    key = fold(u"{}".format(wanted).strip())
    for test in (lambda l: fold(l) == key,
                 lambda l: fold(l.split(u": ", 1)[-1]) == key,
                 lambda l: key in fold(l)):
        hits = [l for l in labels if test(l)]
        if hits:
            return types[hits[0]], hits[0]
    return None, None


# ── Room to Floor ───────────────────────────────────────────────────────────

def room_problem(room):
    """Why no floor can be built under `room`, or None."""
    from Autodesk.Revit.DB import BuiltInParameter
    if room is None:
        return "not in the model"
    if not hasattr(room, "GetBoundarySegments"):
        return "not a room"
    try:
        area = room.get_Parameter(BuiltInParameter.ROOM_AREA)
        if not area or not area.AsDouble():
            return "not placed or not enclosed (area 0)"
    except Exception:
        return "area unreadable"
    return None


class FloorGenerator(object):
    """Creates one floor per room from its boundary (Room to Floor).

    failures: [(room, reason)] of the last generate_floors run.
    """

    def __init__(self, doc):
        self.doc = doc
        self.failures = []
        self.last_error = None

    def generate_floors(self, room_elements, floor_type, offset_mm=0.0, is_structural=False, use_finish=True,
                        progress_callback=None, cancel_check=None):
        from Autodesk.Revit.DB import TransactionGroup
        from Snippets._compat import disposing
        offset_ft = offset_mm / 304.8
        is_structural_bool = bool(is_structural) if is_structural is not None else False
        created_count = 0
        error_count = 0
        new_floors = []
        self.failures = []

        with disposing(TransactionGroup(self.doc, "T3Lab: Room to Floor")) as tg:
            tg.Start()
            total = len(room_elements)
            for idx, room in enumerate(room_elements):
                if cancel_check and cancel_check():
                    break
                if progress_callback:
                    progress_callback(idx, total)
                floor = self._create_one_floor(room, floor_type, offset_ft, is_structural_bool, use_finish)
                if floor:
                    new_floors.append(floor)
                    created_count += 1
                else:
                    error_count += 1
                    self.failures.append((room, self.last_error or "no floor created"))
            tg.Assimilate()

        return new_floors, created_count, error_count

    def _fail(self, reason):
        self.last_error = reason
        return None

    def _create_one_floor(self, room, floor_type, offset_ft, is_structural_bool, use_finish):
        """Create a floor for a single room in its own Transaction. Returns floor or None."""
        from Autodesk.Revit.DB import (BuiltInParameter, CurveArray, CurveLoop, Floor,
                                       SpatialElementBoundaryLocation,
                                       SpatialElementBoundaryOptions, Transaction)
        from Snippets._compat import disposing
        self.last_error = None
        try:
            area_param = room.get_Parameter(BuiltInParameter.ROOM_AREA)
            if not area_param or not area_param.AsDouble():
                return self._fail("not placed or not enclosed (area 0)")

            level_id = room.LevelId
            level = self.doc.GetElement(level_id)
            if not level:
                return self._fail("no level")

            opt = SpatialElementBoundaryOptions()
            opt.SpatialElementBoundaryLocation = (
                SpatialElementBoundaryLocation.Finish if use_finish
                else SpatialElementBoundaryLocation.Center
            )

            room_boundaries = room.GetBoundarySegments(opt)
            if not room_boundaries:
                return self._fail("no boundary")

            new_floor = None

            if _revit_version(self.doc) >= 2022:
                from System.Collections.Generic import List
                with disposing(Transaction(self.doc, "T3Lab: Create Floor")) as t:
                    t.Start()
                    profile = List[CurveLoop]()
                    for loop in room_boundaries:
                        curve_loop = CurveLoop()
                        for seg in loop:
                            curve = seg.GetCurve()
                            if curve:
                                curve_loop.Append(curve)
                        if not curve_loop.IsOpen():
                            profile.Add(curve_loop)

                    if not profile.Count:
                        return self._fail("boundary is not closed")

                    new_floor = Floor.Create(self.doc, profile, floor_type.Id, level_id)

                    if new_floor:
                        if abs(offset_ft) > 0.0001:
                            param = new_floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
                            if param:
                                param.Set(offset_ft)
                        if is_structural_bool:
                            struct_param = new_floor.get_Parameter(BuiltInParameter.FLOOR_PARAM_IS_STRUCTURAL)
                            if struct_param:
                                struct_param.Set(1)

                    _silence(t)
                    t.Commit()

            else:
                floor_shape = room_boundaries[0]
                openings = list(room_boundaries)[1:] if len(room_boundaries) > 1 else []

                with disposing(Transaction(self.doc, "T3Lab: Create Floor")) as t:
                    t.Start()
                    curve_array = CurveArray()
                    for seg in floor_shape:
                        curve = seg.GetCurve()
                        if curve:
                            curve_array.Append(curve)

                    new_floor = self.doc.Create.NewFloor(curve_array, floor_type, level, is_structural_bool)

                    if new_floor and abs(offset_ft) > 0.0001:
                        param = new_floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
                        if param:
                            param.Set(offset_ft)

                    _silence(t)
                    t.Commit()

                if new_floor and openings:
                    with disposing(Transaction(self.doc, "T3Lab: Create Floor Openings")) as t2:
                        t2.Start()
                        for opening in openings:
                            try:
                                opening_curve = CurveArray()
                                for seg in opening:
                                    curve = seg.GetCurve()
                                    if curve:
                                        opening_curve.Append(curve)
                                self.doc.Create.NewOpening(new_floor, opening_curve, True)
                            except Exception:
                                pass
                        t2.Commit()

            return new_floor if new_floor else self._fail("Revit created no floor")

        except Exception as ex:
            from Snippets._compat import short_error
            _log.debug("Error creating floor for room: %s", ex)
            return self._fail(short_error(ex))


# ── Door Threshold ──────────────────────────────────────────────────────────

def _get_wall_thickness_at_point(wall, point):
    """Measure the actual wall thickness at a given point by projecting it
    onto the wall's exterior and interior side faces."""
    from Autodesk.Revit.DB import HostObjectUtils, ShellLayerType, Wall
    try:
        if not isinstance(wall, Wall):
            return None

        ext_refs = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior)
        int_refs = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Interior)
        if not ext_refs or not int_refs or ext_refs.Count == 0 or int_refs.Count == 0:
            return None

        ext_face = wall.GetGeometryObjectFromReference(ext_refs[0])
        int_face = wall.GetGeometryObjectFromReference(int_refs[0])
        if not ext_face or not int_face:
            return None

        ext_proj = ext_face.Project(point)
        int_proj = int_face.Project(point)
        if not ext_proj or not int_proj:
            return None

        thickness = ext_proj.XYZPoint.DistanceTo(int_proj.XYZPoint)
        if thickness > 0.001:
            return thickness
    except Exception:
        pass
    return None


def _get_door_width_ft(door):
    """Get the door's actual width in feet, trying parameters first and
    falling back to measuring the instance's bounding box along its HandOrientation axis."""
    from Autodesk.Revit.DB import BuiltInParameter, XYZ
    for param in (
        lambda: door.get_Parameter(BuiltInParameter.DOOR_WIDTH),
        lambda: door.Symbol.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM),
        lambda: door.LookupParameter("Width"),
        lambda: door.Symbol.LookupParameter("Width"),
    ):
        try:
            p = param()
            if p and p.AsDouble() > 0.001:
                return p.AsDouble()
        except Exception:
            pass

    try:
        bbox = door.get_BoundingBox(None)
        if bbox:
            v = door.HandOrientation
            corners = [
                XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
                XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
            ]
            projections = [c.DotProduct(v) for c in corners]
            width = max(projections) - min(projections)
            if width > 0.001:
                return width
    except Exception:
        pass

    return 900 * MM_TO_FT


def door_label(door):
    """'Door <Mark>' or 'Door <id>'."""
    from Autodesk.Revit.DB import BuiltInParameter
    from Snippets._compat import eid_value
    mark_param = door.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)
    return "Door {}".format(mark_param.AsString() if (mark_param and mark_param.AsString())
                            else eid_value(door.Id))


class ThresholdGenerator(object):
    """Creates one threshold floor per door, wall-thick and door-wide (Door Threshold)."""

    def __init__(self, doc, unit=None):
        from Snippets._units import project_length_unit
        self.doc = doc
        # Unit the error list is written in (the project's length unit).
        self.unit = unit or project_length_unit(doc)

    def generate_thresholds(self, doors, floor_type, offset_mm):
        """(new floors, created, errors, error messages); one Transaction."""
        from Autodesk.Revit.DB import (BuiltInParameter, CurveLoop, Floor, Line, Transaction,
                                       XYZ)
        from Snippets._compat import disposing
        offset_ft = offset_mm * MM_TO_FT
        modern = _revit_version(self.doc) >= 2022

        created_count = 0
        error_count = 0
        new_floors = []
        error_messages = []

        with disposing(Transaction(self.doc, "T3Lab: Door Threshold")) as t:
            t.Start()
            _silence(t)

            for door in doors:
                label = door_label(door)
                try:
                    level_id = door.LevelId
                    level = self.doc.GetElement(level_id)
                    if not level:
                        error_count += 1
                        error_messages.append("{}: no level found".format(label))
                        continue
                    z = level.Elevation

                    w = _get_door_width_ft(door)

                    loc = door.Location
                    if not loc:
                        error_count += 1
                        error_messages.append("{}: no location point".format(label))
                        continue

                    p = XYZ(loc.Point.X, loc.Point.Y, z)

                    wall = door.Host
                    if wall:
                        thickness = _get_wall_thickness_at_point(wall, p)
                        if thickness is None:
                            thickness = wall.Width
                    else:
                        thickness = 150 * MM_TO_FT

                    v = door.HandOrientation
                    u = door.FacingOrientation

                    half_w = w / 2.0
                    half_d = thickness / 2.0

                    if half_w < 0.001 or half_d < 0.001:
                        error_count += 1
                        error_messages.append(
                            "{}: invalid size (width {}, wall thickness {})".format(
                                label, self.unit.show(w),
                                self.unit.show(thickness)))
                        continue

                    p1 = p - v * half_w - u * half_d
                    p2 = p + v * half_w - u * half_d
                    p3 = p + v * half_w + u * half_d
                    p4 = p - v * half_w + u * half_d

                    c1 = Line.CreateBound(p1, p2)
                    c2 = Line.CreateBound(p2, p3)
                    c3 = Line.CreateBound(p3, p4)
                    c4 = Line.CreateBound(p4, p1)

                    loop = CurveLoop()
                    loop.Append(c1)
                    loop.Append(c2)
                    loop.Append(c3)
                    loop.Append(c4)

                    new_floor = None

                    if modern:
                        from System.Collections.Generic import List
                        profile = List[CurveLoop]()
                        profile.Add(loop)
                        new_floor = Floor.Create(self.doc, profile, floor_type.Id, level_id)
                    else:
                        from Autodesk.Revit.DB import CurveArray
                        curve_array = CurveArray()
                        curve_array.Append(c1)
                        curve_array.Append(c2)
                        curve_array.Append(c3)
                        curve_array.Append(c4)
                        new_floor = self.doc.Create.NewFloor(curve_array, floor_type, level, False)

                    if new_floor:
                        if abs(offset_ft) > 0.0001:
                            param = new_floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
                            if param:
                                param.Set(offset_ft)

                        new_floors.append(new_floor)
                        created_count += 1

                except Exception as ex:
                    _log.error("Error creating threshold for %s: %s", label, ex)
                    error_messages.append("{}: {}".format(label, ex))
                    error_count += 1

            t.Commit()

        return new_floors, created_count, error_count, error_messages

