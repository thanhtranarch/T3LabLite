# -*- coding: utf-8 -*-
"""Point Cloud to Model, Revit side: create walls, floors, ceilings, columns,
openings and roofs from detections (Services/point_cloud_analysis.py).

Shared by the Point Cloud to Model window (GUI/PointCloudDialog.py) and the
build_from_point_cloud AI tool (core/tools/from_source.py). No WPF here; the
Revit API names are imported when Revit is present, so the pure half
(from_detections, region_from_args) loads in tests.
Test: dev/test_tools_create_domain.py.
"""

import math

from Services.point_cloud_analysis import (
    DetectedColumn, DetectedFloor, DetectedOpening, DetectedRoof, DetectedStair,
    DetectedWall, mm_to_ft)
from Snippets._compat import clr_namespace, committed, disposing

try:
    import clr
    import Autodesk.Revit.DB as DB
    from System.Collections.Generic import List
    from Autodesk.Revit.DB import (
        BuiltInCategory, BuiltInParameter, CurveArray, CurveLoop, FamilySymbol,
        FilteredElementCollector, Floor, FloorType, Level, Line, Transaction,
        TransactionGroup, Wall, WallType, XYZ)
    try:
        from Autodesk.Revit.DB import StructuralType
    except ImportError:
        from Autodesk.Revit.DB.Structure import StructuralType
except Exception:   # outside Revit: only the pure half is used
    pass

def _footprint_roof(doc, curves, level, roof_type):
    """NewFootPrintRoof; its last parameter is `out ModelCurveArray`."""
    box = getattr(clr, 'Reference', None)
    if box is not None:                     # IronPython: a StrongBox for the out value
        return doc.Create.NewFootPrintRoof(curves, level, roof_type, box[DB.ModelCurveArray]())
    # pythonnet 3 has no clr.Reference: pass a placeholder, get (roof, curves) back.
    out = doc.Create.NewFootPrintRoof(curves, level, roof_type, DB.ModelCurveArray())
    return out[0] if isinstance(out, tuple) else out


try:
    from pyrevit import script as _script
    logger = _script.get_logger()
except Exception:   # outside pyRevit
    import logging
    logger = logging.getLogger('T3Lab.point_cloud_build')


_SWALLOWER = []


def _warning_swallower():
    """IFailuresPreprocessor class, built once on first use (tests never touch .NET).

    Scan-derived geometry is noisy, so nearly every element raises a warning
    (walls off axis, overlapping corners, openings not fully in their host).
    Without it each one stops the batch on a modal dialog with a
    TransactionGroup open. Warnings are deleted; errors are resolved and the
    commit proceeds.
    """
    if _SWALLOWER:
        return _SWALLOWER[0]
    from Autodesk.Revit.DB import (IFailuresPreprocessor, FailureProcessingResult,
                                   FailureSeverity)

    class WarningSwallower(IFailuresPreprocessor):
        __namespace__ = clr_namespace('PointCloudBuild')

        def PreprocessFailures(self, failuresAccessor):
            fail_list = failuresAccessor.GetFailureMessages()
            if fail_list.Count == 0:
                return FailureProcessingResult.Continue

            has_error = False
            for failure in fail_list:
                severity = failure.GetSeverity()
                if severity == FailureSeverity.Warning:
                    failuresAccessor.DeleteWarning(failure)
                elif severity == FailureSeverity.Error:
                    has_error = True
                    try:
                        failuresAccessor.ResolveFailure(failure)
                    except Exception:
                        pass

            if has_error:
                return FailureProcessingResult.ProceedWithCommit
            return FailureProcessingResult.Continue

    _SWALLOWER.append(WarningSwallower)
    return WarningSwallower


def _eid(element):
    from Snippets._compat import eid_value
    return eid_value(element.Id)


class ElementBuilder(object):
    """Creates Revit elements from DetectedElement data inside named Transactions."""

    def __init__(self, document):
        from Snippets._compat import revit_year
        self.doc         = document
        self._year        = revit_year(document)
        self._wall_types  = self._load_wall_types()
        self._floor_types = self._load_floor_types()
        self._levels      = self._load_levels()

    def _start(self, t):
        """Start a transaction with warnings suppressed.

        Must run after Start(): the options object is per-transaction, and a
        rollback has to clear its own failures or the next commit inherits
        them.
        """
        t.Start()
        try:
            opts = t.GetFailureHandlingOptions()
            opts.SetFailuresPreprocessor(_warning_swallower()())
            opts.SetClearAfterRollback(True)
            t.SetFailureHandlingOptions(opts)
        except Exception as ex:
            logger.debug("failure handling options: {}".format(ex))

    def _load_wall_types(self):
        if self.doc is None:
            return []
        return list(FilteredElementCollector(self.doc).OfClass(WallType).ToElements())

    def _load_floor_types(self):
        if self.doc is None:
            return []
        return list(FilteredElementCollector(self.doc).OfClass(FloorType).ToElements())

    def _load_levels(self):
        if self.doc is None:
            return {}
        lvs = FilteredElementCollector(self.doc).OfClass(Level).ToElements()
        return {lv.Name: lv for lv in lvs}

    def _best_wall_type(self, thickness_ft):
        """Find the WallType whose compound structure width is closest to thickness_ft."""
        best      = None
        best_diff = float('inf')
        for wt in self._wall_types:
            try:
                cs = wt.GetCompoundStructure()
                w  = cs.GetTotalWidth() if cs else wt.Width
                d  = abs(w - thickness_ft)
                if d < best_diff:
                    best_diff = d
                    best = wt
            except Exception:
                pass
        return best or (self._wall_types[0] if self._wall_types else None)

    def _best_floor_type(self):
        return self._floor_types[0] if self._floor_types else None

    def _get_level(self, level_name):
        return self._levels.get(level_name) or (
            sorted(self._levels.values(), key=lambda l: l.Elevation)[0]
            if self._levels else None)

    def _set_offset_from_level(self, elem, bip, offset_ft):
        """Set a height-above-level offset parameter, ignoring failures."""
        try:
            p = elem.get_Parameter(bip)
            if p and not p.IsReadOnly:
                p.Set(offset_ft)
        except Exception:
            pass

    def _rect_curve_loop(self, corners_xy, z_ft):
        """Build a closed rectangular CurveLoop from 4 (x, y) corner tuples."""
        cl  = CurveLoop()
        pts = [XYZ(x, y, z_ft) for (x, y) in corners_xy]
        for i in range(len(pts)):
            p1 = pts[i]
            p2 = pts[(i + 1) % len(pts)]
            if p1.DistanceTo(p2) > 0.01:
                cl.Append(Line.CreateBound(p1, p2))
        return cl

    def build_wall(self, elem, height_ft=None):
        d = elem._data
        angle      = d['angle']
        cx, cy     = d['cx'], d['cy']
        length_ft  = d['length_ft']
        thick_ft   = d['thickness_ft']
        if height_ft is None:
            height_ft = mm_to_ft(3000.0)

        half  = length_ft / 2.0
        pt1   = XYZ(cx - math.cos(angle) * half, cy - math.sin(angle) * half, 0.0)
        pt2   = XYZ(cx + math.cos(angle) * half, cy + math.sin(angle) * half, 0.0)
        wline = Line.CreateBound(pt1, pt2)

        lv = self._get_level(elem.LevelName)
        if not lv:
            return None
        wt = self._best_wall_type(thick_ft)
        if not wt:
            return None

        with disposing(Transaction(self.doc, "T3Lab: Create Wall")) as t:
            self._start(t)
            try:
                wall = Wall.Create(self.doc, wline, wt.Id, lv.Id,
                                   height_ft, 0.0, False, False)
                if not committed(t):
                    return None           # rolled back by Revit
                return wall
            except Exception as ex:
                logger.error("build_wall: {}".format(ex))
                t.RollBack()
                return None

    def build_floor(self, elem):
        d  = elem._data
        lv = self._get_level(elem.LevelName)
        if not lv:
            return None
        ft = self._best_floor_type()
        if not ft:
            return None
        cl = self._rect_curve_loop(d['corners_xy'], d['z_ft'])

        with disposing(Transaction(self.doc, "T3Lab: Create Floor")) as t:
            self._start(t)
            try:
                if self._year >= 2022:
                    profile = List[CurveLoop]()
                    profile.Add(cl)
                    floor = Floor.Create(self.doc, profile, ft.Id, lv.Id)
                else:
                    ca = CurveArray()
                    for curve in cl:
                        ca.Append(curve)
                    floor = self.doc.Create.NewFloor(ca, ft, lv, False)
                self._set_offset_from_level(
                    floor, BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM,
                    d['z_ft'] - lv.Elevation)
                if not committed(t):
                    return None           # rolled back by Revit
                return floor
            except Exception as ex:
                logger.error("build_floor: {}".format(ex))
                t.RollBack()
                return None

    def build_ceiling(self, elem):
        """Create a ceiling (Revit 2022+ Ceiling.Create, else floor-based fallback)."""
        d  = elem._data
        lv = self._get_level(elem.LevelName)
        if not lv:
            return None
        cl = self._rect_curve_loop(d['corners_xy'], d['z_ft'])

        with disposing(Transaction(self.doc, "T3Lab: Create Ceiling")) as t:
            self._start(t)
            try:
                if self._year >= 2022:
                    from Autodesk.Revit.DB import Ceiling, CeilingType
                    ceil_types = FilteredElementCollector(self.doc) \
                        .OfClass(CeilingType).ToElements()
                    if not ceil_types:
                        t.RollBack()
                        return None
                    ct = ceil_types[0]
                    profile = List[CurveLoop]()
                    profile.Add(cl)
                    ceiling = Ceiling.Create(self.doc, profile, ct.Id, lv.Id)
                    self._set_offset_from_level(
                        ceiling, BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM,
                        d['z_ft'] - lv.Elevation)
                else:
                    ft = self._best_floor_type()
                    if not ft:
                        t.RollBack()
                        return None
                    ca = CurveArray()
                    for curve in cl:
                        ca.Append(curve)
                    ceiling = self.doc.Create.NewFloor(ca, ft, lv, False)
                if not committed(t):
                    return None           # rolled back by Revit
                return ceiling
            except Exception as ex:
                logger.error("build_ceiling: {}".format(ex))
                t.RollBack()
                return None

    def build_column(self, elem):
        """Place a structural (or architectural) column family instance."""
        d  = elem._data
        lv = self._get_level(elem.LevelName)
        if not lv:
            return None

        col_symbols = (FilteredElementCollector(self.doc)
                       .OfClass(FamilySymbol)
                       .OfCategory(BuiltInCategory.OST_StructuralColumns)
                       .ToElements())
        if not col_symbols:
            col_symbols = (FilteredElementCollector(self.doc)
                           .OfClass(FamilySymbol)
                           .OfCategory(BuiltInCategory.OST_Columns)
                           .ToElements())
        if not col_symbols:
            return None

        sym = col_symbols[0]
        pt  = XYZ(d['cx'], d['cy'], d['z_bot_ft'])

        with disposing(Transaction(self.doc, "T3Lab: Create Column")) as t:
            self._start(t)
            try:
                if not sym.IsActive:
                    sym.Activate()
                col = self.doc.Create.NewFamilyInstance(
                    pt, sym, lv, StructuralType.Column)
                if not committed(t):
                    return None           # rolled back by Revit
                return col
            except Exception as ex:
                logger.error("build_column: {}".format(ex))
                t.RollBack()
                return None

    def build_opening(self, elem, host_wall_revit_elem):
        """Place a door or window family instance in the given host wall."""
        if host_wall_revit_elem is None:
            return None
        d     = elem._data
        angle = d['host_wall_data']['angle']
        cx    = d['host_wall_data']['cx']
        cy    = d['host_wall_data']['cy']
        u_c   = d['u_center']

        ox = cx + math.cos(angle) * u_c
        oy = cy + math.sin(angle) * u_c
        # insertion point at the opening BOTTOM (door threshold / window sill),
        # not mid-height — Revit derives the level offset from the point Z
        oz = d['w_bottom']
        pt = XYZ(ox, oy, oz)

        cat = (BuiltInCategory.OST_Doors
               if elem.Type == 'Door' else BuiltInCategory.OST_Windows)
        syms = (FilteredElementCollector(self.doc)
                .OfClass(FamilySymbol)
                .OfCategory(cat)
                .ToElements())
        if not syms:
            return None
        sym = syms[0]
        lv  = self._get_level(elem.LevelName)
        if not lv:
            return None

        with disposing(Transaction(self.doc, "T3Lab: Create {}".format(elem.Type))) as t:
            self._start(t)
            try:
                if not sym.IsActive:
                    sym.Activate()
                try:
                    from Autodesk.Revit.DB.Structure import StructuralType as ST
                    non_structural = ST.NonStructural
                except ImportError:
                    non_structural = StructuralType.NonStructural
                inst = self.doc.Create.NewFamilyInstance(
                    pt, sym, host_wall_revit_elem, lv, non_structural)
                if not committed(t):
                    return None           # rolled back by Revit
                return inst
            except Exception as ex:
                logger.error("build_opening: {}".format(ex))
                t.RollBack()
                return None

    def build_roof(self, elem):
        """Create a basic FootPrintRoof from the detected boundary."""
        d  = elem._data
        lv = self._get_level(elem.LevelName)
        if not lv:
            return None

        roof_types = (FilteredElementCollector(self.doc)
                      .OfCategory(BuiltInCategory.OST_Roofs)
                      .OfClass(DB.RoofType)
                      .ToElements())
        if not roof_types:
            return None
        rt = roof_types[0]

        z_ft    = d['z_ft']
        corners = d['corners_xy']
        pts     = [XYZ(x, y, z_ft) for (x, y) in corners]
        n       = len(pts)

        with disposing(Transaction(self.doc, "T3Lab: Create Roof")) as t:
            self._start(t)
            try:
                ca = CurveArray()
                for i in range(n):
                    p1 = pts[i]
                    p2 = pts[(i + 1) % n]
                    if p1.DistanceTo(p2) > 0.01:
                        ca.Append(Line.CreateBound(p1, p2))
                roof = _footprint_roof(self.doc, ca, lv, rt)
                self._set_offset_from_level(
                    roof, BuiltInParameter.ROOF_LEVEL_OFFSET_PARAM,
                    z_ft - lv.Elevation)
                if not committed(t):
                    return None           # rolled back by Revit
                return roof
            except Exception as ex:
                logger.error("build_roof: {}".format(ex))
                t.RollBack()
                return None

    def build_stair_marker(self, elem):
        """Stub — stair creation requires the Architecture API; skipped here."""
        return None

    def build_all(self, elements, default_wall_height_ft=None, created=None):
        """
        Create all selected elements in dependency order, in one
        TransactionGroup (one undo step). `created`, a list, receives the
        new element ids. Returns {'counts': {...}, 'errors': int, 'skipped': {...}}.
        """
        if default_wall_height_ft is None:
            default_wall_height_ft = mm_to_ft(3000.0)

        counts = {k: 0 for k in ['Wall', 'Floor', 'Ceiling', 'Column',
                                  'Door', 'Window', 'Stair', 'Roof']}
        skipped       = {'Stair': 0}
        errors        = 0
        created_walls = {}   # id(detected wall _data dict) → Revit Wall element
        wall_order = []      # creation order: IronPython 2.7 dicts are unordered

        def made(element):
            if element and created is not None:
                created.append(_eid(element))
            return element

        with disposing(TransactionGroup(self.doc, "T3Lab: Point Cloud to Model")) as tg:
            tg.Start()
            # Pass 1: walls
            for elem in elements:
                if not elem.Include:
                    continue
                if elem.Type == 'Wall':
                    w = made(self.build_wall(elem, default_wall_height_ft))
                    if w:
                        counts['Wall'] += 1
                        created_walls[id(elem._data)] = w
                        wall_order.append(w)
                    else:
                        errors += 1

            # Pass 2: floors, ceilings, columns, roof
            for elem in elements:
                if not elem.Include:
                    continue
                if elem.Type == 'Floor':
                    r = made(self.build_floor(elem))
                    if r:   counts['Floor'] += 1
                    else:   errors += 1
                elif elem.Type == 'Ceiling':
                    r = made(self.build_ceiling(elem))
                    if r:   counts['Ceiling'] += 1
                    else:   errors += 1
                elif elem.Type == 'Column':
                    r = made(self.build_column(elem))
                    if r:   counts['Column'] += 1
                    else:   errors += 1
                elif elem.Type == 'Roof':
                    r = made(self.build_roof(elem))
                    if r:   counts['Roof'] += 1
                    else:   errors += 1
                elif elem.Type == 'Stair':
                    # build_stair_marker is a stub — creation needs the
                    # Architecture stairs API. Counting it as created reported
                    # "N Stairs" in the summary for elements that were never
                    # placed; track it separately and tell the truth.
                    skipped['Stair'] += 1

            # Pass 3: doors / windows — host into the wall each opening was
            # detected in; fall back to any created wall
            fallback_wall = wall_order[0] if wall_order else None
            for elem in elements:
                if not elem.Include:
                    continue
                if elem.Type not in ('Door', 'Window'):
                    continue
                host = created_walls.get(id(elem._data['host_wall_data']),
                                         fallback_wall)
                r = made(self.build_opening(elem, host))
                if r:   counts[elem.Type] += 1
                else:   errors += 1

            tg.Assimilate()

        return {'counts': counts, 'errors': errors, 'skipped': skipped}


# ── Detections in, as detect_point_cloud_elements returns them (mm) ──────────

KIND_KEYS = ('walls', 'floors', 'ceilings', 'doors', 'windows', 'columns', 'stairs', 'roof')
TYPE_KEYS = {'Wall': 'walls', 'Floor': 'floors', 'Ceiling': 'ceilings', 'Door': 'doors',
             'Window': 'windows', 'Column': 'columns', 'Stair': 'stairs', 'Roof': 'roof'}
_HOST_TOL_FT = 0.2                  # ~60 mm: an opening's host wall centre vs a wall row
_HOST_TOL_DEG = 1.0


def _rows(detections):
    """Rows of a detect_point_cloud_elements result, of its `detections`, or a list."""
    if isinstance(detections, dict):
        inner = detections.get('detections', detections)
        if isinstance(inner, dict):
            rows = []
            for key in KIND_KEYS:
                rows += list(inner.get(key) or [])
            return rows
        detections = inner
    return list(detections or [])


def _ft(value):
    return mm_to_ft(float(value))


def _xy(pair):
    return (_ft(pair[0]), _ft(pair[1]))


def _host_for(row, walls):
    """(host data dict, u along it) for a door / window row: the wall row it was
    detected in when that wall is in the same build, else a free-standing host."""
    loc = row['location_mm']
    x, y = _ft(loc[0]), _ft(loc[1])
    cx, cy = _xy(row['host_wall_center_mm'])
    deg = float(row['host_wall_angle_deg'])
    host = None
    for w in walls:
        d = w._data
        turn = abs(math.degrees(d['angle']) - deg) % 180.0
        if (math.hypot(d['cx'] - cx, d['cy'] - cy) <= _HOST_TOL_FT
                and min(turn, 180.0 - turn) <= _HOST_TOL_DEG):
            host = d
            break
    if host is None:
        host = {'angle': math.radians(deg), 'cx': cx, 'cy': cy}
    u = (x - host['cx']) * math.cos(host['angle']) + (y - host['cy']) * math.sin(host['angle'])
    return host, u


def _element(row, walls):
    kind = row.get('type')
    level = row.get('level') or ''
    if kind == 'Wall':
        (x0, y0), (x1, y1) = _xy(row['start_mm']), _xy(row['end_mm'])
        length = math.hypot(x1 - x0, y1 - y0)
        if length <= 0.01:
            raise ValueError('start_mm equals end_mm')
        return DetectedWall(math.atan2(y1 - y0, x1 - x0), (x0 + x1) / 2.0, (y0 + y1) / 2.0,
                            length, _ft(row['thickness_mm']), level, True, '',
                            _ft(row.get('base_z_mm') or 0.0))
    if kind in ('Floor', 'Ceiling'):
        return DetectedFloor(_ft(row['z_mm']), [_xy(c) for c in row['corners_mm']],
                             kind.lower(), level)
    if kind == 'Column':
        cx, cy = _xy(row['center_mm'])
        return DetectedColumn(cx, cy, _ft(row['width_mm']), _ft(row['depth_mm']),
                              _ft(row['bottom_z_mm']), _ft(row['top_z_mm']), level)
    if kind in ('Door', 'Window'):
        host, u = _host_for(row, walls)
        return DetectedOpening(kind, host, u, _ft(row['location_mm'][2]),
                               _ft(row['width_mm']), _ft(row['height_mm']), level)
    if kind == 'Stair':
        return DetectedStair(_ft(row['bottom_z_mm']), _ft(row['top_z_mm']),
                             int(row.get('tread_count') or 0), level)
    if kind == 'Roof':
        return DetectedRoof(_ft(row['z_mm']), [_xy(c) for c in row['corners_mm']],
                            float(row.get('slope_deg') or 0.0), level)
    return None


def from_detections(detections):
    """(DetectedElement list, problems) from detect_point_cloud_elements rows in mm.
    Walls are read first so each door / window is hosted in the wall it was found in."""
    rows = [r for r in _rows(detections) if isinstance(r, dict)]
    rows.sort(key=lambda r: 0 if r.get('type') == 'Wall' else 1)
    walls, out, problems = [], [], []
    for i, row in enumerate(rows):
        try:
            el = _element(row, walls)
        except (KeyError, TypeError, ValueError, IndexError) as ex:
            problems.append('{} #{}: missing or bad {}'.format(row.get('type'), i + 1, ex))
            continue
        if el is None:
            problems.append('#{}: unknown type "{}"'.format(i + 1, row.get('type')))
            continue
        if row.get('confidence_pct') is not None:
            el.ConfidenceText = u'{}%'.format(row['confidence_pct'])
        out.append(el)
        if el.Type == 'Wall':
            walls.append(el)
    return out, problems


def confidence(element):
    try:
        return int('{}'.format(element.ConfidenceText).rstrip('%'))
    except ValueError:
        return 0


# ── Detection again, with detect_point_cloud_elements' arguments ─────────────
# One choice of cloud and region: core/server.py _pick_point_cloud /
# _point_cloud_region call these.

def pick_cloud(clouds, wanted_id, id_of):
    """(cloud, None) or (None, error dict)."""
    rows = [{'id': id_of(c), 'name': c.Name} for c in clouds]
    if not clouds:
        return None, {'error': 'There is no point cloud in the active document. '
                               'Link one with Insert > Point Cloud, then retry.'}
    if wanted_id is None:
        if len(clouds) == 1:
            return clouds[0], None
        return None, {'error': 'This document has {} point clouds: pass point_cloud_id.'
                               .format(len(clouds)), 'point_clouds': rows}
    for c in clouds:
        if id_of(c) == int(wanted_id):
            return c, None
    return None, {'error': 'No point cloud with id {}.'.format(wanted_id), 'point_clouds': rows}


def region_from_args(args, levels, bbox):
    """((x0, y0, z0), (x1, y1, z1)) in feet, None (whole cloud), or an error dict.
    levels: [(name, elevation ft)]; bbox: the cloud's ((min), (max)) or None."""
    region = args.get('region_mm')
    if region:
        try:
            lo = [float(v) / 304.8 for v in region['min']]
            hi = [float(v) / 304.8 for v in region['max']]
            if len(lo) != 3 or len(hi) != 3:
                raise ValueError
        except (KeyError, TypeError, ValueError):
            return {'error': 'region_mm must be {"min": [x, y, z], "max": [x, y, z]} in mm.'}
        return (tuple(min(a, b) for a, b in zip(lo, hi)),
                tuple(max(a, b) for a, b in zip(lo, hi)))
    name = args.get('level_name')
    if not name:
        return None
    ordered = sorted(levels, key=lambda lv: lv[1])
    idx = next((i for i, lv in enumerate(ordered)
                if lv[0].strip().lower() == '{}'.format(name).strip().lower()), None)
    if idx is None:
        return {'error': 'No Level named "{}".'.format(name), 'levels': [lv[0] for lv in ordered]}
    z0 = ordered[idx][1]
    z1 = ordered[idx + 1][1] if idx + 1 < len(ordered) else z0 + 4000.0 / 304.8
    if bbox is None:
        return {'error': 'The point cloud has no bounding box (is it loaded?).'}
    return ((bbox[0][0], bbox[0][1], z0), (bbox[1][0], bbox[1][1], z1))


def detect(doc, args):
    """Sample and detect like detect_point_cloud_elements: (DetectedElement list,
    source dict) or (None, error dict). Revit main thread only (GetPoints)."""
    from Autodesk.Revit.DB import PointCloudInstance
    from Snippets._compat import eid_value
    from Services import point_cloud_analysis as pca

    clouds = list(FilteredElementCollector(doc).OfClass(PointCloudInstance).ToElements())
    pc, err = pick_cloud(clouds, args.get('point_cloud_id'), lambda c: eid_value(c.Id))
    if err:
        return None, err
    levels = []
    for lv in FilteredElementCollector(doc).OfClass(Level).ToElements():
        try:
            elev = lv.ProjectElevation    # internal origin, like the points
        except Exception:
            elev = lv.Elevation
        levels.append((lv.Name, elev))
    bb = pc.get_BoundingBox(None)
    bbox = None if bb is None else ((bb.Min.X, bb.Min.Y, bb.Min.Z), (bb.Max.X, bb.Max.Y, bb.Max.Z))
    region = region_from_args(args, levels, bbox)
    if isinstance(region, dict):
        return None, region
    budget = pca.clamp_point_budget(args.get('max_points'))
    if region is None:
        pts = pca.extract_full_cloud(pc, budget)
    else:
        (x0, y0, z0), (x1, y1, z1) = region
        center = XYZ((x0 + x1) / 2.0, (y0 + y1) / 2.0, (z0 + z1) / 2.0)
        pts = pca.extract_full_cloud_from_region(
            pc, center, abs(x1 - x0) / 2.0, abs(y1 - y0) / 2.0, abs(z1 - z0) / 2.0, budget)
    source = {'point_cloud_id': eid_value(pc.Id), 'name': pc.Name, 'sampled_points': len(pts or [])}
    if not pts:
        return None, {'error': 'No points came back from the point cloud. Check that it is '
                               'loaded and visible, and that region_mm / level_name overlaps it.',
                      'source': source}
    elements = args.get('elements') or list(pca.DETECTABLE)
    results = pca.PointCloudAnalyzer(pts, doc=doc).run(
        pca.detection_settings(elements, args.get('min_wall_length_mm')))
    wanted = set(elements)
    return [r for r in results if TYPE_KEYS.get(r.Type) in wanted], source
