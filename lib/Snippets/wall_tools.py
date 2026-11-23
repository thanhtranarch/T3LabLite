# -*- coding: utf-8 -*-
"""Wall Adjust Base (re-host to a level, same 3D position) and Wall Cut Profile
(openings where ducts, pipes, trays, beams or columns cross walls).

Revit side of GUI/WallAdjustBaseDialog.py, GUI/WallCutProfileDialog.py and the
rehost_to_level / cut_wall_openings tools (core/tools/model.py). Revit is
imported inside functions (IronPython 2.7 server, tests without Revit).
Test: dev/test_tools_model_domain.py.
"""
from Snippets._compat import clr_namespace, eid_value

_SWALLOWER = None


def silence_warnings(transaction):
    """Dismiss warnings in `transaction` (preprocessor class built once, lazily)."""
    global _SWALLOWER
    if _SWALLOWER is None:
        from Autodesk.Revit.DB import (IFailuresPreprocessor, FailureProcessingResult,
                                       FailureSeverity)

        class WallToolWarnings(IFailuresPreprocessor):
            __namespace__ = clr_namespace('WallTools')

            def PreprocessFailures(self, fa):
                for f in fa.GetFailureMessages():
                    if f.GetSeverity() == FailureSeverity.Warning:
                        fa.DeleteWarning(f)
                return FailureProcessingResult.Continue

        _SWALLOWER = WallToolWarnings
    opts = transaction.GetFailureHandlingOptions()
    opts.SetFailuresPreprocessor(_SWALLOWER())
    transaction.SetFailureHandlingOptions(opts)


def bic_int(name):
    """int of BuiltInCategory.<name>, None when this release lacks it."""
    from Autodesk.Revit.DB import BuiltInCategory
    member = getattr(BuiltInCategory, name, None)
    return int(member) if member is not None else None


def category_int(element):
    try:
        return eid_value(element.Category.Id)
    except Exception:
        return None


# ── Wall Adjust Base ────────────────────────────────────────────────────────

def level_params(elem, adjust_base=True):
    """(level BuiltInParameter, offset BuiltInParameter) that place `elem`, or None."""
    from Autodesk.Revit.DB import (BuiltInCategory, BuiltInParameter, FamilyInstance,
                                   Floor, Wall)
    if isinstance(elem, Wall):
        return ((BuiltInParameter.WALL_BASE_CONSTRAINT if adjust_base else BuiltInParameter.WALL_HEIGHT_TYPE),
                (BuiltInParameter.WALL_BASE_OFFSET if adjust_base else BuiltInParameter.WALL_TOP_OFFSET))
    if isinstance(elem, Floor):
        return BuiltInParameter.LEVEL_PARAM, BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM
    if isinstance(elem, FamilyInstance) and elem.Category:
        cat_id = eid_value(elem.Category.Id)
        if cat_id in (BuiltInCategory.OST_StructuralColumns.value__, BuiltInCategory.OST_Columns.value__):
            return ((BuiltInParameter.FAMILY_BASE_LEVEL_PARAM if adjust_base else BuiltInParameter.FAMILY_TOP_LEVEL_PARAM),
                    (BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM if adjust_base
                     else BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM))
        if cat_id == BuiltInCategory.OST_StructuralFraming.value__:
            return BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM, BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION
    return None


def plan_rehost(doc, elem, new_level, adjust_base=True):
    """{'p_lvl', 'p_off', 'old_level', 'old_offset', 'new_offset'} for moving
    `elem`'s base (or top) to `new_level` at the same elevation, or
    {'reason': ...} when it cannot be done. Reads only."""
    from Autodesk.Revit.DB import Level
    params = level_params(elem, adjust_base)
    if params is None:
        return {'reason': 'not a wall, floor, column or beam'}
    p_lvl = elem.get_Parameter(params[0])
    p_off = elem.get_Parameter(params[1])
    if not p_lvl or not p_off or p_lvl.IsReadOnly or p_off.IsReadOnly:
        return {'reason': 'level or offset parameter missing or read-only'}
    old_lvl_id = p_lvl.AsElementId()
    if eid_value(old_lvl_id) < 0:
        return {'reason': 'top is unconnected' if not adjust_base else 'no level'}
    old_level = doc.GetElement(old_lvl_id)
    if not old_level or not isinstance(old_level, Level):
        return {'reason': 'no level'}
    old_offset = p_off.AsDouble()
    actual_elevation = old_level.Elevation + old_offset
    return {'p_lvl': p_lvl, 'p_off': p_off, 'old_level': old_level, 'old_offset': old_offset,
            'new_offset': actual_elevation - new_level.Elevation}


def _beam_end1(elem, plan):
    """(END1 offset parameter, its elevation) of a beam being rehosted, else None."""
    from Autodesk.Revit.DB import BuiltInCategory, FamilyInstance, BuiltInParameter
    try:
        if not (isinstance(elem, FamilyInstance) and elem.Category is not None and
                eid_value(elem.Category.Id) == BuiltInCategory.OST_StructuralFraming.value__):
            return None
        p = elem.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION)
        if p is None or p.IsReadOnly:
            return None
        return p, plan['old_level'].Elevation + p.AsDouble()
    except Exception:
        return None


def adjust_element(doc, elem, new_level, adjust_base=True):
    """Re-host elem's base (or top) on new_level keeping its elevation; True when set.
    Call inside an open Transaction."""
    try:
        plan = plan_rehost(doc, elem, new_level, adjust_base)
        if 'reason' in plan:
            return False
        end1 = _beam_end1(elem, plan)
        plan['p_lvl'].Set(new_level.Id)
        plan['p_off'].Set(plan['new_offset'])
        if end1 is not None:
            # A beam has two end offsets from its reference level: keep the end too.
            end1[0].Set(end1[1] - new_level.Elevation)
        return True
    except Exception:
        return False


# ── Wall Cut Profile ────────────────────────────────────────────────────────

# Tool key -> (label in the window, BuiltInCategory name).
CUT_CATEGORIES = (
    ('ducts', 'Ducts', 'OST_DuctCurves'),
    ('pipes', 'Pipes', 'OST_PipeCurves'),
    ('cable_trays', 'Cable Trays', 'OST_CableTray'),
    ('structural_framing', 'Structural Framing', 'OST_StructuralFraming'),
    ('columns', 'Columns', 'OST_Columns'),
)
METHOD_FAMILY = "Place Opening Family"
METHOD_OPENING = "Wall Opening"


def categories(keys):
    """[(label, BuiltInCategory)] for tool keys, in CUT_CATEGORIES order."""
    from Autodesk.Revit.DB import BuiltInCategory
    keys = set(keys)
    return [(label, getattr(BuiltInCategory, bic)) for key, label, bic in CUT_CATEGORIES
            if key in keys]


def is_cut_wall(elem):
    """Basic or curtain-free wall with a location curve (the window's wall filter)."""
    from Autodesk.Revit.DB import Wall, WallKind
    if not isinstance(elem, Wall):
        return False
    if elem.WallType and elem.WallType.Kind in (WallKind.Curtain, WallKind.Stacked):
        return False
    loc = elem.Location
    return loc is not None and hasattr(loc, 'Curve') and loc.Curve is not None


def opening_symbols(doc):
    """Wall-hosted or opening-like family types (Generic Models, Windows, Doors), unique."""
    from Autodesk.Revit.DB import (BuiltInCategory, BuiltInParameter,
                                   FilteredElementCollector, FamilySymbol)
    results = []
    cats = [
        BuiltInCategory.OST_GenericModel,
        BuiltInCategory.OST_Windows,
        BuiltInCategory.OST_Doors,
    ]
    for bic in cats:
        try:
            coll = FilteredElementCollector(doc).OfCategory(bic).OfClass(FamilySymbol)
            for sym in coll:
                fam = sym.Family
                if not fam:
                    continue
                p = fam.get_Parameter(BuiltInParameter.FAMILY_HOSTING_BEHAVIOR)
                if p and p.AsInteger() == 1:
                    results.append(sym)
                elif any(kw in fam.Name.lower() for kw in ["opening", "void", "cutout", "penetration", "sleeve"]):
                    results.append(sym)
        except Exception:
            pass
    seen, unique = set(), []
    for sym in results:
        sid = eid_value(sym.Id)
        if sid not in seen:
            seen.add(sid)
            unique.append(sym)
    return unique


def symbol_label(symbol):
    """'Family : Type' as the window lists opening families."""
    from Autodesk.Revit.DB import BuiltInParameter
    try:
        fam_name = symbol.Family.Name
    except Exception:
        fam_name = "Unknown"
    try:
        p = symbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME)
        type_name = p.AsString() if p else symbol.Name
    except Exception:
        type_name = symbol.Name
    return "{} : {}".format(fam_name, type_name)


def collect_cutters(sources, categories):
    """Bounding boxes of the cutting elements, in this model's coordinates.
    sources: [(document, transform or None)]; categories: [(label, BuiltInCategory)]."""
    from Autodesk.Revit.DB import FilteredElementCollector
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
    return cutters


def find_hits(walls, cutters, offset_ft):
    """[(wall, cutter)] where the cutter's centre lies in the wall's box grown by offset_ft."""
    hits = []
    for w in walls:
        w_bb = w.get_BoundingBox(None)
        if not w_bb:
            continue
        w_min = w_bb.Min
        w_max = w_bb.Max
        for item in cutters:
            i_c = item['center']
            if (w_min.X - offset_ft <= i_c.X <= w_max.X + offset_ft and
                    w_min.Y - offset_ft <= i_c.Y <= w_max.Y + offset_ft and
                    w_min.Z - offset_ft <= i_c.Z <= w_max.Z + offset_ft):
                hits.append((w, item))
    return hits


def execute_openings(doc, walls, sources, categories, method, offset_ft, symbol, created_ids=None):
    """Cut walls where cutting elements cross them; one Transaction.
    Returns (created, errors); created_ids (list) receives the new element ids."""
    from Autodesk.Revit.DB import Transaction, XYZ
    from Autodesk.Revit.DB.Structure import StructuralType
    created = 0
    errors = 0
    hits = find_hits(walls, collect_cutters(sources, categories), offset_ft)
    new_ids = []

    t = Transaction(doc, "T3Lab: Wall Cut Profile")
    t.Start()
    silence_warnings(t)            # after Start, as every other T3Lab transaction
    try:
        if symbol and not symbol.IsActive:
            symbol.Activate()
            doc.Regenerate()

        for w, item in hits:
            i_c = item['center']
            try:
                if symbol and METHOD_FAMILY in method:
                    inst = doc.Create.NewFamilyInstance(i_c, symbol, w, StructuralType.NonStructural)
                    for p_name, val in [("Width", max(item['dx'], item['dy']) + offset_ft * 2),
                                        ("Height", item['dz'] + offset_ft * 2)]:
                        p = inst.LookupParameter(p_name)
                        if p and not p.IsReadOnly:
                            p.Set(val)
                    new_ids.append(eid_value(inst.Id))
                    created += 1
                elif METHOD_OPENING in method:
                    p0 = XYZ(i_c.X - 0.5, i_c.Y - 0.5, i_c.Z - item['dz'] * 0.5 - offset_ft)
                    p1 = XYZ(i_c.X + 0.5, i_c.Y + 0.5, i_c.Z + item['dz'] * 0.5 + offset_ft)
                    opening = doc.Create.NewOpening(w, p0, p1)
                    if opening is not None:
                        new_ids.append(eid_value(opening.Id))
                        created += 1
                    else:
                        errors += 1
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

    if created_ids is not None:
        created_ids.extend(new_ids)
    return created, errors
