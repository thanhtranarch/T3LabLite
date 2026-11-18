# -*- coding: utf-8 -*-
"""Build a Revit family from a FamiGen schema (v1 or v2). Revit API only, no WPF.

Shared by the FamiGen dialog (Create Family) and the MCP server
(``famigen_create_family``), so both produce the same family from the same JSON.

Every public entry point must run in Revit API context (a modal command, an
ExternalEvent handler). Each user action is one Transaction in the family
document, rolled back on any error (rules S1/S2/S19).

Geometry parsing (projection onto the sketch plane, endpoint healing, blend
loop alignment, cylinder axis handling) moved here unchanged from
``FamiGenDialog._generate_json_family``. v2 adds, inside the same transaction:

* materials - one Revit ``Material`` per ``materials[]`` entry (colour,
  transparency, shininess, smoothness), one family type parameter per material
  in Materials and Finishes whose default is that material, and every solid's
  ``MATERIAL_ID_PARAM`` associated to its parameter;
* ``subcategory`` - created under the family category and assigned to the form;
* ``parameters[]`` - set when the template has them, created when it does not;
* ``repeat`` - written out first (``expand_schema``), exactly as the preview does;
* ``detail_levels`` - each form's FamilyElementVisibility (coarse/medium/fine);
* ``reference_planes`` / ``dimensions`` / ``geometry[].locks`` - planes, length
  dimensions labelled with their parameter, and solid faces aligned + locked to
  the planes, so the family flexes. A constraint that Revit refuses is reported
  as a warning and never stops the family from being built;
* ``types[]`` - the current type is renamed to the first type, the others are
  created with ``FamilyManager.NewType`` (which copies the current values) and
  given their parameter and material values.
"""
import logging
import math
import os
import re

from Autodesk.Revit.DB import (
    Arc, BuiltInCategory, BuiltInParameter, Category, Color, CurveArrArray,
    CurveArray, Ellipse, FailureProcessingResult, FailureSeverity,
    FamilyElementVisibility, FamilyElementVisibilityType,
    FilteredElementCollector, HermiteSpline, IFailuresPreprocessor,
    IFamilyLoadOptions, FamilySource, Level, Line, Material, Options, Plane,
    PlanarFace, ProfilePlaneLocation, ReferenceArray, ReferencePlane, SaveAsOptions,
    SketchPlane, Solid, Transaction, View, ViewType, XYZ,
)
from System.Collections.Generic import List as _NetList

from Intelligence.family_schema import (
    BUILTIN_PLANES, DETAIL_LEVELS, category_bic_name, category_is_hosted, category_templates,
    expand_schema, material_parameter_name, parse_color, validate_family_schema,
    SUPPORTED_CATEGORIES,
)
from Snippets._compat import add_family_parameter, disposing, family_parameter_kind
from FamilyGen.form_frames import local_to_world, path_ends, profile_at

logger = logging.getLogger('T3Lab.FamilyGen')

SCL = 1.0 / 304.8   # mm -> feet
FULL_TURN = 6.283185307


class FamilyBuildError(Exception):
    """A family could not be created; the message is shown to the user."""


class _WarningSwallower(IFailuresPreprocessor):
    __namespace__ = "T3Lab.FamilyBuilder"

    def PreprocessFailures(self, failuresAccessor):
        for failure in failuresAccessor.GetFailureMessages():
            if failure.GetSeverity() == FailureSeverity.Warning:
                failuresAccessor.DeleteWarning(failure)
        return FailureProcessingResult.Continue


class _OverwriteLoadOptions(IFamilyLoadOptions):
    """Reload over an already loaded family, overwriting its parameter values."""
    __namespace__ = "T3Lab.FamilyBuilder"

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        try:                                   # IronPython: out param is a StrongBox
            overwriteParameterValues.Value = True
            return True
        except AttributeError:                 # pythonnet 3: out values in the tuple
            return (True, True)

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        try:
            source.Value = FamilySource.Family
            overwriteParameterValues.Value = True
            return True
        except AttributeError:
            return (True, FamilySource.Family, True)


def _start(transaction):
    options = transaction.GetFailureHandlingOptions()
    options.SetFailuresPreprocessor(_WarningSwallower())
    transaction.SetFailureHandlingOptions(options)
    transaction.Start()


# ── templates ────────────────────────────────────────────────────────────────

def template_search_dirs(app):
    dirs = []
    try:
        tdir = app.FamilyTemplatePath
        if tdir and os.path.isdir(tdir):
            dirs.append(tdir)
    except Exception:
        pass
    base = r"C:\ProgramData\Autodesk\RVT {}".format(app.VersionNumber)
    for sub in ("English", "", "English-Imperial", "English_I"):
        dirs.append(os.path.join(base, "Family Templates", sub) if sub
                    else os.path.join(base, "Family Templates"))
    return dirs


def find_template(app, category):
    """Full path of the category's .rft, or None."""
    names = category_templates(category)
    for folder in template_search_dirs(app):
        if not os.path.isdir(folder):
            continue
        for name in names:
            path = os.path.join(folder, name)
            if os.path.isfile(path):
                return path
    return None


def resolve_template(app, category):
    """(template_path, recategorize_to) - never a silent category change.

    A missing non-hosted template is replaced by the Generic Model template and
    the family is re-categorised to `category` (reported as a warning). Hosted
    categories (Door, Window) need their own template - the host cut lives in
    it - so a missing one is an error.
    """
    if category not in SUPPORTED_CATEGORIES:
        raise FamilyBuildError("Category '{}' is not supported. Use one of: {}."
                               .format(category, ', '.join(SUPPORTED_CATEGORIES)))
    path = find_template(app, category)
    if path:
        return path, None
    searched = ', '.join(category_templates(category))
    if category_is_hosted(category):
        raise FamilyBuildError(
            "The '{}' family template ({}) was not found in the Revit family template "
            "folders. Install the Revit content for this language, or set the Family "
            "Template path in Revit Options > File Locations, then try again."
            .format(category, searched))
    generic = find_template(app, 'Generic Model')
    if not generic:
        raise FamilyBuildError(
            "No family template was found for '{}' ({}) or Generic Model. Set the Family "
            "Template path in Revit Options > File Locations, then try again."
            .format(category, searched))
    return generic, category


# ── geometry ─────────────────────────────────────────────────────────────────

class JsonFamilyBuilder(object):
    """Turns schema geometry into family forms in an open family document.

    The caller owns the Transaction; `build()` only creates elements.
    """

    _PLANE_AXES = {
        'z': (XYZ.BasisX, XYZ.BasisY),
        'x': (XYZ.BasisY, XYZ.BasisZ),   # arcs on an X-facing plane: 0 rad = +Y
        'y': (XYZ.BasisZ, XYZ.BasisX),   # arcs on a Y-facing plane: 0 rad = +Z
    }
    _SNAP_TOL = 2.0 * SCL   # snap endpoint gaps under ~2 mm
    _MIN_LEN = 1.0 * SCL    # drop segments under ~1 mm (Revit short-curve tol ~0.8 mm)

    def __init__(self, fam_doc):
        self.doc = fam_doc
        self.fm = fam_doc.FamilyManager
        self.report = {
            'built': 0, 'total': 0, 'skipped': [], 'warnings': [],
            'materials_created': [], 'materials_reused': [], 'material_parameters': [],
            'parameters_created': [], 'parameters_set': [], 'subcategories': [],
            'types': [], 'reference_planes': [], 'dimensions': [], 'locks': 0,
        }
        self._forms = {}            # geometry index -> created forms
        self._materials = {}        # schema material name -> (Material, FamilyParameter|None)
        self._subcategories = {}

    # ── family type / parameters ────────────────────────────────────────
    def _ensure_current_type(self, name):
        if self.fm.CurrentType is None:
            self.fm.NewType(name or 'Default')

    def _family_params(self):
        return dict((p.Definition.Name, p) for p in self.fm.Parameters)

    def _set_value(self, fparam, kind, value):
        if kind == 'length':
            self.fm.Set(fparam, float(value) * SCL)
        elif kind == 'number':
            self.fm.Set(fparam, float(value))
        elif kind == 'integer':
            self.fm.Set(fparam, int(value))
        elif kind == 'text':
            self.fm.Set(fparam, u'{}'.format(value))
        else:
            raise ValueError('this parameter type cannot be set from JSON')

    def apply_parameters(self, params):
        existing = self._family_params()
        groups = {'length': 'geometry', 'number': 'data', 'integer': 'data', 'text': 'text'}
        for spec in params or []:
            if not isinstance(spec, dict):
                continue
            name = (spec.get('name') or '').strip()
            if not name:
                continue
            ptype = spec.get('type', 'length')
            fparam = existing.get(name)
            if fparam is None:
                try:
                    fparam = add_family_parameter(self.fm, name, groups.get(ptype, 'data'),
                                                  ptype, bool(spec.get('instance', False)))
                    existing[name] = fparam
                    self.report['parameters_created'].append(name)
                except Exception as ex:
                    self.report['warnings'].append(
                        "Parameter '{}' could not be created: {}".format(name, ex))
                    continue
            value = spec.get('value')
            if value is None:
                continue
            # The template's own parameter type wins over the JSON `type`:
            # a "Width" length parameter still takes millimeters.
            kind = family_parameter_kind(fparam.Definition, getattr(fparam, 'StorageType', None))
            try:
                if kind in ('length', 'number', 'integer') and isinstance(value, str):
                    value = float(value)
                self._set_value(fparam, kind, value)
                self.report['parameters_set'].append(name)
            except Exception as ex:
                self.report['warnings'].append(
                    "Parameter '{}' was not set ({} parameter, value {!r}): {}"
                    .format(name, kind, value, ex))

    # ── family types ────────────────────────────────────────────────────
    def _type_by_name(self, name):
        for ftype in self.fm.Types:
            if ftype.Name == name:
                return ftype
        return None

    def _set_type_value(self, type_name, params, key, value):
        fparam = params.get(key.strip().lower())
        if fparam is None:
            self.report['warnings'].append(
                "Type '{}': parameter '{}' is not in the family; value skipped"
                .format(type_name, key))
            return
        kind = family_parameter_kind(fparam.Definition, getattr(fparam, 'StorageType', None))
        try:
            if kind == 'material':
                entry = self._materials.get((u'{}'.format(value)).strip())
                if entry is None:
                    raise ValueError("material '{}' is not defined".format(value))
                self.fm.Set(fparam, entry[0].Id)
                return
            if kind in ('length', 'number', 'integer') and isinstance(value, str):
                value = float(value)
            self._set_value(fparam, kind, value)
        except Exception as ex:
            self.report['warnings'].append(
                "Type '{}': '{}' not set ({} parameter, value {!r}): {}"
                .format(type_name, key, kind, value, ex))

    def apply_types(self, types):
        """First type = the current one (renamed); the rest via NewType."""
        specs = [t for t in (types or []) if isinstance(t, dict)
                 and isinstance(t.get('name'), str) and t['name'].strip()]
        if not specs:
            return
        params = dict((p.Definition.Name.strip().lower(), p) for p in self.fm.Parameters)
        first = None
        for index, spec in enumerate(specs):
            name = spec['name'].strip()
            try:
                ftype = self._type_by_name(name)
                if ftype is not None:
                    self.fm.CurrentType = ftype
                elif index == 0 and self.fm.CurrentType is not None:
                    self.fm.RenameCurrentType(name)
                    ftype = self.fm.CurrentType
                else:
                    ftype = self.fm.NewType(name)
            except Exception as ex:
                self.report['warnings'].append("Type '{}' could not be created: {}".format(name, ex))
                continue
            if first is None:
                first = ftype
            for key, value in sorted((spec.get('values') or {}).items()):
                if isinstance(key, str) and key.strip():
                    self._set_type_value(name, params, key, value)
            self.report['types'].append(name)
        if first is not None:
            try:
                self.fm.CurrentType = first
            except Exception:
                pass

    # ── materials ───────────────────────────────────────────────────────
    def _material_by_name(self, name):
        with disposing(FilteredElementCollector(self.doc)) as collector:
            for mat in collector.OfClass(Material):
                if mat.Name == name:
                    return mat
        return None

    def apply_materials(self, materials):
        existing = self._family_params()
        for spec in materials or []:
            if not isinstance(spec, dict):
                continue
            name = (spec.get('name') or '').strip()
            rgb = parse_color(spec.get('color'))
            if not name or rgb is None:
                continue
            material = self._material_by_name(name)
            if material is None:
                material = self.doc.GetElement(Material.Create(self.doc, name))
                self.report['materials_created'].append(name)
            else:
                self.report['materials_reused'].append(name)
            material.Color = Color(rgb[0], rgb[1], rgb[2])
            for field, attr, top in (('transparency', 'Transparency', 100),
                                     ('shininess', 'Shininess', 128),
                                     ('smoothness', 'Smoothness', 100)):
                if spec.get(field) is not None:
                    try:
                        setattr(material, attr, int(round(max(0, min(top, float(spec[field]))))))
                    except Exception as ex:
                        self.report['warnings'].append(
                            "Material '{}': {} not set ({})".format(name, field, ex))
            pname = material_parameter_name(spec)
            fparam = existing.get(pname)
            if fparam is None:
                try:
                    fparam = add_family_parameter(self.fm, pname, 'materials', 'material', False)
                    existing[pname] = fparam
                    self.report['material_parameters'].append(pname)
                except Exception as ex:
                    fparam = None
                    self.report['warnings'].append(
                        "Material parameter '{}' could not be created: {}. Parts keep the "
                        "material directly.".format(pname, ex))
            elif family_parameter_kind(fparam.Definition, getattr(fparam, 'StorageType', None)) != 'material':
                self.report['warnings'].append(
                    "Family parameter '{}' exists but is not a Material parameter; parts "
                    "using '{}' keep the material directly.".format(pname, name))
                fparam = None
            else:
                self.report['material_parameters'].append(pname)
            if fparam is not None:
                try:
                    self.fm.Set(fparam, material.Id)
                except Exception as ex:
                    self.report['warnings'].append(
                        "Default of '{}' not set: {}".format(pname, ex))
            self._materials[name] = (material, fparam)

    def _bind_material(self, form, material_name, label):
        entry = self._materials.get(material_name)
        if entry is None:
            self.report['warnings'].append(
                "{}: material '{}' is not defined; it keeps the category default"
                .format(label, material_name))
            return
        material, fparam = entry
        param = form.get_Parameter(BuiltInParameter.MATERIAL_ID_PARAM)
        if param is None:
            return
        try:
            if fparam is not None:
                self.fm.AssociateElementParameterToFamilyParameter(param, fparam)
            else:
                param.Set(material.Id)
        except Exception as ex:
            self.report['warnings'].append("{}: material not applied ({})".format(label, ex))

    def _subcategory(self, name):
        if name in self._subcategories:
            return self._subcategories[name]
        family_category = self.doc.OwnerFamily.FamilyCategory
        found = None
        for sub in family_category.SubCategories:
            if sub.Name == name:
                found = sub
                break
        if found is None:
            found = self.doc.Settings.Categories.NewSubcategory(family_category, name)
            self.report['subcategories'].append(name)
        self._subcategories[name] = found
        return found

    # ── curve parsing (moved from FamiGenDialog) ────────────────────────
    def _plane_info(self, geom_data):
        """Which plane the entry sketches on: ('x'|'y'|'z', value in feet)."""
        if "sketch_plane_x" in geom_data:
            return ('x', geom_data["sketch_plane_x"] * SCL)
        if "sketch_plane_y" in geom_data:
            return ('y', geom_data["sketch_plane_y"] * SCL)
        return ('z', geom_data.get("sketch_plane_z", 0.0) * SCL)

    def _make_sketch_plane(self, plane):
        kind, val = plane
        if kind == 'x':
            base_plane = Plane.CreateByNormalAndOrigin(XYZ.BasisX, XYZ(val, 0, 0))
        elif kind == 'y':
            base_plane = Plane.CreateByNormalAndOrigin(XYZ.BasisY, XYZ(0, val, 0))
        else:
            base_plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ(0, 0, val))
        return SketchPlane.Create(self.doc, base_plane)

    def _json_point(self, coords, plane):
        """mm triple -> XYZ (feet), projected onto the sketch plane if given."""
        x, y, z = coords[0] * SCL, coords[1] * SCL, coords[2] * SCL
        if plane is not None:
            kind, val = plane
            if kind == 'x':
                x = val
            elif kind == 'y':
                y = val
            else:
                z = val
        return XYZ(x, y, z)

    def _json_curve(self, seg, plane, split_full=0):
        """One JSON segment -> Revit Curve (None = unusable).

        With ``split_full`` >= 2 a full circle/ellipse comes back as a LIST of
        arcs - NewBlend rejects single-curve cyclic loops ("internal error code 1").
        """
        try:
            seg_type = seg.get("type", "Line")
            ax, ay = self._PLANE_AXES[plane[0] if plane else 'z']
            if seg_type == "Line":
                p0 = self._json_point(seg["start"], plane)
                p1 = self._json_point(seg["end"], plane)
                if p0.DistanceTo(p1) < self._MIN_LEN:
                    return None
                return Line.CreateBound(p0, p1)
            elif seg_type in ("Arc3P", "ArcThreePoint"):
                p0 = self._json_point(seg["start"], plane)
                p1 = self._json_point(seg["end"], plane)
                pm = self._json_point(seg["mid"], plane)
                if p0.DistanceTo(p1) < self._MIN_LEN:
                    return None
                return Arc.Create(p0, p1, pm)
            elif seg_type == "Spline":
                pts = [self._json_point(p, plane) for p in (seg.get("points") or [])]
                clean = []
                for p in pts:
                    if not clean or clean[-1].DistanceTo(p) >= self._MIN_LEN / 4.0:
                        clean.append(p)
                if len(clean) < 3:
                    return None
                net_pts = _NetList[XYZ]()
                for p in clean:
                    net_pts.Add(p)
                return HermiteSpline.Create(net_pts, False)
            elif seg_type in ("Arc", "Circle"):
                nc = self._json_point(seg["center"], plane)
                r = seg["radius"] * SCL
                if seg_type == "Circle":
                    a0, a1 = 0.0, FULL_TURN
                else:
                    a0 = seg.get("start_angle", 0.0)
                    a1 = seg.get("end_angle", FULL_TURN)
                if abs(a1 - a0) >= 6.2831 and split_full >= 2:
                    step = (a1 - a0) / split_full
                    return [Arc.Create(nc, r, a0 + i * step, a0 + (i + 1) * step, ax, ay)
                            for i in range(split_full)]
                return Arc.Create(nc, r, a0, a1, ax, ay)
            elif seg_type == "Ellipse":
                nc = self._json_point(seg["center"], plane)
                rx = seg["radius_x"] * SCL
                ry = seg["radius_y"] * SCL
                a0 = seg.get("start_angle", 0.0)
                a1 = seg.get("end_angle", FULL_TURN)
                if abs(a1 - a0) >= 6.2831 and split_full >= 2:
                    step = (a1 - a0) / split_full
                    return [Ellipse.CreateCurve(nc, rx, ry, ax, ay,
                                                a0 + i * step, a0 + (i + 1) * step)
                            for i in range(split_full)]
                return Ellipse.CreateCurve(nc, rx, ry, ax, ay, a0, a1)
        except Exception as ex:
            logger.warning("JSON curve skip: %s", ex)
        return None

    @staticmethod
    def _endpoints(curve):
        try:
            return curve.GetEndPoint(0), curve.GetEndPoint(1)
        except Exception:
            return None, None              # closed curve (full circle/ellipse)

    def _heal_loop(self, curves, close=True):
        """Snap sub-2mm endpoint gaps, then bridge remaining spans with lines."""
        if len(curves) < 2:
            return curves
        curves = list(curves)
        n = len(curves)
        pair_count = n if close else n - 1
        for i in range(pair_count):
            j = (i + 1) % n
            a_end = self._endpoints(curves[i])[1]
            b_start = self._endpoints(curves[j])[0]
            if a_end is None or b_start is None:
                continue
            gap = a_end.DistanceTo(b_start)
            if gap <= 1e-9 or gap > self._SNAP_TOL:
                continue
            if isinstance(curves[i], Line):
                a_start = curves[i].GetEndPoint(0)
                if a_start.DistanceTo(b_start) >= self._MIN_LEN:
                    curves[i] = Line.CreateBound(a_start, b_start)
            elif isinstance(curves[j], Line):
                b_end = curves[j].GetEndPoint(1)
                if a_end.DistanceTo(b_end) >= self._MIN_LEN:
                    curves[j] = Line.CreateBound(a_end, b_end)
        healed = []
        for i in range(n):
            healed.append(curves[i])
            if i == n - 1 and not close:
                break
            j = (i + 1) % n
            a_end = self._endpoints(curves[i])[1]
            b_start = self._endpoints(curves[j])[0]
            if a_end is None or b_start is None:
                continue
            if a_end.DistanceTo(b_start) >= self._MIN_LEN:
                healed.append(Line.CreateBound(a_end, b_start))
        return healed

    def _json_loop(self, segs, plane, close=True, split_full=0):
        curves = []
        for seg in segs or []:
            c = self._json_curve(seg, plane, split_full)
            if isinstance(c, list):
                curves.extend(c)
            elif c is not None:
                curves.append(c)
        return self._heal_loop(curves, close)

    @staticmethod
    def _curve_arr(curves):
        arr = CurveArray()
        for c in curves:
            arr.Append(c)
        return arr

    def _json_profile(self, geom_data, plane):
        outer = self._json_loop(geom_data.get("profile", []), plane)
        if not outer:
            return None
        profile = CurveArrArray()
        profile.Append(self._curve_arr(outer))
        for inner in geom_data.get("inner_loops", []):
            inner_curves = self._json_loop(inner, plane)
            if inner_curves:
                profile.Append(self._curve_arr(inner_curves))
        return profile

    def _uv(self, p, plane):
        ax, ay = self._PLANE_AXES[plane[0] if plane else 'z']
        return (p.X * ax.X + p.Y * ax.Y + p.Z * ax.Z,
                p.X * ay.X + p.Y * ay.Y + p.Z * ay.Z)

    def _loop_area(self, curves, plane):
        """Signed area of a loop in sketch-plane UV coords (CCW > 0)."""
        pts = []
        for c in curves:
            try:
                tess = list(c.Tessellate())
            except Exception:
                continue
            for p in tess[:-1]:
                pts.append(self._uv(p, plane))
        if len(pts) < 3:
            return 0.0
        area = 0.0
        for i in range(len(pts)):
            u0, v0 = pts[i]
            u1, v1 = pts[(i + 1) % len(pts)]
            area += u0 * v1 - u1 * v0
        return 0.5 * area

    @staticmethod
    def _reversed_loop(curves):
        return [c.CreateReversed() for c in reversed(curves)]

    def _align_loop_start(self, loop, ref_loop, plane):
        """Rotate a blend loop so its start vertex matches the reference loop's.

        NewBlend pairs first vertices; a mismatch twists the solid or fails
        outright ("internal error code 1")."""
        if len(loop) < 2:
            return loop

        def _starts(curves):
            pts = []
            for c in curves:
                s = self._endpoints(c)[0]
                if s is None:
                    return None
                pts.append(self._uv(s, plane))
            return pts

        ref_pts = _starts(ref_loop)
        pts = _starts(loop)
        if not ref_pts or not pts:
            return loop
        rcu = sum(p[0] for p in ref_pts) / len(ref_pts)
        rcv = sum(p[1] for p in ref_pts) / len(ref_pts)
        cu = sum(p[0] for p in pts) / len(pts)
        cv = sum(p[1] for p in pts) / len(pts)
        ref_ang = math.atan2(ref_pts[0][1] - rcv, ref_pts[0][0] - rcu)
        best_k, best_d = 0, None
        for k in range(len(pts)):
            ang = math.atan2(pts[k][1] - cv, pts[k][0] - cu)
            d = abs(math.atan2(math.sin(ang - ref_ang), math.cos(ang - ref_ang)))
            if best_d is None or d < best_d:
                best_k, best_d = k, d
        if best_k == 0:
            return loop
        return loop[best_k:] + loop[:best_k]

    def _loop_centroid_uv(self, curves, plane):
        us, vs = [], []
        for c in curves:
            try:
                tess = list(c.Tessellate())
            except Exception:
                continue
            for p in tess[:-1]:
                u, v = self._uv(p, plane)
                us.append(u)
                vs.append(v)
        if not us:
            return None
        return (sum(us) / len(us), sum(vs) / len(vs))

    def _loops_congruent(self, loop_a, loop_b, plane):
        """Same area and centroid: a Blend that is really a prism (NewBlend fails)."""
        area_a = abs(self._loop_area(loop_a, plane))
        area_b = abs(self._loop_area(loop_b, plane))
        if area_a <= 0 or area_b <= 0:
            return False
        if abs(area_a - area_b) > 0.02 * max(area_a, area_b):
            return False
        ca = self._loop_centroid_uv(loop_a, plane)
        cb = self._loop_centroid_uv(loop_b, plane)
        if ca is None or cb is None:
            return False
        du, dv = ca[0] - cb[0], ca[1] - cb[1]
        return (du * du + dv * dv) ** 0.5 <= self._SNAP_TOL

    def _loop_plane_offset(self, segs, plane):
        """Constant signed offset (feet) of raw JSON points from the plane."""
        kind, val = plane
        idx = {'x': 0, 'y': 1, 'z': 2}[kind]
        vals = []
        for seg in segs or []:
            for key in ("start", "end", "center", "mid"):
                if key in seg:
                    try:
                        vals.append(seg[key][idx] * SCL)
                    except Exception:
                        pass
            for p in seg.get("points") or []:
                try:
                    vals.append(p[idx] * SCL)
                except Exception:
                    pass
        if not vals:
            return 0.0
        lo, hi = min(vals), max(vals)
        if hi - lo > self._SNAP_TOL:
            return 0.0
        return (lo + hi) / 2.0 - val

    @staticmethod
    def _set_offsets(form, start_ft, end_ft, start_attr, end_attr):
        """Assign two offsets keeping end > start at every step."""
        if end_ft > 0:
            setattr(form, end_attr, end_ft)
            setattr(form, start_attr, start_ft)
        else:
            setattr(form, start_attr, start_ft)
            setattr(form, end_attr, end_ft)

    # ── one geometry entry -> forms ─────────────────────────────────────
    def _build_entry(self, geom_data, label):
        """Create the form(s) for one entry. Returns (forms, skip_reason)."""
        doc = self.doc
        geom_type = geom_data.get("type", "Extrusion")
        is_solid = geom_data.get("is_solid", True)
        plane = self._plane_info(geom_data)
        sketch_plane = self._make_sketch_plane(plane)

        if geom_type == "Extrusion":
            profile = self._json_profile(geom_data, plane)
            if not profile:
                return [], "empty/invalid profile"
            start_ft = geom_data.get("extrusion_start", 0.0) * SCL
            end_ft = geom_data.get("extrusion_end", 1.0) * SCL
            if end_ft < start_ft:
                start_ft, end_ft = end_ft, start_ft
            if end_ft - start_ft < self._MIN_LEN:
                end_ft = start_ft + self._MIN_LEN
            ext = doc.FamilyCreate.NewExtrusion(is_solid, profile, sketch_plane, end_ft - start_ft)
            self._set_offsets(ext, start_ft, end_ft, 'StartOffset', 'EndOffset')
            return [ext], None

        if geom_type == "Blend":
            base_segs = geom_data.get("profile", [])
            top_segs = geom_data.get("top_profile", [])
            base_loop = self._json_loop(base_segs, plane, split_full=max(2, len(top_segs or [])))
            top_loop = self._json_loop(top_segs, plane, split_full=max(2, len(base_segs or [])))
            if not base_loop or not top_loop:
                return [], "Blend needs both 'profile' and 'top_profile'"
            if self._loop_area(base_loop, plane) < 0:
                base_loop = self._reversed_loop(base_loop)
            if self._loop_area(top_loop, plane) < 0:
                top_loop = self._reversed_loop(top_loop)
            base_off = geom_data.get("base_offset")
            top_off = geom_data.get("top_offset")
            base_ft = base_off * SCL if base_off is not None else self._loop_plane_offset(base_segs, plane)
            top_ft = top_off * SCL if top_off is not None else self._loop_plane_offset(top_segs, plane)
            if top_ft < base_ft:
                base_ft, top_ft = top_ft, base_ft
            if top_ft - base_ft < self._MIN_LEN:
                top_ft = base_ft + self._MIN_LEN
            if self._loops_congruent(base_loop, top_loop, plane):
                prism = CurveArrArray()
                prism.Append(self._curve_arr(base_loop))
                ext = doc.FamilyCreate.NewExtrusion(is_solid, prism, sketch_plane, top_ft - base_ft)
                self._set_offsets(ext, base_ft, top_ft, 'StartOffset', 'EndOffset')
                return [ext], None
            top_aligned = self._align_loop_start(top_loop, base_loop, plane)
            attempts = [(top_aligned, base_loop)]
            if top_aligned is not top_loop:
                attempts.append((top_loop, base_loop))
            attempts.append((self._reversed_loop(top_aligned), self._reversed_loop(base_loop)))
            blend, last_err = None, None
            for t_loop, b_loop in attempts:
                try:
                    blend = doc.FamilyCreate.NewBlend(
                        is_solid, self._curve_arr(t_loop), self._curve_arr(b_loop), sketch_plane)
                    break
                except Exception as blend_err:
                    last_err = blend_err
            if blend is None:
                raise last_err
            if base_ft < 0:
                blend.BaseOffset = base_ft
                blend.TopOffset = top_ft
            else:
                blend.TopOffset = top_ft
                blend.BaseOffset = base_ft
            return [blend], None

        if geom_type == "Revolution":
            profile = self._json_profile(geom_data, plane)
            ax_pt = geom_data.get("axis_start")
            bx_pt = geom_data.get("axis_end")
            if not (profile and ax_pt and bx_pt):
                return [], "Revolution needs 'profile', 'axis_start', 'axis_end'"
            p0 = self._json_point(ax_pt, plane)
            p1 = self._json_point(bx_pt, plane)
            if p0.DistanceTo(p1) < self._MIN_LEN:
                return [], "Revolution axis has zero length"
            rev = doc.FamilyCreate.NewRevolution(
                is_solid, profile, sketch_plane, Line.CreateBound(p0, p1),
                geom_data.get("start_angle", 0.0), geom_data.get("end_angle", FULL_TURN))
            return [rev], None

        if geom_type == "SweptBlend":
            return self._swept_blend(geom_data, plane, sketch_plane, is_solid)

        if geom_type == "Sweep":
            path_curves = self._json_loop(geom_data.get("path", []), plane, close=False)
            profile_segs = geom_data.get("profile", [])
            if geom_data.get("profile_local"):
                # Local cross-section -> world, placed at the path start.
                (start_pt, start_t), _ = path_ends(geom_data.get("path", []))
                profile_segs = profile_at(profile_segs, start_pt, start_t)
            prof_curves = self._json_loop(profile_segs, None)
            if not path_curves or not prof_curves:
                return [], "Sweep needs both 'path' and 'profile'"
            prof_arr = CurveArrArray()
            prof_arr.Append(self._curve_arr(prof_curves))
            sweep_profile = doc.Application.Create.NewCurveLoopsProfile(prof_arr)
            try:
                sweep = doc.FamilyCreate.NewSweep(
                    is_solid, self._curve_arr(path_curves), sketch_plane,
                    sweep_profile, 0, ProfilePlaneLocation.Start)
                return [sweep], None
            except Exception:
                # Sharp corners often kill a multi-segment sweep - rebuild it
                # as one sweep per path segment instead of losing the part.
                if len(path_curves) < 2:
                    raise
                created = []
                for pc in path_curves:
                    try:
                        one_path = CurveArray()
                        one_path.Append(pc)
                        seg_prof = CurveArrArray()
                        seg_prof.Append(self._curve_arr(
                            self._json_loop(profile_segs, None)))
                        created.append(doc.FamilyCreate.NewSweep(
                            is_solid, one_path, sketch_plane,
                            doc.Application.Create.NewCurveLoopsProfile(seg_prof),
                            0, ProfilePlaneLocation.Start))
                    except Exception:
                        pass
                if not created:
                    raise
                self.report['warnings'].append("{}: built as {} separate sweep segment(s)"
                                               .format(label, len(created)))
                return created, None

        if geom_type == "Cylinder":
            s = geom_data.get("start")
            e = geom_data.get("end")
            r = geom_data.get("radius")
            if not (s and e and r is not None):
                return [], "Cylinder needs 'start', 'end', 'radius'"
            dx = abs(e[0] - s[0]); dy = abs(e[1] - s[1]); dz = abs(e[2] - s[2])
            tol = 1.0   # mm - treat as axis-aligned within 1 mm
            kind = None
            if dx <= tol and dy <= tol and dz > tol:
                kind = 'z'
            elif dy <= tol and dz <= tol and dx > tol:
                kind = 'x'
            elif dx <= tol and dz <= tol and dy > tol:
                kind = 'y'
            if kind is not None:
                val = {'x': s[0], 'y': s[1], 'z': s[2]}[kind] * SCL
                cyl_plane = (kind, val)
                cyl_sp = self._make_sketch_plane(cyl_plane)
                circ = self._json_curve({"type": "Circle", "center": s, "radius": r}, cyl_plane)
                prof = CurveArrArray()
                prof.Append(self._curve_arr([circ]))
                length = {'x': e[0], 'y': e[1], 'z': e[2]}[kind] * SCL - val
                ext = doc.FamilyCreate.NewExtrusion(is_solid, prof, cyl_sp, abs(length))
                if length >= 0:
                    ext.EndOffset = length
                    ext.StartOffset = 0.0
                else:
                    ext.StartOffset = length
                    ext.EndOffset = 0.0
                return [ext], None
            # Diagonal rod: a Revolution of a rectangle about the rod's own axis
            # (Revit's sweep engine is unreliable here, NewRevolution is not).
            p0 = self._json_point(s, None)
            p1 = self._json_point(e, None)
            if p0.DistanceTo(p1) < self._MIN_LEN:
                return [], "Cylinder has zero length"
            rr = r * SCL
            axis_dir = (p1 - p0).Normalize()
            ref = XYZ.BasisZ if abs(axis_dir.Z) < 0.9 else XYZ.BasisX
            normal = axis_dir.CrossProduct(ref).Normalize()
            out = axis_dir.CrossProduct(normal).Normalize()
            o0 = p0 + out.Multiply(rr)
            o1 = p1 + out.Multiply(rr)
            rect = CurveArray()
            rect.Append(Line.CreateBound(p0, p1))
            rect.Append(Line.CreateBound(p1, o1))
            rect.Append(Line.CreateBound(o1, o0))
            rect.Append(Line.CreateBound(o0, p0))
            prof = CurveArrArray()
            prof.Append(rect)
            cyl_sp = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(normal, p0))
            rev = doc.FamilyCreate.NewRevolution(
                is_solid, prof, cyl_sp, Line.CreateBound(p0, p1), 0.0, FULL_TURN)
            return [rev], None

        return [], "unsupported type '{}'".format(geom_type)

    # ── detail levels ───────────────────────────────────────────────────
    def _set_detail_levels(self, form, levels, label):
        if not levels or set(levels) >= set(DETAIL_LEVELS):
            return
        try:
            vis = FamilyElementVisibility(FamilyElementVisibilityType.Model)
            vis.IsShownInCoarse = 'coarse' in levels
            vis.IsShownInMedium = 'medium' in levels
            vis.IsShownInFine = 'fine' in levels
            form.SetVisibility(vis)
        except Exception as ex:
            self.report['warnings'].append("{}: detail levels not set ({})".format(label, ex))

    # ── reference planes, dimensions, locks ─────────────────────────────
    def _views(self):
        plan = elev = None
        with disposing(FilteredElementCollector(self.doc)) as collector:
            for view in collector.OfClass(View):
                try:
                    if view.IsTemplate:
                        continue
                    if view.ViewType == ViewType.FloorPlan and plan is None:
                        plan = view
                    elif (view.ViewType == ViewType.Elevation and elev is None
                          and abs(view.ViewDirection.Y) > 0.99):
                        elev = view
                except Exception:
                    continue
        return plan, elev

    def _builtin_references(self):
        refs = {}
        with disposing(FilteredElementCollector(self.doc)) as collector:
            for rp in collector.OfClass(ReferencePlane):
                if rp.Name in BUILTIN_PLANES:
                    refs[rp.Name.lower()] = rp.GetReference()
        with disposing(FilteredElementCollector(self.doc)) as collector:
            for level in collector.OfClass(Level):
                refs.setdefault('ref. level', level.GetPlaneReference())
                break
        return refs

    def apply_constraints(self, schema, bbox):
        planes = [p for p in schema.get('reference_planes') or [] if isinstance(p, dict)]
        dims = [d for d in schema.get('dimensions') or [] if isinstance(d, dict)]
        locked = [(i, g) for i, g in enumerate(schema.get('geometry') or [])
                  if isinstance(g, dict) and isinstance(g.get('locks'), dict) and i in self._forms]
        if not (planes or dims or locked):
            return
        plan, elev = self._views()
        if plan is None or elev is None:
            self.report['warnings'].append(
                "The template has no plan or front elevation view; reference planes, "
                "dimensions and locks were skipped.")
            return
        info = dict((n.lower(), (a, o * SCL)) for n, (a, o) in BUILTIN_PLANES.items())
        refs = self._builtin_references()
        big = 3.0   # ft: plane extent either side
        for spec in planes:
            name, axis, off = spec['name'].strip(), spec['axis'], float(spec['offset']) * SCL
            try:
                if axis == 'x':
                    rp = self.doc.FamilyCreate.NewReferencePlane(
                        XYZ(off, -big, 0), XYZ(off, big, 0), XYZ.BasisZ, plan)
                elif axis == 'y':
                    rp = self.doc.FamilyCreate.NewReferencePlane(
                        XYZ(-big, off, 0), XYZ(big, off, 0), XYZ.BasisZ, plan)
                else:
                    rp = self.doc.FamilyCreate.NewReferencePlane(
                        XYZ(-big, 0, off), XYZ(big, 0, off), XYZ.BasisY, elev)
                rp.Name = name
                refs[name.lower()] = rp.GetReference()
                info[name.lower()] = (axis, off)
                self.report['reference_planes'].append(name)
            except Exception as ex:
                self.report['warnings'].append("Reference plane '{}' not created: {}".format(name, ex))
        self.doc.Regenerate()
        # Locks before dimensions: labelling a dimension then flexes the locked faces.
        for index, geom in locked:
            self._lock_faces(index, geom, refs, info, plan, elev)
        params = self._family_params()
        lo = [v * SCL for v in (bbox[0] if bbox else (-500, -500, 0))]
        hi = [v * SCL for v in (bbox[1] if bbox else (500, 500, 1000))]
        margin = 300 * SCL
        for spec in dims:
            pname = (spec.get('parameter') or '').strip()
            a, b = (spec.get('from') or '').strip().lower(), (spec.get('to') or '').strip().lower()
            label = "Dimension '{}'".format(pname)
            if a not in refs or b not in refs:
                self.report['warnings'].append("{}: a plane is missing; skipped".format(label))
                continue
            axis = info[a][0]
            try:
                arr = ReferenceArray()
                arr.Append(refs[a])
                arr.Append(refs[b])
                if axis == 'x':
                    line, view = Line.CreateBound(XYZ(lo[0] - margin, hi[1] + margin, 0),
                                                  XYZ(hi[0] + margin, hi[1] + margin, 0)), plan
                elif axis == 'y':
                    line, view = Line.CreateBound(XYZ(hi[0] + margin, lo[1] - margin, 0),
                                                  XYZ(hi[0] + margin, hi[1] + margin, 0)), plan
                else:
                    line, view = Line.CreateBound(XYZ(hi[0] + margin, lo[1], lo[2] - margin),
                                                  XYZ(hi[0] + margin, lo[1], hi[2] + margin)), elev
                dim = self.doc.FamilyCreate.NewDimension(view, line, arr)
                fparam = params.get(pname)
                if dim is not None and fparam is not None:
                    dim.FamilyLabel = fparam
                self.report['dimensions'].append(pname)
            except Exception as ex:
                self.report['warnings'].append("{} not created or labelled: {}".format(label, ex))

    # ── Void Forms ──────────────────────────────────────────────────────
    def _cut_with_voids(self):
        """Make every void cut the solids it overlaps (Revit's Cut Geometry).

        Revit usually joins a new void to what it touches by itself; this asks
        for the cut explicitly so openings and hollows never stay uncut. A pair
        Revit refuses (already cut, not touching) is simply skipped.
        """
        voids, solids = [], []
        for forms in self._forms.values():
            for form in forms:
                try:
                    (solids if form.IsSolid else voids).append(form)
                except Exception:
                    continue
        if not voids or not solids:
            return
        try:
            from Autodesk.Revit.DB import SolidSolidCutUtils
        except Exception:
            return
        self.doc.Regenerate()
        cuts = 0
        for void in voids:
            vbox = void.get_BoundingBox(None)
            for solid in solids:
                sbox = solid.get_BoundingBox(None)
                if vbox is None or sbox is None or not (
                        vbox.Min.X <= sbox.Max.X and sbox.Min.X <= vbox.Max.X and
                        vbox.Min.Y <= sbox.Max.Y and sbox.Min.Y <= vbox.Max.Y and
                        vbox.Min.Z <= sbox.Max.Z and sbox.Min.Z <= vbox.Max.Z):
                    continue
                try:
                    SolidSolidCutUtils.AddCutBetweenSolids(self.doc, solid, void)
                    cuts += 1
                except Exception:
                    pass
        self.report['void_cuts'] = cuts

    # ── Swept Blend ─────────────────────────────────────────────────────
    def _swept_blend(self, geom_data, plane, sketch_plane, is_solid):
        """Revit Swept Blend: one path curve, a start and an end section.

        The local sections are placed on the planes normal to the path at its
        two ends (FamilyGen.form_frames). If Revit refuses that placement, the
        sections are retried in the XY plane, which is how the Revit UI sketches
        a swept-blend profile before mapping it onto the path.
        """
        doc = self.doc
        path_segs = geom_data.get("path", [])
        path_curves = self._json_loop(path_segs, plane, close=False)
        if len(path_curves) != 1:
            return [], "SweptBlend needs exactly one path curve"
        (p0, t0), (p1, t1) = path_ends(path_segs)

        def _profile(segs):
            arr = CurveArrArray()
            arr.Append(self._curve_arr(self._json_loop(segs, None)))
            return doc.Application.Create.NewCurveLoopsProfile(arr)

        bottom, top = geom_data.get("profile", []), geom_data.get("top_profile", [])
        attempts = (
            ('path ends', profile_at(bottom, p0, t0), profile_at(top, p1, t1)),
            ('XY plane', local_to_world(bottom, [0, 0, 0], [1, 0, 0], [0, 1, 0]),
             local_to_world(top, [0, 0, 0], [1, 0, 0], [0, 1, 0])),
        )
        last_error = None
        for where, seg_a, seg_b in attempts:
            try:
                form = doc.FamilyCreate.NewSweptBlend(
                    is_solid, path_curves[0], sketch_plane, _profile(seg_a), _profile(seg_b))
                if form is not None:
                    if where != 'path ends':
                        self.report['warnings'].append(
                            "{}: swept blend sections placed from the {}".format(
                                geom_data.get('id') or 'SweptBlend', where))
                    return [form], None
            except Exception as ex:
                last_error = ex
        return [], "Revit refused the swept blend ({})".format(last_error)

    def _lock_faces(self, index, geom, refs, info, plan, elev):
        label = geom.get('id') or '#{}'.format(index + 1)
        tol = 1.0 * SCL
        opts = Options()
        opts.ComputeReferences = True
        for key, plane_name in sorted(geom['locks'].items()):
            pkey = (u'{}'.format(plane_name)).strip().lower()
            if pkey not in refs or pkey not in info:
                self.report['warnings'].append("{}: lock {} -> unknown plane".format(label, key))
                continue
            axis, offset = info[pkey]
            sign = -1.0 if key.endswith('_min') else 1.0
            k = 'xyz'.index(axis)
            done = False
            for form in self._forms.get(index, []):
                try:
                    for obj in form.get_Geometry(opts):
                        if not isinstance(obj, Solid):
                            continue
                        for face in obj.Faces:
                            if not isinstance(face, PlanarFace):
                                continue
                            n = face.FaceNormal
                            if [n.X, n.Y, n.Z][k] * sign < 0.99:
                                continue
                            if abs([face.Origin.X, face.Origin.Y, face.Origin.Z][k] - offset) > tol:
                                continue
                            align = self.doc.FamilyCreate.NewAlignment(
                                elev if axis == 'z' else plan, refs[pkey], face.Reference)
                            if align is not None:
                                align.IsLocked = True
                                self.report['locks'] += 1
                                done = True
                            break
                except Exception as ex:
                    self.report['warnings'].append("{}: {} not locked ({})".format(label, key, ex))
                    done = True
                if done:
                    break
            if not done:
                self.report['warnings'].append(
                    "{}: no face lies on plane '{}' for {}; not locked".format(label, plane_name, key))

    def build(self, schema):
        """Materials, parameters, geometry, then constraints. Returns the report."""
        schema = expand_schema(schema)
        self._ensure_current_type(schema.get('family_name'))
        self.apply_materials(schema.get('materials'))
        self.apply_parameters(schema.get('parameters'))
        self.apply_types(schema.get('types'))
        geometry = schema.get("geometry", [])
        self.report['total'] = len(geometry)
        for idx, geom_data in enumerate(geometry):
            label = geom_data.get("id") or "#{} ({})".format(idx + 1, geom_data.get("type"))
            try:
                created, reason = self._build_entry(geom_data, label)
                if not created:
                    self.report['skipped'].append("{}: {}".format(label, reason))
                    continue
                is_solid = geom_data.get("is_solid", True)
                self._forms[idx] = list(created)
                for form in created:
                    self._set_detail_levels(form, geom_data.get('detail_levels'), label)
                    if is_solid and geom_data.get('material'):
                        self._bind_material(form, geom_data['material'], label)
                    if geom_data.get('subcategory'):
                        try:
                            form.Subcategory = self._subcategory(geom_data['subcategory'].strip())
                        except Exception as ex:
                            self.report['warnings'].append(
                                "{}: subcategory '{}' not applied ({})"
                                .format(label, geom_data['subcategory'], ex))
                self.report['built'] += 1
            except Exception as ex:
                msg = "{}".format(ex)
                if "conditions for the inputs" in msg:
                    msg += " [profile likely open or self-intersecting]"
                elif "internal error" in msg.lower():
                    msg += " [blend loops may self-intersect or pair badly]"
                self.report['skipped'].append("{}: {}".format(label, msg))
                logger.warning("JSON geometry skip %s: %s", label, ex)
        try:
            from FamilyGen.preview_mesh import build_preview
            model = build_preview(schema)
            bbox = (model.bbox_min, model.bbox_max) if model.bbox_min is not None else None
        except Exception:
            bbox = None
        self._cut_with_voids()
        try:
            self.apply_constraints(schema, bbox)
        except Exception as ex:
            self.report['warnings'].append("Constraints skipped: {}".format(ex))
        return self.report


# ── public entry points ──────────────────────────────────────────────────────

def _checked(schema):
    errors, warnings = validate_family_schema(schema)
    if errors:
        shown = errors[:8]
        more = len(errors) - len(shown)
        raise FamilyBuildError(
            "The family JSON is not valid, so nothing was created:\n- " + "\n- ".join(shown)
            + ("\n- ... and {} more".format(more) if more else "")
            + "\nFix the JSON (or ask the AI to repair it) and try again.")
    return warnings


def default_output_folder():
    """Documents\\T3Lab\\FamiGen - where MCP-created families go by default."""
    return os.path.join(os.path.expanduser('~'), 'Documents', 'T3Lab', 'FamiGen')


def safe_file_name(name):
    cleaned = re.sub(r'[\\/*?:"<>|]', "_", (name or '').strip()) or 'T3Lab_Family'
    return cleaned[:120]


def build_into_document(fam_doc, schema):
    """Add the schema's geometry to an already open family document (one Transaction)."""
    warnings = _checked(schema)
    builder = JsonFamilyBuilder(fam_doc)
    with disposing(Transaction(fam_doc, "T3Lab - FamiGen")) as t:
        _start(t)
        report = builder.build(schema)
        t.Commit()
    report['warnings'] = list(warnings) + report['warnings']
    return report


def create_family(app, schema, output_folder, project_doc=None, load_into_project=False):
    """New family from the category template -> build -> save .rfa -> optional load.

    Returns the build report with 'saved_path', 'template', 'category' and
    'loaded'. Raises FamilyBuildError with a user-facing message.
    """
    warnings = _checked(schema)
    category = schema['family_category']
    template, recategorize = resolve_template(app, category)
    if not output_folder:
        raise FamilyBuildError("Choose an output folder for the family, then try again.")
    if not os.path.isdir(output_folder):
        try:
            os.makedirs(output_folder)
        except OSError as ex:
            raise FamilyBuildError("The output folder '{}' cannot be created: {}"
                                   .format(output_folder, ex))
    save_path = os.path.join(output_folder, safe_file_name(schema.get('family_name')) + '.rfa')

    fam_doc = app.NewFamilyDocument(template)
    try:
        builder = JsonFamilyBuilder(fam_doc)
        with disposing(Transaction(fam_doc, "T3Lab - FamiGen")) as t:
            _start(t)
            if recategorize:
                bic = getattr(BuiltInCategory, category_bic_name(category), None)
                if bic is None:
                    raise FamilyBuildError("Category '{}' does not exist in this Revit version."
                                           .format(category))
                fam_doc.OwnerFamily.FamilyCategory = Category.GetCategory(fam_doc, bic)
                builder.report['warnings'].append(
                    "No '{}' template was found; the family was built from the Generic "
                    "Model template and its category changed to {}.".format(category, category))
            report = builder.build(schema)
            if report['built'] == 0:
                raise FamilyBuildError(
                    "No part could be built, so the family was not saved:\n- "
                    + "\n- ".join(report['skipped'][:8]))
            t.Commit()
        options = SaveAsOptions()
        options.OverwriteExistingFile = True
        try:
            fam_doc.SaveAs(save_path, options)
        except Exception as ex:
            raise FamilyBuildError(
                "The family was built but could not be saved to '{}': {}. Close that file "
                "if it is open in Revit, or choose another folder.".format(save_path, ex))
        report['loaded'] = False
        if load_into_project and project_doc is not None and not project_doc.IsFamilyDocument:
            try:
                family = fam_doc.LoadFamily(project_doc, _OverwriteLoadOptions())
                report['loaded'] = family is not None
            except Exception as ex:
                report['warnings'].append(
                    "Saved, but not loaded into '{}': {}".format(project_doc.Title, ex))
    finally:
        try:
            fam_doc.Close(False)
        except Exception:
            pass
    report['warnings'] = list(warnings) + report['warnings']
    report.update({'saved_path': save_path, 'template': os.path.basename(template),
                   'category': category})
    return report


# Pure, so FamiGen Studio's controller and its tests can read a report without Revit.
from FamilyGen.builder_report import report_lines  # noqa: E402,F401
