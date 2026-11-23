# -*- coding: utf-8 -*-
"""CAD to Elements planning: from scanned CAD layers (Snippets/_cad_revit.scan_cad)
and plain options to a Plan (count, confirm text, details, execute).

Shared by the CAD to Elements window (GUI/CADToElementsDialog.py reads its
controls, then calls plan_<mode>) and the cad_to_elements AI tool
(core/tools/from_source.py). Lengths are in feet; `fmt(feet)` writes one for
the details text (the window: project unit; the tool: mm). Each execute runs
one Snippets/_cad_revit create_* call = one Transaction.
Test: dev/test_tools_create_domain.py.
"""
import re

from Snippets import _cad_geometry as geo


class PlanError(Exception):
    """An input the caller has to fix; the message is shown as-is."""


_ONE_PLURAL = re.compile(r"\b1 ([A-Za-z][A-Za-z ]*?)s\b")


class Plan(object):
    """What a run will do — shown in the confirm dialog before anything changes."""

    def __init__(self, count, noun, question, ok_text, details, execute, empty_msg=None):
        if count == 1:
            question = _ONE_PLURAL.sub(r"1 \1", question, count=1)
            ok_text = _ONE_PLURAL.sub(r"1 \1", ok_text, count=1)
        self.count = count
        self.noun = noun
        self.question = question
        self.ok_text = ok_text
        self.details = details
        self.execute = execute
        self.empty_msg = empty_msg


def _cr():
    from Snippets import _cad_revit
    return _cad_revit


def mm_text(feet):
    """Default `fmt`: a length in whole mm."""
    return u"{:g} mm".format(round(geo.to_mm(feet), 1))


def segments(layers, merge):
    segs = [s for lg in layers for s in lg.segments]
    return geo.merge_collinear(segs) if merge else segs


def outline_groups(layers, with_holes):
    loops = [lp for lg in layers for lp in lg.loops()]
    polys = [geo.loop_polygon(lp) for lp in loops]
    keep = geo.dedupe_polygons(polys)
    loops = [loops[i] for i in keep]
    polys = [polys[i] for i in keep]
    if not with_holes:
        return [(lp, []) for lp in loops], polys, [[] for _ in loops]
    groups = geo.nest_loops(polys)
    return ([(loops[o], [loops[h] for h in holes]) for o, holes in groups],
            [polys[o] for o, _h in groups],
            [[polys[h] for h in holes] for _o, holes in groups])


def _need(item, what):
    if item is None or item.get("element") is None:
        raise PlanError(u"No {} is available.\nLoad one into the model (or pick one in the "
                        u"options card), then run again.".format(what))
    return item


# ── Walls ──

def plan_wall(doc, layers, level, level_name, height, offset, default_thk, include, merge,
              part_category=None, wall_type=None, match=True, structural=False, fmt=mm_text):
    segs = segments(layers, merge)
    pairs, unpaired = geo.find_parallel_pairs(segs)
    items = [(c[0], c[1], c[2], c[3], c[4]) for c in pairs]
    if include:
        items += [(s[0], s[1], s[2], s[3], default_thk) for s in unpaired]
    thicknesses = []                # one per wall type (1 mm apart), in the project unit
    for size_mm in sorted(geo.group_by_size(items, lambda it: it[4])):
        text = fmt(geo.mm(size_mm))
        if text not in thicknesses:
            thicknesses.append(text)
    details = [u"Wall pairs found: {}".format(len(pairs)),
               u"Single lines: {} ({})".format(len(unpaired), "included" if include else "ignored"),
               u"Thicknesses: {}".format(", ".join(thicknesses[:12]) or "—"),
               u"Height {} · base offset {}".format(fmt(height), fmt(offset))]
    empty = (u"No two parallel lines lie within {} of each other. Tick the layers that hold "
             u"both faces of each wall, or turn on Include unpaired lines."
             .format(fmt(geo.MAX_WALL_THICKNESS)))
    n = len(items)
    if part_category is not None:
        cat = part_category
        profiles = []
        for (x0, y0, x1, y1, t) in items:
            rect = geo.centerline_rect(x0, y0, x1, y1, t / 2.0)
            lp = geo.loop_from_polygon(rect) if rect else None
            if lp:
                profiles.append((lp, [], level.Elevation + offset, height))
        details.append(u"Category: {}".format(cat["name"]))
        return Plan(len(profiles), "wall parts",
                    u"Create {} wall parts on {}?".format(len(profiles), level_name),
                    u"Create {} Parts".format(len(profiles)), u"\n".join(details),
                    lambda p: _cr().create_parts(doc, profiles, cat["bic"],
                                                 "T3Lab: CAD to Wall Parts", p), empty)
    wtype = _need(wall_type, "basic wall type")
    details.append(u"Base type: {}{}".format(wtype["name"], " (copied per thickness)" if match else ""))
    return Plan(n, "walls", u"Create {} walls on {}?".format(n, level_name),
                u"Create {} Walls".format(n), u"\n".join(details),
                lambda p: _cr().create_walls(doc, items, level, wtype["element"], height,
                                             offset, structural, match, p), empty)


# ── Floors / Ceilings ──

def plan_floor(doc, layers, level, level_name, offset, holes, part_category=None,
               part_thickness=None, floor_type=None, structural=False, fmt=mm_text):
    groups, _p, _h = outline_groups(layers, holes)
    n = len(groups)
    n_holes = sum(len(h) for _o, h in groups)
    details = [u"Outlines: {} · openings: {}".format(n, n_holes),
               u"Height offset {}".format(fmt(offset))]
    empty = ("The ticked layers hold no closed outline. Use closed polylines or lines whose "
             "ends meet, or tick the layer that holds the slab edges.")
    if part_category is not None:
        cat, thk = part_category, part_thickness
        profiles = [(o, h, level.Elevation + offset, thk) for o, h in groups]
        details.append(u"Part thickness {} · category {}".format(fmt(thk), cat["name"]))
        return Plan(n, "floor parts", u"Create {} floor parts on {}?".format(n, level_name),
                    u"Create {} Parts".format(n), u"\n".join(details),
                    lambda p: _cr().create_parts(doc, profiles, cat["bic"],
                                                 "T3Lab: CAD to Floor Parts", p), empty)
    ftype = _need(floor_type, "floor type")
    details.append(u"Floor type: {}".format(ftype["name"]))
    return Plan(n, "floors", u"Create {} floors on {}?".format(n, level_name),
                u"Create {} Floors".format(n), u"\n".join(details),
                lambda p: _cr().create_floors(doc, groups, ftype["id"], level, offset,
                                              structural, p), empty)


def plan_ceiling(doc, layers, level, level_name, offset, holes, part_category=None,
                 part_thickness=None, ceiling_type=None, fmt=mm_text):
    groups, _p, _h = outline_groups(layers, holes)
    n = len(groups)
    details = [u"Outlines: {} · openings: {}".format(n, sum(len(h) for _o, h in groups)),
               u"Height offset {}".format(fmt(offset))]
    empty = ("The ticked layers hold no closed outline. Tick the layer that holds the ceiling "
             "outlines (closed polylines or lines whose ends meet).")
    if part_category is not None:
        cat, thk = part_category, part_thickness
        profiles = [(o, h, level.Elevation + offset, thk) for o, h in groups]
        details.append(u"Part thickness {} · category {}".format(fmt(thk), cat["name"]))
        return Plan(n, "ceiling parts", u"Create {} ceiling parts on {}?".format(n, level_name),
                    u"Create {} Parts".format(n), u"\n".join(details),
                    lambda p: _cr().create_parts(doc, profiles, cat["bic"],
                                                 "T3Lab: CAD to Ceiling Parts", p), empty)
    ctype = _need(ceiling_type, "ceiling type")
    details.append(u"Ceiling type: {}".format(ctype["name"]))
    return Plan(n, "ceilings", u"Create {} ceilings on {}?".format(n, level_name),
                u"Create {} Ceilings".format(n), u"\n".join(details),
                lambda p: _cr().create_ceilings(doc, groups, ctype["id"], level, offset, p),
                empty)


# ── Rooms ──

def plan_room(doc, layers, level, level_name, make_lines, make_rooms, view_item, name, skip):
    view = None
    if make_lines:
        if view_item is None:
            raise PlanError(u"{} has no floor plan to hold room separation lines.\nCreate a "
                            u"floor plan for this level, or choose Rooms only."
                            .format(level_name))
        view = view_item["element"]
    edges = [e for lg in layers for e in lg.all_edges()] if make_lines else []
    points, too_small = [], 0
    if make_rooms:
        _g, polys, _h = outline_groups(layers, False)
        points, too_small = geo.room_points(polys)
    if make_rooms:
        n, noun = len(points), "rooms"
        question = u"Create {} rooms on {}{}?".format(
            n, level_name,
            u" with {} room separation lines".format(len(edges)) if make_lines else u"")
        ok = u"Create {} Rooms".format(n)
    else:
        n, noun = len(edges), "room separation lines"
        question = u"Draw {} room separation lines on {}?".format(n, level_name)
        ok = u"Draw {} Lines".format(n)
    details = [u"Closed outlines (one room each): {}{}".format(
                   len(points), u" · too small or too narrow for a room: {}".format(too_small)
                   if too_small else u"") if make_rooms else u"Rooms: none (lines only)",
               u"Separation lines: {}{}".format(len(edges), u" in view " + view.Name if view else "")]
    empty = ("The ticked layers hold no closed outline to put a room in. Tick the wall or room "
             "boundary layers, or choose Lines only.") if make_rooms else \
        "The ticked layers hold no linework."
    return Plan(n, noun, question, ok, u"\n".join(details),
                lambda p: _cr().create_rooms(doc, level, view, edges, points, make_lines,
                                             make_rooms, name, skip, p), empty)


# ── Columns ──

def plan_column(doc, layers, level, level_name, step, height, rect_item, round_item, top_item,
                structural, match, rotate, size_unit=None, fmt=mm_text):
    rect_sym = rect_item["element"] if rect_item else None
    round_sym = round_item["element"] if round_item else None
    fps = geo.filter_footprints([fp for lg in layers for fp in lg.footprints()])
    rects = [f for f in fps if f["shape"] == "rect"]
    rounds = [f for f in fps if f["shape"] == "round"]
    if rects and rect_sym is None:
        raise PlanError("No column type is loaded for rectangles.\nLoad a column family "
                        "(for example Concrete-Rectangular-Column), then run again.")
    items = (rects if rect_sym is not None else []) + (rounds if round_sym is not None else [])
    top_level = top_item["element"] if top_item else None
    sizes = {}
    for f in items:
        nm = geo.footprint_type_name(f, step, size_unit)
        sizes[nm] = sizes.get(nm, 0) + 1
    details = [u"Rectangles: {} · circles: {}{}".format(
        len(rects), len(rounds), u" (skipped — no round type chosen)" if rounds and round_sym is None else u""),
        u"Sizes: {}".format(u", ".join(u"{} ×{}".format(k, v) for k, v in sorted(sizes.items())[:8]) or u"—"),
        u"Top: {}".format(top_item["name"] if top_level is not None else u"unconnected, {}".format(fmt(height)))]
    n = len(items)
    empty = (u"The ticked layers hold no closed rectangle or circle {}–{} across. Tick the "
             u"layer that holds the column outlines.".format(
                 fmt(geo.mm(geo.FOOTPRINT_MIN_MM)), fmt(geo.mm(geo.FOOTPRINT_MAX_MM))))
    return Plan(n, "columns", u"Place {} columns on {}?".format(n, level_name),
                u"Place {} Columns".format(n), u"\n".join(details),
                lambda p: _cr().create_columns(doc, items, level, top_level, height,
                                               rect_sym, round_sym, structural, match, rotate,
                                               step, p, size_unit=size_unit), empty)


# ── Beams ──

def plan_beam(doc, layers, level, level_name, offset, w_min, w_max, merge, part_category=None,
              beam_type=None, match=True, fmt=mm_text):
    segs = segments(layers, merge)
    pairs, _unpaired = geo.find_parallel_pairs(segs, max_sep=w_max,
                                               min_sep=w_min, min_overlap=0.7)
    items = []
    for (x0, y0, x1, y1, sep, _layer) in pairs:
        w = max(50, geo.round_to(geo.to_mm(sep), 50))
        items.append((x0, y0, x1, y1, w, geo.beam_height_for_width(w)))
    sizes = sorted(set("{}x{}".format(it[4], it[5]) for it in items))
    details = [u"Beam pairs found: {}".format(len(items)),
               u"Sizes (mm): {}".format(", ".join(sizes[:10]) or "—"),
               u"Top offset {}".format(fmt(offset))]
    empty = (u"No two parallel lines {}–{} apart overlap enough to be a beam. Tick the "
             u"layer that holds both edges of each beam.".format(fmt(w_min), fmt(w_max)))
    n = len(items)
    if part_category is not None:
        cat = part_category
        top = level.Elevation + offset
        profiles = []
        for (x0, y0, x1, y1, w, h) in items:
            rect = geo.centerline_rect(x0, y0, x1, y1, geo.mm(w) / 2.0)
            lp = geo.loop_from_polygon(rect) if rect else None
            if lp:
                profiles.append((lp, [], top - geo.mm(h), geo.mm(h)))
        details.append(u"Category: {}".format(cat["name"]))
        return Plan(len(profiles), "beam parts",
                    u"Create {} beam parts on {}?".format(len(profiles), level_name),
                    u"Create {} Parts".format(len(profiles)), u"\n".join(details),
                    lambda p: _cr().create_parts(doc, profiles, cat["bic"],
                                                 "T3Lab: CAD to Beam Parts", p), empty)
    sym = _need(beam_type, "structural framing type")
    details.append(u"Type: {}{}".format(sym["name"], " (copied per size)" if match else ""))
    return Plan(n, "beams", u"Create {} beams on {}?".format(n, level_name),
                u"Create {} Beams".format(n), u"\n".join(details),
                lambda p: _cr().create_beams(doc, items, level, sym["element"], offset,
                                             match, p), empty)


# ── Grids ──

def plan_grid(doc, layers, extend, min_len, auto, start_number, start_letter, skip_existing,
              grid_type=None, fmt=mm_text):
    if auto and geo.letter_index(start_letter) == 0 and start_letter.upper() != "A":
        raise PlanError(u"First letter is \"{}\", which is not a grid letter.\nUse letters "
                        u"A–Z without I and O (for example A, C or AA).".format(start_letter))
    axes = geo.collapse_axis_lines(segments(layers, False), min_len=min_len)
    skipped = 0
    existing, taken = _cr().existing_grid_lines(doc)
    if skip_existing:
        kept = [a for a in axes if not any(geo.same_axis(a, ex) for ex in existing)]
        skipped = len(axes) - len(kept)
        axes = kept
    names = geo.name_grids(axes, start_number, start_letter) if auto else None
    lines = [geo.extend_line(a[0], a[1], a[2], a[3], extend) for a in axes]
    type_id = grid_type["id"] if grid_type else None
    numbers = [nm for nm in (names or []) if nm.isdigit()]
    letters = [nm for nm in (names or []) if not nm.isdigit()]
    details = [u"Axes found: {} · already in the model (skipped): {}".format(len(axes) + skipped, skipped)]
    if auto:
        if numbers:
            details.append(u"Numbered {}–{}".format(numbers and min(numbers, key=int), max(numbers, key=int)))
        if letters:
            details.append(u"Lettered {}–{}".format(geo.grid_letter(geo.letter_index(start_letter)),
                                                      geo.grid_letter(geo.letter_index(start_letter) + len(letters) - 1)))
    n = len(lines)
    empty = (u"The ticked layers hold no straight axis line of at least {}{}. Tick the axis "
             u"layer.".format(fmt(min_len), ", or every axis already has a grid" if skipped else ""))
    return Plan(n, "grids", u"Create {} grids?".format(n), u"Create {} Grids".format(n),
                u"\n".join(details),
                lambda p: _cr().create_grids(doc, lines, names, type_id, set(taken), p), empty)


# ── Lines ──

def plan_lines(doc, layers, level_name, view, z, merge, style_item):
    """view: the detail view for detail lines, None for model lines at height z."""
    edges = []
    for lg in layers:
        segs = geo.merge_collinear(lg.segments) if merge else lg.segments
        edges.extend(("L", (s[0], s[1]), (s[2], s[3]), None) for s in segs)
        edges.extend(lg.arc_edges)
        edges.extend(lg.spline_edges)
        for (cx, cy, r) in lg.circles:
            edges.extend(geo.loop_from_circle(cx, cy, r))
    gs = style_item["element"] if style_item else None
    detail = view is not None
    noun = "detail lines" if detail else "model lines"
    where = u"in view {}".format(view.Name) if detail else u"on {}".format(level_name)
    n = len(edges)
    details = [u"Curves: {}".format(n), u"Line style: {}".format(style_item["name"] if style_item else "default")]
    return Plan(n, noun, u"Create {} {} {}?".format(n, noun, where),
                u"Create {} Lines".format(n), u"\n".join(details),
                lambda p: _cr().create_lines(doc, edges, view, z, gs, p),
                "The ticked layers hold no linework.")


# ── MEP runs ──

def plan_mep(doc, layers, level, level_name, key, mep_type, system_id, double, width, height,
             offset, merge, auto_elbow, fmt=mm_text):
    """width: None for double lines (the pair spacing sizes each run)."""
    cat = geo.mep_category(key)
    tp = mep_type
    segs = segments(layers, merge)
    unpaired = []
    if double:
        pairs, unpaired = geo.find_parallel_pairs(segs, max_sep=geo.mm(2500))
        runs = [(c[0], c[1], c[2], c[3], c[4]) for c in pairs]
    else:
        runs = [(s[0], s[1], s[2], s[3], None) for s in segs]
    n = len(runs)
    details = [u"{}: {}".format("Line pairs" if double else "Lines", n),
               u"Type: {}".format(tp["name"]),
               u"Size: {}".format("from the pair spacing" if double else fmt(width)) +
               (u" × {}".format(fmt(height)) if height else u""),
               u"Offset from level {}".format(fmt(offset))]
    if double and unpaired:
        details.append(u"Single lines ignored: {}".format(len(unpaired)))
    noun = cat["noun"]
    return Plan(n, noun, u"Create {} {} on {}?".format(n, noun, level_name),
                u"Create {} {}".format(n, cat["verb"][len("Create "):]), u"\n".join(details),
                lambda p: _cr().create_mep_runs(doc, key, runs, level, offset, tp["id"],
                                                system_id, width, height, auto_elbow,
                                                noun, "T3Lab: CAD to {}".format(cat["label"] + "s"), p),
                "The ticked layers hold no straight line." if not double else
                "No parallel line pairs found. Switch to Single line or tick the layer with both edges.")


# ═══════════════════════════════════════════════════════════════════════════
# A layer mapping (AI tool) -> plans. Lengths in mm; defaults = the window's.
# ═══════════════════════════════════════════════════════════════════════════

MEP_KEYS = tuple(c["key"] for c in geo.MEP_CATEGORIES)
ELEMENTS = ("wall", "floor", "ceiling", "room", "column", "beam", "grid", "lines") + MEP_KEYS
OPTIONS = {
    "wall": dict(height_mm=3000, offset_mm=0, thickness_mm=200, include_unpaired=False,
                 merge=True, match_size=True, structural=False, part_category=None),
    "floor": dict(offset_mm=0, holes=True, structural=False, part_thickness_mm=200,
                  part_category=None),
    "ceiling": dict(offset_mm=2700, holes=True, part_thickness_mm=25, part_category=None),
    "room": dict(mode="both", name="Room", skip_existing=True, view=None),
    "column": dict(height_mm=3000, rounding_mm=10, top_level=None, structural=True,
                   match_size=True, rotate=True, round_type=None),
    "beam": dict(offset_mm=-50, min_width_mm=50, max_width_mm=1500, merge=True,
                 match_size=True, part_category=None),
    "grid": dict(extend_mm=0, min_length_mm=1000, auto_name=True, start_number=1,
                 start_letter="A", skip_existing=True),
    "lines": dict(offset_mm=0, merge=True, view_id=None),
}
for _c in geo.MEP_CATEGORIES:
    OPTIONS[_c["key"]] = dict(width_mm=_c["width"], height_mm=_c["height_mm"],
                              offset_mm=_c["offset"], double=False, merge=True, elbows=True,
                              system=None)


def _fold(text):
    return u" ".join(u"{}".format(text or u"").split()).lower()


def find_named(items, wanted, what, prefer=()):
    """The item named `wanted` (exact, then a 'Family : Type' ending, then a unique
    substring); no name -> the first whose name holds a `prefer` word, else the
    first. PlanError lists the names when nothing or several match."""
    if not items:
        raise PlanError(u"No {} in this model. Load one, then run again.".format(what))
    if wanted is None or u"{}".format(wanted).strip() == u"":
        for word in prefer:
            for it in items:
                if word in _fold(it["name"]):
                    return it
        return items[0]
    key = _fold(wanted)
    tests = ((True, lambda n: n == key), (False, lambda n: n.endswith(u": " + key)),
             (False, lambda n: key in n))
    for exact, test in tests:
        hits = [it for it in items if test(_fold(it["name"]))]
        if len(hits) == 1 or (hits and exact):
            return hits[0]
        if hits:
            raise PlanError(u"{} \"{}\" matches {} names: {}. Use the full name.".format(
                what, wanted, len(hits), u", ".join(h["name"] for h in hits[:10])))
    raise PlanError(u"No {} named \"{}\". Available: {}.".format(
        what, wanted, u", ".join(it["name"] for it in items[:20])))


def group_mapping(mapping):
    """[(entry, [layer names])]: entries with the same element, type, level and
    options share one plan (wall faces may sit on two layers)."""
    import json
    groups, order = {}, []
    for entry in mapping or []:
        if not isinstance(entry, dict):
            raise PlanError(u"Each mapping entry must be an object with layers and element.")
        layers = entry.get("layers")
        if layers is None:
            layers = entry.get("layer")
        if not isinstance(layers, (list, tuple)):
            layers = [layers] if layers else []
        key = json.dumps([entry.get("element"), entry.get("type"), entry.get("level"),
                          entry.get("options") or {}], sort_keys=True, default=str)
        if key not in groups:
            groups[key] = (entry, [])
            order.append(key)
        for name in layers:
            if u"{}".format(name) not in groups[key][1]:
                groups[key][1].append(u"{}".format(name))
    return [groups[k] for k in order]


class Lookups(object):
    """Model lists for one call, each read once (Snippets/_cad_revit get_*)."""

    def __init__(self, doc):
        self.doc = doc
        self._cache = {}

    def get(self, key, read):
        if key not in self._cache:
            self._cache[key] = read()
        return self._cache[key]

    def levels(self):
        return self.get("levels", lambda: _cr().get_levels(self.doc))

    def symbols(self, bic_name):
        cr = _cr()
        return self.get(bic_name, lambda: cr.get_symbols(
            self.doc, getattr(cr.DB.BuiltInCategory, bic_name)))


def _level(look, name):
    levels = look.levels()
    if name is None or u"{}".format(name).strip() == u"":
        try:
            gen = look.doc.ActiveView.GenLevel
            for lv in levels:
                if gen is not None and lv["name"] == gen.Name:
                    return lv
        except Exception:
            pass
    return find_named(levels, name, "level")


def _options(element, given):
    allowed = OPTIONS[element]
    given = given or {}
    if not isinstance(given, dict):
        raise PlanError(u"options must be an object.")
    unknown = sorted(k for k in given if k not in allowed)
    if unknown:
        raise PlanError(u"{} has no option {}. Options and defaults: {}.".format(
            element, u", ".join(unknown), u", ".join(
                u"{}={}".format(k, allowed[k]) for k in sorted(allowed))))
    out = dict(allowed)
    out.update(given)
    return out


def _mm(opts, key):
    try:
        return geo.mm(float(opts[key]))
    except (TypeError, ValueError):
        raise PlanError(u"{} must be a number of millimetres, not \"{}\".".format(key, opts[key]))


def _part(opts):
    if not opts.get("part_category"):
        return None
    return find_named(_cr().get_ds_categories(), opts["part_category"], "part category")


def plan_entry(doc, look, entry, layers, fmt=mm_text):
    """One mapping group -> Plan. `layers`: the LayerGeoms of its layer names."""
    element = u"{}".format(entry.get("element") or u"").lower()
    if element not in ELEMENTS:
        raise PlanError(u"element must be one of: {}.".format(u", ".join(ELEMENTS)))
    cr = _cr()
    o = _options(element, entry.get("options"))
    wanted = entry.get("type")
    if element == "grid":
        grid_type = find_named(look.get("grid", lambda: cr.get_grid_types(doc)), wanted,
                               "grid type") if wanted else None
        return plan_grid(doc, layers, _mm(o, "extend_mm"), _mm(o, "min_length_mm"),
                         bool(o["auto_name"]), int(o["start_number"]),
                         u"{}".format(o["start_letter"] or "A").strip() or "A",
                         bool(o["skip_existing"]), grid_type, fmt=fmt)
    lv = _level(look, entry.get("level"))
    level, level_name = lv["element"], lv["name"]
    if element == "wall":
        part = _part(o)
        wtype = None if part else find_named(
            look.get("wall", lambda: cr.get_basic_wall_types(doc)), wanted, "basic wall type",
            ("generic",))
        return plan_wall(doc, layers, level, level_name, _mm(o, "height_mm"), _mm(o, "offset_mm"),
                         _mm(o, "thickness_mm"), bool(o["include_unpaired"]), bool(o["merge"]),
                         part_category=part, wall_type=wtype, match=bool(o["match_size"]),
                         structural=bool(o["structural"]), fmt=fmt)
    if element in ("floor", "ceiling"):
        part = _part(o)
        thk = _mm(o, "part_thickness_mm") if part else None
        if element == "floor":
            ftype = None if part else find_named(
                look.get("floor", lambda: cr.get_floor_types(doc)), wanted, "floor type")
            return plan_floor(doc, layers, level, level_name, _mm(o, "offset_mm"),
                              bool(o["holes"]), part_category=part, part_thickness=thk,
                              floor_type=ftype, structural=bool(o["structural"]), fmt=fmt)
        ctype = None if part else find_named(
            look.get("ceiling", lambda: cr.get_ceiling_types(doc)), wanted, "ceiling type")
        return plan_ceiling(doc, layers, level, level_name, _mm(o, "offset_mm"), bool(o["holes"]),
                            part_category=part, part_thickness=thk, ceiling_type=ctype, fmt=fmt)
    if element == "room":
        mode = u"{}".format(o["mode"]).lower()
        if mode not in ("both", "rooms", "lines"):
            raise PlanError(u"room mode must be both, rooms or lines.")
        make_lines = mode != "rooms"
        view_item = None
        if make_lines:
            views = cr.get_plan_views(doc, lv["id"])
            if views:
                view_item = find_named(views, o["view"], u"floor plan of {}".format(level_name))
        return plan_room(doc, layers, level, level_name, make_lines, mode != "lines", view_item,
                         u"{}".format(o["name"] or u"").strip(), bool(o["skip_existing"]))
    if element == "column":
        from Snippets._units import paper_unit, project_length_unit
        size_unit = paper_unit(project_length_unit(doc))
        syms = look.symbols("OST_StructuralColumns" if o["structural"] else "OST_Columns")
        rect = find_named(syms, wanted, "column type", ("rectangular", "concrete", "square"))
        rnd = find_named(syms, o["round_type"], "round column type") if o["round_type"] else None
        top = None
        if o["top_level"]:
            above = [x for x in look.levels() if x["elevation"] > lv["elevation"] + 1e-6]
            top = find_named(above, o["top_level"], u"level above {}".format(level_name))
        return plan_column(doc, layers, level, level_name, _mm(o, "rounding_mm"),
                           _mm(o, "height_mm"), rect, rnd, top, bool(o["structural"]),
                           bool(o["match_size"]), bool(o["rotate"]), size_unit=size_unit,
                           fmt=fmt)
    if element == "beam":
        part = _part(o)
        sym = None if part else find_named(look.symbols("OST_StructuralFraming"), wanted,
                                           "structural framing type")
        w_min, w_max = _mm(o, "min_width_mm"), _mm(o, "max_width_mm")
        if w_min <= 0 or w_max < w_min:
            raise PlanError(u"Beam widths need 0 < min_width_mm <= max_width_mm.")
        return plan_beam(doc, layers, level, level_name, _mm(o, "offset_mm"), w_min, w_max,
                         bool(o["merge"]), part_category=part, beam_type=sym,
                         match=bool(o["match_size"]), fmt=fmt)
    if element == "lines":
        style = find_named(look.get("styles", lambda: cr.get_line_styles(doc)), wanted,
                           "line style") if wanted else None
        view, z = None, level.Elevation + _mm(o, "offset_mm")
        if o["view_id"] is not None:
            from Snippets._compat import make_eid
            view = doc.GetElement(make_eid(int(o["view_id"])))
            vt = cr.DB.ViewType
            if view is None or view.ViewType not in (vt.FloorPlan, vt.CeilingPlan,
                                                     vt.EngineeringPlan, vt.AreaPlan,
                                                     vt.DraftingView):
                raise PlanError(u"Detail lines need a plan or drafting view; view_id {} is not "
                                u"one. Omit view_id for model lines.".format(o["view_id"]))
            z = 0.0
            try:
                if view.GenLevel is not None:
                    z = view.GenLevel.Elevation
            except Exception:
                pass
        return plan_lines(doc, layers, level_name, view, z, bool(o["merge"]), style)
    cat = geo.mep_category(element)
    tp = find_named(look.get(element, lambda: cr.get_mep_types(doc, element)), wanted,
                    cat["type_label"])
    system_id = None
    if cat["system"]:
        system_id = find_named(
            look.get(element + "_sys", lambda: cr.get_mep_systems(doc, element)), o["system"],
            u"{} system type".format(cat["label"].lower()))["id"]
    double = bool(cat["double"] and o["double"])
    width = None if double else _mm(o, "width_mm")
    height = _mm(o, "height_mm") if cat["height"] else None
    return plan_mep(doc, layers, level, level_name, element, tp, system_id, double, width,
                    height, _mm(o, "offset_mm"), bool(o["merge"]), bool(o["elbows"]), fmt=fmt)
