# -*- coding: utf-8 -*-
"""Read rebar sets into BVBS records: the Revit side of BVBS Export.

Shared by GUI/BVBSExportDialog.py (preview grid) and the export_bvbs tool
(core/tools/rebar_assembly.py). The writer is Snippets/_bvbs.py. Read-only:
no transaction. Revit is reached only through `doc` / elements passed in, so it
loads outside Revit and under IronPython 2.7. Test: dev/test_bvbs_writer.py.
"""

import re

from Snippets import _bvbs, _rebar
from Snippets._compat import bar_mass_per_metre, eid_value, make_eid, short_error, to_mm

DEFAULT_PLAN = "MODEL"
DEFAULT_REVISION = "a"
DEFAULT_GRADE = "B500B"

# Severity names are the dialog's row styles.
SEV_OK = "Success"
SEV_WARN = "Warning"
SEV_FAIL = "Danger"

PROGRESS_EVERY = 5            # progress callback every N sets while reading centre-lines
LENGTH_TOLERANCE_MM = 1.0     # sets sharing a mark must agree on length within this


def natural_key(text):
    return [int(part) if part.isdigit() else part.lower()
            for part in re.split(r"(\d+)", u"%s" % text)]


def _plural(count, word):
    return u"%d %s%s" % (count, word, u"" if count == 1 else u"s")


class BarItem(object):
    """One preview line: a distinct leg chain under one mark (``record`` set),
    or a set that cannot be written (``record`` None, ``reason`` says why)."""

    def __init__(self, mark, status, severity, record=None, diameter_mm=0.0, quantity=0,
                 length_mm=None, legs=u"", detail=u"", group=u"", ids=(), reason=u"",
                 selected=False):
        self.mark = mark
        self.status = status
        self.severity = severity
        self.record = record
        self.diameter_mm = diameter_mm
        self.quantity = quantity
        self.length_mm = length_mm
        self.legs = legs
        self.detail = detail
        self.group = group
        self.ids = list(ids)
        self.reason = reason
        self.selected = selected

    @property
    def exportable(self):
        return self.record is not None and self.severity != SEV_FAIL


# ── SCOPE ────────────────────────────────────────────────────────────────────

def assembly_maps(index, assemblies):
    """({assembly id: mark}, {assembly id: [RebarRecord members]})."""
    marks = dict((a.id, a.mark) for a in assemblies)
    by_assembly = {}
    for row in index.rows:
        if row.assembly_id is not None and row.assembly_id >= 0:
            by_assembly.setdefault(row.assembly_id, []).append(row)
    return marks, by_assembly


def assembly_rows(index, by_assembly, assembly):
    """Rebar that are members of the assembly plus rebar hosted by its members (A1)."""
    out, seen = [], set()
    candidates = list(by_assembly.get(assembly.id, ()))
    for member in assembly.members:
        candidates.extend(index.by_host.get(member, ()))
    for row in candidates:
        if row.id not in seen:
            seen.add(row.id)
            out.append(row)
    return out


def selection_rows(index, assemblies, by_assembly, ids):
    """Rebar for element ids: rebar itself, the rebar of hosts, the rebar of
    assemblies. Unique, in id order."""
    if index is None:
        return []
    by_id = dict((a.id, a) for a in assemblies)
    out, seen = [], set()

    def add(rows):
        for row in rows:
            if row.id not in seen:
                seen.add(row.id)
                out.append(row)

    for eid in ids:
        row = index.by_id.get(eid)
        if row is not None:
            add([row])
        if eid in index.by_host:
            add(index.by_host[eid])
        assembly = by_id.get(eid)
        if assembly is not None:
            add(assembly_rows(index, by_assembly, assembly))
    return out


def partition_counts(index):
    """{partition: rebar sets} of the shape-able rebar."""
    counts = {}
    for row in index.rows:
        if row.kind == "Rebar" and row.partition:
            counts[row.partition] = counts.get(row.partition, 0) + 1
    return counts


def bar_sets(rows_in_scope, only_shape_driven):
    """(Rebar sets to read, notes on what was left out)."""
    sets = [r for r in rows_in_scope if r.kind == "Rebar"]
    others = len(rows_in_scope) - len(sets)
    hidden = 0
    if only_shape_driven:
        hidden = sum(1 for r in sets if not r.is_shape_driven)
        sets = [r for r in sets if r.is_shape_driven]
    notes = []
    if others:
        notes.append(
            u"%s (area, path, fabric, coupler or in-system rebar) are outside BVBS BF2D and "
            u"are not read." % _plural(others, u"element"))
    if hidden:
        notes.append(
            u"%s hidden by “Shape-driven rebar only”." % _plural(hidden, u"free-form set"))
    return sets, notes


# ── READ EVERY BAR [REVIT] ───────────────────────────────────────────────────

def existing_positions(element):
    try:
        count = int(element.NumberOfBarPositions)
    except Exception:
        count = 0
    if count <= 0:
        return [0]
    out = []
    for position in range(count):
        try:
            if element.DoesBarExistAtPosition(position):
                out.append(position)
        except Exception:
            out.append(position)
    return out or [0]


def shape_at(element, position, diameter_mm):
    """BarShape of one bar position; an unreadable centre-line is a skip reason."""
    try:
        curves = _rebar.centerline_points(element, position)
    except Exception as exc:
        return _bvbs.BarShape(reason=u"centre-line unreadable", error=short_error(exc))
    return _bvbs.chain_to_shape(curves, diameter_mm)


def group_with_failures(shapes):
    """Distinct good shapes with their counts, then one entry per failure reason."""
    groups = _bvbs.group_distinct([s for s in shapes if s.ok])
    failed, order = {}, []
    for shape in shapes:
        if shape.ok:
            continue
        if shape.reason not in failed:
            failed[shape.reason] = [shape, 0]
            order.append(shape.reason)
        failed[shape.reason][1] += 1
    groups.extend((failed[reason][0], failed[reason][1]) for reason in order)
    return groups


def read_shapes(element, diameter_mm, quantity):
    """([(BarShape, count)], positions read, positions in the set) for a rebar set.

    A small set is read bar by bar. A big one is sampled (first / middle /
    last, R11) and counted as uniform when the samples agree; when they
    differ every bar is read and grouped, so a set with varying lengths
    becomes one record per distinct shape with its own count.
    """
    existing = existing_positions(element)
    sample = _bvbs.sample_positions(existing)
    shapes = [shape_at(element, p, diameter_mm) for p in sample]
    agree = len(set((s.key(), s.reason) for s in shapes)) == 1
    if len(sample) < len(existing) and not agree:
        shapes = [shape_at(element, p, diameter_mm) for p in existing]
        return group_with_failures(shapes), len(existing), len(existing)
    if len(sample) < len(existing):
        count = quantity if quantity and quantity > 0 else len(existing)
        return [(shapes[0], count)], len(sample), len(existing)
    return group_with_failures(shapes), len(existing), len(existing)


def revit_length_mm(element):
    """Per-bar length Revit schedules ("Bar Length"), mm; None when unreadable."""
    try:
        from Autodesk.Revit.DB import BuiltInParameter
        built_in = getattr(BuiltInParameter, "REBAR_ELEM_LENGTH", None)
        if built_in is None:
            return None
        param = element.get_Parameter(built_in)
        if param is None or not param.HasValue:
            return None
        value = to_mm(param.AsDouble())
        return value if value > 0 else None
    except Exception:
        return None


def bar_type_info(element, bar_type, cache):
    """(mandrel mm or None, kg/m or None, weight source or None), cached per type + style."""
    style = getattr(element, "Style", None)
    style_name = u"%s" % style if style is not None else u""
    key = (eid_value(bar_type.Id) if bar_type is not None else -1, style_name)
    if key not in cache:
        roll = None
        if bar_type is not None:
            roll = _rebar.mandrel_diameter_mm(bar_type, style)
        value, source = bar_mass_per_metre(bar_type)
        cache[key] = (roll, value, source)
    return cache[key]


def skip_item(row, mark, reason, detail=u""):
    return BarItem(mark or (u"(id %d)" % row.id), u"Skipped: %s" % reason, SEV_FAIL,
                   diameter_mm=row.diameter_mm, quantity=max(int(row.quantity or 1), 1),
                   detail=detail or (u"Element %d — %s." % (row.id, reason)),
                   ids=[row.id], reason=reason)


def read_one(doc, row, group, cache):
    """(entries, skip items) for one rebar set.

    A set whose bars differ becomes one entry per distinct shape; bars that
    cannot be written come back as skip items, never silently dropped.
    """
    mark = _bvbs.build_mark(row.partition, row.number, row.mark)
    if not row.diameter_mm or row.diameter_mm <= 0:
        return [], [skip_item(row, mark, u"no diameter")]
    if not mark:
        return [], [skip_item(
            row, u"", u"no mark",
            u"Element %d has no Rebar Number and no Schedule Mark — number the "
            u"reinforcement in Revit first." % row.id)]
    element = doc.GetElement(make_eid(row.id))
    if element is None:
        return [], [skip_item(row, mark, u"element missing")]

    bar_type = _rebar.bar_type_of(doc, element)
    roll, rev_value, rev_source = bar_type_info(element, bar_type, cache)
    roll_assumed = not roll or roll <= 0
    if roll_assumed:
        roll = 4.0 * row.diameter_mm       # BVBS standard mandrel: 4 x diameter
    revit_length = revit_length_mm(element)

    groups, read, total = read_shapes(element, row.diameter_mm, row.quantity)
    bars_in_groups = sum(count for _, count in groups)
    entries, skips = [], []
    for shape, count in groups:
        if not shape.ok:
            detail = u"Element %d — %s" % (row.id, shape.reason)
            if shape.error:
                detail += u": %s" % shape.error
            if len(groups) > 1:
                detail += u" (%d of %d bars)" % (count, bars_in_groups)
            skip = skip_item(row, mark, shape.reason, detail + u".")
            skip.quantity = count
            skips.append(skip)
            continue
        entries.append({
            "mark": mark, "group": group, "diameter": row.diameter_mm, "shape": shape,
            "count": count, "ids": [row.id], "roll": roll, "roll_assumed": roll_assumed,
            "rev_value": rev_value, "rev_source": rev_source,
            "lengths": [revit_length] if (revit_length and len(groups) == 1) else [],
            "read": read, "total": total})
    return entries, skips


def merge_entries(entries):
    """One entry per (assembly, mark, diameter, shape): quantities add up, as on a schedule."""
    merged, order = {}, []
    for entry in entries:
        key = (entry["group"], entry["mark"], int(round(entry["diameter"] * 10)),
               entry["shape"].key())
        if key not in merged:
            merged[key] = entry
            order.append(key)
            continue
        target = merged[key]
        target["count"] += entry["count"]
        target["ids"].extend(entry["ids"])
        target["lengths"].extend(entry["lengths"])
        target["read"] += entry["read"]
        target["total"] += entry["total"]
    return [merged[key] for key in order]


def entry_item(entry, reused, weight_table, weight_counts):
    """BarItem + BarRecord for one merged entry; counts its weight source."""
    shape = entry["shape"]
    diameter = entry["diameter"]
    lengths = entry["lengths"]
    revit_length = None
    if lengths and max(lengths) - min(lengths) <= LENGTH_TOLERANCE_MM:
        revit_length = sum(lengths) / len(lengths)
    length_ok, length_note = _bvbs.length_check(
        revit_length, shape.outer_sum, shape.bends, diameter, entry["roll"])
    if revit_length and length_ok:
        length_mm, length_source = revit_length, u"Revit bar length"
    else:
        length_mm, length_source = shape.outer_sum, u"sum of the outer legs"

    kg_per_m, source = _bvbs.resolve_weight(
        diameter, entry["rev_value"], entry["rev_source"], weight_table)
    weight = _bvbs.bar_weight_kg(length_mm, kg_per_m)
    weight_counts[source] = weight_counts.get(source, 0) + entry["count"]

    record = _bvbs.BarRecord(
        entry["mark"], diameter, entry["count"], shape.segments,
        weight_kg=weight, roll_diameter_mm=entry["roll"], total_length_mm=length_mm,
        group=entry["group"])
    status, severity, selected = u"Ready", SEV_OK, True
    if not length_ok:
        status, severity, selected = u"Check legs", SEV_WARN, False
    elif reused:
        status, severity, selected = u"Same mark, other shape", SEV_WARN, False

    ids = sorted(set(entry["ids"]))
    parts = [u"Element%s %s" % (u"" if len(ids) == 1 else u"s",
                                u", ".join(str(i) for i in ids[:5])
                                + (u" +%d more" % (len(ids) - 5) if len(ids) > 5 else u"")),
             u"centre-line legs %s" % _bvbs.legs_text(shape.centre_legs),
             u"outer legs %s" % _bvbs.legs_text(shape.segments),
             u"length %d mm (%s)" % (int(round(length_mm)), length_source),
             u"Ø%d, mandrel %d mm%s" % (
                 int(round(diameter)), int(round(entry["roll"])),
                 u" (assumed 4 x Ø: the bar type has no bend diameter)"
                 if entry["roll_assumed"] else u""),
             u"%.3f kg/m from %s" % (kg_per_m, _bvbs.WEIGHT_SOURCE_LABELS.get(source, source))]
    if entry["read"] < entry["total"]:
        parts.append(u"read %d of %d bars, the rest assumed identical"
                     % (entry["read"], entry["total"]))
    if not length_ok:
        parts.append(length_note + u" — compare with the bending schedule, then tick it")
    if reused:
        parts.append(u"this mark has more than one shape in the model — check the "
                     u"numbering, then tick it")
    return BarItem(entry["mark"], status, severity, record=record, diameter_mm=diameter,
                   quantity=entry["count"], length_mm=length_mm,
                   legs=_bvbs.legs_text(shape.segments), detail=u" · ".join(parts),
                   group=entry["group"], ids=ids, selected=selected)


def read_sets(doc, sets, forced_group, asm_marks, weight_table, weight_counts, progress=None):
    """([BarItem] ready first then skipped, stopped) for the rebar sets.

    ``progress(number, total)`` is called every PROGRESS_EVERY sets; False stops
    (stopped = True, the items cover what was read). Never raises per set.
    """
    entries, skips = [], []
    cache = {}
    stopped = False
    for number, row in enumerate(sets, 1):
        if progress is not None and number % PROGRESS_EVERY == 1:
            if progress(number, len(sets)) is False:
                stopped = True
                break
        group = forced_group if forced_group is not None else asm_marks.get(row.assembly_id, u"")
        try:
            found, skipped = read_one(doc, row, group, cache)
        except Exception as exc:
            found, skipped = [], [skip_item(
                row, _bvbs.build_mark(row.partition, row.number, row.mark),
                u"unexpected error", u"Element %d — %s" % (row.id, short_error(exc)))]
        entries.extend(found)
        skips.extend(skipped)

    merged = merge_entries(entries)
    shapes_per_mark = {}
    for entry in merged:
        shapes_per_mark.setdefault((entry["group"], entry["mark"]), set()).add(
            (int(round(entry["diameter"] * 10)), entry["shape"].key()))
    ready = [entry_item(e, len(shapes_per_mark[(e["group"], e["mark"])]) > 1,
                        weight_table, weight_counts)
             for e in merged]
    ready.sort(key=lambda r: (natural_key(r.group), natural_key(r.mark)))
    skips.sort(key=lambda r: (natural_key(r.mark), r.reason))
    return ready + skips, stopped


def apply_fields(records, project, plan, revision, grade):
    """Header fields on every record; returns the plan."""
    for record in records:
        record.project = project
        record.plan = plan
        record.index = revision
        record.grade = grade
    return plan


def skipped_counts(items):
    """{reason: bars} of the items that cannot be written."""
    counts = {}
    for item in items:
        if not item.exportable:
            reason = item.reason or u"skipped"
            counts[reason] = counts.get(reason, 0) + item.quantity
    return counts
