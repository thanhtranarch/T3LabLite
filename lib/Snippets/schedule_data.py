# -*- coding: utf-8 -*-
"""
schedule_data — round-trip schedule VALUES through Excel.

Revit can only export a schedule as delimited text (File > Export > Reports >
Schedule) and cannot read anything back. This module fills that gap:

  export_xlsx()     one worksheet per schedule; column A = Element ID, then one
                    column per visible field, values as Revit displays them
  read_changes()    read an edited workbook, match each worksheet to its
                    schedule by name, list every cell that differs from the
                    model (only editable fields of elements still in it)
  apply_changes()   write the chosen changes in ONE transaction (one Ctrl+Z)

Display values round-trip: lengths are written as Revit shows them (2400 for
2400 mm) and read back with SetValueString, so units stay right.

The Revit-free helpers (sheet naming, the xlsx writer, the diff) are tested in
`dev/test_excel_table.py`.

Author: Tran Tien Thanh
"""

import zipfile
from xml.etree import ElementTree as ET

from Snippets import xlsx_table as xt

EID_HEADER = 'Element ID'
_EID_ALIASES = ('element id', 'elementid', 'element_id', 'id')


# ════════════════════════════════════════════════════════════════════
# REVIT-FREE HELPERS
# ════════════════════════════════════════════════════════════════════

def sheet_names_for(schedule_names):
    """Excel-safe, unique, deterministic worksheet name per schedule name."""
    out, used = {}, set()
    for name in schedule_names:
        base = name
        for ch in '/\\*?[]:':
            base = base.replace(ch, '_')
        base = base.strip("'")[:31] or 'Schedule'
        cand, i = base, 2
        while cand.lower() in used:
            suffix = '~{}'.format(i)
            cand = base[:31 - len(suffix)] + suffix
            i += 1
        used.add(cand.lower())
        out[name] = cand
    return out


def _is_number(text):
    """True only for numbers Excel would show back exactly: '900', '1200.5'.

    '007', '1_000', '1e5', 'Infinity', 15-digit IDs stay text — written as a
    number Excel would reformat them and the import would report a change.
    """
    try:
        num = float(text)
    except (TypeError, ValueError):
        return False
    if num != num or num in (float('inf'), float('-inf')):
        return False
    return xt.format_value(num, 'General') == text


def write_xlsx(path, sheets):
    """sheets: [{'name', 'headers', 'rows', 'locked': set(col index)}].

    Header row bold on grey, frozen; Element ID column and read-only fields
    shaded so the user sees what will not be imported; column widths fitted.
    """
    ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    rel_t = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
    shared, index = [], {}

    def si(text):
        text = '' if text is None else str(text)
        if text not in index:
            index[text] = len(shared)
            shared.append(text)
        return index[text]

    def tostr(el):
        return ET.tostring(el, encoding='unicode')

    # styles: 0 normal · 1 header · 2 locked (grey text on light fill)
    styles = (
        '<styleSheet xmlns="{ns}">'
        '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/><name val="Calibri"/></font>'
        '<font><sz val="11"/><color rgb="FF6B6B6B"/><name val="Calibri"/></font></fonts>'
        '<fills count="4"><fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="gray125"/></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9D9D9"/></patternFill></fill>'
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill></fills>'
        '<borders count="2"><border><left/><right/><top/><bottom/></border>'
        '<border><left/><right/><top/><bottom style="thin"><color rgb="FF808080"/></bottom></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/>'
        '<xf numFmtId="0" fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        '</styleSheet>').format(ns=ns)

    sheet_xml = []
    for sd in sheets:
        headers, rows = sd['headers'], sd['rows']
        locked = sd.get('locked', set())
        ws = ET.Element('worksheet', {'xmlns': ns})
        views = ET.SubElement(ET.SubElement(ws, 'sheetViews'), 'sheetView', {'workbookViewId': '0'})
        ET.SubElement(views, 'pane', {'ySplit': '1', 'topLeftCell': 'A2',
                                      'activePane': 'bottomLeft', 'state': 'frozen'})
        cols = ET.SubElement(ws, 'cols')
        for c, h in enumerate(headers):
            longest = max([len(str(h))] + [len(str(r[c])) for r in rows[:500] if c < len(r)])
            ET.SubElement(cols, 'col', {'min': str(c + 1), 'max': str(c + 1),
                                        'width': str(min(60, max(8, longest + 2))),
                                        'customWidth': '1'})
        data = ET.SubElement(ws, 'sheetData')
        row_el = ET.SubElement(data, 'row', {'r': '1'})
        for c, h in enumerate(headers):
            cel = ET.SubElement(row_el, 'c', {'r': '{}1'.format(xt.col_letter(c + 1)),
                                              't': 's', 's': '1'})
            ET.SubElement(cel, 'v').text = str(si(h))
        for r, row in enumerate(rows):
            rn = r + 2
            row_el = ET.SubElement(data, 'row', {'r': str(rn)})
            for c, val in enumerate(row):
                text = '' if val is None else str(val)
                attrs = {'r': '{}{}'.format(xt.col_letter(c + 1), rn)}
                if c in locked:
                    attrs['s'] = '2'
                if _is_number(text):
                    cel = ET.SubElement(row_el, 'c', attrs)
                    ET.SubElement(cel, 'v').text = text
                elif text:
                    attrs['t'] = 's'
                    cel = ET.SubElement(row_el, 'c', attrs)
                    ET.SubElement(cel, 'v').text = str(si(text))
                elif c in locked:
                    ET.SubElement(row_el, 'c', attrs)
        sheet_xml.append(tostr(ws))

    wb = ET.Element('workbook', {'xmlns': ns, 'xmlns:r': ns_r})
    sh = ET.SubElement(wb, 'sheets')
    for i, sd in enumerate(sheets):
        ET.SubElement(sh, 'sheet', {'name': sd['name'], 'sheetId': str(i + 1),
                                    'r:id': 'rId{}'.format(i + 1)})
    sst = ET.Element('sst', {'xmlns': ns, 'count': str(len(shared)),
                             'uniqueCount': str(len(shared))})
    for text in shared:
        t = ET.SubElement(ET.SubElement(sst, 'si'), 't')
        t.text = text
        if text != text.strip():
            t.set('xml:space', 'preserve')

    rels = ['<Relationship Id="rId{}" Type="{}worksheet" Target="worksheets/sheet{}.xml"/>'.format(
        i + 1, rel_t, i + 1) for i in range(len(sheets))]
    n = len(sheets)
    rels.append('<Relationship Id="rId{}" Type="{}sharedStrings" Target="sharedStrings.xml"/>'.format(n + 1, rel_t))
    rels.append('<Relationship Id="rId{}" Type="{}styles" Target="styles.xml"/>'.format(n + 2, rel_t))
    ct = ['<Override PartName="/xl/worksheets/sheet{}.xml" ContentType="application/'
          'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(i + 1)
          for i in range(n)]
    decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', decl + (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            + ''.join(ct) + '</Types>'))
        zf.writestr('_rels/.rels', decl + (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="{}officeDocument" Target="xl/workbook.xml"/>'
            '</Relationships>').format(rel_t))
        zf.writestr('xl/workbook.xml', decl + tostr(wb))
        zf.writestr('xl/_rels/workbook.xml.rels', decl + (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + ''.join(rels) + '</Relationships>'))
        zf.writestr('xl/sharedStrings.xml', decl + tostr(sst))
        zf.writestr('xl/styles.xml', decl + styles)
        for i, xml in enumerate(sheet_xml):
            zf.writestr('xl/worksheets/sheet{}.xml'.format(i + 1), decl + xml)


def read_sheet_values(wb, sheet_name):
    """(headers, [(element_id, {header: value})]) from one worksheet.

    Raises XlsxError when the sheet has no Element ID column.
    """
    table = wb.read_table(sheet_name, None, include_hidden=True)
    rows = table.value_rows()
    if not rows:
        return [], []
    headers = [h.strip() for h in rows[0]]
    eid_col = next((i for i, h in enumerate(headers) if h.lower() in _EID_ALIASES), None)
    if eid_col is None:
        raise xt.XlsxError("Sheet '{}' has no '{}' column — export the schedule "
                           "from Schedule Manager first, edit, then import.".format(
                               sheet_name, EID_HEADER))
    out = []
    for row in rows[1:]:
        try:
            eid = int(float(row[eid_col]))
        except (ValueError, IndexError):
            continue
        out.append((eid, {headers[i]: row[i] for i in range(len(headers))
                          if i != eid_col and headers[i]}))
    return headers, out


def diff_rows(schedule_name, fields, current, imported):
    """Changes between live values and imported values.

    fields   — [{'name', 'can_edit', ...}] in column order
    current  — {element_id: [value per field]}
    imported — [(element_id, {field name: value})]
    """
    changes = []
    col = {f['name']: i for i, f in enumerate(fields)}
    for eid, values in imported:
        live = current.get(eid)
        if live is None:
            continue
        for name, new in values.items():
            i = col.get(name)
            if i is None or not fields[i].get('can_edit'):
                continue
            new = (new or '').strip()
            old = (live[i] or '').strip() if i < len(live) else ''
            if new != old and not _same_number(new, old):
                changes.append({'schedule': schedule_name, 'element_id': eid,
                                'field': name, 'field_index': i,
                                'old': old, 'new': new})
    return changes


def _same_number(a, b):
    """'2400' vs '2400.0' — Excel re-saves numbers its own way."""
    try:
        return a != '' and b != '' and abs(float(a) - float(b)) < 1e-9
    except ValueError:
        return False


# ════════════════════════════════════════════════════════════════════
# REVIT SIDE
# ════════════════════════════════════════════════════════════════════

def _db():
    from Autodesk.Revit import DB
    return DB


def list_schedules(doc):
    """[(name, ViewSchedule)] — regular schedules only, sorted by name."""
    DB = _db()
    out = []
    for sch in DB.FilteredElementCollector(doc).OfClass(DB.ViewSchedule).ToElements():
        try:
            if sch.IsTemplate or sch.IsTitleblockRevisionSchedule or sch.IsInternalKeynoteSchedule:
                continue
            out.append((sch.Name, sch))
        except Exception:
            pass
    out.sort(key=lambda p: p[0].lower())
    return out


def schedule_fields(schedule):
    """Visible fields: [{'name', 'param_id', 'can_edit', 'on_type'}]."""
    from Snippets._compat import eid_value
    DB = _db()
    out = []
    d = schedule.Definition
    for i in range(d.GetFieldCount()):
        f = d.GetField(i)
        if f.IsHidden:
            continue
        ftype = f.FieldType
        pid = eid_value(f.ParameterId)
        out.append({
            'name': f.GetName(),
            'param_id': pid if pid != -1 else None,
            # Only instance values are written back: a type value is shared by
            # every instance of the type, room / project-info fields belong to
            # another element, formula / count fields have no parameter.
            'can_edit': ftype == DB.ScheduleFieldType.Instance and pid != -1,
            'on_type': ftype == DB.ScheduleFieldType.ElementType,
        })
    return out


def _param_map(element):
    from Snippets._compat import eid_value
    out = {}
    try:
        for p in element.Parameters:
            out.setdefault(eid_value(p.Id), p)
    except Exception:
        pass
    return out


def _find_param(doc, elem, field, cache):
    """Parameter of `field` on the element (type first for type fields)."""
    from Snippets._compat import eid_value
    pid = field.get('param_id')
    if pid is not None:
        holders = [elem]
        try:
            typ = doc.GetElement(elem.GetTypeId())
        except Exception:
            typ = None
        if typ is not None:
            holders = [typ, elem] if field.get('on_type') else [elem, typ]
        for holder in holders:
            key = eid_value(holder.Id)
            if key not in cache:
                cache[key] = _param_map(holder)
            param = cache[key].get(pid)
            if param is not None:
                return param
    try:
        return elem.LookupParameter(field['name'])
    except Exception:
        return None


def _display(param):
    DB = _db()
    from Snippets._compat import eid_value
    try:
        val = param.AsValueString()
        if val:
            return val
        st = param.StorageType
        if st == DB.StorageType.String:
            return param.AsString() or ''
        if st == DB.StorageType.Integer:
            return str(param.AsInteger())
        if st == DB.StorageType.Double:
            return str(round(param.AsDouble(), 6))
        if st == DB.StorageType.ElementId:
            eid = param.AsElementId()
            if eid is not None and eid_value(eid) != -1:
                ref = param.Element.Document.GetElement(eid)
                if ref is not None:
                    return ref.Name
        return ''
    except Exception:
        return ''


def schedule_values(doc, schedule):
    """(fields, {element_id: [display value per field]}, [element_id in order])."""
    from Snippets._compat import eid_value
    DB = _db()
    fields = schedule_fields(schedule)
    values, order, cache = {}, [], {}
    for elem in DB.FilteredElementCollector(doc, schedule.Id).WhereElementIsNotElementType():
        eid = eid_value(elem.Id)
        row = []
        for f in fields:
            p = _find_param(doc, elem, f, cache)
            row.append(_display(p) if p is not None else '')
        values[eid] = row
        order.append(eid)
    return fields, values, order


def _sheet_map(doc, extra=()):
    """Worksheet name per schedule name, computed over EVERY schedule of the
    model so export (ticked ones) and import (all) agree on truncated names."""
    names = [n for n, _ in list_schedules(doc)]
    names += [n for n in extra if n not in names]
    return sheet_names_for(names)


def export_xlsx(doc, schedules, path):
    """Write schedules [(name, ViewSchedule)] to `path`. Returns row count."""
    names = _sheet_map(doc, [n for n, _ in schedules])
    sheets, total = [], 0
    for name, sch in schedules:
        fields, values, order = schedule_values(doc, sch)
        headers = [EID_HEADER] + [f['name'] for f in fields]
        rows = [[str(eid)] + values[eid] for eid in order]
        locked = {0} | {i + 1 for i, f in enumerate(fields) if not f['can_edit']}
        sheets.append({'name': names[name], 'headers': headers, 'rows': rows, 'locked': locked})
        total += len(rows)
    write_xlsx(path, sheets)
    return total


def read_changes(doc, path, schedules, fallback=None):
    """Match each worksheet of `path` to a schedule and diff it.

    schedules — [(name, ViewSchedule)] that may be matched (all of the model)
    fallback  — a schedule used for a single-sheet workbook whose sheet name
                matches no schedule (the user renamed the sheet)
    Returns (changes, report lines).
    """
    by_sheet = {}
    names = _sheet_map(doc, [n for n, _ in schedules])
    for name, sch in schedules:
        by_sheet[names[name].lower()] = (name, sch)
    changes, report = [], []
    with xt.Workbook(path) as wb:
        sheet_list = wb.sheet_names()
        for sheet in sheet_list:
            target = by_sheet.get(sheet.lower())
            if target is None and fallback is not None and len(sheet_list) == 1:
                target = fallback
            if target is None:
                report.append("Sheet '{}': no schedule with that name — skipped.".format(sheet))
                continue
            try:
                _, imported = read_sheet_values(wb, sheet)
            except xt.XlsxError as ex:
                report.append(str(ex))
                continue
            name, sch = target
            fields, values, _ = schedule_values(doc, sch)
            found = diff_rows(name, fields, values, imported)
            missing = sum(1 for eid, _ in imported if eid not in values)
            line = "'{}': {} change(s)".format(name, len(found))
            if missing:
                line += ", {} row(s) not in the schedule any more".format(missing)
            report.append(line + ".")
            for ch in found:
                ch['field_def'] = fields[ch['field_index']]
            changes.extend(found)
    return changes, report


def _set_value(param, value):
    """Write a display value into a parameter. Returns '' or an error."""
    DB = _db()
    if param.IsReadOnly:
        return "read-only"
    st = param.StorageType
    try:
        if st == DB.StorageType.String:
            param.Set(value)
        elif st == DB.StorageType.Integer:
            try:
                param.Set(int(float(value or 0)))
            except ValueError:
                if not param.SetValueString(value):     # "Yes" / enum display names
                    return "'{}' is not a valid value".format(value)
        elif st == DB.StorageType.Double:
            # Exported as AsValueString (project units); read it back the same way.
            if not value:
                param.Set(0.0)
            elif not param.SetValueString(value):
                return "'{}' is not a valid value".format(value)
        else:
            return "element references cannot be set from text"
    except Exception as ex:
        return str(ex)
    return ''


def apply_changes(doc, changes):
    """Write `changes` in one transaction.

    Returns (done, [error lines], [changes not written]).
    """
    from Snippets._compat import disposing, make_eid
    DB = _db()
    done, errors, failed, cache = 0, [], [], {}
    with disposing(DB.Transaction(doc, "T3Lab: Schedule Manager - Apply Excel Changes")) as t:
        t.Start()
        try:
            for ch in changes:
                elem = doc.GetElement(make_eid(ch['element_id']))
                if elem is None:
                    errors.append("ID {}: element no longer exists".format(ch['element_id']))
                    failed.append(ch)
                    continue
                param = _find_param(doc, elem, ch['field_def'], cache)
                if param is None:
                    errors.append("ID {}: '{}' not found".format(ch['element_id'], ch['field']))
                    failed.append(ch)
                    continue
                err = _set_value(param, ch['new'])
                if err:
                    errors.append("ID {}: {} — {}".format(ch['element_id'], ch['field'], err))
                    failed.append(ch)
                else:
                    done += 1
            status = t.Commit()
            if status is not None and status != DB.TransactionStatus.Committed:
                return 0, ["Revit rolled the changes back ({}).".format(status)], list(changes)
        except Exception as ex:
            t.RollBack()
            return 0, ["Nothing was changed — {}".format(ex)], list(changes)
    return done, errors, failed
