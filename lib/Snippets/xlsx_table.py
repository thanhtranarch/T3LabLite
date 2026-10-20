# -*- coding: utf-8 -*-
"""
xlsx_table — read one Excel range WITH its formatting, no Office required.

Pure Python (zipfile + ElementTree): no Revit, no Excel interop, so it runs on
any machine and is unit-tested outside Revit (`python dev/test_excel_table.py`).

What it reads, per cell: displayed text (number formats applied), font (name,
size, bold, italic, underline, colour), solid fill colour, the four borders
(weight + colour), horizontal / vertical alignment, wrap. Per sheet: merged
cells, column widths, row heights, hidden rows / columns, defined names,
Excel Tables ("Format as Table" ranges) and the print area.

Not read (limits, on purpose): images and charts, conditional formatting,
the colours of "Format as Table" styles that live only in the table style,
rotated text, formulas (the value Excel cached last time it saved is used).

`build_layout()` turns a Table into paper-space geometry (inches, y down):
cell rectangles, fills, merged border segments and text boxes — shared by
the drafting-view and the schedule renderers in `excel_table_revit.py`.

Author: Tran Tien Thanh
"""

import colorsys
import datetime
import hashlib
import json
import os
import re
import zipfile
from xml.etree import ElementTree as ET

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS_REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

USED_RANGE = "Used range"

DEFAULT_FONT = {'name': 'Calibri', 'size': 11.0, 'bold': False, 'italic': False,
                'underline': False, 'color': (0, 0, 0)}

# Excel's 64-colour legacy palette (indexed="n"); 64/65 = system fg/bg.
_INDEXED = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
    '000000', 'FFFFFF',
]

# Office 2013+ default theme — used when the workbook has no theme part.
_DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31',
                  'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72']

BORDER_WEIGHT = {
    'hair': 1, 'thin': 1, 'dotted': 1, 'dashed': 1, 'dashDot': 1, 'dashDotDot': 1,
    'medium': 2, 'mediumDashed': 2, 'mediumDashDot': 2, 'mediumDashDotDot': 2,
    'slantDashDot': 2, 'thick': 3, 'double': 3,
}

_BUILTIN_FMT = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
    9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'm/d/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy',
    18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss',
    22: 'm/d/yyyy h:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
}


class XlsxError(Exception):
    """A readable message for the user: what is wrong with the workbook."""


# ════════════════════════════════════════════════════════════════════
# CELL REFERENCES
# ════════════════════════════════════════════════════════════════════

_REF_RE = re.compile(r'^\$?([A-Za-z]{1,3})\$?(\d+)$')


def col_letter(col):
    """1-based column index -> 'A', 'B', ..., 'AA'."""
    out = ""
    while col > 0:
        col -= 1
        out = chr(col % 26 + 65) + out
        col //= 26
    return out


def col_number(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def parse_ref(ref):
    """'B3' -> (row 3, col 2). Raises XlsxError."""
    m = _REF_RE.match((ref or '').strip())
    if not m:
        raise XlsxError("'{}' is not a cell reference (use A1 style).".format(ref))
    return int(m.group(2)), col_number(m.group(1))


def parse_range(text):
    """'A1:F20' / 'Sheet1'!A1:F20 / B3 -> (r1, c1, r2, c2), normalised."""
    text = (text or '').strip()
    if '!' in text:
        text = text.rsplit('!', 1)[1]
    parts = text.split(':')
    if len(parts) == 1:
        r, c = parse_ref(parts[0])
        return r, c, r, c
    if len(parts) != 2:
        raise XlsxError("'{}' is not a range (use A1:F20).".format(text))
    r1, c1 = parse_ref(parts[0])
    r2, c2 = parse_ref(parts[1])
    return min(r1, r2), min(c1, c2), max(r1, r2), max(c1, c2)


def range_text(r1, c1, r2, c2):
    return "{}{}:{}{}".format(col_letter(c1), r1, col_letter(c2), r2)


# ════════════════════════════════════════════════════════════════════
# COLOURS
# ════════════════════════════════════════════════════════════════════

def _hex_rgb(hex6):
    hex6 = hex6[-6:]
    return int(hex6[0:2], 16), int(hex6[2:4], 16), int(hex6[4:6], 16)


def _apply_tint(rgb, tint):
    if not tint:
        return rgb
    h, l, s = colorsys.rgb_to_hls(*[v / 255.0 for v in rgb])
    if tint < 0:
        l = l * (1.0 + tint)
    else:
        l = l * (1.0 - tint) + tint
    r, g, b = colorsys.hls_to_rgb(h, max(0.0, min(1.0, l)), s)
    return int(round(r * 255)), int(round(g * 255)), int(round(b * 255))


def _color(el, theme):
    """<color rgb|theme|indexed tint/> -> (r, g, b) or None (automatic)."""
    if el is None:
        return None
    if el.get('auto') == '1':
        return None
    tint = 0.0
    try:
        tint = float(el.get('tint') or 0)
    except ValueError:
        pass
    rgb = None
    if el.get('rgb'):
        try:
            rgb = _hex_rgb(el.get('rgb'))
        except ValueError:
            return None
    elif el.get('theme') is not None:
        try:
            idx = int(el.get('theme'))
        except ValueError:
            return None
        # Excel swaps the first two pairs: theme 0 = lt1, 1 = dk1, 2 = lt2, 3 = dk2.
        order = [1, 0, 3, 2] + list(range(4, 12))
        if 0 <= idx < len(order) and order[idx] < len(theme):
            rgb = _hex_rgb(theme[order[idx]])
    elif el.get('indexed') is not None:
        try:
            idx = int(el.get('indexed'))
        except ValueError:
            return None
        if idx == 64:
            return None                  # system foreground = automatic
        if 0 <= idx < len(_INDEXED):
            rgb = _hex_rgb(_INDEXED[idx])
    if rgb is None:
        return None
    return _apply_tint(rgb, tint)


def _read_theme(zf, path):
    try:
        root = ET.fromstring(zf.read(path))
    except (KeyError, ET.ParseError):
        return list(_DEFAULT_THEME)
    scheme = root.find('.//{%s}clrScheme' % NS_A)
    if scheme is None:
        return list(_DEFAULT_THEME)
    names = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3',
             'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']
    out = []
    for i, name in enumerate(names):
        node = scheme.find('{%s}%s' % (NS_A, name))
        val = None
        if node is not None:
            srgb = node.find('{%s}srgbClr' % NS_A)
            sysc = node.find('{%s}sysClr' % NS_A)
            if srgb is not None:
                val = srgb.get('val')
            elif sysc is not None:
                val = sysc.get('lastClr')
        out.append(val or _DEFAULT_THEME[i])
    return out


# ════════════════════════════════════════════════════════════════════
# NUMBER FORMATS
# ════════════════════════════════════════════════════════════════════

def _split_sections(code):
    out, cur, quote = [], '', False
    i = 0
    while i < len(code):
        ch = code[i]
        if ch == '"':
            quote = not quote
            cur += ch
        elif ch == '\\' and i + 1 < len(code):
            cur += code[i:i + 2]
            i += 1
        elif ch == ';' and not quote:
            out.append(cur)
            cur = ''
        else:
            cur += ch
        i += 1
    out.append(cur)
    return out


def _strip_brackets(section):
    """Drop [Red] / [<=100] / [$-409]; keep the currency symbol of [$€-407]."""
    def repl(m):
        body = m.group(1)
        if body.startswith('$'):
            return body[1:].split('-')[0]
        if body.lower() in ('h', 'hh', 'm', 'mm', 's', 'ss'):
            return '[' + body + ']'      # elapsed time, kept for the date path
        return ''
    return re.sub(r'\[([^\]]*)\]', repl, section)


def _unquoted(section):
    """The format with quoted / escaped literals removed (token detection)."""
    s = re.sub(r'"[^"]*"', '', section)
    return re.sub(r'\\.', '', s)


def _is_date_format(section):
    s = _unquoted(_strip_brackets(section)).lower()
    s = s.replace('general', '')
    return bool(re.search(r'[dmyhs]', s)) and not re.search(r'[0#?]', s.replace('.0', ''))


def _general(value):
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    text = '%.10g' % value
    if 'e' in text:
        mant, exp = text.split('e')
        return '{}E{}{:02d}'.format(mant, '+' if int(exp) >= 0 else '-', abs(int(exp)))
    return text


def _literal_text(section, text=''):
    """A non-numeric section as Excel shows it; unquoted '@' becomes `text`.
    `_x` pads with a space and `*x` (fill) is dropped, as in _literal_parts
    (2026-10-20: accounting `_(@_)` printed "_(TEXT_)" in Revit)."""
    out, i = [], 0
    section = _strip_brackets(section)
    while i < len(section):
        ch = section[i]
        if ch == '"':
            j = section.find('"', i + 1)
            j = len(section) if j < 0 else j
            out.append(section[i + 1:j])
            i = j + 1
            continue
        if ch in '\\_*' and i + 1 < len(section):
            out.append({'\\': section[i + 1], '_': ' ', '*': ''}[ch])
            i += 2
            continue
        out.append(text if ch == '@' else ch)
        i += 1
    return ''.join(out)


def _literal_parts(section):
    """Split a numeric section into (prefix, number pattern, suffix) literals."""
    i = 0
    tokens = []          # (is_literal, text)
    while i < len(section):
        ch = section[i]
        if ch == '"':
            j = section.find('"', i + 1)
            j = len(section) if j < 0 else j
            tokens.append((True, section[i + 1:j]))
            i = j + 1
            continue
        if ch == '\\' and i + 1 < len(section):
            tokens.append((True, section[i + 1]))
            i += 2
            continue
        if ch == '_' and i + 1 < len(section):
            tokens.append((True, ' '))
            i += 2
            continue
        if ch == '*' and i + 1 < len(section):
            i += 2
            continue
        if ch in 'Ee' and section[i + 1:i + 2] in ('+', '-') and section[i + 1:i + 2]:
            tokens.append((False, ch + section[i + 1]))
            i += 2
            continue
        if ch in '0#?.,%':
            tokens.append((False, ch))
        else:
            tokens.append((True, ch))
        i += 1
    first = next((k for k, t in enumerate(tokens) if not t[0] and t[1] in '0#?'), None)
    if first is None:
        return ''.join(t[1] for t in tokens), '', ''
    last = max(k for k, t in enumerate(tokens) if not t[0])
    prefix = ''.join(t[1] for t in tokens[:first])
    # a '.' or ',' right before the first digit belongs to the pattern
    pattern = ''.join(t[1] for t in tokens[first:last + 1] if not t[0])
    suffix = ''.join(t[1] for t in tokens[last + 1:])
    # literals in the middle of the pattern (rare: "0 m2") go to the suffix
    middle = ''.join(t[1] for t in tokens[first:last + 1] if t[0])
    return prefix, pattern, middle + suffix


def _round_half_up(value, decimals):
    """Excel rounds half away from zero on the shortest decimal repr:
    2.5 -> '3', 0.125 -> '0.13' (Python's format gives '2' and '0.12')."""
    from decimal import Decimal, ROUND_HALF_UP
    q = Decimal(1).scaleb(-decimals)
    return str(Decimal(repr(float(value))).quantize(q, rounding=ROUND_HALF_UP))


def _format_numeric(value, section):
    prefix, pattern, suffix = _literal_parts(section)
    if not pattern:
        return prefix + suffix if (prefix or suffix) else _general(value)
    percent = (prefix + pattern + suffix).count('%')
    value = value * (100 ** percent)
    if 'E' in pattern.upper():
        mant = pattern.upper().split('E')[0].replace(',', '')
        decimals = len(mant.split('.')[1]) if '.' in mant else 0
        text = '{:.{}E}'.format(value, decimals)
        m, e = text.split('E')
        return prefix + '{}E{}{:02d}'.format(m, '+' if int(e) >= 0 else '-', abs(int(e))) + suffix
    # trailing commas scale by 1000 each ("#,##0," = thousands)
    core = pattern.replace('%', '')
    while core.endswith(','):
        value /= 1000.0
        core = core[:-1]
    int_part, _, dec_part = core.partition('.')
    decimals = len(dec_part)
    min_dec = dec_part.count('0')
    thousands = ',' in int_part
    min_int = int_part.count('0')
    text = _round_half_up(abs(value), decimals)
    whole, _, frac = text.partition('.')
    if decimals > min_dec:                     # '#' decimals are optional
        frac = frac.rstrip('0')
        if len(frac) < min_dec:
            frac = frac + '0' * (min_dec - len(frac))
    if whole == '0' and min_int == 0:
        whole = ''
    if len(whole) < min_int:
        whole = '0' * (min_int - len(whole)) + whole
    if thousands and whole:
        whole = '{:,}'.format(int(whole))
    out = whole + ('.' + frac if frac else '')
    if pattern.endswith('%'):
        out += '%' * pattern.count('%')
    return prefix + out + suffix


_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
           'August', 'September', 'October', 'November', 'December']
_DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']


def _serial_to_datetime(value, date1904):
    base = datetime.datetime(1904, 1, 1) if date1904 else datetime.datetime(1899, 12, 30)
    if not date1904 and value < 60:
        value += 1                     # Excel's fictional 1900-02-29
    return base + datetime.timedelta(days=value)


def _format_date(value, section, date1904):
    try:
        dt = _serial_to_datetime(value, date1904)
    except (OverflowError, ValueError):
        return _general(value)
    sec = _strip_brackets(section)
    ampm = bool(re.search(r'am/pm|a/p', sec, re.I))
    tokens = re.findall(r'"[^"]*"|\\.|\[h+\]|\[m+\]|\[s+\]|am/pm|a/p|y+|m+|d+|h+|s+(?:\.0+)?|.', sec, re.I)
    out = []
    for i, tok in enumerate(tokens):
        low = tok.lower()
        if tok.startswith('"'):
            out.append(tok[1:-1])
        elif tok.startswith('\\'):
            out.append(tok[1:])
        elif low.startswith('[h'):
            out.append(str(int(value * 24)))
        elif low.startswith('[m'):
            out.append(str(int(value * 1440)))
        elif low.startswith('[s'):
            out.append(str(int(value * 86400)))
        elif low in ('am/pm', 'a/p'):
            pm = dt.hour >= 12
            out.append(('PM' if pm else 'AM') if low == 'am/pm' else ('P' if pm else 'A'))
        elif low[0] == 'y':
            out.append(str(dt.year) if len(low) > 2 else '%02d' % (dt.year % 100))
        elif low[0] == 'd':
            if len(low) >= 4:
                out.append(_DAYS[dt.weekday()])
            elif len(low) == 3:
                out.append(_DAYS[dt.weekday()][:3])
            else:
                out.append('%0*d' % (len(low), dt.day))
        elif low[0] == 'h':
            h = dt.hour
            if ampm:
                h = h % 12 or 12
            out.append('%0*d' % (min(len(low), 2), h))
        elif low[0] == 's':
            txt = '%02d' % dt.second if len(low.split('.')[0]) > 1 else str(dt.second)
            if '.' in low:
                frac = (dt.microsecond / 1e6)
                txt += ('%.*f' % (len(low.split('.')[1]), frac))[1:]
            out.append(txt)
        elif low[0] == 'm':
            prev = [t.lower() for t in tokens[:i] if re.match(r'[a-z\[]', t, re.I)]
            nxt = [t.lower() for t in tokens[i + 1:] if re.match(r'[a-z\[]', t, re.I)]
            minute = (prev and prev[-1][0] in 'h[') or (nxt and nxt[0][0] == 's')
            if minute and len(low) <= 2:
                out.append('%0*d' % (len(low), dt.minute))
            elif len(low) >= 5:
                out.append(_MONTHS[dt.month - 1][0])
            elif len(low) == 4:
                out.append(_MONTHS[dt.month - 1])
            elif len(low) == 3:
                out.append(_MONTHS[dt.month - 1][:3])
            else:
                out.append('%0*d' % (len(low), dt.month))
        else:
            out.append(tok)
    return ''.join(out)


def format_value(value, fmt_code, date1904=False):
    """Excel-style display text for a numeric cell value."""
    code = fmt_code or 'General'
    sections = _split_sections(code)
    if value < 0 and len(sections) >= 2:
        section, value = sections[1], -value
    elif value == 0 and len(sections) >= 3:
        section = sections[2]
    else:
        section = sections[0]
        if value < 0 and len(sections) == 1 and not _is_date_format(section):
            text = format_value(-value, section, date1904)
            return text if text.startswith('-') else '-' + text
    if not section.strip():
        return ''
    if 'general' in _unquoted(section).lower():
        idx = section.lower().find('general')
        return _literal_text(section[:idx]) + _general(value) + _literal_text(section[idx + 7:])
    if section.strip() == '@':
        return _general(value)
    if _is_date_format(section):
        return _format_date(value, section, date1904)
    return _format_numeric(value, _strip_brackets(section))


def format_text(text, fmt_code):
    """Text cells use the 4th section (or '@' placeholder) when there is one."""
    sections = _split_sections(fmt_code or 'General')
    section = sections[3] if len(sections) >= 4 else (sections[0] if '@' in sections[0] else None)
    if not section or '@' not in _unquoted(section):
        return text
    return _literal_text(section, text)


# ════════════════════════════════════════════════════════════════════
# WORKBOOK
# ════════════════════════════════════════════════════════════════════

class Workbook(object):
    """An opened .xlsx: sheets, names and styles. Use as a context manager."""

    def __init__(self, path):
        if not os.path.isfile(path):
            raise XlsxError("Excel file not found: {}".format(path))
        if not path.lower().endswith(('.xlsx', '.xlsm')):
            raise XlsxError("Only .xlsx / .xlsm workbooks are supported "
                            "(save the file as .xlsx in Excel).")
        self.path = path
        try:
            self.zf = zipfile.ZipFile(path, 'r')
        except (zipfile.BadZipfile, IOError, OSError) as ex:
            raise XlsxError("Cannot open '{}' — {}. If the file is open in Excel "
                            "with unsaved changes, save it first.".format(
                                os.path.basename(path), ex))
        try:
            self._load()
        except XlsxError:
            self.zf.close()
            raise
        except (KeyError, ET.ParseError) as ex:
            self.zf.close()
            raise XlsxError("'{}' is not a valid Excel workbook ({}).".format(
                os.path.basename(path), ex))

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.close()
        return False

    def close(self):
        try:
            self.zf.close()
        except Exception:
            pass

    # ── parts ─────────────────────────────────────────────────────────
    def _xml(self, part):
        return ET.fromstring(self.zf.read(part))

    def _rels(self, part):
        folder, name = part.rsplit('/', 1) if '/' in part else ('', part)
        rel_part = (folder + '/' if folder else '') + '_rels/' + name + '.rels'
        out = {}
        try:
            root = self._xml(rel_part)
        except KeyError:
            return out
        for rel in root.findall('{%s}Relationship' % NS_REL):
            target = rel.get('Target', '')
            if target.startswith('/'):
                target = target[1:]
            else:
                target = os.path.normpath(os.path.join(folder, target)).replace('\\', '/')
            out[rel.get('Id')] = (rel.get('Type', '').rsplit('/', 1)[-1], target)
        return out

    def _load(self):
        wb = self._xml('xl/workbook.xml')
        rels = self._rels('xl/workbook.xml')
        pr = wb.find('{%s}workbookPr' % NS)
        self.date1904 = pr is not None and pr.get('date1904') in ('1', 'true')

        self.sheets = []         # [(name, part, state)]
        for sh in wb.findall('.//{%s}sheet' % NS):
            rid = sh.get('{%s}id' % NS_R)
            part = rels.get(rid, (None, None))[1]
            if part:
                self.sheets.append((sh.get('name', ''), part, sh.get('state', 'visible')))
        if not self.sheets:
            raise XlsxError("The workbook has no worksheets.")

        # defined names: {name: (sheet or None, range text)}
        self.names = {}
        self.print_areas = {}
        for dn in wb.findall('.//{%s}definedName' % NS):
            name, ref = dn.get('name', ''), (dn.text or '').strip()
            if not ref or '#REF' in ref or ',' in ref:
                continue
            sheet = None
            if '!' in ref:
                sheet = ref.rsplit('!', 1)[0].strip("'").replace("''", "'")
            if name == '_xlnm.Print_Area' and dn.get('localSheetId') is not None:
                try:
                    sheet = self.sheets[int(dn.get('localSheetId'))][0]
                except (ValueError, IndexError):
                    continue
                self.print_areas[sheet] = ref.rsplit('!', 1)[-1].replace('$', '')
            elif not name.startswith('_xlnm'):
                self.names[name] = (sheet, ref.rsplit('!', 1)[-1].replace('$', ''))

        theme_part = next((t for k, t in rels.values() if k == 'theme'), 'xl/theme/theme1.xml')
        self.theme = _read_theme(self.zf, theme_part)
        styles_part = next((t for k, t in rels.values() if k == 'styles'), 'xl/styles.xml')
        self._load_styles(styles_part)
        sst_part = next((t for k, t in rels.values() if k == 'sharedStrings'), None)
        self.shared = self._load_shared(sst_part) if sst_part else []

    def _load_shared(self, part):
        try:
            root = self._xml(part)
        except KeyError:
            return []
        out = []
        for si in root.findall('{%s}si' % NS):
            out.append(_rich_text(si))
        return out

    def _load_styles(self, part):
        self.numfmts = dict(_BUILTIN_FMT)
        self.fonts, self.fills, self.borders, self.xfs = [], [], [], []
        try:
            root = self._xml(part)
        except KeyError:
            self.fonts = [dict(DEFAULT_FONT)]
            return
        for nf in root.findall('{%s}numFmts/{%s}numFmt' % (NS, NS)):
            try:
                self.numfmts[int(nf.get('numFmtId'))] = nf.get('formatCode', 'General')
            except (TypeError, ValueError):
                pass
        for f in root.findall('{%s}fonts/{%s}font' % (NS, NS)):
            self.fonts.append(self._font(f))
        for fl in root.findall('{%s}fills/{%s}fill' % (NS, NS)):
            self.fills.append(self._fill(fl))
        for b in root.findall('{%s}borders/{%s}border' % (NS, NS)):
            self.borders.append(self._border(b))
        for xf in root.findall('{%s}cellXfs/{%s}xf' % (NS, NS)):
            al = xf.find('{%s}alignment' % NS)
            self.xfs.append({
                'numfmt': _int(xf.get('numFmtId'), 0),
                'font': _int(xf.get('fontId'), 0),
                'fill': _int(xf.get('fillId'), 0),
                'border': _int(xf.get('borderId'), 0),
                'halign': (al.get('horizontal') if al is not None else None) or 'general',
                'valign': (al.get('vertical') if al is not None else None) or 'bottom',
                'wrap': al is not None and al.get('wrapText') in ('1', 'true'),
                'indent': _int(al.get('indent') if al is not None else None, 0),
            })
        if not self.fonts:
            self.fonts = [dict(DEFAULT_FONT)]

    def _font(self, f):
        def has(tag):
            el = f.find('{%s}%s' % (NS, tag))
            return el is not None and el.get('val', '1') not in ('0', 'false', 'none')
        sz = f.find('{%s}sz' % NS)
        name = f.find('{%s}name' % NS)
        return {
            'name': name.get('val') if name is not None else DEFAULT_FONT['name'],
            'size': _float(sz.get('val') if sz is not None else None, 11.0),
            'bold': has('b'),
            'italic': has('i'),
            'underline': has('u'),
            'color': _color(f.find('{%s}color' % NS), self.theme) or (0, 0, 0),
        }

    def _fill(self, fl):
        pf = fl.find('{%s}patternFill' % NS)
        if pf is not None:
            ptype = pf.get('patternType', 'none')
            if ptype == 'none':
                return None
            fg = _color(pf.find('{%s}fgColor' % NS), self.theme)
            if ptype == 'solid':
                return fg
            # patterns (gray125, lightGrid ...) are approximated by their colour
            return fg if ptype != 'gray125' else None
        gf = fl.find('{%s}gradientFill' % NS)
        if gf is not None:
            stop = gf.find('{%s}stop/{%s}color' % (NS, NS))
            return _color(stop, self.theme)
        return None

    def _border(self, b):
        out = {}
        for side in ('left', 'right', 'top', 'bottom'):
            el = b.find('{%s}%s' % (NS, side))
            style = el.get('style') if el is not None else None
            if style and style in BORDER_WEIGHT:
                col = _color(el.find('{%s}color' % NS), self.theme) or (0, 0, 0)
                out[side] = (BORDER_WEIGHT[style], col)
        return out

    # ── sheet listing ─────────────────────────────────────────────────
    def sheet_names(self, include_hidden=False):
        return [n for n, _, st in self.sheets if include_hidden or st == 'visible']

    def _sheet_part(self, sheet_name):
        if not sheet_name:
            return self.sheets[0][0], self.sheets[0][1]
        for name, part, _ in self.sheets:
            if name == sheet_name:
                return name, part
        for name, part, _ in self.sheets:
            if name.lower() == sheet_name.lower():
                return name, part
        raise XlsxError("Sheet '{}' not found in {}. Sheets: {}.".format(
            sheet_name, os.path.basename(self.path), ", ".join(self.sheet_names(True))))

    def excel_tables(self, sheet_name):
        """{table name: range} for the "Format as Table" ranges of a sheet."""
        name, part = self._sheet_part(sheet_name)
        out = {}
        for kind, target in self._rels(part).values():
            if kind != 'table':
                continue
            try:
                root = self._xml(target)
            except (KeyError, ET.ParseError):
                continue
            tname = root.get('displayName') or root.get('name')
            if tname and root.get('ref'):
                out[tname] = root.get('ref')
        return out

    def range_choices(self, sheet_name):
        """Range names a user can pick for a sheet, Used range first."""
        name, _ = self._sheet_part(sheet_name)
        out = [USED_RANGE]
        if name in self.print_areas:
            out.append('Print_Area')
        out.extend(sorted(self.excel_tables(name)))
        out.extend(sorted(n for n, (sh, _) in self.names.items() if sh in (None, name)))
        return out

    def resolve_range(self, sheet_name, text):
        """'' / Used range / Print_Area / a table or defined name / A1:F20."""
        name, _ = self._sheet_part(sheet_name)
        text = (text or '').strip()
        if not text or text == USED_RANGE:
            return None
        if text == 'Print_Area' and name in self.print_areas:
            return parse_range(self.print_areas[name])
        tables = self.excel_tables(name)
        if text in tables:
            return parse_range(tables[text])
        if text in self.names:
            return parse_range(self.names[text][1])
        return parse_range(text)

    # ── reading ───────────────────────────────────────────────────────
    def read_table(self, sheet_name=None, cell_range=None, include_hidden=False):
        """Read a range of a sheet into a Table (see class Table)."""
        name, part = self._sheet_part(sheet_name)
        bounds = self.resolve_range(name, cell_range)
        try:
            root = self._xml(part)
        except KeyError:
            raise XlsxError("Sheet '{}' is missing from the workbook file.".format(name))
        return _read_sheet(self, root, name, bounds, include_hidden,
                           cell_range or USED_RANGE)

    def read_view(self, sheet_name=None, include_hidden=False, max_rows=400, max_cols=50):
        """The sheet as Excel shows it, from A1, for picking cells by eye.

        Starts at A1 (like Excel), stops at the used range or the limits.
        Table.used_bounds = the real used range; Table.clipped = True when the
        limits cut the sheet off.
        """
        name, part = self._sheet_part(sheet_name)
        try:
            root = self._xml(part)
        except KeyError:
            raise XlsxError("Sheet '{}' is missing from the workbook file.".format(name))
        return _read_sheet(self, root, name, None, include_hidden, USED_RANGE,
                           view_limits=(max_rows, max_cols))


def _int(v, default):
    try:
        return int(v)
    except (TypeError, ValueError):
        return default


def _float(v, default):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _rich_text(node):
    """Text of an <si> / <is>: plain <t> or concatenated <r><t> runs."""
    parts = []
    for child in node:
        tag = child.tag.rsplit('}', 1)[-1]
        if tag == 't':
            parts.append(child.text or '')
        elif tag == 'r':
            t = child.find('{%s}t' % NS)
            if t is not None:
                parts.append(t.text or '')
    return ''.join(parts)


# ════════════════════════════════════════════════════════════════════
# TABLE
# ════════════════════════════════════════════════════════════════════

class Cell(object):
    __slots__ = ('text', 'value', 'kind', 'font', 'fill', 'borders',
                 'halign', 'valign', 'wrap', 'indent')

    def __init__(self):
        self.text = ''          # what Excel shows
        self.value = ''         # raw value as text (numbers unformatted)
        self.kind = 'blank'     # blank | n | s | b | e | d
        self.font = dict(DEFAULT_FONT)
        self.fill = None        # (r, g, b) or None
        self.borders = {}       # side -> (weight 1..3, (r, g, b))
        self.halign = 'left'    # left | center | right
        self.valign = 'bottom'  # top | center | bottom
        self.wrap = False
        self.indent = 0

    def as_dict(self):
        return {k: getattr(self, k) for k in self.__slots__}


class Table(object):
    """A rectangular block of formatted cells, 0-based rows/cols.

    rows / cols       — count after hidden rows/columns were dropped
    cells[r][c]       — Cell
    col_widths_in     — paper width of each column, inches (96 px = 1 in)
    row_heights_in    — paper height of each row, inches (72 pt = 1 in)
    merges            — [(r1, c1, r2, c2)] inclusive, 0-based
    source_rows/cols  — 1-based Excel row / column of each table row / column
    """

    def __init__(self):
        self.sheet = ''
        self.range_name = USED_RANGE
        self.bounds = None
        self.rows = 0
        self.cols = 0
        self.cells = []
        self.col_widths_in = []
        self.row_heights_in = []
        self.merges = []
        self.source_rows = []
        self.source_cols = []
        self.used_bounds = None
        self.clipped = False

    def text_rows(self):
        return [[c.text for c in row] for row in self.cells]

    def value_rows(self):
        return [[c.value for c in row] for row in self.cells]

    def fingerprint(self):
        """Stable hash of everything that is drawn — detects Excel edits."""
        payload = {
            'cells': [[c.as_dict() for c in row] for row in self.cells],
            'w': [round(w, 4) for w in self.col_widths_in],
            'h': [round(h, 4) for h in self.row_heights_in],
            'm': self.merges,
        }
        raw = json.dumps(payload, sort_keys=True, default=list).encode('utf-8')
        return hashlib.sha1(raw).hexdigest()


def _col_px(width):
    """Stored <col width> (characters incl. padding) -> pixels, Calibri 11."""
    return int(((256.0 * width + int(128.0 / 7)) / 256.0) * 7)


def _read_sheet(wb, root, sheet_name, bounds, include_hidden, range_name,
                view_limits=None):
    fmt = root.find('{%s}sheetFormatPr' % NS)
    default_h = _float(fmt.get('defaultRowHeight') if fmt is not None else None, 15.0)
    default_px = 64
    if fmt is not None and fmt.get('defaultColWidth'):
        default_px = _col_px(_float(fmt.get('defaultColWidth'), 8.43))
    elif fmt is not None and fmt.get('baseColWidth'):
        default_px = _int(fmt.get('baseColWidth'), 8) * 7 + 5 + 3

    col_px, col_hidden, col_style = {}, set(), {}
    for col in root.findall('{%s}cols/{%s}col' % (NS, NS)):
        lo, hi = _int(col.get('min'), 1), _int(col.get('max'), 1)
        hi = min(hi, lo + 16384)
        px = _col_px(_float(col.get('width'), 8.43)) if col.get('width') else default_px
        hidden = col.get('hidden') in ('1', 'true')
        for c in range(lo, hi + 1):
            col_px[c] = px
            if hidden:
                col_hidden.add(c)
            if col.get('style'):
                col_style[c] = _int(col.get('style'), 0)

    raw = {}             # (r, c) -> (kind, value text, style index)
    row_ht, row_hidden, row_style = {}, set(), {}
    max_r = max_c = 0
    min_r = min_c = None
    for row in root.findall('{%s}sheetData/{%s}row' % (NS, NS)):
        rn = _int(row.get('r'), 0)
        if row.get('ht'):
            row_ht[rn] = _float(row.get('ht'), default_h)
        if row.get('hidden') in ('1', 'true'):
            row_hidden.add(rn)
        if row.get('s') and row.get('customFormat') in ('1', 'true'):
            row_style[rn] = _int(row.get('s'), 0)
        for cel in row.findall('{%s}c' % NS):
            try:
                r, c = parse_ref(cel.get('r', ''))
            except XlsxError:
                continue
            t = cel.get('t', 'n')
            s = _int(cel.get('s'), 0)
            v = cel.find('{%s}v' % NS)
            if t == 'inlineStr':
                node = cel.find('{%s}is' % NS)
                val, kind = (_rich_text(node) if node is not None else ''), 's'
            elif v is None or v.text is None:
                val, kind = '', 'blank'
            elif t == 's':
                idx = _int(v.text, -1)
                val, kind = (wb.shared[idx] if 0 <= idx < len(wb.shared) else ''), 's'
            elif t == 'str':
                val, kind = v.text, 's'
            elif t == 'b':
                val, kind = ('TRUE' if v.text.strip() == '1' else 'FALSE'), 'b'
            elif t == 'e':
                val, kind = v.text, 'e'
            elif t == 'd':
                val, kind = v.text, 'd'
            else:
                val, kind = v.text, 'n'
            raw[(r, c)] = (kind, val, s)
            styled = s and (s < len(wb.xfs)) and (
                wb.xfs[s]['fill'] or wb.xfs[s]['border'])
            if kind != 'blank' or styled:
                max_r, max_c = max(max_r, r), max(max_c, c)
                min_r = r if min_r is None else min(min_r, r)
                min_c = c if min_c is None else min(min_c, c)

    merges_abs = []
    for mc in root.findall('{%s}mergeCells/{%s}mergeCell' % (NS, NS)):
        try:
            merges_abs.append(parse_range(mc.get('ref', '')))
        except XlsxError:
            pass

    if bounds is None:
        if min_r is None:
            raise XlsxError("Sheet '{}' is empty.".format(sheet_name))
        for (r1, c1, r2, c2) in merges_abs:
            if r1 <= max_r and c1 <= max_c:
                max_r, max_c = max(max_r, r2), max(max_c, c2)
        bounds = (min_r, min_c, max_r, max_c)
    used_bounds = bounds
    clipped = False
    if view_limits:
        max_rows, max_cols = view_limits
        r2 = min(bounds[2], max_rows)
        c2 = min(bounds[3], max_cols)
        clipped = (r2, c2) != (bounds[2], bounds[3])
        bounds = (1, 1, r2, c2)
    r1, c1, r2, c2 = bounds
    if (r2 - r1 + 1) * (c2 - c1 + 1) > 200000:
        raise XlsxError("The range {} is too large ({} rows x {} columns). "
                        "Pick a smaller range or a named range.".format(
                            range_text(*bounds), r2 - r1 + 1, c2 - c1 + 1))

    rows = [r for r in range(r1, r2 + 1) if include_hidden or r not in row_hidden]
    cols = [c for c in range(c1, c2 + 1) if include_hidden or c not in col_hidden]
    if not rows or not cols:
        raise XlsxError("Every row or column of {} is hidden.".format(range_text(*bounds)))

    table = Table()
    table.sheet = sheet_name
    table.range_name = range_name
    table.bounds = bounds
    table.used_bounds = used_bounds
    table.clipped = clipped
    table.source_rows, table.source_cols = rows, cols
    table.rows, table.cols = len(rows), len(cols)
    table.col_widths_in = [col_px.get(c, default_px) / 96.0 for c in cols]
    table.row_heights_in = [row_ht.get(r, default_h) / 72.0 for r in rows]

    for r in rows:
        line = []
        for c in cols:
            cell = Cell()
            kind, val, s = raw.get((r, c), ('blank', '', None))
            if s is None:
                s = row_style.get(r, col_style.get(c, 0))
            _style_cell(wb, cell, kind, val, s)
            line.append(cell)
        table.cells.append(line)

    # merges -> table coordinates, clipped to the range and to visible rows/cols
    row_ix = {r: i for i, r in enumerate(rows)}
    col_ix = {c: i for i, c in enumerate(cols)}
    for (mr1, mc1, mr2, mc2) in merges_abs:
        if mr2 < r1 or mr1 > r2 or mc2 < c1 or mc1 > c2:
            continue
        vr = [row_ix[r] for r in range(max(mr1, r1), min(mr2, r2) + 1) if r in row_ix]
        vc = [col_ix[c] for c in range(max(mc1, c1), min(mc2, c2) + 1) if c in col_ix]
        if not vr or not vc or (len(vr) == 1 and len(vc) == 1):
            continue
        table.merges.append((vr[0], vc[0], vr[-1], vc[-1]))
        # Excel draws the outer border of a merge from its edge cells
        top_left = table.cells[vr[0]][vc[0]]
        for i in vr:
            for j in vc:
                if (i, j) == (vr[0], vc[0]):
                    continue
                other = table.cells[i][j]
                if i == vr[0] and 'top' in other.borders:
                    top_left.borders.setdefault('top', other.borders['top'])
                if j == vc[-1] and 'right' in other.borders:
                    top_left.borders['right'] = other.borders['right']
                if i == vr[-1] and 'bottom' in other.borders:
                    top_left.borders['bottom'] = other.borders['bottom']
    return table


def _style_cell(wb, cell, kind, val, s):
    xf = wb.xfs[s] if 0 <= s < len(wb.xfs) else None
    fmt_code = 'General'
    if xf:
        cell.font = dict(wb.fonts[xf['font']] if xf['font'] < len(wb.fonts) else DEFAULT_FONT)
        cell.fill = wb.fills[xf['fill']] if xf['fill'] < len(wb.fills) else None
        cell.borders = dict(wb.borders[xf['border']]) if xf['border'] < len(wb.borders) else {}
        cell.valign = {'center': 'center', 'top': 'top', 'justify': 'center',
                       'distributed': 'center'}.get(xf['valign'], 'bottom')
        cell.wrap = xf['wrap'] or xf['halign'] in ('justify', 'distributed')
        cell.indent = xf['indent']
        fmt_code = wb.numfmts.get(xf['numfmt'], 'General')
        halign = xf['halign']
    else:
        cell.font = dict(wb.fonts[0]) if wb.fonts else dict(DEFAULT_FONT)
        halign = 'general'

    cell.kind = kind
    cell.value = val
    if kind == 'n':
        try:
            num = float(val)
            cell.text = format_value(num, fmt_code, wb.date1904)
            cell.value = _general(num)
        except ValueError:
            cell.text = val
    elif kind == 's':
        cell.text = format_text(val, fmt_code)
    else:
        cell.text = val

    if halign in ('center', 'centerContinuous', 'distributed'):
        cell.halign = 'center'
    elif halign == 'right':
        cell.halign = 'right'
    elif halign in ('left', 'fill', 'justify'):
        cell.halign = 'left'
    else:                                   # general
        cell.halign = {'n': 'right', 'd': 'right', 'b': 'center', 'e': 'center'}.get(kind, 'left')


def list_sheets(path):
    with Workbook(path) as wb:
        return wb.sheet_names()


def bounds_of_cells(cells):
    """[(row, col)] (1-based Excel) -> 'B2:F20' covering them all, or ''."""
    cells = list(cells)
    if not cells:
        return ''
    rows = [r for r, _ in cells]
    cols = [c for _, c in cells]
    return range_text(min(rows), min(cols), max(rows), max(cols))


def read_table(path, sheet_name=None, cell_range=None, include_hidden=False):
    with Workbook(path) as wb:
        return wb.read_table(sheet_name, cell_range, include_hidden)


# ════════════════════════════════════════════════════════════════════
# LAYOUT  (paper space, inches, origin = top-left, y grows DOWN)
# ════════════════════════════════════════════════════════════════════

class Layout(object):
    """Geometry shared by the Revit renderers.

    width / height  — overall table size, inches
    rects           — [(r, c, x, y, w, h)] one per drawn cell (merged = one)
    fills           — [(x, y, w, h, (r, g, b))]
    segments        — [((x1, y1), (x2, y2), weight, (r, g, b))] merged collinear
    texts           — [{'x','y','w','h','text','font','halign','valign','wrap','r','c'}]
    """

    def __init__(self):
        self.width = 0.0
        self.height = 0.0
        self.rects = []
        self.fills = []
        self.segments = []
        self.texts = []


def build_layout(table, scale=1.0, formatting=True):
    """Place every cell of `table` on paper. scale 1.0 = Excel at 100 %."""
    xs = [0.0]
    for w in table.col_widths_in:
        xs.append(xs[-1] + w * scale)
    ys = [0.0]
    for h in table.row_heights_in:
        ys.append(ys[-1] + h * scale)

    owner = {}                         # (r, c) -> top-left (r, c) of its merge
    span = {}                          # top-left -> (r2, c2)
    for (r1, c1, r2, c2) in table.merges:
        span[(r1, c1)] = (r2, c2)
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                owner[(r, c)] = (r1, c1)

    lay = Layout()
    lay.width, lay.height = xs[-1], ys[-1]

    # ── rects, fills, texts ──────────────────────────────────────────
    for r in range(table.rows):
        for c in range(table.cols):
            if owner.get((r, c), (r, c)) != (r, c):
                continue
            r2, c2 = span.get((r, c), (r, c))
            x, y = xs[c], ys[r]
            w, h = xs[c2 + 1] - x, ys[r2 + 1] - y
            lay.rects.append((r, c, x, y, w, h))
            cell = table.cells[r][c]
            if formatting and cell.fill:
                lay.fills.append((x, y, w, h, cell.fill))
            if not cell.text.strip():
                continue
            tw = w
            # Excel lets unwrapped text run on over EMPTY cells to its right
            # (left aligned) or left (right aligned) — mimic it.
            if not cell.wrap and (r, c) not in span:
                if cell.halign == 'left':
                    k = c + 1
                    while k < table.cols and (r, k) not in owner and not table.cells[r][k].text.strip():
                        k += 1
                    tw = xs[k] - x
                elif cell.halign == 'right':
                    k = c - 1
                    while k >= 0 and (r, k) not in owner and not table.cells[r][k].text.strip():
                        k -= 1
                    tw = xs[c + 1] - xs[k + 1]
                    x = xs[k + 1]
            font = dict(cell.font) if formatting else dict(DEFAULT_FONT)
            font['size'] = font['size'] * scale
            lay.texts.append({
                'x': x, 'y': y, 'w': tw, 'h': h, 'text': cell.text,
                'font': font,
                'halign': cell.halign, 'valign': cell.valign,
                'wrap': cell.wrap, 'indent': cell.indent * scale, 'r': r, 'c': c,
            })

    if not formatting:
        # plain grid when formatting is off
        for x in xs:
            lay.segments.append(((x, 0.0), (x, ys[-1]), 1, (0, 0, 0)))
        for y in ys:
            lay.segments.append(((0.0, y), (xs[-1], y), 1, (0, 0, 0)))
        return lay

    # ── borders: one value per grid edge, the heavier side wins ─────
    def border_of(r, c, side):
        tl = owner.get((r, c), (r, c))
        cell = table.cells[tl[0]][tl[1]]
        r2, c2 = span.get(tl, tl)
        # inner edges of a merge are never drawn
        if side == 'right' and c != c2:
            return None
        if side == 'left' and c != tl[1]:
            return None
        if side == 'bottom' and r != r2:
            return None
        if side == 'top' and r != tl[0]:
            return None
        return cell.borders.get(side)

    def heavier(a, b):
        if not a:
            return b
        if not b:
            return a
        return a if a[0] >= b[0] else b

    # horizontal edges: row line i (0..rows), column c
    h_edges = {}
    for i in range(table.rows + 1):
        for c in range(table.cols):
            above = border_of(i - 1, c, 'bottom') if i > 0 else None
            below = border_of(i, c, 'top') if i < table.rows else None
            if 0 < i < table.rows:
                oa = owner.get((i - 1, c), (i - 1, c))
                ob = owner.get((i, c), (i, c))
                if oa == ob:
                    continue
            b = heavier(above, below)
            if b:
                h_edges[(i, c)] = b
    v_edges = {}
    for j in range(table.cols + 1):
        for r in range(table.rows):
            left = border_of(r, j - 1, 'right') if j > 0 else None
            right = border_of(r, j, 'left') if j < table.cols else None
            if 0 < j < table.cols:
                ol = owner.get((r, j - 1), (r, j - 1))
                orr = owner.get((r, j), (r, j))
                if ol == orr:
                    continue
            b = heavier(left, right)
            if b:
                v_edges[(j, r)] = b

    # merge collinear runs with the same style into one segment
    for i in range(table.rows + 1):
        c = 0
        while c < table.cols:
            b = h_edges.get((i, c))
            if not b:
                c += 1
                continue
            start = c
            while c + 1 < table.cols and h_edges.get((i, c + 1)) == b:
                c += 1
            lay.segments.append(((xs[start], ys[i]), (xs[c + 1], ys[i]), b[0], b[1]))
            c += 1
    for j in range(table.cols + 1):
        r = 0
        while r < table.rows:
            b = v_edges.get((j, r))
            if not b:
                r += 1
                continue
            start = r
            while r + 1 < table.rows and v_edges.get((j, r + 1)) == b:
                r += 1
            lay.segments.append(((xs[j], ys[start]), (xs[j], ys[r + 1]), b[0], b[1]))
            r += 1
    return lay
