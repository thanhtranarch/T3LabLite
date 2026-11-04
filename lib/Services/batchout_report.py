# -*- coding: utf-8 -*-
"""BatchOut export report: one row per queue item x format, written as XLSX or CSV.

Revit-free; tested by `dev/test_batchout_report.py`.
"""

import csv
import os
from datetime import datetime

# (key, header). Order = default column order; a profile's ReportColumns picks
# and reorders these keys.
COLUMNS = (
    ('number', 'Number'),
    ('name', 'Name'),
    ('revision', 'Revision'),
    ('format', 'Format'),
    ('size', 'Paper Size'),
    ('orientation', 'Orientation'),
    ('status', 'Status'),
    ('reason', 'Reason'),
    ('file', 'File'),
    ('file_kb', 'File Size (KB)'),
    ('model', 'Model'),
    ('exported_at', 'Exported At'),
)
DEFAULT_COLUMNS = [key for key, _ in COLUMNS]
_HEADERS = dict(COLUMNS)

# XLSX first: a comma CSV opens as one column in Excel on locales whose list
# separator is ';' (vi-VN, de-DE, ...).
REPORT_FORMATS = ('XLSX', 'CSV')
NOT_EXPORTED = 'Not exported'


def normalize_columns(keys):
    """Known keys in the given order, no duplicates; nothing usable -> defaults."""
    out = []
    for key in keys or ():
        key = str(key).strip().lower()
        if key in _HEADERS and key not in out:
            out.append(key)
    return out or list(DEFAULT_COLUMNS)


def normalize_format(fmt):
    fmt = str(fmt or '').strip().upper()
    return fmt if fmt in REPORT_FORMATS else REPORT_FORMATS[0]


def _file_kb(path):
    try:
        if path and os.path.isfile(path):
            return int(round(os.path.getsize(path) / 1024.0))
    except OSError:
        pass
    return ''


def build_rows(items, reasons=None, model='', exported_at=None):
    """Report rows (dicts keyed by column key) from BatchOut queue rows.

    `reasons`: {(item number, format): text} from the run's failures. A row
    whose status was never set counts as not exported.
    """
    reasons = reasons or {}
    stamp = (exported_at or datetime.now()).strftime('%Y-%m-%d %H:%M')
    rows = []
    for item in items:
        number = str(getattr(item, 'SheetNumber', '') or '')
        fmt = str(getattr(item, 'Format', '') or '')
        status = str(getattr(item, 'Status', '') or '') or NOT_EXPORTED
        path = getattr(item, 'OutputFile', '') or ''
        exported = status == 'Exported'
        rows.append({
            'number': number,
            'name': str(getattr(item, 'SheetName', '') or ''),
            'revision': str(getattr(item, 'Revision', '') or ''),
            'format': fmt,
            'size': str(getattr(item, 'Size', '') or ''),
            'orientation': str(getattr(item, 'Orientation', '') or ''),
            'status': status,
            'reason': '' if exported else reasons.get((number, fmt), ''),
            'file': path if exported else '',
            'file_kb': _file_kb(path) if exported else '',
            'model': model or '',
            'exported_at': stamp,
        })
    return rows


def summarize(rows):
    """{status: count}, for the completion message."""
    counts = {}
    for row in rows:
        counts[row['status']] = counts.get(row['status'], 0) + 1
    return counts


def report_path(folder, fmt, when=None):
    """Timestamped path, so a new run never overwrites the last report."""
    when = when or datetime.now()
    ext = '.csv' if normalize_format(fmt) == 'CSV' else '.xlsx'
    return os.path.join(folder, 'BatchOut_Report_{}{}'.format(
        when.strftime('%Y%m%d_%H%M%S'), ext))


def write_report(path, rows, columns=None, fmt=None):
    """Write `rows` to `path`; returns the path. Raises on I/O errors."""
    columns = normalize_columns(columns)
    headers = [_HEADERS[key] for key in columns]
    table = [['' if row.get(key) is None else row.get(key) for key in columns]
             for row in rows]
    if normalize_format(fmt or os.path.splitext(path)[1].lstrip('.')) == 'CSV':
        # utf-8-sig: Excel needs the BOM to read Vietnamese sheet names.
        with open(path, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(headers)
            writer.writerows(table)
    else:
        from Snippets.schedule_data import write_xlsx
        write_xlsx(path, [{'name': 'BatchOut Report', 'headers': headers,
                           'rows': [[str(v) for v in r] for r in table]}])
    return path
