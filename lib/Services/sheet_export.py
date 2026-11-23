# -*- coding: utf-8 -*-
"""Sheet export without the BatchOut window: PDF, DWG, DXF, DGN from sheets, a
Revit sheet set (ViewSheetSet) and/or a saved BatchOut profile.

Used by the MCP tool `export_sheets` (core/server.py, IronPython 2.7) and the
Assistant (CPython 3); BatchOut itself imports apply_revit_display_fidelity
from here. Reuses BatchOut's Services: batchout_profiles (settings),
batchout_refs (DWG image references), batchout_report (report file).
Native-crash markers and safe mode stay in the BatchOut window.
Doc: docs/agents/sheet-export.md. Test: dev/test_sheet_export.py.
"""
from __future__ import unicode_literals

import logging
import os
import re
import time
from datetime import datetime

try:
    from Autodesk.Revit import DB
except Exception:            # tests / outside Revit
    DB = None

from Services import batchout_profiles, batchout_refs, batchout_report

logger = logging.getLogger('T3Lab.sheet_export')

FORMATS = ('pdf', 'dwg', 'dxf', 'dgn')
DEFAULT_PATTERN = '{SheetNumber}-{SheetName}'            # BatchOut's naming default
PROFILES_FOLDER = os.path.join(os.path.expanduser('~'), 'Documents', 'T3Lab_BatchOut_Profiles')
DEFAULT_ROOT = os.path.join(os.path.expanduser('~'), 'Documents', 'Revit Exports')
RASTER_LADDER = ('Presentation', 'High', 'Medium', 'Low')   # = BatchOut _apply_pdf_quality
ACAD_VERSIONS = (('2018', 'R2018'), ('2013', 'R2013'), ('2010', 'R2010'), ('2007', 'R2007'))
_BAD_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_PLACEHOLDER = re.compile(r'\{([^{}]+)\}')
CLOCK_SLACK = 2.0


# ─── DWG "Match Revit display" (shared with GUI/BatchOutDialog.py) ───

def _enum_member(enum_name, *members):
    """First member of DB.<enum_name> this Revit build has, or None."""
    enum = getattr(DB, enum_name, None)
    for member in members:
        value = getattr(enum, member, None) if enum is not None else None
        if value is not None:
            return value
    return None


def apply_revit_display_fidelity(options, for_sheets=True):
    """Make a DWG export look like the Revit view it came from.

    Returns the applied settings, for the log.
    - Colors = TrueColorPerView (RGB incl. V/G overrides + filters; default
      IndexColors colours by layer table, e.g. AIA magenta text). Falls back to
      TrueColor on builds without it.
    - PropOverrides = ByEntity: element overrides stay on the entity, not the
      category layer.
    - LineScaling: PaperSpace for sheets (PSLTSCALE = 1), ViewScale for views.
    - Hatch background white only when the setup turns it on — never forced.
    Line weights are Revit's own; AutoCAD shows them only with LWDISPLAY on.
    """
    applied = []
    for prop, value, label in (
            ('Colors', _enum_member('ExportColorMode', 'TrueColorPerView', 'TrueColor'),
             'true colors'),
            ('PropOverrides', _enum_member('PropOverrideMode', 'ByEntity'),
             'overrides by entity'),
            ('LineScaling', _enum_member('LineScaling',
                                         'PaperSpace' if for_sheets else 'ViewScale'),
             'linetype scale')):
        if value is None or not hasattr(options, prop):
            continue
        try:
            setattr(options, prop, value)
            applied.append(label)
        except Exception as ex:
            logger.debug("Could not set DWG {}: {}".format(prop, ex))
    try:
        if getattr(options, 'UseHatchBackgroundColor', False):
            options.HatchBackgroundColor = DB.Color(255, 255, 255)
            applied.append('white hatch background')
    except Exception as ex:
        logger.debug("Could not set DWG HatchBackgroundColor: {}".format(ex))
    return applied


# ─── Pure ───

def settings_from(profile=None, **overrides):
    """Export settings: BatchOut defaults < profile dict (ExportProfile keys) <
    overrides that are not None."""
    p = profile or {}
    s = {
        'pattern': p.get('NamingPattern') or DEFAULT_PATTERN,
        'combined': bool(p.get('CombinePDF', False)),
        'hide_ref_planes': bool(p.get('PDFHideRefPlanes', False)),
        'hide_scope_boxes': bool(p.get('PDFHideScopeBoxes', False)),
        'hide_crop_boundaries': bool(p.get('PDFHideCropBoundaries', False)),
        'hide_unreferenced_tags': bool(p.get('PDFHideUnreferencedTags', False)),
        'dwg_version': p.get('DWGVersion') or 'AutoCAD 2013',
        'export_setup': p.get('CADExportSetup') or '',
        'match_display': bool(p.get('CADMatchRevitDisplay', True)),
        'views_on_sheets': bool(p.get('CADExportViewsOnSheets', False)),
        'links_external': bool(p.get('CADExportLinksAsExternal', False)),
        'skip_images': bool(p.get('CADSkipImageReferences', False)),
        'folder': p.get('OutputFolder') or DEFAULT_ROOT,
        'split_by_format': bool(p.get('SplitByFormat', True)) if p else True,
        'report': bool(p.get('WriteReport', False)),
        'report_format': p.get('ReportFormat') or 'XLSX',
        'report_columns': p.get('ReportColumns') or None,
    }
    if s['export_setup'] == 'Use setup from file':          # BatchOut's "none" entry
        s['export_setup'] = ''
    for k, v in overrides.items():
        if v is not None:
            s[k] = v
    return s


def output_folder(settings, fmt):
    """BatchOut layout: <folder>/<PDF|DWG|...> unless the profile turned it off."""
    folder = settings['folder']
    return os.path.join(folder, fmt.upper()) if settings.get('split_by_format') else folder


def placeholders(pattern):
    return set(_PLACEHOLDER.findall(pattern or ''))


def file_stem(pattern, values):
    """Pattern with {Name} replaced from `values` (missing -> ''), Windows-safe."""
    text = _PLACEHOLDER.sub(lambda m: '{}'.format(values.get(m.group(1), '') or ''),
                            pattern or DEFAULT_PATTERN)
    text = _BAD_CHARS.sub('_', text)
    text = re.sub(r'\s+', ' ', text).strip(' .')
    return text[:200] or 'Unnamed'


def reserve(used, folder, stem, ext):
    """Stem unclaimed in this run; ' (2)', ' (3)', ... on collision (as BatchOut)."""
    candidate, n = stem, 1
    while os.path.join(folder, candidate + ext).lower() in used:
        n += 1
        candidate = '{} ({})'.format(stem, n)
    used.add(os.path.join(folder, candidate + ext).lower())
    return candidate


def select(rows, numbers=None, ids=None, prefix=None, member_ids=None):
    """Rows {id, number, name} kept by the filters (all given filters apply);
    returns (kept rows in sheet-number order, numbers / ids asked for but not found)."""
    out = rows
    missing = []
    if member_ids is not None:
        member_ids = set(member_ids)
        out = [r for r in out if r['id'] in member_ids]
    if ids:
        want = set(int(i) for i in ids)
        missing += ['id {}'.format(i) for i in sorted(want - set(r['id'] for r in rows))]
        out = [r for r in out if r['id'] in want]
    if numbers:
        want = dict(('{}'.format(n).strip().lower(), '{}'.format(n)) for n in numbers)
        have = set(r['number'].lower() for r in rows)
        missing += [v for k, v in sorted(want.items()) if k not in have]
        out = [r for r in out if r['number'].lower() in want]
    if prefix:
        p = '{}'.format(prefix).strip().lower()
        out = [r for r in out if r['number'].lower().startswith(p)]
    return sorted(out, key=lambda r: _natural(r['number'])), missing


def _natural(text):
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r'(\d+)', text or '')]


def fresh_file(path, prev_mtime, started):
    """Problem text, or None when Revit wrote a new, non-empty, valid file."""
    if not os.path.isfile(path):
        return 'Revit wrote no file'
    try:
        mtime, size = os.path.getmtime(path), os.path.getsize(path)
    except OSError as ex:
        return 'unreadable: {}'.format(ex)
    if size == 0:
        return 'empty file'
    if mtime < started - CLOCK_SLACK and (prev_mtime is None or mtime == prev_mtime):
        return 'file not rewritten by this export'
    try:
        from core import tool_cli
        r = tool_cli.verify_file(path)
        return None if r.get('ok') else r.get('problem')
    except ImportError:
        return None


# ─── Profiles (BatchOut's folder) ───

def profile_names(folder=None):
    return [name for name, _path, _data in
            batchout_profiles.list_profiles(folder or PROFILES_FOLDER)]


def load_profile(name, folder=None):
    """Profile dict by name (case-insensitive), or None."""
    want = '{}'.format(name or '').strip().lower()
    for pname, _path, data in batchout_profiles.list_profiles(folder or PROFILES_FOLDER):
        if pname.lower() == want:
            return data
    return None


# ─── Revit: reading ───

def _eid(element_id):
    try:
        return int(element_id.Value)
    except AttributeError:
        return int(element_id.IntegerValue)


def sheet_rows(doc):
    """[{id, number, name, sheet}] of every non-placeholder sheet."""
    out = []
    for s in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheet):
        try:
            if s.IsPlaceholder:
                continue
            out.append({'id': _eid(s.Id), 'number': s.SheetNumber or '', 'name': s.Name or '',
                        'sheet': s})
        except Exception:
            continue
    return out


def sheet_sets(doc):
    """[{name, sheets, views}] of the saved Revit sheet sets (read only, no PrintManager)."""
    out = []
    for vss in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet):
        try:
            views = list(vss.Views)
            n_sheets = sum(1 for v in views if isinstance(v, DB.ViewSheet))
            out.append({'name': vss.Name, 'sheets': n_sheets, 'views': len(views) - n_sheets})
        except Exception:
            continue
    return sorted(out, key=lambda r: r['name'].lower())


def set_member_ids(doc, name):
    """(sheet ids of the sheet set named `name`, other views it holds) or (None, 0)."""
    want = '{}'.format(name or '').strip().lower()
    for vss in DB.FilteredElementCollector(doc).OfClass(DB.ViewSheetSet):
        try:
            if (vss.Name or '').lower() != want:
                continue
            views = list(vss.Views)
            ids = [_eid(v.Id) for v in views if isinstance(v, DB.ViewSheet)]
            return ids, len(views) - len(ids)
        except Exception:
            continue
    return None, 0


def export_setups(doc):
    """Names of the DWG/DXF/DGN export setups saved in the model."""
    try:
        return sorted(DB.BaseExportOptions.GetPredefinedSetupNames(doc))
    except Exception:
        return []


def sheet_values(doc, sheet, pattern):
    """Placeholder values for one sheet: BatchOut's standard keys, plus any
    sheet parameter the pattern names."""
    values = {'SheetNumber': sheet.SheetNumber or '', 'SheetName': sheet.Name or '',
              'Date': datetime.now().strftime('%Y%m%d'),
              'Time': datetime.now().strftime('%H%M%S')}
    try:
        pi = doc.ProjectInformation
        values.update(ProjectNumber=pi.Number or '', ProjectName=pi.Name or '',
                      ProjectAddress=pi.Address or '', ClientName=pi.ClientName or '',
                      ProjectStatus=pi.Status or '')
    except Exception:
        pass
    wanted = placeholders(pattern) - set(values)
    if 'Revision' in wanted:
        try:
            p = sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION)
            values['Revision'] = (p.AsString() or '') if p else ''
        except Exception:
            values['Revision'] = ''
    for name in wanted - set(values):
        try:
            p = sheet.LookupParameter(name)
            values[name] = (p.AsString() or p.AsValueString() or '') if p else ''
        except Exception:
            values[name] = ''
    return values


# ─── Revit: options ───

def pdf_options(stem, s):
    o = DB.PDFExportOptions()
    o.Combine = True            # even for one sheet: with False Revit ignores FileName
    o.FileName = stem
    try:
        o.PaperFormat = DB.ExportPaperFormat.Default       # the sheet's own size
    except Exception:
        pass
    rq = getattr(DB, 'RasterQualityType', None)
    for name in RASTER_LADDER:
        value = getattr(rq, name, None) if rq is not None else None
        if value is not None:
            try:
                o.RasterQuality = value
            except Exception:
                pass
            break
    try:
        o.ColorDepth = DB.ColorDepthType.Color
    except Exception:
        pass
    for prop, key in (('HideReferencePlane', 'hide_ref_planes'),
                      ('HideScopeBoxes', 'hide_scope_boxes'),
                      ('HideCropBoundaries', 'hide_crop_boundaries'),
                      ('HideUnreferencedViewTags', 'hide_unreferenced_tags')):
        if s.get(key):
            try:
                setattr(o, prop, True)
            except Exception:
                pass
    return o


def cad_options(doc, fmt, s, for_sheets=True):
    """(options, notes) for dwg / dxf / dgn, from the named setup when given."""
    notes = []
    cls = {'dwg': DB.DWGExportOptions, 'dxf': DB.DXFExportOptions,
           'dgn': DB.DGNExportOptions}[fmt]
    o = None
    if s.get('export_setup'):
        try:
            o = cls.GetPredefinedOptions(doc, s['export_setup'])
        except Exception as ex:
            notes.append('export setup "{}" not applied: {}'.format(s['export_setup'], ex))
    if o is None:
        o = cls()
    if fmt == 'dgn':
        return o, notes
    if s.get('match_display'):
        apply_revit_display_fidelity(o, for_sheets=for_sheets)
    version = '{}'.format(s.get('dwg_version') or '')
    for year, member in ACAD_VERSIONS:
        if year in version:
            value = _enum_member('ACADVersion', member, 'R2013')
            if value is not None:
                try:
                    o.FileVersion = value
                except Exception:
                    pass
            break
    try:
        area = _enum_member('ExportingAreas', 'ExportViewsOnSheets' if s.get('views_on_sheets')
                            else 'DontExportViewsOnSheets')
        if area is not None and hasattr(o, 'ExportingAreas'):
            o.ExportingAreas = area
    except Exception:
        pass
    try:
        if hasattr(o, 'MergedViews'):
            o.MergedViews = not s.get('links_external')
    except Exception:
        pass
    return o, notes


# ─── Revit: export ───

def _ids(elements):
    from System.Collections.Generic import List
    out = List[DB.ElementId]()
    for e in elements:
        out.Add(e.Id)
    return out


class _Row(object):
    """What batchout_report.build_rows reads from a BatchOut queue row."""

    def __init__(self, sheet, fmt):
        self.SheetNumber, self.SheetName = sheet['number'], sheet['name']
        self.Format, self.Status, self.OutputFile = fmt.upper(), '', ''
        self.Revision = self.Size = self.Orientation = ''


def export(doc, fmt, sheets, folder, s):
    """Export `sheets` (sheet_rows entries). Returns {files, rows, notes}.

    One file per sheet; PDF with s['combined'] makes one file. Every file is
    checked fresh, non-empty and with the right header before it counts.
    """
    if not os.path.isdir(folder):
        os.makedirs(folder)
    used, files, notes, rows = set(), [], [], []
    queue = [_Row(sh, fmt) for sh in sheets]
    reasons = {}
    if fmt == 'pdf' and s.get('combined'):
        jobs = [(sheets, '{}-{}_Combined'.format(sheets[0]['number'], sheets[-1]['number'])
                 if len(sheets) > 1 else None)]
    else:
        jobs = [([sh], None) for sh in sheets]
    options = None
    if fmt != 'pdf':
        options, setup_notes = cad_options(doc, fmt, s)
        notes += setup_notes
    for group, combined_name in jobs:
        sheet = group[0]
        stem = combined_name or file_stem(s['pattern'], sheet_values(doc, sheet['sheet'],
                                                                     s['pattern']))
        ext = '.' + fmt
        stem = reserve(used, folder, stem, ext)
        path = os.path.join(folder, stem + ext)
        prev = os.path.getmtime(path) if os.path.exists(path) else None
        started = time.time()
        before = batchout_refs.snapshot(folder) if fmt == 'dwg' and s.get('skip_images') else None
        problem = None
        try:
            ids = _ids(r['sheet'] for r in group)
            if fmt == 'pdf':
                ok = doc.Export(folder, ids, pdf_options(stem, s))
            else:
                ok = doc.Export(folder, stem, ids, options)
            if ok is False:
                problem = 'Revit reported the export as incomplete'
        except Exception as ex:
            problem = '{}'.format(getattr(ex, 'Message', None) or ex).strip().splitlines()[0]
        if before is not None:
            deleted, _failed = batchout_refs.delete_files(
                batchout_refs.written_reference_files(folder, before, keep=[path]))
            if deleted:
                notes.append('{} image reference file(s) deleted'.format(len(deleted)))
        problem = problem or fresh_file(path, prev, started)
        for r in queue:
            if r.SheetNumber in set(g['number'] for g in group):
                if problem:
                    r.Status = 'Failed'
                    reasons[(r.SheetNumber, r.Format)] = problem
                else:
                    r.Status, r.OutputFile = 'Exported', path
        if not problem:
            files.append(path)
    rows = batchout_report.build_rows(queue, reasons, model=doc.Title)
    return {'files': files, 'rows': rows, 'notes': notes}


def write_report(folder, rows, s):
    """BatchOut report next to the files; returns its path or an error text."""
    try:
        path = batchout_report.report_path(folder, s.get('report_format'))
        return batchout_report.write_report(path, rows, s.get('report_columns'),
                                            s.get('report_format')), None
    except Exception as ex:
        return None, 'report not written: {}'.format(ex)
