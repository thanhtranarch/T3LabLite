# -*- coding: utf-8 -*-
"""Services/sheet_export.py + MCP export_sheets + Assistant sheet sets (BatchOut reuse).

Run:  python lite_guard/test_sheet_export.py      (fake Revit API that writes real files)
"""

import ast
import json
import os
import sys
import tempfile
import time
import types

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_sheet_export_test_')

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ── fake Revit API ──────────────────────────────────────────────────────────

class Id(object):
    def __init__(self, v):
        self.Value = v


class ViewSheet(object):
    def __init__(self, i, number, name, rev='', placeholder=False, params=None):
        self.Id, self.SheetNumber, self.Name = Id(i), number, name
        self.IsPlaceholder, self.rev, self.params = placeholder, rev, params or {}

    def get_Parameter(self, bip):
        return Param(self.rev)

    def LookupParameter(self, name):
        return Param(self.params[name]) if name in self.params else None


class View(object):
    def __init__(self, i):
        self.Id = Id(i)


class Param(object):
    def __init__(self, v):
        self.v = v

    def AsString(self):
        return self.v

    def AsValueString(self):
        return self.v


class ViewSheetSet(object):
    def __init__(self, name, views):
        self.Name, self.Views = name, views


class Collector(object):
    def __init__(self, doc):
        self.doc = doc

    def OfClass(self, cls):
        return [e for e in self.doc.elements if isinstance(e, cls)]


class Opts(object):
    setups = {}

    def __init__(self):
        self.setup = None

    @classmethod
    def GetPredefinedOptions(cls, doc, name):
        if name not in cls.setups:
            raise Exception('no setup ' + name)
        o = cls()
        o.setup = name
        return o


class PDFExportOptions(object):
    pass


class DWGExportOptions(Opts):
    Colors = PropOverrides = LineScaling = FileVersion = ExportingAreas = MergedViews = None
    setups = {'AIA'}


class DXFExportOptions(DWGExportOptions):
    pass


class DGNExportOptions(Opts):
    setups = {'V8'}


ns = types.SimpleNamespace
DB = ns(ViewSheet=ViewSheet, ViewSheetSet=ViewSheetSet, FilteredElementCollector=Collector,
        PDFExportOptions=PDFExportOptions, DWGExportOptions=DWGExportOptions,
        DXFExportOptions=DXFExportOptions, DGNExportOptions=DGNExportOptions,
        ExportPaperFormat=ns(Default='default'), RasterQualityType=ns(High='high'),
        ColorDepthType=ns(Color='color'), ACADVersion=ns(R2013='2013', R2018='2018'),
        ExportingAreas=ns(DontExportViewsOnSheets='no', ExportViewsOnSheets='yes'),
        ExportColorMode=ns(TrueColorPerView='tcpv'), PropOverrideMode=ns(ByEntity='entity'),
        LineScaling=ns(PaperSpace='paper', ViewScale='view'),
        BaseExportOptions=ns(GetPredefinedSetupNames=lambda doc: ['V8', 'AIA']),
        BuiltInParameter=ns(SHEET_CURRENT_REVISION=1), ElementId=Id, Color=lambda *a: a)

gen = types.ModuleType('System.Collections.Generic')


class _List(list):
    def Add(self, x):
        self.append(x)


gen.List = {Id: _List}
sys.modules['System'] = types.ModuleType('System')
sys.modules['System.Collections'] = types.ModuleType('System.Collections')
sys.modules['System.Collections.Generic'] = gen


class Doc(object):
    Title = 'Tower.rvt'
    ProjectInformation = ns(Number='P01', Name='Tower', Address='', ClientName='', Status='')

    def __init__(self, elements):
        self.elements, self.calls = elements, []
        self.broken, self.silent = set(), set()

    def Export(self, folder, *args):
        if isinstance(args[-1], PDFExportOptions):
            ids, opts = args
            name, data = opts.FileName + '.pdf', b'%PDF-1.7'
            self.calls.append(('pdf', [i.Value for i in ids], opts))
        else:
            name, ids, opts = args
            ext = {'DWGExportOptions': '.dwg', 'DXFExportOptions': '.dxf',
                   'DGNExportOptions': '.dgn'}[type(opts).__name__]
            data = {'.dwg': b'AC1032', '.dxf': b'0\nSECTION', '.dgn': b'\x08\x09'}[ext]
            name += ext
            self.calls.append((ext, [i.Value for i in ids], opts))
        first = ids[0].Value
        if first in self.broken:
            raise Exception('Revit export failed\nstack')
        if first not in self.silent:
            with open(os.path.join(folder, name), 'wb') as f:
                f.write(data)
            if ext_is_dwg(name):
                with open(os.path.join(folder, 'logo.png'), 'wb') as f:
                    f.write(b'\x89PNG')
        return True


def ext_is_dwg(name):
    return name.endswith('.dwg')


from Services import sheet_export as SE   # noqa: E402
SE.DB = DB

A101 = ViewSheet(101, 'A-101', 'Ground Plan', rev='P02', params={'Drawn By': 'TT'})
A102 = ViewSheet(102, 'A-102', 'First/Plan')
A110 = ViewSheet(110, 'A-110', 'Roof')
S201 = ViewSheet(201, 'S-201', 'Beams')
PH = ViewSheet(301, 'X-1', 'Placeholder', placeholder=True)
SETS = [ViewSheetSet('Kiến trúc', [A110, A101, View(900)]), ViewSheetSet('Kết cấu', [S201])]
doc = Doc([A101, A102, A110, S201, PH] + SETS)

# ── pure ────────────────────────────────────────────────────────────────────
print('[pure]')
s = SE.settings_from()
check('defaults = BatchOut', s['pattern'] == '{SheetNumber}-{SheetName}'
      and s['match_display'] and s['split_by_format'] and not s['combined'])
prof = {'Name': 'Issue', 'NamingPattern': '{ProjectNumber}_{SheetNumber}', 'CombinePDF': True,
        'CADExportSetup': 'Use setup from file', 'OutputFolder': 'D:/out', 'SplitByFormat': False,
        'DWGVersion': 'AutoCAD 2018', 'WriteReport': True, 'ReportFormat': 'CSV'}
s = SE.settings_from(prof, pattern=None, combined=False)
check('profile mapped, overrides win, "Use setup from file" = none',
      s['pattern'] == '{ProjectNumber}_{SheetNumber}' and s['combined'] is False
      and s['export_setup'] == '' and s['folder'] == 'D:/out' and s['report'], s)
check('output folder: BatchOut subfolder only when split',
      SE.output_folder(SE.settings_from(), 'dwg').endswith(os.path.join('Revit Exports', 'DWG'))
      and SE.output_folder(s, 'pdf') == 'D:/out')
check('file_stem: placeholders, bad chars, missing -> empty',
      SE.file_stem('{SheetNumber}-{SheetName}{Nope}', {'SheetNumber': 'A-102',
                                                       'SheetName': 'First/Plan: "B"'})
      == 'A-102-First_Plan_ _B_')
check('file_stem never empty', SE.file_stem('{Nope}', {}) == 'Unnamed')
used = set()
check('reserve: (2) on collision', [SE.reserve(used, 'f', 'A', '.pdf') for _ in range(3)]
      == ['A', 'A (2)', 'A (3)'])
rows = [{'id': i, 'number': n, 'name': ''} for i, n in
        ((1, 'A-10'), (2, 'A-9'), (3, 'S-1'), (4, 'A-100'))]
kept, miss = SE.select(rows, prefix='a')
check('select: prefix, natural order', [r['number'] for r in kept] == ['A-9', 'A-10', 'A-100'])
kept, miss = SE.select(rows, numbers=['s-1', 'Z-9'], ids=[3, 77])
check('select: missing numbers and ids reported', [r['id'] for r in kept] == [3]
      and miss == ['id 77', 'Z-9'], miss)
kept, _m = SE.select(rows, prefix='A', member_ids=[1, 3])
check('select: sheet set ∩ prefix', [r['id'] for r in kept] == [1])

# ── profiles (BatchOut folder, io.open) ─────────────────────────────────────
print('[profiles]')
pdir = tempfile.mkdtemp(prefix='t3lab_profiles_')
with open(os.path.join(pdir, 'issue.json'), 'w', encoding='utf-8') as f:
    json.dump(dict(prof, Name='Phát hành'), f, ensure_ascii=False)
with open(os.path.join(pdir, '_latest.json'), 'w') as f:
    f.write('{}')
check('profile by name, case-insensitive, unicode',
      SE.load_profile('PHÁT HÀNH', pdir)['NamingPattern'] == '{ProjectNumber}_{SheetNumber}')
check('profile names skip bookkeeping files', SE.profile_names(pdir) == ['Phát hành'])
SE.PROFILES_FOLDER = pdir

# ── Revit reads ─────────────────────────────────────────────────────────────
print('[revit reads]')
check('sheet_rows skip placeholders', [r['number'] for r in SE.sheet_rows(doc)]
      == ['A-101', 'A-102', 'A-110', 'S-201'])
check('sheet_sets with counts', SE.sheet_sets(doc) == [
    {'name': 'Kiến trúc', 'sheets': 2, 'views': 1}, {'name': 'Kết cấu', 'sheets': 1, 'views': 0}],
      SE.sheet_sets(doc))
check('set members by name, case-insensitive', SE.set_member_ids(doc, 'kiến trúc') == ([110, 101], 1))
check('unknown set', SE.set_member_ids(doc, 'nope') == (None, 0))
vals = SE.sheet_values(doc, A101, '{SheetNumber}{Revision}{Drawn By}{ProjectNumber}')
check('sheet_values: revision, any parameter, project', vals['Revision'] == 'P02'
      and vals['Drawn By'] == 'TT' and vals['ProjectNumber'] == 'P01', vals)
o = SE.pdf_options('A', SE.settings_from(prof, hide_crop_boundaries=True))
check('pdf options: Combine for FileName, sheet size, quality ladder, colour, hides',
      o.Combine and o.FileName == 'A' and o.PaperFormat == 'default'
      and o.RasterQuality == 'high' and o.ColorDepth == 'color' and o.HideCropBoundaries)
o, notes = SE.cad_options(doc, 'dwg', SE.settings_from(prof, export_setup='AIA'))
check('dwg: setup, Match Revit display, version, areas, merged',
      o.setup == 'AIA' and o.Colors == 'tcpv' and o.LineScaling == 'paper'
      and o.FileVersion == '2018' and o.ExportingAreas == 'no' and o.MergedViews is True
      and not notes, vars(o))
o, notes = SE.cad_options(doc, 'dgn', SE.settings_from(export_setup='Nope'))
check('dgn: missing setup -> default options + note', o.setup is None and 'Nope' in notes[0])

# ── export ──────────────────────────────────────────────────────────────────
print('[export]')
out = tempfile.mkdtemp(prefix='t3lab_out_')
rows = SE.sheet_rows(doc)
doc.broken, doc.silent = {102}, {110}
res = SE.export(doc, 'pdf', rows[:3], out, SE.settings_from())
check('pdf: one file per sheet, failures reported', [os.path.basename(f) for f in res['files']]
      == ['A-101-Ground Plan.pdf'], res['files'])
status = dict((r['number'], (r['status'], r['reason'])) for r in res['rows'])
check('pdf: export error and missing file named', status['A-102'] == (
    'Failed', 'Revit export failed') and status['A-110'] == ('Failed', 'Revit wrote no file'),
      status)
doc.broken, doc.silent = set(), set()
res = SE.export(doc, 'pdf', rows[:3], out, SE.settings_from(combined=True))
check('pdf combined: one call, first-last name', len(res['files']) == 1
      and res['files'][0].endswith('A-101-A-110_Combined.pdf')
      and doc.calls[-1][1] == [101, 102, 110], doc.calls[-1][:2])
res = SE.export(doc, 'dwg', rows[:2], out, SE.settings_from(skip_images=True, export_setup='AIA'))
check('dwg: image references deleted (batchout_refs)', len(res['files']) == 2
      and not os.path.exists(os.path.join(out, 'logo.png'))
      and res['notes'] == ['1 image reference file(s) deleted'] * 2, res['notes'])
for fmt in ('dxf', 'dgn'):
    res = SE.export(doc, fmt, rows[:1], out, SE.settings_from())
    check('{}: exported and header checked'.format(fmt), len(res['files']) == 1, res)
with open(os.path.join(out, 'S-201-Beams.pdf'), 'wb') as f:
    f.write(b'%PDF old')
old = time.time() - 3600
os.utime(os.path.join(out, 'S-201-Beams.pdf'), (old, old))
doc.silent = {201}
res = SE.export(doc, 'pdf', rows[3:], out, SE.settings_from())
check('stale file from an earlier run is not counted', not res['files']
      and 'not rewritten' in res['rows'][0]['reason'], res['rows'])
doc.silent = set()

# ── MCP tool ────────────────────────────────────────────────────────────────
print('[export_sheets tool]')
from core import tool_specs as H       # noqa: E402
from core import write_verify as W     # noqa: E402
r = H.run('export_sheets', doc, {'format': 'pdf'})
check('no selection -> error + sheet sets', 'Say which sheets' in r['error']
      and r['sheet_sets'][1]['name'] == 'Kết cấu')
r = H.run('export_sheets', doc, {'format': 'jpg', 'all': True})
check('bad format', r['error'].startswith('format must'))
r = H.run('export_sheets', doc, {'format': 'pdf', 'sheet_set': 'Nope'})
check('unknown set -> names', 'No sheet set' in r['error'] and len(r['sheet_sets']) == 2)
r = H.run('export_sheets', doc, {'format': 'pdf', 'sheet_numbers': ['A-999']})
check('unknown sheet number', 'A-999' in r['error'])
r = H.run('export_sheets', doc, {'format': 'pdf', 'profile': 'Nope', 'all': True})
check('unknown profile -> names', r['profiles'] == ['Phát hành'], r)
n0 = len(doc.calls)
dest = tempfile.mkdtemp(prefix='t3lab_dest_')
r = H.run('export_sheets', doc, {'format': 'dwg', 'sheet_set': 'Kiến trúc', 'dry_run': True,
                                 'output_folder': dest})
check('dry run: files, setups, note about views, nothing exported',
      r['files'] == ['A-101-Ground Plan.dwg', 'A-110-Roof.dwg'] and r['export_setups'] == ['AIA', 'V8']
      and 'not sheets' in r['note'] and len(doc.calls) == n0 and r['output_folder'] == dest, r)
r = H.run('export_sheets', doc, {'format': 'pdf', 'sheet_set': 'Kiến trúc', 'prefix': 'A-1',
                                 'profile': 'phát hành', 'combined': False, 'output_folder': dest})
check('run with set + profile naming + report', r['success'] and r['exported'] == 2
      and [os.path.basename(f) for f in r['files']] == ['P01_A-101.pdf', 'P01_A-110.pdf']
      and r['profile'] == 'Phát hành' and r['report'].endswith('.csv')
      and os.path.isfile(r['report']), r)
with open(r['report'], encoding='utf-8-sig') as f:
    report = f.read()
check('report written by batchout_report (BOM, headers, rows)',
      report.startswith('Number,Name') and 'A-110' in report, report[:120])
v = W.verify('export_sheets', {}, r, 0, lambda i: True)
check('write_verify re-checks the files', v == {'verified': True, 'checks': 2}, v)
os.remove(r['files'][0])
check('write_verify catches a vanished file', not W.verify('export_sheets', {}, r, 0,
                                                            lambda i: True)['verified'])
doc.broken = {101}
r = H.run('export_sheets', doc, {'format': 'pdf', 'sheet_numbers': ['A-101'],
                                 'output_folder': dest})
check('all failed -> error, success false', not r['success'] and 'error' in r
      and r['failed'][0]['sheet'] == 'A-101', r)
doc.broken = set()

# ── Assistant: direct_export (BatchOut window) + parser ─────────────────────
print('[assistant]')
from Services import batchout_executor as BE   # noqa: E402
items = [ns(Sheet=sh, IsSelected=True, SheetNumber=sh.SheetNumber, SheetName=sh.Name)
         for sh in (A101, A102, A110)]
window = ns(doc=doc, all_sheets=items)
check('BatchOut window: sheet set narrows the selection',
      BE._apply_sheet_set(window, 'Kiến trúc') is None
      and [i.SheetNumber for i in items if i.IsSelected] == ['A-101', 'A-110'])
check('BatchOut window: unknown set names the sets',
      'Kết cấu' in BE._apply_sheet_set(window, 'Nope'))
check('parse: quoted set name (vi + en)',
      BE.parse_export_params(u'xuất pdf bộ “Kiến trúc”')['sheet_set'] == u'Kiến trúc'
      and BE.parse_export_params('export dwg sheet set "A - Plans"')['sheet_set'] == 'A - Plans')
check('parse: unquoted "toàn bộ" is not a set',
      'sheet_set' not in BE.parse_export_params(u'xuất toàn bộ G sheet pdf'))
from Intelligence import t3lab_assistant as TA   # noqa: E402
check('Assistant intent parser keeps the set',
      TA._parse_export_params(u'xuất pdf bộ "Kết cấu"')['sheet_set'] == u'Kết cấu')

# ── server + IronPython syntax ──────────────────────────────────────────────
print('[server]')
import core.server as S   # noqa: E402
check('export_sheets: a file-output spec on the main thread',
      H.get('export_sheets').effect == 'file' and 'export_sheets' in S.T3LabAIServer._WRITE_TOOLS)
for rel in ('Services/sheet_export.py', 'Services/batchout_profiles.py',
            'Services/batchout_report.py', 'Services/batchout_refs.py', 'core/tools/export.py'):
    text = open(os.path.join(LIB, rel), encoding='utf-8').read()
    tree = ast.parse(text)
    bad = [type(n).__name__ for n in ast.walk(tree)
           if isinstance(n, (ast.JoinedStr, ast.Nonlocal, ast.NamedExpr))
           or (isinstance(n, ast.arguments) and (n.kwonlyargs or n.posonlyargs))]
    bad += ['open(encoding=) line {}'.format(n.lineno) for n in ast.walk(tree)
            if isinstance(n, ast.Call) and isinstance(n.func, ast.Name) and n.func.id == 'open'
            and any(k.arg == 'encoding' for k in n.keywords)]
    check('{}: IronPython-safe'.format(rel), not bad, bad)

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
