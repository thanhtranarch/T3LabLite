# -*- coding: utf-8 -*-
"""core/tools/rebar_assembly.py: inspect_rebar, rebar_tools, export_bvbs.

Run:  python lite_guard/test_tools_rebar_domain.py   (no Revit: Snippets Revit calls are faked;
the rules, planning, BVBS reading and the .abs writer are the shipped code)
"""

import json
import os
import shutil
import sys
import tempfile
import types

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_rebar_domain_test_')

from core import tool_specs as TS                        # noqa: E402
from core import write_verify as W                       # noqa: E402
from core.tools import rebar_assembly as RA              # noqa: E402
from Snippets import _assembly as ASM                    # noqa: E402
from Snippets import _rebar as REB                       # noqa: E402
from Snippets import _rebar_check as RC                  # noqa: E402
from Snippets import _drawing_clone as DC                # noqa: E402
from Snippets import _bvbs_read as RD                    # noqa: E402
import Snippets._compat as C                             # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ── fake Revit: one TransactionGroup log, ids are ints ──────────────────────

class Group(object):
    log = []

    def __init__(self, doc, name):
        Group.log.append(('new', name))
        self.ended = False

    def Start(self):
        Group.log.append('start')

    def Assimilate(self):
        Group.log.append('assimilate')
        self.ended = True

    def RollBack(self):
        Group.log.append('rollback')
        self.ended = True

    def HasStarted(self):
        return True

    def HasEnded(self):
        return self.ended

    def Dispose(self):
        pass


db = types.ModuleType('Autodesk.Revit.DB')
db.TransactionGroup = Group
sys.modules.setdefault('Autodesk', types.ModuleType('Autodesk'))
sys.modules.setdefault('Autodesk.Revit', types.ModuleType('Autodesk.Revit'))
sys.modules['Autodesk.Revit.DB'] = db
C.make_eid = lambda v: v
RD.make_eid = lambda v: v


class Doc(object):
    def __init__(self, elements=None, path=''):
        self.elements = dict(elements or {})
        self.PathName = path
        self.ProjectInformation = types.SimpleNamespace(Number='P-100')

    def GetElement(self, i):
        return self.elements.get(i)


Bar = ASM.RebarRecord
Asm = ASM.AssemblyRecord
Host = ASM.HostRecord


def index_of(rows):
    idx = REB.RebarIndex()
    for r in rows:
        idx.add(r)
    return idx


def run(name, doc, args):
    return TS.run(name, doc, args)


# ── helpers ─────────────────────────────────────────────────────────────────
print('[helpers]')
check('ints: numbers kept, junk dropped', RA.ints([1, '2', 'x', None]) == [1, 2]
      and RA.ints(5) == [5] and RA.ints(None) == [])
check('natural sort', sorted(['C-10', 'C-2', 'c-1'], key=RA.natural) == ['c-1', 'C-2', 'C-10'])
recs = [Asm(id=70, mark='C-01', type_id=1), Asm(id=71, mark='C-01', type_id=1),
        Asm(id=72, mark='B-01', type_id=2)]
found, missing = RA.find_assemblies(recs, ['c-01', 72, 'nope', '70'])
check('find_assemblies: mark hits every instance, id, missing reported, no duplicates',
      [r.id for r in found] == [70, 71, 72] and missing == ['nope'], (found, missing))
check('parse_series: Cast Unit defaults', RA.parse_series(None) == (('CU-', 1, 1, 3), None))
check('parse_series: custom', RA.parse_series({'prefix': 'B', 'start': 5, 'step': 5,
                                               'digits': 0})[0] == ('B', 5, 5, 0))
check('parse_series: step 0 / digits 9 / text rejected',
      all(RA.parse_series(v)[1] for v in ({'step': 0}, {'digits': 9}, {'start': 'a'})))
check('type split note only when Revit split or merged',
      RA.type_split_note(0, 0, [], []) is None
      and 'split 1' in RA.type_split_note(1, 0, [70], ['C-01']))
rec = Asm(id=70, members=[10, 11, 1], rebar_ids=[1])
by_host = {10: [Bar(id=2, host_id=10), Bar(id=3, host_id=10, assembly_id=99),
                Bar(id=4, host_id=10, in_group=True)], 11: [Bar(id=5, host_id=11)]}
check('unsynced_rebar_ids: loose, not grouped, not in another assembly',
      ASM.unsynced_rebar_ids(rec, by_host) == [2, 5])

# ── declarations ────────────────────────────────────────────────────────────
print('[declarations]')
SPECS = dict((n, TS.get(n)) for n in ('inspect_rebar', 'rebar_tools', 'export_bvbs'))
for n, spec in sorted(SPECS.items()):
    size = len(json.dumps(spec.registry_entry())) // 4
    check('{}: declared, complete, {} tokens <= 350'.format(n, size),
          spec is not None and not spec.problems() and size <= 350, spec and spec.problems())
check('effects: read / write / file', (SPECS['inspect_rebar'].effect, SPECS['rebar_tools'].effect,
                                       SPECS['export_bvbs'].effect) == ('read', 'write', 'file'))
check('rebar_tools: replace_annotations is destructive, dry run is not',
      TS.is_destructive('rebar_tools', {'operation': 'clone_drawing',
                                        'existing': 'replace_annotations'})
      and not TS.is_destructive('rebar_tools', {'existing': 'replace_annotations',
                                                'dry_run': True})
      and not TS.is_destructive('rebar_tools', {'operation': 'create_cast_units'}))
check('export_bvbs: overwrite is destructive', TS.is_destructive('export_bvbs', {'overwrite': True})
      and not TS.is_destructive('export_bvbs', {}))
check('rebar_tools never claims verify=ids (clone replace deletes)',
      'ids' not in SPECS['rebar_tools'].verify and 'self' in SPECS['rebar_tools'].verify)
check('unknown operation', 'operation must be' in run('rebar_tools', Doc(), {'operation': 'x'})['error']
      and 'operation must be' in run('inspect_rebar', Doc(), {'operation': 'x'})['error'])

# ── inspect_rebar ───────────────────────────────────────────────────────────
print('[inspect_rebar]')
ROWS = [Bar(id=1, host_id=10, partition='P1', number='1', shape='M_00'),   # host 10 in 70: unsynced
        Bar(id=2, host_id=99, partition='P1', number='2', shape='M_00'),   # host gone
        Bar(id=3, host_id=11, number='3', shape='M_00')]                   # no partition
HOSTS = {10: Host(id=10, name='B1', assembly_id=70), 11: Host(id=11, name='B2')}
ASMS = [Asm(id=70, mark='C-01', type_id=1, members=[10])]                  # no drawing


def fake_scan(doc, progress=None):
    data = RC.ScanData()
    data.index, data.hosts, data.assemblies = index_of(ROWS), dict(HOSTS), list(ASMS)
    return data


RC.scan = fake_scan
deep_calls = []
RC.collect_legs = lambda doc, ids, progress=None: (deep_calls.append('legs') or {}, False)
RC.collect_boxes = lambda doc, rows, hosts, progress=None: (deep_calls.append('boxes')
                                                            or RC.BoxData(), False)
r = run('inspect_rebar', Doc(), {})
check('check: every rule ran on the scan', r['by_check'].get('host') == 1
      and r['by_check'].get('member') == 1 and r['by_check'].get('drawing') == 1
      and r['by_check'].get('partition') == 1, r.get('by_check'))
check('check: summary + rows', r['summary']['total'] == 3 and r['issues_shown'] == len(r['issues'])
      and r['issues'][0]['check'] == 'host' and r['issues'][0]['id'] == 2, r['issues'][:2])
check('check: Sync fix points at rebar_tools', r['fix']['sync']['assembly_ids'] == [70]
      and 'sync_assembly_rebar' in r['fix']['sync']['next'], r.get('fix'))
check('check: not deep -> note, no geometry read', 'note' in r and not deep_calls)
r = run('inspect_rebar', Doc(), {'check': 'member', 'deep': True})
check('check filter + deep reads geometry', [i['check'] for i in r['issues']] == ['member']
      and deep_calls == ['boxes'] and 'note' not in r, (r['issues'], deep_calls))
check('bad check / scope', 'check must be' in run('inspect_rebar', Doc(), {'check': 'x'})['error']
      and 'scope must be' in run('inspect_rebar', Doc(), {'scope': 'x'})['error'])

ASM.collect_assemblies = lambda doc: [Asm(id=70, mark='C-10', type_id=1, members=[10, 1],
                                          rebar_ids=[1], view_ids=[500]),
                                      Asm(id=71, mark='C-2', type_id=2, members=[11])]
REB.build_rebar_index = lambda doc, progress=None: index_of(
    [Bar(id=1, host_id=10, assembly_id=70), Bar(id=2, host_id=10), Bar(id=3, host_id=11)])
r = run('inspect_rebar', Doc(), {'operation': 'assemblies'})
check('assemblies: natural order, unsynced counted',
      [(a['mark'], a['unsynced_rebar'], a['views']) for a in r['assemblies']]
      == [('C-2', 1, 0), ('C-10', 1, 1)], r)
check('assemblies: unknown mark -> error', 'No assembly' in run(
    'inspect_rebar', Doc(), {'operation': 'assemblies', 'assemblies': ['Z']})['error'])

# ── create_cast_units ───────────────────────────────────────────────────────
print('[create_cast_units]')
HOST_RECS = [Host(id=10, name='B1'), Host(id=11, name='B2'), Host(id=12, name='B3', in_group=True)]
ASM.host_records_from_ids = lambda doc, ids: [h for h in HOST_RECS if h.id in ids]
TYPES = {}
ASM.assembly_type_map = lambda doc: dict(TYPES)
made = []


def fake_batch(doc, plans, progress=None, prefix='', start=1, step=1, digits=3, created=None):
    out = []
    for p in plans:
        if not p.host_ids:
            out.append(ASM.Result(p.label, 'skipped', 0, p.skips[0][1]))
            continue
        new = 900 + len(made)
        made.append(new)
        created.append(new)
        TYPES[new] = 5000 + new
        out.append(ASM.Result(p.label, 'ok', 1 + len(p.rebar_ids), 'assembly {}'.format(new)))
    return out


ASM.batch_create = fake_batch
base = {'operation': 'create_cast_units'}
r = run('rebar_tools', Doc(), dict(base, element_ids=[10, 11, 12], dry_run=True,
                                   series={'prefix': 'C-'}))
check('create dry run: plan per host, grouped host skipped, series names, nothing made',
      r['dry_run'] and r['assemblies'] == 2 and not made
      and [p.get('name') for p in r['plans']] == ['C-001', 'C-002', None]
      and r['plans'][2]['skipped'] == ASM.SKIP_IN_GROUP and r['plans'][0]['rebar'] == 1, r)
Group.log[:] = []
r = run('rebar_tools', Doc(), dict(base, element_ids=[10, 11]))
check('create: one TransactionGroup, assimilated, ids reported',
      Group.log == [('new', 'T3Lab: Create assemblies'), 'start', 'assimilate']
      and r['created_assembly_ids'] == made and r['success'] and r['operation'] ==
      'create_cast_units' and 'not_in_model_after' not in r, (Group.log, r))
check('create: write_verify passes', W.verify('rebar_tools', {}, r, 0, lambda i: True)['verified'])
ASM.assembly_type_map = lambda doc: {}
r = run('rebar_tools', Doc(), dict(base, element_ids=[10]))
check('create: an assembly missing afterwards is not_in_model_after -> verify fails',
      r['not_in_model_after'] and not W.verify('rebar_tools', {}, r, 0,
                                               lambda i: True)['verified'], r)
check('create: no ids / no host / bad series -> errors',
      'element_ids is required' in run('rebar_tools', Doc(), base)['error']
      and 'valid rebar host' in run('rebar_tools', Doc(), dict(base, element_ids=[5]))['error']
      and 'series' in run('rebar_tools', Doc(), dict(base, element_ids=[10],
                                                     series={'step': 0}))['error'])
r = run('rebar_tools', Doc(), dict(base, element_ids=[12]))
check('create: every host skipped -> error, nothing applied', not r['applied'] and 'error' in r)


def boom(*a, **k):
    raise RuntimeError('Revit said no')


ASM.batch_create = boom
Group.log[:] = []
r = run('rebar_tools', Doc(), dict(base, element_ids=[10]))
check('create: exception -> rolled back, error dict', Group.log[-1] == 'rollback'
      and r['error'] == 'Revit said no', (Group.log, r))

# ── sync_assembly_rebar ─────────────────────────────────────────────────────
print('[sync_assembly_rebar]')
MEMBERS = {70: [10, 1], 71: [11]}
ASM.collect_assemblies = lambda doc: [Asm(id=70, mark='C-01', type_id=1, members=list(MEMBERS[70]),
                                          rebar_ids=[1]),
                                      Asm(id=71, mark='C-02', type_id=2, members=list(MEMBERS[71]))]
IDX = index_of([Bar(id=1, host_id=10, assembly_id=70), Bar(id=2, host_id=10),
                Bar(id=3, host_id=11), Bar(id=4, host_id=11)])
REB.build_rebar_index = lambda doc, progress=None: IDX


def fake_sync(doc, assembly_ids, rebar_index, progress=None, land=True):
    out = []
    for a in assembly_ids:
        hosts = [m for m in MEMBERS[a] if m in (10, 11)]
        rows = [r for h in hosts for r in rebar_index.by_host.get(h, ()) if r.assembly_id < 0]
        for row in rows:
            row.assembly_id = a
            if land and a == 70:
                MEMBERS[a].append(row.id)
        out.append(ASM.Result('C-0{}'.format(a - 69), 'ok', len(rows), 'added'))
    return out


ASM.sync_rebar = fake_sync
s = {'operation': 'sync_assembly_rebar'}
r = run('rebar_tools', Doc(), dict(s, dry_run=True))
check('sync dry run: default = every assembly with loose rebar',
      [(p['assembly'], p['unsynced_rebar']) for p in r['plan']] == [('C-01', 1), ('C-02', 2)], r)
r = run('rebar_tools', Doc(), dict(s, assemblies=['c-01']))
check('sync: by mark, added counted, read back clean', r['rebar_added'] == 1
      and 'not_in_model_after' not in r and Group.log[-1] == 'assimilate', r)
r = run('rebar_tools', Doc(), dict(s, assemblies=[71]))
check('sync: rebar reported added but not a member -> not_in_model_after',
      r['not_in_model_after'] == ['rebar 3 in assembly 71', 'rebar 4 in assembly 71'], r)
r = run('rebar_tools', Doc(), s)
check('sync: nothing loose left -> message, not applied', not r['applied'] and 'message' in r, r)

# ── rename_assemblies ───────────────────────────────────────────────────────
print('[rename_assemblies]')
MARKS = {1: 'X-10', 2: 'X-2', 3: 'Y'}
ASM.collect_assemblies = lambda doc: [Asm(id=70, mark=MARKS[1], type_id=1),
                                      Asm(id=71, mark=MARKS[1], type_id=1),
                                      Asm(id=72, mark=MARKS[2], type_id=2),
                                      Asm(id=73, mark=MARKS[3], type_id=3)]


def fake_rename(doc, type_ids, prefix, start, step, digits):
    names = ASM.series_names(len(type_ids), prefix, start, step, digits)
    out = []
    for t, n in zip(type_ids, names):
        if t != 3:
            MARKS[t] = n
        out.append(ASM.Result(n, 'ok', 1, 'renamed'))
    return out, (0, 0, [])


ASM.rename_series = fake_rename
n = {'operation': 'rename_assemblies'}
r = run('rebar_tools', Doc(), dict(n, assemblies=['x-10', 'X-2', 71], dry_run=True))
check('rename dry run: one name per type, natural order of marks',
      [(p['from'], p['to']) for p in r['plan']] == [('X-2', 'CU-001'), ('X-10', 'CU-002')], r)
r = run('rebar_tools', Doc(), dict(n, assemblies=['X-10', 'X-2', 'Y'], series={'prefix': 'Z'}))
check('rename: read back finds the type Revit did not rename',
      r['renamed'] == 3 and r['not_in_model_after'] == ['Z003'], r)
check('rename: assemblies required', 'required' in run('rebar_tools', Doc(), n)['error'])

# ── clone_drawing ───────────────────────────────────────────────────────────
print('[clone_drawing]')
ASM.collect_assemblies = lambda doc: [
    Asm(id=70, mark='C-01', type_id=1, naming_category='Columns', view_ids=[500], sheet_ids=[600]),
    Asm(id=71, mark='C-02', type_id=2, naming_category='Columns'),
    Asm(id=72, mark='C-03', type_id=3, naming_category='Columns', view_ids=[501]),
    Asm(id=73, mark='B-01', type_id=4, naming_category='Framing'),
    Asm(id=74, mark='C-01', type_id=1, naming_category='Columns')]
VIEWS_AFTER = {70: ([500], [600])}
ASM.collect_assembly_views = lambda doc: dict(VIEWS_AFTER)
DC.assembly_profile = lambda doc, el: {'category': 'Framing' if el == 73 else 'Columns',
                                       'bbox_mm': (400, 400, 3000), 'rebar_count': 8,
                                       'member_count': 9}
DC.read_source = lambda doc, el, idx=None: [DC.ViewSpec(kind='3d', src_view_id=500),
                                            DC.ViewSpec(kind='sheet', src_view_id=600,
                                                        sheet_number='S-1')]
DC.load_presets = lambda path: ({'Danger': DC.CloneSettings(existing=DC.EXISTING_REPLACE)},
                                'Tekla default', [])
DC.replaceable_ids = lambda doc, el, idx=None: [(501, [1, 2, 3])]
cl = {'operation': 'clone_drawing', 'source': 'C-01'}
r = run('rebar_tools', Doc(), dict(cl, dry_run=True))
check('clone dry run: similar clonable targets only (same type, other category, has drawing out)',
      [t['assembly'] for t in r['targets']] == ['C-02'] and r['dry_run'] and 'Clone' in r['plan'], r)
r = run('rebar_tools', Doc(), dict(cl, dry_run=True, assemblies=['C-03', 'B-01']))
check('clone dry run: named targets that all skip -> error listing each status',
      'No target' in r['error'] and [t['status'] for t in r['targets']]
      == ['Will skip: different category', 'Will skip: has drawing'], r)
r = run('rebar_tools', Doc(), dict(cl, assemblies=['B-01']))
check('clone: only unclonable targets -> error with the list', 'No target' in r['error']
      and r['targets'][0]['assembly'] == 'B-01', r)
r = run('rebar_tools', Doc(), dict(cl, dry_run=True, assemblies=['C-03'],
                                   existing='replace_annotations'))
check('clone dry run: replace counts the annotations it would delete',
      r['annotations_to_delete'] == 3 and r['existing'] == 'replace_annotations', r)
r = run('rebar_tools', Doc(), dict(cl, preset='Danger'))
check('clone: a preset that replaces needs existing=replace_annotations explicitly',
      'replaces annotations' in r['error'], r)
r = run('rebar_tools', Doc(), dict(cl, settings={'tags': 'maybe', 'color': 'red'}))
check('clone: invalid settings named with their choices', 'tags' in r['error']
      and 'color' in r['error'] and 'tags=clone|clone_create|skip' in r['error'], r)
r = run('rebar_tools', Doc(), dict(cl, preset='Nope'))
check('clone: unknown preset lists presets', 'Tekla default' in r['error'] and 'Danger' in r['error'])
r = run('rebar_tools', Doc(), dict(cl, source='C-02'))
check('clone: a source without views -> error with sources', r['sources'] == ['C-01 (70)', 'C-03 (72)'], r)


def fake_clone(doc, src_id, target_ids, settings, progress=None, replace_plan=None):
    outs = []
    for t in target_ids:
        o = DC.TargetOutcome(t, 'C-0{}'.format(t - 69))
        o.views_created, o.sheet_id, o.sheet_number = 1, 601, 'S-1-C-02'
        o.unmatched = [(42, 'tag: no matching element in C-02')]
        VIEWS_AFTER[t] = ([700], [601])
        outs.append(o)
    return outs


DC.run_clone = fake_clone
r = run('rebar_tools', Doc({601: object()}), cl)
check('clone: outcome per target, read back clean', r['success'] and r['totals']['ok'] == 1
      and r['targets'][0]['sheet'] == 'S-1-C-02' and r['targets'][0]['unmatched']
      and 'not_in_model_after' not in r, r)
r = run('rebar_tools', Doc({}), cl)
check('clone: sheet reported but missing -> not_in_model_after',
      r['not_in_model_after'] == ['C-02: sheet S-1-C-02'], r)

# ── export_bvbs (real reader + writer, fake rebar) ──────────────────────────
print('[export_bvbs]')


def line_curves(*points):
    return [('Line', a, b) for a, b in zip(points, points[1:])]


L_BAR = line_curves((0, 0, 0), (394, 0, 0), (394, 594, 0))          # outer 400 / 600 at d=12


class FakeRebar(object):
    def __init__(self, i, bars, curves=L_BAR, length_mm=None):
        self.Id, self.positions = i, dict((k, curves) for k in range(bars))
        self.NumberOfBarPositions, self.length_mm, self.Style = bars, length_mm, 'Standard'

    def DoesBarExistAtPosition(self, k):
        return k in self.positions


class BarType(object):
    Id = 900
    StandardBendDiameter = 48 / 304.8
    StirrupTieBendDiameter = 32 / 304.8


REB.centerline_points = lambda el, pos=0: el.positions[pos]
REB.bar_type_of = lambda doc, el: BarType()
RD.revit_length_mm = lambda el: el.length_mm
RD.bar_mass_per_metre = lambda bar_type: (None, None)
ASM.collect_assemblies = lambda doc: [Asm(id=70, mark='C-01', type_id=1, members=[10])]
REB.build_rebar_index = lambda doc, progress=None: index_of([
    Bar(id=1, host_id=10, assembly_id=70, partition='C1', number='1', diameter_mm=12.0,
        quantity=4, is_shape_driven=True),
    Bar(id=2, host_id=11, partition='C1', number='2', diameter_mm=12.0, quantity=2,
        is_shape_driven=True),
    Bar(id=3, host_id=11, partition='C2', number='3', diameter_mm=12.0, quantity=1,
        is_shape_driven=True),
    Bar(id=5, host_id=11, kind='AreaReinforcement')])
out_dir = tempfile.mkdtemp(prefix='t3lab_bvbs_tool_')
bdoc = Doc({1: FakeRebar(1, 4), 2: FakeRebar(2, 2), 3: FakeRebar(3, 1, length_mm=1100)})
r = run('export_bvbs', bdoc, {'output_folder': out_dir, 'dry_run': True})
check('dry run: rows, one file, flagged bar held back, other kinds noted, nothing written',
      r['files'] == ['MODEL.abs (2 marks)'] and r['bars'] == 6 and len(r['rows']) == 3
      and r['held_back'][0]['mark'] == 'C2-3' and r['notes'] and not os.listdir(out_dir), r)
r = run('export_bvbs', bdoc, {'output_folder': out_dir})
path = os.path.join(out_dir, 'MODEL.abs')
check('export: file written, listed, read back', r['success'] and r['files'] == [path]
      and r['exported'] == 1 and r['bars'] == 6 and os.path.isfile(path), r)
check('export: write_verify reads the file', W.verify('export_bvbs', {}, r, 0,
                                                      lambda i: True)['verified'])
with open(path, 'rb') as fh:
    data = fh.read()
check('export: BF2D lines, project number, CR LF', data.count(b'BF2D@HjP-100@rMODEL@ia@')
      == 2 and data.endswith(b'\r\n') and b'@gB500B@' in data, data[:120])
r = run('export_bvbs', bdoc, {'output_folder': out_dir})
check('export: existing file kept without overwrite', 'already exist' in r['error'], r)
r = run('export_bvbs', bdoc, {'output_folder': out_dir, 'overwrite': True,
                              'include_flagged': True, 'grade': 'CB400'})
check('export: overwrite + include_flagged writes all 3 marks', r['success'] and r['bars'] == 7
      and 'held_back' not in r and b'@gCB400@' in open(path, 'rb').read(), r)
r = run('export_bvbs', bdoc, {'output_folder': out_dir, 'assembly': 'c-01', 'per_assembly': True,
                              'plan': 'P'})
check('export: assembly scope, one file per assembly', r['files'] == [
    os.path.join(out_dir, 'P_C-01.abs')] and r['bars'] == 4, r)
r = run('export_bvbs', bdoc, {'output_folder': out_dir, 'partition': 'C9', 'dry_run': True})
check('export: unknown partition lists partitions', 'C1, C2' in r['error'], r)
r = run('export_bvbs', bdoc, {'output_folder': os.path.join(out_dir, 'new'),
                              'element_ids': [11], 'include_flagged': True})
check('export: element_ids (a host) scope, folder created', r['bars'] == 3
      and os.path.isfile(r['files'][0]), r)
r = run('export_bvbs', bdoc, {'output_folder': out_dir, 'element_ids': [999]})
check('export: nothing in scope -> error, no file', 'Nothing to export' in r['error']
      and not r.get('files'), r)
shutil.rmtree(out_dir, ignore_errors=True)

print()
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
