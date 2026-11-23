# -*- coding: utf-8 -*-
"""core/tools/schedules.py + core/tools/links.py: schedule_roundtrip, link_models, cad_imports.

schedule_roundtrip runs end to end on the fake Revit API of lite_guard/test_mana_sched_flow.py;
link_models runs on Snippets/_links.py loaded with stubs (lite_guard/test_batch_link.py) and
fakes for what needs Revit; cad_imports on Snippets/_cad_imports.py with a fake DB.

Run:  python lite_guard/test_tools_data_domain.py      (no Revit)
"""

import json
import logging
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
sys.path.insert(0, os.path.join(REPO, 'lite_guard'))
TMP = tempfile.mkdtemp(prefix='t3lab_data_domain_test_')
os.environ['APPDATA'] = TMP

from core import tool_specs as TS              # noqa: E402
from core import write_verify as W             # noqa: E402
from core.tools import links as LK, schedules as SC   # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


class NS(object):
    def __init__(self, **kw):
        self.__dict__.update(kw)


def fresh_modules(fakes):
    """Install fake Revit modules; drop cached Snippets so they import against them."""
    for name in list(sys.modules):
        if name == 'Snippets' or name.startswith('Snippets.') or name.startswith('Autodesk'):
            del sys.modules[name]
    sys.modules.update(fakes)


class FakeGroup(object):
    log = []

    def __init__(self, doc, name):
        self.name = name

    def Start(self):
        FakeGroup.log.append(('start', self.name))

    def Assimilate(self):
        FakeGroup.log.append(('assimilate', self.name))

    def RollBack(self):
        FakeGroup.log.append(('rollback', self.name))

    def Dispose(self):
        pass


# ── declarations ────────────────────────────────────────────────────────────
print('[specs]')
SPECS = TS.all_specs()
for name in ('schedule_roundtrip', 'link_models', 'cad_imports'):
    spec = SPECS.get(name)
    check('{}: declared, complete'.format(name), spec is not None and not spec.problems(),
          spec and spec.problems())
    size = len(json.dumps(spec.registry_entry())) // 4
    check('{}: registry entry {} tokens (<= 350)'.format(name, size), size <= 350)
check('cad_imports delete is destructive, list / dry run are not',
      TS.is_destructive('cad_imports', {'operation': 'delete'})
      and not TS.is_destructive('cad_imports', {'operation': 'delete', 'dry_run': True})
      and not TS.is_destructive('cad_imports', {'operation': 'list'}))
check('link_models never deletes, so ids read-back is allowed',
      'ids' in SPECS['link_models'].verify and not SPECS['link_models'].destructive)
check('element ids accept "@set" handles',
      SPECS['cad_imports'].properties['element_ids']['items']['type'] == ['integer', 'string'])

# ── pure helpers ────────────────────────────────────────────────────────────
print('[helpers]')
scheds = [('Door Schedule', NS(Id=11)), ('Window Schedule', NS(Id=12))]
picked, missing = SC.pick_schedules(scheds, ['door  schedule', 12, '12', 'Nope'], lambda s: s.Id)
check('pick_schedules: name, id, duplicate once, missing reported',
      [n for n, _s in picked] == ['Door Schedule', 'Window Schedule'] and missing == ['Nope'],
      (picked, missing))
check('default_xlsx_path: one schedule named, unsafe chars replaced',
      SC.default_xlsx_path([('A/B: C', None)]).endswith(os.path.join('XLSX', 'A_B_ C.xlsx'))
      and SC.default_xlsx_path(scheds).endswith('Schedules.xlsx'))

folder = os.path.join(TMP, 'models')
os.makedirs(folder)
for n in ('A.rvt', 'A.0001.rvt', 'B.RVT', 'notes.txt'):
    open(os.path.join(folder, n), 'w').close()
files, problems = LK.expand_paths([folder, os.path.join(folder, 'A.rvt'),
                                   os.path.join(folder, 'C.rvt'), os.path.join(folder, 'notes.txt')])
check('expand_paths: folder -> .rvt without backups, duplicates once, problems named',
      [os.path.basename(f) for f in files] == ['A.rvt', 'B.RVT'] and len(problems) == 2
      and 'file not found' in problems[0] and 'not a .rvt' in problems[1], (files, problems))
recs = [NS(name='Arch.rvt', type_id=5), NS(name='Struct.rvt : 2 : location', type_id=6)]
found, missing = LK.match_links(recs, ['arch', 'STRUCT.rvt', 6, 'MEP'])
check('match_links: .rvt optional, " : " suffix, id, missing',
      found == recs and missing == ['MEP'], (found, missing))
ws = [NS(name='Shared Levels', is_open=True), NS(name='Furniture', is_open=False),
      NS(name='Site', is_open=True)]
to_open, to_close, unknown, changes = LK.workset_plan(ws, ['furniture'], ['site', 'Ghost'])
check('workset_plan: case-insensitive, unknown names, only real changes',
      to_open == ['Furniture'] and to_close == ['Site'] and unknown == ['Ghost']
      and changes == ['Furniture', 'Site'], (to_open, to_close, unknown, changes))
_o, _c, _u, changes = LK.workset_plan(ws, ['Shared Levels'], [])
check('workset_plan: already open = no change', changes == [])

# ── schedule_roundtrip (fake Revit of test_mana_sched_flow) ─────────────────
print('[schedule_roundtrip]')
try:
    import openpyxl
    import test_mana_sched_flow as FLOW
except ImportError:
    openpyxl = None
    print('  skip  openpyxl not installed')

if openpyxl is not None:
    DB, ES = FLOW.build_fake_db()
    DB.TransactionGroup = FakeGroup
    autodesk, revit = types.ModuleType('Autodesk'), types.ModuleType('Autodesk.Revit')
    revit.DB, autodesk.Revit, DB.ExtensibleStorage = DB, revit, ES
    fakes = FLOW.build_fake_system()
    fakes.update({'Autodesk': autodesk, 'Autodesk.Revit': revit, 'Autodesk.Revit.DB': DB,
                  'Autodesk.Revit.DB.ExtensibleStorage': ES})
    saved = dict(sys.modules)
    fresh_modules(fakes)
    flow = FLOW.ManaSchedFlowTest
    flow.DB = DB
    fx = flow.__new__(flow)
    doc, sched, _sheet = fx.make_doc()
    from Snippets import schedule_data as SD    # noqa: E402

    r = TS.run('schedule_roundtrip', doc, {'operation': 'export', 'schedules': ['Nope']})
    check('unknown schedule -> error with the model\'s schedules',
          'Nope' in r['error'] and 'Door Schedule' in r['schedules'], r)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'export'})
    check('export without schedules -> error', r['error'].startswith('Say which'), r)
    out = os.path.join(TMP, 'xl', 'doors')
    r = TS.run('schedule_roundtrip', doc, {'operation': 'export', 'schedules': ['door schedule'],
                                           'xlsx_path': out, 'dry_run': True})
    check('export dry run: path + sheet names, no file', r['dry_run'] and not r['applied']
          and r['xlsx_path'].endswith('doors.xlsx') and not os.path.exists(r['xlsx_path'])
          and r['sheets'] == {'Door Schedule': 'Door Schedule'}, r)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'export', 'schedules': [sched.Id.Value],
                                           'xlsx_path': out})
    path = out + '.xlsx'
    check('export by id: folder made, file written, rows counted',
          r.get('success') and r['rows'] == 2 and r['files'] == [path] and os.path.isfile(path), r)
    v = W.verify('schedule_roundtrip', {}, r, 0, lambda i: True)
    check('write_verify: exported xlsx checked', v and v['verified'] and v['checks'] >= 1, v)

    wb = openpyxl.load_workbook(path)
    sheet = wb['Door Schedule']
    sheet['B2'] = 'D01-NEW'
    sheet['C3'] = 1250
    wb.save(path)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'preview', 'xlsx_path': path})
    got = sorted((c['field'], c['old'], c['new']) for c in r['changes'])
    check('preview: the two edited cells, nothing written',
          got == [('Mark', 'D01', 'D01-NEW'), ('Width', '1200', '1250')] and not r['applied']
          and r['elements'] == 2 and 'field_def' not in r['changes'][0], r)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'preview', 'xlsx_path': path,
                                           'schedules': ['Door Schedule']})
    check('preview limited to a schedule', r['change_count'] == 2, r)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'apply', 'xlsx_path': path,
                                           'dry_run': True})
    check('apply dry run = preview', r['dry_run'] and r['change_count'] == 2 and not r['applied'])
    check('no write on dry run', 'D01-NEW' not in [p.value for e in doc.elements.values()
                                                   for p in (getattr(e, 'Parameters', None) or [])
                                                   if hasattr(p, 'value')])
    r = TS.run('schedule_roundtrip', doc, {'operation': 'apply', 'xlsx_path': path})
    doors = [e for e in doc.elements.values() if getattr(e, 'OwnerViewId', None) == sched.Id
             and not isinstance(e, DB.View)]
    check('apply: written, one transaction, read back clean',
          r.get('success') and r['written'] == 2 and 'not_in_model_after' not in r
          and doors[0].Parameters[0].value == 'D01-NEW' and doc.open_tx is None
          and doc.committed.count('T3Lab: Schedule Manager - Apply Excel Changes') == 1, r)
    v = W.verify('schedule_roundtrip', {}, r, 0, lambda i: True)
    check('write_verify: apply self-check passes', v and v['verified'], v)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'apply', 'xlsx_path': path})
    check('apply again: nothing differs, nothing written',
          r['change_count'] == 0 and not r['applied'] and 'error' not in r, r)

    sheet['B2'] = 'D01-LOST'
    wb.save(path)
    real_apply = SD.apply_changes
    SD.apply_changes = lambda d, changes: (len(changes), [], [])   # Revit "accepts", keeps old
    r = TS.run('schedule_roundtrip', doc, {'operation': 'apply', 'xlsx_path': path})
    SD.apply_changes = real_apply
    check('read-back: a value Revit did not keep is reported',
          r.get('not_in_model_after') and 'D01-LOST' in r['not_in_model_after'][0]
          and 'reads "D01-NEW"' in r['not_in_model_after'][0], r)
    v = W.verify('schedule_roundtrip', {}, r, 0, lambda i: True)
    check('write_verify fails on it', v and v['verified'] is False, v)
    check('read_back: "1250 mm" typed, "1250" shown is the same',
          SD._lead_number('1250 mm') == SD._lead_number('1250')
          and SD._lead_number('D01') != SD._lead_number('D01'))

    # ── table / update_tables ───────────────────────────────────────────────
    book = os.path.join(TMP, 'table.xlsx')
    fx.make_workbook(book)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'table', 'xlsx_path': book,
                                           'kind': 'chart'})
    check('table: bad kind', r['error'].startswith('kind must'), r)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'table', 'xlsx_path': book,
                                           'sheet': 'Doors', 'range': 'A2:C4', 'dry_run': True,
                                           'name': 'Door Table'})
    check('table dry run: size, range, name, nothing built',
          r['dry_run'] and r['rows'] == 3 and r['cols'] == 3 and r['range'] == 'A2:C4'
          and r['view_name'] == 'Door Table' and not doc.committed[-1].startswith('T3Lab: Excel'),
          r)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'table', 'xlsx_path': book,
                                           'range': 'DoorTable', 'name': 'Door Table'})
    view = doc.GetElement(r.get('view_id'))
    check('table: linked drafting view built in one transaction',
          r.get('success') and view is not None and view.Name == 'Door Table'
          and doc.committed[-1] == 'T3Lab: Excel Table - Create' and 'not_in_model_after' not in r,
          r)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'table', 'xlsx_path': book,
                                           'sheet': 'Nope'})
    check('table: unknown sheet -> error', 'error' in r and 'Nope' in r['error'], r)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'update_tables', 'name': 'Ghost'})
    check('update_tables: unknown name lists the linked tables',
          'Ghost' in r['error'] and r['tables'] == ['Door Table'], r)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'update_tables', 'dry_run': True})
    check('update_tables dry run: status per table', r['dry_run']
          and r['tables'][0]['status'] == 'Up to date', r)
    del FakeGroup.log[:]
    r = TS.run('schedule_roundtrip', doc, {'operation': 'update_tables', 'name': 'door table'})
    check('update_tables: redrawn inside one TransactionGroup',
          r.get('success') and r['updated'] == 1 and FakeGroup.log == [
              ('start', 'T3Lab: Excel Table - Update'), ('assimilate', 'T3Lab: Excel Table - Update')]
          and doc.committed[-1] == 'T3Lab: Excel Table - Update', (r, FakeGroup.log))
    r = TS.run('schedule_roundtrip', doc, {'operation': 'preview', 'xlsx_path': 'C:/nope.xlsx'})
    check('missing workbook -> what next', 'Export first' in r['error'], r)
    r = TS.run('schedule_roundtrip', doc, {'operation': 'merge'})
    check('bad operation', r['error'].startswith('operation must'), r)

    sys.modules.clear()
    sys.modules.update(saved)

# ── link_models (Snippets/_links.py with stubs + fakes) ─────────────────────
print('[link_models]')
import test_batch_link as BL                    # noqa: E402

L = BL._load_links_module()
check('placement enum = Snippets._links.PLACEMENTS', tuple(LK.PLACEMENTS) == L.PLACEMENTS)


# link_one itself, on a fake transaction / RevitLinkType
class Tx(object):
    log = []

    def __init__(self, doc, name):
        self.name, self.state = name, 'new'

    def Start(self):
        self.state = 'started'

    def GetFailureHandlingOptions(self):
        return NS(SetFailuresPreprocessor=lambda p: None, SetClearAfterRollback=lambda b: None)

    def SetFailureHandlingOptions(self, o):
        pass

    def Commit(self):
        self.state = 'ended'
        Tx.log.append(('commit', self.name))

    def RollBack(self):
        self.state = 'ended'
        Tx.log.append(('rollback', self.name))

    def HasStarted(self):
        return self.state != 'new'

    def HasEnded(self):
        return self.state == 'ended'


class Eid(object):
    def __init__(self, v):
        self.Value = v

    def __eq__(self, o):
        return isinstance(o, Eid) and o.Value == self.Value

    def __ne__(self, o):
        return not self == o


LOADED, FAILED = 'LinkLoaded', 'LinkNotFound'
lt_elem = NS(AttachmentType=None)
inst_made = []
L.Transaction = Tx
L.ElementId = NS(InvalidElementId=Eid(-1))
L.LinkLoadResultType = NS(LinkLoaded=LOADED, UsedExisting='UsedExisting')
L.AttachmentType = NS(Overlay='Overlay', Attachment='Attachment')
L.ImportPlacement = NS(Origin='O', Shared='S', Site='Si', Centered='C')
L.RevitLinkOptions = lambda rel: ('opts', rel)
L.WarningSwallower = lambda: None
L.RevitLinkType = NS(Create=lambda doc, mp, opts: NS(LoadResult=mp['result'],
                                                     ElementId=Eid(mp.get('type', 50))))


def make_instance(doc, tid, placement):
    inst = NS(Id=Eid(60 + len(inst_made)), Pinned=False, placement=placement)
    inst_made.append(inst)
    return inst


L.RevitLinkInstance = NS(Create=make_instance)
L.set_element_workset = lambda doc, el, wid: (True, 'Moved') if wid != 7 else (False, 'Owned by Bo')
link_doc = NS(GetElement=lambda eid: lt_elem)
res = L.link_one(link_doc, {'result': LOADED}, 'A.rvt', relative=True, placement='shared',
                 overlay=False, pin=True, workset_for=lambda label: 3)
check('link_one: linked, attachment + pin + placement, committed',
      res == {'ok': True, 'status': 'Linked', 'workset_note': '', 'type_id': 50,
              'instance_id': 60} and lt_elem.AttachmentType == 'Attachment'
      and inst_made[0].Pinned and inst_made[0].placement == 'S'
      and Tx.log[-1] == ('commit', 'Link: A.rvt'), (res, Tx.log))
res = L.link_one(link_doc, {'result': LOADED}, 'B.rvt', workset_for=lambda label: 7)
check('link_one: workset refused -> linked with a note', res['ok']
      and res['workset_note'] == 'Owned by Bo')
res = L.link_one(link_doc, {'result': FAILED}, 'C.rvt')
check('link_one: load failure rolled back', not res['ok'] and res['status'] == FAILED
      and Tx.log[-1] == ('rollback', 'Link: C.rvt'))
res = L.link_one(link_doc, {'result': LOADED, 'type': -1}, 'D.rvt')
check('link_one: invalid type rolled back', res['status'] == 'Invalid Type'
      and Tx.log[-1] == ('rollback', 'Link: D.rvt'))
res = L.link_one(link_doc, {}, 'E.rvt')
check('link_one: exception -> first line, rolled back', not res['ok'] and res['status']
      and Tx.log[-1] == ('rollback', 'Link: E.rvt'), res)

# the tool, with the Revit side of _links faked
STATE = {'links': {}, 'worksets': {'Arch.rvt': {'Furniture': False, 'Site': True}}}


def fake_link_one(doc, mp, label, relative=False, placement='origin', overlay=True, pin=False,
                  workset_for=None):
    ws = workset_for(label) if workset_for else -1
    if 'Broken' in label:
        return {'ok': False, 'status': 'LinkNotFound', 'workset_note': '', 'type_id': -1,
                'instance_id': -1}
    if 'Ghost' not in label:
        STATE['links'][mp] = label
    return {'ok': True, 'status': 'Linked', 'workset_note': '' if ws != 9 else 'read-only',
            'type_id': 100 + len(STATE['links']), 'instance_id': 200 + len(STATE['links'])}


def fake_keys(doc):
    return (set(os.path.normpath(p).lower() for p in STATE['links']),
            set(n.lower() for n in STATE['links'].values()))


def fake_records(doc):
    out = []
    for i, (name, ws) in enumerate(sorted(STATE['worksets'].items())):
        ld = NS(ws=ws) if name != 'Unloaded.rvt' else None
        out.append(NS(name=name, type_id=300 + i, link_type=NS(name=name), link_doc=ld,
                      path=STATE.get('path_' + name, 'C:/old/' + name)))
    return out


ENSURED = []
L.link_one = fake_link_one
L.existing_link_keys = fake_keys
L.user_model_path = lambda p: p
L.is_workshared = lambda doc: True
L.ensure_workset = lambda doc, name: ENSURED.append(name) or 4
L.collect_links = fake_records
L.eid_int = lambda v: v
L.get_link_worksets = lambda ld: [L.LinkWorkset(i, n, o, True) for i, (n, o) in enumerate(
    sorted(ld.ws.items()))] if ld is not None else []


def fake_apply(link_type, open_ids, close_ids):
    ws = STATE['worksets'][link_type.name]
    names = sorted(ws)
    for i in open_ids:
        ws[names[i]] = True
    for i in close_ids:
        ws[names[i]] = 'stuck' in STATE      # stuck: Revit keeps it open
    return True, 'Reloaded'


L.apply_link_worksets = fake_apply


def fake_reload_from(link_type, path):
    STATE['path_' + link_type.name] = path if 'Wrong' not in path else 'C:/elsewhere.rvt'
    return True, 'Reloaded from new path'


L.reload_link_from = fake_reload_from
saved = dict(sys.modules)
for name in [n for n in sys.modules if n == 'Snippets' or n.startswith('Snippets.')]:
    del sys.modules[name]
import Snippets                                  # noqa: E402
Snippets._links = L
sys.modules['Snippets._links'] = L
db_stub = types.ModuleType('Autodesk.Revit.DB')
db_stub.TransactionGroup = FakeGroup
sys.modules.update({'Autodesk': types.ModuleType('Autodesk'),
                    'Autodesk.Revit': types.ModuleType('Autodesk.Revit'),
                    'Autodesk.Revit.DB': db_stub})
doc = object()
open(os.path.join(folder, 'Broken.rvt'), 'w').close()

r = TS.run('link_models', doc, {'operation': 'link', 'paths': [os.path.join(folder, 'Nope.rvt')]})
check('link: missing file -> error, nothing linked', 'file not found' in r['error']
      and not STATE['links'], r)
r = TS.run('link_models', doc, {'operation': 'link', 'paths': [folder], 'placement': 'roof'})
check('link: bad placement', r['error'].startswith('placement must'), r)
r = TS.run('link_models', doc, {'operation': 'link', 'paths': [folder], 'dry_run': True})
check('link dry run: plan per file, nothing linked',
      [m['action'] for m in r['models']] == ['link', 'link', 'link'] and not STATE['links'], r)
del FakeGroup.log[:]
r = TS.run('link_models', doc, {'operation': 'link', 'paths': [folder], 'workset_prefix': 'LNK_'})
check('link: two linked, broken reported, one TransactionGroup kept',
      r['linked'] == 2 and r['link_type_ids'] == [101, 102] and r['link_instance_ids'] == [201, 202]
      and [m['status'] for m in r['models']] == ['Linked', 'Linked', 'LinkNotFound']
      and FakeGroup.log == [('start', 'T3Lab: Link Models'), ('assimilate', 'T3Lab: Link Models')]
      and 'not_in_model_after' not in r, (r, FakeGroup.log))
check('link: per-link worksets named prefix + file name', ENSURED == ['LNK_A', 'LNK_B', 'LNK_Broken'],
      ENSURED)
v = W.verify('link_models', {}, r, 0, lambda i: i in (101, 102, 201, 202))
check('write_verify: reported link ids exist', v and v['verified'] and v['checks'] >= 4, v)
r = TS.run('link_models', doc, {'operation': 'link', 'paths': [os.path.join(folder, 'A.rvt')]})
check('link: already linked is skipped by default, nothing to undo',
      r['models'][0]['status'].startswith('skipped') and r['success'] and not r['applied']
      and FakeGroup.log[-1] == ('rollback', 'T3Lab: Link Models'), r)
ghost = os.path.join(folder, 'Ghost.rvt')
open(ghost, 'w').close()
r = TS.run('link_models', doc, {'operation': 'link', 'paths': [ghost]})
check('link read-back: reported linked but not in the model', r.get('not_in_model_after') == [ghost],
      r)

r = TS.run('link_models', doc, {'operation': 'set_worksets', 'links': ['Nope']})
check('set_worksets: unknown link lists the links', r['links'] == ['Arch.rvt'], r)
r = TS.run('link_models', doc, {'operation': 'set_worksets', 'links': ['arch'],
                                'open_worksets': ['Site'], 'close_worksets': ['site']})
check('set_worksets: same name in both -> error', 'no name in both' in r['error'], r)
r = TS.run('link_models', doc, {'operation': 'set_worksets', 'links': ['arch'],
                                'open_worksets': ['furniture'], 'close_worksets': ['Site', 'X'],
                                'dry_run': True})
row = r['links'][0]
check('set_worksets dry run: current state, plan, unknown names',
      row['open'] == ['Furniture'] and row['close'] == ['Site'] and row['not_in_link'] == ['X']
      and row['changes'] == 2 and 'Furniture (closed)' in row['worksets']
      and STATE['worksets']['Arch.rvt'] == {'Furniture': False, 'Site': True}, r)
r = TS.run('link_models', doc, {'operation': 'set_worksets', 'links': [300],
                                'open_worksets': ['furniture'], 'close_worksets': ['Site']})
check('set_worksets: applied and read back',
      r['success'] and r['reloaded'] == 1 and 'not_in_model_after' not in r
      and STATE['worksets']['Arch.rvt'] == {'Furniture': True, 'Site': False}, r)
r = TS.run('link_models', doc, {'operation': 'set_worksets', 'links': ['Arch'],
                                'open_worksets': ['Furniture']})
check('set_worksets: nothing to change -> already so', r['links'][0]['status'] == 'already so'
      and not r['applied'], r)
STATE['stuck'] = True
STATE['worksets']['Arch.rvt']['Site'] = True
r = TS.run('link_models', doc, {'operation': 'set_worksets', 'links': ['Arch'],
                                'close_worksets': ['Site']})
check('set_worksets read-back: a workset Revit kept open is reported',
      r.get('not_in_model_after') == ['Arch.rvt: Site not closed'], r)
STATE['worksets']['Unloaded.rvt'] = {}
r = TS.run('link_models', doc, {'operation': 'set_worksets', 'links': ['Unloaded'],
                                'open_worksets': ['A']})
check('set_worksets: unloaded link says what to do', 'reload it first' in r['links'][0]['status']
      and 'error' in r, r)

new = os.path.join(folder, 'B.RVT')
r = TS.run('link_models', doc, {'operation': 'reload_from', 'links': ['Arch'], 'paths': [new]})
check('reload_from: repointed, read back', r['success'] and r['link_type_id'] == 300
      and 'not_in_model_after' not in r, r)
check('reload_from: write_verify on the type id',
      W.verify('link_models', {}, r, 0, lambda i: i == 300)['verified'])
wrong = os.path.join(folder, 'Wrong.rvt')
open(wrong, 'w').close()
r = TS.run('link_models', doc, {'operation': 'reload_from', 'links': ['Arch'], 'paths': [wrong]})
check('reload_from read-back: link points elsewhere', r.get('not_in_model_after'), r)
r = TS.run('link_models', doc, {'operation': 'reload_from', 'links': ['Arch', 'Unloaded'],
                                'paths': [new]})
check('reload_from: exactly one link', r['error'].startswith('reload_from needs'), r)
sys.modules.clear()
sys.modules.update(saved)

# ── cad_imports (Snippets/_cad_imports.py on a fake DB) ─────────────────────
print('[cad_imports]')


class Inst(object):
    def __init__(self, eid, linked, type_id, view_id, pinned=False):
        self.Id, self.IsLinked, self.type_id = eid, linked, type_id
        self.OwnerViewId, self.Pinned, self.Name = view_id, pinned, 'import'

    def GetTypeId(self):
        return self.type_id


class CadDoc(object):
    def __init__(self):
        self.els = {}

    def GetElement(self, eid):
        return self.els.get(eid)

    def Delete(self, eid):
        el = self.els[eid]
        if getattr(el, 'Pinned', False) or getattr(el, 'stuck', False):
            raise RuntimeError('ElementId cannot be deleted.')
        del self.els[eid]


class CadTx(object):
    def __init__(self, *a):
        self.ended = False

    def Start(self):
        pass

    def Commit(self):
        self.ended = True
        return 'Committed'

    def RollBack(self):
        self.ended = True

    def HasEnded(self):
        return self.ended


class Sub(object):
    def __init__(self, doc):
        pass

    def Start(self):
        pass

    def Commit(self):
        pass

    def RollBack(self):
        pass


class Collector(object):
    def __init__(self, doc):
        self.doc = doc

    def OfClass(self, cls):
        items = [e for e in self.doc.els.values() if isinstance(e, Inst)]
        return NS(ToElements=lambda: items)


logging.getLogger('T3Lab.cad_imports').addHandler(logging.NullHandler())   # expected warnings
cad_db = types.ModuleType('Autodesk.Revit.DB')
cad_db.FilteredElementCollector, cad_db.ImportInstance = Collector, Inst
cad_db.Transaction, cad_db.SubTransaction = CadTx, Sub
cad_db.TransactionStatus = NS(Committed='Committed')
cad_db.ElementId = NS(InvalidElementId=-1)
cad_db.View = NS
saved = dict(sys.modules)
fresh_modules({'Autodesk': types.ModuleType('Autodesk'),
               'Autodesk.Revit': types.ModuleType('Autodesk.Revit'), 'Autodesk.Revit.DB': cad_db})
cdoc = CadDoc()
cdoc.els.update({
    1: Inst(1, False, 10, 5, pinned=True), 2: Inst(2, True, 20, -1), 3: Inst(3, True, 20, 5),
    10: NS(Id=10, Name='Survey.dwg'),
    20: NS(Id=20, Name='Site.dwg', GetExternalFileReference=lambda: None),
    5: NS(Name='Level 1'), 4: NS(Name='a wall')})
r = TS.run('cad_imports', cdoc, {'operation': 'list'})
rows = dict((x['id'], x) for x in r['items'])
check('list: import / link, owner view, pin', r['count'] == 3 and r['imports'] == 1
      and r['links'] == 2 and rows[1]['view'] == 'Level 1' and rows[1]['view_id'] == 5
      and rows[2]['view_id'] is None and rows[2]['view'] == 'All Views (3D)' and rows[1]['pinned'],
      r)
r = TS.run('cad_imports', cdoc, {'operation': 'delete', 'element_ids': [4]})
check('delete: not a CAD item -> error, nothing deleted', 'Not a CAD import' in r['error']
      and len(cdoc.els) == 7, r)
r = TS.run('cad_imports', cdoc, {'operation': 'delete', 'element_ids': ['x']})
check('delete: unresolved handle -> error', 'not an element id' in r['error'], r)
r = TS.run('cad_imports', cdoc, {'operation': 'delete', 'element_ids': [1, 2, 2],
                                 'dry_run': True})
check('delete dry run: items listed once, nothing deleted',
      [x['id'] for x in r['items']] == [1, 2] and r['dry_run'] and len(cdoc.els) == 7, r)
r = TS.run('cad_imports', cdoc, {'operation': 'delete', 'element_ids': [1, 2]})
check('delete: pinned import unpinned and deleted, link type kept while used',
      r['deleted'] == 2 and 1 not in cdoc.els and 2 not in cdoc.els and 20 in cdoc.els
      and 'not_in_model_after' not in r, r)
cdoc.els[3].stuck = True
r = TS.run('cad_imports', cdoc, {'operation': 'delete', 'element_ids': [3]})
check('delete: a refused item is reported failed, not done', 'error' in r
      and r['failed'][0]['name'] == 'Site.dwg' and 'not_in_model_after' not in r, r)
real = sys.modules['Snippets._cad_imports'].delete_cad_items
sys.modules['Snippets._cad_imports'].delete_cad_items = lambda d, items: (len(items), [])
r = TS.run('cad_imports', cdoc, {'operation': 'delete', 'element_ids': [3]})
sys.modules['Snippets._cad_imports'].delete_cad_items = real
check('delete read-back: reported deleted but still there',
      r.get('not_in_model_after') == ['Site.dwg [3] still in the model'], r)
check('write_verify fails on it', W.verify('cad_imports', {'operation': 'delete'}, r, 0,
                                           lambda i: True)['verified'] is False)
sys.modules.clear()
sys.modules.update(saved)

shutil.rmtree(TMP, ignore_errors=True)
print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
