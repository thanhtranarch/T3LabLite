# -*- coding: utf-8 -*-
"""core/tools/families.py + core/tools/standards.py: transfer_families, manage_groups,
manage_worksets, smart_purge as MCP / Assistant tools.

Run:  python lite_guard/test_tools_standards_domain.py      (no Revit: a fake Revit API)
"""

import json
import os
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
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_standards_test_')

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ── fake Revit API (enough to import the shipped modules) ───────────────────

class Status(object):
    Started, Committed, RolledBack, Pending = 'Started', 'Committed', 'RolledBack', 'Pending'


class Txn(object):
    log = []

    def __init__(self, doc=None, name=''):
        self.name, self.status = name, None

    def Start(self):
        self.status = Status.Started
        Txn.log.append(('start', self.name))

    def Commit(self):
        self.status = Status.Committed
        Txn.log.append(('commit', self.name))
        return self.status

    def Assimilate(self):
        self.status = Status.Committed
        Txn.log.append(('assimilate', self.name))
        return self.status

    def RollBack(self):
        self.status = Status.RolledBack
        Txn.log.append(('rollback', self.name))

    def GetStatus(self):
        return self.status

    def HasStarted(self):
        return self.status is not None

    def HasEnded(self):
        return self.status in (Status.Committed, Status.RolledBack)

    def GetFailureHandlingOptions(self):
        return types.SimpleNamespace(SetForcedModalHandling=lambda v: None)

    def SetFailureHandlingOptions(self, o):
        pass

    def Dispose(self):
        pass


class StubModule(types.ModuleType):
    def __getattr__(self, name):
        if name.startswith('__'):
            raise AttributeError(name)
        value = type(str(name), (object,), {})
        setattr(self, name, value)
        return value


db = StubModule('Autodesk.Revit.DB')
db.Transaction = db.TransactionGroup = db.SubTransaction = Txn
db.TransactionStatus = Status
for name, mod in (('Autodesk', StubModule('Autodesk')),
                  ('Autodesk.Revit', StubModule('Autodesk.Revit')),
                  ('Autodesk.Revit.DB', db),
                  ('Autodesk.Revit.DB.Architecture', StubModule('Autodesk.Revit.DB.Architecture')),
                  ('System', StubModule('System')),
                  ('System.Collections', StubModule('System.Collections')),
                  ('System.Collections.Generic', StubModule('System.Collections.Generic'))):
    sys.modules[name] = mod
pyrevit = types.ModuleType('pyrevit')
pyrevit.revit = types.ModuleType('pyrevit.revit')
pyrevit.script = types.ModuleType('pyrevit.script')
pyrevit.script.get_logger = lambda: types.SimpleNamespace(
    warning=lambda *a: None, error=lambda *a: None, debug=lambda *a: None)
sys.modules.update({'pyrevit': pyrevit, 'pyrevit.revit': pyrevit.revit,
                    'pyrevit.script': pyrevit.script})

from core import tool_specs as TS                           # noqa: E402
from core import write_verify as W                          # noqa: E402
from core.tools import families as HF, standards as HS      # noqa: E402
from Snippets import family_transfer as FT                  # noqa: E402
try:
    from Snippets import doc_transfer as DT                 # noqa: E402
except ImportError:      # T3Lab Lite: sources and targets live in family_transfer
    DT = FT
from Snippets import _group_ops as GOPS                     # noqa: E402
from Services import workset_service as WS                  # noqa: E402
from Services.ModelAuditor.smart_purge import purge_scanner as PS   # noqa: E402
import Snippets._compat as C                                # noqa: E402

C.make_eid = lambda v: v


class Doc(object):
    def __init__(self, title, path='', workshared=True):
        self.Title, self.PathName, self.IsWorkshared = title, path, workshared
        self.Application = types.SimpleNamespace(Username='me')
        self.elements = {}

    def Equals(self, other):
        return self is other

    def GetElement(self, i):
        return self.elements.get(i)


# ── transfer_families ───────────────────────────────────────────────────────
print('[transfer_families]')
Fam = FT.FamilyInfo
door = Fam('Door-Single', 'Doors', 1, {'800x2100': 11, '900x2100': 12})
win = Fam('Window-Fixed', 'Windows', 2, {'600x900': 21})
sel, miss_f, miss_t = HF.pick_families([door, win], ['door-single', 'Nope'], ['900X2100', 'Z'])
check('pick_families: case-insensitive, types filtered, misses reported',
      [(f.name, t) for f, t in sel] == [('Door-Single', ['900x2100'])]
      and miss_f == ['Nope'] and miss_t == ['Z'], (sel, miss_f, miss_t))
sel, _f, _t = HF.pick_families([door, win], ['Window-Fixed', 'Door-Single'], [])
check('pick_families: no type_names = every type',
      [t for _f2, t in sel] == [['600x900'], ['800x2100', '900x2100']])

src, tower, podium = Doc('Library.rvt', 'C:/l.rvt'), Doc('Tower.rvt', 'C:/t.rvt'), \
    Doc('Podium.rvt', 'C:/p.rvt')
src.Application = object()
INDEX = {'Tower.rvt': {'Door-Single': {'800x2100'}}, 'Podium.rvt': {}}
LANDED = {'Tower.rvt': {}, 'Podium.rvt': {}}
DT.open_projects = lambda app: [src, tower, podium]
DT._loaded_link_docs = lambda host: []
FT.collect_families = lambda d: [door, win]


def _index(d):
    out = dict((k, set(v)) for k, v in INDEX[d.Title].items())
    for fam, types_ in LANDED[d.Title].items():
        out.setdefault(fam, set()).update(types_)
    return out


FT.target_index = _index


def fake_execute(src_doc, plans, overwrite_values=False, step=None):
    fake_execute.values = overwrite_values
    out = []
    for d, actions in plans:
        res = []
        for a in actions:
            if a.kind == 'skip':
                res.append(FT.Result(a.family.name, 'skipped', 0, a.reason))
                continue
            res.append(FT.Result(a.family.name, 'ok', len(a.types), a.kind))
            if d.Title != 'Podium.rvt' or a.family.name != 'Window-Fixed':
                LANDED[d.Title].setdefault(a.family.name, set()).update(a.types)
        out.append((d, res, None))
    return out


FT.execute_plans = fake_execute
r = TS.run('transfer_families', src, {'target_documents': ['Tower'],
                                      'family_names': ['Door-Single'], 'mode': 'add',
                                      'dry_run': True})
check('dry run: add mode copies only the missing type, nothing landed',
      r.get('dry_run') and r['plans'][0]['copy'] == ['Door-Single (900x2100)']
      and not LANDED['Tower.rvt'], r)
r = TS.run('transfer_families', src, {'target_documents': ['Tower'],
                                      'family_names': ['Door-Single'], 'dry_run': True})
check('dry run: skip mode leaves an existing family alone',
      r['plans'][0]['skip'] == ['Door-Single: already in the target project'], r)
r = TS.run('transfer_families', src, {'target_documents': ['Tower'], 'family_names': ['X']})
check('unknown family -> error', 'family X' in r['error'], r)
r = TS.run('transfer_families', src, {'target_documents': ['Tower'],
                                      'family_names': ['Door-Single'], 'type_names': ['1x1']})
check('unknown type -> error', 'type 1x1' in r['error'], r)
r = TS.run('transfer_families', src, {'target_documents': ['Library'], 'family_names': ['W']})
check('source is not a target', 'none other than the source' in r['error'], r)
r = TS.run('transfer_families', src, {'target_documents': ['Tower'], 'family_names': []})
check('family_names required', 'family_names' in r['error'], r)
r = TS.run('transfer_families', src, {'target_documents': ['Tower'], 'family_names': ['W'],
                                      'mode': 'merge'})
check('bad mode', r['error'].startswith('mode must'), r)
r = TS.run('transfer_families', src, {'target_documents': ['Tower', 'Podium'],
                                      'family_names': ['Door-Single', 'Window-Fixed'],
                                      'mode': 'overwrite', 'overwrite_values': True})
check('run: overwrite + copy counted per target', r['applied'] and r['transferred'] == 4
      and fake_execute.values is True and r['targets'][0]['summary']['ok'] == 2, r)
check('read-back: the family missing in Podium is reported',
      r.get('not_in_model_after') == ['Podium.rvt: Window-Fixed'], r.get('not_in_model_after'))
check('write_verify fails on it', W.verify('transfer_families', {}, r, 0,
                                           lambda i: True)['verified'] is False)
LANDED['Podium.rvt']['Window-Fixed'] = {'600x900'}
r = TS.run('transfer_families', src, {'target_documents': ['Podium'],
                                      'family_names': ['Window-Fixed'], 'mode': 'overwrite'})
check('all landed -> verified', 'not_in_model_after' not in r and W.verify(
    'transfer_families', {}, r, 0, lambda i: True)['verified'], r)
LANDED['Tower.rvt']['Door-Single'] = {'800x2100'}
lost = HF.lost_after('T', FT.plan_transfer([(door, ['900x2100'])], {}, 'skip'),
                     [FT.Result('Door-Single', 'ok', 1, '')], {'Door-Single': {'800x2100'}})
check('lost_after: a missing type is reported', lost == ['T: Door-Single: 900x2100'], lost)

# ── manage_groups ───────────────────────────────────────────────────────────
print('[manage_groups]')


class Rec(object):
    def __init__(self, type_id, name, kind='Model', instances=(), ws=()):
        self.type_id, self.name, self.kind = type_id, name, kind
        self.instances = list(instances)
        self.instance_count = len(self.instances)
        self.member_count = 3 if self.instances else 0
        self.workset_names = list(ws)

    def audit_issues(self):
        return [] if self.instance_count else ['Unused']


MODEL = {}


def collect(doc):
    return [Rec(*MODEL[k]) for k in sorted(MODEL)]


def reset_groups():
    MODEL.clear()
    MODEL.update({1: (1, 'Bath Pod', 'Model', ['i1', 'i2'], ['ARC_Toilets']),
                  2: (2, 'Old Kit', 'Model', [], []),
                  3: (3, 'Tag Set', 'Detail', ['i3'], ['ARC_Misc']),
                  4: (4, 'Spare', 'Detail', [], [])})


reset_groups()
GOPS.collect_group_types = collect
GOPS.user_worksets = lambda doc: [(10, 'ARC_Misc'), (11, 'ARC_Toilets')]


def fake_rename(doc, pairs, progress=None):
    for r, wanted in pairs:
        if wanted != 'Lost Name':
            MODEL[r.type_id] = (r.type_id, wanted) + MODEL[r.type_id][2:]
    return [(r, True, 'Renamed') for r, _w in pairs]


def fake_workset(doc, records, value, include_members=False, progress=None):
    fake_workset.members = include_members
    name = dict(GOPS.user_worksets(doc))[value]
    for r in records:
        MODEL[r.type_id] = MODEL[r.type_id][:4] + ([name],)
    return [(r, len(r.instances), 0, 0, 'Moved') for r in records]


def fake_purge(doc, records, progress=None):
    for r in records:
        MODEL.pop(r.type_id, None)
    return [(r, True, 'Purged') for r in records]


def fake_ungroup(doc, records, progress=None):
    for r in records:
        MODEL[r.type_id] = MODEL[r.type_id][:3] + ([],) + MODEL[r.type_id][4:]
    return [(r, len(r.instances), 0, 'Ungrouped') for r in records]


GOPS.rename_group_types, GOPS.apply_workset = fake_rename, fake_workset
GOPS.purge_group_types, GOPS.ungroup_instances = fake_purge, fake_ungroup
gdoc = Doc('Tower.rvt')

r = TS.run('manage_groups', gdoc, {'operation': 'list', 'kind': 'detail'})
check('list: kind filter, rows with issues', r['count'] == 2 and r['unused'] == 1
      and r['groups'][1] == {'name': 'Spare', 'kind': 'Detail', 'instances': 0,
                             'members': 0, 'issues': ['Unused']}, r)
r = TS.run('manage_groups', gdoc, {'operation': 'list', 'group_names': ['nope']})
check('unknown group -> error + hint', 'nope' in r['error'] and 'list' in r['hint'], r)
r = TS.run('manage_groups', gdoc, {'operation': 'rename', 'kind': 'model', 'prefix': 'GRP_',
                                   'dry_run': True})
check('rename dry run: prefix rule, nothing changed',
      r['renames'] == [{'from': 'Bath Pod', 'to': 'GRP_Bath Pod'},
                       {'from': 'Old Kit', 'to': 'GRP_Old Kit'}]
      and MODEL[1][1] == 'Bath Pod', r)
r = TS.run('manage_groups', gdoc, {'operation': 'rename',
                                   'new_names': {'bath pod': 'Spare', 'Old Kit': 'Kit|2',
                                                 'Tag Set': 'Tags'}})
check('rename: clash with a kept name and illegal chars skipped, rest renamed',
      r['renamed'] == 1 and MODEL[3][1] == 'Tags' and MODEL[1][1] == 'Bath Pod'
      and sorted(s['group'] for s in r['skipped']) == ['Bath Pod', 'Old Kit'], r)
r = TS.run('manage_groups', gdoc, {'operation': 'rename', 'new_names': {'Tags': 'Lost Name'}})
check('rename read-back: name not in the model is reported',
      r['not_in_model_after'] == ['Lost Name'], r)
check('check_renames: two records asking for one name',
      len(HS.check_renames([(Rec(8, 'a'), 'x'), (Rec(9, 'b'), 'X')], [])[1]) == 1)
r = TS.run('manage_groups', gdoc, {'operation': 'rename', 'group_names': ['Spare']})
check('rename without a rule -> error', 'new_names' in r['error'], r)
r = TS.run('manage_groups', gdoc, {'operation': 'set_workset', 'group_names': ['Bath Pod'],
                                   'workset': 'arc_misc', 'include_members': True})
check('set_workset: moved, members flag passed, read back ok',
      r['moved'] == 2 and fake_workset.members and MODEL[1][4] == ['ARC_Misc']
      and 'not_in_model_after' not in r, r)
r = TS.run('manage_groups', gdoc, {'operation': 'set_workset', 'group_names': ['Bath Pod'],
                                   'workset': 'Nope'})
check('set_workset: unknown workset -> names', r['worksets'] == ['ARC_Misc', 'ARC_Toilets'], r)
r = TS.run('manage_groups', gdoc, {'operation': 'set_workset', 'workset': 'ARC_Misc'})
check('set_workset: needs a scope', 'group_names' in r['error'], r)
r = TS.run('manage_groups', Doc('Solo', workshared=False),
           {'operation': 'set_workset', 'group_names': ['Bath Pod'], 'workset': 'ARC_Misc'})
check('set_workset: not workshared', 'not workshared' in r['error'], r)
r = TS.run('manage_groups', gdoc, {'operation': 'purge_unused', 'dry_run': True})
check('purge dry run lists the unused types only', r['unused'] == ['Old Kit', 'Spare']
      and 2 in MODEL, r)
r = TS.run('manage_groups', gdoc, {'operation': 'purge_unused', 'kind': 'model'})
check('purge: unused model types deleted and gone', r['deleted'] == 1 and 2 not in MODEL
      and 4 in MODEL and 'not_in_model_after' not in r, r)
r = TS.run('manage_groups', gdoc, {'operation': 'ungroup'})
check('ungroup needs group_names', 'group_names' in r['error'], r)
r = TS.run('manage_groups', gdoc, {'operation': 'ungroup', 'group_names': ['Bath Pod']})
check('ungroup: instances counted, none left', r['ungrouped'] == 2 and MODEL[1][3] == [], r)
GOPS.ungroup_instances = lambda doc, recs, progress=None: [(x, 1, 0, 'Ungrouped') for x in recs]
r = TS.run('manage_groups', gdoc, {'operation': 'ungroup', 'group_names': ['Tags']})
check('ungroup read-back: an instance still grouped is reported',
      'still grouped' in r['not_in_model_after'][0], r)


def boom(*a, **k):
    raise RuntimeError('Transaction was not committed: RolledBack')


GOPS.purge_group_types = boom
r = TS.run('manage_groups', gdoc, {'operation': 'purge_unused'})
check('transaction failure -> error, nothing changed', 'Nothing was changed' in r['error'], r)
r = TS.run('manage_groups', gdoc, {'operation': 'explode'})
check('bad operation', r['error'].startswith('operation must'), r)

# ── manage_worksets ─────────────────────────────────────────────────────────
print('[manage_worksets]')


class Ws(object):
    def __init__(self, i, name, owner=''):
        self.Id, self.Name, self.Owner = i, name, owner


WORKSETS = []
VIEWS = set()


def reset_worksets():
    WORKSETS[:] = [Ws(1, 'ARC_Floor'), Ws(2, 'ARC_Old'), Ws(3, 'ARC_Tmp'), Ws(4, 'ARC_Them', 'bob')]
    VIEWS.clear()
    VIEWS.add('ARC_Floor')


reset_worksets()
WS.get_user_worksets = lambda doc: list(WORKSETS)
WS.is_workset_name_unique = lambda doc, n: True


def fake_ws_rename(doc, wid, new):
    for w in WORKSETS:
        if w.Id == wid:
            w.Name = new
    return True, ''


def fake_ws_delete(doc, name, keep):
    if name == 'ARC_Tmp':
        return False, 'Revit said no'
    WORKSETS[:] = [w for w in WORKSETS if w.Name != name]
    return True, None


def fake_views(doc):
    made = [w.Name for w in WORKSETS if w.Name not in VIEWS]
    VIEWS.update(made)
    return made, [w.Name for w in WORKSETS if w.Name not in made], None


WS.rename_workset, WS.delete_workset = fake_ws_rename, fake_ws_delete
WS.create_workset_views = fake_views
WS.existing_3d_view_names = lambda doc: set(VIEWS)
wdoc = Doc('Tower.rvt')

r = TS.run('manage_worksets', wdoc, {'operation': 'rename', 'worksets': ['arc_old'],
                                     'new_name': 'ARC_New', 'dry_run': True})
check('rename dry run', r['renames'] == [{'from': 'ARC_Old', 'to': 'ARC_New'}]
      and WORKSETS[1].Name == 'ARC_Old', r)
r = TS.run('manage_worksets', wdoc, {'operation': 'rename', 'worksets': ['ARC_Old'],
                                     'new_name': 'ARC_Floor'})
check('rename onto an existing name -> validator message', 'ARC_Floor' in r['error'], r)
r = TS.run('manage_worksets', wdoc, {'operation': 'rename', 'worksets': ['ARC_Them'],
                                     'new_name': 'Mine'})
check('rename: owned by another user -> error', 'owned by bob' in r['error'], r)
r = TS.run('manage_worksets', wdoc, {'operation': 'rename', 'worksets': ['ARC_Old'],
                                     'new_name': ' ARC_New '})
check('rename: trimmed name written and read back', r['applied']
      and WORKSETS[1].Name == 'ARC_New' and 'not_in_model_after' not in r, r)
r = TS.run('manage_worksets', wdoc, {'operation': 'delete', 'worksets': ['ARC_New']})
check('delete needs reassign_to', 'reassign_to' in r['error'], r)
r = TS.run('manage_worksets', wdoc, {'operation': 'delete', 'worksets': ['ARC_New'],
                                     'reassign_to': 'ARC_New'})
check('delete: reassign_to cannot be deleted', 'also being deleted' in r['error'], r)
r = TS.run('manage_worksets', wdoc, {'operation': 'delete', 'worksets': ['Ghost'],
                                     'reassign_to': 'ARC_Floor'})
check('delete: unknown workset -> names', 'Ghost' in r['error'] and r['worksets'], r)
Txn.log[:] = []
r = TS.run('manage_worksets', wdoc, {'operation': 'delete', 'worksets': ['ARC_New', 'ARC_Tmp'],
                                     'reassign_to': 'ARC_Floor'})
check('delete: one TransactionGroup assimilated, failure reported',
      r['deleted'] == ['ARC_New'] and r['failed'][0]['workset'] == 'ARC_Tmp'
      and Txn.log == [('start', 'T3Lab: Delete Worksets'),
                      ('assimilate', 'T3Lab: Delete Worksets')], (r, Txn.log))
Txn.log[:] = []
r = TS.run('manage_worksets', wdoc, {'operation': 'delete', 'worksets': ['ARC_Tmp'],
                                     'reassign_to': 'ARC_Floor'})
check('delete: nothing deleted -> group rolled back', not r['applied']
      and Txn.log[-1] == ('rollback', 'T3Lab: Delete Worksets'), Txn.log)
r = TS.run('manage_worksets', wdoc, {'operation': 'create_views', 'dry_run': True})
check('create_views dry run lists worksets without a view',
      r['create'] == ['ARC_Tmp', 'ARC_Them'] and r['skipped'] == 1, r)
r = TS.run('manage_worksets', wdoc, {'operation': 'create_views'})
check('create_views: created and read back', r['created'] == ['ARC_Tmp', 'ARC_Them']
      and 'not_in_model_after' not in r, r)
solo = Doc('Solo.rvt', workshared=False)
r = TS.run('manage_worksets', solo, {'operation': 'delete', 'worksets': ['x'],
                                     'reassign_to': 'y'})
check('not workshared -> enable first', 'enable_worksharing' in r['error'], r)
r = TS.run('manage_worksets', solo, {'operation': 'enable_worksharing', 'dry_run': True})
check('enable dry run names the worksets and warns', r['worksets'] == list(
    WS.WORKSHARING_WORKSETS) and 'Cannot be undone' in r['note'], r)


def fake_enable(doc):
    doc.IsWorkshared = True
    return True


WS.enable_worksharing = fake_enable
r = TS.run('manage_worksets', solo, {'operation': 'enable_worksharing'})
check('enable: applied, read back', r['applied'] and solo.IsWorkshared
      and 'not_in_model_after' not in r, r)
r = TS.run('manage_worksets', solo, {'operation': 'enable_worksharing'})
check('enable twice -> error', 'already workshared' in r['error'], r)

# ── smart_purge ─────────────────────────────────────────────────────────────
print('[smart_purge]')
pdoc = Doc('Tower.rvt')
SCAN = {'MaterialScanner': [{'id': 101, 'name': 'Old Brick', 'can_delete': True},
                            {'id': 102, 'name': 'Glass', 'can_delete': True},
                            {'id': 50, 'name': 'Default', 'can_delete': False}],
        'LinePatternScanner': [{'id': 201, 'name': 'Dash 2', 'can_delete': True}]}


class Scanner(object):
    def __init__(self, items):
        self.items = items

    def scan(self):
        print('scanning ... (must not reach the output)')
        return [dict(i, element=Elem(i['id'])) for i in self.items]


def fake_create_scanner(name, doc):
    if name == 'FillPatternScanner':
        raise RuntimeError('fill patterns broke')
    return Scanner(SCAN.get(name, []))


PS.create_scanner = fake_create_scanner


class Elem(object):
    def __init__(self, i):
        self.Id = i


for i in (101, 102, 50, 201):
    pdoc.elements[i] = Elem(i)


def fake_delete(eid):
    if eid == 201:
        return None
    if eid != 102:                      # 102 "deleted" but stays in the model
        pdoc.elements.pop(eid, None)
    return types.SimpleNamespace(Count=1)


pdoc.Delete = fake_delete
r = TS.run('smart_purge', pdoc, {'categories': ['materials', 'Line Patterns', 'fill_patterns']})
check('default is a dry run: counts, deletable, items, nothing deleted',
      r['dry_run'] and r['categories'][0] == {
          'category': 'materials', 'found': 3, 'deletable': 2,
          'items': ['Old Brick [101]', 'Glass [102]']} and 101 in pdoc.elements, r)
check('a scanner that fails is reported, not fatal',
      r['scan_errors'] == ['fill_patterns: fill patterns broke'], r)
check('other categories named for the next call', 'unused_families' in r['other_categories'])
r = TS.run('smart_purge', pdoc, {'categories': ['vibes']})
check('unknown category -> ids', 'vibes' in r['error'] and 'materials' in r['categories'], r)
r = TS.run('smart_purge', pdoc, {'exclude_ids': ['@set']})
check('exclude_ids must be numbers', 'exclude_ids' in r['error'], r)
r = TS.run('smart_purge', pdoc, {})
check('no categories = the default-checked set', 'wall_types' in r['other_categories']
      and 'materials' in [c['category'] for c in r['categories']], r)
Txn.log[:] = []
r = TS.run('smart_purge', pdoc, {'categories': ['materials', 'line_patterns'],
                                 'exclude_ids': [102], 'dry_run': False})
check('purge: one TransactionGroup, excluded and protected kept',
      Txn.log[0] == ('start', 'T3Lab: Smart Purge') and Txn.log[-1][0] == 'assimilate'
      and 101 not in pdoc.elements and 102 in pdoc.elements and 50 in pdoc.elements, Txn.log)
check('purge: failure reported with reason', r['deleted'] == 1
      and r['failed'] == [{'item': 'Dash 2 [201]', 'reason': 'Delete returned no IDs'}], r)
r = TS.run('smart_purge', pdoc, {'categories': ['materials'], 'dry_run': False})
check('read-back: deleted but still in the model is reported',
      r['not_in_model_after'] == ['deleted Glass [102] (still in the model)'], r)
check('write_verify fails on it', W.verify('smart_purge', {'dry_run': False}, r, 0,
                                           lambda i: True)['verified'] is False)
SCAN.clear()
r = TS.run('smart_purge', pdoc, {'categories': ['materials'], 'dry_run': False})
check('nothing to purge -> unchanged message', not r['applied'] and 'unchanged' in r['message'], r)

# ── declarations / server ───────────────────────────────────────────────────
print('[specs]')
import core.server as S   # noqa: E402
srv = object.__new__(S.T3LabAIServer)
for tool in ('transfer_families', 'manage_groups', 'manage_worksets', 'smart_purge'):
    spec = TS.get(tool)
    size = len(json.dumps(spec.registry_entry())) // 4
    check('{}: complete, main thread, ~{} tokens (< 350)'.format(tool, size),
          not spec.problems() and tool in S.T3LabAIServer._WRITE_TOOLS and size < 350
          and 'self' in spec.verify)
check('transfer_families: overwrite is destructive, its dry run is not',
      srv.is_destructive('transfer_families', {'mode': 'overwrite'})
      and not srv.is_destructive('transfer_families', {'mode': 'overwrite', 'dry_run': True})
      and not srv.is_destructive('transfer_families', {'mode': 'add'}))
check('manage_groups: purge_unused / ungroup destructive, rename not',
      srv.is_destructive('manage_groups', {'operation': 'ungroup'})
      and srv.is_destructive('manage_groups', {'operation': 'purge_unused'})
      and not srv.is_destructive('manage_groups', {'operation': 'rename'}))
check('manage_worksets: delete / enable destructive, create_views not',
      srv.is_destructive('manage_worksets', {'operation': 'delete'})
      and srv.is_destructive('manage_worksets', {'operation': 'enable_worksharing'})
      and not srv.is_destructive('manage_worksets', {'operation': 'create_views'}))
check('smart_purge: destructive only with dry_run=false',
      srv.is_destructive('smart_purge', {'dry_run': False})
      and not srv.is_destructive('smart_purge', {})
      and not srv.is_destructive('smart_purge', {'dry_run': True}))
check('no document', 'No active document' in TS.run('smart_purge', None, {})['error'])

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
