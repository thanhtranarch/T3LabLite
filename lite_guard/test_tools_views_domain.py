# -*- coding: utf-8 -*-
"""core/tools/view_sync.py + core/tools/room_views.py: Crop Sync, Datum Sync and
SheetGen as MCP tools (sync_views, generate_room_views).

Run:  python lite_guard/test_tools_views_domain.py      (no Revit: a fake API stands in)
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
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_views_domain_test_')

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


# ── fake Autodesk.Revit.DB ──────────────────────────────────────────────────
STATUS = NS(Started='Started', Committed='Committed', RolledBack='RolledBack',
            Pending='Pending', Uninitialized='Uninitialized')
LOG = []


class Tx(object):
    """Transaction / TransactionGroup / SubTransaction."""

    def __init__(self, doc=None, name='sub'):
        self.name, self.status = name, STATUS.Uninitialized

    def Start(self):
        self.status = STATUS.Started
        LOG.append(('start', self.name))
        return self.status

    def GetStatus(self):
        return self.status

    def HasStarted(self):
        return self.status != STATUS.Uninitialized

    def HasEnded(self):
        return self.status in (STATUS.Committed, STATUS.RolledBack)

    def Commit(self):
        self.status = STATUS.Committed
        LOG.append(('commit', self.name))
        return self.status

    Assimilate = Commit

    def RollBack(self):
        self.status = STATUS.RolledBack
        LOG.append(('rollback', self.name))

    def Dispose(self):
        pass

    def GetFailureHandlingOptions(self):
        return NS(SetForcedModalHandling=lambda v: None, SetClearAfterRollback=lambda v: None)

    def SetFailureHandlingOptions(self, options):
        pass


VT = NS(FloorPlan='FloorPlan', CeilingPlan='CeilingPlan', Elevation='Elevation',
        Section='Section', Detail='Detail', AreaPlan='AreaPlan', EngineeringPlan='EngPlan',
        ThreeD='ThreeD', DrawingSheet='DrawingSheet')
BIP = NS(ROOM_NUMBER='num', ROOM_NAME='name', SHEET_NUMBER='sheetnum', SYMBOL_NAME_PARAM='sym',
         VIEWER_BOUND_FAR_CLIPPING='far', VIEWER_BOUND_OFFSET_FAR='faroff',
         VIEWER_ANNOTATION_CROP_ACTIVE='anno')


class XYZ(object):
    def __init__(self, x=0.0, y=0.0, z=0.0):
        self.X, self.Y, self.Z = x, y, z


class Param(object):
    def __init__(self, value=''):
        self.value, self.IsReadOnly = value, False

    def AsString(self):
        return self.value

    def Set(self, value):
        self.value = value
        return True


class El(object):
    kinds = ()
    category = None

    def __init__(self, eid, name='', **kw):
        self.Id, self.Name, self.params = eid, name, {}
        self.__dict__.update(kw)

    def get_Parameter(self, key):
        return self.params.get(key)

    def LookupParameter(self, name):
        return self.params.get(name)


class View(El):
    kinds = ('View',)
    IsTemplate = False
    CropBoxActive = False
    CropBoxVisible = False
    ViewTemplateId = None
    Scale = 50


class CropView(View):
    def GetCropRegionShapeManager(self):
        return object()


class ViewPlan(CropView):
    kinds = ('View', 'ViewPlan')


class ViewSheet(View):
    kinds = ('View', 'ViewSheet')
    ViewType = VT.DrawingSheet

    @property
    def SheetNumber(self):
        return self.params[BIP.SHEET_NUMBER].value

    @staticmethod
    def Create(doc, tb_id):
        sheet = ViewSheet(doc.new_id(), '')
        sheet.params[BIP.SHEET_NUMBER] = Param('')
        return doc.add(sheet)


class ViewFamilyType(El):
    kinds = ('ViewFamilyType',)


class Collector(object):
    def __init__(self, doc, view_id=None):
        self.items = [] if view_id is not None else list(doc.model.values())

    def OfClass(self, cls):
        self.items = [e for e in self.items if cls.__name__ in e.kinds]
        return self

    def OfCategory(self, cat):
        self.items = [e for e in self.items if e.category == cat]
        return self

    def WhereElementIsElementType(self):
        return self

    WhereElementIsNotElementType = WhereElementIsElementType

    def ToElements(self):
        return list(self.items)

    def __iter__(self):
        return iter(self.items)


class Marker(object):
    MaximumViewCount = 4
    DIRS = [(0, 1), (1, 0), (0, -1), (-1, 0)]

    def __init__(self, doc):
        self.doc, self.Id = doc, doc.new_id()

    def IsAvailableIndex(self, idx):
        return True

    def CreateElevation(self, doc, plan_id, idx):
        dx, dy = self.DIRS[idx]
        ev = CropView(doc.new_id(), 'Elevation 1', ViewType=VT.Elevation,
                      ViewDirection=XYZ(dx, dy, 0))
        return doc.add(ev)


class Viewport(object):
    made = []

    @staticmethod
    def Create(doc, sheet_id, view_id, point):
        vp = NS(Id=doc.new_id(), sheet=sheet_id, view=view_id,
                GetBoxOutline=lambda: NS(MinimumPoint=XYZ(0, 0), MaximumPoint=XYZ(0.2, 0.2)),
                SetBoxCenter=lambda p: None)
        Viewport.made.append(vp)
        return vp


def _plan_create(doc, type_id, level_id):
    vt = doc.model[type_id]
    return doc.add(ViewPlan(doc.new_id(), '', ViewType=vt.view_type, GenLevel=None))


DB = types.ModuleType('Autodesk.Revit.DB')
DB.__dict__.update(
    Transaction=Tx, TransactionGroup=Tx, SubTransaction=lambda doc: Tx(doc, 'sub'),
    TransactionStatus=STATUS, ViewType=VT, BuiltInParameter=BIP,
    BuiltInCategory=NS(OST_Rooms='rooms', OST_TitleBlocks='titleblocks'),
    ViewFamily=NS(Elevation='ElevationFamily'), FilteredElementCollector=Collector,
    View=View, ViewPlan=NS(Create=_plan_create, __name__='ViewPlan'), ViewSheet=ViewSheet,
    ViewFamilyType=ViewFamilyType, Viewport=Viewport, XYZ=XYZ,
    BoundingBoxXYZ=lambda: NS(Min=None, Max=None),
    ElevationMarker=NS(CreateElevationMarker=lambda doc, t, p, s: Marker(doc)),
    ElementTransformUtils=NS(RotateElement=lambda *a: None),
    Line=NS(CreateBound=lambda a, b: (a, b)), DatumExtentType=NS(), DatumEnds=NS(),
    Grid=El, Level=El)
revit_pkg = types.ModuleType('Autodesk.Revit')
revit_pkg.DB = DB
sys.modules.setdefault('Autodesk', types.ModuleType('Autodesk'))
sys.modules['Autodesk.Revit'] = revit_pkg
sys.modules['Autodesk.Revit.DB'] = DB

from core import tool_specs as TS            # noqa: E402
from core import write_verify as W           # noqa: E402
from core.tools import room_views as RV, view_sync as VS   # noqa: E402
import Snippets._compat as C                 # noqa: E402
from Snippets import crop_sync as CS         # noqa: E402
from Snippets import datum_sync as DS        # noqa: E402
from Snippets import sheetgen as SG          # noqa: E402

C.make_eid = lambda v: v                     # element ids are plain ints here
SG._API['DB'] = DB


class Doc(object):
    def __init__(self):
        self.model, self.next = {}, 1000
        self.ActiveView = None

    def new_id(self):
        self.next += 1
        return self.next

    def add(self, el):
        self.model[el.Id] = el
        return el

    def GetElement(self, eid):
        return self.model.get(eid)

    def Regenerate(self):
        pass


# ── pure helpers ────────────────────────────────────────────────────────────
print('[helpers]')
check('to_ints keeps numbers, reports the rest', VS.to_ints([1, '22', 'x', None]) ==
      ([1, 22], ['x', None]))
check('flags default true, false only when said', VS.flags({'shape': False}, 'crop') ==
      {'shape': False, 'crop_active': True, 'crop_visible': True, 'annotation_crop': True})
names = ['A1 Metric: A1', 'A1 Metric: A3', 'Cover: A3']
check('pick_name: exact, case-insensitive', RV.pick_name('a1 metric: a1', names, C.fold) ==
      'A1 Metric: A1')
check('pick_name: type part when unique, None when ambiguous',
      RV.pick_name('A1', names, C.fold) == 'A1 Metric: A1'
      and RV.pick_name('A3', names, C.fold) is None)
check('offset: default 1 m, rejects negative / text / inf', RV._offset_ft(None) == 1000 / 304.8
      and RV._offset_ft(-1) is None and RV._offset_ft('x') is None
      and RV._offset_ft(float('inf')) is None and RV._offset_ft(0) == 0.0)
room = NS(Name='Kitchen', Number='101', RoomType='')
unit = NS(Name='UNIT A', Number='7', RoomType='T2')
check('sheetgen names', SG.build_view_name(room) == 'ENLARGED PLAN - Kitchen - (#101)'
      and SG.build_view_name(room, 2).endswith(' - Copy 2')
      and SG.build_view_name(unit) == 'ENLARGED PLAN - TYPE T2 (UNIT A)'
      and SG.build_sheet_number(room, 'P') == 'EPL-101-P')
check('direction label from ViewDirection',
      [SG.direction_label(NS(ViewDirection=XYZ(*d))) for d in Marker.DIRS] ==
      ['South', 'West', 'North', 'East'])
usable = SG.usable_rect(0, 0, 2.759, 1.949, SG.SHEET_MARGIN_FT, 'right', 0.23)
check('usable rect: margin and right strip', abs(usable[2] - (2.759 - 0.132 - 0.23)) < 1e-9)
check('clamp: too big centers and says so', SG.clamp_center(0, 0, 9, 9, usable)[2] is False)

# ── sync_views: crop ────────────────────────────────────────────────────────
print('[sync_views crop]')
doc = Doc()
src = doc.add(CropView(1, 'L1 - Plan', ViewType=VT.FloorPlan))
t1 = doc.add(CropView(2, 'L2 - Plan', ViewType=VT.FloorPlan))
t2 = doc.add(CropView(3, 'L3 - Plan', ViewType=VT.FloorPlan))
tpl = doc.add(CropView(4, 'Template', ViewType=VT.FloorPlan, IsTemplate=True))
sheet = doc.add(View(5, 'A101', ViewType=VT.DrawingSheet))
doc.ActiveView = src
CS.get_view_crop_data = lambda v: CS.ViewCropData(is_active=True, is_visible=False,
                                                  is_shaped=True, width=8000.0, height=6000.0)
APPLIED = []


def fake_crop(doc_, view, data, apply_shape=True, set_active=True, set_visible=True,
              sync_annotation_crop=True):
    APPLIED.append((view.Id, apply_shape))
    if view.Id == 3 and BROKEN.get('raise'):
        raise RuntimeError('crop shape is not valid')
    if set_active and not BROKEN.get('inactive'):
        view.CropBoxActive = True
    if set_visible:
        view.CropBoxVisible = data.is_visible
    return True


BROKEN = {}
CS.sync_crop_to_view = fake_crop
args = {'operation': 'crop', 'view_ids': [2, '3', 4, 5, 1, 99], 'dry_run': True}
r = TS.run('sync_views', doc, args)
reasons = [s['reason'] for s in r.get('skipped', [])]
check('dry run: targets listed, nothing applied', r.get('dry_run') and not APPLIED
      and r['targets'] == ['L2 - Plan [2]', 'L3 - Plan [3]'], r)
check('dry run: template, sheet, source and unknown id skipped with reasons',
      reasons == ['is a view template', 'has no crop region', 'is the source view',
                  'not a view in this model'], reasons)
check('source summary', r['source_data'] == {'shape': 'non-rectangular',
                                             'size_mm': [8000, 6000]}, r['source_data'])
del LOG[:]
r = TS.run('sync_views', doc, {'operation': 'crop', 'view_ids': [2, 3], 'shape': False})
check('run: both views updated, flags passed', r['success'] and r['updated_view_ids'] == [2, 3]
      and APPLIED == [(2, False), (3, False)], r)
check('one TransactionGroup, one Transaction per view, assimilated',
      [x for x in LOG if x[1] == 'T3Lab: Crop Sync'] ==
      [('start', 'T3Lab: Crop Sync'), ('commit', 'T3Lab: Crop Sync')]
      and len([x for x in LOG if x == ('commit', 'Crop Sync to L2 - Plan')]) == 1, LOG)
v = W.verify('sync_views', {}, r, 0, doc.model.__contains__)
check('write_verify: ids exist, self-check clean', v['verified'] and v['checks'] >= 3, v)
for view in (t1, t2):
    view.CropBoxActive = False
BROKEN['inactive'] = True
r = TS.run('sync_views', doc, {'operation': 'crop', 'view_ids': [2]})
check('read-back: crop not active reported', r['not_in_model_after'] ==
      ['L2 - Plan [2]: crop region not active'], r)
check('write_verify fails on it', W.verify('sync_views', {}, r, 0,
                                           doc.model.__contains__)['verified'] is False)
BROKEN.clear()
BROKEN['raise'] = True
r = TS.run('sync_views', doc, {'operation': 'crop', 'view_ids': [2, 3]})
check('one view fails, the other still synced', r['updated_view_ids'] == [2]
      and r['failed'] == [{'view': 'L3 - Plan [3]', 'reason': 'crop shape is not valid'}], r)
CS.sync_crop_to_view = lambda *a, **k: False
del LOG[:]
r = TS.run('sync_views', doc, {'operation': 'crop', 'view_ids': [2]})
check('nothing changed -> group rolled back, error', 'error' in r and not r['applied']
      and ('rollback', 'T3Lab: Crop Sync') in LOG, r)
check('bad operation', TS.run('sync_views', doc, {'operation': 'x', 'view_ids': [2]})[
    'error'].startswith('operation must'))
check('bad view ids', 'view_ids must' in TS.run('sync_views', doc, {
    'operation': 'crop', 'view_ids': ['abc']})['error'])
r = TS.run('sync_views', doc, {'operation': 'crop', 'source_view_id': 5, 'view_ids': [2]})
check('source without crop -> friendly error', 'has no crop region' in r['error'], r)
r = TS.run('sync_views', doc, {'operation': 'crop', 'view_ids': [4, 5]})
check('no usable target -> error with skipped', 'No target view' in r['error']
      and len(r['skipped']) == 2, r)

# ── sync_views: datums ──────────────────────────────────────────────────────
print('[sync_views datums]')
sec = doc.add(CropView(6, 'Section A', ViewType=VT.Section))
three_d = doc.add(CropView(7, '3D', ViewType=VT.ThreeD))
GRIDS = {1: {'A': DS.GridViewData('A', curves=[1], bubble0=True, bubble0_vis=False,
                                  bubble1=True, bubble1_vis=True)},
         2: {'A': DS.GridViewData('A', curves=[1], bubble0=True, bubble0_vis=True,
                                  bubble1=True, bubble1_vis=True)}}
LEVELS_READ = []
DS.get_view_grids_data = lambda d, view: GRIDS.get(view.Id, {})
DS.get_view_levels_data = lambda d, view: LEVELS_READ.append(view.Id) or {}


def fake_grids(d, view, data, sync_extents=True, sync_bubbles=True, sync_leaders=True):
    if sync_bubbles and not BROKEN.get('bubbles'):
        GRIDS[view.Id]['A'].bubble0_vis = data['A'].bubble0_vis
    return 1


DS.sync_grids_to_view = fake_grids
DS.sync_levels_to_view = lambda *a, **k: 0
r = TS.run('sync_views', doc, {'operation': 'datums', 'view_ids': [2, 6], 'dry_run': True})
check('plan source: section skipped, levels not read', r['targets'] == ['L2 - Plan [2]']
      and r['skipped'][0]['reason'] == 'not a plan view like the source' and not LEVELS_READ
      and r['source_data'] == {'grids': 1, 'levels': 0}, r)
r = TS.run('sync_views', doc, {'operation': 'datums', 'view_ids': [2]})
check('run: grid updates counted, bubbles read back', r['success'] and r['grid_updates'] == 1
      and 'not_in_model_after' not in r and GRIDS[2]['A'].bubble0_vis is False, r)
GRIDS[2]['A'].bubble0_vis = True
BROKEN['bubbles'] = True
r = TS.run('sync_views', doc, {'operation': 'datums', 'view_ids': [2]})
check('read-back: bubble that did not change', r['not_in_model_after'] ==
      ['L2 - Plan [2]: grid A bubble End0'], r)
BROKEN.clear()
r = TS.run('sync_views', doc, {'operation': 'datums', 'source_view_id': 7, 'view_ids': [2]})
check('3D source -> friendly error', 'not a plan, elevation or section' in r['error'], r)

# ── generate_room_views (sheetgen on the fake API) ─────────────────────────
print('[generate_room_views]')


def model():
    d = Doc()
    for i, (fam, vt) in enumerate((('Floor Plan', VT.FloorPlan), ('Ceiling Plan', VT.CeilingPlan),
                                   ('Interior Elevation', VT.Elevation)), 1):
        e = d.add(ViewFamilyType(10 + i, fam, FamilyName=fam, view_type=vt,
                                 ViewFamily='ElevationFamily' if vt == VT.Elevation else fam))
        e.params[BIP.SYMBOL_NAME_PARAM] = Param(fam)
    d.add(View(20, 'Enlarged 1:50', ViewType=VT.FloorPlan, IsTemplate=True))
    tb = d.add(El(30, 'A1', category='titleblocks', Family=NS(Name='A1 Metric')))
    tb.params[BIP.SYMBOL_NAME_PARAM] = Param('A1')
    level = NS(Id=40)
    d.add(ViewPlan(41, 'Level 1', ViewType=VT.FloorPlan, GenLevel=level))
    d.add(ViewPlan(42, 'ENLARGED PLAN - Kitchen - (#101)', ViewType=VT.FloorPlan,
                   GenLevel=level))
    old = d.add(ViewSheet(43, 'old'))
    old.params[BIP.SHEET_NUMBER] = Param('EPL-101')
    for rid, num, name, placed in ((50, '101', 'Kitchen', True), (51, '102', 'Bath', True),
                                   (52, '103', 'Store', False)):
        rm = d.add(El(rid, name, category='rooms', LevelId=40,
                      Location=NS(Point=XYZ(5, 5, 0)) if placed else None))
        rm.params[BIP.ROOM_NUMBER] = Param(num)
        rm.params[BIP.ROOM_NAME] = Param(name)
        rm.get_BoundingBox = (lambda v: NS(Min=XYZ(0, 0, 0), Max=XYZ(10, 8, 3)))
    return d


d = model()
before = len(d.model)
r = TS.run('generate_room_views', d, {'room_ids': [50, 52, 30], 'ceiling': True,
                                       'elevations': True, 'layout': 'separate',
                                       'templates': {'floor': 'enlarged 1:50'},
                                       'dry_run': True})
row = r['rooms'][0]
check('dry run: nothing created', len(d.model) == before and r['dry_run'] and not r['applied'])
check('dry run: existing plan name gets "(2)", ceiling keeps the base name',
      row['views'][:2] == ['ENLARGED PLAN - Kitchen - (#101) (2)',
                           'ENLARGED PLAN - Kitchen - (#101)'], row)
check('dry run: four elevation names', len([v for v in row['views']
                                            if v.startswith('INTERIOR ELEV')]) == 4, row)
check('dry run: sheet numbers, existing EPL-101 untouched', row['sheets'] == [
    'EPL-101-P - ENLARGED PLAN - Kitchen - (#101) — Plans',
    'EPL-101-E - ENLARGED PLAN - Kitchen - (#101) — Elevations'], row)
check('dry run: unplaced room and non-room skipped', [s['reason'] for s in r['skipped']] ==
      ['not placed', 'not a room in this model'] and r['view_count'] == 6
      and r['sheet_count'] == 2, r)
check('template matched case-insensitively, default title block',
      r['templates'] == {'floor': 'Enlarged 1:50'} and r['titleblock'] == 'A1 Metric: A1', r)
del LOG[:]
Viewport.made[:] = []
r = TS.run('generate_room_views', d, {'room_ids': [50, 51], 'elevations': True,
                                       'layout': 'combined', 'titleblock': 'A1',
                                       'templates': {'floor': 'Enlarged 1:50'}})
check('run: 2 rooms x (plan + 4 elevations), one sheet each', r['success']
      and r['view_count'] == 10 and r['sheet_count'] == 2 and len(Viewport.made) == 10, r)
check('one outer TransactionGroup assimilated',
      LOG[0] == ('start', 'T3Lab: Room views and sheets')
      and LOG[-1] == ('commit', 'T3Lab: Room views and sheets'), LOG[:2] + LOG[-2:])
plan = d.model[r['view_ids'][0]]
check('plan: template applied, cropped, unique name', plan.ViewTemplateId == 20
      and plan.CropBoxActive and plan.Name == 'ENLARGED PLAN - Kitchen - (#101) (2)', plan.Name)
check('elevations named from their direction', sorted(r['rooms'][0]['views'][1:]) == sorted(
    'INTERIOR ELEV - Kitchen - {} (101)'.format(x) for x in ('North', 'East', 'South', 'West')),
    r['rooms'][0])
check('sheet numbers', [s.split(' - ')[0] for row in r['rooms'] for s in row['sheets']] ==
      ['EPL-101-2', 'EPL-102'], r['rooms'])
v = W.verify('generate_room_views', {}, r, 0, d.model.__contains__, changes=(12, 0, 0))
check('write_verify: views and sheets exist', v['verified'], v)
r = TS.run('generate_room_views', d, {'room_ids': [50], 'layout': 'combined',
                                       'titleblock': 'Nope'})
check('unknown title block -> names', 'No title block' in r['error']
      and r['titleblocks'] == ['A1 Metric: A1'], r)
r = TS.run('generate_room_views', d, {'room_ids': [50], 'templates': {'floor': 'X'}})
check('unknown template -> names', r['templates'] == ['Enlarged 1:50'], r)
check('templates with a wrong key', 'templates takes' in TS.run(
    'generate_room_views', d, {'room_ids': [50], 'templates': {'plan': 'X'}})['error'])
check('nothing asked', 'Nothing to create' in TS.run(
    'generate_room_views', d, {'room_ids': [50], 'floor': False})['error'])
check('bad layout / offset', 'layout must' in TS.run(
    'generate_room_views', d, {'room_ids': [50], 'layout': 'grid'})['error']
      and 'offset_mm' in TS.run('generate_room_views', d, {'room_ids': [50],
                                                           'offset_mm': -5})['error'])
check('only unplaced rooms', 'No placed room' in TS.run(
    'generate_room_views', d, {'room_ids': [52]})['error'])
del d.model[13]                                   # no elevation type
check('missing view type named', 'no elevation view type' in TS.run(
    'generate_room_views', d, {'room_ids': [50], 'elevations': True})['error'])
d = model()
d.model[50].get_BoundingBox = lambda v: None
del LOG[:]
r = TS.run('generate_room_views', d, {'room_ids': [50]})
check('no view made -> rollback + error + reason', 'error' in r and not r['applied']
      and 'no bounding box' in r['failed'][0]
      and LOG[-1] == ('rollback', 'T3Lab: Room views and sheets'), r)

# ── declarations + server ───────────────────────────────────────────────────
print('[specs]')
for name in ('sync_views', 'generate_room_views'):
    spec = TS.get(name)
    size = len(json.dumps(spec.registry_entry())) // 4
    check('{}: complete'.format(name), not spec.problems(), spec.problems())
    check('{}: registry entry {} tokens <= 350'.format(name, size), size <= 350)
    check('{}: never destructive'.format(name), not spec.is_destructive({}))
import core.server as S   # noqa: E402
for name in ('sync_views', 'generate_room_views'):
    check('{}: runs on the main thread'.format(name), name in S.T3LabAIServer._WRITE_TOOLS)
check('no document', 'No active document' in TS.run('sync_views', None, {})['error'])

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
