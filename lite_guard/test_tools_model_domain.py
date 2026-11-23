# -*- coding: utf-8 -*-
"""core/tools/model.py: auto_join, room_to_floor, rehost_to_level,
cut_wall_openings, create_door_thresholds against a fake Revit API.

The shipped Snippets/finish_floors.py, Snippets/wall_tools.py and
Services/join_service.py run here; only Autodesk / System / pyrevit are faked.

Run:  python lite_guard/test_tools_model_domain.py
"""
import ast
import io
import logging
import os
import sys
import tempfile
import types
from types import SimpleNamespace

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
sys.path.insert(0, os.path.join(REPO, 'lite_guard'))
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_model_tools_test_')

logging.getLogger('T3Lab.finish_floors').addHandler(logging.NullHandler())
logging.getLogger('T3Lab.finish_floors').propagate = False
FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ── fake Revit API ──────────────────────────────────────────────────────────
class Cat(int):
    @property
    def value__(self):
        return int(self)


BIC = dict(OST_Walls=-2000011, OST_Floors=-2000032, OST_Doors=-2000023, OST_Rooms=-2000160,
           OST_StructuralColumns=-2001330, OST_Columns=-2000100,
           OST_StructuralFraming=-2001320, OST_StructuralFoundation=-2001300,
           OST_Roofs=-2000035, OST_Ceilings=-2000038, OST_GenericModel=-2000151,
           OST_Windows=-2000014, OST_DuctCurves=-2008000, OST_PipeCurves=-2008044,
           OST_CableTray=-2008130)


class AnyName(object):
    def __getattr__(self, name):
        if name.startswith('__'):
            raise AttributeError(name)
        return name


class Id(object):
    def __init__(self, v):
        self.Value = self.IntegerValue = int(v)

    def __eq__(self, other):
        return isinstance(other, Id) and other.Value == self.Value

    def __ne__(self, other):
        return not self == other

    def __hash__(self):
        return hash(self.Value)


class XYZ(object):
    def __init__(self, x, y, z):
        self.X, self.Y, self.Z = float(x), float(y), float(z)

    def __add__(self, o):
        return XYZ(self.X + o.X, self.Y + o.Y, self.Z + o.Z)

    def __sub__(self, o):
        return XYZ(self.X - o.X, self.Y - o.Y, self.Z - o.Z)

    def __mul__(self, k):
        return XYZ(self.X * k, self.Y * k, self.Z * k)

    def DotProduct(self, o):
        return self.X * o.X + self.Y * o.Y + self.Z * o.Z


class Box(object):
    def __init__(self, a, b):
        self.Min, self.Max = XYZ(*a), XYZ(*b)


class P(object):
    """Parameter: number, string or element id."""

    def __init__(self, value, read_only=False):
        self.value, self.IsReadOnly = value, read_only

    def AsDouble(self):
        return float(self.value)

    def AsString(self):
        return self.value

    def AsInteger(self):
        return int(self.value)

    def AsElementId(self):
        return self.value

    def Set(self, value):
        self.value = value


LOG = []
_NEXT = [1000]


class El(object):
    category = None

    def __init__(self, i=None, name='', params=None, bic=None):
        if i is None:
            _NEXT[0] += 1
            i = _NEXT[0]
        self.Id, self.Name = Id(i), name
        self.params = dict(params or {})
        bic = bic or self.category
        self.Category = SimpleNamespace(Id=Id(BIC[bic])) if bic else None
        self.box = None

    def get_Parameter(self, bip):
        return self.params.get(bip)

    def LookupParameter(self, name):
        return self.params.get(name)

    def get_BoundingBox(self, view):
        return self.box


class Level(El):
    def __init__(self, i, name, elevation):
        El.__init__(self, i, name)
        self.Elevation = elevation


class Wall(El):
    category = 'OST_Walls'

    def __init__(self, i, box, kind='Basic', **kw):
        El.__init__(self, i, 'Wall {}'.format(i), **kw)
        self.box = Box(*box)
        self.WallType = SimpleNamespace(Kind=kind)
        self.Location = SimpleNamespace(Curve='line')
        self.Width = 0.5


class Floor(El):
    category = 'OST_Floors'

    @staticmethod
    def Create(doc, profile, type_id, level_id):
        LOG.append(('Floor.Create', profile.Count, type_id.Value, level_id.Value))
        if getattr(doc, 'refuse_floors', False):
            raise RuntimeError('Revit says no')
        return doc.add(Floor(params={'FLOOR_HEIGHTABOVELEVEL_PARAM': P(0.0),
                                     'FLOOR_PARAM_IS_STRUCTURAL': P(0)}))


class FloorType(El):
    def __init__(self, i, family, name):
        El.__init__(self, i, name, params={'Type Name': P(name)})
        self.FamilyName = family
        self.is_type = True


class FamilyInstance(El):
    pass


class FamilySymbol(El):
    category = 'OST_GenericModel'

    def __init__(self, i, family, name, hosted=True):
        El.__init__(self, i, name, params={'ALL_MODEL_TYPE_NAME': P(name)})
        self.Family = SimpleNamespace(Name=family, get_Parameter=lambda bip: P(1 if hosted else 0))
        self.IsActive, self.is_type = True, True


class RevitLinkInstance(El):
    def __init__(self, i, ldoc, shift):
        El.__init__(self, i, 'link')
        self.ldoc, self.shift = ldoc, shift

    def GetLinkDocument(self):
        return self.ldoc

    def GetTotalTransform(self):
        return SimpleNamespace(OfPoint=lambda p: p + self.shift)


class Room(El):
    category = 'OST_Rooms'

    def __init__(self, i, level, area=10.0, loops=1):
        El.__init__(self, i, 'Room', params={'ROOM_AREA': P(area), 'Number': P('10{}'.format(i)),
                                             'Name': P('Office')})
        self.LevelId, self.loops, self.asked = level.Id, loops, []

    def GetBoundarySegments(self, opt):
        self.asked.append(opt.SpatialElementBoundaryLocation)
        seg = SimpleNamespace(GetCurve=lambda: 'curve')
        return [[seg, seg, seg]] * self.loops


class Door(FamilyInstance):
    category = 'OST_Doors'

    def __init__(self, i, level, host, mark):
        FamilyInstance.__init__(self, i, 'Door', params={'DOOR_WIDTH': P(3.0),
                                                       'ALL_MODEL_MARK': P(mark)})
        self.LevelId, self.Host = level.Id, host
        self.Location = SimpleNamespace(Point=XYZ(10, 0, 0))
        self.HandOrientation, self.FacingOrientation = XYZ(1, 0, 0), XYZ(0, 1, 0)


class CurveLoop(object):
    def __init__(self):
        self.curves = []

    def Append(self, c):
        self.curves.append(c)

    def IsOpen(self):
        return not self.curves


class NetList(list):
    def Add(self, x):
        self.append(x)

    @property
    def Count(self):
        return len(self)


class _Generic(object):
    def __getitem__(self, t):
        return NetList


class Txn(object):
    kind = 'Transaction'

    def __init__(self, doc, name):
        self.name, self.state = name, None
        LOG.append((self.kind, name))

    def Start(self):
        self.state = 'started'
        LOG.append(('start', self.name))

    def Commit(self):
        self.state = 'committed'
        LOG.append(('commit', self.name))
        return 'Committed'

    def RollBack(self):
        self.state = 'rolledback'
        LOG.append(('rollback', self.name))

    def Assimilate(self):
        self.state = 'assimilated'
        LOG.append(('assimilate', self.name))

    def HasStarted(self):
        return self.state is not None

    def HasEnded(self):
        return self.state not in (None, 'started')

    def GetStatus(self):
        return 'Started' if self.state == 'started' else self.state

    def GetFailureHandlingOptions(self):
        return SimpleNamespace(SetFailuresPreprocessor=lambda p: None,
                               SetForcedModalHandling=lambda v: None)

    def SetFailureHandlingOptions(self, opts):
        pass

    def Dispose(self):
        pass


class Group(Txn):
    kind = 'TransactionGroup'


class Collector(object):
    def __init__(self, doc, view_id=None):
        self.items = [e for e in doc.model.values()]

    def OfClass(self, cls):
        self.items = [e for e in self.items if isinstance(e, cls)]
        return self

    def OfCategory(self, cat):
        self.items = [e for e in self.items if e.Category is not None
                      and e.Category.Id.Value == int(cat)]
        return self

    def WhereElementIsNotElementType(self):
        self.items = [e for e in self.items if not getattr(e, 'is_type', False)]
        return self

    def ToElements(self):
        return list(self.items)

    def __iter__(self):
        return iter(list(self.items))


class Creator(object):
    def __init__(self, doc):
        self.doc = doc

    def NewFamilyInstance(self, point, symbol, host, structural):
        return self.doc.add(FamilyInstance(params={'Width': P(0.0), 'Height': P(0.0)},
                                           bic='OST_GenericModel'))

    def NewOpening(self, wall, p0, p1):
        return self.doc.add(El(name='Opening'))


class Doc(object):
    def __init__(self, title='Tower.rvt'):
        self.Title, self.model = title, {}
        self.Create = Creator(self)
        self.ActiveView = self.add(El(5, 'Level 1 plan'))

    def add(self, el):
        self.model[el.Id.Value] = el
        return el

    def GetElement(self, i):
        return self.model.get(i.Value if isinstance(i, Id) else int(i))

    def Regenerate(self):
        pass


def install_fakes():
    db = types.ModuleType('Autodesk.Revit.DB')
    names = dict(
        BuiltInCategory=type('BuiltInCategory', (), dict((k, Cat(v)) for k, v in BIC.items())),
        BuiltInParameter=AnyName(), Transaction=Txn, TransactionGroup=Group,
        TransactionStatus=SimpleNamespace(Committed='Committed', Started='Started'),
        FilteredElementCollector=Collector, IFailuresPreprocessor=object,
        FailureProcessingResult=SimpleNamespace(Continue=0),
        FailureSeverity=SimpleNamespace(Warning=1, Error=2),
        SpatialElementBoundaryOptions=lambda: SimpleNamespace(
            SpatialElementBoundaryLocation=None),
        SpatialElementBoundaryLocation=SimpleNamespace(Finish='Finish', Center='Center'),
        CurveLoop=CurveLoop, CurveArray=CurveLoop, Floor=Floor, FloorType=FloorType,
        Level=Level, Wall=Wall, FamilyInstance=FamilyInstance, FamilySymbol=FamilySymbol,
        RevitLinkInstance=RevitLinkInstance, XYZ=XYZ,
        WallKind=SimpleNamespace(Basic='Basic', Curtain='Curtain', Stacked='Stacked'),
        Line=SimpleNamespace(CreateBound=lambda a, b: (a, b)),
        HostObjectUtils=SimpleNamespace(GetSideFaces=lambda w, s: None),
        ShellLayerType=SimpleNamespace(Exterior=0, Interior=1),
        JoinGeometryUtils=None, BoundingBoxIntersectsFilter=None, Outline=None, Options=None,
        ViewDetailLevel=None, Solid=type('Solid', (), {}),
        GeometryInstance=type('GeometryInstance', (), {}),
        GeometryElement=type('GeometryElement', (), {}), BooleanOperationsUtils=None,
        BooleanOperationsType=None, ElementId=Id)
    for k, v in names.items():
        setattr(db, k, v)
    structure = types.ModuleType('Autodesk.Revit.DB.Structure')
    structure.StructuralType = SimpleNamespace(NonStructural=0)
    db.Structure = structure
    autodesk = sys.modules.setdefault('Autodesk', types.ModuleType('Autodesk'))
    revit = sys.modules.setdefault('Autodesk.Revit', types.ModuleType('Autodesk.Revit'))
    autodesk.Revit, revit.DB = revit, db
    sys.modules['Autodesk.Revit.DB'] = db
    sys.modules['Autodesk.Revit.DB.Structure'] = structure
    system = sys.modules.setdefault('System', types.ModuleType('System'))
    coll = sys.modules.setdefault('System.Collections', types.ModuleType('System.Collections'))
    gen = types.ModuleType('System.Collections.Generic')
    gen.List = _Generic()
    system.Collections, coll.Generic = coll, gen
    sys.modules['System.Collections.Generic'] = gen
    pyrevit = types.ModuleType('pyrevit')

    class _Revit(object):
        doc = None

        @property
        def uidoc(self):
            raise AssertionError('the tool must not read the UI selection')

    pyrevit.revit = _Revit()
    pyrevit.script = SimpleNamespace(get_logger=lambda: logging.getLogger('test'))
    sys.modules['pyrevit'] = pyrevit


install_fakes()
import Snippets._compat as C               # noqa: E402
C.make_eid = lambda v: v if isinstance(v, Id) else Id(int(v))
C.revit_year = lambda doc=None: 2025

from core import tool_specs as TS          # noqa: E402
from core.tools import model as M          # noqa: E402
from core import write_verify as W         # noqa: E402
from Services import join_service as JS    # noqa: E402
from Snippets import finish_floors as FF   # noqa: E402
from Snippets import wall_tools as WT      # noqa: E402


def exists_in(doc):
    return lambda i: doc.GetElement(i) is not None


# ── pure helpers ────────────────────────────────────────────────────────────
print('[helpers]')
check('int_ids: ints, digit strings, duplicates, bad', M.int_ids([3, '4', 3, 'x', ' 5 ']) ==
      ([3, 4, 5], ['x']))
check('int_ids: one value, None', M.int_ids(7) == ([7], []) and M.int_ids(None) == ([], []))
rules, bad = M.parse_rules([{'priority': 'walls', 'join_with': 'FLOORS'},
                            {'priority': 'Beams', 'join_with': 'Walls'}],
                           JS.CATEGORY_NAMES, JS.DEFAULT_RULES)
check('parse_rules: case-insensitive, unknown reported',
      rules == [{'priority': 'Walls', 'join_with': 'Floors'}] and bad == ['Beams cuts Walls'],
      (rules, bad))
check('parse_rules: none = the 7 default rules',
      M.parse_rules(None, JS.CATEGORY_NAMES, JS.DEFAULT_RULES)[0] == JS.DEFAULT_RULES)
types_ = {'Floor: Generic 150': 'g150', 'Floor: Finish 50': 'f50', 'Slab: Generic 150': 's'}
check('pick_floor_type: default = first label',
      FF.pick_floor_type(types_) == ('f50', 'Floor: Finish 50'))
check('pick_floor_type: full label, type name, part', FF.pick_floor_type(types_, 'slab: generic 150')
      == ('s', 'Slab: Generic 150') and FF.pick_floor_type(types_, 'finish 50')[0] == 'f50'
      and FF.pick_floor_type(types_, 'Generic')[0] == 'g150')
check('pick_floor_type: no match / no types', FF.pick_floor_type(types_, 'Tiles') == (None, None)
      and FF.pick_floor_type({}, None) == (None, None))
lv = [Level(1, 'Level 1', 0.0), Level(2, 'Level 10', 30.0), Level(3, 'Roof', 40.0)]
check('pick_name: exact before part, id', M.pick_name(lv, 'level 1', lambda l: l.Name) is lv[0]
      and M.pick_name(lv, 'roo', lambda l: l.Name) is lv[2]
      and M.pick_name(lv, '2', lambda l: l.Name) is lv[1]
      and M.pick_name(lv, 'Basement', lambda l: l.Name) is None)


# ── room_to_floor ───────────────────────────────────────────────────────────
print('[room_to_floor]')
doc = Doc()
L1 = doc.add(Level(10, 'Level 1', 0.0))
doc.add(FloorType(20, 'Floor', 'Generic 150'))
doc.add(FloorType(21, 'Floor', 'Finish 50'))
r1, r2 = doc.add(Room(31, L1)), doc.add(Room(32, L1, loops=2))
unplaced = doc.add(Room(33, L1, area=0.0))
wall = doc.add(Wall(40, ((0, -1, 0), (20, 1, 10))))
del LOG[:]
r = TS.run('room_to_floor', doc, {'room_ids': [31, '32', 33, 40, 99], 'dry_run': True})
check('dry run: rooms listed, problems skipped with reason, nothing made',
      r['dry_run'] and r['rooms'] == 2 and len(r['room_list']) == 2
      and [s['reason'] for s in r['skipped']] == ['not placed or not enclosed (area 0)',
                                                  'not a room', 'not in the model']
      and not [e for e in LOG if e[0] == 'Floor.Create'], r)
check('default floor type = first label', r['floor_type'] == 'Floor: Finish 50', r)
r = TS.run('room_to_floor', doc, {'room_id': 31, 'floor_type': 'Tiles'})
check('unknown floor type -> list', 'Tiles' in r['error'] and 'Floor: Generic 150' in
      r['floor_types'], r)
check('no rooms -> hint', 'room_ids' in TS.run('room_to_floor', doc, {})['error'])
r = TS.run('room_to_floor', doc, {'room_ids': [31, 32], 'floor_type': 'generic 150',
                                  'offset_mm': 304.8, 'structural': True, 'boundary': 'center'})
floors = [doc.GetElement(i) for i in r.get('floor_ids', [])]
check('run: one floor per room, ids reported', r['success'] and r['floors_created'] == 2
      and len(floors) == 2 and all(floors), r)
check('options reach the floor: type, offset in ft, structural, centre boundary',
      [e[2] for e in LOG if e[0] == 'Floor.Create'] == [20, 20]
      and floors[0].params['FLOOR_HEIGHTABOVELEVEL_PARAM'].value == 1.0
      and floors[0].params['FLOOR_PARAM_IS_STRUCTURAL'].value == 1
      and r1.asked[-1] == 'Center', (LOG, r1.asked))
check('a room with an inner loop keeps both loops', [e[1] for e in LOG
                                                    if e[0] == 'Floor.Create'][1] == 2)
check('one TransactionGroup assimilated (one undo)',
      [e for e in LOG if e[0] in ('TransactionGroup', 'assimilate')][:1] ==
      [('TransactionGroup', 'T3Lab: Room to Floor')] and ('assimilate', 'T3Lab: Room to Floor')
      in LOG)
v = W.verify('room_to_floor', {}, r, 0, exists_in(doc), changes=(2, 0, 0))
check('write_verify: floors exist + a change', v['verified'] and v['checks'] == 3, v)
doc.refuse_floors = True
r = TS.run('room_to_floor', doc, {'room_ids': [31]})
check('Revit refuses -> error with the reason per room', 'No floor' in r['error']
      and r['failed'] == [{'room': 31, 'reason': 'Revit says no'}], r)
doc.refuse_floors = False


# ── create_door_thresholds ──────────────────────────────────────────────────
print('[create_door_thresholds]')
d1, d2 = doc.add(Door(51, L1, wall, 'D1')), doc.add(Door(52, L1, None, ''))
del LOG[:]
r = TS.run('create_door_thresholds', doc, {'element_ids': [51, 52, 31], 'dry_run': True})
check('dry run: doors listed by mark / id, non-doors skipped', r['dry_run']
      and r['door_list'] == ['Door D1', 'Door 52'] and r['skipped'][0]['element'] == 31
      and not [e for e in LOG if e[0] == 'Floor.Create'], r)
r = TS.run('create_door_thresholds', doc, {'element_ids': [51, 52], 'offset_mm': 15})
check('run: one threshold per door inside one group', r['success'] and r['created_count'] == 2
      and all(doc.GetElement(i) for i in r['threshold_ids'])
      and ('assimilate', 'T3Lab: Door Threshold') in LOG, (r, LOG))
v = W.verify('create_door_thresholds', {}, r, 0, exists_in(doc), changes=(2, 0, 0))
check('write_verify: thresholds exist', v['verified'], v)
check('not a door -> error', 'placed door' in TS.run('create_door_thresholds', doc,
                                                     {'element_ids': [40]})['error'])
doc.refuse_floors = True
del LOG[:]
r = TS.run('create_door_thresholds', doc, {'element_ids': [51]})
check('nothing created -> group rolled back, reasons listed', 'error' in r
      and ('rollback', 'T3Lab: Door Threshold') in LOG and 'Door D1' in r['failed'][0], r)
doc.refuse_floors = False


# ── rehost_to_level ─────────────────────────────────────────────────────────
print('[rehost_to_level]')
L2 = doc.add(Level(11, 'Level 2', 10.0))
w2 = doc.add(Wall(41, ((0, 0, 0), (1, 1, 1)), params={
    'WALL_BASE_CONSTRAINT': P(L1.Id), 'WALL_BASE_OFFSET': P(12.0),
    'WALL_HEIGHT_TYPE': P(Id(-1)), 'WALL_TOP_OFFSET': P(0.0)}))
r = TS.run('rehost_to_level', doc, {'element_ids': [41], 'level': 'Level 9'})
check('unknown level -> list', r['levels'][:2] == ['Level 1', 'Level 2'], r)
r = TS.run('rehost_to_level', doc, {'element_ids': [41, 40, 31], 'level': 'level 2',
                                    'dry_run': True})
check('dry run: offsets in mm, unsupported skipped', r['dry_run']
      and r['elements'] == [{'id': 41, 'from_level': 'Level 1', 'offset_mm': 3657.6,
                             'new_offset_mm': 609.6}]
      and [s['element'] for s in r['skipped']] == [40, 31]
      and w2.params['WALL_BASE_CONSTRAINT'].value == L1.Id, r)
r = TS.run('rehost_to_level', doc, {'element_ids': [41], 'level': 'Level 2', 'end': 'top'})
check('top unconnected -> reason', r['skipped'][0]['reason'] == 'top is unconnected', r)
del LOG[:]
r = TS.run('rehost_to_level', doc, {'element_ids': ['41'], 'level': '11'})
check('run: level + offset set, same elevation', r['success'] and r['element_ids'] == [41]
      and w2.params['WALL_BASE_CONSTRAINT'].value == L2.Id
      and w2.params['WALL_BASE_OFFSET'].value == 2.0, (r, w2.params))
check('one group around one transaction', [e for e in LOG if e[0] in ('assimilate', 'commit')]
      == [('commit', 'T3Lab: Auto Adjust Offset'), ('assimilate', 'T3Lab: Rehost to Level')], LOG)
v = W.verify('rehost_to_level', {}, r, 0, exists_in(doc))
check('read-back clean -> verified', v['verified'] and 'not_in_model_after' not in r, (r, v))
orig = WT.adjust_element
WT.adjust_element = lambda d, el, lvl, base=True: True      # claims done, sets nothing
r = TS.run('rehost_to_level', doc, {'element_ids': [41], 'level': 'Level 1'})
WT.adjust_element = orig
check('read-back catches an element still on the old level',
      r['not_in_model_after'] == [41] and not W.verify('rehost_to_level', {}, r, 0,
                                                       exists_in(doc))['verified'], r)


# ── cut_wall_openings ───────────────────────────────────────────────────────
print('[cut_wall_openings]')
cdoc = Doc('MEP host.rvt')
cw = cdoc.add(Wall(60, ((0, -0.5, 0), (20, 0.5, 10))))
cdoc.add(Wall(61, ((0, 5, 0), (20, 6, 10)), kind='Curtain'))
duct = cdoc.add(El(70, 'Duct', bic='OST_DuctCurves'))
duct.box = Box((4, -3, 4), (6, 3, 5))                      # crosses wall 60
far = cdoc.add(El(71, 'Duct', bic='OST_DuctCurves'))
far.box = Box((4, 30, 4), (6, 33, 5))
col = cdoc.add(El(72, 'Column', bic='OST_Columns'))
col.box = Box((9, -1, 0), (10, 1, 10))
ldoc = Doc('Plumbing.rvt')
pipe = ldoc.add(El(80, 'Pipe', bic='OST_PipeCurves'))
pipe.box = Box((-1, -40, 2), (1, -36, 3))                  # crosses once moved by +38 in Y
cdoc.add(RevitLinkInstance(90, ldoc, XYZ(2, 38, 0)))
cdoc.add(FamilySymbol(95, 'Wall Opening Void', '200x200'))
cdoc.add(FamilySymbol(96, 'Window Fixed', '900x1200', hosted=False))
r = TS.run('cut_wall_openings', cdoc, {'dry_run': True})
check('dry run (active view, default categories): duct crossing counted, curtain wall skipped',
      r['dry_run'] and r['walls'] == 1 and r['openings_planned'] == 1
      and r['by_category'] == {'Ducts': 1} and r['opening_families'] ==
      ['Wall Opening Void : 200x200'], r)
r = TS.run('cut_wall_openings', cdoc, {'dry_run': True, 'include_links': True,
                                       'categories': ['pipes', 'columns']})
check('links (transformed) and chosen categories', r['by_category'] == {'Columns': 1, 'Pipes': 1}
      and r['scanned'] == ['MEP host.rvt', 'Plumbing.rvt'], r)
check('bad category / method', 'Unknown categories' in TS.run(
    'cut_wall_openings', cdoc, {'categories': ['beams']})['error']
      and 'method' in TS.run('cut_wall_openings', cdoc, {'method': 'profile'})['error'])
r = TS.run('cut_wall_openings', cdoc, {'element_ids': [61], 'dry_run': True})
check('only curtain walls -> error with skipped', 'No wall to cut' in r['error']
      and r['skipped'][0]['element'] == 61, r)
check('unknown opening family -> list', TS.run('cut_wall_openings', cdoc, {
    'opening_family': 'Sleeve'})['opening_families'] == ['Wall Opening Void : 200x200'])
del LOG[:]
r = TS.run('cut_wall_openings', cdoc, {'element_ids': [60], 'clearance_mm': 50})
made = cdoc.GetElement(r['opening_ids'][0]) if r.get('opening_ids') else None
check('run: family placed and sized (box + 2 x clearance)', r['success'] and r['created_count']
      == 1 and made is not None and abs(made.params['Width'].value - (6 + 100 / 304.8)) < 1e-9
      and abs(made.params['Height'].value - (1 + 100 / 304.8)) < 1e-9, r)
check('one group around the transaction', ('assimilate', 'T3Lab: Wall Cut Profile') in LOG
      and ('commit', 'T3Lab: Wall Cut Profile') in LOG, LOG)
v = W.verify('cut_wall_openings', {}, r, 0, exists_in(cdoc), changes=(1, 0, 0))
check('write_verify: opening exists', v['verified'], v)
r = TS.run('cut_wall_openings', cdoc, {'element_ids': [60], 'method': 'wall_opening',
                                       'categories': ['columns']})
check('wall_opening method', r['success'] and cdoc.GetElement(r['opening_ids'][0]).Name ==
      'Opening', r)
r = TS.run('cut_wall_openings', cdoc, {'element_ids': [60], 'categories': ['cable_trays']})
check('nothing crosses -> error, group rolled back', 'no cutting element' in r['error']
      and LOG[-1][0] == 'rollback', (r, LOG[-3:]))


# ── auto_join ───────────────────────────────────────────────────────────────
print('[auto_join]')
A = [El(i, 'Wall', bic='OST_Walls') for i in (1, 2)]
B = [El(i, 'Floor', bic='OST_Floors') for i in (3, 4)]
JOINED = {(1, 3)}
CALLS = []


def fake_collect(document, bic, scope, view_id=None, selected_ids=None):
    CALLS.append((int(bic), scope, selected_ids))
    pool = A if int(bic) == BIC['OST_Walls'] else B
    if scope == 'Selected Elements':
        pool = [e for e in pool if e.Id.Value in set(selected_ids)]
    return pool


JS._collect_elements = fake_collect
JS._get_intersecting_elements = lambda d, el, bic, scope, view_id=None: (
    B if int(bic) == BIC['OST_Floors'] else [a for a in A if a is not el])
JS._same_design_option = lambda a, b: set([a.Id.Value, b.Id.Value]) != set([2, 4])
JS.JoinGeometryUtils = SimpleNamespace(
    AreElementsJoined=lambda d, a, b: (min(a.Id.Value, b.Id.Value),
                                       max(a.Id.Value, b.Id.Value)) in JOINED)
jdoc = Doc('Joins.rvt')
r = TS.run('auto_join', jdoc, {'rules': [{'priority': 'Walls', 'join_with': 'Floors'},
                                         {'priority': 'Floors', 'join_with': 'Walls'}],
                               'dry_run': True})
check('dry run: pairs per rule, other design option and repeats not counted twice',
      r['dry_run'] and r['per_rule'][0] == {'rule': 'Walls cuts Floors', 'pairs': 4,
                                            'to_change': 2, 'unchanged': 1,
                                            'other_design_option': 1}
      and r['per_rule'][1] == {'rule': 'Floors cuts Walls', 'pairs': 4, 'to_change': 0,
                               'unchanged': 3, 'other_design_option': 1}
      and r['pairs_to_change'] == 2
      and r['scope'] == 'Entire Project', r)
r = TS.run('auto_join', jdoc, {'mode': 'unjoin', 'dry_run': True,
                               'rules': [{'priority': 'Walls', 'join_with': 'Floors'}]})
check('unjoin dry run counts joined pairs', r['pairs_to_change'] == 1, r)
del CALLS[:]
TS.run('auto_join', jdoc, {'element_ids': ['2'], 'dry_run': True})
check('element_ids -> Selected Elements scope with those ids', CALLS[0][1:] ==
      ('Selected Elements', [2]), CALLS[:1])
check('bad rule / scope / ids', 'Beams' in TS.run('auto_join', jdoc, {
    'rules': [{'priority': 'Beams', 'join_with': 'Walls'}]})['error']
      and 'scope' in TS.run('auto_join', jdoc, {'scope': 'view'})['error']
      and 'empty' in TS.run('auto_join', jdoc, {'element_ids': []})['error'])

# run_join itself (shipped code): selected_ids reach the collector, no UI selection read.
JS.Transaction = Txn
JS.JoinFailuresPreprocessor = lambda: None
done = []
JS.JoinGeometryUtils = SimpleNamespace(
    AreElementsJoined=lambda d, a, b: (a.Id.Value, b.Id.Value) in done,
    JoinGeometry=lambda d, a, b: done.append((a.Id.Value, b.Id.Value)),
    IsCuttingElementInJoin=lambda d, a, b: True, SwitchJoinOrder=lambda d, a, b: None)
del CALLS[:]
r = TS.run('auto_join', jdoc, {'element_ids': [1], 'protect_embedded': False,
                               'rules': [{'priority': 'Walls', 'join_with': 'Floors'}]})
check('run: selected ids joined, counts returned', r['success'] and r['pairs_changed'] == 2
      and done == [(1, 3), (1, 4)] and CALLS[0][2] == [1], (r, CALLS))
v = W.verify('auto_join', {}, r, 0, exists_in(jdoc), changes=(0, 4, 0))
check('write_verify: must_change passes', v['verified'], v)
r = TS.run('auto_join', jdoc, {'element_ids': [1], 'protect_embedded': False,
                               'rules': [{'priority': 'Walls', 'join_with': 'Floors'}]})
check('nothing left to join -> success false + message (no false read-back alarm)',
      r['success'] is False and 'Nothing to join' in r['message']
      and W.verify('auto_join', {}, r, 0, exists_in(jdoc), changes=(0, 0, 0)) is None, r)


# ── declarations / server ───────────────────────────────────────────────────
print('[server]')
import core.server as S   # noqa: E402
import server_registry    # noqa: E402
srv = object.__new__(S.T3LabAIServer)
NEW = ('auto_join', 'room_to_floor', 'rehost_to_level', 'cut_wall_openings',
       'create_door_thresholds')
for tool in NEW:
    spec = TS.get(tool)
    check('{}: a complete spec on the main thread'.format(tool), spec is not None
          and not spec.problems() and tool in S.T3LabAIServer._WRITE_TOOLS)
check('room_to_floor: no hand-written registry entry left',
      'room_to_floor' not in server_registry.literal())
check('room_to_floor: old argument names still accepted',
      set(['room_ids', 'room_id', 'floor_type']) <= set(TS.get('room_to_floor').properties))
check('unjoin is destructive, join and its dry run are not',
      srv.is_destructive('auto_join', {'mode': 'unjoin'})
      and not srv.is_destructive('auto_join', {'mode': 'unjoin', 'dry_run': True})
      and not srv.is_destructive('auto_join', {}))
check('id arrays accept "@set" handles', all(
    TS.get(t).properties[k]['items']['type'] == ['integer', 'string']
    for t, k in (('auto_join', 'element_ids'), ('room_to_floor', 'room_ids'),
                 ('rehost_to_level', 'element_ids'), ('cut_wall_openings', 'element_ids'),
                 ('create_door_thresholds', 'element_ids'))))
for tool in NEW:
    size = len('{}'.format(TS.get(tool).registry_entry())) // 4
    check('{}: registry entry ~{} tokens (< 350)'.format(tool, size), size < 350)

print('[dialogs use the shared code]')
for dialog, module, gone in (('RoomToFloorDialog.py', 'Snippets.finish_floors', 'FloorGenerator'),
                             ('DoorThresholdDialog.py', 'Snippets.finish_floors',
                              'ThresholdGenerator'),
                             ('WallAdjustBaseDialog.py', 'Snippets.wall_tools', None),
                             ('WallCutProfileDialog.py', 'Snippets.wall_tools', None)):
    src = io.open(os.path.join(LIB, 'GUI', dialog), encoding='utf-8').read()
    tree = ast.parse(src)
    imports = [n.module for n in tree.body if isinstance(n, ast.ImportFrom)]
    classes = [n.name for n in tree.body if isinstance(n, ast.ClassDef)]
    check('{} imports {}'.format(dialog, module), module in imports
          and (gone is None or gone not in classes))

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
