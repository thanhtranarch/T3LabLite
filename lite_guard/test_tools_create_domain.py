# -*- coding: utf-8 -*-
"""core/tools/from_source.py + core/tools/dimensions.py: build_from_point_cloud,
cad_to_elements and auto_dimension, plus the Revit-side modules the windows now
share with them (Services/point_cloud_build, Snippets/_cad_plan, _auto_dimension).

Run:  python lite_guard/test_tools_create_domain.py      (fake Revit API, no Revit)
"""
import ast
import io
import math
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
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_create_domain_test_')

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


# ── fake Revit: TransactionGroup + ids ──────────────────────────────────────

class Group(object):
    log = []

    def __init__(self, doc, name):
        Group.log.append(name)

    def Start(self):
        Group.log.append('start')

    def Assimilate(self):
        Group.log.append('assimilate')

    def RollBack(self):
        Group.log.append('rollback')

    def HasStarted(self):
        return True

    def HasEnded(self):
        return Group.log[-1] in ('assimilate', 'rollback')

    def Dispose(self):
        pass


class Id(object):
    def __init__(self, v):
        self.Value = self.IntegerValue = v


db = types.ModuleType('Autodesk.Revit.DB')
db.TransactionGroup = Group
sys.modules.setdefault('Autodesk', types.ModuleType('Autodesk'))
sys.modules.setdefault('Autodesk.Revit', types.ModuleType('Autodesk.Revit'))
sys.modules['Autodesk.Revit.DB'] = db
import Snippets._compat as C   # noqa: E402
C.make_eid = lambda v: v

from core import tool_specs as TS                       # noqa: E402
from core.tools import from_source as FS, dimensions as DM   # noqa: E402
from Services import point_cloud_analysis as PCA        # noqa: E402
from Services import point_cloud_build as PCB           # noqa: E402
from Snippets import _cad_geometry as geo               # noqa: E402
from Snippets import _cad_plan as CP                    # noqa: E402

MM = 1.0 / 304.8

# ── point cloud: detections in, the ElementBuilder out ──────────────────────
print('[point cloud]')
wall = PCA.DetectedWall(math.radians(30), 10.0, 5.0, 4000 * MM, 200 * MM, 'L1', True, '', 0.0)
other = PCA.DetectedWall(0.0, 40.0, 5.0, 3000 * MM, 150 * MM, 'L1', False, '', 0.0)
door = PCA.DetectedOpening('Door', wall._data, 700 * MM, 0.0, 900 * MM, 2100 * MM, 'L1')
floor = PCA.DetectedFloor(0.0, [(0, 0), (10, 0), (10, 8), (0, 8)], 'floor', 'L1')
col = PCA.DetectedColumn(3.0, 4.0, 400 * MM, 400 * MM, 0.0, 10.0, 'L1')
stair = PCA.DetectedStair(0.0, 10.0, 18, 'L1')
roof = PCA.DetectedRoof(30.0, [(0, 0), (10, 0), (10, 8), (0, 8)], 12.0, 'L2')
detected = [wall, other, door, floor, col, stair, roof]
summary = PCA.summarize_detections(detected)
els, problems = PCB.from_detections(summary)
by = dict((e.Type, e) for e in els)
check('from_detections: every row back, no problem', len(els) == 7 and not problems,
      (len(els), problems))
w = [e for e in els if e.Type == 'Wall' and e.ConfidenceText == '85%'][0]
check('wall geometry round-trips (mm precision)',
      abs(w._data['cx'] - 10.0) < 1e-3 and abs(w._data['length_ft'] - 4000 * MM) < 1e-3
      and abs(math.degrees(w._data['angle']) % 180 - 30) < 0.01, w._data)
d = by['Door']
check('door hosted in ITS wall row (same dict -> build_all hosts it)',
      d._data['host_wall_data'] is w._data, d._data)
cx = w._data['cx'] + math.cos(w._data['angle']) * d._data['u_center']
check('door insertion point unchanged',
      abs(cx - (10.0 + math.cos(math.radians(30)) * 700 * MM)) < 1e-3, cx)
check('floor / column / roof / stair data', by['Floor']._data['corners_xy'][2] == (10.0, 8.0)
      and abs(by['Column']._data['width_ft'] - 400 * MM) < 1e-6
      and by['Roof'].LevelName == 'L2' and by['Stair']._data['tread_count'] == 18)
els2, problems = PCB.from_detections({'walls': [{'type': 'Wall', 'start_mm': [0, 0]}],
                                      'doors': [{'type': 'Door'}]})
check('bad rows reported, not raised', not els2 and len(problems) == 2, problems)

server_like = {}


class FakeBuilder(object):
    def __init__(self, doc):
        self._levels = {'L1': 'level'}

    def build_all(self, elements, height, created=None):
        server_like['height'] = height
        server_like['types'] = [e.Type for e in elements]
        counts = {}
        for i, e in enumerate(elements):
            if e.Type == 'Stair':
                continue
            counts[e.Type] = counts.get(e.Type, 0) + 1
            created.append(500 + i)
        return {'counts': counts, 'errors': 1, 'skipped': {'Stair': 1}}


PCB.ElementBuilder = FakeBuilder
r = TS.run('build_from_point_cloud', object(), {'detections': summary, 'dry_run': True})
check('dry run: counts per type, stairs and unknown levels named',
      r['dry_run'] and r['planned']['Wall'] == 2 and r['planned']['Stair'] == 1
      and 'stair' in r['note'] and r['levels_not_in_model'] == ['L2'] and 'types' not in server_like,
      r)
r = TS.run('build_from_point_cloud', object(),
           {'detections': summary['detections'], 'elements': ['walls', 'doors'],
            'min_confidence': 75, 'wall_height_mm': 2800})
check('kinds + min_confidence filter, wall height in feet',
      server_like['types'] == ['Wall'] and r['below_min_confidence'] == 2
      and abs(server_like['height'] - 2800 * MM) < 1e-9, (server_like, r))
check('created ids + failed hint', r['created_ids'] == [500] and r['failed'] == 1
      and 'Load one' in r['hint'] and r['success'], r)
r = TS.run('build_from_point_cloud', object(), {'detections': summary, 'elements': ['pipes']})
check('unknown kind -> accepted list', 'pipes' in r['error'] and 'walls' in r['accepted'], r)
PCB.detect = lambda doc, args: (None, {'error': 'There is no point cloud in the active document.'})
r = TS.run('build_from_point_cloud', object(), {})
check('no detections -> detect again; its error comes back', 'no point cloud' in r['error'], r)
PCB.detect = lambda doc, args: ([], {'point_cloud_id': 9})
r = TS.run('build_from_point_cloud', object(), {})
check('nothing detected -> what next', 'Nothing to build' in r['error'], r)

# pick_cloud / region_from_args stay in step with the server's detect tool.
import core.server as S   # noqa: E402
srv = object.__new__(S.T3LabAIServer)


class Cloud(object):
    def __init__(self, i, name):
        self.Id, self.Name = Id(i), name

    def get_BoundingBox(self, view):
        return types.SimpleNamespace(Min=types.SimpleNamespace(X=0.0, Y=1.0, Z=2.0),
                                     Max=types.SimpleNamespace(X=50.0, Y=40.0, Z=30.0))


clouds = [Cloud(7, 'A'), Cloud(8, 'B')]
idof = lambda c: c.Id.Value     # noqa: E731
HAS_PC_TOOLS = hasattr(srv, '_pick_point_cloud')      # Lite's server has no point cloud tools
for wanted in (None, 8, 99) if HAS_PC_TOOLS else ():
    mine = PCB.pick_cloud(clouds, wanted, idof)
    theirs = srv._pick_point_cloud(clouds, wanted, lambda i: i.Value)
    check('pick_cloud == server ({})'.format(wanted), (mine[0] is theirs[0])
          and bool(mine[1]) == bool(theirs[1]))
levels = [('L1', 0.0), ('L2', 12.0)]
for args in (({}, {'level_name': 'l2'}, {'level_name': 'L1'}, {'level_name': 'X'},
              {'region_mm': {'min': [0, 0, 0], 'max': [3048, -3048, 304.8]}}, {'region_mm': 3})
             if HAS_PC_TOOLS else ()):
    mine = PCB.region_from_args(args, levels, ((0.0, 1.0, 2.0), (50.0, 40.0, 30.0)))
    theirs, err = srv._point_cloud_region(clouds[0], args, levels)
    same = (isinstance(mine, dict) and err is not None) or mine == theirs
    check('region_from_args == server {}'.format(sorted(args)), same, (mine, theirs, err))

# ── CAD to Elements ─────────────────────────────────────────────────────────
print('[cad]')


class Layer(object):
    def __init__(self, name, segments=(), loops=()):
        self.name = name
        self.segments = [tuple(s) + (name,) for s in segments]
        self._loops = list(loops)
        self.closed_polys, self.circles, self.arc_edges, self.spline_edges = [], [], [], []
        self.n_curves = len(self.segments) + len(self._loops)

    def loops(self):
        return self._loops

    def footprints(self):
        return []

    def all_edges(self):
        return [('L', (s[0], s[1]), (s[2], s[3]), None) for s in self.segments]


T = 200 * MM
walls_a = Layer('A-WALL', [(0, 0, 20, 0), (0, 20, 20, 20)])
walls_b = Layer('A-WALL-2', [(0, T, 20, T), (0, 20 + T, 20, 20 + T)])
slab = Layer('S-SLAB', loops=[geo.loop_from_polygon([(0, 0), (20, 0), (20, 20), (0, 20)])])
axes = Layer('GRID', [(0, -5, 0, 25), (10, -5, 10, 25), (-5, 0, 25, 0)])
LAYERS = dict((lg.name, lg) for lg in (walls_a, walls_b, slab, axes, Layer('EMPTY')))


class Run(object):
    def __init__(self, n):
        self.created, self.failed, self.skipped, self.warnings = n, 0, 0, 2
        self.notes, self.errors = [], []


CALLS = []
cr = types.ModuleType('Snippets._cad_revit')
cr.DB = types.SimpleNamespace(BuiltInCategory=types.SimpleNamespace(
    OST_StructuralFraming='frame', OST_StructuralColumns='scol', OST_Columns='col'),
    ViewType=types.SimpleNamespace(FloorPlan=1, CeilingPlan=2, EngineeringPlan=3, AreaPlan=4,
                                   DraftingView=5))
LV1 = types.SimpleNamespace(Elevation=0.0, Name='Level 1')
cr.get_cad_instances = lambda doc: [{'element': 'cad', 'id': 77, 'name': 'plan.dwg [Linked]'}]
cr.scan_cad = lambda doc, el: LAYERS
cr.get_levels = lambda doc: [{'element': LV1, 'id': 1, 'name': 'Level 1', 'elevation': 0.0},
                             {'element': 'lv2', 'id': 2, 'name': 'Level 2', 'elevation': 10.0}]
cr.get_basic_wall_types = lambda doc: [{'element': 'wt1', 'id': 11, 'name': 'Basic 200'},
                                       {'element': 'wt2', 'id': 12, 'name': 'Generic - 200mm'}]
cr.get_floor_types = lambda doc: [{'element': 'ft', 'id': 21, 'name': 'Floor : Slab 200'}]
cr.get_grid_types = lambda doc: [{'element': 'gt', 'id': 31, 'name': '6.5mm Bubble'}]
cr.existing_grid_lines = lambda doc: ([(0, -5, 0, 25)], {'1'})
cr.get_ds_categories = lambda: [{'name': 'Walls', 'bic': 'w'}, {'name': 'Floors', 'bic': 'f'}]


def _create(kind):
    def go(doc, *args, **kw):
        CALLS.append((kind, args))
        return Run(len(args[0]))
    return go


cr.create_walls = _create('walls')
cr.create_floors = _create('floors')
cr.create_grids = _create('grids')
cr.create_parts = _create('parts')
sys.modules['Snippets._cad_revit'] = cr
import Snippets   # noqa: E402
Snippets._cad_revit = cr

items = cr.get_basic_wall_types(None)
check('find_named: no name prefers "generic"', CP.find_named(items, None, 'x', ('generic',))['id'] == 12)
check('find_named: case / spaces / substring', CP.find_named(items, ' basic  200', 'x')['id'] == 11
      and CP.find_named(cr.get_floor_types(None), 'slab 200', 'x')['id'] == 21)
try:
    CP.find_named(items, '200', 'wall type')
    check('find_named: ambiguous raises', False)
except CP.PlanError as ex:
    check('find_named: ambiguous names both', 'Basic 200' in str(ex) and 'Generic' in str(ex))
groups = CP.group_mapping([{'layers': ['A-WALL'], 'element': 'wall'},
                           {'layer': 'A-WALL-2', 'element': 'wall'},
                           {'layers': ['S-SLAB'], 'element': 'floor'}])
check('group_mapping: same element/type/level/options planned together',
      [g[1] for g in groups] == [['A-WALL', 'A-WALL-2'], ['S-SLAB']])
plan = CP.plan_wall(None, [walls_a, walls_b], LV1, 'Level 1', 3000 * MM, 0.0, T, False, True,
                    wall_type=items[0])
check('plan_wall: faces on two layers pair into 2 walls', plan.count == 2
      and plan.question == 'Create 2 walls on Level 1?' and 'Thicknesses: 200 mm' in plan.details,
      (plan.count, plan.details))
one = CP.Plan(1, 'grids', 'Create 1 grids?', 'Create 1 Grids', '', None)
check('Plan singular text (moved from the dialog)', one.question == 'Create 1 grid?'
      and one.ok_text == 'Create 1 Grid')

r = TS.run('cad_to_elements', object(), {})
check('scan: layers with geometry, levels, elements', r['cad']['id'] == 77
      and [x['layer'] for x in r['layers']] == ['A-WALL', 'A-WALL-2', 'GRID', 'S-SLAB']
      and r['empty_layers'] == 1 and 'Level 2' in r['levels'] and 'duct' in r['elements'], r)
MAP = [{'layers': ['A-WALL', 'A-WALL-2'], 'element': 'wall', 'level': 'level 1'},
       {'layers': ['S-SLAB'], 'element': 'floor', 'options': {'offset_mm': 50}},
       {'layers': ['GRID'], 'element': 'grid'}]


class CadDoc(object):
    ActiveView = types.SimpleNamespace(GenLevel=types.SimpleNamespace(Name='Level 1'))


r = TS.run('cad_to_elements', CadDoc(), {'operation': 'create', 'mapping': MAP, 'dry_run': True})
check('dry run: count per entry, nothing created',
      [(x['element'], x['count']) for x in r['planned']] == [('wall', 2), ('floor', 1),
                                                              ('grid', 2)]
      and r['would_create'] == 5 and not CALLS, r)
check('grid dry run skips the axis already in the model',
      'already in the model (skipped): 1' in r['planned'][2]['details'][0], r['planned'][2])
r = TS.run('cad_to_elements', CadDoc(), {'operation': 'create', 'mapping': MAP})
check('create: one TransactionGroup, assimilated, results per entry',
      Group.log[-3:] == ['T3Lab: CAD to Elements', 'start', 'assimilate'] and r['created_count'] == 5
      and [x['status'] for x in r['results']] == ['ok'] * 3, (Group.log, r))
check('floor offset in feet reached create_floors', abs(CALLS[1][1][3] - 50 * MM) < 1e-9,
      CALLS[1])
r = TS.run('cad_to_elements', CadDoc(), {'operation': 'create', 'mapping': [
    {'layers': ['A-WALL'], 'element': 'wall', 'options': {'height': 3}}]})
check('unknown option -> options with defaults', 'height_mm=3000' in r['error'], r)
r = TS.run('cad_to_elements', CadDoc(), {'operation': 'create', 'mapping': [
    {'layers': ['NOPE'], 'element': 'wall'}]})
check('unknown layer -> the layers', 'NOPE' in r['error'] and 'A-WALL' in r['layers'], r)
r = TS.run('cad_to_elements', CadDoc(), {'operation': 'create', 'mapping': [
    {'layers': ['EMPTY'], 'element': 'floor'}]})
check('nothing found -> error, no transaction', 'Nothing to create' in r['error']
      and r['planned'][0]['empty'], r)
cr.create_walls = lambda *a: (_ for _ in ()).throw(RuntimeError('boom'))
r = TS.run('cad_to_elements', CadDoc(), {'operation': 'create', 'mapping': MAP[:1]})
check('a failing create -> rollback, error', Group.log[-1] == 'rollback' and 'boom' in
      r['results'][0]['error'] and 'unchanged' in r['error'], r)

# ── auto_dimension ──────────────────────────────────────────────────────────
print('[auto_dimension]')


class View(object):
    def __init__(self, i, name, vt='FloorPlan', template=False):
        self.Id, self.Name, self.ViewType, self.IsTemplate = Id(i), name, vt, template


class Dim(object):
    def __init__(self, i, segs):
        self.Id = Id(i)
        self.Segments = [types.SimpleNamespace(Value=v) for v in segs]


VIEWS = {1: View(1, 'L1 Plan'), 2: View(2, 'Section A', 'Section'), 3: View(3, '3D', 'ThreeD'),
         4: View(4, 'Tpl', template=True)}
DIMCALLS = []
ad = types.ModuleType('Snippets._auto_dimension')
ad.TARGET_CATEGORIES = (('walls', 'OST_Walls'), ('structural_columns', 'OST_StructuralColumns'),
                        ('columns', 'OST_Columns'), ('windows', 'OST_Windows'),
                        ('doors', 'OST_Doors'), ('lifts', 'OST_MechanicalEquipment'))
ad._is_valid_view = lambda v: v.ViewType in ('FloorPlan', 'Section')
ad.linear_dimension_types = lambda doc: [('Linear 2.5mm', 'dt1'), ('Diagonal', 'dt2')]
ad.count_in_view = lambda doc, v, bic: {'OST_Grids': 4 if v.Id.Value == 1 else 0}.get(bic, 3)


def fake_dim(doc, view, *args, **kw):
    DIMCALLS.append((view.Name, args))
    if view.Id.Value == 2:
        kw['on_error']("Transaction failed for view 'Section A': x")
        return (0, 1, [], [])
    return (2, 3, ['ref rejected'], [Dim(901, [0.5, 5.0]), Dim(902, [4.0])])


ad.dim_one_view = fake_dim
sys.modules['Snippets._auto_dimension'] = ad
Snippets._auto_dimension = ad


class DimDoc(object):
    ActiveView = VIEWS[1]

    def GetElement(self, i):
        return VIEWS.get(i)


opts, err = DM.read_options({})
check('defaults = the window: all but facade, L1/L2/L3 500/1000/1500, both sides, 300 min',
      err is None and opts['do_walls'] and opts['do_lifts'] and not opts['do_facade']
      and abs(opts['l1_feet'] - 500 * MM) < 1e-9 and abs(opts['l3_feet'] - 1500 * MM) < 1e-9
      and opts['both_sides'] and opts['check_min'] and opts['run_x'] and opts['run_y'], opts)
opts, err = DM.read_options({'offsets_mm': [400, 800, 1200], 'direction': 'x',
                             'check_mode': True, 'targets': ['grids']})
check('three offsets, X only, check mode turns both sides off', opts['is_3level']
      and not opts['run_y'] and not opts['both_sides'] and not opts['do_walls'])
check('bad options -> message', DM.read_options({'offsets_mm': [1, 2]})[1]
      and DM.read_options({'targets': ['rebar']})[1])
r = TS.run('auto_dimension', DimDoc(), {'view_ids': [1, 2, 3, 4], 'dry_run': True})
check('dry run: per-view counts, bad views skipped, nothing made',
      r['dry_run'] and [v['view'] for v in r['views']] == ['L1 Plan', 'Section A']
      and r['views'][0]['grids'] == 4 and r['views'][0]['walls'] == 3
      and r['views'][1]['measures'] == 'grids and levels'
      and len(r['skipped_views']) == 2 and not DIMCALLS, r)
r = TS.run('auto_dimension', DimDoc(), {'dim_type': 'nope'})
check('unknown dim type -> names', 'Linear 2.5mm' in r['error'], r)
r = TS.run('auto_dimension', DimDoc(), {'view_ids': [1, 2], 'dim_type': 'diagonal'})
args = DIMCALLS[0][1]
check('run: same positional args the window passes (wall faces, dim type, L1..L3)',
      args[8] == 'ext' and args[9] == 'dt2' and abs(args[14] - 1000 * MM) < 1e-9
      and len(args) == 24, args)
check('one TransactionGroup assimilated; ids, errors and short segments reported',
      Group.log[-3:] == ['T3Lab: Auto Dimension', 'start', 'assimilate']
      and r['dimension_ids'] == [901, 902] and 'Section A' in r['errors'][0]
      and r['warning'].startswith('1 segment'), r)
r = TS.run('auto_dimension', DimDoc(), {'view_ids': [2]})
check('nothing made -> rollback + what next', Group.log[-1] == 'rollback'
      and 'two grids' in r['error'], r)
r = TS.run('auto_dimension', DimDoc(), {'view_ids': [3]})
check('no valid view -> error', 'No view to dimension' in r['error'], r)

# ── registration, contract, windows still delegate ──────────────────────────
print('[contract]')
for name in ('build_from_point_cloud', 'cad_to_elements', 'auto_dimension'):
    spec = TS.get(name)
    size = len(__import__('json').dumps(spec.registry_entry())) // 4
    check('{}: complete spec, main thread, not destructive, {} tokens <= 350'.format(name, size),
          not spec.problems() and name in S.T3LabAIServer._WRITE_TOOLS and size <= 350
          and not srv.is_destructive(name, {}))


def src(rel):
    return io.open(os.path.join(LIB, rel), encoding='utf-8').read()


pc = src('GUI/PointCloudDialog.py')
check('PointCloudDialog imports ElementBuilder, defines no builder',
      'from Services.point_cloud_build import ElementBuilder' in pc
      and 'class ElementBuilder' not in pc and 'class WarningSwallower' not in pc)
cad = src('GUI/CADToElementsDialog.py')
check('CADToElementsDialog plans through Snippets/_cad_plan',
      all('cp.plan_{}('.format(m) in cad for m in ('wall', 'floor', 'ceiling', 'room', 'column',
                                                     'beam', 'grid', 'lines', 'mep'))
      and 'cp.PlanError' in cad and 'class Plan(' not in cad)
adlg = src('GUI/AutoDimensionDialog.py')
check('AutoDimensionDialog delegates to Snippets/_auto_dimension',
      'return dim_one_view(self.doc, view,' in adlg and 'FilteredElementCollector(self.doc, view.Id)'
      not in adlg)

print('[ironpython]')
for rel in ('core/tools/from_source.py', 'core/tools/dimensions.py', 'Snippets/_cad_plan.py',
            'Snippets/_auto_dimension.py', 'Services/point_cloud_build.py'):
    tree = ast.parse(src(rel))
    # f(a, *rest) is 2.7; a, *rest = x and [*a] are not.
    bad = [type(n).__name__ for n in ast.walk(tree)
           if isinstance(n, (ast.JoinedStr, ast.Nonlocal, ast.NamedExpr))
           or (isinstance(n, ast.Starred) and not isinstance(n.ctx, ast.Load))
           or (isinstance(n, (ast.List, ast.Tuple, ast.Set))
               and any(isinstance(e, ast.Starred) for e in n.elts))
           or (isinstance(n, ast.arguments) and (n.kwonlyargs or n.posonlyargs))]
    check('{}: no py3-only syntax'.format(rel), not bad, bad)

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
