# -*- coding: utf-8 -*-
"""T3Lab Data Hub: snapshot export (core/hub_snapshot.py, against a fake Revit
API), the SQLite store (core/hub_store.py) and the t3lab_hub gateway tool
(Lite core/bridge.py). Runs on Linux, no Revit. docs/data-hub.md."""
import ast
import gzip
import json
import os
import subprocess
import sys
import tempfile
import types
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_hub_appdata_')
os.environ['T3LAB_HUB_DIR'] = tempfile.mkdtemp(prefix='t3lab_hub_')

from core import hub_snapshot as snap  # noqa: E402
from core import hub_store as store    # noqa: E402
import core.bridge as B                # noqa: E402


# ─── Fake Revit API (only what hub_snapshot reads) ─────────────────────────────

class Eid(object):
    def __init__(self, value):
        self.Value = value


class Unit(object):
    def __init__(self, name):
        self.TypeId = 'autodesk.unit.unit:{}-1.0.1'.format(name)


class Param(object):
    def __init__(self, name, storage, value, text=None, unit=None):
        self.Definition = types.SimpleNamespace(Name=name)
        self.StorageType, self._value, self._text, self._unit = storage, value, text, unit
        self.HasValue = value is not None

    def AsString(self):
        return self._value

    def AsInteger(self):
        return self._value

    def AsDouble(self):
        return self._value

    def AsElementId(self):
        return self._value

    def AsValueString(self):
        return self._text

    def GetUnitTypeId(self):
        if self._unit is None:
            raise RuntimeError('not measurable')
        return Unit(self._unit)


class Elem(object):
    ViewSpecific = False

    def __init__(self, doc, eid, category, cat_type='Model', name='', type_id=-1,
                 level_id=-1, params=(), family=None, cat_id=1):
        self.doc, self.Id, self.UniqueId = doc, Eid(eid), 'uid-{}'.format(eid)
        self.Category = (types.SimpleNamespace(Name=category, CategoryType=cat_type,
                                               Id=Eid(cat_id)) if category else None)
        self.Name, self.FamilyName = name, family
        self._type, self.LevelId = Eid(type_id), Eid(level_id)
        self.CreatedPhaseId, self.WorksetId = Eid(-1), None
        self._params = list(params)

    def GetTypeId(self):
        return self._type

    def GetOrderedParameters(self):
        return self._params

    def GetParameters(self, name):
        return [p for p in self._params if p.Definition.Name == name]

    def get_Parameter(self, _bip):
        return None


class Warning_(object):
    def __init__(self, text, ids):
        self._text, self._ids = text, ids

    def GetDescriptionText(self):
        return self._text

    def GetSeverity(self):
        return 'Warning'

    def GetFailingElements(self):
        return [Eid(i) for i in self._ids]


class Doc(object):
    Title, PathName = 'Tower A.rvt', 'C:\\Projects\\Tower A.rvt'
    IsWorkshared = IsModelInCloud = IsModified = False

    def __init__(self):
        self.instances, self.types, self.warnings = [], [], []
        self.Application = types.SimpleNamespace(VersionNumber='2025', VersionBuild='25.0')

    def GetElement(self, eid):
        for e in self.instances + self.types:
            if e.Id.Value == eid.Value:
                return e
        return None

    def GetWarnings(self):
        return self.warnings


class Collector(object):
    def __init__(self, doc):
        self.doc = doc

    def WhereElementIsNotElementType(self):
        return list(self.doc.instances)

    def WhereElementIsElementType(self):
        return list(self.doc.types)


DB = types.ModuleType('Autodesk.Revit.DB')
DB.StorageType = types.SimpleNamespace(String='S', Integer='I', Double='D', ElementId='E')
DB.CategoryType = types.SimpleNamespace(Model='Model')
DB.BuiltInCategory = types.SimpleNamespace(OST_Sheets=-2003100, OST_Levels=-2000240,
                                           OST_Grids=-2000220, OST_ProjectInformation=-2003101)
DB.BuiltInParameter = types.SimpleNamespace()
DB.ElementId = types.SimpleNamespace(InvalidElementId=Eid(-1))
DB.FilteredElementCollector = Collector
DB.UnitUtils = types.SimpleNamespace(
    ConvertFromInternalUnits=lambda raw, unit: raw * 304.8 if 'millimeters' in unit.TypeId else raw)
for name in ('Autodesk', 'Autodesk.Revit'):
    sys.modules.setdefault(name, types.ModuleType(name))
sys.modules['Autodesk.Revit'].DB = DB
sys.modules['Autodesk.Revit.DB'] = DB


def make_doc(height_ft=10.0, mark='W1', extra_wall=False, warnings=1):
    doc = Doc()
    doc.types = [Elem(doc, 50, 'Walls', name='Generic - 200mm', family='Basic Wall',
                      params=[Param('Fire Rating', 'S', '2h')])]
    level = Elem(doc, 30, 'Levels', cat_type='Annotation', name='Level 1', cat_id=-2000240)
    wall = Elem(doc, 100, 'Walls', name='Generic - 200mm', type_id=50, level_id=30, params=[
        Param('Unconnected Height', 'D', height_ft, unit='millimeters'),
        Param('Mark', 'S', mark), Param('Comments', 'S', None),
        Param('Structural', 'I', 1, text='Yes'), Param('Count', 'I', 3, text='3'),
        Param('Phase Created', 'E', Eid(7), text='New Construction'),
        Param('Ratio', 'D', 0.5)])
    door = Elem(doc, 101, 'Doors', name='D1', level_id=30,
                params=[Param('Mark', 'S', 'D-01')])
    tag = Elem(doc, 200, 'Wall Tags', cat_type='Annotation', name='tag')
    sketch = Elem(doc, 201, '<Sketch>', name='line')
    nocat = Elem(doc, 202, None)
    doc.instances = [level, wall, door, tag, sketch, nocat]
    if extra_wall:
        doc.instances.append(Elem(doc, 102, 'Walls', name='Generic - 200mm', type_id=50,
                                  level_id=30, params=[Param('Mark', 'S', 'W2')]))
    doc.warnings = [Warning_('Highlighted walls overlap.', [100, 101])] * warnings
    return doc


def read_lines(path):
    opener = gzip.open if path.endswith('.gz') else open
    with opener(path, 'rt', encoding='utf-8') as f:
        return [json.loads(line) for line in f]


class Clock(object):
    def __init__(self, start=1763366400.0, step=0.0):     # 2025-11-17T08:00:00Z
        self.now, self.step = start, step

    def __call__(self):
        self.now += self.step
        return self.now


class SnapshotTests(unittest.TestCase):
    def setUp(self):
        self.inbox = tempfile.mkdtemp(prefix='t3lab_inbox_')

    def test_helpers(self):
        self.assertEqual(snap.model_key('C:\\A.rvt'), snap.model_key(' c:\\a.RVT '))
        self.assertEqual(len(snap.model_key('x')), 12)
        self.assertEqual(snap.snapshot_name('abc', 0), 'abc-19700101T000000Z')
        self.assertEqual(snap.clean_names([' Walls', 'walls', 'Doors']), ['walls', 'doors'])
        self.assertIsNone(snap.clean_names([]))
        self.assertEqual(snap.clamp_seconds(9999), snap.MAX_SECONDS_CAP)
        self.assertEqual(snap.clamp_seconds('x'), snap.DEFAULT_MAX_SECONDS)
        self.assertEqual(snap.int_or_text('Yes', 1), 'Yes')
        self.assertEqual(snap.int_or_text('3', 9), 3)
        self.assertEqual(snap.int_or_text('', 9), 9)

    def test_publish_writes_a_complete_snapshot(self):
        result = snap.publish(make_doc(), clock=Clock(), folder=self.inbox)
        self.assertTrue(result['published'])
        self.assertEqual((result['elements'], result['types'], result['warnings']), (3, 1, 1))
        self.assertEqual(os.listdir(self.inbox), [result['snapshot'] + '.ndjson'])
        lines = read_lines(result['file'])
        header, end = lines[0], lines[-1]
        self.assertEqual((header['t'], header['format'], header['title']),
                         ('header', 'hub/1', 'Tower A.rvt'))
        self.assertEqual(header['identity_kind'], 'file')
        self.assertEqual(header['taken_at'], '2025-11-17T08:00:00Z')
        self.assertTrue(end['complete'])
        self.assertEqual(end['units'], {'Unconnected Height': 'millimeters'})
        by_uid = dict((r['u'], r) for r in lines if r['t'] == 'e')
        self.assertEqual(sorted(by_uid), ['uid-100', 'uid-101', 'uid-30', 'uid-50'])
        wall = by_uid['uid-100']
        self.assertEqual((wall['c'], wall['f'], wall['y'], wall['l'], wall['k']),
                         ('Walls', 'Basic Wall', 'Generic - 200mm', 'Level 1', 'i'))
        self.assertEqual(wall['p'], {
            'Unconnected Height': 3048.0, 'Mark': 'W1', 'Structural': 'Yes', 'Count': 3,
            'Phase Created': 'New Construction', 'Ratio': 0.5})
        self.assertEqual(by_uid['uid-50']['k'], 't')
        self.assertEqual(by_uid['uid-50']['p'], {'Fire Rating': '2h'})
        warning = [r for r in lines if r['t'] == 'w'][0]
        self.assertEqual(warning['u'], ['uid-100', 'uid-101'])

    def test_scope_arguments(self):
        result = snap.publish(make_doc(), {'categories': ['walls', 'Wall Tags'],
                                           'parameters': ['Mark']},
                              clock=Clock(), folder=self.inbox)
        records = [r for r in read_lines(result['file']) if r['t'] == 'e']
        self.assertEqual(sorted(r['c'] for r in records), ['Wall Tags', 'Walls', 'Walls'])
        self.assertEqual([r.get('p') for r in records if r['u'] == 'uid-100'], [{'Mark': 'W1'}])

    def test_timeout_publishes_nothing(self):
        doc = make_doc()
        doc.instances = doc.instances * 100
        result = snap.publish(doc, {'max_seconds': 5}, clock=Clock(step=10.0), folder=self.inbox)
        self.assertEqual(result['error_code'], 'hub_publish_timeout')
        self.assertEqual(os.listdir(self.inbox), [])

    def test_writer_abort_leaves_no_file(self):
        writer = snap.SnapshotWriter(self.inbox, 'k-20250101T000000Z')
        writer.write({'t': 'header'})
        writer.abort()
        self.assertEqual(os.listdir(self.inbox), [])


class StoreCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix='t3lab_hubstore_')
        self.hub = store.Hub(self.root)
        self.clock = Clock()

    def publish(self, doc, arguments=None, hours=1):
        self.clock.now += hours * 3600
        result = snap.publish(doc, arguments, clock=self.clock, folder=self.hub.inbox)
        self.assertTrue(result.get('published'), result)
        return result


class StoreTests(StoreCase):
    def test_ingest_versions_only_what_changed(self):
        self.publish(make_doc())
        first = self.hub.ingest_pending()
        self.assertEqual(first['errors'], [])
        self.assertEqual(first['ingested'][0]['seq'], 1)
        self.assertIsNone(first['ingested'][0]['since_previous'])
        self.assertEqual(self.hub.pending(), [])
        archived = os.listdir(os.path.join(self.hub.archive, first['ingested'][0]['model_key']))
        self.assertEqual(len(archived), 1)
        self.assertTrue(archived[0].endswith('.ndjson.gz'))

        self.publish(make_doc(mark='W1-rev', extra_wall=True))
        second = self.hub.ingest_pending()['ingested'][0]
        self.assertEqual(second['since_previous'],
                         {'new': 1, 'changed': 1, 'unchanged': 3, 'removed': 0})
        import sqlite3
        con = sqlite3.connect(self.hub.db_path)
        rows = con.execute('SELECT COUNT(*) FROM elements').fetchone()[0]
        con.close()
        self.assertEqual(rows, 6)      # 4 + new wall + new version of the changed wall

    def test_summary(self):
        self.publish(make_doc())
        self.hub.ingest_pending()
        models = self.hub.summary(now=self.clock.now + 7200)
        self.assertEqual(models['models'][0]['model'], 'Tower A.rvt')
        self.assertEqual(models['models'][0]['age_hours'], 2.0)
        one = self.hub.summary('tower')
        self.assertEqual((one['elements'], one['types'], one['warnings']), (3, 1, 1))
        self.assertEqual(sorted((c['value'], c['count']) for c in one['by_category']),
                         [('Doors', 1), ('Levels', 1), ('Walls', 1)])
        self.assertEqual(one['top_warnings'][0]['count'], 1)

    def test_query_filters_groups_and_pages(self):
        self.publish(make_doc(extra_wall=True))
        self.hub.ingest_pending()
        walls = self.hub.query(category='walls', select=['Mark'])
        self.assertEqual(walls['total'], 2)
        self.assertEqual([r['params']['Mark'] for r in walls['rows']], ['W1', 'W2'])
        self.assertEqual(walls['rows'][0]['level'], 'Level 1')
        self.assertIn('as_of', walls['snapshot'])
        tall = self.hub.query(where=[{'param': 'Unconnected Height', 'op': '>', 'value': 3000}])
        self.assertEqual([r['uid'] for r in tall['rows']], ['uid-100'])
        none = self.hub.query(category='Walls', where=[{'param': 'Unconnected Height',
                                                        'op': 'empty'}])
        self.assertEqual([r['uid'] for r in none['rows']], ['uid-102'])
        groups = self.hub.query(group_by=['category'], sum_param='Unconnected Height')
        self.assertEqual(groups['groups'][0], {'category': 'Walls', 'count': 2, 'sum': 3048.0})
        self.assertEqual(groups['unit'], 'millimeters')
        by_param = self.hub.query(category='Walls', group_by='param:Structural')
        self.assertEqual(set((g['param:Structural'], g['count']) for g in by_param['groups']),
                         {(None, 1), ('Yes', 1)})
        page = self.hub.query(limit=2)
        self.assertEqual((len(page['rows']), page['next_offset'], page['total']), (2, 2, 4))
        rest = self.hub.query(limit=2, offset=2)
        self.assertEqual((len(rest['rows']), rest['next_offset']), (2, None))
        types_ = self.hub.query(kind='type')
        self.assertEqual([r['family'] for r in types_['rows']], ['Basic Wall'])
        with self.assertRaises(store.HubError) as bad:
            self.hub.query(where=[{'param': 'Mark', 'op': 'like'}])
        self.assertEqual(bad.exception.payload['error_code'], 'bad_arguments')

    def test_match_ops(self):
        self.assertTrue(store.match('3048.0', '=', 3048))
        self.assertTrue(store.match('Yes', '=', 'yes'))
        self.assertTrue(store.match(None, '!=', 'x') is True)
        self.assertTrue(store.match('Fire 2h', 'contains', '2H'))
        self.assertTrue(store.match(None, 'not_contains', 'x'))
        self.assertFalse(store.match('abc', '>', 1))
        self.assertFalse(store.match(True, '>', 0))

    def test_diff(self):
        self.publish(make_doc(warnings=1))
        doc = make_doc(height_ft=12.0, mark='W1', extra_wall=True, warnings=0)
        doc.instances = [e for e in doc.instances if e.Id.Value != 101]      # door deleted
        self.publish(doc)
        self.hub.ingest_pending()
        d = self.hub.diff()
        self.assertEqual((d['added'], d['removed'], d['changed']), (1, 1, 1))
        self.assertEqual((d['since']['seq'], d['to']['seq']), (1, 2))
        change = d['changed_samples'][0]
        self.assertEqual(change['uid'], 'uid-100')
        self.assertEqual(change['changes'], {'Unconnected Height': [3048.0, 3657.6]})
        self.assertEqual(d['changed_parameters'], [{'name': 'Unconnected Height', 'elements': 1}])
        self.assertEqual(d['added_samples'][0]['uid'], 'uid-102')
        self.assertEqual(d['removed_samples'][0]['category'], 'Doors')
        self.assertEqual((d['warnings']['resolved'], d['warnings']['new']), (1, 0))
        walls = dict((c['category'], c) for c in d['by_category'])['Walls']
        self.assertEqual((walls['added'], walls['changed'], walls['removed']), (1, 1, 0))
        only_doors = self.hub.diff(category='Doors')
        self.assertEqual((only_doors['added'], only_doors['removed'], only_doors['changed']),
                         (0, 1, 0))
        with self.assertRaises(store.HubError):
            self.hub.diff(since=2, to=2)

    def test_diff_compares_only_the_common_scope(self):
        self.publish(make_doc())
        self.publish(make_doc(height_ft=11.0, mark='W9'),
                     {'categories': ['Walls'], 'parameters': ['Mark']})
        self.hub.ingest_pending()
        d = self.hub.diff()
        self.assertEqual((d['added'], d['removed'], d['changed']), (0, 0, 1))
        self.assertEqual(d['changed_samples'][0]['changes'], {'Mark': ['W1', 'W9']})
        self.assertIn('different scopes', d['notes'][0])

    def test_one_snapshot_cannot_diff(self):
        self.publish(make_doc())
        self.hub.ingest_pending()
        with self.assertRaises(store.HubError) as one:
            self.hub.diff()
        self.assertEqual(one.exception.payload['error_code'], 'hub_one_snapshot')

    def test_empty_hub_and_models(self):
        with self.assertRaises(store.HubError) as empty:
            self.hub.query()
        self.assertEqual(empty.exception.payload['error_code'], 'hub_empty')
        self.publish(make_doc())
        other = make_doc()
        other.Title, other.PathName = 'Podium.rvt', 'C:\\Projects\\Podium.rvt'
        self.publish(other)
        self.hub.ingest_pending()
        with self.assertRaises(store.HubError) as many:
            self.hub.query()
        self.assertEqual(many.exception.payload['error_code'], 'hub_model_required')
        self.assertEqual(self.hub.query(model='podium')['model'], 'Podium.rvt')
        allowed = lambda title: title.startswith('Tower')        # noqa: E731
        self.assertEqual(self.hub.query(allowed=allowed)['model'], 'Tower A.rvt')
        with self.assertRaises(store.HubError):
            self.hub.query(model='podium', allowed=allowed)
        self.assertEqual(len(self.hub.summary(allowed=allowed)['models']), 1)

    def test_bad_and_older_files_leave_the_inbox(self):
        os.makedirs(self.hub.inbox)
        with open(os.path.join(self.hub.inbox, 'k-20250101T000000Z.ndjson'), 'w') as f:
            f.write(json.dumps({'t': 'header', 'format': 'hub/1', 'model_key': 'k'}) + '\n')
        out = self.hub.ingest_pending()
        self.assertEqual(out['errors'][0]['error_code'], 'hub_bad_snapshot')
        self.assertEqual(self.hub.pending(), [])
        self.assertEqual(os.listdir(os.path.join(self.root, 'rejected')),
                         ['k-20250101T000000Z.ndjson'])

        newer = self.publish(make_doc(mark='new'), hours=10)
        self.hub.ingest_pending()
        self.clock.now -= 20 * 3600
        older = self.publish(make_doc(mark='old'), hours=0)
        out = self.hub.ingest_pending()
        self.assertEqual(out['errors'][0]['error_code'], 'hub_out_of_order')
        self.assertEqual(self.hub.pending(), [])
        rebuilt = self.hub.rebuild()
        self.assertEqual(rebuilt['ingested'], [older['snapshot'], newer['snapshot']])
        self.assertEqual(rebuilt['errors'], [])
        d = self.hub.diff()
        self.assertEqual(d['changed_samples'][0]['changes'], {'Mark': ['old', 'new']})

    def test_inbox_file_of_an_ingested_snapshot_is_archived_not_lost(self):
        result = self.publish(make_doc())
        with open(result['file'], 'rb') as f:
            data = f.read()
        self.hub.ingest_pending()
        key = result['model_key']
        gz = os.path.join(self.hub.archive, key, result['snapshot'] + '.ndjson.gz')
        os.remove(gz)                               # bridge died between commit and archive
        with open(result['file'], 'wb') as f:
            f.write(data)
        again = self.hub.ingest_pending()
        self.assertTrue(again['ingested'][0]['already_ingested'])
        self.assertTrue(os.path.exists(gz))
        self.assertEqual(self.hub.summary()['models'][0]['snapshots'], 1)

    def test_many_elements(self):
        doc = make_doc()
        wall = doc.instances[1]
        doc.instances = [Elem(doc, 1000 + i, 'Walls', name='W', type_id=50, level_id=-1,
                              params=[Param('Mark', 'S', 'M{}'.format(i % 50)),
                                      Param('Area', 'D', float(i % 7))])
                         for i in range(20000)] + [wall]
        self.publish(doc)
        doc.instances[10]._params[0] = Param('Mark', 'S', 'changed')
        self.publish(doc)
        out = self.hub.ingest_pending()
        self.assertEqual(out['ingested'][1]['since_previous'],
                         {'new': 0, 'changed': 1, 'unchanged': 20001, 'removed': 0})
        groups = self.hub.query(category='Walls', group_by=['param:Mark'], sum_param='Area')
        self.assertEqual(groups['total'], 20001)
        self.assertEqual(groups['groups_total'], 52)
        self.assertEqual(self.hub.diff()['changed'], 1)


class BridgeTests(StoreCase):
    def setUp(self):
        StoreCase.setUp(self)
        os.environ['T3LAB_HUB_DIR'] = self.root
        self.current = {'port': 1}

    def call(self, arguments):
        request = {'jsonrpc': '2.0', 'id': 1, 'method': 'tools/call',
                   'params': {'name': 't3lab_hub', 'arguments': arguments}}
        return B._tool_payload(B._handle(request, self.current, 't'))

    def test_listed_with_the_bridge_tools(self):
        def no_revit(*a, **k):
            raise ConnectionRefusedError('no Revit in this test')
        real = B._forward
        B._forward = no_revit
        self.addCleanup(setattr, B, '_forward', real)
        tools = B._handle_tools_list({'id': 1}, self.current, 't')['result']['tools']
        self.assertIn('t3lab_hub', [t['name'] for t in tools])

    def test_summary_query_diff_through_the_bridge(self):
        self.publish(make_doc())
        self.publish(make_doc(mark='W1b'))
        out = self.call({'action': 'summary'})
        self.assertEqual(out['models'][0]['snapshots'], 2)       # inbox ingested on read
        rows = self.call({'action': 'query', 'options': {'category': 'Walls',
                                                         'select': ['Mark']}})
        self.assertEqual(rows['rows'][0]['params'], {'Mark': 'W1b'})
        self.assertEqual(self.call({'action': 'query', 'category': 'Doors'})['total'], 1)
        diff = self.call({'action': 'diff'})
        self.assertEqual(diff['changed_samples'][0]['changes'], {'Mark': ['W1', 'W1b']})
        self.assertEqual(self.call({'action': 'nope'})['error_code'], 'bad_arguments')

    def test_results_are_compact_json(self):
        self.publish(make_doc())
        request = {'jsonrpc': '2.0', 'id': 1, 'method': 'tools/call',
                   'params': {'name': 't3lab_hub', 'arguments': {'action': 'summary'}}}
        text = B._handle(request, self.current, 't')['result']['content'][0]['text']
        self.assertNotIn('\n', text)

    def test_empty_hub_says_how_to_publish(self):
        out = self.call({'action': 'query'})
        self.assertEqual(out['error_code'], 'hub_empty')
        self.assertIn('publish', out['hint'])

    def test_publish_runs_in_revit_then_ingests(self):
        seen = {}

        def fake_forward(request, current, token, timeout=None):
            seen['name'] = request['params']['name']
            seen['args'] = request['params']['arguments']
            return B._tool_response(request['id'], snap.publish(
                make_doc(), seen['args'], clock=self.clock, folder=self.hub.inbox))

        real = B._forward
        B._forward = fake_forward
        self.addCleanup(setattr, B, '_forward', real)
        out = self.call({'action': 'publish', 'options': {'categories': ['Walls']}})
        self.assertEqual(seen, {'name': 'hub_publish', 'args': {'categories': ['Walls']}})
        self.assertTrue(out['published'])
        self.assertEqual(out['hub']['seq'], 1)
        self.assertNotIn('file', out)
        self.assertEqual(self.hub.pending(), [])

        def revit_error(request, current, token, timeout=None):
            return B._tool_response(request['id'], {'error': 'No document'})
        B._forward = revit_error
        self.assertEqual(self.call({'action': 'publish'})['error'], 'No document')

    def test_cli(self):
        self.publish(make_doc())
        bridge = os.path.join(LIB, 'core', 'bridge.py')
        proc = subprocess.run([sys.executable, '-I', bridge, 'hub', 'summary'],
                              capture_output=True, text=True, encoding='utf-8', timeout=60,
                              env=dict(os.environ))
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertEqual(json.loads(proc.stdout)['models'][0]['model'], 'Tower A.rvt')
        proc = subprocess.run([sys.executable, '-I', bridge, 'hub', 'query', '{"category": "Doors"}'],
                              capture_output=True, text=True, encoding='utf-8', timeout=60,
                              env=dict(os.environ))
        self.assertEqual(json.loads(proc.stdout)['total'], 1)
        proc = subprocess.run([sys.executable, '-I', bridge, 'hub', 'rebuild'],
                              capture_output=True, text=True, encoding='utf-8', timeout=60,
                              env=dict(os.environ))
        self.assertEqual(len(json.loads(proc.stdout)['ingested']), 1)


class WiringTests(unittest.TestCase):
    def read(self, *parts):
        with open(os.path.join(LIB, *parts), encoding='utf-8') as f:
            return f.read()

    def test_server_registers_hub_publish(self):
        src = self.read('core', 'server.py')
        self.assertIn("'name': 'hub_publish'", src)
        self.assertIn("elif tool_name == 'hub_publish':", src)

    def test_hub_publish_is_read_only_for_the_policy(self):
        from Intelligence.tool_schema import READ_ONLY_TOOL_NAMES
        self.assertIn('hub_publish', READ_ONLY_TOOL_NAMES)

    def test_store_is_deployed_next_to_the_bridge(self):
        self.assertIn("('hub_store.py', 'hub_store.py')", self.read('bridge_deploy.py'))

    def test_store_and_snapshot_agree_on_the_folder_and_format(self):
        self.assertEqual(snap.hub_dir(), store.hub_dir())
        self.assertEqual(snap.FORMAT, store.FORMAT)

    def test_same_hub_code_as_t3lab_dev(self):
        """hub_store / hub_snapshot are shared with t3lab_dev: only the test
        pointers in their docstrings may differ, so snapshots stay readable
        by both extensions."""
        for name in ('hub_store.py', 'hub_snapshot.py'):
            src = self.read('core', name)
            self.assertIn("FORMAT = 'hub/1'", src)


if __name__ == '__main__':
    unittest.main()
