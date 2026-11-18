# -*- coding: utf-8 -*-
"""FamiGen Studio: controller protocol, page build, wiring and schema v2 extras.

No Revit, no WebView2, no network. The controller runs with fake adapters; the
page is checked statically (every id the script uses exists, every message
type has a handler on the other side). docs/famigen-studio-spec.md.
"""
import base64
import copy
import json
import os
import re
import sys
import tempfile
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='famigen_studio_')

from FamilyGen import guidance                        # noqa: E402
from FamilyGen import studio_controller as sc         # noqa: E402
from FamilyGen import studio_page                     # noqa: E402
from FamilyGen import studio_link                     # noqa: E402
from FamilyGen import templates as tpl                # noqa: E402
from FamilyGen.builder_report import report_lines     # noqa: E402
from FamilyGen.proposals import ProposalStore         # noqa: E402
from Intelligence import family_schema as fs          # noqa: E402

STUDIO = os.path.join(LIB, 'FamilyGen', 'studio')
GOOD = copy.deepcopy(fs.EXAMPLE_SCHEMA)
PNG = base64.b64encode(b'\x89PNG\r\n\x1a\n' + b'0' * 64).decode('ascii')


def submit(family, call_id='toolu_1', summary='Done.'):
    return json.dumps({'content': [{'type': 'tool_use', 'id': call_id, 'name': 'submit_family',
                                    'input': {'family': family, 'summary': summary}}],
                       'stop_reason': 'tool_use',
                       'usage': {'input_tokens': 10, 'output_tokens': 5}})


class FakeHttp(object):
    def __init__(self, *replies):
        self.replies = list(replies)
        self.bodies = []
        self.cancelled = 0

    def start(self, url, headers, body, callback):
        self.bodies.append(json.loads(body))
        self.headers = headers
        status, raw = self.replies.pop(0)
        callback(status, raw)

    def cancel(self):
        self.cancelled += 1


class FakeRevit(object):
    into_family = False

    def __init__(self):
        self.files = []
        self.created = []

    def pick_files(self):
        return self.files

    def pick_json_file(self):
        return None

    def save_json_file(self, name, text):
        return None

    def pick_folder(self, initial):
        return r'C:\Families'

    def copy_text(self, text):
        self.copied = text

    def close(self):
        pass

    def window_state(self, action):
        self.state = action

    def create_family(self, schema, folder, load, callback):
        self.created.append((schema['family_name'], folder, load))
        callback({'built': 3, 'total': 3, 'skipped': [], 'warnings': [], 'materials_created': ['Oak'],
                  'materials_reused': [], 'material_parameters': [], 'parameters_created': [],
                  'parameters_set': [], 'subcategories': [], 'types': ['Oak Top'],
                  'reference_planes': ['Top'], 'dimensions': ['Height'], 'locks': 2,
                  'saved_path': folder + r'\x.rfa'}, None)


class FakeClaude(object):
    def __init__(self, key='sk-test'):
        self.key = key

    def api_key(self):
        return self.key

    def models(self):
        return ['claude-opus-5-5', 'claude-sonnet-5-5']

    def default_model(self):
        return 'claude-opus-5-5'


class Prefs(dict):
    def set(self, key, value):
        self[key] = value

    def save(self):
        pass


def controller(http=None, revit=None, claude=None, **prefs):
    posts = []
    ctl = sc.StudioController(posts.append, http or FakeHttp(), revit or FakeRevit(),
                              claude or FakeClaude(), Prefs(prefs),
                              store=ProposalStore(tempfile.mkdtemp()),
                              overlay_reader=lambda c: u'SPEC FOR ' + c, pyrevit_version='7.1.0')
    return ctl, posts


def of_type(posts, kind):
    return [p for p in posts if p.get('type') == kind]


class ControllerTests(unittest.TestCase):
    def test_init(self):
        ctl, posts = controller()
        ctl.handle({'type': 'ready'})
        init = of_type(posts, 'init')[0]
        self.assertEqual(init['categories'][0], 'Auto')
        self.assertEqual(init['category'], 'Auto')
        self.assertTrue(init['claude']['ready'])
        self.assertEqual(init['claude']['model'], 'claude-opus-5-5')
        self.assertTrue(init['templates'])
        self.assertIn('MEP: Piping & supports', init['template_groups'])

    def test_no_key_explains_and_never_sends(self):
        http = FakeHttp()
        ctl, posts = controller(http=http, claude=FakeClaude(key=None))
        ctl.handle({'type': 'ready'})
        self.assertFalse(of_type(posts, 'init')[0]['claude']['ready'])
        ctl.handle({'type': 'send', 'text': 'A table'})
        self.assertEqual(http.bodies, [])
        self.assertIn('LLMs Setting', of_type(posts, 'chat')[-1]['text'])

    def test_send_auto_detects_category_and_offers_spec_tool(self):
        http = FakeHttp((200, submit(GOOD)))
        ctl, posts = controller(http=http)
        ctl.handle({'type': 'send', 'text': 'Side table', 'json': ''})
        body = http.bodies[0]
        self.assertEqual([t['name'] for t in body['tools']], ['submit_family', 'get_category_spec'])
        self.assertIn('CATEGORY GUIDE', body['system'][0]['text'])
        self.assertEqual(of_type(posts, 'detected')[0]['category'], 'Furniture')
        self.assertEqual(json.loads(of_type(posts, 'schema')[0]['json'])['family_name'], 'Side Table 500')
        self.assertTrue(of_type(posts, 'preview')[0]['ok'])
        self.assertEqual(posts[-1]['tone'], 'success')
        self.assertFalse(ctl.busy)
        self.assertEqual(len(ctl.store.ids()), 1)               # kept in My families

    def test_category_spec_tool_round_trip(self):
        guide = json.dumps({'content': [{'type': 'tool_use', 'id': 'g1', 'name': 'get_category_spec',
                                         'input': {'category': 'Furniture'}}],
                            'stop_reason': 'tool_use', 'usage': {}})
        http = FakeHttp((200, guide), (200, submit(GOOD, 'toolu_2')))
        ctl, posts = controller(http=http)
        ctl.handle({'type': 'send', 'text': 'Side table'})
        result = http.bodies[1]['messages'][-1]['content'][0]
        self.assertEqual(result['tool_use_id'], 'g1')
        self.assertEqual(result['content'], 'SPEC FOR Furniture')
        self.assertTrue(of_type(posts, 'schema'))

    def test_fixed_category_puts_spec_in_system_prompt(self):
        http = FakeHttp((200, submit(GOOD)))
        ctl, _ = controller(http=http, category='Furniture')
        ctl.handle({'type': 'send', 'text': 'Side table'})
        body = http.bodies[0]
        self.assertEqual([t['name'] for t in body['tools']], ['submit_family'])
        self.assertIn('SPEC FOR Furniture', body['system'][0]['text'])

    def test_reference_check_round(self):
        http = FakeHttp((200, submit(GOOD)),
                        (200, json.dumps({'content': [{'type': 'text', 'text': 'Matches the photo.'}],
                                          'stop_reason': 'end_turn', 'usage': {}})))
        ctl, posts = controller(http=http)
        ctl.handle({'type': 'attach_data', 'name': 'photo.png', 'media_type': 'image/png', 'data': PNG})
        self.assertEqual(of_type(posts, 'attachments')[-1]['items'][0]['name'], 'photo.png')
        ctl.handle({'type': 'send', 'text': 'Model this'})
        request = of_type(posts, 'snapshot_request')[0]
        self.assertTrue(ctl.busy)
        ctl.handle({'type': 'snapshot', 'token': request['token'], 'views': ['iso'],
                    'images': ['data:image/png;base64,' + PNG]})
        check = http.bodies[1]['messages'][-1]['content']
        self.assertEqual(check[0]['type'], 'tool_result')                # owed result first
        self.assertEqual(check[1]['type'], 'image')
        self.assertIn('Matches', check[-1]['text'])
        self.assertEqual(of_type(posts, 'chat')[-1]['text'], 'Matches the photo.')
        self.assertFalse(ctl.busy)
        ctl.handle({'type': 'snapshot', 'token': request['token'], 'images': [PNG]})  # stale: ignored
        self.assertEqual(len(http.bodies), 2)

    def test_check_can_be_turned_off(self):
        http = FakeHttp((200, submit(GOOD)))
        ctl, posts = controller(http=http)
        ctl.handle({'type': 'set_check', 'on': False})
        ctl.handle({'type': 'attach_data', 'name': 'p.png', 'media_type': 'image/png', 'data': PNG})
        ctl.handle({'type': 'send', 'text': 'Model this'})
        self.assertEqual(of_type(posts, 'snapshot_request'), [])

    def test_attach_rejects_unsupported_files(self):
        ctl, posts = controller()
        ctl.handle({'type': 'attach_data', 'name': 'x.exe', 'media_type': 'application/x-msdownload', 'data': PNG})
        self.assertEqual(posts[-1]['items'], [])
        self.assertEqual(of_type(posts, 'status')[-1]['tone'], 'warning')
        text = base64.b64encode(u'Width 600 mm'.encode('utf-8')).decode('ascii')
        ctl.handle({'type': 'attach_data', 'name': 'spec.txt', 'media_type': 'text/plain', 'data': text})
        self.assertEqual(posts[-1]['items'][0]['kind'], 'text')

    def test_cancel_resets_the_conversation(self):
        class Slow(FakeHttp):
            def start(self, url, headers, body, callback):
                self.bodies.append(json.loads(body))
                self.later = callback
        http = Slow()
        ctl, posts = controller(http=http)
        ctl.handle({'type': 'send', 'text': 'A table'})
        self.assertTrue(ctl.busy)
        ctl.handle({'type': 'cancel'})
        self.assertFalse(ctl.busy)
        self.assertIsNone(ctl.session)
        self.assertEqual(http.cancelled, 1)
        http.later(200, submit(GOOD))                      # late answer is dropped
        self.assertEqual(of_type(posts, 'schema'), [])

    def test_create_validates_then_builds(self):
        revit = FakeRevit()
        ctl, posts = controller(revit=revit)
        ctl.handle({'type': 'create', 'json': '{"family_name": "x", "geometry": []}'})
        self.assertEqual(revit.created, [])
        self.assertEqual(of_type(posts, 'chat')[-1]['role'], 'error')
        ctl.handle({'type': 'create', 'json': json.dumps(GOOD), 'load': True})
        self.assertEqual(revit.created, [('Side Table 500', r'C:\Families', True)])
        created = of_type(posts, 'created')[-1]
        self.assertIn('Parametric: 1 reference plane(s), 1 labelled dimension(s), 2 locked face(s).',
                      created['lines'])

    def test_history_lists_and_loads(self):
        http = FakeHttp((200, submit(GOOD)))
        ctl, posts = controller(http=http, category='Furniture')
        ctl.handle({'type': 'send', 'text': 'Side table'})
        ctl.handle({'type': 'history'})
        items = of_type(posts, 'history')[-1]['items']
        self.assertEqual(items[0]['family_name'], 'Side Table 500')
        ctl.handle({'type': 'load_history', 'id': items[0]['id']})
        self.assertEqual(of_type(posts, 'schema')[-1]['source'], 'history')

    def test_unknown_message_and_handler_errors_never_raise(self):
        ctl, posts = controller()
        ctl.handle({'type': 'nope'})
        ctl.handle('not a dict')
        self.assertEqual(posts[-1]['tone'], 'warning')


class PreviewTests(unittest.TestCase):
    def schema(self):
        s = copy.deepcopy(GOOD)
        s['parameters'].append({'name': 'Height', 'type': 'length', 'value': 550})
        s['reference_planes'] = [{'name': 'Top', 'axis': 'z', 'offset': 550}]
        s['dimensions'] = [{'parameter': 'Height', 'from': 'Ref. Level', 'to': 'Top'}]
        s['geometry'][0]['locks'] = {'z_max': 'Top'}
        s['geometry'][1]['repeat'] = {'count': 3, 'step': [100, 0, 0]}
        s['geometry'][1]['detail_levels'] = ['fine']
        return s

    def test_payload_has_planes_dimensions_levels_and_repeats(self):
        p = sc.preview_payload(json.dumps(self.schema()))
        self.assertTrue(p['ok'], p['errors'])
        self.assertEqual(len(p['parts']), 5)                   # 1 + 3 legs + base
        self.assertEqual([x['label'] for x in p['parts']][1:4], ['Leg 1', 'Leg 2', 'Leg 3'])
        self.assertEqual(p['parts'][1]['levels'], ['fine'])
        self.assertEqual(p['dimensions'][0]['value'], 550.0)
        self.assertIn('Top', [pl['name'] for pl in p['planes'] if not pl['template']])
        self.assertEqual(p['warnings'], [])

    def test_lock_off_the_plane_is_reported(self):
        s = self.schema()
        s['geometry'][0]['locks'] = {'z_min': 'Ref. Level'}     # the top's underside is at 520
        p = sc.preview_payload(json.dumps(s))
        self.assertTrue(any('will not lock' in w for w in p['warnings']))

    def test_parse_error_and_empty(self):
        self.assertTrue(sc.preview_payload('{oops')['parse_error'])
        self.assertFalse(sc.preview_payload('')['ok'])


class SchemaExtrasTests(unittest.TestCase):
    def test_expand_repeat_moves_points_and_sketch_planes(self):
        s = copy.deepcopy(GOOD)
        s['geometry'][0]['repeat'] = {'count': 2, 'step': [0, 0, 100]}
        s['geometry'][1]['repeat'] = {'at': [[0, 0, 0], [50, 60, 0]]}
        e = fs.expand_schema(s)
        self.assertNotIn('repeat', json.dumps(e))
        self.assertEqual(e['geometry'][1]['sketch_plane_z'], 100.0)
        self.assertEqual(e['geometry'][3]['start'], [50.0, 60.0, 0.0])
        self.assertIn('repeat', s['geometry'][0])               # input untouched

    def test_bad_constraints_and_repeats_are_errors(self):
        s = copy.deepcopy(GOOD)
        s['reference_planes'] = [{'name': 'Ref. Level', 'axis': 'z', 'offset': 0},
                                 {'name': 'P', 'axis': 'w', 'offset': 1}]
        s['dimensions'] = [{'parameter': 'Designer', 'from': 'P', 'to': 'Nope'}]
        s['geometry'][0]['locks'] = {'x_min': 'Ref. Level'}
        s['geometry'][1]['repeat'] = {'count': 1, 'step': [1, 0, 0]}
        s['geometry'][2]['detail_levels'] = ['ultra']
        text = '\n'.join(fs.validate_family_schema(s)[0])
        for expected in ('already exists', 'must be "x", "y" or "z"', 'a dimension needs length',
                         "unknown plane 'Nope'", 'x_min needs an x plane', 'from 2 to',
                         'detail_levels'):
            self.assertIn(expected, text)

    def test_auto_prompt_lists_every_category(self):
        prompt = fs.build_system_prompt('Auto', delivery='tool')
        for category in fs.SUPPORTED_CATEGORIES:
            self.assertIn(category, prompt)
        self.assertEqual({c for c, _ in fs.CATEGORY_GUIDE}, set(fs.SUPPORTED_CATEGORIES))

    def test_every_category_has_a_spec_and_template_rows_are_valid(self):
        folder = os.path.join(REPO, 'T3Lab.tab', 'Standards & Families.panel', 'FamiGen.pushbutton', 'prompts')
        for category in fs.SUPPORTED_CATEGORIES:
            path = os.path.join(folder, guidance.category_slug(category) + '.md')
            self.assertTrue(os.path.isfile(path), category)
            with open(path, encoding='utf-8') as handle:
                self.assertNotIn('does **not** read `reference_planes`', handle.read())
        for row in tpl.templates():
            self.assertIn(row['category'], fs.SUPPORTED_CATEGORIES, row['title'])
            self.assertIn(row['group'], tpl.groups())

    def test_report_lines_mention_types_and_constraints(self):
        lines = report_lines({'built': 1, 'total': 1, 'types': ['A', 'B'], 'reference_planes': ['L'],
                              'dimensions': ['W'], 'locks': 2})
        self.assertIn('Types: A, B.', lines)
        self.assertTrue(any(l.startswith('Parametric:') for l in lines))


class PageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.html = studio_page.inline_page()
        with open(os.path.join(STUDIO, 'studio.js'), encoding='utf-8') as handle:
            cls.js = handle.read()
        with open(os.path.join(STUDIO, 'studio.html'), encoding='utf-8') as handle:
            cls.template = handle.read()
        with open(os.path.join(LIB, 'FamilyGen', 'studio_controller.py'), encoding='utf-8') as handle:
            cls.py = handle.read()

    def test_one_self_contained_file(self):
        self.assertNotIn('/*__', self.html)
        self.assertNotIn('__T3_LOGO_DATA__', self.html)
        self.assertIn('src="data:image/png;base64,', self.html)
        self.assertEqual(self.html.count('<script>'), 2)
        self.assertNotRegex(self.html, r'<script[^>]+src=')        # nothing loaded from outside
        self.assertNotRegex(self.html, r'https?://(?!threejs\.org)[^"\s]+\.(js|css)')

    def test_every_element_id_used_by_the_script_exists(self):
        ids = set(re.findall(r'id="([^"]+)"', self.template))
        used = set(re.findall(r"\$\('([^']+)'\)", self.js))
        self.assertEqual(sorted(used - ids), [])

    def test_messages_have_handlers_on_both_sides(self):
        sent = set(re.findall(r"bridge\.send\(\{ type: '([a-z_]+)'", self.js))
        handled_py = set(re.findall(r'def _on_([a-z_]+)\(self, message\)', self.py))
        self.assertEqual(sorted(sent - handled_py), [])
        posted = set(re.findall(r"'type': '([a-z_]+)'", self.py))
        posted |= {'preview'}                                   # payload['type'] = 'preview'
        handled_js = set(re.findall(r'^    ([a-z_]+): function', self.js, re.M))
        self.assertEqual(sorted(posted - handled_js), [])

    def test_version_gate(self):
        self.assertFalse(studio_page.requirement((6, 5, 5))[0])
        self.assertIn('6.5.5', studio_page.requirement((6, 5, 5))[1])
        self.assertTrue(studio_page.requirement((7, 0, 1))[0])
        self.assertFalse(studio_page.requirement(None)[0])

    def test_three_js_is_the_vendored_r158(self):
        with open(os.path.join(STUDIO, 'vendor', 'LICENSE-three.txt'), encoding='utf-8') as handle:
            self.assertIn('r158', handle.read())
        self.assertIn('"158"', self.html)                      # three.js REVISION


class McpModeTests(unittest.TestCase):
    def setUp(self):
        os.environ['APPDATA'] = tempfile.mkdtemp(prefix='famigen_link_')

    def test_link_files(self):
        self.assertFalse(studio_link.studio_open())
        studio_link.beat({'category': 'Furniture'})
        self.assertTrue(studio_link.studio_open())
        self.assertFalse(studio_link.studio_open(clock=lambda: 1e12))
        studio_link.leave()
        self.assertFalse(studio_link.studio_open())
        self.assertIsNone(studio_link.active_client())
        studio_link.note_activity('claude-ai', 'famigen_get_schema')
        self.assertEqual(studio_link.active_client(), 'Claude Desktop')
        studio_link.post_message('Make it 2 m long', ['a.png'])
        self.assertEqual(studio_link.unread_count(), 1)
        self.assertEqual(studio_link.take_unread()[0]['files'], ['a.png'])
        self.assertEqual(studio_link.take_unread(), [])

    def test_patch_family_by_part_id(self):
        out, notes = studio_link.patch_family(GOOD, {'family_name': 'Side Table 600'},
                                              [{'id': 'Leg', 'radius': 20},
                                               {'id': 'Base', 'remove': True},
                                               {'id': 'Foot', 'type': 'Cylinder', 'start': [0, 0, 0],
                                                'end': [0, 0, 10], 'radius': 40, 'material': 'Oak'}])
        self.assertEqual(out['family_name'], 'Side Table 600')
        self.assertEqual([g['id'] for g in out['geometry']], ['Top', 'Leg', 'Foot'])
        self.assertEqual(out['geometry'][1]['radius'], 20)
        self.assertEqual(GOOD['geometry'][1]['radius'], 25)          # input untouched
        self.assertEqual(fs.validate_family_schema(out)[0], [])
        self.assertEqual(len(notes), 4)
        with self.assertRaises(ValueError):
            studio_link.patch_family(GOOD, parts=[{'id': 'Ghost', 'radius': 1}])
        with self.assertRaises(ValueError):
            studio_link.patch_family(GOOD, {'geometry': []})

    def test_controller_switches_to_the_ai_app_and_back(self):
        http = FakeHttp()
        ctl, posts = controller(http=http)
        ctl.tick(now=100.0)
        self.assertTrue(studio_link.studio_open())                   # heartbeat
        self.assertEqual(ctl.engine(), 'claude')
        studio_link.note_activity('claude-code', 'famigen_get_studio_state')
        ctl.tick(now=101.0)
        self.assertEqual(ctl.engine(), 'mcp')
        self.assertEqual(of_type(posts, 'engine')[-1]['client'], 'Claude Code')
        self.assertIn('continue in FamiGen', of_type(posts, 'chat')[-1]['text'])
        ctl.handle({'type': 'send', 'text': 'Make the legs thinner', 'json': json.dumps(GOOD)})
        self.assertEqual(http.bodies, [])                            # no API tokens used
        self.assertEqual(studio_link.take_unread()[0]['text'], 'Make the legs thinner')
        self.assertEqual(studio_link.read_state()['family']['family_name'], 'Side Table 500')
        ctl.handle({'type': 'set_engine', 'engine': 'claude'})
        self.assertEqual(ctl.engine(), 'claude')
        ctl.close_link()
        self.assertFalse(studio_link.studio_open())

    def test_new_mcp_proposal_appears_in_studio(self):
        ctl, posts = controller()
        studio_link.note_activity('claude-ai', 'famigen_propose_family')
        ctl.tick(now=1.0)
        ctl.store.save(GOOD, source='mcp', note='Side table from Claude Desktop')
        ctl.tick(now=2.0)
        self.assertEqual(of_type(posts, 'schema')[-1]['source'], 'mcp')
        self.assertEqual(of_type(posts, 'chat')[-1]['text'], 'Side table from Claude Desktop')
        count = len(of_type(posts, 'schema'))
        ctl.tick(now=3.0)                                            # shown once
        self.assertEqual(len(of_type(posts, 'schema')), count)

    def test_friendly_words_for_people(self):
        bad = copy.deepcopy(GOOD)
        bad['geometry'][1]['material'] = 'Pine'
        p = sc.preview_payload(json.dumps(bad))
        self.assertTrue(p['errors'][0].startswith('$.geometry[1].material'))   # Claude keeps paths
        self.assertTrue(p['issues']['errors'][0].startswith("Part 'Leg' (material): unknown"))
        self.assertEqual(sc.friendly_model('claude-opus-5-5'), 'Claude Opus 5.5')
        self.assertEqual(sc.friendly_model('claude-sonnet-5'), 'Claude Sonnet 5')
        self.assertEqual(sc.friendly_model('gpt-x'), 'gpt-x')

    def test_server_and_toolsets_know_the_studio_tools(self):
        with open(os.path.join(LIB, 'core', 'server.py'), encoding='utf-8') as handle:
            server = handle.read()
        from Intelligence.tool_schema import READ_ONLY_TOOL_NAMES
        for name in ('famigen_get_studio_state', 'famigen_validate_family',
                     'famigen_patch_family', 'famigen_list_templates'):
            self.assertIn("'{0}': {{\n                'name': '{0}'".format(name), server)
            self.assertIn("elif tool_name == '{}':".format(name), server)
            self.assertIn(name, READ_ONLY_TOOL_NAMES)
        enum = re.search(r"'enum': \['Generic Model'.*?\]", server, re.S).group(0)
        for category in fs.SUPPORTED_CATEGORIES:
            self.assertIn("'{}'".format(category), enum)


class ContinuityTests(unittest.TestCase):
    """Tube (smooth curved members) and the loose-part check (docs §5, §9)."""

    def hoop(self, **kw):
        g = {'id': 'Hoop', 'type': 'Tube', 'material': 'Oak', 'radius': 24,
             'points': [[0, 0, 0], [0, 0, 700], [300, 0, 900], [600, 0, 700], [600, 0, 0]]}
        g.update(kw)
        s = copy.deepcopy(GOOD)
        s['geometry'] = [g]
        return s

    def test_tube_is_valid_and_expands_to_overlapping_cylinders(self):
        s = self.hoop()
        self.assertEqual(fs.validate_family_schema(s)[0], [])
        parts = fs.expand_schema(s)['geometry']
        cyls = [p for p in parts if p['type'] == 'Cylinder']
        self.assertEqual(len(cyls), 4 * fs.TUBE_SUBDIVISIONS)        # 4 spans, spline
        self.assertTrue(all(p['material'] == 'Oak' and p['radius'] == 24 for p in cyls))
        self.assertEqual(cyls[0]['start'], [0.0, 0.0, 0.0])           # the line starts on the floor
        # consecutive segments overlap (each grows past the shared point)
        a_end, b_start = cyls[3]['end'], cyls[4]['start']
        self.assertNotEqual(a_end, b_start)

    def test_sharp_corner_gets_a_ball_smooth_bend_does_not(self):
        sharp = fs.expand_schema(self.hoop(smooth=False, points=[[0, 0, 0], [0, 0, 700], [600, 0, 700]]))
        self.assertEqual([p['id'] for p in sharp['geometry'] if 'joint' in p['id']], ['Hoop joint 1'])
        smooth = fs.expand_schema(self.hoop())
        self.assertEqual([p for p in smooth['geometry'] if 'joint' in p['id']], [])

    def test_tube_rules(self):
        bad = self.hoop(points=[[0, 0, 0]], radius=0, smooth='yes')
        text = '\n'.join(fs.validate_family_schema(bad)[0])
        self.assertIn('points', text)
        self.assertIn('radius', text)
        self.assertIn('true or false', text)

    def test_repeat_moves_tube_points(self):
        s = self.hoop(repeat={'count': 2, 'step': [1000, 0, 0]})
        cyl = [p for p in fs.expand_schema(s)['geometry'] if p['id'] == 'Hoop 2 1'][0]
        self.assertEqual(cyl['start'][0], 1000.0)

    def test_floating_part_is_reported_by_name(self):
        from FamilyGen.preview_mesh import build_preview
        s = copy.deepcopy(GOOD)
        self.assertFalse([w for w in build_preview(s).warnings if w.startswith('Not connected')])
        s['geometry'].append({'id': 'Lamp', 'type': 'Cylinder', 'material': 'Oak',
                              'start': [0, 0, 900], 'end': [0, 0, 1000], 'radius': 50})
        loose = [w for w in build_preview(s).warnings if w.startswith('Not connected')]
        self.assertEqual(len(loose), 1)
        self.assertIn("'Lamp'", loose[0])

    def test_polish_round_sends_loose_parts_back_once(self):
        from FamilyGen import claude_designer as designer
        loose = copy.deepcopy(GOOD)
        loose['geometry'].append({'id': 'Lamp', 'type': 'Cylinder', 'material': 'Oak',
                                  'start': [0, 0, 900], 'end': [0, 0, 1000], 'radius': 50})
        session = designer.ClaudeFamilySession('Furniture', 'claude-opus-5-5')
        session.begin_turn('Side table')
        first = session.handle_http(200, submit(loose, 'a'))
        self.assertEqual(first.kind, 'continue')
        result = first.request['messages'][-1]['content'][0]
        self.assertEqual(result['tool_use_id'], 'a')
        self.assertNotIn('is_error', result)
        self.assertIn("'Lamp'", result['content'])
        second = session.handle_http(200, submit(loose, 'c'))      # still loose: shown, not looped
        self.assertEqual(second.kind, 'schema')
        self.assertTrue(any('Not connected' in w for w in second.warnings))


class FormToolboxTests(unittest.TestCase):
    """Sweep with local profiles, SweptBlend and voids (docs §5.2)."""

    U = [{'type': 'Line', 'start': [a[0], a[1], 0], 'end': [b[0], b[1], 0]} for a, b in
         zip([[-300, 0], [300, 0], [300, 200], [-300, 200]], [[300, 0], [300, 200], [-300, 200], [-300, 0]])]

    def family(self, *geometry):
        s = copy.deepcopy(GOOD)
        s['geometry'] = list(geometry)
        return s

    def test_frame_rule(self):
        from FamilyGen.form_frames import frame
        u, v = frame([1, 0, 0])
        self.assertEqual([round(c, 6) for c in u], [0, 1, 0])     # sideways = +Y
        self.assertEqual([round(c, 6) for c in v], [0, 0, 1])     # up = +Z
        u, v = frame([0, 0, 1])                                    # vertical path
        self.assertEqual([round(c, 6) for c in u], [1, 0, 0])

    def test_local_profile_to_world_and_arc_tangents(self):
        from FamilyGen.form_frames import local_to_world, segment_ends
        world = local_to_world([{'type': 'Circle', 'center': [0, 0, 0], 'radius': 100}],
                               [10, 20, 30], [0, 1, 0], [0, 0, 1])
        self.assertEqual([w['type'] for w in world], ['Arc3P', 'Arc3P'])
        self.assertEqual(world[0]['start'], [10.0, 120.0, 30.0])
        (s, ts), (e, te) = segment_ends({'type': 'Arc3P', 'start': [0, 0, 0], 'end': [100, 0, 0],
                                         'mid': [50, 0, 50]})
        self.assertGreater(ts[2], 0.9)                             # leaves upwards
        self.assertLess(te[2], -0.9)                               # arrives downwards

    def test_local_sweep_preview_sits_on_the_path(self):
        from FamilyGen.preview_mesh import build_preview
        sweep = {'id': 'Trough', 'type': 'Sweep', 'material': 'Oak', 'profile_local': True,
                 'sketch_plane_y': 0, 'profile': self.U,
                 'path': [{'type': 'Line', 'start': [0, 0, 0], 'end': [1000, 0, 0]}]}
        model = build_preview(self.family(sweep))
        self.assertEqual(fs.validate_family_schema(self.family(sweep))[0], [])
        self.assertEqual([round(v) for v in model.bbox_min], [0, -300, 0])
        self.assertEqual([round(v) for v in model.bbox_max], [1000, 300, 200])

    def test_swept_blend_rules_and_preview(self):
        from FamilyGen.preview_mesh import build_preview
        sb = {'id': 'Chute', 'type': 'SweptBlend', 'material': 'Oak', 'sketch_plane_y': 0,
              'path': [{'type': 'Arc3P', 'start': [0, 0, 1000], 'end': [1000, 0, 0], 'mid': [700, 0, 700]}],
              'profile': [{'type': 'Circle', 'center': [0, 0, 0], 'radius': 300}],
              'top_profile': [{'type': 'Ellipse', 'center': [0, 0, 0], 'radius_x': 450, 'radius_y': 150}]}
        self.assertEqual(fs.validate_family_schema(self.family(sb))[0], [])
        model = build_preview(self.family(sb))
        self.assertEqual(model.warnings, [])
        self.assertAlmostEqual(model.size_mm()[1], 900, delta=5)  # widest section = end ellipse
        two = dict(sb, path=sb['path'] + [{'type': 'Line', 'start': [1000, 0, 0], 'end': [2000, 0, 0]}])
        self.assertIn('ONE segment', '\n'.join(fs.validate_family_schema(self.family(two))[0]))
        spline = dict(sb, path=[{'type': 'Spline', 'points': [[0, 0, 0], [1, 0, 1], [2, 0, 0]]}])
        self.assertTrue(fs.validate_family_schema(self.family(spline))[0])
        rep = dict(sb, repeat={'count': 2, 'step': [10, 0, 0]})
        self.assertIn('cannot repeat', '\n'.join(fs.validate_family_schema(self.family(rep))[0]))

    def test_contract_names_every_form(self):
        contract = fs.schema_contract()
        for form in ('Extrusion', 'Blend', 'Revolution', 'Sweep', 'SweptBlend', 'Cylinder', 'Tube', 'Void'):
            self.assertIn(form, contract['form_guide'])
        prompt = fs.build_system_prompt('Site', delivery='tool')
        self.assertIn('use the whole Revit toolbox', prompt)
        self.assertIn('SweptBlend', prompt)

    def test_studio_can_hide_voids(self):
        with open(os.path.join(STUDIO, 'studio.js'), encoding='utf-8') as handle:
            js = handle.read()
        self.assertIn("m.part.solid || $('chk-voids').checked", js)


class ReferenceAnalysisTests(unittest.TestCase):
    """Pipe elbows (Tube bend_radius) and the record_analysis -> camera check flow."""

    def test_bend_radius_makes_round_elbows_without_balls(self):
        pipe = {'id': 'Pipe', 'type': 'Tube', 'material': 'Oak', 'radius': 150, 'bend_radius': 300,
                'points': [[0, 0, 0], [0, 0, 2500], [0, -1300, 2500], [0, -1300, 2000]]}
        s = copy.deepcopy(GOOD)
        s['geometry'] = [pipe]
        self.assertEqual(fs.validate_family_schema(s)[0], [])
        parts = fs.expand_schema(s)['geometry']
        self.assertFalse([p for p in parts if 'joint' in p['id']])
        self.assertGreater(len(parts), 10)                         # two arcs of short segments
        from FamilyGen.preview_mesh import build_preview
        model = build_preview(s)
        self.assertEqual([round(v) for v in model.bbox_min], [-150, -1450, 0])
        self.assertLess(model.bbox_max[2], 2651)                   # the elbow stays inside the corner
        too_tight = dict(pipe, bend_radius=100)
        s['geometry'] = [too_tight]
        self.assertIn('at least the tube radius', '\n'.join(fs.validate_family_schema(s)[0]))

    def test_analysis_tool_only_with_references_then_continue(self):
        from FamilyGen import claude_designer as designer
        plain = designer.ClaudeFamilySession('Furniture', 'claude-opus-5-5')
        body = plain.begin_turn('A side table')
        self.assertNotIn('record_analysis', [t['name'] for t in body['tools']])
        ref = designer.ClaudeFamilySession('Furniture', 'claude-opus-5-5')
        att = {'name': 'p.png', 'kind': 'image', 'media_type': 'image/png', 'size': 10, 'data': PNG}
        body = ref.begin_turn('Model this', attachments=[att])
        self.assertIn('record_analysis', [t['name'] for t in body['tools']])
        self.assertIn('REFERENCE ANALYSIS', body['system'][0]['text'])
        analysis = {'camera': {'azimuth_deg': -110, 'elevation_deg': 30},
                    'overall_size_mm': [500, 500, 550], 'components': [{'name': 'Top', 'form': 'Extrusion'}]}
        reply = json.dumps({'content': [{'type': 'tool_use', 'id': 'an1', 'name': 'record_analysis',
                                         'input': analysis}], 'stop_reason': 'tool_use', 'usage': {}})
        out = ref.handle_http(200, reply)
        self.assertEqual(out.kind, 'continue')
        self.assertEqual(out.analysis['camera']['azimuth_deg'], -110)
        owed = out.request['messages'][-1]['content'][0]
        self.assertEqual((owed['tool_use_id'], owed.get('is_error')), ('an1', None))

    def test_controller_shows_analysis_and_checks_from_the_reference_camera(self):
        analysis = {'camera': {'azimuth_deg': -112, 'elevation_deg': 31}, 'view_type': 'render',
                    'overall_size_mm': [500, 500, 550], 'scale_basis': 'seat height',
                    'components': [{'name': 'Leg', 'form': 'Cylinder', 'count': 4}]}
        first = json.dumps({'content': [{'type': 'tool_use', 'id': 'an1', 'name': 'record_analysis',
                                         'input': analysis}], 'stop_reason': 'tool_use', 'usage': {}})
        http = FakeHttp((200, first), (200, submit(GOOD, 's1')))
        ctl, posts = controller(http=http)
        ctl.handle({'type': 'attach_data', 'name': 'photo.png', 'media_type': 'image/png', 'data': PNG})
        ctl.handle({'type': 'send', 'text': 'Model this'})
        card = [p for p in of_type(posts, 'chat') if p.get('role') == 'analysis'][0]
        self.assertIn('Leg ×4 · Cylinder', card['assumptions'])
        self.assertEqual(of_type(posts, 'ref_camera')[0]['azimuth_deg'], -112.0)
        request = of_type(posts, 'snapshot_request')[0]
        self.assertEqual(request['views'], ['ref', 'iso'])
        self.assertEqual(request['camera']['elevation_deg'], 31.0)


if __name__ == '__main__':
    unittest.main()
