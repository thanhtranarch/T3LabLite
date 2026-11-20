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
        ctl, posts = controller(http=http, claude=FakeClaude(key=None), engine='claude')
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
        self.assertEqual(of_type(posts, 'status')[-1]['tone'], 'success')
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
        self.assertIn('listen command', of_type(posts, 'chat')[-1]['text'])
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
        for call in ('c', 'd')[:designer.MAX_POLISH - 1]:
            self.assertEqual(session.handle_http(200, submit(loose, call)).kind, 'continue')
        last = session.handle_http(200, submit(loose, 'z'))        # still loose: shown, not looped
        self.assertEqual(last.kind, 'schema')
        self.assertTrue(any('Not connected' in w for w in last.warnings))


class FormToolboxTests(unittest.TestCase):
    """Sweep with local profiles, SweptBlend and voids (docs §5.2)."""

    U = [{'type': 'Line', 'start': [a[0], a[1], 0], 'end': [b[0], b[1], 0]} for a, b in
         zip([[-300, 0], [300, 0], [300, 200], [-300, 200]], [[300, 0], [300, 200], [-300, 200], [-300, 0]])]

    def family(self, *geometry):
        s = copy.deepcopy(GOOD)
        s['geometry'] = list(geometry)
        s.pop('joints', None)                    # the example's joints name the example's parts
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


class T3StyleTests(unittest.TestCase):
    """studio.css stays inside the T3 design system (docs/ui-governance/08-design-system-authority.md)."""

    TOKENS = {'#18181B', '#FFFFFF', '#27272A', '#52525B', '#71717A', '#9A9AA2', '#DCDCE0', '#A1A1AA',
              '#F4F4F6', '#E4E4E7', '#FAFAFB', '#F1F1F4', '#EAF8F0', '#22A85C', '#157038', '#FFF6E6',
              '#F5CE5A', '#8A6308', '#FDECEC', '#F87171', '#D23B3B', '#F6D5D5', '#ECECEF', '#C2410C',
              '#F59E0B', '#2563EB', '#D97706', '#059669', '#DC2626', '#7C3AED', '#0891B2', '#C026D3',
              '#65A30D', '#C7C7CE', '#000000'}       # palette + Plot.* + Plot.Off + Primary hover

    @classmethod
    def setUpClass(cls):
        with open(os.path.join(STUDIO, 'studio.css'), encoding='utf-8') as handle:
            raw = handle.read()
        # LIQUID GLASS: the owner's waiver (2026-11-19) for the chat column; checked on its own.
        glass = re.search(r'/\* ── LIQUID GLASS ──.*?/\* END LIQUID GLASS \*/', raw, re.S)
        cls.glass = re.sub(r'/\*.*?\*/', '', glass.group(0), flags=re.S) if glass else ''
        raw = raw.replace(glass.group(0), '') if glass else raw
        # TIMELINE GEOMETRY: exact pixels for the dots and segments (2026-11-20), checked below.
        geometry = re.search(r'/\* ── TIMELINE GEOMETRY ──.*?/\* END TIMELINE GEOMETRY \*/', raw, re.S)
        cls.geometry = re.sub(r'/\*.*?\*/', '', geometry.group(0), flags=re.S) if geometry else ''
        raw = raw.replace(geometry.group(0), '') if geometry else raw
        cls.css = re.sub(r'/\*.*?\*/', '', raw, flags=re.S)       # rules, not comments

    def test_timeline_geometry_stays_on_the_timeline(self):
        self.assertIn('::before', self.geometry)
        for _, sel in re.findall(r'(^|})\s*([^{}]+)\{', self.geometry):
            for part in sel.split(','):
                self.assertTrue(part.strip().startswith('.timeline .msg'), part)
        self.assertNotIn('transform', self.geometry)        # one rounding for dot and segment

    def test_liquid_glass_stays_in_the_chat_column(self):
        self.assertIn('backdrop-filter', self.glass)
        selectors = re.findall(r'(^|})\s*([^{}]+)\{', self.glass)
        for _, sel in selectors:
            for part in sel.split(','):
                self.assertTrue(part.strip().startswith('.pane.chat'), part)
        self.assertNotIn('#', re.sub(r'rgba\([^)]*\)', '', self.glass))   # no hex: tokens or rgba of tokens

    def test_only_t3_colours_and_light_only(self):
        used = {h.upper() for h in re.findall(r'#[0-9A-Fa-f]{6}\b', self.css)}
        self.assertEqual(sorted(used - self.TOKENS), [])
        self.assertNotIn('prefers-color-scheme', self.css)

    def test_type_scale(self):
        sizes = set(re.findall(r'font-size:\s*([0-9.]+)px', self.css)) | \
            set(re.findall(r'--fs-[a-z]+:\s*([0-9.]+)px', self.css))
        self.assertLessEqual(sizes, {'19', '15', '13', '11.5', '11', '12.5'})
        self.assertLessEqual(set(re.findall(r'font-weight:\s*(\d+)', self.css)), {'400', '600'})

    def test_spacing_in_steps_of_four(self):
        bad = []
        for prop, value in re.findall(r'(?<![-\w])(padding|margin|gap|top|left|right|bottom|inset)'
                                      r'(?:-(?:top|left|right|bottom))?:\s*([^;]+);', self.css):
            bad += ['{}:{}'.format(prop, n) for n in re.findall(r'(-?\d+(?:\.\d+)?)px', value)
                    if float(n) % 4]
        self.assertEqual(bad, [])

    def test_no_effects_and_t3_radii(self):
        for effect in ('box-shadow', 'transition', 'gradient', 'filter:'):
            self.assertNotIn(effect, self.css)
        radii = set(re.findall(r'border-radius:\s*([^;]+);', self.css))
        self.assertLessEqual(radii, {'var(--r-window)', 'var(--r-control)', 'var(--r-pill)', '18px', '50%'})
        self.assertIn('--h-row: 26px', self.css)
        self.assertIn('outline: 1px solid var(--ink)', self.css)


def _box(part_id, x0, y0, x1, y1, z0, z1, **extra):
    pts = [[x0, y0, 0], [x1, y0, 0], [x1, y1, 0], [x0, y1, 0]]
    geom = {'id': part_id, 'type': 'Extrusion', 'material': 'Oak',
            'profile': [{'type': 'Line', 'start': a, 'end': b} for a, b in zip(pts, pts[1:] + pts[:1])],
            'extrusion_start': z0, 'extrusion_end': z1}
    geom.update(extra)
    return geom


def _rod(part_id, start, end, radius, **extra):
    geom = {'id': part_id, 'type': 'Cylinder', 'material': 'Oak', 'start': start, 'end': end,
            'radius': radius}
    geom.update(extra)
    return geom


class MeaningTests(unittest.TestCase):
    """Concept, roles, hollow parts, joints and the spatial check (docs §5.3)."""

    def family(self, *geometry, **extra):
        s = copy.deepcopy(GOOD)
        s['geometry'] = list(geometry)
        s.pop('joints', None)
        s.update(extra)
        return s

    def found(self, family, prefix):
        from FamilyGen.preview_mesh import build_preview
        return [w for w in build_preview(family).warnings if w.startswith(prefix)]

    # ── schema ───────────────────────────────────────────────────────────
    def test_example_declares_its_meaning_and_passes_every_check(self):
        from FamilyGen.spatial_check import PREFIXES
        from FamilyGen.preview_mesh import build_preview
        self.assertEqual(fs.validate_ai_schema(GOOD, 'Furniture'), [])
        self.assertTrue(GOOD['concept']['what'] and GOOD['joints'])
        self.assertEqual([w for w in build_preview(GOOD).warnings if w.startswith(PREFIXES)], [])

    def test_ai_families_must_state_their_concept(self):
        bare = copy.deepcopy(GOOD)
        del bare['concept']
        self.assertEqual(fs.validate_family_schema(bare)[0], [])            # hand-written JSON still loads
        self.assertTrue(any('$.concept' in e for e in fs.validate_ai_schema(bare, 'Furniture')))
        bad = dict(GOOD, concept={'use': 'x'})
        self.assertTrue(any('$.concept.what' in e for e in fs.validate_family_schema(bad)[0]))

    def test_role_wall_and_joint_rules(self):
        top = _box('Top', -300, -300, 300, 300, 700, 730)
        errors = fs.validate_family_schema(self.family(dict(top, role='shelf')))[0]
        self.assertTrue(any('.role' in e for e in errors))
        errors = fs.validate_family_schema(self.family(dict(top, wall=3)))[0]
        self.assertTrue(any('.wall' in e and 'Cylinder and Tube' in e for e in errors))
        pipe = _rod('Pipe', [0, 0, 0], [0, 0, 700], 50, wall=50)
        self.assertTrue(any('thinner than the radius' in e for e in fs.validate_family_schema(self.family(pipe))[0]))
        stale = self.family(top, joints=[{'a': 'Top', 'b': 'Gone', 'at': [0, 0, 700]}])
        errors, warnings = fs.validate_family_schema(stale)
        self.assertEqual(errors, [])
        self.assertTrue(any("'Gone'" in w for w in warnings))
        self.assertTrue(any("'Gone'" in w for w in self.found(stale, 'Joint')))
        same = self.family(top, joints=[{'a': 'Top', 'b': 'Top', 'at': [0, 0, 700]}])
        self.assertTrue(any('two different parts' in e for e in fs.validate_family_schema(same)[0]))

    def test_prompt_contract_and_category_guides_teach_the_meaning(self):
        prompt = fs.build_system_prompt('Auto', delivery='tool')
        for block in ('MEANING', 'SOLID OR HOLLOW', 'JOINTS', 'SPACE', '"concept" (required)', '"wall"'):
            self.assertIn(block, prompt)
        rules = fs.schema_contract()['meaning_rules']
        self.assertEqual(rules['roles']['values'], list(fs.ROLES))
        folder = os.path.join(REPO, 'T3Lab.tab', 'Standards & Families.panel', 'FamiGen.pushbutton', 'prompts')
        for name in os.listdir(folder):
            with open(os.path.join(folder, name), encoding='utf-8') as handle:
                self.assertIn('## MEANING, SOLID / HOLLOW, JOINTS', handle.read(), name)
        from FamilyGen import claude_designer as designer
        self.assertIn('THINK IN THIS ORDER', designer._WORKFLOW)
        self.assertIn('connections', designer.ANALYSIS_TOOL['input_schema']['properties'])

    # ── hollow ───────────────────────────────────────────────────────────
    def test_hollow_tube_keeps_its_bore_through_every_piece(self):
        from FamilyGen.preview_mesh import build_preview
        tube = {'id': 'Pipe', 'type': 'Tube', 'material': 'Oak', 'role': 'conduit', 'radius': 60, 'wall': 5,
                'points': [[0, 0, 0], [0, 0, 600], [500, 0, 600]], 'smooth': False}
        parts = fs.expand_schema(self.family(tube))['geometry']
        self.assertTrue(all(p['wall'] == 5 for p in parts if p['type'] == 'Cylinder'))
        ball = [p for p in parts if p['type'] == 'Revolution'][0]
        self.assertEqual(len(ball['profile']), 4)                 # a shell: outer arc, inner arc, two lines
        self.assertNotIn('wall', ball)
        solid = build_preview(self.family(_rod('Rod', [0, 0, 0], [0, 0, 600], 60)))
        hollow = build_preview(self.family(_rod('Rod', [0, 0, 0], [0, 0, 600], 60, wall=5)))
        self.assertGreater(len(hollow.meshes[0].triangles), len(solid.meshes[0].triangles))
        self.assertEqual(self.found(self.family(tube), 'Hollow'), [])

    def test_containers_and_conduits_left_solid_are_reported(self):
        base = _box('Base', -400, -400, 400, 400, 0, 20)
        lump = _box('Basin', -300, -300, 300, 300, 20, 400, role='container')
        self.assertTrue(any("'Basin'" in w for w in self.found(self.family(base, lump), 'Hollow')))
        cavity = _box('Cavity', -280, -280, 280, 280, 60, 420, is_solid=False)
        cavity.pop('material')
        self.assertEqual(self.found(self.family(base, lump, cavity), 'Hollow'), [])
        rod = _rod('Drain', [0, 0, 20], [0, 0, 300], 40, role='conduit')
        self.assertTrue(any("'Drain'" in w for w in self.found(self.family(base, rod), 'Hollow')))
        self.assertEqual(self.found(self.family(base, dict(rod, wall=3)), 'Hollow'), [])
        bowl = {'id': 'Bowl', 'type': 'Revolution', 'material': 'Oak', 'role': 'container', 'sketch_plane_y': 0,
                'axis_start': [0, 0, 0], 'axis_end': [0, 0, 100],
                'profile': [{'type': 'Line', 'start': a, 'end': b} for a, b in zip(
                    [[0, 0, 20], [300, 0, 20], [300, 0, 300], [280, 0, 300], [280, 0, 40], [0, 0, 40]],
                    [[300, 0, 20], [300, 0, 300], [280, 0, 300], [280, 0, 40], [0, 0, 40], [0, 0, 20]])]}
        self.assertEqual(self.found(self.family(base, bowl), 'Hollow'), [])

    # ── space ────────────────────────────────────────────────────────────
    def test_contact_is_surface_to_surface_not_bounding_boxes(self):
        base = _box('Base', -500, -500, 500, 500, 0, 20)
        post = _rod('Post', [-400, 0, 20], [-400, 0, 900], 30)
        # A diagonal rail whose BOX overlaps the post, but which hangs 150 mm away from it.
        rail = _rod('Rail', [-250, 0, 300], [400, 0, 900], 20)
        self.assertTrue(any("'Rail'" in w for w in self.found(self.family(base, post, rail), 'Not connected')))
        fixed = _rod('Rail', [-400, 0, 880], [400, 0, 880], 20)
        self.assertEqual(self.found(self.family(base, post, fixed), 'Not connected'), [])

    def test_a_part_set_into_another_is_connected(self):
        wall = _box('Wall', -500, 0, 500, 200, 0, 2000)
        glass = _box('Glass', -200, 90, 200, 110, 800, 1500)         # inside the wall, 90 mm from any face
        self.assertEqual(self.found(self.family(wall, glass), 'Not connected'), [])
        crossed = _box('Joist', -100, -600, 100, 600, 2000, 2100)    # lies across the top: no vertex near
        self.assertEqual(self.found(self.family(wall, crossed), 'Not connected'), [])

    def test_declared_joints_are_checked_against_the_geometry(self):
        deck = _box('Deck', 0, 0, 2000, 2000, 1000, 1100, role='surface')
        post = _rod('Post', [100, 100, 0], [100, 100, 1000], 60, role='support')
        slide = _rod('Slide', [2600, 1000, 1100], [4200, 1000, 0], 250, role='access', wall=10)
        missed = self.family(deck, post, slide, joints=[
            {'a': 'Slide', 'b': 'Deck', 'at': [2000, 1000, 1100]},
            {'a': 'Slide', 'b': 'ground', 'at': [4200, 1000, 0]}])
        found = self.found(missed, 'Joint')
        self.assertEqual(len(found), 1)
        self.assertIn("'Slide' is", found[0])
        self.assertIn("with 'Deck' at (2000, 1000, 1100)", found[0])
        fixed = copy.deepcopy(missed)
        fixed['geometry'][2]['start'] = [1990, 1000, 1340]
        self.assertEqual(self.found(fixed, 'Joint'), [])
        self.assertEqual(self.found(fixed, 'Access'), [])
        unlanded = dict(fixed, joints=fixed['joints'][:1])
        self.assertTrue(any("'Slide'" in w for w in self.found(unlanded, 'Access')))

    def test_bores_and_headroom_must_stay_clear(self):
        ground = _box('Ground', -3000, -3000, 3000, 3000, -100, 0)
        tube = _rod('Tube', [-2000, 0, 400], [2000, 0, 400], 400, role='conduit', wall=10)
        post = _rod('Post', [500, 0, 0], [500, 0, 1500], 60, role='support')
        found = self.found(self.family(ground, tube, post), 'Crossing')
        self.assertTrue(found and "'Post' is in the way of 'Tube'" in found[0])
        beside = _rod('Post', [500, 700, 0], [500, 700, 1500], 60, role='support')
        self.assertEqual(self.found(self.family(ground, tube, beside), 'Crossing'), [])
        joined = self.family(ground, tube, post, joints=[{'a': 'Tube', 'b': 'Post', 'at': [500, 0, 400]}])
        self.assertEqual(self.found(joined, 'Crossing'), [])
        # A pipe through a wall: reported, unless a void opens the way.
        wall = _box('Wall', -100, -1500, 100, 1500, 0, 2000)
        self.assertTrue(self.found(self.family(ground, tube, wall), 'Crossing'))
        hole = _rod('Opening', [-150, 0, 400], [150, 0, 400], 420, is_solid=False)
        hole.pop('material')
        self.assertEqual(self.found(self.family(ground, tube, wall, hole), 'Crossing'), [])
        # A slab 600 above a stair tread leaves no room to walk.
        tread = _box('Tread', 0, 0, 900, 300, 0, 40, role='access')
        slab = _box('Slab', -200, -200, 1100, 500, 640, 760)
        legs = _rod('Leg', [-150, -150, 0], [-150, -150, 640], 40)
        found = self.found(self.family(ground, tread, slab, legs), 'Crossing')
        self.assertTrue(found and "'Slab' is in the way of 'Tread'" in found[0])

    def test_spatial_findings_go_back_to_claude(self):
        from FamilyGen import claude_designer as designer
        family = self.family(_box('Deck', 0, 0, 2000, 2000, 1000, 1100, role='surface'),
                             _rod('Post', [100, 100, 0], [100, 100, 1000], 60, role='support'),
                             _rod('Slide', [2600, 1000, 1100], [4200, 1000, 0], 250, role='access', wall=10),
                             joints=[{'a': 'Slide', 'b': 'Deck', 'at': [2000, 1000, 1100]},
                                     {'a': 'Slide', 'b': 'ground', 'at': [4200, 1000, 0]}])
        session = designer.ClaudeFamilySession('Furniture', 'claude-opus-5-5')
        session.begin_turn('Deck with a slide')
        first = session.handle_http(200, submit(family, 'a'))
        self.assertEqual(first.kind, 'continue')
        text = first.request['messages'][-1]['content'][0]['content']
        self.assertIn('Joint:', text)
        self.assertIn('fix the geometry', text)

    def test_summary_and_page_show_the_concept(self):
        info = fs.schema_summary(GOOD)
        self.assertEqual(info['concept']['what'], GOOD['concept']['what'])
        self.assertEqual(info['joints'], 3)
        with open(os.path.join(STUDIO, 'studio.js'), encoding='utf-8') as handle:
            script = handle.read()
        self.assertIn("'Concept'", script)
        self.assertIn("'Joints checked'", script)


class HostChromeTests(unittest.TestCase):
    """One title bar: the host hides pyRevit's, or the page hides its own."""

    def test_page_drops_its_title_bar_when_the_host_keeps_one(self):
        with open(os.path.join(STUDIO, 'studio.js'), encoding='utf-8') as handle:
            script = handle.read()
        with open(os.path.join(STUDIO, 'studio.css'), encoding='utf-8') as handle:
            css = handle.read()
        self.assertIn("classList.toggle('host-chrome', m.frameless === false)", script)
        self.assertIn('.host-chrome .winctrls', css)
        self.assertIn("action: 'drag'", script)

    def test_controller_reports_frameless_and_routes_drag(self):
        revit = FakeRevit()
        posted = []
        controller = sc.StudioController(post=posted.append, http=FakeHttp(), revit=revit,
                                         claude=FakeClaude(), prefs=Prefs(), store=None,
                                         pyrevit_version='7.1.0')
        revit.frameless = False
        controller.handle({'type': 'ready'})
        init = [m for m in posted if m.get('type') == 'init'][0]
        self.assertIs(init['frameless'], False)
        calls = []
        revit.window_state = calls.append
        controller.handle({'type': 'window', 'action': 'drag'})
        self.assertEqual(calls, ['drag'])
        with open(os.path.join(LIB, 'FamilyGen', 'studio_host.py'), encoding='utf-8') as handle:
            host = handle.read()
        for needle in ('def _go_frameless', 'UseNoneWindowStyle', 'def _drag_window',
                       "prefs.get('frameless', True)"):
            self.assertIn(needle, host)


class OutputWindowTests(unittest.TestCase):
    """The Studio lives in pyRevit 7's output window, which every other T3Lab
    tool has swapped for a silent stand-in under CPython."""

    def test_the_stand_in_would_pass_every_check_and_show_nothing(self):
        import _cpython_bootstrap as boot
        fake = boot._NullOutput()
        self.assertTrue(hasattr(fake, 'on_message'))
        self.assertTrue(hasattr(fake, 'post_message'))
        self.assertIsNone(fake.open_page('<html></html>'))

    def test_the_studio_asks_pyrevit_for_the_real_window(self):
        with open(os.path.join(LIB, 'FamilyGen', 'studio_host.py'), encoding='utf-8') as handle:
            host = handle.read()
        opener = host[host.index('def open_studio('):]
        self.assertIn('output = _pyrevit_output()', opener)
        self.assertNotIn('script.get_output()', opener)
        helper = host[host.index('def _pyrevit_output('):host.index('def open_studio(')]
        self.assertIn('pyrevit_output.PyRevitOutputWindow()', helper)
        self.assertNotIn('get_output()\n', helper.split('"""')[-1])


class RibbonButtonTests(unittest.TestCase):
    """What the Assistant learns about the FamiGen button from its script."""

    def test_the_button_is_known_by_its_own_names_only(self):
        # Services/tool_discovery turns every `from GUI.<X>Dialog import` in a
        # script into keywords for that button. FamiGen borrowed a helper from
        # GUI.ManaSelectDialog and "Mana Select" then matched two tools, so the
        # Assistant opened neither.
        from Services import tool_discovery as discovery
        script = os.path.join(os.path.dirname(LIB), 'T3Lab.tab',
                              'Standards & Families.panel', 'FamiGen.pushbutton', 'script.py')
        if not os.path.isfile(script):          # t3lab_dev: one T3Lab_Dev.tab
            sys.path.insert(0, os.path.join(REPO, 'dev'))
            import tabdir
            script = tabdir.bundle_path('FamiGen.pushbutton', 'script.py')
        title, _doc, xamls = discovery._read_meta(script)
        self.assertEqual(title, 'FamiGen')
        self.assertEqual(xamls, ['FamiGen'])
        keywords = discovery._gen_keywords(title, 'FamiGen.pushbutton', xamls)
        self.assertEqual(sorted(keywords), ['fami', 'famigen', 'gen'])


def _flex_family():
    """A 1000 x 600 x 400 box whose right face follows a 'Width' dimension."""
    fam = copy.deepcopy(GOOD)
    fam['parameters'] = [{'name': 'Width', 'type': 'length', 'value': 1000},
                         {'name': 'Depth', 'type': 'length', 'value': 600},
                         {'name': 'Note', 'type': 'text', 'value': 'x'}]
    fam['types'] = [{'name': 'W1000', 'values': {}}, {'name': 'W1400', 'values': {'Width': 1400}}]
    fam['reference_planes'] = [{'name': 'Left', 'axis': 'x', 'offset': -500},
                               {'name': 'Right', 'axis': 'x', 'offset': 500}]
    fam['dimensions'] = [{'parameter': 'Width', 'from': 'Left', 'to': 'Right'}]
    fam['geometry'] = [_box('Body', -500, -300, 500, 300, 0, 400, material='Oak',
                            locks={'x_min': 'Left', 'x_max': 'Right'}),
                       _box('Loose', -100, -100, 100, 100, 400, 450, material='Oak')]
    fam.pop('joints', None)
    return fam


class OutboxTests(unittest.TestCase):
    """pyRevit drops post_message while the page loads; the outbox must not (2026-11-19)."""

    def setUp(self):
        self.Outbox = studio_page.Outbox

    def test_held_until_the_page_can_take_it_then_sent_in_order(self):
        ready, got, woke = [False], [], []

        def send(msg):
            if not ready[0]:
                return False
            got.append(msg)
            return True
        box = self.Outbox(send, wake=lambda: woke.append(1))
        box.post({'type': 'init'})
        box.post({'type': 'status'})
        self.assertEqual((got, box.pending), ([], 2))
        self.assertTrue(woke)
        ready[0] = True
        self.assertTrue(box.flush())
        self.assertEqual([m['type'] for m in got], ['init', 'status'])

    def test_a_message_that_raises_does_not_block_the_rest(self):
        got = []

        def send(msg):
            if msg.get('bad'):
                raise TypeError('not JSON')
            got.append(msg)
            return None                       # older pyRevit returned nothing: delivered
        box = self.Outbox(send)
        box.post({'bad': True})
        box.post({'type': 'ok'})
        self.assertEqual((got, box.pending), ([{'type': 'ok'}], 0))

    def test_the_host_posts_through_the_outbox(self):
        with open(os.path.join(LIB, 'FamilyGen', 'studio_host.py'), encoding='utf-8') as handle:
            host = handle.read()
        self.assertIn('post=outbox.post', host)
        self.assertNotIn('post=lambda msg: output.post_message(msg)', host)


class FlexTests(unittest.TestCase):
    """Editing a length flexes the preview the way Revit flexes the type."""

    def test_type_value_moves_the_free_plane_and_stretches_locked_parts_only(self):
        fam = _flex_family()
        text = json.dumps(fam)
        base = sc.preview_payload(text)
        wide = sc.preview_payload(text, 'W1400')
        self.assertEqual(base['size_mm'][0], 1000.0)
        self.assertEqual(wide['size_mm'][0], 1400.0)
        self.assertEqual(wide['flexed'], ['right'])
        self.assertEqual([d['value'] for d in wide['dimensions']], [1400.0])
        loose = [p for p in wide['parts'] if p['label'].startswith('Loose')][0]
        xs = loose['pos'][0::3]
        self.assertEqual((min(xs), max(xs)), (-100.0, 100.0))       # no locks: it stays

    def test_template_planes_never_move(self):
        from FamilyGen import flex
        fam = _flex_family()
        fam['dimensions'] = [{'parameter': 'Width', 'from': 'Right', 'to': 'Center (Left/Right)'}]
        moved = flex.plane_offsets(fam, {'width': 800})
        self.assertEqual(moved['center (left/right)'], (0.0, 0.0))
        self.assertEqual(moved['right'][1], 800.0)

    def test_set_param_edits_the_type_and_reflexes(self):
        ctl, posts = controller()
        ctl.current_json = json.dumps(_flex_family())
        ctl.handle({'type': 'set_param', 'param': 'width', 'type_name': 'W1000', 'value': 1200})
        schema = json.loads(of_type(posts, 'schema')[-1]['json'])
        self.assertEqual(schema['types'][0]['values'], {'Width': 1200.0})
        self.assertEqual(of_type(posts, 'preview')[-1]['size_mm'][0], 1200.0)
        self.assertEqual(of_type(posts, 'status')[-1]['tone'], 'success')

    def test_a_parameter_that_drives_nothing_says_so(self):
        ctl, posts = controller()
        ctl.current_json = json.dumps(_flex_family())
        ctl.handle({'type': 'set_param', 'param': 'Depth', 'type_name': '', 'value': 900})
        self.assertEqual(json.loads(ctl.current_json)['parameters'][1]['value'], 900.0)
        self.assertIn('drives no labelled dimension', of_type(posts, 'status')[-1]['text'])
        ctl.handle({'type': 'set_param', 'param': 'Width', 'type_name': '', 'value': -5})
        self.assertEqual(of_type(posts, 'status')[-1]['tone'], 'warning')

    def test_switching_type_reflexes(self):
        ctl, posts = controller()
        ctl.current_json = json.dumps(_flex_family())
        ctl.handle({'type': 'set_type', 'name': 'W1400'})
        self.assertEqual(of_type(posts, 'preview')[-1]['size_mm'][0], 1400.0)


class ProfileTests(unittest.TestCase):
    """Category parameter profiles and the fit score (docs/famigen-param-profiles.md)."""

    def test_every_category_has_a_profile_and_three_templates(self):
        from FamilyGen import param_profiles as pp
        self.assertTrue(pp.covers_every_category())
        counts = {}
        for t in tpl.templates():
            counts[t['category']] = counts.get(t['category'], 0) + 1
            self.assertIn('Make it parametric', t['prompt'])
        self.assertEqual(sorted(c for c in fs.SUPPORTED_CATEGORIES if counts.get(c, 0) < 3), [])
        self.assertEqual(len(set(t['id'] for t in tpl.templates())), len(tpl.templates()))

    def test_driven_present_missing_and_score(self):
        from FamilyGen import param_profiles as pp
        fam = _flex_family()
        fit = pp.evaluate(fam)
        states = dict((r['name'], r['state']) for r in fit['rows'] if r['required'])
        self.assertEqual(states, {'Width': 'driven', 'Depth': 'present', 'Height': 'missing'})
        self.assertEqual(fit['score'], 50)
        brief = pp.brief(fam)
        self.assertIn('Depth', brief)
        self.assertIn('Height', brief)
        self.assertNotIn('Width (', brief)

    def test_make_parametric_sends_the_brief_to_claude(self):
        http = FakeHttp((200, submit(GOOD)))
        ctl, posts = controller(http=http)
        ctl.current_json = json.dumps(_flex_family())
        ctl.handle({'type': 'make_parametric'})
        self.assertIn('Height', of_type(posts, 'chat')[0]['text'])
        self.assertTrue(http.bodies)


class ProgressTests(unittest.TestCase):
    def test_a_send_walks_the_strip(self):
        http = FakeHttp((200, submit(GOOD)))
        ctl, posts = controller(http=http)
        ctl.handle({'type': 'ready'})
        self.assertEqual(of_type(posts, 'init')[0]['progress']['steps'][0]['state'], 'todo')
        ctl.handle({'type': 'send', 'text': 'Side table', 'json': ''})
        last = dict((s['id'], s['state']) for s in of_type(posts, 'progress')[-1]['steps'])
        self.assertEqual(last, {'brief': 'done', 'draft': 'done', 'check': 'done', 'review': 'wait',
                                'compare': 'skip', 'create': 'todo'})
        ctl.handle({'type': 'create', 'json': ctl.current_json, 'load': False})
        last = dict((s['id'], s['state']) for s in of_type(posts, 'progress')[-1]['steps'])
        self.assertIn(last['create'], ('done', 'warn'))
        self.assertEqual(last['review'], 'done')


class McpReceiptTests(unittest.TestCase):
    def setUp(self):
        os.environ['APPDATA'] = tempfile.mkdtemp(prefix='famigen_receipt_')

    def test_sent_read_answered(self):
        ctl, posts = controller(engine='mcp')
        ctl.handle({'type': 'send', 'text': 'Make it 2 m long', 'json': ''})
        mine = of_type(posts, 'chat')[0]
        self.assertTrue(mine['mcp_id'])
        self.assertIn('not read', mine['mcp_state'])
        ctl.tick()
        self.assertEqual(of_type(posts, 'mcp_status'), [])
        studio_link.take_unread()                    # the AI app read it (famigen_get_studio_state)
        ctl.tick()
        self.assertEqual(of_type(posts, 'mcp_status')[0]['state'], 'read')
        self.assertEqual(ctl._mcp_waiting, [])


class ChatInputTests(unittest.TestCase):
    def setUp(self):
        with open(os.path.join(STUDIO, 'studio.js'), encoding='utf-8') as handle:
            self.js = handle.read()

    def test_enter_sends_shift_enter_breaks_and_ime_is_respected(self):
        self.assertIn("e.key === 'Enter' && !e.shiftKey", self.js)
        self.assertIn('e.isComposing', self.js)
        self.assertNotIn("(e.ctrlKey || e.metaKey)) { e.preventDefault(); send(); }", self.js)

    def test_ready_is_repeated_until_init_and_send_never_locks_before_it(self):
        self.assertIn('if (state.inited) return;', self.js)
        self.assertIn('return !state.inited ||', self.js)


class McpFirstTests(unittest.TestCase):
    """Auto puts MCP first, then the API key; the app listens instead of "continue"."""

    def setUp(self):
        os.environ['APPDATA'] = tempfile.mkdtemp(prefix='famigen_first_')

    def test_auto_order(self):
        ctl, _ = controller()
        self.assertEqual(ctl.engine(), 'claude')                 # key, no app
        studio_link.note_activity('claude-ai', 'list_open_documents')
        ctl.tick()
        self.assertEqual(ctl.engine(), 'mcp')                    # any tool call: app connected
        ctl2, _ = controller(claude=FakeClaude(key=None))
        self.assertEqual(ctl2.engine(), 'mcp')                   # no key: MCP rather than nothing
        ctl3, _ = controller(engine='claude')
        studio_link.note_activity('claude-ai', 'famigen_get_schema')
        ctl3.tick()
        self.assertEqual(ctl3.engine(), 'claude')                # an explicit choice wins

    def test_wait_returns_messages_or_times_out(self):
        t = [0.0]
        clock = lambda: t[0]

        def sleep(d):
            t[0] += d
        self.assertEqual(studio_link.wait_unread(3, clock=clock, sleep=sleep), [])
        self.assertGreaterEqual(t[0], 3)
        studio_link.post_message('Make it 2 m long')
        got = studio_link.wait_unread(30, clock=clock, sleep=sleep)
        self.assertEqual([m['text'] for m in got], ['Make it 2 m long'])
        self.assertEqual(studio_link.unread_count(), 0)

    def test_listener_and_listen_command(self):
        self.assertIsNone(studio_link.listener())
        studio_link.note_listening('claude-ai')
        self.assertEqual(studio_link.listener(), 'Claude Desktop')
        self.assertIn('famigen_wait_for_message', studio_link.LISTEN_COMMAND)
        self.assertIn('past chats', studio_link.LISTEN_COMMAND)
        ctl, posts = controller()
        ctl.tick()
        self.assertTrue(of_type(posts, 'engine')[-1]['listening'])
        ctl.handle({'type': 'send', 'text': 'Taller', 'json': ''})
        self.assertFalse([p for p in of_type(posts, 'chat') if p.get('listen')])   # no nagging

    def test_server_waits_off_the_revit_thread(self):
        with open(os.path.join(LIB, 'core', 'server.py'), encoding='utf-8') as handle:
            server = handle.read()
        head = server[server.index('    def _execute_tool(self'):server.index('if self._external_event:', server.index('    def _execute_tool(self'))]
        self.assertIn("if tool_name == 'famigen_wait_for_message':", head)
        self.assertNotIn('continue in FamiGen', server)

    def test_mcp_fallback_moves_the_message_to_claude(self):
        http = FakeHttp((200, submit(GOOD)))
        ctl, posts = controller(http=http)
        studio_link.note_activity('claude-ai', 'list_open_documents')
        ctl.tick()
        ctl.handle({'type': 'send', 'text': 'Side table', 'json': ''})
        mcp_id = [p for p in of_type(posts, 'chat') if p.get('mcp_id')][0]['mcp_id']
        ctl.handle({'type': 'mcp_fallback', 'id': mcp_id})
        self.assertTrue(http.bodies)
        self.assertEqual(studio_link.unread_count(), 0)
        self.assertEqual(of_type(posts, 'mcp_status')[-1]['state'], 'moved')


class FinishingTests(unittest.TestCase):
    """Fable review leftovers and MCP replies (2026-11-19)."""

    def setUp(self):
        os.environ['APPDATA'] = tempfile.mkdtemp(prefix='famigen_finish_')
        with open(os.path.join(STUDIO, 'studio.js'), encoding='utf-8') as handle:
            self.js = handle.read()
        with open(os.path.join(STUDIO, 'studio.html'), encoding='utf-8') as handle:
            self.html = handle.read()

    def test_mcp_reply_in_words_reaches_the_chat(self):
        ctl, posts = controller(engine='mcp')
        ctl.handle({'type': 'send', 'text': 'What size?', 'json': ''})
        studio_link.post_reply('Which width: 1800 or 2100?', 'claude-ai')
        ctl.tick()
        answer = [p for p in of_type(posts, 'chat') if p.get('role') == 'assistant'][-1]
        self.assertIn('1800', answer['text'])
        self.assertIn('Claude Desktop', answer['meta'])
        self.assertEqual(of_type(posts, 'mcp_status')[-1]['state'], 'answered')
        self.assertEqual(studio_link.take_replies(), [])

    def test_category_change_keeps_the_chat(self):
        ctl, posts = controller()
        ctl.session = object()                      # a conversation is running
        ctl.handle({'type': 'set_category', 'category': 'Door'})
        self.assertIsNone(ctl.session)
        self.assertEqual(of_type(posts, 'chat_reset'), [])
        self.assertIn('fresh conversation', of_type(posts, 'chat')[-1]['text'])

    def test_edit_after_create_reopens_create(self):
        ctl, posts = controller()
        ctl.current_json = json.dumps(_flex_family())
        ctl._steps = {'create': 'done', 'review': 'done'}
        ctl.handle({'type': 'set_param', 'param': 'Width', 'type_name': 'W1000', 'value': 1100})
        last = dict((s['id'], s['state']) for s in of_type(posts, 'progress')[-1]['steps'])
        self.assertEqual((last['create'], last['review']), ('todo', 'wait'))

    def test_office_profiles_override(self):
        from FamilyGen import param_profiles as pp
        path = os.path.join(os.environ['APPDATA'], 'T3LabAI')
        os.makedirs(path)
        with open(os.path.join(path, 'famigen_profiles.json'), 'w', encoding='utf-8') as handle:
            json.dump({'Furniture': {'required': [['Width', 'x', 'w'], ['Seat Height', 'z', 's']]}}, handle)
        names = [r['name'] for r in pp.evaluate(_flex_family())['rows']]
        self.assertEqual(names, ['Width', 'Seat Height'])
        with open(os.path.join(path, 'famigen_profiles.json'), 'w', encoding='utf-8') as handle:
            handle.write('{broken')
        self.assertIn('Depth', [r['name'] for r in pp.evaluate(_flex_family())['rows']])

    def test_no_emoji_or_text_glyph_icons(self):
        for glyph in (u'📎', u'📝', u'📄', u'▸', u'▾', u'✕', u'⌕', u"'×'"):
            self.assertNotIn(glyph, self.js + self.html)

    def test_narrow_window_keeps_the_inspector_reachable(self):
        self.assertIn('id="btn-inspect"', self.html)
        self.assertIn("classList.toggle('show-inspect'", self.js)

    def test_drawings_hide_the_size_along_the_view(self):
        self.assertIn("{ front: 'y', right: 'x', top: 'z' }", self.js)


SP_TEXT = (u'# This is a Revit shared parameter file.\n'
           u'# Do not edit manually.\n'
           u'*META\tVERSION\tMINVERSION\nMETA\t2\t1\n'
           u'*GROUP\tID\tNAME\nGROUP\t1\tT3 Dimensions\n'
           u'*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\tGROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\tHIDEWHENNOVALUE\n'
           u'PARAM\t0a1b2c3d-0000-4000-8000-00000000abcd\tWidth\tLENGTH\t\t1\t1\tOverall width\t1\t0\n'
           u'PARAM\t1a1b2c3d-0000-4000-8000-00000000abcd\tSill Height\tautodesk.spec.aec:length-2.0.0\t\t1\t1\t\t1\t0\n'
           u'PARAM\t2a1b2c3d-0000-4000-8000-00000000abcd\tNote\tTEXT\t\t1\t1\t\t1\t0\n')


class SharedParamTests(unittest.TestCase):
    """Instance / type scopes and Revit shared parameters (docs/famigen-param-profiles.md)."""

    def setUp(self):
        from FamilyGen import shared_params, param_profiles
        self.sp, self.pp = shared_params, param_profiles
        self.dir = tempfile.mkdtemp(prefix='famigen_sp_')

    def write(self, encoding):
        path = os.path.join(self.dir, 'sp_{}.txt'.format(encoding))
        with open(path, 'w', encoding=encoding, newline='\r\n') as handle:
            handle.write(SP_TEXT)
        return path

    def test_reads_utf16_and_utf8_files(self):
        for encoding in ('utf-16', 'utf-8'):
            defs = self.sp.read(self.write(encoding))
            self.assertEqual([d['name'] for d in defs], ['Width', 'Sill Height', 'Note'])
            self.assertEqual([d['kind'] for d in defs], ['length', 'length', 'text'])
            self.assertEqual(defs[0]['group'], 'T3 Dimensions')
        self.assertEqual(self.sp.read(os.path.join(self.dir, 'missing.txt')), [])

    def test_find_by_guid_then_name_with_a_matching_kind(self):
        defs = self.sp.parse(SP_TEXT)
        self.assertEqual(self.sp.find(defs, 'x', '1A1B2C3D-0000-4000-8000-00000000ABCD')['name'], 'Sill Height')
        self.assertEqual(self.sp.find(defs, 'width', kind='length')['guid'][:8], '0a1b2c3d')
        self.assertIsNone(self.sp.find(defs, 'Width', kind='text'))

    def window(self):
        return {'family_category': 'Window', 'family_name': 'W', 'parameters': [
            {'name': 'Width', 'type': 'length', 'value': 900},
            {'name': 'Sill Height', 'type': 'length', 'value': 900},
            {'name': 'Note', 'type': 'text', 'value': 'x'}]}

    def test_scope_grading_and_apply_standards(self):
        defs = self.sp.parse(SP_TEXT)
        fit = self.pp.evaluate(self.window(), defs)
        sill = [r for r in fit['rows'] if r['name'] == 'Sill Height'][0]
        self.assertEqual((sill['scope'], sill['scope_ok'], sill['in_file']), ('instance', False, True))
        self.assertEqual((fit['scope_issues'], fit['shareable']), (1, 2))
        fixed, changes = self.pp.apply_standards(self.window(), defs)
        params = dict((p['name'], p) for p in fixed['parameters'])
        self.assertTrue(params['Sill Height']['instance'])
        self.assertFalse(params['Width'].get('instance', False))
        self.assertEqual(params['Width']['guid'], '0a1b2c3d-0000-4000-8000-00000000abcd')
        self.assertTrue(params['Note']['shared'])
        self.assertEqual(len(changes), 4)
        self.assertEqual(fs.validate_family_schema(dict(GOOD, parameters=fixed['parameters']))[0], [])
        self.assertEqual(self.pp.apply_standards(fixed, defs)[1], [])            # idempotent
        self.assertNotIn('shared', self.window()['parameters'][0])                # input untouched

    def test_schema_checks_shared_and_guid(self):
        bad = dict(GOOD, parameters=[{'name': 'Width', 'type': 'length', 'value': 1,
                                      'shared': 'yes', 'guid': 'nope'}])
        errors = fs.validate_family_schema(bad)[0]
        self.assertTrue(any('.shared' in e for e in errors))
        self.assertTrue(any('.guid' in e for e in errors))

    def test_controller_applies_standards_without_ai(self):
        revit = FakeRevit()
        revit.shared_parameter_file = self.write('utf-16')
        http = FakeHttp()
        ctl, posts = controller(http=http, revit=revit)
        ctl.handle({'type': 'ready'})
        self.assertEqual(of_type(posts, 'init')[0]['shared_file']['count'], 3)
        ctl.current_json = json.dumps(dict(GOOD, family_category='Window',
                                           parameters=self.window()['parameters']))
        ctl.handle({'type': 'apply_standards'})
        schema = json.loads(of_type(posts, 'schema')[-1]['json'])
        self.assertTrue([p for p in schema['parameters'] if p['name'] == 'Sill Height'][0]['instance'])
        self.assertEqual(http.bodies, [])
        fit = of_type(posts, 'preview')[-1]['param_fit']
        self.assertEqual((fit['scope_issues'], fit['shareable']), (0, 0))
        self.assertIn('Applied the parameter standard', of_type(posts, 'chat')[-1]['text'])

    def test_builder_adds_shared_parameters_version_safe(self):
        with open(os.path.join(LIB, 'FamilyGen', 'builder.py'), encoding='utf-8') as handle:
            builder = handle.read()
        with open(os.path.join(LIB, 'Snippets', '_compat.py'), encoding='utf-8') as handle:
            compat = handle.read()
        self.assertIn('add_shared_family_parameter(self.fm, definition', builder)
        self.assertIn('OpenSharedParameterFile()', compat)
        helper = compat[compat.index('def add_shared_family_parameter'):compat.index('def shared_definitions')]
        self.assertLess(helper.index('_DB.GroupTypeId'), helper.index("'BuiltInParameterGroup')"))
        self.assertIn("'shared_parameters'", builder)


class IronPythonSourceTests(unittest.TestCase):
    """The MCP server runs these modules in IronPython 2, which refuses a non-ASCII
    source file without a coding line (famigen_wait_for_message failed on
    family_schema.py's box-drawing comments, 2026-11-20)."""

    def test_non_ascii_sources_declare_utf8(self):
        bad = []
        for root, _dirs, files in os.walk(LIB):
            for name in files:
                if not name.endswith('.py'):
                    continue
                path = os.path.join(root, name)
                with open(path, 'rb') as handle:
                    raw = handle.read()
                if any(b > 127 for b in raw) and not any(
                        re.search(rb'coding[:=]\s*([-\w.]+)', line) for line in raw.split(b'\n')[:2]):
                    bad.append(os.path.relpath(path, LIB))
        self.assertEqual(bad, [])


if __name__ == '__main__':
    unittest.main()
