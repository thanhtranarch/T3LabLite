# -*- coding: utf-8 -*-
"""FamiGen Studio's direct Claude session (lib/FamilyGen/claude_designer.py)
and the schema v2 family types (lib/Intelligence/family_schema.py).

No Revit, no network: the Messages API is a scripted fake. Checks the request
shape, the validate -> repair loop, the append-only transcript, the optional
feature fallbacks on a 400, refusals and the output limit. docs/famigen-studio-spec.md.
"""
import copy
import json
import os
import sys
import tempfile
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'lib'))

from FamilyGen import claude_designer as cd          # noqa: E402
from Intelligence import family_schema as fs          # noqa: E402

GOOD = copy.deepcopy(fs.EXAMPLE_SCHEMA)


def tool_response(family, summary='Built it.', assumptions=None, call_id='toolu_1',
                  stop='tool_use', thinking=True, usage=None):
    content = []
    if thinking:
        content.append({'type': 'thinking', 'thinking': '', 'signature': 'sig-abc'})
    content.append({'type': 'tool_use', 'id': call_id, 'name': 'submit_family',
                    'input': {'family': family, 'summary': summary,
                              'assumptions': assumptions or []}})
    return {'content': content, 'stop_reason': stop,
            'usage': usage or {'input_tokens': 100, 'output_tokens': 50,
                               'cache_read_input_tokens': 10,
                               'cache_creation_input_tokens': 5}}


def text_response(text, stop='end_turn'):
    return {'content': [{'type': 'text', 'text': text}], 'stop_reason': stop,
            'usage': {'input_tokens': 20, 'output_tokens': 10}}


class FakeApi(object):
    """Returns scripted (status, body) pairs and records every request body."""

    def __init__(self, *replies):
        self.replies = list(replies)
        self.bodies = []

    def __call__(self, body):
        self.bodies.append(json.loads(json.dumps(body)))   # snapshot, like the wire
        status, payload = self.replies.pop(0)
        return status, payload if isinstance(payload, str) else json.dumps(payload)


def session(**kw):
    return cd.ClaudeFamilySession('Furniture', 'claude-opus-5-5', **kw)


class RequestShapeTests(unittest.TestCase):
    def test_first_request(self):
        s = session()
        body = s.begin_turn('Side table 500 mm')
        self.assertEqual(body['model'], 'claude-opus-5-5')
        self.assertEqual(body['tool_choice'], {'type': 'auto'})   # forced choice 400s on 5.5
        self.assertEqual([t['name'] for t in body['tools']], ['submit_family'])
        self.assertEqual(body['output_config'], {'effort': 'high'})
        self.assertEqual(body['fallbacks'], 'default')
        self.assertEqual(body['cache_control'], {'type': 'ephemeral'})
        self.assertNotIn('thinking', body)            # Opus 5.5: omit, never disable
        self.assertNotIn('temperature', body)
        self.assertIn('submit_family', body['system'][0]['text'])
        self.assertIn('"Furniture"', body['system'][0]['text'])
        self.assertEqual(body['messages'][0]['role'], 'user')
        h = s.headers('KEY')
        self.assertEqual(h['x-api-key'], 'KEY')
        self.assertEqual(h['anthropic-beta'], cd.FALLBACK_BETA)

    def test_rejects_bad_setup(self):
        with self.assertRaises(ValueError):
            cd.ClaudeFamilySession('Spaceship', 'claude-opus-5-5')
        with self.assertRaises(ValueError):
            cd.ClaudeFamilySession('Furniture', None)
        with self.assertRaises(ValueError):
            session().begin_turn('   ')

    def test_attachments_become_content_blocks(self):
        folder = tempfile.mkdtemp()
        png = os.path.join(folder, 'chair.png')
        pdf = os.path.join(folder, 'spec.pdf')
        with open(png, 'wb') as f:
            f.write(b'\x89PNG\r\n\x1a\nxx')
        with open(pdf, 'wb') as f:
            f.write(b'%PDF-1.4 x')
        atts = [cd.load_attachment(png), cd.load_attachment(pdf)]
        body = session().begin_turn('Model this chair', attachments=atts)
        kinds = [b['type'] for b in body['messages'][0]['content']]
        self.assertEqual(kinds, ['image', 'document', 'text'])
        self.assertEqual(body['messages'][0]['content'][0]['source']['media_type'], 'image/png')
        self.assertIn('chair.png', body['messages'][0]['content'][2]['text'])
        bad = os.path.join(folder, 'model.dwg')
        open(bad, 'wb').close()
        with self.assertRaises(ValueError):
            cd.load_attachment(bad)


class TurnTests(unittest.TestCase):
    def test_valid_family_first_try(self):
        api = FakeApi((200, tool_response(GOOD, assumptions=['Height 550 mm'])))
        s = session()
        out = cd.run_turn(s, api, 'Side table')
        self.assertEqual(out.kind, 'schema')
        self.assertEqual(out.errors, [])
        self.assertEqual(out.assumptions, ['Height 550 mm'])
        self.assertEqual(s.last_schema['family_name'], 'Side Table 500')
        self.assertEqual(s.usage['requests'], 1)
        self.assertEqual(s.usage['output_tokens'], 50)
        self.assertIsNotNone(s.usage_summary()['usd'])

    def test_repair_round_then_valid(self):
        broken = copy.deepcopy(GOOD)
        broken['geometry'][0]['material'] = 'Pine'
        api = FakeApi((200, tool_response(broken, call_id='toolu_a')),
                      (200, tool_response(GOOD, call_id='toolu_b')))
        s = session()
        out = cd.run_turn(s, api, 'Side table')
        self.assertEqual(out.kind, 'schema')
        self.assertEqual(out.errors, [])
        self.assertEqual(out.repairs, 1)
        repair_msg = api.bodies[1]['messages'][-1]
        self.assertEqual(repair_msg['role'], 'user')
        block = repair_msg['content'][0]
        self.assertEqual(block['type'], 'tool_result')
        self.assertEqual(block['tool_use_id'], 'toolu_a')
        self.assertTrue(block['is_error'])
        self.assertIn("unknown material 'Pine'", block['content'])

    def test_gives_up_after_max_repairs_with_a_draft(self):
        broken = copy.deepcopy(GOOD)
        broken['family_category'] = 'Casework'
        api = FakeApi(*[(200, tool_response(broken, call_id='t%d' % i)) for i in range(3)])
        out = cd.run_turn(session(), api, 'Side table')
        self.assertEqual(out.kind, 'schema')
        self.assertTrue(out.errors)
        self.assertEqual(out.repairs, cd.MAX_REPAIRS)
        self.assertEqual(len(api.bodies), cd.MAX_REPAIRS + 1)

    def test_text_reply_is_a_question(self):
        api = FakeApi((200, text_response('Which width do you need?')))
        out = cd.run_turn(session(), api, 'A thing')
        self.assertEqual(out.kind, 'reply')
        self.assertIn('width', out.text)

    def test_refusal(self):
        api = FakeApi((200, {'content': [], 'stop_reason': 'refusal',
                             'stop_details': {'category': 'cyber'}}))
        out = cd.run_turn(session(), api, 'x')
        self.assertEqual(out.kind, 'error')
        self.assertIn('declined', out.message)

    def test_output_limit_asks_for_a_compact_family(self):
        cut = tool_response({'family_name': 'x'}, stop='max_tokens', call_id='cut')
        api = FakeApi((200, cut), (200, tool_response(GOOD, call_id='ok')))
        out = cd.run_turn(session(), api, 'Huge kitchen')
        self.assertEqual(out.kind, 'schema')
        last = api.bodies[1]['messages'][-1]['content']
        self.assertEqual(last[0]['tool_use_id'], 'cut')
        self.assertIn('output limit', last[1]['text'])

    def test_400_drops_the_refused_feature_and_retries(self):
        err = {'type': 'error', 'error': {'type': 'invalid_request_error',
                                          'message': 'fallbacks: Extra inputs are not permitted'}}
        api = FakeApi((400, err), (200, tool_response(GOOD)))
        s = session()
        out = cd.run_turn(s, api, 'Side table')
        self.assertEqual(out.kind, 'schema')
        self.assertNotIn('fallbacks', api.bodies[1])
        self.assertNotIn('anthropic-beta', s.headers('k'))
        self.assertIn('output_config', api.bodies[1])

    def test_http_errors_are_plain_words(self):
        for status, expect in ((401, 'API key'), (429, 'rate limit'), (529, 'busy'),
                               (0, 'could not be reached')):
            api = FakeApi((status, '{"error":{"type":"x","message":"m"}}'))
            out = cd.run_turn(session(), api, 'x')
            self.assertEqual(out.kind, 'error')
            self.assertIn(expect, out.message)


class TranscriptTests(unittest.TestCase):
    def test_append_only_and_tool_results_first(self):
        """Each request's messages extend the previous request's messages
        unchanged (preserved thinking), and the next user turn starts with the
        tool_result owed to the accepted submit_family call."""
        api = FakeApi((200, tool_response(GOOD, call_id='first')),
                      (200, tool_response(GOOD, call_id='second')))
        s = session()
        cd.run_turn(s, api, 'Side table')
        edited = copy.deepcopy(GOOD)
        edited['family_name'] = 'Side Table 600'
        cd.run_turn(s, api, 'Make the top thicker', current_json=json.dumps(edited))
        first, second = api.bodies
        self.assertEqual(second['messages'][:len(first['messages'])], first['messages'])
        assistant = second['messages'][1]
        self.assertEqual(assistant['content'][0]['signature'], 'sig-abc')   # thinking kept
        user2 = second['messages'][2]['content']
        self.assertEqual(user2[0]['type'], 'tool_result')
        self.assertEqual(user2[0]['tool_use_id'], 'first')
        self.assertNotIn('is_error', user2[0])
        self.assertIn('Side Table 600', user2[-1]['text'])      # user's edit is sent
        self.assertIn('Make the top thicker', user2[-1]['text'])

    def test_unchanged_json_is_not_resent(self):
        api = FakeApi((200, tool_response(GOOD)), (200, text_response('ok')))
        s = session()
        cd.run_turn(s, api, 'Side table')
        cd.run_turn(s, api, 'Is it stable?', current_json=json.dumps(GOOD))
        self.assertNotIn('Current family JSON', api.bodies[1]['messages'][-1]['content'][-1]['text'])


class FamilyTypeTests(unittest.TestCase):
    def test_example_is_valid_and_has_types(self):
        errors, warnings = fs.validate_family_schema(fs.EXAMPLE_SCHEMA)
        self.assertEqual((errors, warnings), ([], []))
        self.assertEqual(fs.type_names(fs.EXAMPLE_SCHEMA), ['Oak Top', 'Steel Top'])
        self.assertEqual(fs.schema_summary(fs.EXAMPLE_SCHEMA)['types'], ['Oak Top', 'Steel Top'])

    def test_type_value_rules(self):
        bad = copy.deepcopy(GOOD)
        bad['types'] = [{'name': 'A', 'values': {'Top Material': 'Pine', 'Top Thickness': 'x',
                                                 'Designer': 3, 'Width': 900, 'Flag': True}},
                        {'name': 'a'}, 'nope']
        errors, warnings = fs.validate_family_schema(bad)
        text = '\n'.join(errors)
        self.assertIn('must name one of materials[]', text)
        self.assertIn('Top Thickness', text)
        self.assertIn('must be a string for a text parameter', text)
        self.assertIn('duplicates another type', text)
        self.assertIn('$.types[2]', text)
        self.assertIn('Flag', text)
        self.assertTrue(any('Width' in w for w in warnings))   # template parameter: warning only

    def test_schema_for_type_swaps_materials_without_mutating(self):
        before = json.dumps(GOOD, sort_keys=True)
        steel = fs.schema_for_type(GOOD, 'Steel Top')
        self.assertEqual(steel['geometry'][0]['material'], 'Black Steel')
        self.assertEqual(GOOD['geometry'][0]['material'], 'Oak')
        self.assertEqual(json.dumps(GOOD, sort_keys=True), before)
        self.assertIs(fs.schema_for_type(GOOD, 'Nope'), GOOD)
        self.assertIs(fs.schema_for_type(GOOD, None), GOOD)

    def test_contract_and_prompt_mention_types(self):
        self.assertIn('type_rules', fs.schema_contract())
        self.assertIn('TYPES', fs.build_system_prompt('Furniture'))
        self.assertTrue(fs.build_system_prompt('Furniture').startswith(
            'You create JSON for T3Lab FamiGen in Autodesk Revit.\nReturn ONLY one JSON object'))


if __name__ == '__main__':
    unittest.main()
