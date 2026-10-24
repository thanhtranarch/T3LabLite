#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Remote MCP servers in the T3Lab Assistant -- no Revit, no network.

    python3 lite_guard/test_assistant_mcp.py

Covers:
  - Intelligence/remote_mcp.py: validation (https only, names), tokens kept
    or replaced on save, request_parts() for the Claude request
  - Intelligence/claude_provider.py: mcp_servers + mcp_toolset in the agent
    payload (after the cached tool list), mcp_tool_use blocks kept in the
    transcript but never turned into client tool calls
  - Intelligence/agent_loop.py: a pause_turn from server-run tools is
    resumed instead of ending the request

Plain CPython 3, standard library only.
"""
import os
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True

_TMP = tempfile.mkdtemp(prefix='t3lab_amcp_test_')
os.environ['APPDATA'] = _TMP
os.environ['HOME'] = _TMP
os.environ['USERPROFILE'] = _TMP

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
if LIB not in sys.path:
    sys.path.insert(0, LIB)

from Intelligence import remote_mcp                    # noqa: E402
from Intelligence import claude_provider as cp         # noqa: E402
from Intelligence.agent_loop import AgentLoop          # noqa: E402


class FakeSettings(object):
    """The two agents-block accessors remote_mcp uses."""

    def __init__(self):
        self.agents = {}

    def get_agent_option(self, key, default=None):
        return self.agents.get(key, default)

    def set_agent_option(self, key, value):
        self.agents[key] = value
        return True


class RemoteMcpTests(unittest.TestCase):

    def setUp(self):
        self.s = FakeSettings()

    def test_https_only_and_names(self):
        _saved, err = remote_mcp.save([{'name': 'wiki', 'url': 'http://x.example'}], self.s)
        self.assertIn('https://', err)
        _saved, err = remote_mcp.save([{'name': 'a', 'url': 'https://a.example'},
                                       {'name': 'a', 'url': 'https://b.example'}], self.s)
        self.assertIn('already uses', err)
        self.assertEqual(self.s.agents, {})                 # nothing written on error
        self.assertEqual(remote_mcp.clean_name('Company Wiki!'), 'company-wiki')

    def test_save_load_and_request_parts(self):
        saved, err = remote_mcp.save([
            {'name': 'Company Wiki', 'url': ' https://wiki.example/mcp ', 'token': 'tok1'},
            {'name': 'linear', 'url': 'https://mcp.linear.app/sse', 'enabled': False}], self.s)
        self.assertIsNone(err)
        self.assertEqual([r['name'] for r in saved], ['company-wiki', 'linear'])
        self.assertTrue(saved[0]['has_token'])
        self.assertNotIn('token', saved[0])                 # load() never shows it
        servers, toolsets = remote_mcp.request_parts(self.s)
        self.assertEqual(servers, [{'type': 'url', 'url': 'https://wiki.example/mcp',
                                    'name': 'company-wiki', 'authorization_token': 'tok1'}])
        self.assertEqual(toolsets, [{'type': 'mcp_toolset',
                                     'mcp_server_name': 'company-wiki'}])

    def test_token_kept_unless_given(self):
        remote_mcp.save([{'name': 'wiki', 'url': 'https://w.example', 'token': 'old'}], self.s)
        remote_mcp.save([{'name': 'wiki', 'url': 'https://w2.example'}], self.s)
        servers, _ = remote_mcp.request_parts(self.s)
        self.assertEqual(servers[0]['authorization_token'], 'old')
        self.assertEqual(servers[0]['url'], 'https://w2.example')
        remote_mcp.save([{'name': 'kb', 'previous_name': 'wiki',
                          'url': 'https://w2.example'}], self.s)
        servers, _ = remote_mcp.request_parts(self.s)
        self.assertEqual((servers[0]['name'], servers[0]['authorization_token']),
                         ('kb', 'old'))                        # kept across a rename
        remote_mcp.save([{'name': 'kb', 'url': 'https://w2.example', 'token': ''}], self.s)
        servers, _ = remote_mcp.request_parts(self.s)
        self.assertNotIn('authorization_token', servers[0])

    def test_broken_settings_never_raise(self):
        self.s.agents[remote_mcp.SETTING_KEY] = ['junk', {'name': 'x y', 'url': 'https://a'}]
        self.assertEqual(remote_mcp.request_parts(self.s), ([], []))
        self.assertEqual(len(remote_mcp.load(self.s)), 1)


class ClaudePayloadTests(unittest.TestCase):

    def setUp(self):
        self.p = cp.ClaudeProvider()
        self._orig = (cp.ClaudeProvider._remote_mcp_parts, cp.ClaudeProvider._mcp_on)
        self.parts = ([{'type': 'url', 'url': 'https://w.example', 'name': 'wiki'}],
                      [{'type': 'mcp_toolset', 'mcp_server_name': 'wiki'}])
        cp.ClaudeProvider._remote_mcp_parts = staticmethod(lambda: self.parts)
        cp.ClaudeProvider._mcp_on = lambda _self, model: _self._mcp_mode.get(model) != 'off'

    def tearDown(self):
        cp.ClaudeProvider._remote_mcp_parts = staticmethod(self._orig[0])
        cp.ClaudeProvider._mcp_on = self._orig[1]

    def _payload(self, model='claude-opus-5-5'):
        tools = [{'name': 'list_levels', 'input_schema': {'type': 'object'}}]
        return self.p._agent_payload(model, 'sys', [{'role': 'user', 'content': 'hi'}],
                                     tools, 1000, stream=True)

    def test_servers_and_toolsets_after_cached_tools(self):
        payload = self._payload()
        self.assertEqual(payload['mcp_servers'], self.parts[0])
        self.assertEqual(payload['tools'][-1], self.parts[1][0])
        self.assertIn('cache_control', payload['tools'][0])  # breakpoint unchanged
        self.assertNotIn('cache_control', payload['tools'][-1])

    def test_refused_model_gets_none(self):
        self.p._mcp_mode['claude-opus-5-5'] = 'off'
        payload = self._payload()
        self.assertNotIn('mcp_servers', payload)
        self.assertEqual(len(payload['tools']), 1)
        self.assertTrue(cp.ClaudeProvider._mcp_rejected(
            Exception('400: mcp_servers: could not connect')))

    def test_mcp_blocks_are_transcript_only(self):
        blocks = [
            {'type': 'mcp_tool_use', 'id': 'mcptoolu_1', 'name': 'search',
             'server': 'wiki', 'parts': ['{"q": ', '"BEP"}']},
            {'type': 'raw', 'parts': [], 'raw': {
                'type': 'mcp_tool_result', 'tool_use_id': 'mcptoolu_1',
                'is_error': False, 'content': [{'type': 'text', 'text': 'found'}]}},
            {'type': 'text', 'parts': ['The BEP says A-101.']},
        ]
        res = cp.ClaudeProvider._agent_result_from_blocks(blocks, 'end_turn')
        self.assertEqual(res['tool_calls'], [])
        self.assertEqual(res['mcp_calls'], [{'server': 'wiki', 'name': 'search'}])
        content = res['assistant_msg']['content']
        self.assertEqual(content[0], {'type': 'mcp_tool_use', 'id': 'mcptoolu_1',
                                      'name': 'search', 'server_name': 'wiki',
                                      'input': {'q': 'BEP'}})
        self.assertEqual(content[1]['type'], 'mcp_tool_result')
        self.assertEqual(res['text'], 'The BEP says A-101.')


class PauseTurnTests(unittest.TestCase):

    def test_pause_turn_is_resumed(self):
        replies = [
            {'text': 'Searching the wiki...', 'tool_calls': [], 'stop_reason': 'pause_turn',
             'assistant_msg': {'role': 'assistant', 'content': [
                 {'type': 'mcp_tool_use', 'id': 'm1', 'name': 's',
                  'server_name': 'wiki', 'input': {}}]}},
            {'text': 'Done: A-101.', 'tool_calls': [], 'stop_reason': 'end_turn',
             'assistant_msg': {'role': 'assistant', 'content': [
                 {'type': 'text', 'text': 'Done: A-101.'}]}},
        ]
        seen = []

        class Provider(object):
            def chat_agent(self, system_prompt, messages, tools, on_delta=None,
                           max_tokens=0):
                seen.append([m['role'] for m in messages])
                return replies.pop(0)

            def agent_tool_results(self, calls, results):
                return []

        loop = AgentLoop(Provider(), lambda *a, **k: {}, [])
        out = loop.run([], 'sys', 'what does the BEP say?')
        self.assertEqual(len(seen), 2)
        # The paused assistant turn is sent back as is, with no new user turn.
        self.assertEqual(seen[1][-1], 'assistant')
        self.assertEqual((out['status'], out['text']), ('done', 'Done: A-101.'))
        self.assertEqual(out['iterations'], 2)


if __name__ == '__main__':
    unittest.main(verbosity=1)
