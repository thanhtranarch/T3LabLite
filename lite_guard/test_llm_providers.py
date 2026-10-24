#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""LLM provider layer tests -- no Revit, no network.

    python3 lite_guard/test_llm_providers.py

Covers the OpenAI-compatible providers (lib/Intelligence/):
  - openai_provider.py: live model list filtering, default picks, reasoning
    models (max_completion_tokens, reasoning_effort, no temperature)
  - gemini_provider.py: endpoint URLs, Bearer header, "models/" ids, filtering
  - custom_provider.py: Base URL joining, Azure api-key header, typed model
  - openai_compat_provider.py: parameter-rejection retries, streaming agent
    turns with blocking fallback, tool-call parsing (incl. Gemini's
    extra_content thought signature)
  - llm_router.py: registration + FALLBACK_CHAIN, badge labels
  - config/settings.py: endpoints.custom round trip, Custom key slot
  - context_budget.py: Gemini / Custom windows

Every HTTP helper is replaced by a recorder, so nothing leaves the machine.
Plain CPython 3, standard library only.
"""
import json
import os
import shutil
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True

# settings.json lives in %APPDATA%\T3LabAI: point it at a scratch folder
# BEFORE anything imports config.settings.
_TMP = tempfile.mkdtemp(prefix='t3lab_llm_test_')
os.environ['APPDATA'] = _TMP
os.environ['HOME'] = _TMP
os.environ['USERPROFILE'] = _TMP

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
if LIB not in sys.path:
    sys.path.insert(0, LIB)

from config import settings as settings_mod                   # noqa: E402
from Intelligence import llm_provider                         # noqa: E402
from Intelligence import openai_compat_provider as compat     # noqa: E402
from Intelligence import context_budget                       # noqa: E402
from Intelligence import llm_router                           # noqa: E402
from Intelligence.openai_provider import OpenAIProvider       # noqa: E402
from Intelligence.gemini_provider import GeminiProvider       # noqa: E402
from Intelligence.custom_provider import (CustomProvider,     # noqa: E402
                                          normalize_base_url, is_azure_url)
from Intelligence.deepseek_provider import DeepSeekProvider   # noqa: E402


def _fresh_settings():
    """New settings singleton over an empty settings.json."""
    settings_mod.T3LabAISettings._instance = None
    path = os.path.join(_TMP, 'T3LabAI', 'settings.json')
    if os.path.exists(path):
        os.remove(path)
    return settings_mod.T3LabAISettings()


class FakeHTTP(object):
    """Records every request; answers from queues. Patched into both the
    compat module (chat / models) and llm_provider (the agent helpers)."""

    def __init__(self):
        self.gets = []        # (url, headers)
        self.posts = []       # (url, payload, headers)
        self.streams = []     # (url, payload, headers)
        self.get_reply = None
        self.post_replies = []     # str or Exception, FIFO
        self.stream_replies = []   # list of lines or Exception, FIFO

    def http_get_auth(self, url, headers=None, timeout_ms=8000):
        self.gets.append((url, dict(headers or {})))
        return self.get_reply

    def http_post(self, url, payload, headers=None, timeout_ms=60000):
        self.posts.append((url, json.loads(json.dumps(payload)),
                           dict(headers or {})))
        reply = self.post_replies.pop(0)
        if isinstance(reply, Exception):
            raise reply
        return reply

    def http_post_stream(self, url, payload, headers=None, on_line=None,
                         timeout_ms=120000):
        self.streams.append((url, json.loads(json.dumps(payload)),
                             dict(headers or {})))
        reply = self.stream_replies.pop(0)
        if isinstance(reply, Exception):
            raise reply
        for line in reply:
            on_line(line)
        return True


_NAMES = ('http_get_auth', 'http_post', 'http_post_stream')


class ProviderTestCase(unittest.TestCase):

    def setUp(self):
        self.settings = _fresh_settings()
        self.http = FakeHTTP()
        self._saved = []
        for mod in (compat, llm_provider):
            for name in _NAMES:
                self._saved.append((mod, name, getattr(mod, name)))
                setattr(mod, name, getattr(self.http, name))
        self._saved.append((compat, 'HAS_HTTP', compat.HAS_HTTP))
        compat.HAS_HTTP = True

    def tearDown(self):
        for mod, name, value in self._saved:
            setattr(mod, name, value)


def _models(*ids):
    return json.dumps({'object': 'list',
                       'data': [{'id': i, 'object': 'model'} for i in ids]})


def _chat_reply(content=u'OK', tool_calls=None):
    msg = {'role': 'assistant', 'content': content}
    if tool_calls:
        msg['tool_calls'] = tool_calls
    return json.dumps({'choices': [{'index': 0, 'message': msg,
                                    'finish_reason': 'stop'}]})


def _sse(*objs):
    return ['data: ' + json.dumps(o) for o in objs] + ['data: [DONE]']


# ─── OpenAI ───────────────────────────────────────────────────────────────────

class OpenAIModelListTests(ProviderTestCase):

    def test_non_chat_models_are_filtered_and_newest_first(self):
        self.settings.set_api_key('OpenAI', 'sk-test')
        self.http.get_reply = _models(
            'gpt-4.1', 'gpt-5-mini', 'gpt-5.5', 'o4-mini', 'gpt-image-1',
            'gpt-5-codex', 'gpt-4o-transcribe', 'gpt-5-pro',
            'text-embedding-3-small', 'gpt-4o-mini-tts', 'omni-moderation-latest',
            'gpt-4o-search-preview', 'dall-e-3', 'whisper-1',
            'gpt-4o-realtime-preview', 'sora-2', 'o3-deep-research',
            'computer-use-preview', 'gpt-5-nano')
        p = OpenAIProvider()
        models = p.get_models()
        self.assertEqual(models, ['o4-mini', 'gpt-5.5', 'gpt-5-nano',
                                  'gpt-5-mini', 'gpt-4.1'])
        url, headers = self.http.gets[0]
        self.assertEqual(url, 'https://api.openai.com/v1/models')
        self.assertEqual(headers, {'Authorization': 'Bearer sk-test'})

    def test_default_picks_prefer_the_newest_gpt_family(self):
        p = OpenAIProvider()
        live = ['o4-mini', 'gpt-5.5', 'gpt-5-nano', 'gpt-5-mini',
                'gpt-5-chat-latest', 'gpt-4.1', 'gpt-4o']
        # Plain chat: the newest family's mini, never an o-series mini.
        self.assertEqual(p._pick_model(live, prefer_vision=False), 'gpt-5-mini')
        # Images / agent turns: the newest full model (no nano, no chat-latest).
        self.assertEqual(p._pick_model(live, prefer_vision=True), 'gpt-5.5')
        self.assertEqual(p._pick_model(['gpt-6-luna', 'gpt-5.5'], True), 'gpt-6-luna')
        self.assertIsNone(p._pick_model([], True))
        p._cached_models = live
        self.assertEqual(p.pick_fast_model(), 'gpt-5-nano')

    def test_no_key_means_no_models_and_no_request(self):
        p = OpenAIProvider()
        self.assertEqual(p.get_models(), [])
        self.assertFalse(p.check_health())
        self.assertEqual(self.http.gets, [])
        self.assertIsNone(p.chat([], u'sys', u'hi'))
        self.assertIn(u'no API key', p.get_last_error())


class OpenAIReasoningPayloadTests(ProviderTestCase):

    def setUp(self):
        ProviderTestCase.setUp(self)
        self.settings.set_api_key('OpenAI', 'sk-test')
        self.p = OpenAIProvider()
        self.p.set_model('gpt-5.5')

    def test_reasoning_model_detection(self):
        for m in ('o1', 'o3-mini', 'o4-mini', 'gpt-5', 'gpt-5.5', 'gpt-5-nano',
                  'gpt-6-astra', 'openai/gpt-5'):
            self.assertTrue(compat.is_openai_reasoning_model(m), m)
        for m in ('gpt-4o', 'gpt-4.1-mini', 'gpt-5-chat-latest', 'chatgpt-4o-latest',
                  'omni', 'llama-3.3-70b', '', None):
            self.assertFalse(compat.is_openai_reasoning_model(m), m)

    def test_chat_sends_max_completion_tokens_and_no_temperature(self):
        self.http.post_replies.append(_chat_reply(u'Hello'))
        self.assertEqual(self.p.chat([], u'sys', u'hi', max_tokens=600), u'Hello')
        url, payload, headers = self.http.posts[0]
        self.assertEqual(url, 'https://api.openai.com/v1/chat/completions')
        self.assertNotIn('max_tokens', payload)
        self.assertNotIn('temperature', payload)
        # Reasoning tokens count against the cap: raised to the floor.
        self.assertEqual(payload['max_completion_tokens'],
                         compat.REASONING_MIN_TOKENS)
        # Default effort: the model's own (nothing sent).
        self.assertNotIn('reasoning_effort', payload)

    def test_deep_reasoning_sends_high_effort(self):
        self.settings.set_agent_option('extended_thinking', True)
        self.http.post_replies.append(_chat_reply(u'Deep'))
        self.p.chat([], u'sys', u'hi', max_tokens=1500)
        payload = self.http.posts[0][1]
        self.assertEqual(payload['reasoning_effort'], 'high')
        self.assertEqual(payload['max_completion_tokens'],
                         compat.REASONING_HIGH_TOKENS)

    def test_utility_call_sends_low_effort(self):
        self.settings.set_agent_option('extended_thinking', True)
        self.http.post_replies.append(_chat_reply(u'Connected OK'))
        self.p.chat([], u'sys', u'ping', max_tokens=120)
        payload = self.http.posts[0][1]
        self.assertEqual(payload['reasoning_effort'], 'low')
        self.assertEqual(payload['max_completion_tokens'],
                         compat.REASONING_UTILITY_TOKENS)

    def test_non_reasoning_model_keeps_its_budget_and_no_effort(self):
        self.settings.set_agent_option('extended_thinking', True)
        self.p.set_model('gpt-4.1')
        self.http.post_replies.append(_chat_reply(u'Hi'))
        self.p.chat([], u'sys', u'hi', max_tokens=400)
        payload = self.http.posts[0][1]
        self.assertEqual(payload['max_completion_tokens'], 400)
        self.assertNotIn('reasoning_effort', payload)
        self.assertNotIn('temperature', payload)

    def test_rejected_reasoning_effort_is_dropped_and_remembered(self):
        self.settings.set_agent_option('extended_thinking', True)
        self.p.set_model('o1-mini')
        self.http.post_replies.extend([
            RuntimeError(u"API error: {\"error\": {\"message\": \"Unsupported "
                         u"parameter: 'reasoning_effort' is not supported with "
                         u"this model.\"}}"),
            _chat_reply(u'Answer'),
            _chat_reply(u'Again'),
        ])
        self.assertEqual(self.p.chat([], u'sys', u'hi', max_tokens=900), u'Answer')
        self.assertEqual(self.http.posts[0][1]['reasoning_effort'], 'high')
        self.assertNotIn('reasoning_effort', self.http.posts[1][1])
        self.p.chat([], u'sys', u'hi', max_tokens=900)
        self.assertNotIn('reasoning_effort', self.http.posts[2][1])

    def test_rejected_max_completion_tokens_falls_back_to_max_tokens(self):
        self.p.set_model('gpt-4o')
        self.http.post_replies.extend([
            RuntimeError(u"API error: Unrecognized request argument supplied: "
                         u"max_completion_tokens"),
            _chat_reply(u'Old deployment'),
        ])
        self.assertEqual(self.p.chat([], u'sys', u'hi'), u'Old deployment')
        self.assertIn('max_tokens', self.http.posts[1][1])
        self.assertNotIn('max_completion_tokens', self.http.posts[1][1])

    def test_too_large_cap_is_retried_with_the_named_limit(self):
        self.p.set_model('gpt-4o')
        self.http.post_replies.extend([
            RuntimeError(u"API error: max_completion_tokens is too large: 16000. "
                         u"This model supports at most 4096 completion tokens"),
            _chat_reply(u'Fits'),
        ])
        self.assertEqual(self.p.chat([], u'sys', u'hi', max_tokens=16000), u'Fits')
        self.assertEqual(self.http.posts[1][1]['max_completion_tokens'], 4096)

    def test_empty_reasoning_reply_records_an_error(self):
        self.http.post_replies.append(_chat_reply(None))
        self.assertIsNone(self.p.chat([], u'sys', u'hi'))
        self.assertIn(u'empty reply', self.p.get_last_error())

    def test_images_become_image_url_blocks(self):
        self.http.post_replies.append(_chat_reply(u'A wall'))
        content = [{'type': 'text', 'text': u'what is this'},
                   {'type': 'image', 'source': {'type': 'base64',
                                                'media_type': 'image/png',
                                                'data': 'QUJD'}}]
        self.p.chat([], u'sys', content)
        user = self.http.posts[0][1]['messages'][-1]
        self.assertEqual(user['content'][1],
                         {'type': 'image_url',
                          'image_url': {'url': 'data:image/png;base64,QUJD'}})


# ─── Agent turns (tool calls) ────────────────────────────────────────────────

class AgentTurnTests(ProviderTestCase):

    TOOLS = [{'type': 'function', 'function': {
        'name': 'get_levels', 'description': 'List levels',
        'parameters': {'type': 'object', 'properties': {}}}}]

    def setUp(self):
        ProviderTestCase.setUp(self)
        self.settings.set_api_key('OpenAI', 'sk-test')
        self.p = OpenAIProvider()
        self.p.set_model('gpt-5.5')

    def test_streaming_agent_turn_reassembles_tool_calls(self):
        self.http.stream_replies.append(_sse(
            {'choices': [{'index': 0, 'delta': {'content': u'Checking '}}]},
            {'choices': [{'index': 0, 'delta': {'tool_calls': [
                {'index': 0, 'id': 'call_1', 'type': 'function',
                 'function': {'name': 'get_levels', 'arguments': '{"cat'}}]}}]},
            {'choices': [{'index': 0, 'delta': {'tool_calls': [
                {'index': 0, 'function': {'arguments': 'egory": "L"}'}}]}}]},
            {'choices': [{'index': 0, 'delta': {},
                          'finish_reason': 'tool_calls'}]},
        ))
        deltas = []
        res = self.p.chat_agent(u'sys', [{'role': 'user', 'content': u'levels?'}],
                                self.TOOLS, on_delta=deltas.append,
                                max_tokens=1500)
        self.assertEqual(deltas, [u'Checking '])
        self.assertEqual(res['tool_calls'], [{'id': 'call_1', 'name': 'get_levels',
                                              'args': {'category': 'L'}}])
        self.assertEqual(res['stop_reason'], 'tool_use')
        self.assertEqual(res['model'], 'gpt-5.5')
        payload = self.http.streams[0][1]
        self.assertTrue(payload['stream'])
        self.assertIn('max_completion_tokens', payload)
        self.assertNotIn('max_tokens', payload)
        self.assertEqual(payload['tools'], self.TOOLS)
        self.assertEqual(self.http.posts, [])   # no blocking call needed

    def test_stream_failure_falls_back_to_blocking(self):
        self.http.stream_replies.append(IOError('proxy closed the stream'))
        self.http.post_replies.append(_chat_reply(u'', tool_calls=[
            {'id': 'call_9', 'type': 'function',
             'function': {'name': 'get_levels', 'arguments': '{}'}}]))
        res = self.p.chat_agent(u'sys', [{'role': 'user', 'content': u'x'}],
                                self.TOOLS)
        self.assertEqual(res['tool_calls'][0]['name'], 'get_levels')
        self.assertEqual(res['assistant_msg']['tool_calls'][0]['id'], 'call_9')
        self.assertEqual(len(self.http.posts), 1)

    def test_agent_images_are_converted_without_mutating_the_loop_list(self):
        self.http.stream_replies.append(_sse(
            {'choices': [{'index': 0, 'delta': {'content': u'Seen'},
                          'finish_reason': 'stop'}]}))
        msgs = [{'role': 'user', 'content': [
            {'type': 'image', 'source': {'type': 'base64',
                                         'media_type': 'image/jpeg', 'data': 'Zg=='}},
            {'type': 'text', 'text': u'what?'}]}]
        self.p.chat_agent(u'sys', msgs, self.TOOLS)
        sent = self.http.streams[0][1]['messages'][-1]['content'][0]
        self.assertEqual(sent['type'], 'image_url')
        self.assertEqual(msgs[0]['content'][0]['type'], 'image')

    def test_tool_results_are_role_tool_messages(self):
        out = OpenAIProvider.agent_tool_results(
            [{'id': 'a'}, {'id': 'b'}], [u'{"ok": true}'])
        self.assertEqual(out[0], {'role': 'tool', 'tool_call_id': 'a',
                                  'content': u'{"ok": true}'})
        self.assertEqual(out[1]['content'], u'{"cancelled": true}')

    def test_blocking_helper_parses_object_and_string_arguments(self):
        self.http.post_replies.append(_chat_reply(u'<think>x</think>Done', tool_calls=[
            {'id': 'c1', 'type': 'function',
             'function': {'name': 'a', 'arguments': '{"n": 1}'}},
            {'id': 'c2', 'type': 'function',
             'function': {'name': 'b', 'arguments': {'m': 2}}},
            {'id': 'c3', 'type': 'function',
             'function': {'name': 'c', 'arguments': 'not json'}}]))
        res = llm_provider.openai_chat_agent(
            'https://x/v1/chat/completions', {}, 'm', u'sys', [], self.TOOLS)
        self.assertEqual(res['text'], u'Done')
        self.assertEqual([c['args'] for c in res['tool_calls']],
                         [{'n': 1}, {'m': 2}, {}])
        self.assertIn('max_tokens', self.http.posts[0][1])   # default unchanged


# ─── Gemini ──────────────────────────────────────────────────────────────────

class GeminiTests(ProviderTestCase):

    def setUp(self):
        ProviderTestCase.setUp(self)
        self.settings.set_api_key('Gemini', 'AIza-test')
        self.p = GeminiProvider()

    def test_models_endpoint_header_and_filtering(self):
        self.http.get_reply = _models(
            'models/gemini-2.5-flash', 'models/gemini-2.5-pro',
            'models/gemini-3.8-flash', 'models/gemini-flash-lite-latest',
            'models/gemini-embedding-001', 'models/gemini-2.5-flash-image',
            'models/gemini-2.5-flash-preview-tts', 'models/imagen-4.0-generate-001',
            'models/gemma-3-27b-it', 'models/veo-3.0-generate-001',
            'models/gemini-2.5-flash-native-audio-latest',
            'models/gemini-live-2.5-flash', 'models/aqa')
        models = self.p.get_models()
        self.assertEqual(models, ['gemini-flash-lite-latest', 'gemini-3.8-flash',
                                  'gemini-2.5-pro', 'gemini-2.5-flash'])
        url, headers = self.http.gets[0]
        self.assertEqual(
            url, 'https://generativelanguage.googleapis.com/v1beta/openai/models')
        self.assertEqual(headers, {'Authorization': 'Bearer AIza-test'})
        # Default: stable non-lite Flash; fast calls: Flash-Lite.
        self.assertEqual(self.p.get_active_model(), 'gemini-3.8-flash')
        self.assertEqual(self.p.pick_fast_model(), 'gemini-flash-lite-latest')

    def test_chat_url_and_thinking_budget(self):
        self.p.set_model('gemini-2.5-flash')
        self.http.post_replies.append(_chat_reply(u'Xin chào'))
        self.assertEqual(self.p.chat([], u'sys', u'hi', max_tokens=120), u'Xin chào')
        url, payload, headers = self.http.posts[0]
        self.assertEqual(
            url,
            'https://generativelanguage.googleapis.com/v1beta/openai/chat/completions')
        self.assertEqual(headers, {'Authorization': 'Bearer AIza-test'})
        self.assertEqual(payload['model'], 'gemini-2.5-flash')
        # Gemini counts thinking in max_tokens: a tiny cap is raised.
        self.assertEqual(payload['max_tokens'], compat.REASONING_UTILITY_TOKENS)
        self.assertEqual(payload['reasoning_effort'], 'low')
        self.assertNotIn('temperature', payload)

    def test_gemini_2_0_is_not_a_thinking_model(self):
        self.assertFalse(self.p._is_reasoning('gemini-2.0-flash'))
        self.assertTrue(self.p._is_reasoning('gemini-3.8-flash'))
        self.p.set_model('gemini-2.0-flash')
        self.http.post_replies.append(_chat_reply(u'ok'))
        self.p.chat([], u'sys', u'hi', max_tokens=120)
        payload = self.http.posts[0][1]
        self.assertEqual(payload['max_tokens'], 120)
        self.assertNotIn('reasoning_effort', payload)

    def test_thought_signature_survives_a_streamed_tool_call(self):
        self.p.set_model('gemini-3.8-flash')
        sig = {'google': {'thought_signature': 'c2lnbmF0dXJl'}}
        self.http.stream_replies.append(_sse(
            {'choices': [{'index': 0, 'delta': {'role': 'assistant', 'tool_calls': [
                {'index': 0, 'id': 'function-call-1', 'type': 'function',
                 'extra_content': sig,
                 'function': {'name': 'get_levels', 'arguments': '{}'}}]},
                'finish_reason': 'tool_calls'}]}))
        res = self.p.chat_agent(u'sys', [{'role': 'user', 'content': u'x'}], [])
        self.assertEqual(res['assistant_msg']['tool_calls'][0]['extra_content'], sig)
        self.assertEqual(res['tool_calls'][0]['name'], 'get_levels')

    def test_vision_and_tools_supported(self):
        self.assertTrue(self.p.SUPPORTS_VISION)
        self.assertTrue(self.p.SUPPORTS_NATIVE_TOOLS)


# ─── Custom (OpenAI-compatible) ─────────────────────────────────────────────

class CustomProviderTests(ProviderTestCase):

    def setUp(self):
        ProviderTestCase.setUp(self)
        self.p = CustomProvider()

    def _configure(self, url, model=u'', azure=False, key=None):
        self.settings.set_provider_endpoint('custom', base_url=url, model=model,
                                            azure=azure)
        if key is not None:
            self.settings.set_api_key('Custom', key)
        self.p.reload_credentials()

    def test_base_url_joining(self):
        for typed in ('https://openrouter.ai/api/v1', 'https://openrouter.ai/api/v1/',
                      ' https://openrouter.ai/api/v1// ',
                      'https://openrouter.ai/api/v1/chat/completions',
                      'https://openrouter.ai/api/v1/models'):
            self.assertEqual(normalize_base_url(typed), 'https://openrouter.ai/api/v1')
        self.assertEqual(compat.join_url('http://h:8000/v1/', '/models'),
                         'http://h:8000/v1/models')
        self._configure('https://openrouter.ai/api/v1/', key='or-key')
        self.assertEqual(self.p._url('chat/completions'),
                         'https://openrouter.ai/api/v1/chat/completions')

    def test_unconfigured_custom_is_inert(self):
        self.assertFalse(self.p.is_configured())
        self.assertFalse(self.p.check_health())
        self.assertEqual(self.p.get_models(), [])
        self.assertIsNone(self.p.chat([], u'sys', u'hi'))
        self.assertIn(u'no server URL', self.p.get_last_error())
        self.assertEqual(self.http.gets + self.http.posts, [])

    def test_azure_header_choice(self):
        self.assertTrue(is_azure_url('https://res.openai.azure.com/openai/v1/'))
        self.assertTrue(is_azure_url('https://res.services.ai.azure.com:443/openai/v1'))
        self.assertFalse(is_azure_url('https://azure.com.evil.example/v1'))
        self.assertFalse(is_azure_url('https://openrouter.ai/api/v1'))

        self._configure('https://res.openai.azure.com/openai/v1/', key='az-key')
        self.assertEqual(self.p._headers('az-key'), {'api-key': 'az-key'})
        self._configure('https://gateway.corp.local/v1', azure=True, key='az-key')
        self.assertEqual(self.p._headers('az-key'), {'api-key': 'az-key'})
        self._configure('https://api.groq.com/openai/v1', key='gsk')
        self.assertEqual(self.p._headers('gsk'), {'Authorization': 'Bearer gsk'})
        self.assertEqual(self.p._headers(None), {})

    def test_azure_request_uses_deployment_name_and_api_key_header(self):
        self._configure('https://res.openai.azure.com/openai/v1/',
                        model=u'my-gpt5-deploy', key='az-key')
        self.http.get_reply = _models('gpt-4o', 'gpt-5')   # base models only
        self.assertEqual(self.p.get_models(), ['my-gpt5-deploy', 'gpt-4o', 'gpt-5'])
        self.http.post_replies.append(_chat_reply(u'Hi from Azure'))
        self.assertEqual(self.p.chat([], u'sys', u'hi'), u'Hi from Azure')
        url, payload, headers = self.http.posts[0]
        self.assertEqual(url,
                         'https://res.openai.azure.com/openai/v1/chat/completions')
        self.assertEqual(headers, {'api-key': 'az-key'})
        self.assertEqual(payload['model'], 'my-gpt5-deploy')

    def test_server_without_models_endpoint_uses_typed_model(self):
        self._configure('http://vllm.corp:8000/v1', model=u'qwen3-32b')
        self.http.get_reply = None            # /models unavailable
        self.assertEqual(self.p.get_models(), ['qwen3-32b'])
        self.assertTrue(self.p.check_health())
        self.http.post_replies.append(_chat_reply(u'ok'))
        self.p.chat([], u'sys', u'hi', max_tokens=300)
        url, payload, headers = self.http.posts[0]
        self.assertEqual(headers, {})          # no key → no auth header
        self.assertEqual(payload['max_tokens'], 300)
        self.assertNotIn('max_completion_tokens', payload)
        self.assertNotIn('reasoning_effort', payload)

    def test_openai_reasoning_model_through_a_gateway(self):
        self._configure('https://openrouter.ai/api/v1', model=u'openai/gpt-5',
                        key='or')
        self.http.post_replies.append(_chat_reply(u'ok'))
        self.p.chat([], u'sys', u'hi', max_tokens=600)
        payload = self.http.posts[0][1]
        self.assertIn('max_completion_tokens', payload)
        self.assertNotIn('max_tokens', payload)
        self.assertNotIn('temperature', payload)

    def test_vision_follows_the_model(self):
        self._configure('https://openrouter.ai/api/v1', model=u'google/gemini-3.8-flash')
        self.assertTrue(self.p.SUPPORTS_VISION)
        self._configure('https://api.groq.com/openai/v1', model=u'llama-3.3-70b-versatile')
        self.assertFalse(self.p.SUPPORTS_VISION)
        # A text-only model never receives image blocks.
        msgs = [{'role': 'user', 'content': [
            {'type': 'image', 'source': {'type': 'base64', 'media_type': 'image/png',
                                         'data': 'QQ=='}},
            {'type': 'text', 'text': u'hi'}]}]
        self.assertEqual(self.p._agent_messages(msgs)[0]['content'],
                         [{'type': 'text', 'text': u'hi'}])


# ─── Settings ────────────────────────────────────────────────────────────────

class SettingsTests(unittest.TestCase):

    def test_endpoint_round_trip_survives_a_reload(self):
        s = _fresh_settings()
        self.assertEqual(s.get_provider_endpoint('custom'),
                         {'base_url': u'', 'model': u'', 'azure': False})
        self.assertTrue(s.set_provider_endpoint(
            'custom', base_url=u' https://res.openai.azure.com/openai/v1/ ',
            model=u'deploy-1', azure=True))
        settings_mod.T3LabAISettings._instance = None
        s2 = settings_mod.T3LabAISettings()
        self.assertEqual(s2.get_provider_endpoint('custom'),
                         {'base_url': u'https://res.openai.azure.com/openai/v1/',
                          'model': u'deploy-1', 'azure': True})
        # Partial update keeps the other fields.
        s2.set_provider_endpoint('custom', model=u'deploy-2')
        ep = s2.get_provider_endpoint('custom')
        self.assertEqual((ep['base_url'], ep['model'], ep['azure']),
                         (u'https://res.openai.azure.com/openai/v1/', u'deploy-2', True))
        # The URL is plain config, not under api_keys.
        with open(s2._settings_file) as f:
            raw = json.load(f)
        self.assertIn('custom', raw['endpoints'])
        self.assertNotIn('Custom', raw.get('api_keys', {}))

    def test_old_settings_file_without_endpoints_block(self):
        s = _fresh_settings()
        with open(s._settings_file, 'w') as f:
            json.dump({'api_keys': {}, 'active_provider': 'claude'}, f)
        settings_mod.T3LabAISettings._instance = None
        s2 = settings_mod.T3LabAISettings()
        self.assertEqual(s2.get_provider_endpoint('custom')['base_url'], u'')

    def test_custom_and_gemini_keys_use_the_protected_key_path(self):
        from config import secret_store
        secret_store.set_protector((lambda b: b[::-1], lambda b: b[::-1]))
        try:
            s = _fresh_settings()
            s.set_api_key('Gemini', u'AIza-secret')
            s.set_api_key('Custom', u'custom-secret')
            with open(s._settings_file) as f:
                raw = json.load(f)
            self.assertTrue(raw['api_keys']['Gemini'].startswith(secret_store.PREFIX))
            self.assertNotIn('custom-secret', json.dumps(raw))
            self.assertEqual(s.get_api_key('Custom'), u'custom-secret')
        finally:
            secret_store.set_protector(None)


# ─── Router & context windows ───────────────────────────────────────────────

class RouterTests(unittest.TestCase):

    def setUp(self):
        _fresh_settings()
        llm_router.LLMRouter._instance = None

    def tearDown(self):
        llm_router.LLMRouter._instance = None

    def test_new_providers_are_registered_and_in_the_fallback_chain(self):
        chain = llm_router.LLMRouter.FALLBACK_CHAIN
        for name in ('claude', 'openai', 'gemini', 'deepseek', 'custom',
                     'ollama', 'lmstudio'):
            self.assertIn(name, chain)
        # Local providers stay last; cloud ones keep their old order.
        self.assertEqual(chain[-2:], ['ollama', 'lmstudio'])
        self.assertLess(chain.index('openai'), chain.index('gemini'))
        router = llm_router.LLMRouter()
        self.assertIsInstance(router.get_provider('gemini'), GeminiProvider)
        self.assertIsInstance(router.get_provider('custom'), CustomProvider)
        self.assertIsInstance(router.get_provider('openai'), OpenAIProvider)
        self.assertIsInstance(router.get_provider('deepseek'), DeepSeekProvider)
        self.assertEqual(router.get_provider_names(), chain)
        self.assertIn('custom', router.get_remote_provider_names())
        self.assertTrue(router.switch_provider('gemini', persist=False))
        self.assertEqual(router.get_active_name(), 'gemini')

    def test_instant_status_uses_configuration_not_the_network(self):
        s = settings_mod.T3LabAISettings()
        s.set_provider_endpoint('custom', base_url=u'http://vllm.corp:8000/v1')
        snap = llm_router.LLMRouter().get_status_instant()
        self.assertTrue(snap['custom']['available'])
        self.assertFalse(snap['gemini']['available'])
        self.assertFalse(snap['custom']['probed'])

    def test_badge_labels(self):
        self.assertEqual(llm_router._gpt_label('gpt-5.5'), 'GPT-5.5')
        self.assertEqual(llm_router._gpt_label('gpt-5-mini'), 'GPT-5 mini')
        self.assertEqual(llm_router._gpt_label('gpt-4o-mini'), 'GPT-4o mini')
        self.assertEqual(llm_router._gpt_label('o4-mini'), 'o4-mini')
        self.assertEqual(llm_router._gpt_label(''), 'GPT')


class ContextWindowTests(unittest.TestCase):

    def test_windows(self):
        cw = context_budget.context_window
        self.assertEqual(cw('gemini', 'gemini-3.8-flash'), context_budget.GEMINI_WINDOW)
        self.assertEqual(cw('custom', 'openai/gpt-5'), 400000)
        self.assertEqual(cw('custom', 'google/gemini-2.5-pro'),
                         context_budget.GEMINI_WINDOW)
        self.assertEqual(cw('custom', 'my-deployment'), context_budget.UNKNOWN_WINDOW)
        self.assertEqual(cw('custom', None), context_budget.UNKNOWN_WINDOW)


def tearDownModule():
    shutil.rmtree(_TMP, ignore_errors=True)


if __name__ == '__main__':
    unittest.main(verbosity=2)
