#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""MCP server, bridge and policy tests -- no Revit needed.

    python3 lite_guard/test_mcp_server.py

Covers what an external AI app (Claude Desktop, Codex ... via bridge.py) can
reach:
  - core/mcp_policy.py: read-only mode, code execution, destructive-call
    consent, annotations, isError, Origin and token checks
  - core/server.py: the real HTTP handler on a local port (auth, -32700,
    202 for notifications, ping, tools/list without the request lock),
    tools/list filtering, tools/call blocking, run_batch validation
  - core/bridge.py: re-reading the token after a 401
  - core/file_watcher.py: off by default, never runs an old task file
  - tracking/service.py: code, secrets and full paths are not sent

Revit is not loaded here, so the ExternalEvent and the confirmation dialog
are not exercised; the Revit side must be checked in Revit.

Plain CPython 3, standard library only. CI runs it next to guard.py
(.github/workflows/lite-guard.yml).
"""
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import unittest
import urllib.error
import urllib.request

sys.dont_write_bytecode = True

# Every module below reads %APPDATA% (token, settings) and ~ (data folder):
# point both at a scratch folder BEFORE importing them.
_TMP = tempfile.mkdtemp(prefix='t3lab_mcp_test_')
os.environ['APPDATA'] = _TMP
os.environ['HOME'] = _TMP
os.environ['USERPROFILE'] = _TMP

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
if LIB not in sys.path:
    sys.path.insert(0, LIB)

from core import mcp_policy as mp      # noqa: E402
from core import paths                 # noqa: E402
from core import server as srv_mod     # noqa: E402
from core import bridge                # noqa: E402
from Intelligence import tool_schema   # noqa: E402
from tracking.service import redact_arguments, TelemetryService  # noqa: E402

READ_ONLY = tool_schema.READ_ONLY_TOOL_NAMES

# Tool calls below go through the usage tracking hook: never let a test run
# send rows to the real dashboard.
paths.save_settings({'tracking_enabled': False})


def _set_settings(**values):
    data = paths.load_settings() or {}
    data.update(values)
    paths.save_settings(data)


def _clear_policy():
    data = paths.load_settings() or {}
    for key in (mp.READ_ONLY_KEY, mp.CONFIRM_KEY, mp.CODE_KEY):
        data.pop(key, None)
    paths.save_settings(data)


_SERVER = None


def _server():
    """One registry for the whole run (building it is the slow part)."""
    global _SERVER
    if _SERVER is None:
        _SERVER = srv_mod.T3LabAIServer()
    return _SERVER


class PolicyTests(unittest.TestCase):

    def test_defaults(self):
        self.assertEqual(mp.policy_from_settings({}), mp.DEFAULT_POLICY)
        self.assertEqual(mp.policy_from_settings(None), mp.DEFAULT_POLICY)

    def test_bad_values_fall_back(self):
        p = mp.policy_from_settings({mp.CODE_KEY: 'sometimes',
                                     mp.READ_ONLY_KEY: 'maybe',
                                     mp.CONFIRM_KEY: 'no'})
        self.assertEqual(p['code_execution'], 'ask')
        self.assertFalse(p['read_only'])
        self.assertFalse(p['confirm_destructive'])
        self.assertEqual(mp.policy_from_settings({mp.CODE_KEY: ' OFF '})['code_execution'], 'off')

    def test_negotiate_protocol(self):
        self.assertEqual(mp.negotiate_protocol('2025-03-26'), '2025-03-26')
        self.assertEqual(mp.negotiate_protocol('2024-11-05'), '2024-11-05')
        self.assertEqual(mp.negotiate_protocol('2099-01-01'), mp.LATEST_PROTOCOL)
        self.assertEqual(mp.negotiate_protocol(None), mp.LATEST_PROTOCOL)

    def test_check_call(self):
        ro = frozenset(['list_levels'])
        default = dict(mp.DEFAULT_POLICY)
        self.assertEqual(mp.check_call('list_levels', default, ro, False)[0], 'run')
        self.assertEqual(mp.check_call('create_level', default, ro, False)[0], 'run')
        self.assertEqual(mp.check_call('delete_element', default, ro, True)[0], 'confirm')
        self.assertEqual(mp.check_call('send_code_to_revit', default, ro, False)[0], 'confirm')
        no_confirm = dict(default, confirm_destructive=False)
        self.assertEqual(mp.check_call('delete_element', no_confirm, ro, True)[0], 'run')
        read_only = dict(default, read_only=True)
        self.assertEqual(mp.check_call('create_level', read_only, ro, False)[0], 'block')
        self.assertEqual(mp.check_call('list_levels', read_only, ro, False)[0], 'run')
        allow = dict(default, code_execution='allow')
        self.assertEqual(mp.check_call('send_code_to_revit', allow, ro, False)[0], 'run')
        self.assertEqual(mp.check_call('send_code_to_revit', dict(allow, read_only=True), ro, False)[0],
                         'block')
        off = dict(default, code_execution='off')
        decision, message = mp.check_call('send_code_to_revit', off, ro, False)
        self.assertEqual(decision, 'block')
        self.assertIn('turned off', message)

    def test_visibility(self):
        ro = frozenset(['list_levels'])
        default = dict(mp.DEFAULT_POLICY)
        self.assertTrue(mp.tool_visible('send_code_to_revit', default, ro))
        self.assertFalse(mp.tool_visible('send_code_to_revit', dict(default, code_execution='off'), ro))
        self.assertFalse(mp.tool_visible('create_level', dict(default, read_only=True), ro))
        self.assertTrue(mp.tool_visible('list_levels', dict(default, read_only=True), ro))

    def test_annotations(self):
        ro = frozenset(['list_levels'])
        self.assertEqual(mp.tool_annotations('list_levels', ro),
                         {'readOnlyHint': True, 'openWorldHint': False})
        self.assertEqual(mp.tool_annotations('create_level', ro),
                         {'readOnlyHint': False, 'openWorldHint': False, 'destructiveHint': False})
        self.assertTrue(mp.tool_annotations('delete_element', ro)['destructiveHint'])
        tool = {'name': 'list_levels', 'inputSchema': {}}
        entry = mp.list_entry(tool, ro)
        self.assertNotIn('annotations', tool, 'the registry entry itself must not change')
        self.assertTrue(entry['annotations']['readOnlyHint'])

    def test_is_error_result(self):
        self.assertTrue(mp.is_error_result({'error': 'x'}))
        self.assertTrue(mp.is_error_result({'success': False}))
        self.assertFalse(mp.is_error_result({'success': True, 'error': 'warning only'}))
        self.assertFalse(mp.is_error_result({'error': ''}))
        self.assertFalse(mp.is_error_result({'ok': 1}))
        self.assertFalse(mp.is_error_result([1, 2]))

    def test_origin(self):
        for ok in (None, '', 'http://localhost', 'http://localhost:3000',
                   'https://127.0.0.1:8443', 'http://[::1]:5173'):
            self.assertTrue(mp.origin_allowed(ok), ok)
        for bad in ('null', 'https://evil.example', 'http://localhost.evil.example',
                    'http://127.0.0.1.evil.example', 'file://'):
            self.assertFalse(mp.origin_allowed(bad), bad)

    def test_token(self):
        self.assertTrue(mp.token_matches('Bearer abc', 'abc'))
        self.assertFalse(mp.token_matches('Bearer abd', 'abc'))
        self.assertFalse(mp.token_matches('', 'abc'))
        self.assertFalse(mp.token_matches('Bearer ', ''))

    def test_confirm_text_shows_code(self):
        main, content, expanded = mp.confirm_text('send_code_to_revit',
                                                  {'code': 'x = 1\ny = 2'})
        self.assertIn('Python', main)
        self.assertIn('2 line', content)
        self.assertEqual(expanded, 'x = 1\ny = 2')
        _main, _content, long_expanded = mp.confirm_text(
            'send_code_to_revit', {'code': 'a' * 5000})
        self.assertIn('more characters not shown', long_expanded)

    def test_save_policy_round_trip(self):
        try:
            p = mp.save_policy(read_only=True, code_execution='OFF')
            self.assertTrue(p['read_only'])
            self.assertEqual(p['code_execution'], 'off')
            self.assertEqual(mp.load_policy(), p)
            with self.assertRaises(ValueError):
                mp.save_policy(code_execution='sometimes')
        finally:
            _clear_policy()


class RegistryTests(unittest.TestCase):

    def test_run_batch_steps_are_real_tools(self):
        s = _server()
        self.assertIn('run_batch', s._tools)
        missing = sorted(s._BATCH_STEP_TOOLS - set(s._tools))
        self.assertEqual(missing, [], 'run_batch allows tools that do not exist')
        for banned in ('send_code_to_revit', 'run_batch', 'manage_document',
                       'export_sheets_pdf', 'set_active_view', 'select_elements',
                       'switch_active_document', 'close_document', 'famigen_create_family'):
            self.assertNotIn(banned, s._BATCH_STEP_TOOLS)
        self.assertIn('run_batch', s._WRITE_TOOLS)
        self.assertIn('run_batch', tool_schema.LOCAL_HIDDEN_TOOLS)

    def test_point_cloud_tools_are_gone(self):
        # Removed 2026-10-20: no MCP client called them; the ribbon wizard stays.
        s = _server()
        for name in ('list_point_clouds', 'analyze_point_cloud', 'detect_point_cloud_elements'):
            self.assertNotIn(name, s._tools)
            self.assertNotIn(name, s._WRITE_TOOLS)
            self.assertNotIn(name, READ_ONLY)
            self.assertNotIn(name, bridge.LARGE_RESULT_TOOLS)

    def test_destructive_hint_tools_are_real(self):
        s = _server()
        self.assertEqual(sorted(mp.DESTRUCTIVE_HINT_TOOLS - set(s._tools)), [])

    def test_is_destructive(self):
        s = _server()
        self.assertTrue(s.is_destructive('close_document', {'path_or_title': 'A'}))
        self.assertFalse(s.is_destructive('close_document', {'path_or_title': 'A', 'save': True}))
        self.assertTrue(s.is_destructive('run_batch', {'steps': [
            {'name': 'create_level', 'arguments': {}},
            {'name': 'delete_element', 'arguments': {'element_ids': [1]}}]}))
        self.assertFalse(s.is_destructive('run_batch', {'steps': [
            {'name': 'create_level', 'arguments': {}}]}))

    def test_batch_problem(self):
        s = _server()
        self.assertIsNone(s._batch_problem({'steps': [
            {'name': 'list_levels'}, {'name': 'create_level', 'arguments': {}}]}))
        for bad, needle in (({}, 'non-empty'),
                            ({'steps': [{'arguments': {}}]}, 'needs a "name"'),
                            ({'steps': [{'name': 'send_code_to_revit'}]}, 'cannot run inside'),
                            ({'steps': [{'name': 'list_levels', 'args': {}}]}, 'unknown keys'),
                            ({'steps': [{'name': 'list_levels', 'arguments': []}]}, 'must be an object'),
                            ({'steps': [{'name': 'list_levels'}] * 51}, 'at most')):
            problem = s._batch_problem(bad)
            self.assertIsNotNone(problem, bad)
            self.assertIn(needle, problem['error'])
        unknown = s._batch_problem({'steps': [
            {'name': 'create_level', 'arguments': {'not_a_real_argument': 1}}]})
        self.assertIn('Step 1', unknown['error'])
        self.assertIn('not_a_real_argument', unknown['error'])


class ToolCallTests(unittest.TestCase):

    def setUp(self):
        _clear_policy()
        self.s = _server()
        self.calls = []

        def fake_execute(name, arguments, confirm=None):
            self.calls.append((name, arguments, confirm))
            return {'error': 'boom'} if name == 'create_grid' else {'ok': name}
        self.s._execute_tool = fake_execute
        self.s._external_event = None

    def tearDown(self):
        del self.s._execute_tool
        _clear_policy()

    def _call(self, name, arguments=None):
        return self.s._handle_tool_call({'name': name, 'arguments': arguments or {}})

    def test_runs_and_sets_is_error(self):
        res = self._call('create_level')
        self.assertFalse(res['isError'])
        self.assertEqual(json.loads(res['content'][0]['text']), {'ok': 'create_level'})
        self.assertTrue(self._call('create_grid')['isError'])

    def test_read_only_blocks_writes(self):
        _set_settings(**{mp.READ_ONLY_KEY: True})
        res = self._call('create_level')
        self.assertTrue(res['isError'])
        self.assertIn('Read-only mode', res['content'][0]['text'])
        self.assertEqual(self.calls, [])
        self.assertFalse(self._call('list_levels')['isError'])

    def test_code_off_blocks(self):
        _set_settings(**{mp.CODE_KEY: 'off'})
        res = self._call('send_code_to_revit', {'code': 'result = 1'})
        self.assertTrue(res['isError'])
        self.assertEqual(self.calls, [])

    def test_confirmation_never_runs_without_ui_thread(self):
        # No ExternalEvent = no dialog = the call must not run unasked.
        res = self._call('delete_element', {'element_ids': [1]})
        self.assertTrue(res['isError'])
        self.assertIn('confirm', res['content'][0]['text'])
        res = self._call('send_code_to_revit', {'code': 'result = 1'})
        self.assertTrue(res['isError'])
        self.assertEqual(self.calls, [])

    def test_confirmation_is_passed_to_the_ui_thread(self):
        self.s._external_event = object()
        try:
            self._call('delete_element', {'element_ids': [1]})
            self._call('create_level', {})
        finally:
            self.s._external_event = None
        (name1, _a1, confirm1), (name2, _a2, confirm2) = self.calls
        self.assertEqual(name1, 'delete_element')
        self.assertTrue(confirm1 and 'delete_element' in confirm1[0])
        self.assertIsNone(confirm2)

    def test_confirm_can_be_turned_off(self):
        _set_settings(**{mp.CONFIRM_KEY: False})
        res = self._call('delete_element', {'element_ids': [1]})
        self.assertFalse(res['isError'])
        self.assertIsNone(self.calls[0][2])

    def test_tools_list_follows_policy(self):
        names = {t['name']: t for t in self.s._handle_tools_list()['tools']}
        self.assertIn('send_code_to_revit', names)
        self.assertIn('run_batch', names)
        self.assertTrue(names['list_levels']['annotations']['readOnlyHint'])
        self.assertTrue(names['delete_element']['annotations']['destructiveHint'])
        self.assertFalse(names['create_level']['annotations']['destructiveHint'])
        self.assertNotIn('annotations', self.s._tools['list_levels'])

        _set_settings(**{mp.READ_ONLY_KEY: True})
        names = {t['name'] for t in self.s._handle_tools_list()['tools']}
        self.assertNotIn('create_level', names)
        self.assertNotIn('send_code_to_revit', names)
        self.assertIn('list_levels', names)
        self.assertTrue(names <= set(READ_ONLY))

        _clear_policy()
        _set_settings(**{mp.CODE_KEY: 'off'})
        names = {t['name'] for t in self.s._handle_tools_list()['tools']}
        self.assertNotIn('send_code_to_revit', names)
        self.assertIn('create_level', names)

    def test_initialize(self):
        res = self.s._handle_initialize({'protocolVersion': '2025-06-18'})
        self.assertEqual(res['protocolVersion'], '2025-06-18')
        self.assertEqual(self.s._handle_initialize({})['protocolVersion'], mp.LATEST_PROTOCOL)


class ToolTaskTests(unittest.TestCase):

    def test_give_up_before_start_means_never_runs(self):
        task = srv_mod._ToolTask('create_level', {})
        self.assertTrue(task.give_up())
        self.assertFalse(task.claim('ui'))

    def test_give_up_while_running_reports_running(self):
        task = srv_mod._ToolTask('create_level', {})
        self.assertTrue(task.claim('ui'))
        self.assertFalse(task.give_up())

    def test_late_yes_does_not_run(self):
        task = srv_mod._ToolTask('delete_element', {}, confirm=('a', 'b', 'c'))
        self.assertTrue(task.claim('ui'))
        task.awaiting_user = True
        self.assertTrue(task.give_up())          # caller gives up mid-dialog
        self.assertFalse(task.proceed_after_confirm())

    def test_yes_in_time_runs(self):
        task = srv_mod._ToolTask('delete_element', {}, confirm=('a', 'b', 'c'))
        self.assertTrue(task.claim('ui'))
        task.awaiting_user = True
        self.assertTrue(task.proceed_after_confirm())


class HttpTests(unittest.TestCase):
    """The real MCPRequestHandler on a free local port."""

    @classmethod
    def setUpClass(cls):
        cls.s = _server()
        cls.http = srv_mod._ThreadedHTTPServer(('127.0.0.1', 0), srv_mod.MCPRequestHandler)
        cls.http.mcp_server = cls.s
        cls.port = cls.http.server_address[1]
        cls.thread = threading.Thread(target=cls.http.serve_forever)
        cls.thread.daemon = True
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.http.shutdown()
        cls.http.server_close()

    def _post(self, body, token=True, headers=None):
        data = body if isinstance(body, bytes) else json.dumps(body).encode('utf-8')
        h = {'Content-Type': 'application/json'}
        if token:
            h['Authorization'] = 'Bearer ' + self.s.token
        h.update(headers or {})
        req = urllib.request.Request('http://127.0.0.1:{}/mcp'.format(self.port),
                                     data=data, headers=h)
        try:
            with urllib.request.urlopen(req, timeout=5) as r:
                raw = r.read()
                return r.status, (json.loads(raw) if raw else None)
        except urllib.error.HTTPError as e:
            raw = e.read()
            return e.code, (json.loads(raw) if raw else None)

    def test_needs_token(self):
        code, body = self._post({'jsonrpc': '2.0', 'id': 1, 'method': 'ping'}, token=False)
        self.assertEqual(code, 401)
        self.assertEqual(body['error']['code'], -32001)

    def test_ping(self):
        code, body = self._post({'jsonrpc': '2.0', 'id': 7, 'method': 'ping'})
        self.assertEqual((code, body['id'], body['result']), (200, 7, {}))

    def test_parse_error(self):
        code, body = self._post(b'{not json')
        self.assertEqual(code, 400)
        self.assertEqual(body['error']['code'], -32700)

    def test_batch_array_rejected(self):
        code, body = self._post([{'jsonrpc': '2.0', 'id': 1, 'method': 'ping'}])
        self.assertEqual(code, 400)
        self.assertEqual(body['error']['code'], -32600)

    def test_notification_gets_202_and_no_body(self):
        code, body = self._post({'jsonrpc': '2.0', 'method': 'notifications/initialized'})
        self.assertEqual((code, body), (202, None))

    def test_foreign_origin_refused(self):
        code, _body = self._post({'jsonrpc': '2.0', 'id': 1, 'method': 'ping'},
                                 headers={'Origin': 'https://evil.example'})
        self.assertEqual(code, 403)

    def test_tools_list_does_not_wait_for_a_running_call(self):
        # A slow tools/call holds the request lock; tools/list must still answer.
        with srv_mod._MCP_REQUEST_LOCK:
            t0 = time.time()
            code, body = self._post({'jsonrpc': '2.0', 'id': 2, 'method': 'tools/list'})
            self.assertLess(time.time() - t0, 3)
        self.assertEqual(code, 200)
        self.assertGreater(len(body['result']['tools']), 50)

    def test_unknown_method(self):
        code, body = self._post({'jsonrpc': '2.0', 'id': 3, 'method': 'resources/list'})
        self.assertEqual(body['error']['code'], -32601)


class BridgeTokenTests(unittest.TestCase):
    """bridge.py picks up a token written after it started."""

    def setUp(self):
        self.s = _server()
        self.http = srv_mod._ThreadedHTTPServer(('127.0.0.1', 0), srv_mod.MCPRequestHandler)
        self.http.mcp_server = self.s
        self.port = self.http.server_address[1]
        t = threading.Thread(target=self.http.serve_forever)
        t.daemon = True
        t.start()
        bridge._LATEST_TOKEN[0] = ''

    def tearDown(self):
        self.http.shutdown()
        self.http.server_close()
        bridge._LATEST_TOKEN[0] = ''

    def test_stale_token_is_refreshed(self):
        self.assertEqual(bridge._read_token(), self.s.token)
        resp = bridge._post_mcp(self.port, {'jsonrpc': '2.0', 'id': 1, 'method': 'ping'},
                                'stale-token', timeout=5)
        self.assertEqual(resp['result'], {})
        self.assertEqual(bridge._LATEST_TOKEN[0], self.s.token)

    def test_wrong_token_on_disk_still_fails(self):
        path = os.path.join(_TMP, 'T3LabAI', 'mcp_token.txt')
        with open(path) as f:
            good = f.read()
        try:
            with open(path, 'w') as f:
                f.write('also-wrong')
            with self.assertRaises(urllib.error.HTTPError):
                bridge._post_mcp(self.port, {'jsonrpc': '2.0', 'id': 1, 'method': 'ping'},
                                 'stale-token', timeout=5)
        finally:
            with open(path, 'w') as f:
                f.write(good)


class FileWatcherTests(unittest.TestCase):

    def setUp(self):
        from core import file_watcher as fw
        self.fw = fw
        os.makedirs(fw.T3LAB_DATA_DIR, exist_ok=True)
        for p in (fw.TASK_FILE, fw.RESULT_FILE, fw.TASK_PY_FILE, fw.RESULT_TXT):
            if os.path.exists(p):
                os.remove(p)
        self.w = fw.TaskFileWatcher()

    def tearDown(self):
        self.w.stop()
        time.sleep(0.6)
        self.fw.set_auto_start(False)

    def _wait_result(self, timeout=3.0):
        deadline = time.time() + timeout
        while time.time() < deadline:
            if os.path.exists(self.fw.RESULT_FILE):
                with open(self.fw.RESULT_FILE) as f:
                    return json.load(f)
            time.sleep(0.1)
        return None

    def test_off_by_default(self):
        self.assertFalse(self.fw.auto_start_enabled())
        self.assertFalse(self.w.start_if_enabled())
        self.assertFalse(self.w.is_running)
        self.fw.set_auto_start(True)
        self.assertTrue(self.w.start_if_enabled())

    def test_old_task_is_not_run(self):
        with open(self.fw.TASK_FILE, 'w') as f:
            json.dump({'task_id': 'left-over', 'code': 'result = "old"'}, f)
        self.w.start()
        self.assertIsNone(self._wait_result(1.5), 'a task written before start() ran')
        with open(self.fw.TASK_FILE, 'w') as f:
            json.dump({'task_id': 'new', 'code': 'result = "new"'}, f)
        res = self._wait_result()
        self.assertIsNotNone(res)
        self.assertEqual((res['task_id'], res['output']), ('new', 'new'))


class TelemetryTests(unittest.TestCase):

    def test_redaction(self):
        out = redact_arguments({
            'code': 'import os\nos.remove("x")',
            'output_folder': 'C:\\Users\\someone\\Client A\\Exports',
            'file_path': '/home/someone/proj/Model.rvt',
            'api_key': 'sk-123',
            'view_name': 'Level 1',
            'steps': [{'name': 'x', 'arguments': {'script': 'a', 'folder': 'D:\\a\\b'}}],
        })
        self.assertEqual(out['code'], '(code not sent, 2 lines)')
        self.assertEqual(out['output_folder'], 'Exports')
        self.assertEqual(out['file_path'], 'Model.rvt')
        self.assertEqual(out['api_key'], '(not sent)')
        self.assertEqual(out['view_name'], 'Level 1')
        self.assertEqual(out['steps'][0]['arguments'], {'script': '(code not sent, 1 lines)',
                                                        'folder': 'b'})

    def test_payload_has_no_code(self):
        sent = []
        orig = TelemetryService._dispatch_async
        # Sending is stubbed out, so tracking can be on for this one call.
        TelemetryService._dispatch_async = classmethod(lambda cls, p: sent.append(p))
        _set_settings(tracking_enabled=True)
        try:
            TelemetryService.record_mcp_call(
                'send_code_to_revit', {'code': 'SECRET_MARKER = 1'}, {'success': True}, 5)
        finally:
            _set_settings(tracking_enabled=False)
            TelemetryService._dispatch_async = orig
        self.assertEqual(len(sent), 1)
        self.assertNotIn('SECRET_MARKER', json.dumps(sent[0]))


def tearDownModule():
    shutil.rmtree(_TMP, ignore_errors=True)


if __name__ == '__main__':
    unittest.main(verbosity=1)
