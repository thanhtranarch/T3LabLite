# -*- coding: utf-8 -*-
"""Harness Studio (core/harness_studio.py + .html): the owner's seat on the
board, flows, the runner thread and the local HTTP API, through Lite's real
bridge. Linux, no Revit, no browser. docs/harness.md. From t3lab_dev's
dev/test_harness_studio.py; the callers are Lite bridge `current`s (no profiles)."""
import json
import os
import sys
import tempfile
import threading
import unittest
from urllib.error import HTTPError
from urllib.request import ProxyHandler, Request, build_opener

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_studio_appdata_')
os.environ['T3LAB_HARNESS_DIR'] = tempfile.mkdtemp(prefix='t3lab_studio_')

from core import harness_agents as agents   # noqa: E402
from core import harness_store as store     # noqa: E402
from core import harness_studio as studio   # noqa: E402
import core.bridge as B                     # noqa: E402

OPEN = build_opener(ProxyHandler({})).open


def _client(agent, transport='stdio'):
    return {'port': 1, 'client': agent, 'agent_explicit': True, 'validated': True,
            'transport': transport}


def cli_gateway(agent='cli'):
    """What `bridge.py harness studio` hands the Studio: calls of the CLI."""
    current = _client(agent, 'cli')

    def gateway(tool, arguments):
        resp = B._handle(B._call_request(1, tool, arguments), current, None)
        payload = B._tool_payload(resp)
        return payload if payload is not None else resp
    return gateway


class Case(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix='t3lab_studio_case_')
        os.environ['T3LAB_HARNESS_DIR'] = self.root
        self.board = store.Harness(self.root)
        self.studio = studio.Studio(store, agents, cli_gateway(), B.HARNESS_TOOL,
                                    't3lab_approval', token='secret')

    def code(self, fn, *args, **kwargs):
        with self.assertRaises(store.HarnessError) as ctx:
            fn(*args, **kwargs)
        return ctx.exception.payload['error_code']


class OwnerSeatTests(Case):
    def test_owner_plans_and_removes_without_a_seat(self):
        sid = self.board.open('claude-desktop', 'Goal', 'PRJ.rvt')['session']
        self.assertEqual(self.code(self.board.plan, 'stranger', [{'title': 'x'}]), 'not_a_member')
        out = self.board.plan('ignored', [{'key': 'a', 'title': 'A'},
                                         {'key': 'b', 'title': 'B', 'depends_on': ['a']}],
                              as_owner=True)
        self.assertEqual(out['planned'], ['a', 'b'])
        self.assertEqual(self.code(self.board.remove, 'x', 'a', as_owner=True), 'bad_dependency')
        self.assertTrue(self.board.remove('x', 'b', as_owner=True)['removed'])
        self.assertEqual([t['task'] for t in self.board.status(sid)['tasks']], ['a'])

    def test_started_task_cannot_be_removed(self):
        self.board.open('lead', 'Goal')
        self.board.plan('lead', [{'key': 'a', 'title': 'A'}])
        self.board.join('w')
        self.board.next('w')
        self.assertEqual(self.code(self.board.remove, 'lead', 'a'), 'task_started')
        self.assertEqual(self.code(self.board.remove, 'w', 'a'), 'role_denied')

    def test_owner_opens_despite_a_master_and_leads(self):
        store.save_config(self.root, master='claude-desktop')
        self.assertEqual(self.code(self.board.open, 'codex', 'Goal'), 'role_denied')
        out = self.board.open('codex', 'Goal', as_owner=True)
        self.assertEqual(self.board.status(out['session'])['lead'], store.OWNER)

    def test_owner_retries_notes_and_closes(self):
        sid = self.board.open('lead', 'Goal', review=False)['session']
        self.board.plan('lead', [{'key': 'a', 'title': 'A'}])
        self.board.join('w')
        self.board.next('w')
        self.board.report('w', 'a', 'failed', 'boom')
        self.assertEqual(self.board.retry('x', 'a', as_owner=True)['status'], 'todo')
        self.board.note('x', 'Use Level 1 only', as_owner=True)
        self.assertEqual(self.board.status(sid)['notes'][0]['by'], store.OWNER)
        self.assertTrue(self.board.close('x', as_owner=True)['closed'])

    def test_full_status_carries_what_the_canvas_edits(self):
        sid = self.board.open('lead', 'Goal')['session']
        self.board.plan('lead', [{'key': 'a', 'title': 'A', 'detail': 'Do it',
                                  'scope': {'levels': ['Level 1']}}])
        task = self.board.status(sid, full=True)['tasks'][0]
        self.assertEqual((task['detail'], task['scope']['levels']), ('Do it', ['level 1']))
        self.assertNotIn('detail', self.board.status(sid)['tasks'][0])
        self.assertEqual(self.board.sessions()[0]['session'], sid)


class GatewayTests(Case):
    def test_owner_flag_only_counts_from_the_cli(self):
        sid = self.board.open('claude-desktop', 'Goal')['session']
        args = {'action': 'plan', 'owner': True, 'session': sid, 'tasks': [{'title': 'x'}]}
        out = B._tool_payload(B._handle(B._call_request(1, B.HARNESS_TOOL, args),
                                        _client('codex'), None))
        self.assertEqual(out['error_code'], 'not_a_member')
        self.assertEqual(self.studio.call('plan', {'session': sid, 'tasks': [{'title': 'x'}]})
                         ['planned'], ['t1'])

    def test_studio_actions_change_the_board(self):
        sid = self.studio.act('open', {'goal': 'Audit doors', 'revit_document': 'PRJ.rvt'})['session']
        self.studio.act('plan', {'session': sid, 'tasks': [
            {'key': 'a', 'title': 'A'}, {'key': 'b', 'title': 'B', 'depends_on': ['a']}]})
        self.board.join('codex')
        self.board.next('codex')
        self.board.report('codex', 'a', 'done', '3 doors')
        self.assertEqual(self.studio.act('review', {'session': sid, 'task': 'a',
                                                    'verdict': 'accept'})['status'], 'done')
        self.assertTrue(self.studio.act('remove', {'session': sid, 'task': 'b'})['removed'])
        state = self.studio.state()
        self.assertEqual((state['session'], state['board']['lead']), (sid, store.OWNER))
        self.assertEqual(state['board']['tasks'][0]['summary'], '3 doors')
        self.assertEqual(state['stats']['counts'], {'done': 1})
        with self.assertRaises(studio.StudioError):
            self.studio.act('run', {})          # agents start through /api/run only
        with self.assertRaises(studio.StudioError):
            self.studio.act('setup', {'master': 'x'})

    def test_agents_are_added_as_a_person_and_reserved_names_refused(self):
        out = self.studio.act('agent_add', {'name': 'console-qa', 'config': {
            'platform': 'command', 'command': ['echo', '{prompt}']}})
        self.assertNotIn('error', out)
        self.assertIn('console-qa', self.studio.state()['agents'])
        self.assertEqual(self.studio.act('agent_add', {'name': 'codex', 'config': {
            'platform': 'command', 'command': ['x', '{prompt}']}})['error_code'],
                         'reserved_name')


class FlowTests(Case):
    def test_save_apply_delete(self):
        sid = self.studio.act('open', {'goal': 'Audit doors'})['session']
        self.studio.act('plan', {'session': sid, 'tasks': [
            {'key': 'a', 'title': 'A', 'assignee': 'codex'},
            {'key': 'b', 'title': 'B', 'depends_on': ['a'], 'scope': {'write': False}}]})
        self.board.join('codex')
        self.board.next('codex')                    # a flow never carries status or owner
        self.assertEqual(self.studio.flow('save', {'name': 'Door-Audit', 'session': sid}),
                         {'saved': 'door-audit', 'tasks': 2})
        flow = self.studio.flow('get', {'name': 'door-audit'})
        self.assertEqual(sorted(flow['tasks'][0]), ['assignee', 'key', 'scope', 'title'])
        self.assertEqual(self.studio.state()['flows'][0]['tasks'], 2)
        out = self.studio.flow('apply', {'name': 'door-audit', 'revit_document': 'B.rvt'})
        self.assertNotEqual(out['session'], sid)
        again = self.board.status(out['session'], full=True)
        self.assertEqual((again['goal'], again['document']), ('Audit doors', 'B.rvt'))
        self.assertEqual([(t['task'], t['status'], t.get('owner')) for t in again['tasks']],
                         [('a', 'todo', None), ('b', 'todo', None)])
        self.assertEqual(again['tasks'][1]['depends_on'], ['a'])
        self.studio.flow('delete', {'name': 'door-audit'})
        self.assertEqual(self.studio.state()['flows'], [])

    def test_names_cannot_leave_the_flows_folder(self):
        for name in ('../x', 'a/b', '', 'x' * 41, '..'):
            with self.assertRaises(studio.StudioError):
                studio.save_flow(self.root, name, 'g', [{'key': 'a', 'title': 'A'}])
        with self.assertRaises(studio.StudioError):
            studio.save_flow(self.root, 'ok', 'g', [])


class RunnerTests(unittest.TestCase):
    def run_with(self, answers, **kwargs):
        calls = []

        def call(action, args):
            calls.append((action, dict(args)))
            return answers.pop(0) if answers else {'ran': [{'agent': 'a', 'all_done': True}]}
        runs = studio.Runs(call, wait_s=0.01, **kwargs)
        runs.start('s1', 2)
        runs.join(5)
        return runs.view(), calls

    def test_rounds_until_the_board_is_done(self):
        view, calls = self.run_with([
            {'ran': [{'agent': 'a', 'task': 't1', 'status': 'done', 'summary': 'ok'}]},
            {'ran': [{'agent': 'a', 'task': None, 'all_done': False}]},
            {'ran': [{'agent': 'a', 'task': None, 'all_done': True}]}])
        self.assertEqual((view['running'], view['rounds'], view['ended']),
                         (False, 3, 'Every task is done.'))
        self.assertEqual([i['task'] for i in view['log']], ['t1'])
        self.assertEqual(calls[0], ('run', {'session': 's1', 'parallel': 2}))

    def test_ends_with_the_reason(self):
        self.assertIn('No launched agent', self.run_with([{'ran': []}])[0]['ended'])
        self.assertEqual(self.run_with([{'error': 'Revit is closed.'}])[0]['ended'],
                         'Revit is closed.')
        view, _calls = self.run_with([{'ran': [{'agent': 'a', 'task': None}]}] * 9, max_idle=2)
        self.assertIn('Nothing moved', view['ended'])

    def test_one_run_at_a_time_and_stop(self):
        gate = threading.Event()

        def call(action, args):
            gate.wait(5)
            return {'ran': [{'agent': 'a', 'task': 't', 'status': 'done'}]}
        runs = studio.Runs(call)
        runs.start(None)
        with self.assertRaises(studio.StudioError):
            runs.start(None)
        self.assertTrue(runs.stop()['stopping'])
        gate.set()
        runs.join(5)
        self.assertEqual(runs.view()['ended'], 'Stopped.')


class HttpTests(Case):
    def setUp(self):
        Case.setUp(self)
        self.httpd = studio.make_server(self.studio, 0)
        self.base = 'http://127.0.0.1:{}'.format(self.httpd.server_address[1])
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()
        self.addCleanup(self.httpd.server_close)
        self.addCleanup(self.httpd.shutdown)

    def ask(self, path, body=None, token='secret', headers=None):
        headers = dict({studio.TOKEN_HEADER: token} if token else {}, **(headers or {}))
        data = None if body is None else json.dumps(body).encode('utf-8')
        try:
            with OPEN(Request(self.base + path, data=data, headers=headers), timeout=5) as r:
                return r.status, r.read(), r.headers
        except HTTPError as ex:
            return ex.code, ex.read(), ex.headers

    def test_page_is_served_and_self_contained(self):
        code, body, headers = self.ask('/', token=None)
        page = body.decode('utf-8')
        self.assertEqual(code, 200)
        self.assertIn('Harness Studio', page)
        self.assertIn("frame-ancestors 'none'", headers['Content-Security-Policy'])
        for banned in ('innerHTML', 'outerHTML', 'insertAdjacentHTML', 'document.write',
                       'src="http', 'href="http', '@import'):
            self.assertNotIn(banned, page.replace('never innerHTML', ''))

    def test_token_and_host_guard_the_api(self):
        self.assertEqual(self.ask('/api/state', token=None)[0], 401)
        self.assertEqual(self.ask('/api/state', token='wrong')[0], 401)
        self.assertEqual(self.ask('/api/act', {'action': 'open', 'args': {'goal': 'x'}},
                                  token=None)[0], 401)
        self.assertEqual(self.ask('/api/state', headers={'Host': 'evil.example'})[0], 403)
        self.assertEqual(self.ask('/api/state', headers={'Origin': 'https://evil.example'})[0], 403)
        self.assertEqual(self.board.sessions(), [])

    def test_state_act_and_errors_round_trip(self):
        code, body, _h = self.ask('/api/act', {'action': 'open', 'args': {'goal': 'Audit'}})
        sid = json.loads(body)['session']
        out = json.loads(self.ask('/api/act', {'action': 'plan', 'args': {
            'session': sid, 'tasks': [{'title': 'A', 'depends_on': ['nope']}]}})[1])
        self.assertEqual(out['error_code'], 'bad_dependency')
        self.assertEqual(json.loads(self.ask('/api/act', {'action': 'boom'})[1])['error_code'],
                         'bad_arguments')
        self.assertEqual(self.ask('/api/act', [1])[0], 400)
        state = json.loads(self.ask('/api/state?session=' + sid)[1])
        self.assertEqual((state['session'], state['board']['goal']), (sid, 'Audit'))
        self.assertEqual(self.ask('/api/nope', {})[0], 404)

    def test_a_running_studio_is_found_again(self):
        with open(os.path.join(self.root, studio.RUNNING_FILE), 'w') as f:
            json.dump({'port': self.httpd.server_address[1], 'token': 'secret'}, f)
        self.assertEqual(studio._running(self.root), self.base + '/?k=secret')
        with open(os.path.join(self.root, studio.RUNNING_FILE), 'w') as f:
            json.dump({'port': self.httpd.server_address[1], 'token': 'stale'}, f)
        self.assertEqual(studio._running(self.root), '')


class McpTemplateTests(Case):
    """AI apps reach templates and saved agents through t3lab_harness (owner, 2026-11-20)."""

    def call(self, agent, **args):
        return B._tool_payload(B._handle(B._call_request(1, B.HARNESS_TOOL, args),
                                         _client(agent), None))

    def test_person_facing_app_saves_and_reuses(self):
        sid = self.call('claude-desktop', action='open', goal='Audit doors')['session']
        self.call('claude-desktop', action='plan', session=sid, tasks=[
            {'key': 'a', 'title': 'List'}, {'key': 'b', 'title': 'Fix', 'depends_on': ['a']}])
        self.assertEqual(self.call('claude-desktop', action='flow_save', name='doors',
                                   session=sid)['saved'], 'doors')
        self.assertEqual(self.call('codex', action='flows')['flows'][0]['tasks'], 2)
        out = self.call('claude-desktop', action='flow_apply', name='doors', goal='Doors, B.rvt',
                        revit_document='B.rvt')
        self.assertNotEqual(out['session'], sid)
        self.assertEqual([t['task'] for t in out['tasks']], ['a', 'b'])
        self.assertEqual(self.board.status(out['session'])['lead'], 'claude-desktop')
        added = self.call('claude-desktop', action='agent_add', name='qa-bot', config={
            'platform': 'command', 'command': ['claude', '-p', '{prompt}'], 'role': 'reviewer'})
        self.assertNotIn('error', added)
        self.assertIn('qa-bot', self.call('claude-desktop', action='agents')['agents'])
        self.assertIn('templates', self.call('codex', action='help'))

    def test_unattended_agent_cannot_save_or_open(self):
        sid = self.call('claude-desktop', action='open', goal='Audit doors')['session']
        self.call('claude-desktop', action='plan', session=sid, tasks=[{'key': 'a', 'title': 'A'}])
        self.call('claude-desktop', action='flow_save', name='doors', session=sid)
        for args in ({'action': 'flow_save', 'name': 'x', 'session': sid},
                     {'action': 'flow_delete', 'name': 'doors'},
                     {'action': 'flow_apply', 'name': 'doors'},
                     {'action': 'agent_add', 'name': 'x', 'config': {'platform': 'mcp'}}):
            self.assertEqual(self.call('openclaw', **args)['error_code'], 'role_denied', args)
        self.assertEqual(self.call('openclaw', action='flow_get', name='doors')['name'], 'doors')


class WiringTests(unittest.TestCase):
    def test_bridge_ships_the_studio(self):
        import bridge_deploy
        names = [dst for _src, dst in bridge_deploy.COMPANIONS]
        self.assertIn('harness_studio.py', names)
        self.assertIn(studio.PAGE_FILE, names)
        self.assertIsNotNone(B._hst)

    def test_ribbon_button_opens_the_studio(self):
        panel = os.path.join(REPO, 'T3Lab.tab', 'Support.panel')
        with open(os.path.join(panel, 'bundle.yaml'), encoding='utf-8') as f:
            self.assertIn('- Harness', f.read())        # pyRevit builds only listed buttons
        with open(os.path.join(panel, 'Harness.pushbutton', 'script.py'), encoding='utf-8') as f:
            script = f.read()
        self.assertTrue(script.startswith('#! python3'))
        self.assertIn('init_cpython_paths()', script)       # Lite's bootstrap + click tracking
        self.assertIn('open_studio()', script)

    def test_revit_side_launchers(self):
        from Intelligence import harness_client
        from Services import harness_settings as settings
        seen = []
        settings.open_studio('py', 'bridge.py', popen=lambda argv, **kw: seen.append(argv))
        self.assertEqual(seen[0][:4], ['py', 'bridge.py', 'harness', 'studio'])
        self.assertTrue(seen[0][4].startswith('--idle-exit='))
        started = []
        text = harness_client.handle('/harness studio', studio=lambda: started.append(1))
        self.assertEqual(started, [1])
        self.assertIn('browser', text)

        def broken():
            raise OSError('no python')
        self.assertIn('Could not start', harness_client.handle('/harness studio', studio=broken))

    def test_bridge_studio_calls_only_the_harness(self):
        """`bridge.py harness studio`: the page acts as the CLI, and Lite has no
        approval queue for it to reach."""
        seen = {}
        real = B._hst.serve

        def serve(store_, agents_, gateway, harness_tool, approval_tool, **kwargs):
            seen.update(gateway=gateway, kwargs=kwargs)
            return 0
        B._hst.serve = serve
        self.addCleanup(setattr, B._hst, 'serve', real)
        self.assertEqual(B._harness_studio(['--no-open', '--idle-exit=10'],
                                           _client('cli', 'cli'), None), 0)
        self.assertEqual((seen['kwargs']['open_browser'], seen['kwargs']['idle_exit_min'],
                          seen['kwargs']['approvals']()), (False, 10.0, []))
        gateway = seen['gateway']
        self.assertEqual(gateway('t3lab_approval', {'action': 'list'})['error_code'],
                         'bad_arguments')
        self.assertEqual(gateway('set_parameter', {})['error_code'], 'bad_arguments')
        opened = gateway('t3lab_harness', {'action': 'open', 'goal': 'g', 'owner': True})
        self.assertEqual(B._hb.Harness().status(opened['session'])['lead'], store.OWNER)


if __name__ == '__main__':
    unittest.main(verbosity=1)
