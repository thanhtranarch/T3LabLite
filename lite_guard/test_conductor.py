"""core/conductor.py: fair order, wait budget, step aside, and Lite's server and
bridge wiring. From t3lab_dev's dev/test_conductor.py. docs/harness.md."""
import json
import os
import sys
import threading
import time
import unittest

LIB = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'lib')
sys.path.insert(0, os.path.abspath(LIB))

from core import conductor as C  # noqa: E402


class Clock(object):
    def __init__(self):
        self.now = 1000.0

    def __call__(self):
        return self.now


def ticket(agent, lane=None, method='tools/call', wait_s=None, tool='t'):
    return C.Ticket(method, agent, tool, None, lane, wait_s)


def enqueue(cond, t):
    """Put a ticket in the queue without blocking (ordering tests only)."""
    with cond._cond:
        cond._seq += 1
        t.seq = cond._seq
        t.arrived = cond._clock()
        cond._waiting.append(t)


class OrderTests(unittest.TestCase):
    def setUp(self):
        self.clock = Clock()
        self.cond = C.Conductor(clock=self.clock)

    def names(self):
        return [t.agent for t in self.cond.order()]

    def test_classes(self):
        enqueue(self.cond, ticket('auto', 'auto'))
        enqueue(self.cond, ticket('worker', 'task'))
        enqueue(self.cond, ticket('person', 'person'))
        enqueue(self.cond, ticket('lead', 'lead'))
        enqueue(self.cond, ticket('list', method='tools/list'))
        self.assertEqual(self.names(), ['list', 'person', 'lead', 'worker', 'auto'])

    def test_unknown_lane_is_worker_class(self):
        self.assertEqual(ticket('x').klass, C.CLASS_TASK)
        self.assertEqual(ticket('x', 'nonsense').klass, C.CLASS_TASK)

    def test_round_robin_within_class(self):
        self.cond._served = {'a': 5, 'b': 2}
        for name in ('a', 'a', 'b', 'c'):
            enqueue(self.cond, ticket(name, 'task'))
        # c never served, then b (served longer ago), then a's two calls in order
        self.assertEqual(self.names(), ['c', 'b', 'a', 'a'])

    def test_aging_lifts_a_waiting_call(self):
        enqueue(self.cond, ticket('auto', 'auto'))
        self.clock.now += C.AGING_S * 2 + 1
        enqueue(self.cond, ticket('lead', 'lead'))
        self.assertEqual(self.names(), ['auto', 'lead'])

    def test_ticket_for_reads_bridge_meta(self):
        t = C.ticket_for({'method': 'tools/call', 'params': {
            'name': 'create_wall', '_meta': {
                't3lab/agent': 'codex', 't3lab/lane': 'lead', 't3lab/task': 's1/a',
                't3lab/queue_wait_s': 45}}})
        self.assertEqual((t.agent, t.tool, t.task, t.klass, t.deadline),
                         ('codex', 'create_wall', 's1/a', C.CLASS_LEAD, 45.0))
        t = C.ticket_for({'method': 'tools/call', 'params': {'_meta': {
            't3lab/queue_wait_s': 'x'}}})
        self.assertIsNone(t.deadline)
        self.assertEqual(C.ticket_for({'method': 'tools/list'}).klass, C.CLASS_CONTROL)


class SlotTests(unittest.TestCase):
    def test_one_holder_and_fair_handover(self):
        cond = C.Conductor()
        first = ticket('lead', 'lead')
        cond.acquire(first)
        started = []
        threads = []
        for name, lane in (('auto', 'auto'), ('worker', 'task'), ('person', 'person')):
            t = ticket(name, lane)

            def run(t=t):
                with cond.slot(t):
                    started.append(t.agent)
                    time.sleep(0.02)
            th = threading.Thread(target=run)
            th.start()
            threads.append(th)
            time.sleep(0.05)
        self.assertEqual(len(cond.snapshot()['waiting']), 3)
        self.assertEqual(cond.snapshot()['running']['agent'], 'lead')
        cond.release(first)
        for th in threads:
            th.join(5)
        self.assertEqual(started, ['person', 'worker', 'auto'])
        self.assertIsNone(cond.snapshot()['running'])

    def test_busy_when_budget_runs_out(self):
        cond = C.Conductor()
        holder = ticket('claude', 'person', tool='create_wall')
        cond.acquire(holder)
        late = ticket('codex', 'task', wait_s=0.2)
        t0 = time.time()
        with self.assertRaises(C.Busy) as ctx:
            cond.acquire(late)
        self.assertLess(time.time() - t0, 2)
        info = ctx.exception.info
        self.assertEqual(info['error_code'], 'revit_busy')
        self.assertTrue(info['nothing_ran'] and info['retryable'])
        self.assertEqual(info['running']['agent'], 'claude')
        self.assertEqual(info['running']['tool'], 'create_wall')
        self.assertEqual(cond.snapshot()['waiting'], [])     # left the queue
        cond.release(holder)
        cond.acquire(ticket('codex', 'task', wait_s=0.2))    # free again

    def test_free_slot_never_busy(self):
        cond = C.Conductor()
        cond.acquire(ticket('a', 'task', wait_s=0))

    def test_step_aside_lets_others_run_then_comes_back_first(self):
        cond = C.Conductor()
        order = []
        approving = ticket('claude', 'person', tool='delete_elements')
        cond.acquire(approving)
        others = []
        for name in ('codex', 'gemini'):
            t = ticket(name, 'lead')

            def run(t=t):
                with cond.slot(t):
                    order.append(t.agent)
                    time.sleep(0.05)
            th = threading.Thread(target=run)
            th.start()
            others.append(th)
        time.sleep(0.05)
        with cond.step_aside():
            time.sleep(0.02)     # the person decides; codex gets Revit meanwhile
            self.assertEqual(cond.snapshot()['running']['agent'], 'codex')
        # back: ahead of gemini, which waited the whole time
        order.append('claude-back')
        self.assertIs(cond._holder, approving)
        cond.release(approving)
        for th in others:
            th.join(5)
        self.assertEqual(order, ['codex', 'claude-back', 'gemini'])

    def test_step_aside_without_slot_is_a_no_op(self):
        cond = C.Conductor()
        with cond.step_aside():
            pass
        self.assertIsNone(cond._holder)

    def test_summary_has_no_task_titles(self):
        cond = C.Conductor()
        t = C.Ticket('tools/call', 'codex', 'create_wall', 's1/a', 'task', None)
        cond.acquire(t)
        summary = cond.summary()
        self.assertEqual((summary['busy'], summary['agent'], summary['tool'], summary['waiting']),
                         (True, 'codex', 'create_wall', 0))
        self.assertNotIn('task', summary)

    def test_busy_result_shape(self):
        res = C.busy_result({'error': 'x', 'error_code': 'revit_busy'})
        self.assertTrue(res['isError'])
        self.assertEqual(json.loads(res['content'][0]['text'])['error_code'], 'revit_busy')



class WiringTests(unittest.TestCase):
    """Source checks: Lite's server uses the conductor (no Revit here)."""

    def read(self, rel):
        with open(os.path.join(LIB, rel), encoding='utf-8') as f:
            return f.read()

    def test_server_uses_conductor(self):
        src = self.read('core/server.py')
        self.assertNotIn('_MCP_REQUEST_LOCK', src, msg='old bare lock is back')
        self.assertIn('_CONDUCTOR.acquire(ticket)', src)
        self.assertIn('_CONDUCTOR.release(ticket)', src)
        self.assertIn("'queue': _CONDUCTOR.summary()", src)

    def test_conductor_is_ironpython_safe(self):
        import ast
        tree = ast.parse(self.read('core/conductor.py'))
        for node in ast.walk(tree):
            self.assertNotIsInstance(node, ast.JoinedStr)
            if isinstance(node, ast.arguments):
                self.assertEqual(node.kwonlyargs, [])
            if isinstance(node, ast.Raise):
                self.assertIsNone(node.cause)

    def test_bridge_and_conductor_agree_on_the_meta_keys(self):
        import importlib.util
        spec = importlib.util.spec_from_file_location(
            'bridge_conductor_test', os.path.join(LIB, 'core', 'bridge.py'))
        bridge = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(bridge)
        self.assertEqual((bridge.LANE_META, bridge.QUEUE_WAIT_META, bridge.TASK_META,
                          bridge.AGENT_META),
                         (C.LANE_META, C.WAIT_META, C.TASK_META, C.AGENT_META))
        self.assertEqual(set(C.LANE_CLASS), set(['lead', 'person', 'task', 'auto']))
        self.assertLess(bridge.QUEUE_WAIT_S, 60)          # under the client's tool timeout


if __name__ == '__main__':
    unittest.main()
