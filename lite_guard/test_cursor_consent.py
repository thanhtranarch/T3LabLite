# -*- coding: utf-8 -*-
"""Consent on the Agent Cursors bar instead of a Revit dialog, and the
one-time offer to turn the cursors on (core/agent_cursors.py ask / answer /
offer, core/server.py _ask_on_bar, core/agent_cursors_overlay.py).

Revit and WPF are not loaded: the Allow / Deny flow is driven through the
same callbacks the bar uses. The Revit side must be checked in Revit.
Plain CPython 3."""
import os
import sys
import tempfile
import threading
import unittest

sys.dont_write_bytecode = True
_TMP = tempfile.mkdtemp(prefix='t3lab_consent_test_')
os.environ['APPDATA'] = _TMP
os.environ['HOME'] = _TMP
os.environ['USERPROFILE'] = _TMP

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
sys.path.insert(0, LIB)

from core import agent_cursors as AC   # noqa: E402
from core import paths                 # noqa: E402
from core import server as S           # noqa: E402

paths.save_settings({'tracking_enabled': False})


def read(*parts):
    with open(os.path.join(LIB, *parts), encoding='utf-8') as f:
        return f.read()


class FakeEvent(object):
    def __init__(self):
        self.raised = 0

    def Raise(self):
        self.raised += 1


class FakeHandler(object):
    def __init__(self):
        import queue
        self.tasks = queue.Queue()
        self.server = type('Srv', (), {})()
        self.server._external_event = FakeEvent()


class BarConsentTests(unittest.TestCase):
    def setUp(self):
        self.asked = []
        self.real_ask = AC.ask

        def fake_ask(app, agent, tool, args, main, content, expanded, always, on_answer):
            self.asked.append({'agent': agent, 'tool': tool, 'main': main,
                               'always': always, 'on_answer': on_answer})
            return 'ask{}'.format(len(self.asked))
        AC.ask = fake_ask
        self.addCleanup(setattr, AC, 'ask', self.real_ask)
        self.handler = FakeHandler()

    def task(self, tool='send_code_to_revit'):
        t = S._ToolTask(tool, {'code': 'x', 'description': 'Number the risers'},
                        confirm=('Run code?', 'Number the risers', 'x'), agent='claude-ai')
        self.assertTrue(t.claim('ui'))
        return t

    def test_allow_requeues_the_task_and_raises_the_event(self):
        t = self.task()
        self.assertTrue(S._ask_on_bar(self.handler, None, t))
        self.assertEqual(t.ask_rid, 'ask1')
        self.assertTrue(t.awaiting_user)
        self.assertFalse(t.done.is_set())
        self.assertEqual(self.asked[0]['main'], 'Run code?')
        self.asked[0]['on_answer'](True, False)
        self.assertTrue(t.answered)
        self.assertIs(self.handler.tasks.get_nowait(), t)
        self.assertEqual(self.handler.server._external_event.raised, 1)
        self.assertFalse(t.done.is_set())             # runs on the UI thread next
        self.assertEqual(t.claimed_by, 'ui')

    def test_deny_finishes_the_task_without_running_it(self):
        t = self.task()
        S._ask_on_bar(self.handler, None, t)
        self.asked[0]['on_answer'](False, False)
        self.assertTrue(t.done.is_set())
        self.assertTrue(t.result['declined_by_user'])
        self.assertTrue(self.handler.tasks.empty())

    def test_allow_after_the_caller_gave_up_never_runs(self):
        t = self.task()
        S._ask_on_bar(self.handler, None, t)
        self.assertTrue(t.give_up())                   # caller timed out while asking
        self.asked[0]['on_answer'](True, False)
        self.assertTrue(t.done.is_set())
        self.assertIn('too late', t.result['error'])
        self.assertTrue(self.handler.tasks.empty())

    def test_no_bar_falls_back_to_the_dialog(self):
        AC.ask = lambda *a, **k: ''
        t = self.task()
        self.assertFalse(S._ask_on_bar(self.handler, None, t))
        self.assertFalse(t.awaiting_user)

    def test_always_allow_text_reaches_the_bar(self):
        t = self.task()
        S._ask_on_bar(self.handler, None, t)
        self.assertTrue(self.asked[0]['always'])       # mcp_policy.dont_ask_text for code


class CursorApiTests(unittest.TestCase):
    def test_ask_needs_cursors_on(self):
        AC.set_enabled(False)
        self.assertEqual(AC.ask(None, 'claude-ai', 'delete_element', {}, 'm', 'c',
                                on_answer=lambda a, b: None), '')

    def test_answer_runs_once_and_cancel_drops(self):
        got = []
        AC._asks['askX'] = lambda approved, always: got.append((approved, always))
        AC.answer('askX', True, True)
        AC.answer('askX', False)
        self.assertEqual(got, [(True, True)])
        AC._asks['askY'] = lambda approved, always: got.append('never')
        AC.cancel_ask('askY')
        AC.answer('askY', True)
        self.assertEqual(got, [(True, True)])
        self.assertEqual(AC.pending_asks(), 0)

    def test_offer(self):
        main, content = AC.offer_text('claude-ai')
        self.assertEqual(main, 'Follow Claude in Revit?')
        self.assertIn('asks for your OK on its bar', content)
        AC._state['offered'] = False
        self.assertFalse(AC.offer(None, ''))           # not an MCP call
        AC.set_enabled(True)
        self.assertFalse(AC.offer(None, 'claude-ai'))  # already on
        AC.set_enabled(False)
        AC._state['offered'] = True
        self.assertFalse(AC.offer(None, 'claude-ai'))  # once per Revit session


class WiringTests(unittest.TestCase):
    def test_execute_asks_on_the_bar_before_the_dialog(self):
        src = read('core', 'server.py')
        body = src.split('        def Execute(self, app):', 1)[1].split('        def GetName', 1)[0]
        self.assertLess(body.index('agent_cursors.offer(app, task.agent)'),
                        body.index('_ask_on_bar(self, app, task)'))
        self.assertLess(body.index('_ask_on_bar(self, app, task)'),
                        body.index('_ask_user_to_confirm(task)'))
        self.assertIn('if task.answered:', body)
        self.assertIn("if task.claimed_by != 'ui':", body)

    def test_timeout_takes_the_card_off(self):
        self.assertIn('agent_cursors.cancel_ask(task.ask_rid)', read('core', 'server.py'))

    def test_overlay_keeps_the_question_in_sight(self):
        src = read('core', 'agent_cursors_overlay.py')
        front = src.split('    def _on_front_tick(self, sender, e):', 1)[1].split('\ndef ', 1)[0]
        self.assertIn('if window is self.hud and self._asks:', front)
        hide = src.split('    def hide(self):', 1)[1].split('    def ', 1)[0]
        self.assertIn('self._answer(item[\'rid\'], False)', hide)
        self.assertIn("'answer': answer", read('core', 'agent_cursors.py'))

    def test_glyphs_exist_in_segoe_mdl2(self):
        for glyph in ('\\uE7BA', '\\uE70D', '\\uE73E', '\\uE711'):
            self.assertIn(glyph, read('core', 'agent_cursors_overlay.py'))


if __name__ == '__main__':
    unittest.main()
