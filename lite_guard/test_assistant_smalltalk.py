#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Small talk is answered instantly and never swallows a request.

    python3 lite_guard/test_assistant_smalltalk.py

Ported with the fast-greeting change from t3lab_dev
(dev/test_assistant_routing.py). Plain CPython 3, standard library only.
"""
import io
import os
import sys
import unittest

sys.dont_write_bytecode = True

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
DIALOG = os.path.join(LIB, 'GUI', 'T3LabAssistantDialog.py')
sys.path.insert(0, LIB)

from Intelligence import nlu_engine as N  # noqa: E402


def _method(src, header):
    body = src.split(header, 1)[1]
    return body.split('\n    def ', 1)[0]


class SmallTalkNLU(unittest.TestCase):

    def test_greetings_are_instant(self):
        for q in (u"hello", u"hey there", u"Hi T3Lab", u"chào em",
                  u"xin chào mọi người", u"chào buổi sáng"):
            r = N.classify(q)
            self.assertEqual(r.get('intent'), 'greet', q)
            self.assertTrue(r.get('_instant'), q)
            self.assertTrue(N.is_pure_smalltalk(q), q)

    def test_a_greeting_never_swallows_a_request(self):
        """"hi, export the G sheets" was answered "Hello!"."""
        r = N.classify(u"hi, export the G sheets")
        self.assertEqual(r.get('intent'), 'export_direct', r)
        for q in (u"hi, tô đỏ tường", u"hello can you help me rename views"):
            r = N.classify(q) or {}
            self.assertNotEqual(r.get('intent'), 'greet', q)
            self.assertFalse(r.get('_instant'), q)
            self.assertFalse(N.is_pure_smalltalk(q), q)

    def test_bye_says_goodbye(self):
        for q, word in ((u"bye", u"Goodbye"), (u"tạm biệt", u"Tạm biệt"),
                        (u"good night", u"Goodbye")):
            self.assertIn(word, N.classify(q).get('message'), q)


class SmallTalkDialog(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        with io.open(DIALOG, encoding='utf-8') as f:
            cls.src = f.read()

    def test_small_talk_skips_the_worker_turn(self):
        body = _method(self.src, '    def _process_input(self):')
        i_fast = body.find('self._try_instant_smalltalk(raw, attached)')
        i_busy = body.find('self._set_busy(True)')
        i_thread = body.find('Thread(ThreadStart(_route))')
        self.assertTrue(0 <= i_fast < i_busy < i_thread,
                        (i_fast, i_busy, i_thread))
        fast = _method(self.src, '    def _try_instant_smalltalk(')
        for needle in ('is_pure_smalltalk(raw)', 'routing.is_continuation(',
                       "result.get('_instant')",
                       'routing.learned_pattern_wins(', '_forced_skill_id'):
            self.assertIn(needle, fast)
        for probe in ('has_local_llm(', 'has_api_key(', 'Thread('):
            self.assertNotIn(probe, fast)

    def test_provider_probes_wait_for_the_nlu(self):
        """has_local_llm() is an HTTP probe; it ran before the NLU."""
        body = _method(self.src, '    def _route_input(self, raw, attached):')
        i_run = body.find('self.Dispatcher.Invoke(Action(_run_nlu))')
        i_probe = body.find('has_local_llm()')
        i_used = body.find("use_local        = _llm['local']")
        self.assertGreater(i_run, 0)
        self.assertGreaterEqual(body.rfind('def _llm_ready():', 0, i_probe), 0)
        self.assertGreater(i_used, i_run)


if __name__ == '__main__':
    unittest.main(verbosity=2)
