# -*- coding: utf-8 -*-
"""The T3Lab Assistant's /harness commands (lib/Intelligence/harness_client.py)
and their wiring in the Assistant. Lite: the pane looks and reviews, it never
takes a task. Linux, no Revit. docs/harness.md."""
import json
import os
import sys
import tempfile
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_harness_client_appdata_')

from Intelligence import harness_client as H     # noqa: E402


class Runner(object):
    """Stands in for Services.harness_settings.run_bridge."""

    def __init__(self, *answers):
        self.answers = list(answers)
        self.calls = []

    def __call__(self, argv, timeout=None):
        self.calls.append(argv)
        return self.answers.pop(0) if self.answers else (0, {})


BOARD = {'sessions': [{'session': 's1', 'goal': 'Audit doors', 'document': 'PRJ.rvt',
                       'counts': {'review': 1, 'todo': 1},
                       'tasks': [{'task': 'a', 'title': 'List doors', 'status': 'review',
                                  'owner': 'codex'},
                                 {'task': 'b', 'title': 'Fix marks', 'status': 'todo'}]}]}


class Commands(unittest.TestCase):
    def test_parse(self):
        self.assertEqual(H.parse('/harness'), ('status', ''))
        self.assertEqual(H.parse('  /harness Redo a  also level 2'), ('redo', 'a  also level 2'))
        self.assertIsNone(H.parse('/harnessx'))
        self.assertIsNone(H.parse('open the harness'))
        self.assertIsNone(H.handle('hello'))

    def test_board_text(self):
        text = H.handle('/harness', board=BOARD)
        self.assertIn('**Audit doors**', text)
        self.assertIn('GPT · List doors (waiting for review)', text)
        self.assertIn('/harness accept', text)
        self.assertIn('No harness session is open (master: claude-desktop)',
                      H.handle('/harness', board={'master': 'claude-desktop'}))

    def test_review_as_the_person_at_the_machine(self):
        runner = Runner((0, {'status': 'todo'}))
        self.assertEqual(H.handle('/harness redo a also level 2', runner=runner),
                         '`a` is now todo.')
        argv = runner.calls[0]
        self.assertEqual(argv[:3], ['harness', 'review', '--agent=t3lab-assistant'])
        self.assertEqual(json.loads(argv[argv.index('--json') + 1]),
                         {'task': 'a', 'verdict': 'redo', 'owner': True, 'note': 'also level 2'})
        self.assertIn('Which task', H.handle('/harness accept', runner=Runner()))

    def test_errors_are_plain_sentences(self):
        refused = Runner((1, {'error': 'Task a is not waiting for review.',
                              'error_code': 'not_in_review', 'hint': 'action=status'}))
        self.assertEqual(H.handle('/harness accept a', runner=refused),
                         'Task a is not waiting for review. action=status')
        self.assertIn('did not answer', H.handle('/harness accept a',
                                                 runner=Runner((None, 'x'))))
        self.assertIn('Unknown /harness command', H.handle('/harness dance'))

    def test_the_pane_never_takes_a_task(self):
        runner = Runner()
        for action in H.MEMBER_ACTIONS:
            self.assertIn('not in T3Lab Lite', H.handle('/harness ' + action, runner=runner))
        self.assertEqual(runner.calls, [])

    def test_the_bridge_reads_the_pane_as_a_person(self):
        from core import bridge as B
        self.assertIn(H.AGENT, B.HUMAN_AGENTS)          # owner=true review is accepted


class Wiring(unittest.TestCase):
    def read(self, *parts):
        with open(os.path.join(LIB, *parts), encoding='utf-8') as f:
            return f.read()

    def test_command_is_routed_before_the_llm(self):
        dialog = self.read('GUI', 'T3LabAssistantDialog.py')
        self.assertIn('self._commands._try_harness_command(raw)', dialog)
        self.assertIn("u'harness'.startswith(q)", dialog)       # in the / suggestions
        commands = self.read('GUI', 'assistant', 'commands.py')
        self.assertIn('def _try_harness_command(self, raw):', commands)
        self.assertIn('harness_client.handle(raw)', commands)


if __name__ == '__main__':
    unittest.main()
