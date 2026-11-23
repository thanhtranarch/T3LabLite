# -*- coding: utf-8 -*-
"""T3Lab BIM Harness: the task board (core/harness_store.py), agents from other
platforms (core/harness_agents.py) and the t3lab_harness tool of Lite's bridge
(core/bridge.py: identity, write check, stamps, CLI). Runs on Linux, no Revit,
no network: the Anthropic client and the webhook are fakes. docs/harness.md.
Store and agent tests are t3lab_dev's dev/test_harness.py; the bridge tests are
Lite's own (no profiles, tiers or approval queue)."""
import json
import os
import subprocess
import sys
import tempfile
import time
import types
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_harness_appdata_')
os.environ['T3LAB_HARNESS_DIR'] = tempfile.mkdtemp(prefix='t3lab_harness_')

from core import harness_agents as agents   # noqa: E402
from core import harness_store as store     # noqa: E402
import core.bridge as B                     # noqa: E402


class Clock(object):
    def __init__(self):
        self.t = 1000.0

    def __call__(self):
        return self.t


class BoardCase(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp(prefix='t3lab_harness_case_')
        self.clock = Clock()
        self.board = store.Harness(self.root, now=self.clock)

    def code(self, fn, *args, **kwargs):
        with self.assertRaises(store.HarnessError) as ctx:
            fn(*args, **kwargs)
        return ctx.exception.payload['error_code']

    def open_with(self, tasks, review=True):
        sid = self.board.open('lead', 'Audit levels 1-2', 'PRJ.rvt', review)['session']
        self.board.plan('lead', tasks)
        return sid


class StoreTests(BoardCase):
    def test_open_plan_next_report_review(self):
        sid = self.open_with([{'key': 'a', 'title': 'Check doors'}])
        self.assertEqual(self.board.join('w1')['role'], 'worker')
        offer = self.board.next('w1')
        self.assertEqual(offer['task']['task'], 'a')
        self.assertEqual(offer['goal'], 'Audit levels 1-2')
        self.assertEqual(offer['document'], 'PRJ.rvt')
        self.assertEqual(self.board.report('w1', 'a', 'done', '12 doors ok')['status'], 'review')
        self.assertEqual(self.board.review('lead', 'a', 'accept')['status'], 'done')
        self.assertTrue(self.board.next('w1')['all_done'])
        status = self.board.status(sid)
        self.assertEqual(status['counts'], {'done': 1})
        self.assertEqual(status['tasks'][0]['summary'], '12 doors ok')

    def test_session_is_found_without_an_id(self):
        self.assertEqual(self.code(self.board.next, 'w1'), 'session_required')
        self.assertEqual(self.board.status(), {'sessions': []})
        first = self.board.open('lead', 'g1')['session']
        self.assertEqual(self.board.join('w1')['session'], first)
        second = self.board.open('lead2', 'g2')['session']
        # w1 is a member of one session only: that one is theirs.
        self.assertEqual(self.board.status(member='w1')['session'], first)
        self.assertEqual(self.code(self.board.next, 'stranger'), 'session_required')
        self.assertEqual(len(self.board.status(member='stranger')['sessions']), 2)
        self.assertEqual(self.board.status(second)['goal'], 'g2')

    def test_roles(self):
        self.open_with([{'key': 'a', 'title': 'x'}])
        self.board.join('w1')
        self.assertEqual(self.code(self.board.plan, 'w1', [{'title': 'y'}]), 'role_denied')
        self.assertEqual(self.code(self.board.join, 'w1', 'lead'), 'role_denied')
        self.assertEqual(self.code(self.board.next, 'nobody'), 'not_a_member')
        self.assertEqual(self.code(self.board.close, 'w1'), 'role_denied')
        self.board.next('w1')
        self.board.report('w1', 'a')
        self.assertEqual(self.code(self.board.review, 'w1', 'a', 'accept'), 'role_denied')
        self.board.join('w1', 'reviewer')
        self.assertEqual(self.code(self.board.review, 'w1', 'a', 'accept'), 'role_denied')
        self.board.join('r1', 'reviewer')
        self.assertEqual(self.board.review('r1', 'a', 'accept')['status'], 'done')

    def test_dependencies(self):
        self.open_with([{'key': 'a', 'title': 'first'},
                        {'key': 'b', 'title': 'second', 'depends_on': ['a'],
                         'scope': {'write': False}}], review=False)
        self.board.join('w1')
        self.board.join('w2')
        self.assertEqual(self.board.next('w1')['task']['task'], 'a')
        wait = self.board.next('w2')
        self.assertIsNone(wait['task'])
        self.assertEqual(wait['waiting'], [{'task': 'b', 'why': 'waits for a'}])
        self.board.report('w1', 'a')
        self.assertEqual(self.board.next('w2')['task']['task'], 'b')

    def test_bad_plans(self):
        self.board.open('lead', 'g')
        self.assertEqual(self.code(self.board.plan, 'lead', []), 'bad_arguments')
        self.assertEqual(self.code(self.board.plan, 'lead', [{'key': 'a'}]), 'bad_arguments')
        self.assertEqual(self.code(self.board.plan, 'lead',
                                   [{'key': 'a', 'title': 'x', 'depends_on': ['zz']}]),
                         'bad_dependency')
        self.assertEqual(self.code(self.board.plan, 'lead',
                                   [{'key': 'a', 'title': 'x', 'depends_on': ['b']},
                                    {'key': 'b', 'title': 'y', 'depends_on': ['a']}]),
                         'bad_dependency')
        self.assertEqual(self.board.status()['tasks'], [])      # a refused plan saves nothing
        many = [{'title': 't'} for _ in range(store.MAX_TASKS + 1)]
        self.assertEqual(self.code(self.board.plan, 'lead', many), 'too_many_tasks')

    def test_replan_only_touches_todo_tasks(self):
        self.open_with([{'key': 'a', 'title': 'old'}, {'key': 'b', 'title': 'keep'}])
        self.board.plan('lead', [{'key': 'a', 'title': 'new'}])
        self.assertEqual([t['title'] for t in self.board.status()['tasks']], ['new', 'keep'])
        self.board.join('w1')
        self.board.next('w1')
        self.assertEqual(self.code(self.board.plan, 'lead', [{'key': 'a', 'title': 'again'}]),
                         'task_started')

    def test_write_scopes_do_not_overlap(self):
        self.open_with([
            {'key': 'l1', 'title': 'walls L1', 'scope': {'levels': ['Level 1']}},
            {'key': 'l1b', 'title': 'doors L1', 'scope': {'levels': ['level 1']}},
            {'key': 'l2', 'title': 'walls L2', 'scope': {'levels': ['Level 2']}},
            {'key': 'read', 'title': 'count rooms', 'scope': {'write': False}},
            {'key': 'all', 'title': 'purge'}], review=False)
        for name in ('w1', 'w2', 'w3', 'w4'):
            self.board.join(name)
        self.assertEqual(self.board.next('w1')['task']['task'], 'l1')
        self.assertEqual(self.board.next('w2')['task']['task'], 'l2')    # l1b waits
        self.assertEqual(self.board.next('w3')['task']['task'], 'read')
        wait = self.board.next('w4')
        self.assertIsNone(wait['task'])
        self.assertEqual(sorted(w['task'] for w in wait['waiting']), ['all', 'l1b'])
        self.assertIn('same scope as l1', wait['waiting'][0]['why'])
        self.board.report('w1', 'l1')
        self.assertEqual(self.board.next('w4')['task']['task'], 'l1b')

    def test_scope_rules(self):
        conflict = store.scopes_conflict
        clean = store.clean_scope
        self.assertTrue(conflict(clean({}), clean({'levels': ['L1']})))
        self.assertFalse(conflict(clean({'write': False}), clean({})))
        self.assertFalse(conflict(clean({'document': 'A.rvt'}), clean({'document': 'B.rvt'})))
        self.assertTrue(conflict(clean({'document': 'a.rvt'}), clean({}), 'A.rvt'))
        self.assertFalse(conflict(clean({'categories': ['Walls'], 'levels': ['L1']}),
                                  clean({'categories': ['Doors'], 'levels': ['L1']})))
        self.assertTrue(conflict(clean({'categories': ['Walls']}), clean({'levels': ['L1']})))
        self.assertFalse(conflict(clean({'element_ids': [1, 2]}), clean({'element_ids': [3]})))
        self.assertEqual(clean({'levels': 'L1', 'junk': 1}), {'write': True, 'levels': ['l1']})

    def test_one_task_per_member_and_ownership(self):
        self.open_with([{'key': 'a', 'title': 'x', 'scope': {'write': False}},
                        {'key': 'b', 'title': 'y', 'scope': {'write': False}}])
        self.board.join('w1')
        self.board.join('w2')
        self.board.next('w1')
        again = self.board.next('w1')
        self.assertTrue(again['already_claimed'])
        self.assertEqual(again['task']['task'], 'a')
        self.assertEqual(self.code(self.board.report, 'w2', 'a'), 'not_owner')
        self.assertEqual(self.code(self.board.report, 'w1', 'a', 'finished'), 'bad_arguments')
        self.assertEqual(self.code(self.board.report, 'w1', 'nope'), 'task_not_found')

    def test_lease_expires_and_progress_renews(self):
        self.open_with([{'key': 'a', 'title': 'x'}])
        self.board.join('w1')
        self.board.join('w2')
        self.board.next('w1')
        self.clock.t += store.LEASE_S - 10
        self.board.report('w1', 'a', 'progress', 'half way')
        self.clock.t += store.LEASE_S - 10
        self.assertIsNone(self.board.next('w2')['task'])         # renewed: still w1's
        self.clock.t += 20
        taken = self.board.next('w2')['task']
        self.assertEqual((taken['task'], taken['attempts']), ('a', 2))
        self.assertEqual(self.code(self.board.report, 'w1', 'a'), 'not_owner')

    def test_redo_failed_retry(self):
        self.open_with([{'key': 'a', 'title': 'x'}])
        self.board.join('w1')
        self.board.next('w1')
        self.board.report('w1', 'a', 'done', 'v1')
        self.assertEqual(self.board.review('lead', 'a', 'redo', 'missed L2')['status'], 'todo')
        self.assertEqual(self.code(self.board.review, 'lead', 'a', 'accept'), 'not_in_review')
        self.assertEqual(self.board.status(task='a')['notes'],
                         [{'by': 'lead', 'text': 'missed L2', 'task': 'a'}])
        self.board.next('w1')
        self.assertEqual(self.board.report('w1', 'a', 'failed', 'Revit error')['status'], 'failed')
        self.assertIsNone(self.board.next('w1')['task'])
        self.assertEqual(self.board.retry('lead', 'a')['status'], 'todo')
        self.assertEqual(self.board.next('w1')['task']['attempts'], 3)

    def test_assignee_and_lead_tasks(self):
        self.open_with([{'key': 'a', 'title': 'for w2', 'assignee': 'w2'},
                        {'key': 'b', 'title': 'lead decides', 'assignee': 'lead',
                         'scope': {'write': False}}], review=False)
        self.board.join('w1')
        self.board.join('w2')
        self.assertEqual(self.board.next('w1')['waiting'][0]['why'], 'assigned to w2')
        self.assertEqual(self.board.next('lead')['task']['task'], 'b')
        self.assertEqual(self.board.next('w2')['task']['task'], 'a')

    def test_notes_and_close(self):
        sid = self.open_with([{'key': 'a', 'title': 'x'}])
        self.board.join('w1')
        for i in range(store.MAX_NOTES_SHOWN + 3):
            self.board.note('w1', 'note {}'.format(i))
        notes = self.board.status()['notes']
        self.assertEqual(len(notes), store.MAX_NOTES_SHOWN)
        self.assertEqual(notes[-1], {'by': 'w1', 'text': 'note 12'})
        self.assertEqual(self.code(self.board.note, 'w1', ''), 'bad_arguments')
        self.assertTrue(self.board.close('lead')['closed'])
        self.assertEqual(self.code(self.board.next, 'w1', sid), 'session_closed')
        self.assertEqual(self.board.status(sid)['status'], 'closed')

    def test_long_text_is_cut(self):
        self.open_with([{'key': 'a', 'title': 'T' * 999, 'detail': 'D' * 9999}])
        task = self.board.status(task='a')['task']
        self.assertEqual((len(task['title']), len(task['detail'])), (200, 4000))

    def test_board_snapshot(self):
        self.assertEqual(store.read_board(self.root), {})
        sid = self.open_with([{'key': 'a', 'title': 'x', 'scope': {'document': 'B.rvt'}},
                              {'key': 'b', 'title': 'y', 'scope': {'write': False}}])
        self.board.join('w1', platform='anthropic')
        self.board.next('w1')
        board = store.read_board(self.root)
        self.assertEqual(board['format'], store.BOARD_FORMAT)
        session = board['sessions'][0]
        self.assertEqual((session['session'], session['document'], session['lead']),
                         (sid, 'PRJ.rvt', 'lead'))
        self.assertEqual(session['counts'], {'claimed': 1, 'todo': 1})
        self.assertEqual(session['members'][1],
                         {'member': 'w1', 'role': 'worker', 'platform': 'anthropic'})
        a, b = session['tasks']
        self.assertEqual((a['status'], a['owner'], a['document'], a['write'], a['started']),
                         ('claimed', 'w1', 'B.rvt', True, self.clock.t))
        self.assertEqual((b['status'], b['write']), ('todo', False))
        self.assertIs(store.read_board(self.root), board)           # cached by mtime
        stamp = os.stat(os.path.join(self.root, store.BOARD_FILE)).st_mtime_ns
        self.clock.t += 5
        self.assertTrue(self.board.next('w1')['already_claimed'])   # nothing changed...
        self.board.note('w1', 'still here')
        self.assertEqual(os.stat(os.path.join(self.root, store.BOARD_FILE)).st_mtime_ns,
                         stamp)                                     # ...so no rewrite
        self.assertEqual(store.task_of('w1', board),
                         {'session': sid, 'task': 'a', 'title': 'x', 'member': 'w1',
                          'document': 'B.rvt'})
        self.assertIsNone(store.task_of('lead', board))
        self.board.close('lead')
        self.assertEqual(store.read_board(self.root)['sessions'], [])
        self.assertEqual([n for n in os.listdir(self.root) if n.endswith('.tmp')], [])

    def test_task_of_one_client_with_two_chats(self):
        self.open_with([{'key': 'a', 'title': 'x', 'scope': {'write': False}},
                        {'key': 'b', 'title': 'y', 'scope': {'write': False}}])
        self.board.join('claude-desktop/arch')
        self.board.join('claude-desktop/mep')
        self.board.next('claude-desktop/arch')
        board = store.read_board(self.root)
        self.assertEqual(store.task_of('claude-desktop', board)['member'], 'claude-desktop/arch')
        self.assertIsNone(store.task_of('claude', board))            # not a prefix match
        self.board.next('claude-desktop/mep')
        # Two chats, two tasks: which chat is calling is unknown.
        self.assertIsNone(store.task_of('claude-desktop', store.read_board(self.root)))

    def test_call_element_ids(self):
        ids = store.call_element_ids
        self.assertEqual(ids({'element_ids': [12, '34', 'x', '@walls', True]}), ['12', '34'])
        # view_id / level_id / type_id are references, not what the call edits.
        self.assertEqual(ids({'element_id': 7, 'view_id': '9', 'name': '55', 'count': 3}), ['7'])
        self.assertEqual(ids({'elements': [{'id': 1}, {'id': 2, 'host_wall_id': 3}], 'id': 99}),
                         ['1', '2'])
        self.assertEqual(ids({'ids': [5], 'views': [{'id': 6}]}), ['5'])
        self.assertEqual(ids({'options': {'element_ids': [8]}}), ['8'])
        self.assertEqual(ids({'code': 'doc.Delete(ElementId(5))'}), [])
        self.assertEqual(ids(None), [])

    def refusal(self, agent, document='', element_ids=()):
        out = self.board.write_refusal(agent, document, element_ids)
        return out and out['error_code']

    def test_writes_stay_inside_the_task(self):
        self.assertIsNone(self.refusal('w1'))                        # no session at all
        self.open_with([
            {'key': 'ids', 'title': 'fixed set', 'assignee': 'w1',
             'scope': {'element_ids': [1, 2], 'levels': ['L1']}},
            {'key': 'other', 'title': 'theirs', 'assignee': 'w2',
             'scope': {'element_ids': [3, 4], 'levels': ['L9']}},
            {'key': 'lvl', 'title': 'a level', 'assignee': 'w5', 'scope': {'levels': ['L5']}},
            {'key': 'ro', 'title': 'read only', 'assignee': 'w3', 'scope': {'write': False}},
            {'key': 'b', 'title': 'other file', 'assignee': 'w4',
             'scope': {'document': 'B.rvt'}}], review=False)
        for name in ('w1', 'w2', 'w3', 'w4', 'w5', 'idle'):
            self.board.join(name)
        self.assertIsNone(self.refusal('stranger'))                  # not a member
        self.assertIsNone(self.refusal('lead'))                      # the person's client
        # A worker without a task cannot write in the session's files...
        self.assertEqual(self.refusal('idle'), 'scope_denied')
        self.assertEqual(self.refusal('idle', 'prj.RVT'), 'scope_denied')
        self.assertEqual(self.refusal('idle', 'B.rvt'), 'scope_denied')
        self.assertIsNone(self.refusal('idle', 'Other.rvt'))         # ...other files are free
        for name in ('w1', 'w2', 'w3', 'w4', 'w5'):
            self.assertIsNotNone(self.board.next(name)['task'], name)
        # element_ids: inside, outside, unknown, and another member's.
        self.assertIsNone(self.refusal('w1', '', ['1', 2]))
        self.assertIsNone(self.refusal('w1'))                        # send_code: no ids known
        out = self.board.write_refusal('w1', '', [1, 8, 9])
        self.assertEqual((out['error_code'], out['outside_ids'], out['task'], out['nothing_ran']),
                         ('scope_denied', ['8', '9'], 'ids', True))
        # Levels are not checked per call, but another member's elements are.
        self.assertIsNone(self.refusal('w5', '', [77]))
        out = self.board.write_refusal('w5', '', [3, 77])
        self.assertEqual((out['error_code'], out['task'], out['claimed_ids'], out['retryable']),
                         ('scope_claimed', 'other', ['3'], True))
        self.assertEqual(self.refusal('w1', 'B.rvt', [1]), 'scope_denied')   # wrong document
        self.assertEqual(self.refusal('w3'), 'scope_denied')         # read-only task
        self.assertIn('read-only', self.board.write_refusal('w3')['error'])
        self.assertIsNone(self.refusal('w4', 'B.rvt'))
        self.assertEqual(self.refusal('w4', 'PRJ.rvt'), 'scope_denied')
        # A reported task no longer licenses writes.
        self.board.report('w4', 'b')
        self.assertEqual(self.refusal('w4', 'B.rvt'), 'scope_denied')
        # A session nobody touched for ACTIVE_S stops limiting anyone.
        self.clock.t += store.ACTIVE_S + 60
        self.assertIsNone(self.refusal('idle'))
        self.board.close('lead')
        self.assertIsNone(self.refusal('w3'))

    def test_chats_of_one_client_share_the_scope_check(self):
        self.open_with([{'key': 'a', 'title': 'x', 'scope': {'element_ids': [1]}}])
        self.board.join('codex/a')
        self.board.next('codex/a')
        self.assertIsNone(self.refusal('codex', '', [1]))
        self.assertEqual(self.refusal('codex', '', [2]), 'scope_denied')
        self.assertIsNone(self.refusal('cod', '', [2]))

    def test_stats(self):
        sid = self.open_with([{'key': 'a', 'title': 'x', 'scope': {'write': False}},
                              {'key': 'b', 'title': 'y', 'scope': {'write': False}},
                              {'key': 'c', 'title': 'z', 'scope': {'write': False}}])
        self.board.join('w1')
        self.board.join('w2')
        self.board.next('w1')                   # a: t=1000
        self.board.next('w2')                   # b: t=1000
        self.clock.t += 60
        self.board.report('w1', 'a')            # 60 s
        self.clock.t += 40
        self.board.report('w2', 'b', 'failed')  # 100 s
        self.board.review('lead', 'a', 'redo')
        self.board.next('w1')                   # a again: t=1100
        self.clock.t += 20
        self.board.report('w1', 'a')            # 20 s
        stats = self.board.stats()
        self.assertEqual((stats['session'], stats['tasks']), (sid, 3))
        self.assertEqual(stats['counts'], {'review': 1, 'failed': 1, 'todo': 1})
        self.assertEqual((stats['attempts'], stats['extra_attempts']), (3, 1))
        self.assertEqual((stats['wall_s'], stats['work_s'], stats['parallelism']),
                         (120.0, 120.0, 1.0))
        self.assertEqual(stats['by_member'], {'w1': {'tasks': 1, 'work_s': 20.0},
                                              'w2': {'tasks': 1, 'work_s': 100.0}})
        empty = store.Harness(tempfile.mkdtemp(), now=self.clock)
        empty.open('lead', 'g')
        self.assertEqual((empty.stats()['wall_s'], empty.stats()['parallelism']), (0.0, None))

    def test_owner_review_needs_no_seat(self):
        self.open_with([{'key': 'a', 'title': 'x'}])
        self.board.join('w1')
        self.board.next('w1')
        self.board.report('w1', 'a')
        self.assertEqual(self.code(self.board.review, 'someone', 'a', 'accept'), 'not_a_member')
        out = self.board.review('ignored', 'a', 'redo', 'check level 2', as_owner=True)
        self.assertEqual(out['status'], 'todo')
        self.assertEqual(self.board.status(task='a')['notes'][-1]['by'], store.OWNER)

    def test_schema_1_database_is_upgraded(self):
        import sqlite3
        root = tempfile.mkdtemp(prefix='t3lab_harness_v1_')
        db = sqlite3.connect(os.path.join(root, 'harness.db'))
        db.executescript(store._SCHEMA.replace('    started REAL, finished REAL,\n', ''))
        db.execute("INSERT INTO meta VALUES ('schema', '1')")
        db.commit()
        db.close()
        board = store.Harness(root, now=self.clock)
        board.open('lead', 'g')
        board.plan('lead', [{'key': 'a', 'title': 'x'}])
        board.join('w1')
        self.assertEqual(board.next('w1')['task']['task'], 'a')
        self.assertEqual(store.read_board(root)['sessions'][0]['tasks'][0]['started'],
                         self.clock.t)

    def test_setup_file(self):
        self.assertEqual(store.load_config(self.root),
                         {'master': '', 'exclusive': False, 'write_mode': 'scoped'})
        saved = store.save_config(self.root, master='Claude-Desktop')
        self.assertEqual(saved, {'master': 'claude-desktop', 'exclusive': True,
                                 'write_mode': 'scoped'})                # master -> exclusive
        self.assertEqual(store.save_config(self.root, write_mode='serial')['write_mode'],
                         'serial')
        self.assertFalse(store.save_config(self.root, exclusive=False)['exclusive'])
        self.assertEqual(store.load_config(self.root),
                         {'master': 'claude-desktop', 'exclusive': False, 'write_mode': 'serial'})
        self.assertEqual(self.code(store.save_config, self.root, write_mode='fast'),
                         'bad_arguments')
        self.assertEqual(self.code(store.save_config, self.root, master='not an id!'),
                         'bad_arguments')
        self.assertEqual(store.save_config(self.root, master=''),
                         {'master': '', 'exclusive': False, 'write_mode': 'serial'})
        with open(os.path.join(self.root, store.CONFIG_FILE), 'w') as f:
            f.write('{broken')
        self.assertEqual(store.load_config(self.root)['master'], '')    # broken file = open

    def test_only_the_master_opens_sessions(self):
        store.save_config(self.root, master='claude-desktop', write_mode='serial')
        self.assertEqual(self.code(self.board.open, 'codex', 'g'), 'role_denied')
        self.assertEqual(self.code(self.board.open, 'codex/x', 'g', agent='codex'), 'role_denied')
        out = self.board.open('claude-desktop/lead', 'g', agent='claude-desktop')
        self.assertEqual((out['exclusive'], out['write_mode']), (True, 'serial'))   # the setup
        own = self.board.open('claude-desktop', 'g2', exclusive=False, write_mode='scoped')
        self.assertEqual((own['exclusive'], own['write_mode']), (False, 'scoped'))  # per session
        self.assertEqual(self.code(self.board.open, 'claude-desktop', 'g', write_mode='x'),
                         'bad_arguments')
        board = store.read_board(self.root)
        self.assertEqual(board['master'], 'claude-desktop')
        self.assertEqual([(x['exclusive'], x['write_mode']) for x in board['sessions']],
                         [(True, 'serial'), (False, 'scoped')])

    def test_exclusive_session_binds_agents_that_never_joined(self):
        sid = self.board.open('claude-desktop', 'Audit', 'PRJ.rvt', exclusive=True)['session']
        self.board.plan('claude-desktop', [{'key': 'a', 'title': 'x', 'assignee': 'codex'}])
        out = self.board.write_refusal('antigravity', 'PRJ.rvt')
        self.assertEqual((out['error_code'], out['lead'], out['session'], out['nothing_ran']),
                         ('harness_managed', 'claude-desktop', sid, True))
        self.assertIn('action=join', out['hint'])
        self.assertEqual(self.refusal('antigravity'), 'harness_managed')     # document unknown
        self.assertIsNone(self.refusal('antigravity', 'Other.rvt'))          # another file
        self.assertIsNone(self.refusal('claude-desktop', 'PRJ.rvt'))         # the master
        self.board.join('antigravity')
        self.assertEqual(self.refusal('antigravity', 'PRJ.rvt'), 'scope_denied')   # no task yet
        self.board.join('codex')
        self.board.next('codex')
        self.assertIsNone(self.refusal('codex', 'PRJ.rvt'))
        self.board.close('claude-desktop')
        self.assertIsNone(self.refusal('antigravity', 'PRJ.rvt'))
        # Not exclusive: agents outside the session are left alone.
        self.board.open('claude-desktop', 'Open one', 'PRJ.rvt', exclusive=False)
        self.assertIsNone(self.refusal('antigravity', 'PRJ.rvt'))

    def test_serial_write_mode_is_one_writer_per_document(self):
        self.board.open('lead', 'g', 'PRJ.rvt', review=False, write_mode='serial')
        self.board.plan('lead', [
            {'key': 'l1', 'title': 'walls L1', 'scope': {'levels': ['L1']}},
            {'key': 'l2', 'title': 'walls L2', 'scope': {'levels': ['L2']}},
            {'key': 'b', 'title': 'other file', 'scope': {'document': 'B.rvt'}},
            {'key': 'r', 'title': 'read', 'scope': {'write': False}}])
        for name in ('w1', 'w2', 'w3', 'w4'):
            self.board.join(name)
        self.assertEqual(self.board.next('w1')['task']['task'], 'l1')
        self.assertEqual(self.board.next('w2')['task']['task'], 'b')         # another document
        self.assertEqual(self.board.next('w3')['task']['task'], 'r')         # readers never wait
        wait = self.board.next('w4')
        self.assertEqual(wait['waiting'], [{'task': 'l2',
                                            'why': 'one writer at a time: after l1 (w1)'}])
        self.board.report('w1', 'l1')
        self.assertEqual(self.board.next('w4')['task']['task'], 'l2')
        self.assertEqual(self.board.status()['write_mode'], 'serial')
        self.assertTrue(store.scopes_conflict({'levels': ['l1']}, {'levels': ['l2']}, serial=True))
        self.assertFalse(store.scopes_conflict({'write': False}, {}, serial=True))

    def test_schema_2_database_is_upgraded(self):
        import sqlite3
        root = tempfile.mkdtemp(prefix='t3lab_harness_v2_')
        db = sqlite3.connect(os.path.join(root, 'harness.db'))
        db.executescript(store._SCHEMA.replace(', exclusive INTEGER, write_mode TEXT', ''))
        db.execute("INSERT INTO sessions VALUES ('old00001', 'old goal', 'A.rvt', 'lead', 'open', "
                   '1, 5.0, NULL)')
        db.execute("INSERT INTO members VALUES ('old00001', 'lead', 'lead', '', 5.0, 5.0)")
        db.commit()
        db.close()
        board = store.Harness(root, now=self.clock)
        status = board.status('old00001')
        self.assertEqual((status['exclusive'], status['write_mode']), (False, 'scoped'))
        self.assertIsNone(board.write_refusal('stranger', 'A.rvt'))
        self.assertTrue(board.open('lead', 'new', exclusive=True)['exclusive'])

    def test_two_processes_share_the_board(self):
        self.open_with([{'key': 'a', 'title': 'x'}], review=False)
        code = ('import sys; sys.path.insert(0, {lib!r}); from core import harness_store as s; '
                'b = s.Harness({root!r}); b.join("other"); print(b.next("other")["task"]["task"])'
                .format(lib=LIB, root=self.root))
        out = subprocess.run([sys.executable, '-c', code], capture_output=True, text=True)
        self.assertEqual(out.stdout.strip(), 'a', out.stderr)
        self.assertEqual(self.board.status()['tasks'][0]['owner'], 'other')


# ─── Agents from other platforms ───────────────────────────────────────────────

class Ev(types.SimpleNamespace):
    pass


def text_event(text):
    return Ev(type='agent.message', content=[Ev(type='text', text=text)])


def idle(reason='end_turn'):
    return Ev(type='session.status_idle', stop_reason=Ev(type=reason))


class FakeAnthropic(object):
    """client.beta.sessions.{create, events.send, events.stream} only."""

    def __init__(self, events):
        self.sent, self.created = [], None
        outer = self

        class Stream(object):
            def __enter__(self):
                return iter(events)

            def __exit__(self, *exc):
                return False

        def create(**kwargs):
            outer.created = kwargs
            return Ev(id='sesn_1')

        def send(session_id, events):
            outer.sent.extend(events)

        self.beta = Ev(sessions=Ev(create=create, events=Ev(
            send=send, stream=lambda session_id: Stream())))


class AgentTests(BoardCase):
    ANT = {'platform': 'anthropic', 'agent_id': 'agent_1', 'environment_id': 'env_1'}

    def test_registry_validation(self):
        code = lambda *a: self._agent_code(*a)      # noqa: E731
        self.assertEqual(code('X', self.ANT), 'bad_arguments')
        self.assertEqual(code('bot', {'platform': 'nope'}), 'bad_arguments')
        self.assertEqual(code('bot', {'platform': 'anthropic'}), 'bad_arguments')
        self.assertEqual(code('bot', dict(self.ANT, api_key='sk-x')), 'bad_arguments')
        self.assertEqual(code('bot', dict(self.ANT, role='lead')), 'bad_arguments')
        self.assertEqual(code('bot', {'platform': 'command', 'command': 'claude -p x'}),
                         'bad_arguments')
        self.assertEqual(code('bot', {'platform': 'command', 'command': ['claude', '-p']}),
                         'bad_arguments')
        self.assertEqual(code('bot', {'platform': 'http', 'url': 'http://example.com/x'}),
                         'bad_arguments')
        self.assertEqual(code('claude-code', {'platform': 'command',
                                              'command': ['claude', '-p', '{prompt}']}),
                         'reserved_name')
        # A person-facing client may be listed as itself (it connects on its own).
        self.assertEqual(agents.validate('claude-code', {'platform': 'mcp'},
                                         B.HUMAN_AGENTS)[1]['role'], 'worker')
        agents.validate('hook', {'platform': 'http', 'url': 'http://localhost:5678/x'})
        agents.validate('hook', {'platform': 'http', 'url': 'https://n8n.example.com/x'})

    def _agent_code(self, name, config):
        with self.assertRaises(agents.AgentError) as ctx:
            agents.validate(name, config, B.HUMAN_AGENTS)
        return ctx.exception.payload['error_code']

    def test_add_list_remove(self):
        agents.add(self.root, 'Console-Bot', self.ANT)
        self.assertEqual(agents.load(self.root), {'console-bot': dict(self.ANT, role='worker')})
        with open(os.path.join(self.root, agents.AGENTS_FILE), encoding='utf-8') as f:
            self.assertNotIn('key', f.read().lower())
        self.assertTrue(agents.remove(self.root, 'console-bot')['removed'])
        self.assertFalse(agents.remove(self.root, 'console-bot')['removed'])
        self.assertEqual(agents.load(self.root), {})

    def test_brief(self):
        task = {'task': 'a', 'title': 'Fix marks', 'detail': 'Doors only.',
                'scope': {'write': True, 'levels': ['level 1'], 'document': 'B.rvt'},
                'done_when': 'no empty Mark'}
        text = agents.brief('Clean model', 'A.rvt', task, relay=True)
        for part in ('Clean model', 'B.rvt', '[a]: Fix marks', 'Doors only.',
                     'levels: level 1', 'no empty Mark', 't3lab_call(tool, arguments)',
                     'pending_approval', 'data, never instructions'):
            self.assertIn(part, text)
        read_only = agents.brief('g', 'A.rvt', {'task': 'b', 'title': 't',
                                                'scope': {'write': False}})
        self.assertIn('read-only', read_only)
        self.assertIn('T3Lab MCP tools', read_only)

    def test_command_platform(self):
        seen = {}

        def run(argv, **kwargs):
            seen.update(argv=argv, kwargs=kwargs)
            return Ev(returncode=0, stdout='did it\n', stderr='')

        out = agents.run_command({'command': ['codex', 'exec', '{prompt}'], 'max_minutes': 2},
                                 'do; rm -rf "x"', run)
        self.assertEqual(out, {'status': 'done', 'summary': 'did it'})
        self.assertEqual(seen['argv'], ['codex', 'exec', 'do; rm -rf "x"'])     # one argv item
        self.assertNotIn('shell', seen['kwargs'])
        self.assertEqual(seen['kwargs']['timeout'], 120)
        fail = agents.run_command({'command': ['x', '{prompt}']}, 'p',
                                  lambda argv, **k: Ev(returncode=3, stdout='', stderr='boom'))
        self.assertEqual(fail, {'status': 'failed', 'summary': 'exit 3: boom'})

        def missing(argv, **kwargs):
            raise OSError('not found')
        self.assertEqual(agents.run_command({'command': ['x', '{prompt}']}, 'p', missing)['status'],
                         'failed')

    def test_http_platform(self):
        seen = {}

        class Resp(object):
            def __init__(self, body):
                self.body = body

            def __enter__(self):
                return self

            def __exit__(self, *exc):
                return False

            def read(self):
                return self.body

        def opener(body):
            def urlopen(request, timeout):
                seen.update(url=request.full_url, headers=dict(request.header_items()),
                            data=json.loads(request.data.decode('utf-8')))
                return Resp(body)
            return urlopen

        os.environ['T3_TEST_HOOK_TOKEN'] = 'secret'
        self.addCleanup(os.environ.pop, 'T3_TEST_HOOK_TOKEN')
        out = agents.run_http({'url': 'https://h.example/x', 'token_env': 'T3_TEST_HOOK_TOKEN'},
                              {'task': {'task': 'a'}}, opener(b'{"summary": "ok 3"}'))
        self.assertEqual(out, {'status': 'done', 'summary': 'ok 3'})
        self.assertEqual(seen['headers']['Authorization'], 'Bearer secret')
        self.assertEqual(seen['data'], {'task': {'task': 'a'}})
        out = agents.run_http({'url': 'https://h.example/x'}, {},
                              opener(b'{"status": "failed", "summary": "no"}'))
        self.assertEqual(out['status'], 'failed')
        self.assertNotIn('Authorization', seen['headers'])
        self.assertEqual(agents.run_http({'url': 'https://h.example/x'}, {},
                                         opener(b'plain text'))['summary'], 'plain text')

    def test_anthropic_platform_relays_the_custom_tool(self):
        events = [
            text_event('Looking at the model.'),
            Ev(type='agent.custom_tool_use', id='sevt_1', name='t3lab_call',
               input={'tool': 'get_levels', 'arguments': {'revit_document': 'A.rvt'}}),
            idle('requires_action'),
            Ev(type='agent.custom_tool_use', id='sevt_2', name='bash', input={}),
            Ev(type='agent.tool_use', id='sevt_3', evaluated_permission='ask'),
            text_event('3 levels checked.'),
            idle(),
            text_event('never read'),
        ]
        client = FakeAnthropic(events)
        calls = []

        def call_tool(tool, arguments):
            calls.append((tool, arguments))
            return {'levels': 3}

        out = agents.run_anthropic(dict(self.ANT, version=4, max_usd=2.5, vault_ids=['vlt_1']),
                                   'the brief', call_tool, client)
        self.assertEqual(out['status'], 'done')
        self.assertEqual(out['summary'], '3 levels checked.')
        self.assertEqual(out['tool_calls'], 2)
        self.assertTrue(out['trace'].endswith('/workspaces/default/sessions/sesn_1'))
        self.assertEqual(calls, [('get_levels', {'revit_document': 'A.rvt'})])
        self.assertEqual(client.created['agent'], {'type': 'agent', 'id': 'agent_1', 'version': 4})
        self.assertEqual(client.created['budget'], {'type': 'limit', 'max_list_cost': {
            'amount': '250', 'currency': 'USD'}})
        self.assertEqual(client.created['vault_ids'], ['vlt_1'])
        kinds = [e['type'] for e in client.sent]
        self.assertEqual(kinds, ['user.message', 'user.custom_tool_result',
                                 'user.custom_tool_result', 'user.tool_confirmation'])
        self.assertEqual(json.loads(client.sent[1]['content'][0]['text']), {'levels': 3})
        self.assertEqual(json.loads(client.sent[2]['content'][0]['text'])['error_code'],
                         'tool_denied')
        self.assertEqual(client.sent[3], {'type': 'user.tool_confirmation',
                                          'tool_use_id': 'sevt_3', 'result': 'deny'})

    def test_anthropic_platform_failures(self):
        run = lambda events, **kw: agents.run_anthropic(      # noqa: E731
            self.ANT, 'p', lambda t, a: {}, FakeAnthropic(events), **kw)
        self.assertEqual(run([idle()])['status'], 'failed')              # no message at all
        self.assertIn('Budget', run([text_event('x'), idle('budget_reached')])['summary'])
        self.assertEqual(run([Ev(type='session.status_terminated')])['status'], 'failed')
        self.assertIn('stream closed', run([text_event('x')])['summary'])
        ticks = iter([0, 10 ** 6])
        client = FakeAnthropic([text_event('x'), text_event('y')])
        out = agents.run_anthropic(self.ANT, 'p', lambda t, a: {}, client,
                                   clock=lambda: next(ticks))
        self.assertIn('Timed out', out['summary'])
        self.assertEqual(client.sent[-1], {'type': 'user.interrupt'})

    def test_runner_round(self):
        self.open_with([{'key': 'a', 'title': 'cli task', 'scope': {'write': False}},
                        {'key': 'b', 'title': 'cloud task', 'scope': {'write': False}},
                        {'key': 'c', 'title': 'hook task', 'scope': {'write': False}}])
        agents.add(self.root, 'a-cli', {'platform': 'command', 'command': ['x', '{prompt}']})
        agents.add(self.root, 'b-cloud', self.ANT)
        agents.add(self.root, 'c-hook', {'platform': 'http', 'url': 'https://h.example/x'})
        agents.add(self.root, 'desktop', {'platform': 'mcp'})        # never launched
        asked = []

        def call_tool_for(name):
            asked.append(name)
            return lambda tool, arguments: {}

        def boom(config, payload):
            raise RuntimeError('socket closed')

        runners = {
            'command': lambda config, prompt: {'status': 'done', 'summary': 'cli ok'},
            'anthropic': lambda config, prompt, call_tool: {
                'status': 'done', 'summary': 'cloud ok', 'platform_session': 'sesn_9'},
            'http': boom}
        ran = agents.run_once(self.board, self.root, call_tool_for, runners=runners)
        self.assertEqual([(r['agent'], r['task'], r['status']) for r in ran],
                         [('a-cli', 'a', 'done'), ('b-cloud', 'b', 'done'),
                          ('c-hook', 'c', 'failed')])
        self.assertEqual(asked, ['b-cloud'])
        status = self.board.status()
        self.assertEqual(status['counts'], {'review': 2, 'failed': 1})
        self.assertEqual(self.board.status(task='b')['task']['result'],
                         {'platform_session': 'sesn_9'})
        self.assertIn('RuntimeError: socket closed', self.board.status(task='c')['task']['summary'])
        self.assertEqual(sorted(m['member'] for m in status['members']),
                         ['a-cli', 'b-cloud', 'c-hook', 'lead'])
        again = agents.run_once(self.board, self.root, call_tool_for, only='a-cli',
                                runners=runners)
        self.assertEqual(again, [{'agent': 'a-cli', 'task': None, 'all_done': False}])

    def test_runner_works_the_agents_at_the_same_time(self):
        import threading
        self.open_with([{'key': k, 'title': k, 'scope': {'write': False}} for k in 'abc'],
                       review=False)
        for name in ('one', 'two', 'three'):
            agents.add(self.root, name, {'platform': 'command', 'command': ['x', '{prompt}']})
        together = threading.Barrier(3, timeout=10)

        def run(config, prompt):
            together.wait()             # passes only if all three are running at once
            return {'status': 'done', 'summary': 'ok'}

        ran = agents.run_once(self.board, self.root, lambda name: None,
                              runners={'command': run})
        self.assertEqual([(r['agent'], r['status']) for r in ran],
                         [('one', 'done'), ('three', 'done'), ('two', 'done')])
        self.assertEqual(self.board.status()['counts'], {'done': 3})

        # parallel=1: one after another (a barrier of 2 would never fill).
        self.board.plan('lead', [{'key': k, 'title': k, 'scope': {'write': False}}
                                 for k in 'de'])
        running, peak = [], []

        def serial(config, prompt):
            running.append(1)
            peak.append(len(running))
            running.pop()
            return {'status': 'done', 'summary': 'ok'}

        agents.run_once(self.board, self.root, lambda name: None,
                        runners={'command': serial}, parallel=1)
        self.assertEqual(peak, [1, 1])


# ─── Gateway tool ──────────────────────────────────────────────────────────────


def _client(agent, **extra):
    """A bridge `current` whose caller is `agent` (as --agent would name it)."""
    return dict({'port': 1, 'client': agent, 'agent_explicit': True, 'validated': True},
                **extra)


class BridgeTests(unittest.TestCase):
    def setUp(self):
        os.environ['T3LAB_HARNESS_DIR'] = tempfile.mkdtemp(prefix='t3lab_harness_bridge_')
        self.human = _client('claude-desktop')
        self.bot = _client('openclaw')
        self.cli = _client('cli', transport='cli')
        # No Revit here: what the active document is, per test.
        self.docs = []
        real = B._documents_on
        B._documents_on = lambda port, token: list(self.docs)
        self.addCleanup(setattr, B, '_documents_on', real)
        B._save_cached_tools([
            {'name': 'set_parameter', 'inputSchema': {'type': 'object', 'properties': {
                'element_ids': {}, 'dry_run': {}}}},
            {'name': 'get_levels', 'annotations': {'readOnlyHint': True},
             'inputSchema': {'type': 'object', 'properties': {}}}])
        self.addCleanup(B._save_cached_tools, [])

    def call(self, arguments, current=None):
        return B._tool_payload(B._handle(B._call_request(1, 't3lab_harness', arguments),
                                         current or self.human, 't'))

    def gate(self, current, tool, arguments):
        return B._harness_write_error(tool, current, arguments, 't')

    def worker_with_task(self, scope=None, agent='openclaw'):
        """A session on PRJ.rvt where `agent` holds task `a`."""
        self.call({'action': 'open', 'goal': 'g', 'revit_document': 'PRJ.rvt'})
        self.call({'action': 'plan', 'tasks': [
            {'key': 'a', 'title': 'Fix marks', 'scope': scope or {}}]})
        worker = _client(agent)
        self.call({'action': 'join'}, worker)
        self.call({'action': 'next'}, worker)
        return worker

    def test_tool_is_listed_and_never_gated(self):
        real = B._forward
        B._forward = lambda *a, **k: (_ for _ in ()).throw(ConnectionRefusedError())
        self.addCleanup(setattr, B, '_forward', real)
        listed = B._handle({'jsonrpc': '2.0', 'id': 1, 'method': 'tools/list'}, self.human, 't')
        names = [t['name'] for t in listed['result']['tools']]
        self.assertEqual(names.count('t3lab_harness'), 1)
        self.assertTrue(B._is_read_tool('t3lab_harness'))
        self.assertTrue(B._is_read_tool('get_levels'))
        self.assertFalse(B._is_read_tool('set_parameter'))
        self.assertFalse(B._is_read_tool('a_tool_nobody_listed'))     # unknown: a write

    def test_full_flow_through_the_bridge(self):
        sid = self.call({'action': 'open', 'goal': 'Check sheets',
                         'revit_document': 'PRJ.rvt'})['session']
        planned = self.call({'action': 'plan', 'options': {'tasks': [
            {'key': 'a', 'title': 'List sheets', 'scope': {'write': False}}]}})
        self.assertEqual(planned['planned'], ['a'])
        self.assertEqual(self.call({'action': 'join'}, self.bot)['document'], 'PRJ.rvt')
        self.assertEqual(self.call({'action': 'next'}, self.bot)['task']['task'], 'a')
        self.call({'action': 'report', 'task': 'a', 'summary': '40 sheets'}, self.bot)
        self.assertEqual(self.call({'action': 'review', 'task': 'a', 'verdict': 'accept'},
                                   self.bot)['error_code'], 'role_denied')
        self.assertEqual(self.call({'action': 'review', 'task': 'a',
                                    'verdict': 'accept'})['status'], 'done')
        status = self.call({'action': 'status', 'session': sid}, self.bot)
        self.assertEqual(status['counts'], {'done': 1})
        self.assertEqual([m['member'] for m in status['members']], ['claude-desktop', 'openclaw'])

    def test_two_chats_of_one_client(self):
        self.call({'action': 'open', 'goal': 'g', 'as': 'Lead!'})
        self.call({'action': 'plan', 'as': 'lead', 'tasks': [{'key': 'a', 'title': 'x'}]})
        self.assertEqual(self.call({'action': 'plan', 'tasks': [{'title': 'y'}]})['error_code'],
                         'not_a_member')
        self.call({'action': 'join', 'as': 'arch'})
        self.assertEqual(self.call({'action': 'next', 'as': 'arch'})['task']['task'], 'a')
        members = [m['member'] for m in self.call({'action': 'status'})['members']]
        self.assertEqual(members, ['claude-desktop/lead', 'claude-desktop/arch'])

    def test_unattended_agents_cannot_open_or_launch(self):
        for action in ('open', 'agents', 'agent_add', 'agent_remove', 'run', 'flow_save',
                       'flow_delete'):
            out = self.call({'action': action, 'goal': 'g'}, self.bot)
            self.assertEqual((out['error_code'], out['nothing_ran']), ('role_denied', True), action)

    def test_agent_registry_and_run_from_the_cli_only(self):
        added = self.call({'action': 'agent_add', 'name': 'console-bot', 'config': {
            'platform': 'anthropic', 'agent_id': 'agent_1', 'environment_id': 'env_1'}})
        self.assertTrue(added['saved'])
        listed = self.call({'action': 'agents'})
        self.assertEqual(list(listed['agents']), ['console-bot'])
        self.assertEqual(listed['relay_tool']['name'], 't3lab_call')
        for reserved in ('codex', 'antigravity', 'mcp-client', 'cli'):
            self.assertEqual(self.call({'action': 'agent_add', 'name': reserved, 'config': {
                'platform': 'command', 'command': ['x', '{prompt}']}})['error_code'],
                'reserved_name', reserved)
        self.assertEqual(self.call({'action': 'run'})['error_code'], 'use_cli')
        self.call({'action': 'agent_remove', 'name': 'console-bot'})
        self.call({'action': 'open', 'goal': 'g'})
        self.assertEqual(self.call({'action': 'run'}, self.cli), {'ran': []})

    def test_launched_agents_only_read(self):
        """Lite has no approval queue: an agent nobody watches never writes."""
        forwarded = []
        real = B._forward
        B._forward = lambda request, current, token, timeout=None: (
            forwarded.append((request['params']['name'], current['client'])) or
            B._tool_response(request['id'], {'ok': True}))
        self.addCleanup(setattr, B, '_forward', real)
        call_tool = B._harness_call_tool_for(self.human, 't')('console-bot')
        out = call_tool('set_parameter', {'element_ids': [1]})
        self.assertEqual((out['error_code'], out['nothing_ran']), ('tool_denied', True))
        self.assertEqual(call_tool('get_levels', {}), {'ok': True})
        self.assertEqual(call_tool('t3lab_harness', {'action': 'status'}), {'sessions': []})
        self.assertEqual(forwarded, [('get_levels', 'console-bot')])   # its own name, not ours
        self.assertNotIn('console-bot', B.HUMAN_AGENTS)

    def test_calls_carry_the_task_for_the_cursors(self):
        request = B._call_request(1, 'set_parameter', {})
        self.assertIs(B._stamp_harness(request, self.human), request)  # harness unused: as before
        worker = self.worker_with_task()
        meta = B._stamp_harness(request, worker)['params']['_meta']
        sid = self.call({'action': 'status'})['session']
        self.assertEqual(meta['t3lab/task'], sid + '/a')
        self.assertEqual(meta['t3lab/task_title'], 'Fix marks')
        self.assertEqual((meta['t3lab/lane'], meta['t3lab/queue_wait_s']), ('task', 45))
        self.assertNotIn('t3lab/member', meta)              # member == agent id
        lead = B._stamp_harness(request, self.human)['params']['_meta']
        self.assertEqual(lead['t3lab/lane'], 'lead')
        self.assertNotIn('t3lab/task', lead)
        self.assertEqual(B._stamp_harness(request, _client('cursor'))['params']['_meta'][
            't3lab/lane'], 'person')
        self.assertEqual(B._stamp_harness(request, _client('nightly-bot'))['params']['_meta'][
            't3lab/lane'], 'auto')

    def test_a_chat_of_a_client_is_stamped_as_member(self):
        self.call({'action': 'open', 'goal': 'g', 'as': 'lead'})
        self.call({'action': 'plan', 'as': 'lead',
                   'tasks': [{'key': 'a', 'title': 'x', 'scope': {'write': False}}]})
        self.call({'action': 'join', 'as': 'arch'})
        self.call({'action': 'next', 'as': 'arch'})
        meta = B._stamp_harness(B._call_request(1, 'get_levels', {}), self.human)
        self.assertEqual(meta['params']['_meta']['t3lab/member'], 'claude-desktop/arch')

    def test_bridge_refuses_writes_outside_the_task(self):
        idle = _client('nightly-bot')
        self.assertIsNone(self.gate(idle, 'set_parameter', {'element_ids': [1]}))   # no session
        worker = self.worker_with_task({'element_ids': [1, 2]})
        self.call({'action': 'join'}, idle)
        self.docs = [{'title': 'PRJ.rvt', 'is_active': True}]
        self.assertIsNone(self.gate(worker, 'set_parameter', {'element_ids': [1]}))
        out = self.gate(worker, 'set_parameter', {'element_ids': [1, 5]})
        self.assertEqual((out['error_code'], out['outside_ids'], out['nothing_ran']),
                         ('scope_denied', ['5'], True))
        self.docs = [{'title': 'Other.rvt', 'is_active': True}]     # another file is active
        self.assertEqual(self.gate(worker, 'set_parameter', {'element_ids': [1]})['error_code'],
                         'scope_denied')
        self.docs = [{'title': 'PRJ.rvt', 'is_active': True}]
        # A preview changes nothing; reads and the bridge's own tools are never limited.
        self.assertIsNone(self.gate(worker, 'set_parameter', {'element_ids': [5], 'dry_run': True}))
        self.assertIsNone(self.gate(worker, 'get_levels', {'element_ids': [5]}))
        self.assertIsNone(self.gate(worker, 't3lab_harness', {'action': 'status'}))
        # A worker with no task; the lead (a person's client) stays free.
        out = self.gate(idle, 'set_parameter', {})
        self.assertEqual(out['error_code'], 'scope_denied')
        self.assertNotIn('revit_document', out['hint'])             # Lite pins no document
        self.assertIn('switch_active_document', out['hint'])
        self.assertIsNone(self.gate(self.human, 'set_parameter', {'element_ids': [5]}))
        # The refusal reaches the agent as a normal tool result, before Revit.
        real = B._forward
        B._forward = lambda *a, **k: self.fail('a refused call must not reach Revit')
        self.addCleanup(setattr, B, '_forward', real)
        answer = B._tool_payload(B._handle(
            B._call_request(1, 'set_parameter', {'element_ids': [9]}), worker, 't'))
        self.assertEqual(answer['error_code'], 'scope_denied')

    def test_a_broken_board_never_blocks_a_call(self):
        worker = self.worker_with_task({'write': False})
        self.assertEqual(self.gate(worker, 'set_parameter', {})['error_code'], 'scope_denied')
        real = B._hb.Harness.write_refusal
        B._hb.Harness.write_refusal = lambda *a, **k: 1 / 0
        self.addCleanup(setattr, B._hb.Harness, 'write_refusal', real)
        self.assertIsNone(self.gate(worker, 'set_parameter', {}))

    def test_next_waits_for_a_task(self):
        self.call({'action': 'open', 'goal': 'g'})
        self.call({'action': 'join'}, self.bot)
        board = B._hb.Harness()
        naps = []

        def sleep(seconds):
            naps.append(seconds)
            if len(naps) == 2:              # the lead plans while the worker waits
                board.plan('claude-desktop', [{'key': 'a', 'title': 'late'}])

        out = B._harness_next(board, 'openclaw', None, 30, sleep)
        self.assertEqual((out['task']['task'], naps), ('a', [2, 2]))
        B._harness_next(board, 'openclaw', None, 0, sleep)
        self.assertEqual(len(naps), 2)                      # wait_s=0: one look, no sleep
        self.assertEqual(B.HARNESS_WAIT_MAX_S, 50)

    def test_stats_and_owner_review(self):
        worker = self.worker_with_task({'write': False})
        self.call({'action': 'report', 'task': 'a', 'summary': 'ok'}, worker)
        stats = self.call({'action': 'stats'}, worker)
        self.assertEqual((stats['tasks'], stats['counts']), (1, {'review': 1}))
        # owner=true only counts from the CLI on this machine.
        denied = self.call({'action': 'review', 'task': 'a', 'verdict': 'accept', 'owner': True},
                           _client('cli'))
        self.assertEqual(denied['error_code'], 'not_a_member')
        done = self.call({'action': 'review', 'task': 'a', 'verdict': 'accept', 'owner': True},
                         self.cli)
        self.assertEqual(done['status'], 'done')

    def setup_master(self, **changes):
        return self.call(dict({'action': 'setup'}, **changes), self.cli)

    def test_setup_is_changed_at_the_machine_only(self):
        shown = self.call({'action': 'setup'}, self.bot)
        self.assertEqual((shown['setup']['master'], shown['you']), ('', 'openclaw'))
        denied = self.call({'action': 'setup', 'master': 'openclaw'}, self.human)
        self.assertEqual((denied['error_code'], denied['nothing_ran']), ('use_cli', True))
        self.assertEqual(B._hb.load_config()['master'], '')
        out = self.setup_master(master='claude-desktop', write_mode='serial')
        self.assertEqual(out['setup'], {'master': 'claude-desktop', 'exclusive': True,
                                        'write_mode': 'serial'})
        self.assertEqual(self.setup_master(write_mode='nope')['error_code'], 'bad_arguments')

    def test_claude_gpt_antigravity_with_one_master(self):
        """The owner's case: three apps on one model, Claude leads."""
        self.setup_master(master='claude-desktop', write_mode='serial')
        gpt, gemini = _client('codex'), _client('antigravity')
        self.assertEqual(self.call({'action': 'open', 'goal': 'g'}, gpt)['error_code'],
                         'role_denied')
        self.call({'action': 'open', 'goal': 'Clean the door schedule',
                   'revit_document': 'PRJ.rvt'})
        self.call({'action': 'plan', 'tasks': [
            {'key': 'marks', 'title': 'Fix door marks', 'assignee': 'codex'},
            {'key': 'fire', 'title': 'Fill fire rating', 'assignee': 'antigravity'},
            {'key': 'check', 'title': 'List doors without a room', 'assignee': 'antigravity',
             'scope': {'write': False}}]})

        def write(current):
            return self.gate(current, 'set_parameter', {})

        # GPT has not joined: the active document is the session's, so it is told to join.
        self.docs = [{'title': 'PRJ.rvt', 'is_active': True}]
        out = write(gpt)
        self.assertEqual((out['error_code'], out['lead']), ('harness_managed', 'claude-desktop'))
        self.assertIn('switch_active_document', out['hint'])
        self.docs = [{'title': 'Link.rvt', 'is_active': True}]
        self.assertIsNone(write(gpt))                                   # not the session's file
        self.docs = [{'title': 'PRJ.rvt', 'is_active': True}]
        self.call({'action': 'join'}, gpt)
        self.call({'action': 'join'}, gemini)
        self.assertEqual(self.call({'action': 'next'}, gpt)['task']['task'], 'marks')
        # One writer at a time: Antigravity gets its read task, its write task waits.
        self.assertEqual(self.call({'action': 'next'}, gemini)['task']['task'], 'check')
        self.assertIsNone(write(gpt))
        self.assertEqual(write(gemini)['error_code'], 'scope_denied')   # holds a read-only task
        self.assertIsNone(write(self.human))                            # the master is free
        self.call({'action': 'report', 'task': 'check', 'summary': '3 doors'}, gemini)
        waiting = self.call({'action': 'next'}, gemini)
        self.assertEqual(waiting['waiting'][0]['why'], 'one writer at a time: after marks (codex)')
        self.call({'action': 'report', 'task': 'marks', 'summary': 'done'}, gpt)
        self.assertEqual(write(gpt)['error_code'], 'scope_denied')      # reported: no more writes
        self.call({'action': 'review', 'task': 'marks', 'verdict': 'accept'})
        self.assertEqual(self.call({'action': 'next'}, gemini)['task']['task'], 'fire')
        self.assertIsNone(write(gemini))

    def test_master_may_be_an_unattended_agent(self):
        self.assertEqual(self.call({'action': 'open', 'goal': 'g'}, self.bot)['error_code'],
                         'role_denied')
        self.setup_master(master='openclaw')
        opened = self.call({'action': 'open', 'goal': 'g'}, self.bot)
        self.assertEqual((opened['role'], opened['exclusive']), ('lead', True))
        # Leading the board does not let it launch agents.
        self.assertEqual(self.call({'action': 'run'}, self.bot)['error_code'], 'role_denied')

    def test_client_identity(self):
        """--agent wins; else a known clientInfo name, else the name as sent."""
        def agent(name, **current):
            current = dict({'port': 1}, **current)
            B._handle({'jsonrpc': '2.0', 'id': 1, 'method': 'initialize',
                       'params': {'clientInfo': {'name': name}}}, current, 't')
            return B._agent_id(current)

        for name, expected in (('codex-mcp-client', 'codex'), ('Antigravity', 'antigravity'),
                               ('claude-ai', 'claude-desktop'), ('claude-code', 'claude-code'),
                               ('cursor-vscode', 'cursor'), ('some-new-client', 'some-new-client')):
            self.assertEqual(agent(name), expected, name)
        self.assertEqual(agent('codex-mcp-client', client='antigravity', agent_explicit=True),
                         'antigravity')
        self.assertEqual(B._agent_id({'port': 1}), 'mcp-client')      # nothing known
        self.assertTrue(set(B.CLIENT_AGENTS.values()) <= set(B.HUMAN_AGENTS))
        self.assertTrue(set(a for _n, a in B.HOST_APPS) <= set(B.HUMAN_AGENTS))
        # The cursor label stays what it was: only the board id is mapped.
        current = {'port': 1}
        B._handle({'jsonrpc': '2.0', 'id': 1, 'method': 'initialize',
                   'params': {'clientInfo': {'name': 'claude-ai'}}}, current, 't')
        self.assertEqual(current['client'], 'claude-ai')

    def test_errors_and_help(self):
        self.assertEqual(self.call({'action': 'dance'})['error_code'], 'bad_arguments')
        self.assertIn('plan', self.call({'action': 'help'}, self.bot)['lead'])
        self.assertEqual(self.call({'action': 'next'})['error_code'], 'session_required')

    def test_companions_are_deployed(self):
        sys.path.insert(0, LIB)
        import bridge_deploy
        for name in ('harness_store.py', 'harness_agents.py', 'harness_studio.py',
                     'harness_studio.html'):
            self.assertIn((name, name), bridge_deploy.COMPANIONS)

    def test_the_revit_server_cannot_import_the_board(self):
        """The server runs IronPython 2.7 inside Revit: sqlite3 and the runner
        stay in the bridge process."""
        import warnings
        sys.path.insert(0, os.path.join(REPO, 'lite_guard'))
        import audit_ipy_compat
        with warnings.catch_warnings():
            warnings.simplefilter('ignore', ResourceWarning)    # the audit's own file reads
            reachable = set(audit_ipy_compat.reachable())
        self.assertIn('core/conductor.py', reachable)
        for name in ('harness_store', 'harness_agents', 'harness_studio', 'bridge'):
            self.assertNotIn('core/{}.py'.format(name), reachable)

    def test_cli(self):
        env = dict(os.environ, PYTHONPATH='')
        bridge = os.path.join(LIB, 'core', 'bridge.py')

        def cli(*args):
            done = subprocess.run([sys.executable, bridge, 'harness'] + list(args),
                                  capture_output=True, text=True, env=env)
            return done.returncode, json.loads(done.stdout.strip().splitlines()[-1]
                                               if args[:2] == ('run', '--loop')
                                               else done.stdout)

        code, out = cli('setup', '--master=claude-desktop', '--write-mode=serial',
                        '--exclusive=on')
        self.assertEqual((code, out['setup']),
                         (0, {'master': 'claude-desktop', 'exclusive': True,
                              'write_mode': 'serial'}))
        code, out = cli('setup', '--master=')
        self.assertEqual((code, out['setup']['master']), (0, ''))
        path = os.path.join(tempfile.mkdtemp(), 'agent.json')
        with open(path, 'w', encoding='utf-8-sig') as f:
            json.dump({'name': 'night-bot', 'config': {'platform': 'mcp'}}, f)
        code, out = cli('agent_add', '--json-file=' + path)
        self.assertTrue(out['saved'], out)
        code, opened = cli('open', '--json', '{"goal": "from the cli"}')
        self.assertEqual((code, opened['role']), (0, 'lead'))
        code, status = cli('status', '--session=' + opened['session'])
        self.assertEqual((code, status['lead']), (0, 'cli'))
        code, out = cli('next', '--session=nope')
        self.assertEqual((code, out['error_code']), (1, 'session_not_found'))
        code, out = cli('open', '--json', 'not json')
        self.assertEqual((code, out['error_code']), (2, 'bad_arguments'))
        code, out = cli('run', '--loop')
        self.assertEqual((code, out), (1, {'ran': []}))        # nobody launched is registered


if __name__ == '__main__':
    unittest.main()
