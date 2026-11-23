# -*- coding: utf-8 -*-
"""
CPython 3 test: agent cursor overlay helpers (lib/core/agent_cursors.py).

Run:  python3 lite_guard/test_agent_cursors.py      Exit 0 = all pass.

Covers the pure parts: element / view ids taken from tool arguments/results,
the action of a call, the orthographic view -> screen projection, which open
views are visible (tiled vs tabbed), agent colours, the setting and options,
and that server.py calls before() ahead of the tool and notify() only after
the caller got its result.
The WPF overlay itself needs Revit: NEEDS VERIFICATION there.
"""
import ast
import os
import re
import shutil
import sys
import tempfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)

_TMP = tempfile.mkdtemp(prefix='t3cursors_')
os.environ['APPDATA'] = _TMP

import core.agent_cursors as AC   # noqa: E402

BUTTON = 'T3Lab.tab/Support.panel/Settings.pulldown/AgentCursors.pushbutton/script.py'
SERVER = 'lib/core/server.py'
FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok   ' + name)
    else:
        print('  FAIL ' + name + (' -- ' + str(detail) if detail else ''))
        FAILURES.append(name)


def test_extract_ids():
    print('[extract_ids]')
    ids = AC.extract_ids({'element_id': 12, 'value': 'x'}, {'created_ids': [13, '14', 12]})
    check('args + result, unique, ordered', ids == [12, 13, 14], ids)
    ids = AC.extract_ids({'view_id': 5, 'type_id': 6, 'level_id': 7, 'host_id': 8})
    check('view/type/level ids skipped, host kept', ids == [8], ids)
    ids = AC.extract_ids({'id': 99, 'elements': [{'id': 1}, {'id': 2, 'name': 'W'}]})
    check('bare id only inside an elements list', ids == [1, 2], ids)
    ids = AC.extract_ids({'element_ids': list(range(1, 50))})
    check('capped at MAX_TARGETS', len(ids) == AC.MAX_TARGETS, len(ids))
    ids = AC.extract_ids({'element_id': True, 'ElementId': '-1', 'elementIds': ['a', 3.0]})
    check('bools, negatives and text rejected', ids == [3], ids)
    check('None payloads are fine', AC.extract_ids(None, None) == [])


def test_view_ids():
    print('[extract_view_ids]')
    ids = AC.extract_view_ids({'view_id': 5, 'element_id': 9})
    check('view id from arguments', ids == [5], ids)
    ids = AC.extract_view_ids({'view_ids': [3, 4], 'sheet_id': '7'})
    check('view list + sheet', ids == [3, 4, 7], ids)
    ids = AC.extract_view_ids({}, {'view_id': 11}, 'create')
    check('created view from the result', ids == [11], ids)
    ids = AC.extract_view_ids({}, {'views': [{'id': 1}, {'id': 2}]}, 'read')
    check('a read result (list_views) is not followed', ids == [], ids)
    ids = AC.extract_view_ids({'view_ids': list(range(1, 20))})
    check('capped at MAX_FOLLOW_VIEWS', len(ids) == AC.MAX_FOLLOW_VIEWS, len(ids))
    check('view keys are not element ids', AC.extract_ids({'new_view_id': 5}) == [])


def test_actions():
    print('[classify_action]')
    cases = (
        ('revit_get_element_info', None, 'read'),
        ('list_levels', None, 'read'),
        ('create_level', None, 'create'),
        ('place_wall', None, 'create'),
        ('tag_all_walls', None, 'create'),
        ('famigen_create_family', None, 'create'),
        ('set_parameter', None, 'modify'),
        ('move_elements', None, 'modify'),
        ('delete_element', None, 'delete'),
        ('purge_unused', None, 'delete'),
        ('set_active_view', None, 'view'),
        ('select_elements', None, 'view'),
        ('revit_override_color', None, 'view'),
        ('export_sheets_pdf', None, 'export'),
        ('send_code_to_revit', {'code': 'x'}, 'code'),
        ('operate_element', {'operation': 'hide'}, 'view'),
        ('operate_element', {'operation': 'pin'}, 'modify'),
        ('manage_view', {'operation': 'get_properties'}, 'read'),
        ('manage_view', {'operation': 'set_scale'}, 'modify'),
        ('manage_document', {'operation': 'save'}, 'export'),
        ('edit_elements', {'operation': 'mirror'}, 'modify'),
        ('mystery_tool', None, 'modify'),
    )
    for tool, args, want in cases:
        got = AC.classify_action(tool, args)
        check('{} {} -> {}'.format(tool, (args or {}).get('operation', ''), want), got == want, got)
    check('created ids in the result -> create',
          AC.classify_action('mystery_tool', None, {'created_ids': [1]}) == 'create')
    keys = [a[0] for a in AC.ACTIONS]
    check('every action has a style', all(len(AC.action_style(k)) == 3 for k in keys))
    check('unknown action styled as modify', AC.action_style('x') == AC.action_style('modify'))
    check('counts text in ACTIONS order',
          AC.counts_text({'create': 3, 'read': 1}) == 'Read 1 · Create 3',
          AC.counts_text({'create': 3, 'read': 1}))


def test_captions_and_tour():
    print('[describe_call / badge_text / tour]')
    cases = (
        (('list_open_documents', None, 0, False), 'Listing open documents'),
        (('revit_get_element_info', {'element_id': 1}, 1, False), 'Reading element info · 1 element'),
        (('move_elements', {}, 3, False), 'Moving elements · 3 elements'),
        (('operate_element', {'operation': 'hide'}, 5, False), 'Hiding elements · 5 elements'),
        (('send_code_to_revit', {}, 0, False), 'Running code in Revit'),
        (('send_code_to_revit', {'description': ' Number the  stair risers ', 'code': 'x'}, 0, False),
         'Number the stair risers'),
        (('send_code_to_revit', {'code': '# check door marks\nresult = 1'}, 0, False),
         'Check door marks'),
        (('delete_element', {}, 2, True), 'Failed: Deleting element · 2 elements'),
        (('famigen_create_family', None, 0, False), 'Creating family'),
        (('frobnicate_walls', None, 0, False), 'Frobnicating walls'),
    )
    for args, want in cases:
        got = AC.describe_call(*args)
        check('caption ' + args[0], got == want, got)
    long = AC.describe_call('get_' + '_'.join(['parameter'] * 12))
    check('long caption cut with an ellipsis', len(long) <= 46 and long.endswith('…'), long)
    check('empty tool -> Working', AC.describe_call('') == 'Working')
    code_cases = (
        ('# -*- coding: utf-8 -*-\n# Renumber the treads\nx = 1', 'Renumber the treads'),
        ('"""List rooms without area."""\nresult = 1', 'List rooms without area'),
        ('from Autodesk.Revit.DB import *\nt = Transaction(doc, "Renumber risers")\nt.Start()',
         'Renumber risers'),
        ('c = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StairsRailing)\n'
         'for r in c: r.LookupParameter("Mark").Set("R")', 'Editing stairs railing'),
        ('els = FilteredElementCollector(doc).OfClass(Level)\nresult = [e.Name for e in els]',
         'Reading levels'),
        ('result = 1 + 1', ''),
    )
    for code, want in code_cases:
        got = AC.summarize_code(code)
        check('code summary ' + (want or 'none'), got == want, got)
    check('badge with category', AC.badge_text('Walls', '1200 × 300') == 'Walls · 1200 × 300')
    check('badge without category', AC.badge_text('', '10 × 10') == '10 × 10')
    sched = AC.tour_schedule([(0, 0), (10, 5), (20, 20)], 300, 200)
    check('tour: arrive / leave in order',
          [s_[2:] for s_ in sched] == [(300, 500), (800, 1000), (1300, 1500)], sched)
    check('one stop glides MOVE_MS', AC.hop_ms(1) == AC.MOVE_MS)
    check('many stops glide faster, never under 260 ms',
          AC.hop_ms(4) < AC.hop_ms(1) and AC.hop_ms(50) == 260)
    rect = (100.0, 50.0, 900.0, 650.0)
    pts = AC.review_points(rect, 7, 3)
    check('review path: n points', len(pts) == 3)
    check('review path inside the view',
          all(rect[0] < x < rect[2] and rect[1] < y < rect[3] for x, y in pts), pts)
    check('review path stable per seed', pts == AC.review_points(rect, 7, 3))
    check('another seed, another path', pts != AC.review_points(rect, 8, 3))


def test_activity_helpers():
    print('[goal / changes / conflicts / off-screen / summary]')
    check('goal one line', AC.clean_goal('  Renumber\n  doors  ') == 'Renumber doors')
    long = AC.clean_goal('x' * 200)
    check('goal capped', len(long) == AC.NARRATE_MAX and long.endswith('…'))
    check('change text', AC.change_text('Comments', '', 'Checked') == 'Comments: "" → Checked',
          AC.change_text('Comments', '', 'Checked'))
    check('change text shortens long values',
          len(AC.change_text('Mark', 'a' * 50, 'b')) < 60)
    touched = {}
    check('first edit: no clash', AC.track_touches(touched, 'claude', [1, 2], 'modify') == [])
    check('same agent again: no clash', AC.track_touches(touched, 'claude', [1], 'modify') == [])
    check('reads never clash', AC.track_touches(touched, 'codex', [1], 'read') == [])
    clash = AC.track_touches(touched, 'codex', [1, 3], 'delete')
    check('another agent edits 1: clash with claude', clash == [(1, ['claude'])], clash)
    rect = (0.0, 0.0, 400.0, 300.0)
    x, y, ang = AC.edge_point(rect, (1000.0, 150.0))
    check('edge point on the right edge', abs(x - 382.0) < 1e-6 and abs(y - 150.0) < 1e-6
          and abs(ang) < 1e-6, (x, y, ang))
    marks = AC.offscreen_markers(rect, [(1000, 150), (900, 160), (-50, 150), (200, 100)])
    check('one marker per side, inside points ignored',
          sorted(m[3] for m in marks) == [1, 2], marks)
    check('markers sit inside the view', all(0 <= m[0] <= 400 and 0 <= m[1] <= 300 for m in marks))
    lines = AC.summary_lines({'claude-desktop': {'read': 2, 'create': 1}}, 3)
    check('summary lines', lines == ['Claude: Read 2 · Create 1', '3 elements changed'], lines)


def test_perspective():
    print('[perspective projection]')
    cam = {'eye': (0, 0, 0), 'forward': (0, 1, 0), 'up': (0, 0, 1), 'right': (1, 0, 0),
           'c0': (-10, 10, -5), 'c1': (10, 10, 5)}
    rect = (0.0, 0.0, 200.0, 100.0)
    x, y = AC.project_perspective((0, 10, 0), cam, rect)
    check('centre of view -> centre of screen', (x, y) == (100.0, 50.0), (x, y))
    x, y = AC.project_perspective((10, 20, 0), cam, rect)
    check('twice as far, half the offset', abs(x - 150.0) < 1e-9, x)
    check('behind the eye -> None', AC.project_perspective((0, -5, 0), cam, rect) is None)
    corners = [(a, b, c) for a in (-1, 1) for b in (9, 11) for c in (-1, 1)]
    box = AC.perspective_box(corners, cam, rect)
    check('box around the centre', box is not None and box[0] < 100 < box[2], box)
    check('box behind the eye -> None',
          AC.perspective_box([(0, -1, 0)] * 8, cam, rect) is None)


def test_narrate_and_glow():
    print('[narrate / glow helpers]')
    check('narrate schema', AC.NARRATE_SCHEMA['name'] == AC.NARRATE_TOOL == 't3lab_narrate'
          and AC.NARRATE_SCHEMA['inputSchema']['required'] == []
          and set(AC.NARRATE_SCHEMA['inputSchema']['properties']) == {'message', 'tasks', 'step', 'done'})
    res = AC.narrate('claude-ai', '  Renumber doors on Level 2 ')
    check('narrate stores the goal', res['success'] and res['goal'] == 'Renumber doors on Level 2'
          and AC._session['goals'].get('claude-desktop') == res['goal'], res)
    res = AC.narrate('codex', 'Renumber stairs', ['Read the stairs', '  Number   risers ',
                                                  'x' * 80, ''], None)
    check('plan stored, step 1', res['tasks'][:2] == ['Read the stairs', 'Number risers']
          and len(res['tasks']) == 3 and len(res['tasks'][2]) == AC.TASK_MAX
          and res['current'] == 'Step 1/3 · Read the stairs', res)
    res = AC.narrate('codex', step=2)
    check('step advances, goal kept', res['goal'] == 'Renumber stairs'
          and res['current'] == 'Step 2/3 · Number risers', res)
    check('plan rows', [r[0] for r in AC.plan_rows(AC._plan_of('codex'))] ==
          ['done', 'current', 'todo'])
    res = AC.narrate('codex', done=True)
    check('done: every row done, no current task', res['current'] == 'all done'
          and AC.current_task(AC._plan_of('codex')) == ''
          and [r[0] for r in AC.plan_rows(AC._plan_of('codex'))] == ['done'] * 3, res)
    check('plans are per agent', AC.current_task(AC._plan_of('claude-ai')) == '')
    check('nothing to show is an error', AC.narrate('cursor')['error_code'] == 'bad_arguments')
    check('too many tasks capped', len(AC.clean_tasks(['t'] * 20)) == AC.PLAN_MAX)
    AC.waiting('claude-ai', 'delete_element', {'element_ids': [1]}, True)   # no overlay: quiet
    check('waiting never raises without an overlay', True)
    import core.agent_cursors_glow as G
    check('glow: 8 corners', len(G.corners((0, 0, 0), (1, 2, 3))) == 8)
    idx = G.line_indices(2)
    check('glow: 12 edges per box, second box offset by 8',
          len(idx) == 24 and idx[12] == (8, 9) and max(max(p) for p in idx) == 15, idx[:3])
    check('select_touched quiet without Revit', AC.select_touched() is False)


def test_views():
    print('[visible_views / union_rect / moved]')
    a, b = (0, 0, 500, 400), (500, 0, 1000, 400)
    check('tabbed views: only the active one', AC.visible_views([a, a, a], 1) == [1])
    check('tiled views: all of them', AC.visible_views([a, b], 0) == [0, 1])
    check('no active view -> none', AC.visible_views([a, b], None) == [])
    tiny = (0, 0, 10, 10)
    check('tiny windows skipped', AC.visible_views([a, (600, 0, 610, 10)], 0) == [0])
    check('union', AC.union_rect([a, b]) == (0, 0, 1000, 400))
    check('moved', AC.moved((0, 0, 10, 10), (5, 0, 15, 10)))
    check('not moved within tolerance', not AC.moved((0, 0, 10, 10), (1, 1, 11, 11)))
    del tiny


def test_projection():
    print('[project / screen_box]')
    rect = (100.0, 50.0, 500.0, 350.0)          # 400 x 300 px
    origin, right, up = (0, 0, 0), (1, 0, 0), (0, 1, 0)
    x, y = AC.project((0, 0, 0), origin, right, up, 40.0, 30.0, rect)
    check('model min -> bottom-left', (x, y) == (100.0, 350.0), (x, y))
    x, y = AC.project((40, 30, 9), origin, right, up, 40.0, 30.0, rect)
    check('model max -> top-right, z ignored in plan', (x, y) == (500.0, 50.0), (x, y))
    corners = [(a, b, c) for a in (10, 20) for b in (10, 20) for c in (0, 3)]
    box = AC.screen_box(corners, origin, right, up, 40.0, 30.0, rect)
    check('box', box == (200.0, 150.0, 300.0, 250.0), box)
    far = [(a, b, 0) for a in (100, 110) for b in (100, 110)]
    check('outside the view -> None',
          AC.screen_box(far, origin, right, up, 40.0, 30.0, rect) is None)
    edge = [(a, b, 0) for a in (-5, 5) for b in (5, 10)]
    box = AC.screen_box(edge, origin, right, up, 40.0, 30.0, rect)
    check('clipped to the view', box is not None and box[0] == 100.0, box)
    # Section: right = +X, up = +Z
    x, y = AC.project((20, 7, 15), origin, (1, 0, 0), (0, 0, 1), 40.0, 30.0, rect)
    check('section uses Z for up', (x, y) == (300.0, 200.0), (x, y))
    check('size label in mm', AC.size_label(1.0, 0.5) == '305 × 152', AC.size_label(1.0, 0.5))


def test_agent_style():
    print('[agent_style]')
    check('claude desktop', AC.agent_style('claude-desktop')[0] == 'Claude')
    check('codex shows as GPT', AC.agent_style('codex')[0] == 'GPT')
    check('missing agent -> "AI app", never a guessed app', AC.agent_style(None) == AC.AGENTS['mcp-client'])
    check('MCP client name claude-ai -> Claude', AC.agent_style('claude-ai')[0] == 'Claude')
    check('MCP client name codex-mcp-client -> GPT', AC.agent_style('codex-mcp-client')[0] == 'GPT')
    a, b = AC.agent_style('my-bot'), AC.agent_style('my-bot')
    check('unknown agent: stable colour', a == b and a[1] in AC._SPARE, a)
    for key, (_tag, color) in AC.AGENTS.items():
        if not (len(color) == 7 and color.startswith('#')):
            check('colour of ' + key, False, color)


def test_agent_names():
    print('[agent names: raw MCP client ids read as names]')
    st = AC.agent_style
    check('an unseen client id is never guessed as Claude (2026-11-17: Antigravity)',
          st('local-agent-mode-t3lab')[0] != 'Claude'
          and AC.agent_key('Local-Agent-Mode-T3Lab_Lite') != 'claude-desktop')
    check('Connect names the apps', st('antigravity')[0] == 'Antigravity'
          and st('codex')[0] == 'GPT' and st('claude-desktop')[0] == 'Claude')
    check('no name at all reads "AI app"', st('mcp-client')[0] == 'AI app')
    check('suffixed known clients', st('claude-code-cli')[0] == 'Claude Code'
          and st('gemini-cli')[0] == 'Antigravity' and st('codex_vscode')[0] == 'GPT'
          and st('cursor-agent')[0] == 'Cursor')
    check('a prefix needs a separator', AC.agent_key('cursorless') == 'cursorless'
          and AC.agent_key('claudette') == 'claudette')
    pa = AC.pretty_agent
    check('noise and version dropped, title case',
          pa('acme-bot-mcp-client') == 'Acme Bot' and pa('my_bot-mcp') == 'My Bot'
          and pa('revit.helper/1.2.0') == 'Revit Helper' and pa('console-qa') == 'Console QA',
          [pa('acme-bot-mcp-client'), pa('my_bot-mcp'), pa('revit.helper/1.2.0'), pa('console-qa')])
    check('capped at NAME_MAX', len(pa('a-very-long-agent-name-that-goes-on')) <= AC.NAME_MAX
          and pa('a-very-long-agent-name-that-goes-on').endswith(u'\u2026'))
    check('nothing left -> Agent', pa('') == 'Agent'
          and pa(None) == 'Agent')
    check('unknown agent shows prettified', st('acme-bot-mcp-client')[0] == 'Acme Bot'
          and st('acme-bot-mcp-client')[1] in AC._SPARE)


def test_visible_work():
    print('[visible work: one cursor, tags, scan, created label, change capture]')
    rect = (0, 0, 1000, 800)
    # tags
    check('tag right-below by default', AC.place_tags([(100, 100, 200, 20)], rect) == [(12.0, 12.0)])
    dx, dy = AC.place_tags([(900, 100, 200, 20)], rect)[0]
    check('tag flips left at the right edge', dx == -204.0 and 900 + dx >= 0 and dy == 12.0, (dx, dy))
    dx, dy = AC.place_tags([(100, 790, 200, 20)], rect)[0]
    check('tag flips above at the bottom edge', dy == -24.0 and dx == 12.0, (dx, dy))
    dx, dy = AC.place_tags([(950, 790, 200, 20)], rect)[0]
    check('corner: left and above', (dx, dy) == (-204.0, -24.0), (dx, dy))
    items = [(300, 300, 220, 20), (310, 304, 220, 20), (305, 302, 220, 20)]
    out = AC.place_tags(items, rect)
    boxes = [(x + dx, y + dy, x + dx + w, y + dy + h) for (x, y, w, h), (dx, dy) in zip(items, out)]
    check('tags of cursors side by side never overlap',
          not any(AC._overlaps(a, b) for i, a in enumerate(boxes) for b in boxes[i + 1:]), boxes)
    check('the first (active) cursor keeps the default spot', out[0] == (12.0, 12.0))
    check('all inside the view', all(b[0] >= 0 and b[1] >= 0 and b[2] <= 1000 and b[3] <= 800
                                     for b in boxes))
    check('no view rect: default spot', AC.place_tags([(5000, 5000, 200, 20)]) == [(12.0, 12.0)])
    # scan near the last place
    pts = AC.scan_points((600, 400), rect, 7, 3)
    check('scan stays near where the cursor worked',
          len(pts) == 3 and all(abs(x - 600) <= AC.SCAN_RADIUS and abs(y - 400) <= AC.SCAN_RADIUS
                                for x, y in pts), pts)
    check('scan: same seed, same path; other seed, other path',
          pts == AC.scan_points((600, 400), rect, 7, 3) and pts != AC.scan_points((600, 400), rect, 8, 3))
    check('scan: one point = the anchor (Less motion)', AC.scan_points((600, 400), rect, 7, 1) == [(600, 400)])
    check('scan: kept inside the view',
          all(24 <= x <= 976 and 24 <= y <= 776 for x, y in AC.scan_points((2000, -50), rect, 3, 3)))
    check('scan: no anchor -> middle of the view', AC.scan_points(None, rect, 1, 1) == [(500.0, 400.0)])
    # labels
    wl = AC.work_label
    check('created label', wl(['Walls', 'Walls', 'Walls'], 'create') == u'Walls \u00d73 created', wl(['Walls'] * 3, 'create'))
    check('two categories + the rest',
          wl(['Doors', 'Walls', 'Walls', 'Floors', 'Roofs', ''], 'modify')
          == u'Walls \u00d72 \u00b7 Doors \u00d71 \u00b7 +3 more changed',
          wl(['Doors', 'Walls', 'Walls', 'Floors', 'Roofs', ''], 'modify'))
    check('no elements -> no label', wl([], 'create') == '' and wl(None, 'create') == '')
    # targets
    mt = AC.merge_targets
    check('named ids first, then added', mt([5, 6], [6, 7, 8], [99]) == [5, 6, 7, 8])
    check('modified only when nothing else', mt([], [], [3, 4]) == [3, 4] and mt([1], [], [3]) == [1])
    check('targets capped', len(mt([], list(range(100)))) == AC.MAX_TARGETS)
    ids = AC.extract_ids({}, {'created_elements': [{'id': 11}, {'id': 12}], 'results': [{'id': 13}]})
    check('ids inside created_elements / results lists', ids == [11, 12, 13], ids)
    # change capture
    class Args(object):
        def GetAddedElementIds(self):
            return [101, 102, 101]
        def GetModifiedElementIds(self):
            return list(range(1, 500))
    AC._session['call'] = None
    AC._on_changed(None, Args())
    check('no call running: nothing captured', AC._session['call'] is None)
    AC._session['call'] = call = {'added': [], 'modified': []}
    try:
        AC._on_changed(None, Args())
        AC._on_changed(None, Args())
    finally:
        AC._session['call'] = None
    check('added ids captured once each', call['added'] == [101, 102], call['added'])
    check('modified ids capped', len(call['modified']) == AC.MAX_CAPTURE)
    AC._on_changed(None, object())      # no call: quiet
    acsrc = open(os.path.join(LIB, 'core', 'agent_cursors.py'), encoding='utf-8').read()
    check('DocumentChanged hooked once, old delegate detached on reload',
          'application.DocumentChanged -= old' in acsrc
          and 'application.DocumentChanged += delegate' in acsrc and '_WATCH_KEY' in acsrc)
    before = acsrc.split('def before(app, task):', 1)[1].split('\ndef ', 1)[0]
    check('before(): capture starts, named elements go to working()',
          "_session['call'] = {'added': [], 'modified': []}" in before
          and 'targets=[b[\'box\'] for b in boxes[:MAX_STOPS]]' in before)
    notify = acsrc.split('def notify(app, task):', 1)[1].split('\ndef ', 1)[0]
    check('notify(): capture ends, created label + fresh ids go to show()',
          "call, _session['call'] = _session.get('call'), None" in notify
          and 'created=label, fresh=fresh' in notify)
    check('notify(): a call that started always ends on the bar (show, else done)',
          'finally:\n        if started and not handed:\n            _done(' in notify
          and 'finally:\n                if not ended:\n                    _done(' in notify
          and notify.index('started = True') < notify.index("call, _session['call']")
          and notify.index("'show', agent") < notify.index('ended = True\n            except')
          and notify.index('handed = True') < notify.index('_when_settled(app, finish)'))
    check('before(): no view to draw in, the bar still says Working once it is up',
          "elif _state.get('overlay') is not None:" in before
          and "describe_call(tool, args, len(ids)), tool, None," in before)

    class Host(object):
        alive = True

        def __init__(self):
            self.calls = []

        def call(self, name, *args, **kwargs):
            if name != 'set_theme':     # notify() syncs the theme on every call
                self.calls.append((name, args, kwargs))

    class App(object):
        ActiveUIDocument = None         # no document: nothing to snapshot

    class Call(object):
        tool_name, arguments, result, exception, audit = 'list_open_documents', {}, {}, None, None
        agent = 'codex'
    was_on, host = AC.is_enabled(), Host()
    AC.set_enabled(True)
    AC._state['overlay'] = host
    try:
        AC.before(App(), Call())
        AC.notify(App(), Call())
        names = [c[0] for c in host.calls]
        check('no document: the bar is told the call started and ended',
              names == ['working', 'done'] and host.calls[0][1][0] == 'codex'
              and host.calls[0][1][4] is None and host.calls[1][1][:4]
              == ('codex', 'read', 'list_open_documents', False), host.calls)
        del host.calls[:]
        bad = Call()
        bad.result = {'error': 'boom'}
        AC.notify(App(), bad)
        check('a failed call ends as failed', host.calls and host.calls[0][0] == 'done'
              and host.calls[0][1][3] is True, host.calls)
        del host.calls[:]
        real_options, AC.options = AC.options, lambda: 1 / 0
        try:
            AC.notify(App(), Call())
        finally:
            AC.options = real_options
        check('notify() failing half way still ends the call',
              [c[0] for c in host.calls] == ['done'], host.calls)
        del host.calls[:]
        real_snapshot, AC._snapshot = AC._snapshot, lambda *a, **k: 1 / 0
        try:
            AC.notify(App(), Call())
        finally:
            AC._snapshot = real_snapshot
        check('the deferred half failing still ends the call, once',
              [c[0] for c in host.calls] == ['done'], host.calls)
        del host.calls[:]
        held = Call()
        held.result = {'error_code': 'paused', 'error': 'Paused'}
        AC.notify(App(), held)
        check('a call held by Pause is reported as held, not as done',
              [c[0] for c in host.calls] == ['held'], host.calls)
    finally:
        AC._state['overlay'] = None
        AC.set_enabled(was_on)
    check('plan / section views are boxed (IsPerspective is View3D only)',
          "getattr(view, 'IsPerspective', False)" in acsrc and 'view.IsPerspective' not in acsrc)
    osrc = open(os.path.join(LIB, 'core', 'agent_cursors_overlay.py'), encoding='utf-8').read()
    check('one cursor per agent (no agent|action key)',
          "'{}|{}'.format(agent, action)" not in osrc and 'self._cursors[agent] = cur' in osrc
          and 'def _accent(self, cur, action):' in osrc)
    check('tags laid out after every move',
          osrc.count('self._layout_tags(agent)') >= 3 and 'place_tags(items, self._bounds())' in osrc)
    # The drawing part of show() (On model) lives in _draw_step.
    show = osrc.split('    def _draw_step(self, agent,', 1)[1].split('\n    def ', 1)[0]
    check('show(): scan near the cursor, label + build trail for fresh elements',
          "scan_points((cur['x'], cur['y'])" in show and 'review_points' not in osrc
          and 'self._draw_label(' in show and 'self._build_trail(' in show
          and 'self._clear_focus(agent)' in show)
    trail = show.split('self._draw_label(', 1)[1].split('self._tour(', 1)[0]
    check('Less motion: no build trail', 'if not reduce and' in trail)
    working = osrc.split('    def working(self, agent,', 1)[1].split('\n    def ', 1)[0]
    check('working(): the bar says Working first, then the guarded drawing',
          working.index('self._working[agent] = ') < working.index('self._draw_working(')
          < working.index('self._show_hud(rect)') and 'elif rect is not None:' in working)
    drawing = osrc.split('    def _draw_working(self, agent,', 1)[1].split('\n    def ', 1)[0]
    check('_draw_working(): frames what the call names and goes there',
          'self._draw_focus(agent, b, acolor)' in drawing and 'self._tour(cur,' in drawing)
    shown = osrc.split('    def show(self, agent,', 1)[1].split('\n    def ', 1)[0]
    check('show(): the drawing on the model is guarded, the bar always gets the step',
          shown.index('self._working.pop(agent, None)') < shown.index('self._draw_step(')
          < shown.index('self._log_step(') < shown.index('self._show_hud('))
    check('caption capped on the tag', 'short_text(text, TAG_MAX)' in osrc and AC.TAG_MAX <= 40)


def test_setting():
    print('[setting]')
    check('off by default', AC.is_enabled() is False)
    AC.set_enabled(True)
    check('on after set_enabled(True)', AC.is_enabled() is True)
    AC.set_enabled(False)
    check('off again', AC.is_enabled() is False)
    opts = AC.options()
    osrc = open(os.path.join(LIB, 'core', 'agent_cursors_overlay.py'), encoding='utf-8').read()
    check('icon font falls back (no tofu on the bar)',
          "ICON_FONTS = 'Segoe Fluent Icons, Segoe MDL2 Assets'" in osrc
          and "FontFamily(ICON_FONTS if glyph else 'Segoe UI')" in osrc
          and "FontFamily('Segoe MDL2 Assets'" not in osrc)
    import re as _re
    tsv = os.path.join(REPO, 'dev', 'icons', 'mdl2_codepoints.tsv')    # t3lab_dev only
    if os.path.exists(tsv):
        known = set(l.split('\t', 1)[0].strip().upper() for l in open(tsv, encoding='utf-8')
                    if l[:1] != '#')
        used = set(c.upper() for c in _re.findall(r'\\u(E[0-9A-Fa-f]{3})', osrc))
        check('every bar glyph is in Segoe MDL2', used <= known, sorted(used - known))
    check('options default: bar only (owner, 2026-11-18), the rest on for On model',
          opts == {'model': False, 'follow': True, 'zoom': True, 'select': True, 'tile': True,
                   'glow': True, 'trail': True, 'ripple': True, 'labels': True, 'keep': True},
          opts)
    check('Effects: Glow and the three the overlay already drew, now one switch each',
          dict(AC.OPTION_GROUPS)['Effects'] == ('glow', 'trail', 'ripple', 'labels'))
    check('Less motion is no option any more: Animation Off replaced it',
          'reduce' not in opts and 'reduce' not in AC.OPTION_LABELS
          and 'agent_cursors_reduce_motion' not in [k for _n, k, _d in AC.OPTIONS])
    eff = AC.effective(opts)
    check('bar only: nothing touches the model, Keep bar stays',
          eff['keep'] and not any(eff[n] for n in ('follow', 'zoom', 'select', 'tile', 'glow',
                                                   'trail', 'ripple', 'labels')))
    check('On model: the saved options act',
          AC.effective(dict(opts, model=True)) == dict(opts, model=True))
    src = open(os.path.join(LIB, 'core', 'agent_cursors_overlay.py'), encoding='utf-8').read()
    tick = src.split('    def _on_tick(self, sender, e):', 1)[1].split('    def ', 1)[0]
    check('keep bar: never auto-hidden after the summary',
          'now - self._summary >= SUMMARY_S and not self._keep()' in tick)
    summary = src.split('    def _enter_summary(self, keep=False):', 1)[1].split('\n    def ', 1)[0]
    leave = src.split('    def _leave(self):', 1)[1].split('\n    def ', 1)[0]
    check('idle: cursors and boxes always leave (Keep bar keeps only the bar)',
          'if not keep:' not in summary and 'self._leave()' in summary
          and 'self.canvas.Children.Clear()' in leave and 'self.win.Hide()' in leave)
    check('idle: Glow taken off the model through the main thread',
          ".get('idle')" in summary and "'idle': agents_idle" in open(
              os.path.join(LIB, 'core', 'agent_cursors.py'), encoding='utf-8').read())
    front = src.split('    def _on_front_tick(self, sender, e):', 1)[1].split('    def ', 1)[0]
    check('a resting bar never brings the cursor window back', 'self._resting' not in front)
    AC._session['select_event'] = None
    check('ui action without an event is a no-op', AC.agents_idle() is False)

    class _Event(object):
        raised = 0

        def Raise(self):
            _Event.raised += 1
    AC._session['select_event'] = _Event()
    try:
        AC.agents_idle()
        AC.select_touched()
        check('ui actions queue in order for the main thread',
              AC._session['ui_actions'] == ['idle', 'select'] and _Event.raised == 2)
    finally:
        AC._session['select_event'] = None
        AC._session['ui_actions'] = []
    check('an unverified client id is mapped to no app at all',
          AC.agent_style('local-agent-mode-t3lab')[0] not in ('Claude', 'Antigravity'))
    check('no name at all reads "AI app", not Claude',
          AC.agent_style('')[0] == 'AI app' and AC.agent_style(None)[0] == 'AI app')
    check('animation calm by default: the cursors move (owner, 2026-11-19)',
          AC.speed() == 'calm' and AC.speed_factor() > 1.0 and not AC.motion_off())
    AC._paths().set_setting('agent_cursors_reduce_motion', True)
    AC._cache['mtime'] = None
    check('the old Less motion key is ignored (the Settings dialog saved it for everyone)',
          AC.speed() == 'calm' and not AC.motion_off())
    AC.set_speed('fast')
    check('speed saved', AC.speed() == 'fast' and AC.speed_factor() < 1.0)
    AC.set_speed('off')
    check('Off: no motion, timings as Normal',
          AC.motion_off() and AC.speed_factor() == AC.speed_factor('normal') == 1.0)
    check('speed cycles through Off', AC.next_speed('calm') == 'normal'
          and AC.next_speed('fast') == 'off' and AC.next_speed('off') == 'calm')
    AC.set_speed('nonsense')
    check('unknown speed -> the default', AC.speed() == AC.SPEED_DEFAULT == 'calm')
    AC.set_option('tile', True)
    check('option saved', AC.options()['tile'] is True)
    AC.set_option('tile', False)
    check('options text', AC.options_text().startswith('On model off · Follow on'),
          AC.options_text())
    text = AC.pros_cons_text()
    check('notice lists benefits and drawbacks',
          'Benefits' in text and 'Drawbacks' in text
          and all(p in text for p in AC.PROS) and all(c in text for c in AC.CONS))

    class Task(object):
        tool_name = 'set_parameter'
        arguments = {'element_id': 1}
        result = {'ok': True}
        exception = None
        audit = None
    AC.notify(None, Task())     # disabled: must return quietly without Revit
    AC.set_enabled(True)
    AC.notify(None, Task())     # enabled, no Revit app: swallowed, no raise
    AC.before(None, Task())
    check('notify never raises', True)
    AC.set_enabled(False)


def test_server_hook():
    print('[server hook]')
    src = open(os.path.join(LIB, 'core', 'server.py'), encoding='utf-8').read()
    tree = ast.parse(src)
    execute = None
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef) and node.name == 'MCPExternalEventHandler':
            execute = next(n for n in node.body
                           if isinstance(n, ast.FunctionDef) and n.name == 'Execute')
    check('Execute found', execute is not None)
    if execute is None:
        return
    body = ast.get_source_segment(src, execute)
    done = body.find('task.done.set()')
    hook = body.find('agent_cursors.notify(app, task)')
    pre = body.find('agent_cursors.before(app, task)')
    run = body.find('_execute_tool_in_context(')
    check('notify is called', hook > 0)
    check('after task.done.set()', 0 < done < hook, (done, hook))
    check('before() is called ahead of the tool', 0 < pre < run, (pre, run))
    if 'def _external_policy_check(' not in src:        # dev server
        check('t3lab_narrate answered before the ExternalEvent queue',
              0 < src.find("if tool_name == 't3lab_narrate':") < src.find('task = _ToolTask(tool_name, arguments, expect_document'))
        check('waiting() around the approval request',
              src.find('agent_cursors.waiting(agent, tool_name, arguments, True)')
              < src.find('ok, how, wait_ms = broker.request(')
              < src.find('agent_cursors.waiting(agent, tool_name, arguments, False)'))
    check('narrate tool registered', 'NARRATE_SCHEMA' in src)
    if 'def _external_policy_check(' in src:            # Lite server
        check('t3lab_narrate answered before the ExternalEvent queue',
              0 < src.find('if tool_name == agent_cursors.NARRATE_TOOL:') < src.find('task = _ToolTask(tool_name, arguments'))
        check('narrate gets the plan arguments',
              "args.get('tasks')" in src and "bool(args.get('done'))" in src)
        check('waiting() around the confirm dialog',
              src.find('agent_cursors.waiting(task.agent, task.tool_name, task.arguments, True)')
              < src.find('refusal = _ask_user_to_confirm(task)')
              < src.find('agent_cursors.waiting(task.agent, task.tool_name, task.arguments, False)'))
    # IronPython 2.7 parses the module too: no f-strings / py3-only syntax.
    ac_src = ''.join(open(os.path.join(LIB, 'core', f), encoding='utf-8').read()
                     for f in ('agent_cursors.py', 'agent_cursors_overlay.py',
                               'agent_cursors_glow.py'))
    check('no f-strings (IronPython)', not re.search(r'(?<![\w])[fF][rR]?[\'"]', ac_src))
    check('no nonlocal (IronPython)', 'nonlocal' not in ac_src)


def test_agent_stamp():
    print('[bridge stamps the AI app, server keeps it on the task]')
    import core.bridge as B
    current = {'port': 48884}
    B._remember_client({'method': 'initialize',
                        'params': {'clientInfo': {'name': 'claude-ai', 'version': '1'}}}, current)
    req = {'jsonrpc': '2.0', 'id': 3, 'method': 'tools/call',
           'params': {'name': 'set_parameter', 'arguments': {}, '_meta': {'x': 1}}}
    out = B._stamp_agent(req, current)
    meta = out['params']['_meta']
    check('clientInfo name stamped', meta.get(B.AGENT_META) == 'claude-ai', meta)
    check('other _meta kept', meta.get('x') == 1, meta)
    check('request itself not changed', 't3lab/agent' not in req['params']['_meta'])
    check('no client -> request unchanged', B._stamp_agent(req, {}) is req)
    check('stamped name shows as Claude', AC.agent_style(meta[B.AGENT_META])[0] == 'Claude')
    src = open(os.path.join(LIB, 'core', 'server.py'), encoding='utf-8').read()
    check('server reads t3lab/agent', "meta.get('t3lab/agent')" in src)
    check('external call passes agent to the task',
          'confirm=confirm, agent=agent)' in src
          and '_ToolTask(tool_name, arguments, confirm=confirm, agent=agent)' in src)
    bsrc = open(os.path.join(LIB, 'core', 'bridge.py'), encoding='utf-8').read()
    check('bridge stamps every tools/call once, before its own handlers',
          'request = _stamp_agent(request, current)' in bsrc
          and bsrc.index('request = _stamp_agent(request, current)')
          < bsrc.index("if name == 'list_open_documents':"))
    check('bridge probes are plumbing: no cursor, no offer',
          AC._internal(type('T', (), {'agent': AC.BRIDGE_AGENT, 'tool_name': 'list_open_documents'})())
          and not AC._internal(type('T', (), {'agent': 'codex', 'tool_name': 'list_open_documents'})())
          and AC.offer(None, AC.BRIDGE_AGENT) is False)
    check('bridge names its own document probe',
          "'_meta': {'t3lab/agent': BRIDGE_AGENT}" in bsrc
          and "BRIDGE_AGENT = '{}'".format(AC.BRIDGE_AGENT) in bsrc)
    import importlib.util as _ilu
    _spec = _ilu.spec_from_file_location('t3_bridge', os.path.join(LIB, 'core', 'bridge.py'))
    BR = _ilu.module_from_spec(_spec)
    _spec.loader.exec_module(BR)
    home = 'C:\\Users\\claude'
    check('bridge: the launching app names the caller',
          BR.host_agent(['C:\\Python312\\python.exe',
                         home + '\\AppData\\Local\\Programs\\Antigravity\\Antigravity.exe'],
                        home) == 'antigravity'
          and BR.host_agent([home + '\\AppData\\Local\\AnthropicClaude\\app-1\\claude.exe'],
                            home) == 'claude-desktop'
          and BR.host_agent(['Codex.exe']) == 'codex' and BR.host_agent(['Cursor.exe']) == 'cursor')
    check('bridge: nearest ancestor wins, the user name names nobody',
          BR.host_agent(['Cursor.exe', 'Antigravity.exe']) == 'cursor'
          and BR.host_agent([home + '\\AppData\\Local\\Programs\\Python\\python.exe',
                             'C:\\Windows\\explorer.exe'], home) == '')
    BR._log_client = lambda current, name='': None      # no file / stderr in tests

    def named(info, **current):
        BR._remember_client({'params': {'clientInfo': {'name': info}}}, current)
        return current.get('client')
    check('bridge: --agent > known clientInfo > launching app > other clientInfo',
          named('claude-ai', client='antigravity', agent_explicit=True) == 'antigravity'
          and named('claude-code', host='antigravity', client='antigravity') == 'claude-code'
          and named('local-agent-mode-x', host='antigravity', client='antigravity') == 'antigravity'
          and named('some-app') == 'some-app' and named('') is None)
    check('bridge: ancestors off Windows are simply unknown',
          os.name == 'nt' or BR._ancestor_images() == [])
    check('server: a call with no name is "AI app", never claude-desktop',
          "or 'mcp-client'" in src and "or 'claude-desktop'" not in src)
    check('--agent from Connect wins over the self-declared clientInfo',
          "arg.startswith('--agent=')" in bsrc and "current['agent_explicit'] = True" in bsrc
          and "if current.get('agent_explicit'):" in bsrc)


def test_settings_dialog():
    print('[settings window: BatchOut-style cards, a checkbox per option]')
    btn = open(os.path.join(REPO, *BUTTON.split('/')), encoding='utf-8').read()
    check('the button opens the settings window and saves what it returns',
          'from GUI.AgentCursorsDialog import ask' in btn
          and 'ask(on, agent_cursors.options(), agent_cursors.pros_cons_sections())' in btn
          and 'agent_cursors.set_option(name, value)' in btn)
    check('window cannot be built: the plain checkbox dialog, which says why',
          'confirm_options(' in btn and "'(settings window unavailable: {})'.format(ex)" in btn
          and 'for name, _key, _default in agent_cursors.OPTIONS' in btn
          and "('enabled', 'Show agent cursors'" in btn)
    check('no motion row: motion is not an option (owner, 2026-11-19)',
          'animate' not in btn.lower() and 'set_speed' not in btn and 'Less motion' not in btn)
    sections = AC.pros_cons_sections()
    check('sections: Benefits / Drawbacks, every line, a T3 glyph each',
          [(h, lines) for h, lines, _g in sections]
          == [('Benefits', list(AC.PROS)), ('Drawbacks', list(AC.CONS))]
          and [g for _h, _l, g in sections] == ['\uE73E', '\uE7BA'])
    xaml = open(os.path.join(LIB, 'GUI', 'Tools', 'AgentCursorsSettings.xaml'),
                encoding='utf-8').read()
    body = xaml.split('</Window.Resources>', 1)[1]
    names = re.findall(r'x:Name="opt_(\w+)"', body)
    check('every option has its checkbox (a new option fails here until it gets one)',
          set(n for n, _k, _d in AC.OPTIONS) <= set(names) and len(names) == len(set(names)),
          (sorted(names), [n for n, _k, _d in AC.OPTIONS]))
    import xml.etree.ElementTree as ET
    ui, xname = '{http://schemas.microsoft.com/winfx/2006/xaml/presentation}', \
        '{http://schemas.microsoft.com/winfx/2006/xaml}Name'
    groups, boxes, group = {}, {}, None
    for node in ET.fromstring(xaml.encode('utf-8')).iter():      # document order
        if node.tag == ui + 'TextBlock' and node.get('FontSize') == '12.5':
            group = node.get('Text')
        elif node.tag == ui + 'CheckBox' and (node.get(xname) or '').startswith('opt_'):
            groups.setdefault(group, []).append(node.get(xname)[4:])
            boxes[node.get(xname)[4:]] = (node.get('Content'), node.get('ToolTip'))
    check("the bar's own groups, in its order, each with the same options (owner, 2026-11-19)",
          [(g, groups.get(g)) for g, _n in AC.OPTION_GROUPS]
          == [(g, list(names)) for g, names in AC.OPTION_GROUPS], groups)
    root = ET.fromstring(xaml.encode('utf-8'))
    parent = dict((child, node) for node in root.iter() for child in node)
    named = dict((node.get(xname), node) for node in root.iter() if node.get(xname))

    def left(value):            # WPF thickness: "16", "12,8" or "0,4,0,16"
        return float((value or '0').split(',')[0])
    check('checkboxes line up: the on/off box sits where the option boxes do (owner, 2026-11-19)',
          left(parent[named['chk_enabled']].get('Margin')) == left(named['options_body'].get('Padding'))
          and all(left(named['opt_' + n].get('Margin')) == 0 for n in boxes),
          (parent[named['chk_enabled']].get('Margin'), named['options_body'].get('Padding')))
    check('the folded card starts on the same line',
          left(named['details_header'][0].get('Margin')) == left(named['details_body'].get('Padding')))
    check('each checkbox is called what its switch is called on the bar',
          [(n, boxes.get(n, ('', ''))[0]) for n, _k, _d in AC.OPTIONS]
          == [(n, AC.OPTION_LABELS[n][0]) for n, _k, _d in AC.OPTIONS], boxes)
    check('and says what it does: the same text as on the bar (tooltip; "On model" spells it out)',
          all(boxes.get(n, ('', ''))[1] == AC.OPTION_LABELS[n][1] if n != 'model'
              else AC.OPTION_LABELS[n][1] in body for n, _k, _d in AC.OPTIONS))
    def title_bar(text):
        bar = text.split('<Border Grid.Row="0" Style="{StaticResource T3.TitleBar}">', 1)[1]
        bar = bar.split('<!-- ══════════════════ ROW 1', 1)[0]
        return re.sub(r'<TextBlock Text="[^"]*" (Style="\{StaticResource T3\.(Title|Subtitle)\}")',
                      r'<TextBlock Text="…" \1', ' '.join(bar.split()))
    batchout = open(os.path.join(LIB, 'GUI', 'Tools', 'ExportManager.xaml'), encoding='utf-8').read()
    check("title bar = BatchOut's, letter for letter but for its two texts (owner, 2026-11-19)",
          title_bar(xaml) == title_bar(batchout) and title_bar(xaml).count('Text="…"') == 2)
    check('so: logo, minimize / maximize / close (X = Esc), the title area drags the window',
          all(t in body for t in ('x:Name="logo_image"', 'Click="minimize_button_clicked"',
                                  'Click="maximize_button_clicked"', 'Click="close_button_clicked"'))
          and re.search(r'x:Name="btn_close"[^>]*IsCancel="True"', body) is not None
          and 'ResizeMode="CanResize"' in xaml and 'ShowInTaskbar="True"' in xaml)
    check('footer kept: copyright, Cancel + one primary with Enter',
          body.count('Style="{StaticResource T3.Copyright}"') == 1
          and 'Style="{StaticResource T3.FooterBar}"' in body
          and body.index('x:Name="btn_cancel"') < body.index('x:Name="btn_ok"')
          and body.count('T3.Button.Primary') == 1 and 'IsDefault="True" Click="ok_clicked"' in body)
    dsrc0 = open(os.path.join(LIB, 'GUI', 'WPF_Base.py'), encoding='utf-8').read()
    check('the three window buttons are handled by the base window',
          all('def {}(self, sender=None, e=None):'.format(h) in dsrc0 for h in (
              'minimize_button_clicked', 'maximize_button_clicked', 'close_button_clicked')))
    check('BatchOut card: sunken header with the checkbox, two columns and a rule between',
          'x:Name="chk_enabled"' in body and body.count('{StaticResource T3.SurfaceSunken}') == 2
          and body.count('<ColumnDefinition Width="32"/>') == 2
          and len(re.findall(r'<Border [^>]*Width="1"\s+Background="\{StaticResource T3\.Border\}"',
                             body)) == 2)
    check('benefits card is folded until its header is clicked',
          'MouseLeftButtonUp="details_header_clicked"' in body
          and re.search(r'x:Name="details_body"[^>]*Visibility="Collapsed"', body) is not None
          and 'x:Name="pnl_benefits"' in body and 'x:Name="pnl_drawbacks"' in body)
    check('"On model" row is there for T3LabLite and hidden until OPTIONS has it',
          re.search(r'x:Name="row_model"[^>]*Visibility="Collapsed"', body) is not None
          and 'x:Name="opt_model"' in body)
    check('all user-facing text is ASCII English',
          all(ord(c) < 128 for t in re.findall(r'(?:Text|Content|ToolTip)="([^"&]*)"', body)
              for c in t))
    dsrc = open(os.path.join(LIB, 'GUI', 'AgentCursorsDialog.py'), encoding='utf-8').read()
    check('window class: one base, no settings I/O, T3 names looked up guarded',
          'class AgentCursorsSettingsWindow(_Window):' in dsrc and 'from core' not in dsrc
          and 'import agent_cursors' not in dsrc and 'set_option' not in dsrc
          and "self.FindResource('T3.Ink' if on else 'T3.Border')" in dsrc)
    tsrc = open(os.path.join(LIB, 'GUI', 'T3Dialog.py'), encoding='utf-8').read()
    check('the shared T3Dialog is left as it was (plain checkbox rows)',
          "FindResource('T3.CheckBox')" in tsrc and 'def confirm_options(' in tsrc
          and 'sections' not in tsrc)


def test_action_band():
    print('[action bands: colour rule, rows, fills, counts bar]')
    acts = [a[0] for a in AC.ACTIONS]
    for theme in ('dark', 'light'):
        pal = AC.hud_palette(theme)
        wash = AC.band_wash(theme)
        worst = min(AC.contrast(AC.band_style(a, theme)[2],
                                AC.blend(AC.band_style(a, theme)[2], pal['bg'], wash))
                    for a in acts)
        check(theme + ': every action glyph reads 3:1 on its own wash', worst >= 3.0, worst)
        text = min(AC.contrast(pal['text'], AC.blend(AC.band_style(a, theme, s, f)[2], pal['bg'],
                                                     wash * 2))
                   for a in acts for s in ('done', 'waiting') for f in (False, True))
        check(theme + ': caption reads 4.5:1 even on a filled (replay) band', text >= 4.5, text)
        for state, failed in (('waiting', False), ('done', True)):
            color = AC.band_style('delete', theme, state, failed)[2]
            check('{}: {} band reads 3:1'.format(theme, 'failed' if failed else state),
                  AC.contrast(color, AC.blend(color, pal['bg'], wash)) >= 3.0)
    check('raw purple is too dim on the dark bar: the reason bar_tone exists',
          AC.contrast('#7C3AED', AC.blend('#7C3AED', AC.HUD_PALETTES['dark']['bg'],
                                          AC.BAND_WASH['dark'])) < 3.0
          and AC.bar_tone('#7C3AED', 'dark') != '#7C3AED')
    check('a colour that already reads is left alone',
          AC.bar_tone('#7C3AED', 'light') == '#7C3AED' and AC.bar_tone('#D97706', 'dark') == '#D97706')
    check('tone stays the same hue family (moved towards white, not replaced)',
          AC.contrast(AC.bar_tone('#2563EB', 'dark'), '#FFFFFF') < AC.contrast('#2563EB', '#FFFFFF'))
    check('blend', AC.blend('#FFFFFF', '#000000', 0) == '#000000'
          and AC.blend('#FFFFFF', '#000000', 255) == '#FFFFFF'
          and AC.blend('#FF0000', '#0000FF', 128) == '#80007F')
    check('contrast', round(AC.contrast('#000000', '#FFFFFF'), 1) == 21.0
          and AC.contrast('#777777', '#777777') == 1.0)

    label, glyph, color = AC.band_style('create', 'light')
    check('band = the action label, glyph and colour',
          (label, glyph) == AC.action_style('create')[:2]
          and color == AC.bar_tone(AC.action_style('create')[2], 'light'))
    flabel, fglyph, fcolor = AC.band_style('create', 'light', failed=True)
    check('failed: red and a cross, the label still names the action',
          flabel == 'Create' and fglyph != glyph
          and fcolor == AC.bar_tone(AC.HUD_PALETTES['light']['bad'], 'light'))
    check('waiting: amber', AC.band_style('create', 'dark', 'waiting')[2]
          == AC.bar_tone(AC.HUD_PALETTES['dark']['warn'], 'dark'))
    check('unknown action -> modify', AC.band_style(None, 'dark')[0] == 'Modify')

    br = AC.band_rows
    check('nothing yet: no band', br() == ([], 0))
    last = {'agent': 'codex', 'action': 'delete', 'caption': 'Failed: Deleting element',
            'failed': True, 't': 1.0}
    rows, more = br(last=last)
    check('idle: the last step', more == 0 and len(rows) == 1 and rows[0]['state'] == 'done'
          and rows[0]['failed'] and rows[0]['note'] == 'Delete' and rows[0]['progress'] is None,
          rows)
    rows, _more = br(last=last, replay=(3, 20))
    check('replay: the step, "Replay 3 / 20" and its progress',
          rows[0]['state'] == 'replay' and rows[0]['progress'] == (3, 20)
          and rows[0]['note'] == 'Replay 3 / 20' and rows[0]['caption'] == last['caption'])
    working = [('codex', 'delete', 'Deleting'), ('claude-desktop', 'read', 'Listing')]
    rows, more = br(working, last=last, replay=(3, 20))
    check('live calls win over the last step and a replay; one band per agent, agent order',
          more == 0 and [(r['agent'], r['state'], r['note']) for r in rows]
          == [('claude-desktop', 'running', 'Read'), ('codex', 'running', 'Delete')], rows)
    rows, _more = br(working, [('codex', 'delete', 'Deleting')])
    check('a call waiting for the OK: its band says Waiting',
          [(r['agent'], r['state'], r['note']) for r in rows]
          == [('claude-desktop', 'running', 'Read'), ('codex', 'waiting', 'Waiting')], rows)
    many = [('agent-{}'.format(k), 'read', 'x') for k in range(5)]
    rows, more = br(many)
    check('more agents than bands: the rest is counted', len(rows) == AC.BAND_MAX and more == 2)

    bm = AC.band_motion
    one = br([('claude-desktop', 'read', 'Listing')])[0]
    motions, seen = bm(one, {})
    check('motion: a band that was not there enters', motions == ['enter'])
    check('motion: a plain redraw moves nothing', bm(one, seen)[0] == [None])
    two = br([('claude-desktop', 'read', 'Listing'), ('codex', 'delete', 'Deleting')])[0]
    check('motion: only the new agent enters', bm(two, seen)[0] == [None, 'enter'])
    done = br(last={'agent': 'claude-desktop', 'action': 'read', 'caption': 'Listing'})[0]
    check('motion: running -> done pops', bm(done, seen)[0] == ['pop'])
    check('motion: a failed step shakes', bm(br(last=last)[0], {'codex': ('x',)})[0] == ['shake'])
    step3, step4 = br(last=last, replay=(3, 20))[0], br(last=last, replay=(4, 20))[0]
    check('motion: every replay step moves, even with the same caption',
          bm(step4, bm(step3, {})[1])[0] == ['shake'] and bm(step3, bm(step3, {})[1])[0] == [None])
    check('motion: only the agents on the bar are remembered',
          sorted(bm(two, {'gone': ('x',)})[1]) == ['claude-desktop', 'codex'])

    fw = AC.fill_width
    check('fill width', fw(1, 4) == AC.BAR_W / 4 and fw(20, 20) == AC.BAR_W
          and fw(0, 5) == 0.0 and fw(0, 5, minimum=4.0) == 4.0)
    check('fill width never leaves the bar', fw(9, 5) == AC.BAR_W and fw(-1, 5) == 0.0
          and fw(1, 0) == 0.0 and fw(None, 3) == 0.0)

    cb = AC.count_bar
    check('no counts: no bar', cb({}) == [] and cb(None) == [] and cb({'read': 0}) == [])
    check('one action fills the bar', cb({'code': 9}) == [('code', AC.COUNT_BAR_W)])
    bar = cb({'create': 1, 'read': 3}, 41.0, 1.0)
    check('widths follow the counts, ACTIONS order, gap included',
          bar == [('read', 30.0), ('create', 10.0)], bar)
    bar = cb({'read': 1, 'code': 27}, 56.0, 1.0, 3.0)
    check('1 call in 28 still shows (minimum width)',
          bar[0] == ('read', 3.0) and abs(sum(w for _a, w in bar) + 1.0 - 56.0) < 1e-9, bar)
    full = cb(dict((a, 1 + k * 40) for k, a in enumerate(acts)), 56.0, 1.0, 3.0)
    check('every action at once: all shown, still the same width',
          len(full) == len(acts) and min(w for _a, w in full) >= 3.0
          and abs(sum(w for _a, w in full) + (len(acts) - 1) - 56.0) < 1e-9, full)
    tight = cb({'read': 1, 'create': 1, 'modify': 1}, 8.0, 1.0, 3.0)
    check('no room for the minimum: an equal split', [w for _a, w in tight] == [2.0, 2.0, 2.0], tight)

    osrc = open(os.path.join(LIB, 'core', 'agent_cursors_overlay.py'), encoding='utf-8').read()
    check('the bar is built from these helpers (no second colour rule in the overlay)',
          'band_rows(working, waiting, last, replay)' in osrc and 'def _action_band(' in osrc
          and 'def _current_card(' not in osrc and '396.0' not in osrc
          and 'fill_width(done, total, BAR_W' in osrc)
    check('Animation Off = no motion', 'return motion_off()' in osrc
          and "options().get('reduce')" not in osrc)
    check('no motion control on the bar (owner, 2026-11-19): the key is the only switch',
          'set_speed' not in osrc and 'SPEEDS' not in osrc
          and [g for g, _n in AC.OPTION_GROUPS] == ['Show', 'View', 'Selection', 'Effects', 'Bar'])
    check('bar motion is decoration: skipped when off, never fatal',
          'def _motion(self, fn, *args):' in osrc
          and 'if self._reduce():\n            return\n        try:\n            fn(*args)' in osrc)


def test_hud_redesign():
    print('[bar redesign: state, progress, fold, place, pause, minimal cursor]')
    hs = AC.hud_state
    check('state: question > pause > working > idle',
          hs({}, {}, [1], True) == 'waiting' and hs({}, {'k': 1}, [], False) == 'waiting'
          and hs({'k': 1}, {}, [], True) == 'paused' and hs({'k': 1}, {}, [], False) == 'working'
          and hs({}, {}, [], False) == 'idle')
    check('status text', AC.status_text('working', 8.6) == 'Working 8 s'
          and AC.status_text('working', 0.2) == 'Working' and AC.status_text('paused') == 'Paused')
    check('plan progress', AC.plan_progress({'tasks': list('abcde'), 'step': 4}) == (3, 5)
          and AC.plan_progress({'tasks': list('ab'), 'step': 9}) == (2, 2)
          and AC.plan_progress({}) is None and AC.plan_progress(None) is None)
    check('ago text', [AC.ago_text(x) for x in (0, 9, 125, 7300)] == ['now', '9 s', '2 min', '2 h'])
    check('one-line bar text', AC.compact_line('Moving 12 walls', (3, 5)) == u'Moving 12 walls \u00b7 4/5'
          and AC.compact_line('x' * 80).endswith(u'\u2026') and len(AC.compact_line('x' * 80)) == 36)
    check('footer', AC.footer_text(1, 0, 0) == '1 action'
          and AC.footer_text(12, 1, 2) == u'12 actions \u00b7 1 failed \u00b7 2 views')
    sc = AC.should_collapse
    check('fold: hover opens, user fold wins, idle folds, a question opens',
          not sc(True, 99, True, []) and sc(True, 0, False, [1])
          and sc(False, AC.COLLAPSE_S, False, []) and not sc(False, 1, False, [])
          and not sc(False, 99, False, [1]))
    check('offset parse', AC.parse_offset('12,-30') == (12.0, -30.0) and AC.parse_offset('') == (0, 0)
          and AC.parse_offset('bad') == (0, 0))
    view = (0, 0, 1000, 800)
    check('bar placed bottom centre', AC.hud_place(view, (420, 200)) == (290.0, 584.0))
    check('bar kept inside the view', AC.hud_place(view, (420, 200), (5000, 5000)) == (576, 596)
          and AC.hud_place(view, (420, 200), (-5000, -5000)) == (4, 4))
    check('cluster', AC.union_box([(1, 5, 3, 9), (0, 6, 8, 7)]) == (0, 5, 8, 9)
          and AC.cluster_stops(AC.MAX_STOPS + 1) and not AC.cluster_stops(AC.MAX_STOPS))
    for theme in ('dark', 'light'):
        pal = AC.hud_palette(theme)
        check('palette ' + theme + ' complete',
              set(pal) == set(AC.HUD_PALETTES['dark']) and all(pal[k] for k in AC.STATES.values()
                                                              for k in [k[1]]))
    grouped = [n for _g, names in AC.OPTION_GROUPS for n in names]
    check('options drawer lists every option once',
          sorted(grouped) == sorted(n for n, _k, _d in AC.OPTIONS) and len(grouped) == len(set(grouped)))
    AC.set_paused(False)
    check('not paused: nothing held', AC.paused_refusal('set_parameter', True) is None)
    AC.set_paused(True)
    try:
        held = AC.paused_refusal('set_parameter', True)
        check('paused: a write is held, nothing ran',
              held and held['error_code'] == 'paused' and held['nothing_ran'] and 'Resume' in held['error'])
        check('paused: reads and narrate still run',
              AC.paused_refusal('list_views', False) is None
              and AC.paused_refusal(AC.NARRATE_TOOL, True) is None)
    finally:
        AC.set_paused(False)
    osrc = open(os.path.join(LIB, 'core', 'agent_cursors_overlay.py'), encoding='utf-8').read()
    check('minimal cursor: bare tip + one tag (name · what it does)',
          'CURSOR_PATH' in osrc and "'tool_pill'" not in osrc and "cur['tool']" not in osrc
          and 'def _cursor_state(self, cur, state, caption=None):' in osrc)
    check('avatar on the bar is the cursor', 'def _mark(self, color, size=16):' in osrc
          and 'Geometry.Parse(CURSOR_PATH)' in osrc.split('def _mark(', 1)[1].split('    def ', 1)[0])
    check('options drawer replaces the chips', 'for group, names in OPTION_GROUPS' in osrc
          and 'WrapPanel()\n        self._toggles' not in osrc)
    check('header: pause, options, fold, close', all(t in osrc for t in (
        "'Pause: hold model changes", "'Options'", "'Fold to one line", "'Turn agent cursors off")))
    cursor = osrc.split('    def _cursor(self, agent, action=None):', 1)[1].split('\n    def ', 1)[0]
    check('no white outline on the cursor, its avatar, box corners or edge arrows',
          'Stroke = self._brush(' not in cursor and "handle.Fill = self._brush(color)" in osrc
          and "arrow.Stroke = self._brush('#FFFFFF')" not in osrc)
    check('working: a pulsing dot, not a dashed ring', 'StrokeDashArray' not in cursor)
    check('drag to move, remembered', 'self.hud.DragMove()' in osrc and 'set_hud_offset(' in osrc
          and 'e.ClickCount == 2' in osrc)
    check('buttons never start a drag', osrc.count("'MouseLeftButtonDown', _swallow") >= 2)
    check('theme follows Revit', 'def set_theme(self, theme):' in osrc)
    acsrc = open(os.path.join(LIB, 'core', 'agent_cursors.py'), encoding='utf-8').read()
    check('glow switches itself off when Revit refuses it', "set_option('glow', False)" in acsrc)
    ssrc = open(os.path.join(REPO, *SERVER.split('/')), encoding='utf-8').read()
    gate = ssrc.index('agent_cursors.paused_refusal(task.tool_name, is_write)')
    check('server holds writes before they run',
          gate < ssrc.index('self.server._write_in_progress = True', gate)
          and gate < ssrc.index('self.server._execute_tool_in_context(', gate))


class _FakeEvent(object):
    def __init__(self):
        self.handlers = []

    def __iadd__(self, fn):
        self.handlers.append(fn)
        return self

    def __isub__(self, fn):
        self.handlers = [h for h in self.handlers if h is not fn]
        return self


class _FakeArgs(object):
    def __init__(self):
        self.no_delay = 0

    def SetRaiseWithoutDelay(self):
        self.no_delay += 1


class _FakeApp(object):
    def __init__(self):
        self.Idling = _FakeEvent()

    def idle(self):
        args = _FakeArgs()
        for h in list(self.Idling.handlers):
            h(self, args)
        return args


def test_stability():
    print('[stability: front rule, settle before drawing, glow redraw]')
    view = (0, 0, 1000, 800)
    fh = AC.front_hides
    check('Revit in front: shown', not fh(True, False, None, view))
    check('AI app beside Revit / other screen: shown',
          not fh(False, False, (1000, 0, 1900, 800), view)
          and not fh(False, False, (2000, 0, 3000, 900), view))
    check('app over the views: hidden', fh(False, False, (200, 100, 900, 700), view))
    check('a sliver of overlap is not covering', not fh(False, False, (995, 0, 1900, 800), view))
    check('Revit minimized: hidden even if in front', fh(True, True, None, view))
    check('unknown front window or canvas: shown',
          not fh(False, False, None, view) and not fh(False, False, (0, 0, 9, 9), None))

    g1 = [{'id': 5, 'rect': (0, 0, 800, 600), 'origin': (1.0, 2.0, 0.0), 'w': 10.0, 'h': 8.0}]
    g2 = [dict(g1[0], rect=(0, 0, 400, 600))]
    check('layout key: same layout equal, tiled differs',
          AC.layout_key(g1, 0) == AC.layout_key([dict(g1[0])], 0)
          and AC.layout_key(g1, 0) != AC.layout_key(g2, 0)
          and AC.layout_key(g1, 0) != AC.layout_key(g1, None))

    layouts = [(None, [], None)]
    saved = AC._open_views
    AC._open_views = lambda app: layouts[-1]
    try:
        shown = []
        app = _FakeApp()
        layouts.append((None, g1, 0))
        check('deferred: hooked on Idling', AC._when_settled(app, lambda: shown.append('a')))
        args = app.idle()
        check('deferred: not drawn while the layout is new', shown == [] and args.no_delay == 1)
        layouts.append((None, g2, 0))           # Tile lands
        app.idle()
        check('deferred: not drawn while views still move', shown == [])
        app.idle()                              # same layout twice: settled
        check('deferred: drawn once settled, handler removed',
              shown == ['a'] and app.Idling.handlers == [])
        app2 = _FakeApp()
        AC._when_settled(app2, lambda: shown.append('b'))
        AC._flush_pending()                     # the next call starts
        check('next call flushes the pending show first',
              shown == ['a', 'b'] and app2.Idling.handlers == [])
        app3 = _FakeApp()
        AC._when_settled(app3, lambda: shown.append('c'))
        for _ in range(AC.SETTLE_MAX_TICKS):
            layouts.append((None, [dict(g1[0], w=float(len(layouts)))], 0))   # never settles
            app3.idle()
        check('never settles: drawn after SETTLE_MAX_TICKS', shown[-1] == 'c')
    finally:
        AC._open_views = saved
        AC._flush_pending()

    acsrc = open(os.path.join(LIB, 'core', 'agent_cursors.py'), encoding='utf-8').read()
    notify = acsrc.split('def notify(app, task):', 1)[1].split('\ndef ', 1)[0]
    check('notify: tile decided before measuring, show deferred after follow / tile',
          notify.index('tiled = _tile(app)') < notify.index('def finish():')
          and '(followed or tiled) and _when_settled(app, finish)' in notify)
    before = acsrc.split('def before(app, task):', 1)[1].split('\ndef ', 1)[0]
    check('before: flushes the previous show', '_flush_pending()' in before)
    gsrc = open(os.path.join(LIB, 'core', 'agent_cursors_glow.py'), encoding='utf-8').read()
    check('glow: no full view redraw when nothing new is outlined',
          'if not same:\n' in gsrc.split('def update(app, ids):', 1)[1])
    osrc = open(os.path.join(LIB, 'core', 'agent_cursors_overlay.py'), encoding='utf-8').read()
    place = osrc.split('    def _place(self, rect):', 1)[1].split('\n    def ', 1)[0]
    check('overlay: no Show() while Revit is covered (no flash per call)',
          'if self._front_ok:' in place and 'self._clear_effects()' in place)
    check('overlay: hides by covering, not by focus', 'front_hides(' in osrc
          and '_process_in_front' not in osrc)


def test_bar_only():
    print('[bar only by default: no cursor, boxes or effects on the model]')
    acsrc = open(os.path.join(LIB, 'core', 'agent_cursors.py'), encoding='utf-8').read()
    notify = acsrc.split('def notify(app, task):', 1)[1].split('\ndef ', 1)[0]
    check('notify: options pass through effective()', 'opts = effective(options())' in notify
          and "_snapshot(app, ids if opts['model'] else [], pre)" in notify)
    before = acsrc.split('def before(app, task):', 1)[1].split('\ndef ', 1)[0]
    check('before: no boxes measured without On model',
          'model and ids and action in' in before and 'elif model and ids:' in before)
    osrc = open(os.path.join(LIB, 'core', 'agent_cursors_overlay.py'), encoding='utf-8').read()

    def body(name):
        return osrc.split('    def {}('.format(name), 1)[1].split('\n    def ', 1)[0]
    show = body('show')
    check('show: draws on the model only with On model, the bar always',
          'if self._model():\n            try:\n                self._draw_step(' in show
          and 'self._bar_only()' in show
          and 'self._log_step(' in show and 'self._show_hud(' in show
          and 'self._cursor(' not in show and 'self._place(' not in show)
    working = body('working')
    check('working: no cursor or canvas without On model, the bar always',
          'if not self._model():\n            self._bar_only()' in working
          and 'elif rect is not None:\n            try:\n                self._draw_working(' in working
          and 'self._cursor(' not in working and 'self._place(' not in working
          and 'self._show_hud(rect)' in working
          and 'self._cursor(' in body('_draw_working'))
    check('ask: no canvas without On model', 'if self._model():\n            self._place(rect)'
          in body('ask'))
    check('waiting: no cursor made without On model',
          'self.win.IsVisible and self._model()' in body('waiting'))
    check('bar only: canvas hidden, front tick never brings it back',
          'self.win.Hide()' in body('_bar_only')
          and 'if window is self.win and not model:' in body('_on_front_tick'))
    groups = dict(AC.OPTION_GROUPS)
    check('options drawer: On model first', AC.OPTION_GROUPS[0] == ('Show', ('model',))
          and 'model' in AC.OPTION_LABELS and groups)


def test_harness_hud():
    """Several agents, one task board (docs/harness.md)."""
    import json
    import types
    print('[harness on the HUD]')
    check('task text from the bridge stamp',
          AC.harness_task_text('a1b2c3d4/doors', 'Check  the doors') == 'Task doors · Check the doors')
    check('no task -> empty', AC.harness_task_text(None, 'x') == '' and
          AC.harness_task_text('s/k', '') == 'Task k')
    check('long task title cut', len(AC.harness_task_text('s/k', 'x' * 200)) <= AC.TASK_MAX + 10)

    tag, colour = AC.agent_style('claude-desktop/Arch')
    check('a chat of a client has its own tag', tag == 'Claude · arch', tag)
    check('and its own colour', colour != AC.agent_style('claude-desktop')[1]
          and colour != AC.agent_style('claude-desktop/mep')[1], colour)

    # What core/server.py puts on a _ToolTask from the bridge's stamps.
    task = types.SimpleNamespace(agent='claude-desktop', member='claude-desktop/arch',
                                 task_ref='s1/a', task_title='Fix marks', source='mcp',
                                 tool_name='set_parameter')
    check('cursor is the harness member', AC._agent_of(task) == 'claude-desktop/arch')
    check('cursor names the harness task', AC._task_text(task, 'claude-desktop/arch')
          == 'Task a · Fix marks')
    AC.narrate('claude-desktop/arch', 'goal', tasks=['one', 'two'])
    check('the agent\'s own plan step wins', AC._task_text(task, 'claude-desktop/arch')
          == 'Step 1/2 · one')
    plain = types.SimpleNamespace(agent='codex', tool_name='x')
    check('no member, no task -> plain agent', AC._agent_of(plain) == 'codex'
          and AC._task_text(plain, 'codex') == '')

    folder = tempfile.mkdtemp(prefix='t3board_')
    old_dir = os.environ.get('T3LAB_HARNESS_DIR')
    os.environ['T3LAB_HARNESS_DIR'] = folder
    path = os.path.join(folder, AC.BOARD_FILE)
    check('no board file -> {}', AC.read_board() == {} and AC.board_rows({}) == [])
    board = {'format': 'harness-board/1', 'updated': 1.0, 'sessions': [{
        'session': 'a1b2c3d4', 'goal': 'Audit levels', 'document': 'A.rvt',
        'counts': {'claimed': 2, 'review': 1, 'done': 1, 'todo': 1, 'failed': 1},
        'tasks': [
            {'task': 'r', 'title': 'Count rooms', 'status': 'review', 'owner': 'codex'},
            {'task': 'd', 'title': 'Old', 'status': 'done', 'owner': 'codex'},
            {'task': 'a', 'title': 'Fix marks', 'status': 'claimed',
             'owner': 'claude-desktop/arch', 'document': 'a.RVT'},
            {'task': 'b', 'title': 'Check links', 'status': 'claimed',
             'owner': 'console-qa', 'document': 'B.rvt'},
            {'task': 'f', 'title': 'Purge', 'status': 'failed', 'owner': 'codex'},
            {'task': 't', 'title': 'Later', 'status': 'todo', 'owner': None}]},
        {'session': 'empty', 'goal': 'no tasks yet', 'tasks': []}]}
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(board, f)
    read = AC.read_board()
    check('board read from board.json', read.get('updated') == 1.0)
    check('board cached until the file changes', AC.read_board() is read)
    rows = AC.board_rows(read)
    check('one HUD block per session with tasks', [r['session'] for r in rows] == ['a1b2c3d4'])
    check('board summary', rows[0]['summary'] == '2 working · 1 in review · 1 failed · 1/6 done',
          rows[0]['summary'])
    check('board rows: working, review, failed; done / todo left out',
          rows[0]['rows'] == [('current', 'Claude · arch · Fix marks', ''),
                              ('current', 'Console QA · Check links', 'B.rvt'),
                              ('review', 'GPT · Count rooms', ''),
                              ('failed', 'GPT · Purge', '')], rows[0]['rows'])
    check('board rows capped', AC.board_rows(read, limit=2)[0]['more'] == 2)
    with open(path, 'w', encoding='utf-8') as f:
        f.write('{"updated": 2.0, "sessions": [')           # the bridge is mid-write
    check('half-written board keeps the last good one', AC.read_board().get('updated') == 1.0)
    os.remove(path)
    check('board gone -> {}', AC.read_board() == {})
    shutil.rmtree(folder, ignore_errors=True)
    if old_dir is None:
        os.environ.pop('T3LAB_HARNESS_DIR', None)
    else:
        os.environ['T3LAB_HARNESS_DIR'] = old_dir

    with open(os.path.join(LIB, 'core', 'agent_cursors_overlay.py'), encoding='utf-8') as f:
        osrc = f.read()
    check('the HUD polls board.json every tick', 'self._poll_board()' in osrc
          and 'board_rows(board)' in osrc)
    check('the HUD shows the board section', 'full.Children.Add(self._board_panel)' in osrc)
    with open(os.path.join(LIB, 'core', 'server.py'), encoding='utf-8') as f:
        ssrc = f.read()
    check('the server hands the stamps to the task',
          'task.task_ref, task.task_title, task.member = harness' in ssrc)


if __name__ == '__main__':
    try:
        test_extract_ids()
        test_view_ids()
        test_actions()
        test_captions_and_tour()
        test_activity_helpers()
        test_perspective()
        test_narrate_and_glow()
        test_views()
        test_projection()
        test_agent_style()
        test_agent_names()
        test_visible_work()
        test_setting()
        test_server_hook()
        test_agent_stamp()
        test_settings_dialog()
        test_action_band()
        test_hud_redesign()
        test_stability()
        test_bar_only()
        test_harness_hud()
    finally:
        shutil.rmtree(_TMP, ignore_errors=True)
    print()
    if FAILURES:
        print('FAILED: {}'.format(', '.join(FAILURES)))
        sys.exit(1)
    print('All agent cursor tests passed.')
