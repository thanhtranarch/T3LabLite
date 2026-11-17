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
    check('missing agent -> Claude', AC.agent_style(None) == AC.AGENTS['claude-desktop'])
    check('MCP client name claude-ai -> Claude', AC.agent_style('claude-ai')[0] == 'Claude')
    check('MCP client name codex-mcp-client -> GPT', AC.agent_style('codex-mcp-client')[0] == 'GPT')
    a, b = AC.agent_style('my-bot'), AC.agent_style('my-bot')
    check('unknown agent: stable colour', a == b and a[1] in AC._SPARE, a)
    for key, (_tag, color) in AC.AGENTS.items():
        if not (len(color) == 7 and color.startswith('#')):
            check('colour of ' + key, False, color)


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
    check('options default: every bar option on (owner, 2026-11-17)',
          opts == {'follow': True, 'zoom': True, 'select': True, 'tile': True,
                   'reduce': True, 'glow': True, 'keep': True}, opts)
    src = open(os.path.join(LIB, 'core', 'agent_cursors_overlay.py'), encoding='utf-8').read()
    tick = src.split('    def _on_tick(self, sender, e):', 1)[1].split('    def ', 1)[0]
    check('keep bar: never auto-hidden after the summary',
          'now - self._summary >= SUMMARY_S and not self._keep()' in tick)
    summary = src.split('    def _enter_summary(self, keep=False):', 1)[1].split('    def ', 1)[0]
    check('keep bar: idle only dims the cursors, windows stay',
          'if not keep:' in summary and summary.index('if not keep:') < summary.index('self.win.Hide()'))
    check('speed normal by default', AC.speed() == 'normal' and AC.speed_factor() == 1.0)
    AC.set_speed('calm')
    check('speed saved', AC.speed() == 'calm' and AC.speed_factor() > 1.0)
    check('speed cycles', AC.next_speed('calm') == 'normal' and AC.next_speed('fast') == 'calm')
    AC.set_speed('nonsense')
    check('unknown speed -> normal', AC.speed() == 'normal')
    AC.set_option('tile', True)
    check('option saved', AC.options()['tile'] is True)
    AC.set_option('tile', False)
    check('options text', AC.options_text().startswith('Follow on'), AC.options_text())
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
    check('bridge forwards the stamped call',
          '_forward(_stamp_agent(request, current), current, token)' in bsrc)


def test_settings_dialog():
    print('[settings dialog: every option is a checkbox]')
    import re as _re
    btn = open(os.path.join(REPO, *BUTTON.split('/')), encoding='utf-8').read()
    check('Agent Cursors uses confirm_options', 'confirm_options(' in btn and 'confirm(' not in btn.replace('confirm_options(', ''))
    check('rows come from OPTIONS (new options appear by themselves)',
          'for name, _key, _default in agent_cursors.OPTIONS' in btn and 'set_option(name, value)' in btn)
    check('on: an untickable on/off switch', "('enabled', 'Show agent cursors'" in btn)
    dsrc = open(os.path.join(LIB, 'GUI', 'T3Dialog.py'), encoding='utf-8').read()
    check('T3Dialog builds T3.CheckBox rows', "FindResource('T3.CheckBox')" in dsrc
          and 'def confirm_options(' in dsrc and 'option_values' in dsrc)
    xaml = open(os.path.join(LIB, 'GUI', 'Tools', 'T3Dialog.xaml'), encoding='utf-8').read()
    rows = _re.findall(r'Grid\.Row="(\d)"', xaml.split('txt_message', 1)[1].split('ROW 2: FOOTER', 1)[0])
    check('message row 0, pnl_options row 1, details row 2, input row 3',
          'x:Name="pnl_options" Grid.Row="1"' in xaml and rows[:4] == ['0', '1', '2', '3'], rows)


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
        test_setting()
        test_server_hook()
        test_agent_stamp()
        test_settings_dialog()
        test_hud_redesign()
    finally:
        shutil.rmtree(_TMP, ignore_errors=True)
    print()
    if FAILURES:
        print('FAILED: {}'.format(', '.join(FAILURES)))
        sys.exit(1)
    print('All agent cursor tests passed.')
