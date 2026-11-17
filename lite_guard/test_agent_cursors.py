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
    check('options default: follow + zoom on, select + tile off',
          opts == {'follow': True, 'zoom': True, 'select': False, 'tile': False}, opts)
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
    # IronPython 2.7 parses the module too: no f-strings / py3-only syntax.
    ac_src = open(AC.__file__, encoding='utf-8').read()
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


if __name__ == '__main__':
    try:
        test_extract_ids()
        test_view_ids()
        test_actions()
        test_captions_and_tour()
        test_views()
        test_projection()
        test_agent_style()
        test_setting()
        test_server_hook()
        test_agent_stamp()
    finally:
        shutil.rmtree(_TMP, ignore_errors=True)
    print()
    if FAILURES:
        print('FAILED: {}'.format(', '.join(FAILURES)))
        sys.exit(1)
    print('All agent cursor tests passed.')
