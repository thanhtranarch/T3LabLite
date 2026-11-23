# -*- coding: utf-8 -*-
"""Tool Hub recipes over MCP and the CLI (bridge t3lab_recipe). Each step is one
tools/call through the same path as an AI app's: argument check, harness write
check, then Revit (its policy). Ported from t3lab_dev dev/test_bridge_recipes.py;
Lite has no tiers / toolsets / t3lab_call.

Run:  python lite_guard/test_bridge_recipes.py      (no Revit: Revit calls are captured)
"""
import contextlib
import io
import json
import os
import sys
import tempfile

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
sys.path.insert(0, os.path.join(REPO, 'lite_guard'))
APPDATA = tempfile.mkdtemp(prefix='t3lab_lite_recipes_')
os.environ['APPDATA'] = APPDATA
os.environ['T3LAB_HARNESS_DIR'] = tempfile.mkdtemp(prefix='t3lab_lite_recipes_harness_')
os.environ.pop('T3LAB_AGENT', None)

import core.bridge as B                       # noqa: E402
import core.server as S                       # noqa: E402
import bridge_deploy as D                     # noqa: E402
import server_registry                        # noqa: E402
from Intelligence.toolhub import recipes as R  # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


SRV = object.__new__(S.T3LabAIServer)
SRV._tools = server_registry.registry()
B._save_cached_tools(SRV._handle_tools_list()['tools'])


def recipe(slug, steps, is_write, status='published', tamper=False):
    payload = {'kind': 'recipe', 'name': slug, 'is_write': is_write,
               'plan': {'steps': steps},
               'inputs_schema': {'type': 'object', 'properties': {'level': {'type': 'string'}},
                                 'required': []}}
    sha = R.payload_sha(payload)
    if tamper:
        payload = dict(payload, is_write=False)
    return {'slug': slug, 'kind': 'recipe', 'status': status, 'name': slug.title(),
            'payload': payload, 'sha256': sha}


MARK = {'tool': 'set_parameter', 'args': {'element_id': 1, 'parameter_name': 'Mark',
                                         'value': '{{input.level}}'}}
rows = {
    'levels': recipe('levels', [{'tool': 'list_levels', 'args': {}}], False),
    'mark': recipe('mark', [{'tool': 'list_levels', 'args': {}}, MARK], True),
    'typo': recipe('typo', [{'tool': 'set_parameter', 'args': {
        'element_id': 1, 'parameter_nam': 'Mark', 'value': 'A'}}], True),
    'tampered': recipe('tampered', [{'tool': 'list_levels', 'args': {}}], True, tamper=True),
    'draft': recipe('draft', [{'tool': 'list_levels', 'args': {}}], False, status='draft'),
    'loop': recipe('loop', [{'tool': 't3lab_recipe', 'args': {'action': 'list'}}], False),
}
folder = os.path.join(APPDATA, 'T3LabAI', 'toolhub')
catalog = os.path.join(folder, 'catalog.json')

print('[listing]')
names = [t['name'] for t in B._handle_tools_list({'id': 1}, {'port': 1}, 't')['result']['tools']]
check('not listed while the catalog has no recipe (no tokens spent)', 't3lab_recipe' not in names)
os.makedirs(folder)
with open(catalog, 'w', encoding='utf-8') as f:
    json.dump({'rows': rows}, f)
names = [t['name'] for t in B._handle_tools_list({'id': 1}, {'port': 1}, 't')['result']['tools']]
check('listed once a published recipe is synced', names.count('t3lab_recipe') == 1)
check('recipes module loaded by path, not through the Intelligence package',
      B._rc is not None and B._rc.__name__ == 'toolhub_recipes')

seen = []
docs = []


def _capture(request, current, token, **k):
    seen.append((request['params']['name'], request['params'].get('arguments'),
                 (request['params'].get('_meta') or {}).get(B.AGENT_META)))
    return B._tool_response(request.get('id'), {'success': True,
                                                'tool': request['params']['name']})


B._forward = _capture
B._documents_on = lambda port, token: list(docs)
person = {'port': 1, 'client': 'claude-desktop', 'agent_explicit': True}


def call(cur, args):
    resp = B._handle({'jsonrpc': '2.0', 'id': 1, 'method': 'tools/call',
                      'params': {'name': 't3lab_recipe', 'arguments': args}}, cur, 't')
    return json.loads(resp['result']['content'][0]['text'])


print('[list / show]')
r = call(person, {'action': 'list'})
check('only published recipes whose hash matches the review',
      sorted(x['slug'] for x in r['recipes']) == ['levels', 'loop', 'mark', 'typo'], r)
check('changes_model flag', dict((x['slug'], x['changes_model']) for x in r['recipes'])['mark'])
r = call(person, {'action': 'show', 'slug': 'mark'})
check('show lists steps and inputs', [s['tool'] for s in r['steps']] == ['list_levels',
                                                                       'set_parameter']
      and 'level' in r['inputs_schema']['properties'], r)
check('tampered recipe refused', call(person, {'action': 'show', 'slug': 'tampered'})
      ['error_code'] == 'recipe_not_found')
check('unknown action -> bad_arguments', call(person, {'action': 'zap', 'slug': 'mark'})
      ['error_code'] == 'bad_arguments')

print('[run]')
r = call(person, {'action': 'run', 'slug': 'mark', 'inputs': {'level': 'L1'}})
check('write recipe needs confirm=true, nothing ran', r['error_code'] == 'confirmation_required'
      and r['nothing_ran'] and not seen, r)
r = call(person, {'action': 'run', 'slug': 'mark', 'inputs': {'level': 'L1'}, 'confirm': 'yes'})
check('confirm must be true, not just truthy', r.get('error_code') == 'confirmation_required'
      and not seen, r)
r = call(person, {'action': 'run', 'slug': 'mark', 'inputs': {'level': 'L1'}, 'confirm': True})
check('every step forwarded as its own call, inputs rendered, caller stamped', r.get('success')
      and seen == [('list_levels', {}, 'claude-desktop'),
                   ('set_parameter', {'element_id': 1, 'parameter_name': 'Mark', 'value': 'L1'},
                    'claude-desktop')], (r, seen))
del seen[:]
r = call(person, {'action': 'run', 'slug': 'typo', 'confirm': True})
check('a step with a bad argument is refused before Revit, recipe stops',
      'stopped the recipe' in r.get('error', '') and 'parameter_name' in r['error']
      and not seen, r)
r = call(person, {'action': 'run', 'slug': 'loop'})
check('a step cannot call t3lab_recipe', 'cannot call t3lab_recipe' in json.dumps(r)
      and not seen, r)

# Harness: an exclusive session; antigravity has no task there.
docs[:] = [{'title': 'PRJ.rvt', 'is_active': True}]
B._handle({'jsonrpc': '2.0', 'id': 1, 'method': 'tools/call', 'params': {
    'name': 't3lab_harness', 'arguments': {'action': 'open', 'goal': 'g',
                                           'revit_document': 'PRJ.rvt', 'exclusive': True}}},
    person, 't')
outsider = {'port': 1, 'client': 'antigravity', 'agent_explicit': True}
r = call(outsider, {'action': 'run', 'slug': 'mark', 'inputs': {'level': 'L1'}, 'confirm': True})
check('harness write check refuses the write step, the read step ran',
      'stopped the recipe' in r.get('error', '') and 'harness' in r['error'].lower()
      and [s[0] for s in seen] == ['list_levels'], (r, seen))
B._handle({'jsonrpc': '2.0', 'id': 1, 'method': 'tools/call', 'params': {
    'name': 't3lab_harness', 'arguments': {'action': 'close'}}}, person, 't')
del seen[:]

# Revit's policy (read-only mode, a "No" in the confirmation) answers with an error.
B._forward = lambda request, current, token, **k: B._tool_response(request.get('id'), (
    {'success': True} if request['params']['name'] == 'list_levels'
    else {'error': 'Read-only mode is on.', 'error_code': 'read_only'}))
r = call(person, {'action': 'run', 'slug': 'mark', 'inputs': {'level': 'L1'}, 'confirm': True})
check('a step refused in Revit stops the recipe', 'Read-only mode' in r.get('error', '')
      and r['steps'][-1] == {'tool': 'set_parameter', 'ok': False}, r)
B._forward = _capture

caller = B._harness_call_tool_for(person, 't')('nightly-bot')
r = caller('t3lab_recipe', {'action': 'run', 'slug': 'mark', 'confirm': True})
check('a launched agent cannot run a recipe (it only reads)', r.get('error_code') == 'tool_denied'
      and not seen, r)

print('[cli]')


def cli(*argv):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        rc = B._cli(list(argv))
    return rc, buf.getvalue()


rc, out = cli('call', 't3lab_recipe', '--action=list', '--out=compact')
check('CLI lists recipes', rc == 0 and '"mark"' in out, out)
rc, out = cli('call', 't3lab_recipe', '--action=run', '--slug=mark', '--inputs={"level":"L2"}')
check('CLI write recipe without --confirm: exit 1, nothing ran', rc == 1
      and 'confirmation_required' in out and not seen, out)
rc, out = cli('call', 't3lab_recipe', '--action=run', '--slug=mark', '--inputs={"level":"L2"}',
              '--confirm')
check('CLI run: steps stamped "cli"', rc == 0 and [s[2] for s in seen] == ['cli', 'cli']
      and seen[-1][1]['value'] == 'L2', (out, seen))
rc, out = cli('call', 't3lab_recipe', '--action=rnu')
check('CLI enum checked from the schema: exit 2', rc == 2 and 'list|show|run' in out, out)

print('[deploy]')
check('recipes.py deploys next to the bridge as toolhub_recipes.py',
      ('../Intelligence/toolhub/recipes.py', 'toolhub_recipes.py') in D.COMPANIONS)
target = tempfile.mkdtemp(prefix='t3lab_lite_recipes_deploy_')
_path, error, written = D.deploy(target)
check('deploy writes it', error is None and 'toolhub_recipes.py' in written
      and 'tool_cli.py' in written, (error, written))

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
