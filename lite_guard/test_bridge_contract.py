# -*- coding: utf-8 -*-
"""Lite bridge + server hooks ported from t3lab_dev: the argument check before
Revit, t3lab_verify, spec tools in the registry, Lite's policy on them, the
read-back hook, and the two Lite differences (requires, family_transfer API).

Run:  python lite_guard/test_bridge_contract.py      (no Revit)
"""
import json
import os
import sys
import tempfile
import types

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
sys.path.insert(0, os.path.join(REPO, 'lite_guard'))
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_lite_contract_')

import core.bridge as B                    # noqa: E402
import core.server as S                    # noqa: E402
from core import mcp_policy as P           # noqa: E402
from core import tool_specs as TS          # noqa: E402
import server_registry                     # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


REGISTRY = server_registry.registry()
srv = object.__new__(S.T3LabAIServer)
srv._tools = REGISTRY
B._save_cached_tools(list(REGISTRY.values()))

# ── bridge: argument check + t3lab_verify ───────────────────────────────────
print('[bridge]')
seen = []
B._forward = lambda request, current, token, **k: (
    seen.append(request['params']['name']) or B._tool_response(request.get('id'), {'ok': True}))
cur = {'port': B.PORT_MIN}


def call(name, args):
    resp = B._dispatch({'jsonrpc': '2.0', 'id': 1, 'method': 'tools/call',
                        'params': {'name': name, 'arguments': args}}, cur, 't')
    return json.loads(resp['result']['content'][0]['text'])


r = call('set_parameter', {'element_id': 1, 'parameter_nam': 'Mark', 'value': 'A'})
check('typo refused before Revit, names the parameter', not seen
      and r['error_code'] == 'bad_arguments' and 'parameter_name' in r['error']
      and r['nothing_ran'], r)
r = call('set_parameter', {'element_id': 1})
check('missing required refused', not seen and 'missing required' in r['error'], r)
call('set_parameter', {'element_id': 1, 'parameter_name': 'Mark', 'value': 'A'})
check('valid call forwarded', seen == ['set_parameter'], seen)
call('manage_worksets', {'operation': 'rename', 'name': 'A', 'new_name': 'B'})
check('spec tool forwarded', seen[-1] == 'manage_worksets')
out = tempfile.mkdtemp(prefix='t3lab_lite_out_')
good, bad = os.path.join(out, 'A.pdf'), os.path.join(out, 'B.pdf')
open(good, 'wb').write(b'%PDF-1.7')
open(bad, 'wb').write(b'<html>')
n = len(seen)
r = call('t3lab_verify', {'paths': [out], 'ext': 'pdf', 'expect': 2})
check('t3lab_verify answered by the bridge, bad file listed', len(seen) == n
      and r['failed'] == 1 and not r['ok'] and 'B.pdf' in json.dumps(r), r)
check('t3lab_verify is listed', 't3lab_verify' in [t['name'] for t in B.BRIDGE_TOOLS])
tc, B._tc = B._tc, None
check('no tool_cli next to the bridge: calls still go through',
      B._arg_error('set_parameter', {}) is None
      and call('t3lab_verify', {'paths': [out]})['error_code'] == 'verify_unavailable')
B._tc = tc
import bridge_deploy                       # noqa: E402
check('tool_cli.py deploys with the bridge',
      ('tool_cli.py', 'tool_cli.py') in bridge_deploy.COMPANIONS
      and os.path.isfile(os.path.join(LIB, 'core', 'tool_cli.py')))

# ── server: specs in the registry, policy, read-back ────────────────────────
print('[server]')
specs = TS.all_specs()
check('every spec tool is in the registry', all(n in REGISTRY for n in specs))
check('main-thread spec tools are write tools',
      TS.names(main_thread=True) <= S.T3LabAIServer._WRITE_TOOLS)
check('room_to_floor is a spec now, still a batch step',
      'room_to_floor' in specs and 'room_to_floor' in S.T3LabAIServer._BATCH_STEP_TOOLS)
src = open(os.path.join(LIB, 'core', 'server.py'), encoding='utf-8').read()
check('no hand-written room_to_floor left', "elif tool_name == 'room_to_floor':" not in src
      and "'room_to_floor': {" not in src)
body = src[src.index('    def _execute_tool_in_context('):]
check('spec dispatch sits before run_batch and the branch chain',
      body.index('tool_specs.run(tool_name, doc, arguments)')
      < body.index("if tool_name == 'run_batch':") < body.index("if tool_name == 'revit_get_active_view':"))
check('handler reads the model back after a write, tolerates an old server object',
      "verify = getattr(self.server, '_verify_write', None) if is_write else None" in src)
ro = srv._read_only_tool_names()
policy = P.policy_from_settings({})
for name, args, want in (
        ('manage_worksets', {'operation': 'delete', 'names': ['A'], 'reassign_to': 'B'}, 'confirm'),
        ('manage_worksets', {'operation': 'rename'}, 'run'),
        ('smart_purge', {'dry_run': False}, 'confirm'),
        ('smart_purge', {}, 'run'),
        ('cad_imports', {'operation': 'delete', 'element_ids': [1]}, 'confirm'),
        ('transfer_families', {'mode': 'overwrite'}, 'confirm'),
        ('export_sheets', {'format': 'pdf', 'all': True}, 'run'),
        ('inspect_rebar', {}, 'run')):
    decision, _msg = P.check_call(name, policy, ro, srv.is_destructive(name, args))
    check('policy: {} {} -> {}'.format(name, sorted(args.items())[:1], want), decision == want,
          decision)
locked = dict(policy, read_only=True)
blocked = [n for n, s in specs.items() if s.effect != 'read'
           and P.check_call(n, locked, ro, False)[0] != 'block']
check('read-only mode blocks every spec write tool', not blocked, blocked)
check('confirm text is readable for a spec tool',
      'Manage worksets' in '{}'.format(P.confirm_text('manage_worksets', {'operation': 'delete'})))

# ── Lite differences ────────────────────────────────────────────────────────
print('[lite]')
check('transfer_view_templates is not registered (Lite ships no such window)',
      'transfer_view_templates' not in specs and 'transfer_view_templates' in TS.SKIPPED)
check('no spec failed to load', not TS.LOAD_ERRORS, TS.LOAD_ERRORS)
from core.tools import families as HF      # noqa: E402
Result = types.SimpleNamespace
calls = []


class FT(object):                          # Lite's family_transfer: one target per call
    @staticmethod
    def execute_plan(src_doc, dst, actions, overwrite_values=False):
        calls.append((dst, overwrite_values))
        if dst == 'B':
            raise RuntimeError('Revit rolled it back\nstack')
        return ['r']


rows = HF._run_plans(FT, 'src', [('A', ['a']), ('B', ['b'])], True)
check('family transfer without execute_plans: one call per target, error kept as text',
      rows == [('A', ['r'], None), ('B', [], 'Revit rolled it back')]
      and calls == [('A', True), ('B', True)], rows)
FT.execute_plans = staticmethod(lambda s, plans, overwrite_values=False: 'many')
check('execute_plans used when the module has it', HF._run_plans(FT, 's', [], False) == 'many')

# T3Lab and T3Lab Lite share one CPython engine; claim_lib re-imports a lib module
# after the other one's click: a fixed __namespace__ raises "Duplicate type name".
from Snippets._compat import clr_namespace   # noqa: E402
check('clr_namespace is unique per load',
      clr_namespace('X').startswith('T3Lab.X_') and clr_namespace('X') != clr_namespace('X'))
for rel in ('Snippets/finish_floors.py', 'Snippets/wall_tools.py', 'Snippets/_links.py',
            'Snippets/_auto_dimension.py', 'Services/point_cloud_build.py',
            'Snippets/_rebar_wizard.py', 'Snippets/_cad_revit.py', 'Snippets/family_transfer.py',
            'Snippets/_drawing_clone.py', 'Services/revit_context.py', 'Services/join_service.py',
            'GUI/ManaSelectDialog.py', 'GUI/FamiGenDialog.py', 'GUI/TextToElementDialog.py',
            'GUI/ManaLocaDialog.py', 'FamilyGen/builder.py', 'core/agent_cursors_glow.py',
            'core/agent_cursors.py', 'core/file_watcher.py', 'core/server.py'):
    lines = [l.strip() for l in open(os.path.join(LIB, *rel.split('/')), encoding='utf-8')
             if '__namespace__' in l and '=' in l]
    check('{}: .NET class namespace is per load'.format(rel),
          lines and all('clr_namespace(' in l for l in lines), lines)

import glob   # noqa: E402
import re     # noqa: E402
PER_LOAD = ('clr_namespace(', 'uuid.uuid4()', '_NS_SUFFIX')
fixed = []
for path in glob.glob(os.path.join(LIB, '**', '*.py'), recursive=True):
    with open(path, encoding='utf-8', errors='replace') as fh:
        for no, line in enumerate(fh, 1):
            if re.match(r'\s*__namespace__\s*=', line) and not any(m in line for m in PER_LOAD):
                fixed.append('{}:{}'.format(os.path.relpath(path, LIB), no))
check('no fixed __namespace__ left in lib/', not fixed, fixed)

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
