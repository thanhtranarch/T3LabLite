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

# T3Lab and T3Lab Lite share one CPython engine and both ship these modules: a
# fixed __namespace__ raises "Duplicate type name" after the other one's click.
from Snippets._compat import clr_namespace   # noqa: E402
check('clr_namespace is unique per load',
      clr_namespace('X').startswith('T3Lab.X_') and clr_namespace('X') != clr_namespace('X'))
for rel in ('Snippets/finish_floors.py', 'Snippets/wall_tools.py', 'Snippets/_links.py',
            'Snippets/_auto_dimension.py', 'Services/point_cloud_build.py'):
    lines = [l.strip() for l in open(os.path.join(LIB, *rel.split('/')), encoding='utf-8')
             if '__namespace__' in l and '=' in l]
    check('{}: .NET class namespace is per load'.format(rel),
          lines and all('clr_namespace(' in l for l in lines), lines)


# IronPython 2.7 runtime traps fixed in the shared modules (same checks as t3lab_dev).
def _src(rel):
    return open(os.path.join(LIB, *rel.split('/')), encoding='utf-8').read()


check('smart_purge reads type names through elem_name (default types stay protected)',
      "hasattr(wall_type, 'Name')" not in _src(
          'Services/ModelAuditor/smart_purge/purge_scanner_elements.py'))
check('CSV report: BOM written once', "encoding='utf-8-sig'" not in _src(
    'Services/batchout_report.py'))
check('profile scan catches IOError on 2.7', 'except (EnvironmentError, ValueError)' in _src(
    'Services/batchout_profiles.py'))
check('JSON with accented text (IronPython)', 'ensure_ascii=False' in _src('Snippets/_cad_plan.py')
      and 'jsonsafe.dumps' in _src('Snippets/xlsx_table.py')
      and 'jsonsafe.dumps(info)' in _src('Snippets/excel_table_revit.py'))


# Headless tools: a Revit error rolls back instead of a modal dialog holding Revit.
from core import revit_failures as RF       # noqa: E402


class _Msg(object):
    def __init__(self, sev, text):
        self.sev, self.text = sev, text

    def GetSeverity(self):
        return self.sev

    def GetDescriptionText(self):
        return self.text


class _Opts(object):
    def SetClearAfterRollback(self, on):
        self.clear = on


class _FA(object):
    def __init__(self, msgs):
        self.msgs, self.deleted, self.opts = msgs, [], _Opts()

    def GetFailureMessages(self):
        return self.msgs

    def DeleteWarning(self, m):
        self.deleted.append(m.text)

    def GetFailureHandlingOptions(self):
        return self.opts

    def SetFailureHandlingOptions(self, o):
        pass


class _Args(object):
    def __init__(self, fa):
        self.fa, self.result = fa, None

    def GetFailuresAccessor(self):
        return self.fa

    def SetProcessingResult(self, r):
        self.result = r


fs = RF.FailureSession(rollback_result='ROLLBACK')
fa = _FA([_Msg('Warning', 'Walls overlap'), _Msg('Error', 'Cannot join')])
args = _Args(fa)
fs.on_failures(None, args)
res = RF.annotate({'success': True}, fs)
check('failure guard: warning dismissed, error rolled back without a dialog',
      fa.deleted == ['Walls overlap'] and args.result == 'ROLLBACK' and fa.opts.clear
      and res['success'] is False and res['revit_failures']['errors_rolled_back'] == ['Cannot join'],
      res)
check('handler wraps spec write tools in the guard',
      'is_write and tool_specs.get(task.tool_name) is not None' in src
      and 'revit_failures.guard(app)' in src)

# No fixed .NET namespace left in lib/ (core/server.py runs only under IronPython).
import re as _re                           # noqa: E402
_pat = _re.compile(r"""__namespace__ *= *["']T3Lab\.[A-Za-z0-9_.]+["'] *($|#)""")
fixed = []
for root, _d, files in os.walk(LIB):
    for name in files:
        if name.endswith('.py'):
            path = os.path.join(root, name)
            rel = os.path.relpath(path, LIB)
            if rel == os.path.join('core', 'server.py'):
                continue
            for n, line in enumerate(open(path, encoding='utf-8'), 1):
                if _pat.search(line):
                    fixed.append('{}:{}'.format(rel, n))
check('every lib .NET class namespace is per load (clr_namespace)', not fixed, fixed)

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
