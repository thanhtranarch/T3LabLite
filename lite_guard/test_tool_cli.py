# -*- coding: utf-8 -*-
"""core/tool_cli.py + Lite's bridge CLI (tools / help / call flags / instances /
verify / skill), the argument contract on the MCP path and in the Assistant loop,
and the CLI taking the same gate as an MCP tools/call (argument check, harness
write check, agent stamp). Ported from t3lab_dev dev/test_tool_cli.py; Lite has
no tiers / profiles / t3lab_call.

Run:  python lite_guard/test_tool_cli.py      (no Revit: the manifest comes from server.py)
"""

import contextlib
import io
import json
import os
import subprocess
import sys
import tempfile
import time

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
BRIDGE = os.path.join(LIB, 'core', 'bridge.py')
sys.path.insert(0, LIB)
sys.path.insert(0, os.path.join(REPO, 'lite_guard'))
APPDATA = tempfile.mkdtemp(prefix='t3lab_lite_toolcli_')
os.environ['APPDATA'] = APPDATA
os.environ['T3LAB_HARNESS_DIR'] = tempfile.mkdtemp(prefix='t3lab_lite_toolcli_harness_')
os.environ.pop('T3LAB_AGENT', None)

import core.bridge as B          # noqa: E402
import core.server as S          # noqa: E402
from core import tool_cli as T   # noqa: E402
import server_registry           # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


SRV = object.__new__(S.T3LabAIServer)
SRV._tools = server_registry.registry()
MANIFEST = SRV._handle_tools_list()['tools']
B._save_cached_tools(MANIFEST)
BY_NAME = dict((t['name'], t) for t in MANIFEST)
SET_PARAM = BY_NAME['set_parameter']['inputSchema']

SCHEMA = {'type': 'object', 'required': ['level_name'], 'properties': {
    'level_name': {'type': 'string'},
    'elevation': {'type': 'number'},
    'count': {'type': 'integer'},
    'ids': {'type': 'array', 'items': {'type': 'integer'}},
    'scope': {'type': 'string', 'enum': ['instance', 'type']},
    'dry_run': {'type': 'boolean'},
    'options': {'type': 'object'},
    'format': {'type': 'string'}}}

# ── schema + contract ───────────────────────────────────────────────────────
print('[contract]')
for shape, tools in (('mcp', [{'name': 'x', 'inputSchema': SCHEMA}]),
                     ('anthropic', [{'name': 'x', 'input_schema': SCHEMA}]),
                     ('openai', [{'type': 'function',
                                  'function': {'name': 'x', 'parameters': SCHEMA}}])):
    check('schema_of reads ' + shape, T.schema_of(tools, 'x') is SCHEMA)
check('schema_of unknown -> None', T.schema_of([{'name': 'y'}], 'x') is None)
p = T.arg_problems(SCHEMA, {'elevation': 3})
check('missing required reported', p == [{'kind': 'missing', 'param': 'level_name'}], p)
p = T.arg_problems(SCHEMA, {'level_name': 'L1', 'elevaton': 3})
check('near miss -> did you mean', p == [{'kind': 'unknown', 'param': 'elevaton',
                                         'did_you_mean': 'elevation'}], p)
check('unrelated extra arg passes (non-strict)',
      T.arg_problems(SCHEMA, {'level_name': 'L1', 'zzz': 1}) == [])
check('strict reports every unknown',
      T.arg_problems(SCHEMA, {'level_name': 'L1', 'zzz': 1}, strict=True)
      == [{'kind': 'unknown', 'param': 'zzz'}])
check('_meta ignored', T.arg_problems(SCHEMA, {'level_name': 'L1', '_meta': {}},
                                      strict=True) == [])
check('additionalProperties: no unknown check', T.arg_problems(
    dict(SCHEMA, additionalProperties=True), {'level_name': 'L', 'levl_nme': 1}) == [])
check('empty string counts as missing', T.arg_problems(SCHEMA, {'level_name': ''})
      == [{'kind': 'missing', 'param': 'level_name'}])
txt = T.problems_text('make_level', T.arg_problems(SCHEMA, {'elevaton': 1}))
check('problem text names both', 'missing required "level_name"' in txt
      and 'did you mean "elevation"' in txt and 'Nothing ran' in txt, txt)
p = T.arg_problems(SET_PARAM, {'element_id': 1, 'parameter_nam': 'Mark', 'value': 'A'})
check('real schema: set_parameter typo caught', {'kind': 'unknown', 'param': 'parameter_nam',
                                                 'did_you_mean': 'parameter_name'} in p, p)

# ── flags ───────────────────────────────────────────────────────────────────
print('[flags]')
args, errs = T.parse_flags(
    ['make_level', '--level_name=Tầng 1', '--elevation=3.5', '--count=4', '--ids=1,2',
     '--ids=3', '--scope=type', '--dry-run', '--options={"a":1}', '--agent=x',
     '--json', '{"q":1}', '--out=tsv', '--pick=a', '--port=48885'], SCHEMA)
check('typed values', args == {'level_name': 'Tầng 1', 'elevation': 3.5, 'count': 4,
                               'ids': [1, 2, 3], 'scope': 'type', 'dry_run': True,
                               'options': {'a': 1}} and not errs, (args, errs))
args, errs = T.parse_flags(['--no-dry_run', '--ids=[4,5]', '--elevation=3'], SCHEMA)
check('--no-flag, JSON array, int-looking number', args == {
    'dry_run': False, 'ids': [4, 5], 'elevation': 3} and not errs, (args, errs))
args, errs = T.parse_flags(['--count=x', '--scope=both', '--dry_run=maybe'], SCHEMA)
check('type and enum errors, nothing kept', args == {} and len(errs) == 3, (args, errs))
args, _e = T.parse_flags(['--leveel_name=L'], SCHEMA)
check('unknown flag kept for the strict check', args == {'leveel_name': 'L'})
args, _e = T.parse_flags(['--format=pdf'], SCHEMA)
check('--format is a tool flag (not global)', args == {'format': 'pdf'}, args)

# ── discovery ───────────────────────────────────────────────────────────────
print('[discovery]')
res = T.search(MANIFEST, 'sheet')
check('search ranks names first', res['tools'] and 'sheet' in res['tools'][0].split(' ')[0],
      res['tools'][:3])
check('search lines are one sentence', all(len(x) < 160 for x in res['tools']))
h = T.help_text(BY_NAME['set_parameter'])
check('help: usage + required params', h.startswith('usage: bridge.py call set_parameter')
      and '--parameter_name=<string>' in h and 'required' in h, h)
full = len(json.dumps(MANIFEST))
one = len(T.help_text(BY_NAME['set_parameter'])) + len(json.dumps(T.search(MANIFEST, 'parameter')))
print('        manifest {} chars vs search + help {} chars ({:.1%})'.format(full, one, one / full))
check('discovery costs < 5% of the full manifest', one < full * 0.05)

# ── output ──────────────────────────────────────────────────────────────────
print('[output]')
sheets = {'count': 200, 'sheets': [
    {'id': 100000 + i, 'number': 'A-{:03d}'.format(i), 'name': 'Floor Plan Level {}'.format(i),
     'revision': 'P0{}'.format(i % 3), 'issued': i % 2 == 0} for i in range(200)]}
pretty, tsv = T.render(sheets), T.render(sheets, 'tsv')
check('tsv under half of pretty', len(tsv) < len(pretty) * 0.5)
check('tsv: header once + row per item', tsv.splitlines()[0] == '# count: 200'
      and len(tsv.splitlines()) == 202, tsv[:200])
check('compact round-trips', json.loads(T.render(sheets, 'compact')) == sheets)
check('pick keeps listed columns only', T.pick(sheets, ['number'])['sheets'][0]
      == {'number': 'A-000'})

# ── verify ──────────────────────────────────────────────────────────────────
print('[verify]')
out_dir = tempfile.mkdtemp(prefix='t3lab_lite_exports_')


def _w(name, data, age_min=0):
    path = os.path.join(out_dir, name)
    with open(path, 'wb') as f:
        f.write(data)
    if age_min:
        old = time.time() - age_min * 60
        os.utime(path, (old, old))
    return path


good_pdf = _w('A-101.pdf', b'%PDF-1.7\n...')
_w('A-102.pdf', b'<html>error page</html>')
_w('A-101.dwg', b'')
good_dwg = _w('A-102.dwg', b'AC1032' + b'\x00' * 20)
_w('old.pdf', b'%PDF-1.4', age_min=120)
rep = T.verify([out_dir], exts='pdf', since_min=30, expect=2)
check('folder + ext + since: 2 recent pdfs, one bad', rep['checked'] == 2
      and rep['failed'] == 1 and not rep['ok'], rep)
check('all good -> ok', T.verify([good_pdf, good_dwg], expect=2)['ok'])

# ── MCP path: refused before Revit ──────────────────────────────────────────
print('[mcp path]')
seen = []
real_forward = B._forward


def _capture(request, current, token, **k):
    seen.append(request)
    return B._tool_response(request.get('id'), {'success': True})


B._forward = _capture
cur = {'port': 1, 'client': 'claude-code', 'agent_explicit': True}


def _call(name, args):
    return {'jsonrpc': '2.0', 'id': 1, 'method': 'tools/call',
            'params': {'name': name, 'arguments': args}}


pl = B._tool_payload(B._handle(_call('set_parameter', {'element_id': 1, 'parameter_nam': 'Mark',
                                                       'value': 'A'}), cur, 't'))
check('typo refused, not forwarded', not seen and pl['error_code'] == 'bad_arguments'
      and 'parameter_name' in pl['error'] and pl['nothing_ran'], pl)
B._handle(_call('set_parameter', {'element_id': 1, 'parameter_name': 'Mark', 'value': 'A'}),
          cur, 't')
check('valid call forwarded', len(seen) == 1)

# ── Assistant loop: same contract ───────────────────────────────────────────
print('[assistant]')
from Intelligence import agent_loop as AL   # noqa: E402
tools = [{'name': 'set_parameter', 'input_schema': SET_PARAM}]
ref = AL._arg_refusal(tools, 'set_parameter', {'element_id': 1, 'value': 'A'})
check('assistant refusal carries the schema', ref and ref['nothing_ran'], ref)
check('valid args: no refusal', AL._arg_refusal(
    tools, 'set_parameter', {'element_id': 1, 'parameter_name': 'M', 'value': 'A'}) is None)

# ── CLI in process: the same gate as an MCP tools/call ──────────────────────
print('[cli gate]')


def cli(*argv):
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        rc = B._cli(list(argv))
    return rc, buf.getvalue()


del seen[:]
rc, out = cli('call', 'set_parameter', '--element_id=1', '--parameter_name=Mark', '--value=A')
meta = seen[-1]['params'].get('_meta') or {} if seen else {}
check('valid call: forwarded once, exit 0', rc == 0 and len(seen) == 1, (rc, out))
check('no --agent: stamped "cli" (person-facing, CLI on Agent Cursors)',
      meta.get(B.AGENT_META) == 'cli' and 'cli' in B.HUMAN_AGENTS, meta)
rc, out = cli('call', 'set_parameter', '--element_id=1', '--parameter_name=M', '--value=A',
              '--agent=claude-code')
check('--agent names the caller', seen[-1]['params']['_meta'][B.AGENT_META] == 'claude-code')
os.environ['T3LAB_AGENT'] = 'codex'
cli('call', 'set_parameter', '--element_id=1', '--parameter_name=M', '--value=A')
os.environ.pop('T3LAB_AGENT')
check('T3LAB_AGENT names the caller', seen[-1]['params']['_meta'][B.AGENT_META] == 'codex')

n = len(seen)
rc, out = cli('call', 'set_parameter', '--element_id=1', '--value=A', '--parameter_nam=Mark')
check('typo flag refused, exit 2, nothing forwarded', rc == 2 and len(seen) == n
      and 'did you mean' in out, out)
rc, out = cli('call', 'set_parameter', '--json', '{"element_id": 1}')
check('missing required in --json refused, exit 2, nothing forwarded', rc == 2
      and len(seen) == n and 'missing required' in out, out)
# Without the CLI's own flag check the MCP gate (_dispatch -> _arg_error) still refuses.
real_flags = B._cli_flags
B._cli_flags = lambda *a: None
rc, out = cli('call', 'set_parameter', '--json',
              '{"element_id": 1, "parameter_nam": "Mark", "value": "A"}')
B._cli_flags = real_flags
check('MCP argument check refuses a CLI call: exit 2, nothing forwarded', rc == 2
      and len(seen) == n and json.loads(out)['error_code'] == 'bad_arguments', out)
rc, out = cli('call', 'list_levels', '--json', 'not json')
check('bad --json -> exit 2', rc == 2 and 'bad_arguments' in out, out)
rc, out = cli('call')
check('call without a tool -> usage, exit 2', rc == 2 and 'usage' in out)

# Harness: an exclusive session binds every other agent, the CLI included.
docs = [{'title': 'PRJ.rvt', 'is_active': True}]
real_docs = B._documents_on
B._documents_on = lambda port, token: list(docs)
lead = {'port': 1, 'client': 'claude-desktop', 'agent_explicit': True}
B._handle(_call('t3lab_harness', {'action': 'open', 'goal': 'g', 'revit_document': 'PRJ.rvt',
                                  'exclusive': True}), lead, 't')
n = len(seen)
rc, out = cli('call', 'set_parameter', '--element_id=1', '--parameter_name=M', '--value=A',
              '--agent=antigravity')
pl = json.loads(out)
check('harness write check refuses a CLI call: exit 1, nothing forwarded', rc == 1
      and len(seen) == n and pl['error_code'] == 'harness_managed' and pl['nothing_ran'], out)
rc, out = cli('call', 'purge_unused', '--dry_run', '--agent=antigravity')
check('a dry run passes the harness check', rc == 0 and len(seen) == n + 1, out)
B._handle(_call('t3lab_harness', {'action': 'close'}), lead, 't')
B._documents_on = real_docs

# Exit 1 for a tool / Revit error; the answer comes from _answer, as on stdio.
B._forward = lambda request, current, token, **k: B._tool_response(
    request.get('id'), {'error': 'Read-only mode is on.', 'error_code': 'read_only'})
rc, out = cli('call', 'set_parameter', '--element_id=1', '--parameter_name=M', '--value=A')
check('Revit refusal -> exit 1', rc == 1 and 'read_only' in out, out)
B._forward = lambda request, current, token, **k: {'jsonrpc': '2.0', 'id': 1, 'result': {
    'isError': True, 'content': [{'type': 'text', 'text': 'boom'}]}}
rc, out = cli('call', 'list_levels')
check('isError -> exit 1', rc == 1, out)
answered = []
real_answer = B._answer
B._answer = lambda *a: answered.append(a[0]['params']['name']) or real_answer(*a)
B._forward = _capture
rc, out = cli('instances')
B._answer = real_answer
check('instances goes through _answer', answered == ['list_revit_instances'], answered)
B._forward = real_forward

# ── CLI as a process ────────────────────────────────────────────────────────
print('[cli]')


def _proc(*args, **env):
    proc = subprocess.run([sys.executable, '-I', BRIDGE] + list(args), capture_output=True,
                          text=True, encoding='utf-8', timeout=60,
                          env=dict(os.environ, APPDATA=APPDATA, **env))
    return proc.returncode, proc.stdout


rc, out = _proc('tools', 'level')
check('tools <words> from the cache (Revit closed)', rc == 0 and 'level' in out, out[:300])
rc, out = _proc('tools', 'level', '--out=compact')
check('tools --out=compact is one line', rc == 0 and len(out.strip().splitlines()) == 1)
rc, out = _proc('tools', 'verify')
check('bridge tools are listed too', rc == 0 and 't3lab_verify' in out, out)
rc, out = _proc('help', 'set_parameter')
check('help <tool>', rc == 0 and '--parameter_name=<string>' in out, out)
rc, out = _proc('help', 'set_paramter')
check('help unknown tool -> candidates, exit 2', rc == 2 and 'set_parameter' in out, out)
rc, out = _proc('call', 'set_parameter', '--element_id=1', '--parameter_name=M', '--value=A',
                '--colour=red')
check('unknown flag refused (strict)', rc == 2 and 'unknown \\"colour\\"' in out, out)
rc, out = _proc('call', 'set_parameter', '--element_id=1', '--parameter_name=M', '--value=A')
check('valid flags reach the bridge (Revit down) -> exit 1', rc == 1 and
      json.loads(out).get('error_code') == 'revit_unavailable', out)
rc, out = _proc('verify', out_dir, '--ext=pdf', '--since=30', '--out=compact')
check('verify CLI: bad pdf -> exit 1', rc == 1 and json.loads(out)['failed'] == 1, out)
rc, out = _proc('verify', good_pdf, good_dwg, '--expect=2')
check('verify CLI: all good -> exit 0', rc == 0 and json.loads(out)['ok'], out)
rc, out = _proc('verify')
check('verify without paths -> usage, exit 2', rc == 2 and 'usage' in out)
rc, out = _proc('verify', out_dir, '--expect=x')
check('verify with a bad number -> exit 2', rc == 2, out)

# ── skill: tool_cli's text, Lite's lines ────────────────────────────────────
print('[skill]')
check('every Lite replacement still matches tool_cli.SKILL_MD',
      all(old in T.SKILL_MD for old, _new in B.SKILL_LITE),
      [old[:40] for old, _new in B.SKILL_LITE if old not in T.SKILL_MD])
skills = tempfile.mkdtemp(prefix='t3lab_lite_skills_')
rc, out = _proc('skill', '--install=' + skills)
path = os.path.join(skills, 't3lab-cli', 'SKILL.md')
check('skill --install writes SKILL.md', rc == 0 and os.path.isfile(path), out)
body = io.open(path, encoding='utf-8').read() if os.path.isfile(path) else ''
check('skill has AgentSkills frontmatter', body.startswith('---\nname: t3lab-cli\n'))
for word in ('tier', 't3lab_result', 'capabilities', 'revit_document'):
    check('skill names no dev-only feature: ' + word, word not in body)
check('skill points at switch_active_document', 'switch_active_document' in body)
rc, printed = _proc('skill')
check('skill prints the same text', rc == 0 and printed.rstrip('\n') == body.rstrip('\n'))

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
