# -*- coding: utf-8 -*-
"""core/tool_specs.py: one declaration per tool reaches every consumer.

Checks every spec in core/tools/<domain>.py is complete, loads without Revit,
and shows up where its fields say (MCP registry, dispatch, _WRITE_TOOLS,
confirmation, tiers, Assistant subsets, read-back, Tool Hub titles, ribbon).

Run:  python lite_guard/test_tool_specs.py
"""
import ast
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
import guard                             # noqa: E402
TABS = [os.path.join(guard.ROOT, guard.TAB_DIR)]
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_specs_test_')

from core import tool_specs as TS        # noqa: E402
import server_registry                    # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


SPECS = TS.all_specs()
TOOLS_DIR = os.path.join(LIB, 'core', 'tools')

# ── loading ─────────────────────────────────────────────────────────────────
print('[loading]')
check('every domain module loads (outside Revit)', not TS.LOAD_ERRORS, TS.LOAD_ERRORS)
check('specs found', len(SPECS) >= 3, sorted(SPECS))
REVIT_ROOTS = ('Autodesk', 'System', 'clr', 'pyrevit', 'RevitServices')
for name in sorted(os.listdir(TOOLS_DIR)):
    if not name.endswith('.py'):
        continue
    tree = ast.parse(open(os.path.join(TOOLS_DIR, name), encoding='utf-8').read())
    top = []
    for node in tree.body:
        if isinstance(node, ast.Import):
            top += [a.name for a in node.names]
        elif isinstance(node, ast.ImportFrom):
            top.append(node.module or '')
    bad = [m for m in top if m.split('.')[0] in REVIT_ROOTS
           or m.startswith('GUI') or m.startswith('Snippets') or m.startswith('Services')]
    check('{}: no Revit / GUI / Snippets import at module level'.format(name), not bad, bad)
    if name.startswith('_'):
        continue
    src = open(os.path.join(TOOLS_DIR, name), encoding='utf-8').read()
    if 'Transaction(' in src:
        reads = [n for n, s in SPECS.items() if s.domain == name[:-3] and s.effect == 'read'
                 and s.main_thread is False]
        check('{}: opens a Transaction, so no spec runs off the main thread'.format(name),
              not reads, reads)

# ── declarations ────────────────────────────────────────────────────────────
print('[declarations]')
for name, spec in sorted(SPECS.items()):
    check('{}: complete'.format(name), not spec.problems(), spec.problems())
literal = server_registry.literal()
check('no spec duplicates a hand-written tool', not set(SPECS) & set(literal),
      sorted(set(SPECS) & set(literal)))
buttons = set()
for tab in TABS:
    for root, dirs, _files in os.walk(tab):
        buttons.update(d for d in dirs if d.endswith('.pushbutton'))
for name, spec in sorted(SPECS.items()):
    for button in spec.ribbon:
        check('{}: ribbon button {} exists'.format(name, button), button in buttons)

# ── consumers ───────────────────────────────────────────────────────────────
print('[consumers]')
import core.server as S                               # noqa: E402
from core import write_verify as W                    # noqa: E402
from Intelligence import tool_schema as TSC           # noqa: E402
from Intelligence.agents import specialists as SP     # noqa: E402
from Intelligence.toolhub import catalog              # noqa: E402

registry = server_registry.registry()
srv = object.__new__(S.T3LabAIServer)
srv._tools = registry
groups = {'read': SP.READ_TOOLS, 'modify': SP.MODIFY_TOOLS, 'export': SP.EXPORT_TOOLS,
          'modeling': SP.MODELING_TOOLS, 'qa': SP.QA_TOOLS, 'multidoc': SP.MULTIDOC_TOOLS,
          'essential': TSC.ESSENTIAL_TOOL_NAMES, 'hidden': TSC.LOCAL_HIDDEN_TOOLS,
          'external': TSC._EXTERNAL_ONLY_TOOLS}
modes = {'ids': W.ID_TOOLS, 'must_change': W.MUST_CHANGE, 'self': W.SELF_CHECKED,
         'files': W.FILE_LIST_TOOLS}
titles = catalog.titles()
for name, spec in sorted(SPECS.items()):
    check('{}: in the MCP registry'.format(name), registry.get(name) == spec.registry_entry())
    check('{}: main thread <-> _WRITE_TOOLS'.format(name),
          (name in S.T3LabAIServer._WRITE_TOOLS) == spec.main_thread)
    # Lite has no tiers: its policy reads the read-only list (core/mcp_policy.py).
    read_only = name in srv._read_only_tool_names()
    check('{}: read-only list <-> effect read'.format(name), read_only == (spec.effect == 'read'))
    check('{}: file output listed'.format(name),
          (name in TSC.FILE_OUTPUT_TOOL_NAMES) == (spec.effect == 'file'))
    check('{}: placed {}'.format(name, spec.placement),
          all(name in groups[p] for p in spec.placement))
    check('{}: read-back {}'.format(name, spec.verify), all(name in modes[m] for m in spec.verify))
    check('{}: Tool Hub title'.format(name), titles.get(name, {}).get('title_vi') == spec.title_vi)
    if isinstance(spec.destructive, tuple):
        field, values = spec.destructive
        check('{}: destructive when {} is {}'.format(name, field, values),
              srv.is_destructive(name, {field: values[0]})
              and not srv.is_destructive(name, {field: values[0], 'dry_run': True}))
    elif not spec.destructive:
        check('{}: not destructive'.format(name), not srv.is_destructive(name, {}))
check('by_ribbon maps buttons to tools',
      all(n in SPECS for names in TS.by_ribbon().values() for n in names))

# ── run ─────────────────────────────────────────────────────────────────────
print('[run]')


class Boom(Exception):
    Message = '.NET says no\nstack'


demo = TS.ToolSpec('demo_tool', 'Demo.', lambda doc, args: (_ for _ in ()).throw(Boom()),
                   properties={'dry_run': {'type': 'boolean'}}, verify=('ids',),
                   title_vi='Thử')
TS.all_specs()['demo_tool'] = demo
check('exception -> error dict, first line, tool name',
      TS.run('demo_tool', object(), {}) == {'error': '.NET says no', 'tool': 'demo_tool'})
demo.run = lambda doc, args: 42
check('non-dict result wrapped', TS.run('demo_tool', object(), {}) == {'result': 42})
check('no document -> friendly error', 'No active document' in TS.run('demo_tool', None, {})['error'])
demo.run = lambda doc, args: args
check('args copied, never the caller\'s dict', TS.run('demo_tool', 1, {'a': 1}) == {'a': 1})
check('unknown tool', TS.run('nope', 1, {})['error'].startswith('Unknown tool'))
del TS.all_specs()['demo_tool']
bad = TS.ToolSpec('x', 'X', None, effect='write', placement='nowhere', verify=('vibes',),
                  destructive=('mode', ('wipe',)), properties={'mode': {'enum': ['keep']}},
                  required=['ghost'])
problems = ' | '.join(bad.problems())
for word in ('placement', 'verify must', 'run is not callable', 'title_vi', 'required',
             'destructive values', 'dry_run'):
    check('problems() catches: {}'.format(word), word in problems, problems)

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
