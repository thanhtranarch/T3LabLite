# -*- coding: utf-8 -*-
"""
CPython 3 test: no function in the MCP code re-imports or re-assigns a name
the module already imported. No Revit needed.

Python makes such a name local for the WHOLE function, so every use that runs
before the local import -- or in another if/elif branch -- raises
UnboundLocalError. That is how every external tools/call died in
MCPExternalEventHandler.Execute (a local `from core import agent_cursors`
below an `agent_cursors.offer(...)`), and how ai_element_filter silently lost
its `view` field (a local `from Snippets._compat import elem_name` in another
branch of _execute_tool_in_context).

The same trap without a module import: a name a function imports locally
but reads on an earlier line (2026-11-19: T3Lab's Execute did, so none of its
MCP calls finished). Checked over all of lib/.

Run:  python3 lite_guard/test_mcp_shadowing.py   (exit 0 = all pass)
"""
import ast
import os
import symtable
import sys
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

FILES = (
    'lib/core/server.py',
    'lib/core/bridge.py',
    'lib/core/mcp_policy.py',
    'lib/core/mcp_wiki.py',
    'lib/core/mcp_prompts.py',
    'lib/core/mcp_ipy_host.py',
    'lib/core/call_summary.py',
    'lib/core/agent_cursors.py',
    'lib/core/paths.py',
    'lib/bridge_deploy.py',
    'lib/Services/mcp_service.py',
    'lib/GUI/MCPControlDialog.py',
)

# Reviewed: the local import is the first statement that touches the name
# and every use follows it in the same block.
ALLOWED = {
    ('lib/core/server.py', '_ask_user_to_confirm', 'mcp_policy'),
    ('lib/core/server.py', '_ask_on_bar', 'mcp_policy'),
    ('lib/core/server.py', 'start_server', 'time'),
}


def shadowed(rel):
    path = os.path.join(REPO, rel)
    with open(path, encoding='utf-8') as f:
        src = f.read()
    top = symtable.symtable(src, rel, 'exec')
    imported = {s.get_name() for s in top.get_symbols() if s.is_imported()}
    out = []

    def walk(table):
        for child in table.get_children():
            if child.get_type() == 'function':
                for sym in child.get_symbols():
                    name = sym.get_name()
                    if (name in imported and sym.is_local()
                            and not sym.is_parameter()
                            and (rel, child.get_name(), name) not in ALLOWED):
                        out.append('{}:{} {}() makes {!r} local'.format(
                            rel, child.get_lineno(), child.get_name(), name))
            walk(child)

    walk(top)
    return out


def read_before_local_import(rel, src=None):
    """Functions that read a name on a line above their own import of it."""
    if src is None:
        with open(os.path.join(REPO, rel), encoding='utf-8') as f:
            src = f.read()
    tree = ast.parse(src, rel)
    out = []
    scopes = (ast.FunctionDef, ast.AsyncFunctionDef, ast.Lambda, ast.ClassDef)

    def scan(fn):
        imported, bound, read = {}, {}, {}

        def visit(node):
            for child in ast.iter_child_nodes(node):
                if isinstance(child, scopes):
                    continue                    # its own scope
                if isinstance(child, (ast.Import, ast.ImportFrom)):
                    for alias in child.names:
                        name = (alias.asname or alias.name).split('.')[0]
                        imported.setdefault(name, child.lineno)
                        bound[name] = min(bound.get(name, child.lineno), child.lineno)
                elif isinstance(child, ast.Name):
                    lines = read if isinstance(child.ctx, ast.Load) else bound
                    lines[child.id] = min(lines.get(child.id, child.lineno), child.lineno)
                visit(child)

        visit(fn)
        params = set(a.arg for a in fn.args.args + fn.args.kwonlyargs)
        for name in imported:
            if name not in params and name in read and read[name] < bound[name]:
                out.append('{}:{} {}() reads {!r} on line {}, imports it on line {}'.format(
                    rel, fn.lineno, fn.name, name, read[name], bound[name]))

    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            scan(node)
    return out


def lib_files():
    root = os.path.join(REPO, 'lib')
    for folder, _dirs, names in os.walk(root):
        for name in sorted(names):
            if name.endswith('.py'):
                yield os.path.relpath(os.path.join(folder, name), REPO).replace(os.sep, '/')


class ShadowingTests(unittest.TestCase):
    def test_no_name_read_before_its_local_import(self):
        problems = []
        for rel in lib_files():
            try:
                problems.extend(read_before_local_import(rel))
            except SyntaxError:
                pass            # not this gate's job
        self.assertEqual(problems, [], '\n' + '\n'.join(problems))

    def test_the_rule_sees_a_read_above_the_import(self):
        src = ('def Execute(self, app):\n'
               '    held = agent_cursors.paused_refusal()\n'
               '    try:\n'
               '        from core import agent_cursors\n'
               '    except Exception:\n'
               '        pass\n')
        found = read_before_local_import('probe.py', src)
        self.assertEqual(len(found), 1, found)
        self.assertIn("reads 'agent_cursors' on line 2, imports it on line 4", found[0])

    def test_no_module_import_shadowed_in_a_function(self):
        problems = []
        for rel in FILES:
            if os.path.isfile(os.path.join(REPO, rel)):
                problems.extend(shadowed(rel))
        self.assertEqual(problems, [], '\n' + '\n'.join(problems))


if __name__ == '__main__':
    sys.exit(unittest.main())
