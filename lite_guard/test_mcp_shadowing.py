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

Run:  python3 lite_guard/test_mcp_shadowing.py   (exit 0 = all pass)
"""
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


class ShadowingTests(unittest.TestCase):
    def test_no_module_import_shadowed_in_a_function(self):
        problems = []
        for rel in FILES:
            if os.path.isfile(os.path.join(REPO, rel)):
                problems.extend(shadowed(rel))
        self.assertEqual(problems, [], '\n' + '\n'.join(problems))


if __name__ == '__main__':
    sys.exit(unittest.main())
