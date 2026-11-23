# -*- coding: utf-8 -*-
"""dev/audit_ipy_compat.py: catches Python-3-only code the IronPython MCP server can load.

Run:  python lite_guard/test_audit_ipy_compat.py
"""
import os
import sys

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'lite_guard'))
sys.path.insert(0, os.path.join(REPO, 'lib'))
import audit_ipy_compat as A   # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


CASES = {
    'I1': 'x = f"{a}"',
    'I2': 'k = name.casefold()',
    'I3': 'h = open(p, "w", encoding="utf-8")',
    'I4': 'def f(a, *, b):\n    pass',
    'I5': 'import pathlib',
}
for code, src in CASES.items():
    hits = A.scan_source(src)
    check('{} caught'.format(code), [h[1] for h in hits] == [code], hits)
for src, what in (('def f():\n    nonlocal x', 'nonlocal'),
                  ('def f(a: int):\n    pass', 'annotation'),
                  ('class A(B):\n    def f(self):\n        super().f()', 'zero-arg super'),
                  ('raise ValueError() from None', 'raise from'),
                  ('import os\nos.replace(a, b)', 'os.replace'),
                  ('import os\nos.makedirs(p, exist_ok=True)', 'exist_ok'),
                  ('print("a", end="")', 'print kw without __future__')):
    check('{} caught'.format(what), A.scan_source(src), src)
check('io.open is fine', not A.scan_source('import io\nh = io.open(p, "w", encoding="utf-8")'))
check('print kw fine with print_function',
      not A.scan_source('from __future__ import print_function\nprint("a", end="")'))
check('# ipy-ok: exempts a guarded line',
      not A.scan_source('k = name.casefold()  # ipy-ok: guarded'))
reach = A.reachable()
check('server + spec modules are in scope',
      {'core/server.py', 'core/tool_specs.py', 'core/tools/export.py'} <= reach)
check('lazy imports inside functions followed', 'Services/sheet_export.py' in reach
      and 'Snippets/_rebar_wizard.py' in reach)
check('the repo is clean', A.run(quiet=True) == 0)

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
