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

def codes(src):
    return [h[1] for h in A.scan_source(src)]


# I6 json.dump(s) must say ensure_ascii=False (or go through core.jsonsafe)
for src in ('import json\ns = json.dumps(d)',
            'import json\ns = json.dumps(d, ensure_ascii=True)',
            'import json as _j\n_j.dump(d, fh, indent=2)',
            'from json import dumps\ns = dumps(d)',
            'import json\ns = json.dumps(d, **kw)'):
    check('I6 caught: {!r}'.format(src.split('\n')[-1]), codes(src) == ['I6'], codes(src))
for src in ('import json\ns = json.dumps(d, ensure_ascii=False)',
            'from core import jsonsafe\ns = jsonsafe.dumps(d)',
            'import json\nd = json.loads(s)'):
    check('I6 fine: {!r}'.format(src.split('\n')[-1]), not codes(src), codes(src))

# I7 except OSError around open(): IOError slips through on 2.7
for src in ('try:\n    h = open(p)\nexcept OSError:\n    pass',
            'import io\ntry:\n    h = io.open(p)\nexcept (OSError, ValueError):\n    pass'):
    check('I7 caught', codes(src) == ['I7'], (src, codes(src)))
for src in ('try:\n    h = open(p)\nexcept EnvironmentError:\n    pass',
            'try:\n    h = open(p)\nexcept (IOError, OSError):\n    pass',
            'import os\ntry:\n    os.remove(p)\nexcept OSError:\n    pass'):
    check('I7 fine', not codes(src), (src, codes(src)))

# I8 writing utf-8-sig / utf-16: BOM before every write under IronPython
for src in ('import io\nh = io.open(p, "w", encoding="utf-8-sig")',
            'import io\nh = io.open(p, mode="a", encoding="UTF-16")',
            'import codecs\nh = codecs.open(p, "w", "utf-8-sig")'):
    check('I8 caught', codes(src) == ['I8'], (src, codes(src)))
for src in ('import io\nh = io.open(p, "r", encoding="utf-8-sig")',
            'import io\nh = io.open(p, encoding="utf-8-sig")',
            'import io\nh = io.open(p, "w", encoding="utf-8")'):
    check('I8 fine', not codes(src), (src, codes(src)))

# I9 list comprehension leaks over an `except .. as e` name on 2.7
leak = ('try:\n    f()\nexcept Exception as e:\n'
        '    if x not in [e[0] for e in errs]:\n        errs.append(str(e))')
check('I9 caught', codes(leak) == ['I9'], codes(leak))
for src in ('try:\n    f()\nexcept Exception as e:\n'
            '    if x not in [err[0] for err in errs]:\n        errs.append(str(e))',
            'try:\n    f()\nexcept Exception as e:\n    log(str(e))\n    n = [e for e in errs]',
            'try:\n    f()\nexcept Exception as e:\n    s = set(e for e in errs)\n    log(e)'):
    check('I9 fine', not codes(src), (src, codes(src)))

# I10 \u escape in a plain str literal without unicode_literals
check('I10 caught', codes("t = s.lstrip('\\ufeff')") == ['I10'])
for src in ("t = s.lstrip(u'\\ufeff')", "t = r'\\ufeff'", "t = '\\\\ufeff'",
            "from __future__ import unicode_literals\nt = s.lstrip('\\ufeff')"):
    check('I10 fine: {}'.format(src.split('\n')[-1]), not codes(src), codes(src))

check('relative import followed (smart_purge scanners)',
      'Services/ModelAuditor/smart_purge/purge_scanner_elements.py' in A.reachable())
check('relative package resolves', A._package('a/b/c.py', 1) == 'a.b'
      and A._package('a/b/__init__.py', 1) == 'a.b' and A._package('a/b/c.py', 2) == 'a')
check('top-level lib modules followed', {'ribbon_memory.py', 'open_timing.py'} <= A.reachable())
check('io.open is fine',not A.scan_source('import io\nh = io.open(p, "w", encoding="utf-8")'))
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
