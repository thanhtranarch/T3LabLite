# -*- coding: utf-8 -*-
"""The CI workflow file is one GitHub can load, and it runs everything.

A workflow GitHub cannot parse fails in 0 s without running a step and shows
no check on the pull request, so nothing else would notice (2026-11-17).
No PyYAML here: the checks are textual, on the simple layout this file has.

Run: python lite_guard/test_workflow.py
"""
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
WORKFLOWS = os.path.join(REPO, '.github', 'workflows')

FAILED = []


def check(label, ok, detail=''):
    print('  {} {}{}'.format('ok  ' if ok else 'FAIL', label,
                             ('\n       ' + detail) if detail and not ok else ''))
    if not ok:
        FAILED.append(label)


def duplicate_keys(text):
    """[(line, key)] for a key written twice in one block mapping."""
    out, stack = [], []            # stack of (indent, set of keys)
    for number, raw in enumerate(text.split('\n'), 1):
        line = raw.split(' #', 1)[0].rstrip() if not raw.lstrip().startswith('#') else ''
        if not line.strip():
            continue
        indent = len(line) - len(line.lstrip())
        body = line.lstrip()
        item = body.startswith('- ')
        if item:                   # a list item opens a new mapping two columns in
            while stack and stack[-1][0] >= indent:
                stack.pop()
            indent, body = indent + 2, body[2:].lstrip()
        match = re.match(r'([A-Za-z_][\w.-]*)\s*:(\s|$)', body)
        if not match:
            continue
        while stack and stack[-1][0] > indent:
            stack.pop()
        if not stack or stack[-1][0] < indent or item:
            stack.append((indent, set()))
        keys = stack[-1][1]
        if match.group(1) in keys:
            out.append((number, match.group(1)))
        keys.add(match.group(1))
    return out


def main():
    print('the checker sees a duplicated key')
    bad = "steps:\n  - name: a\n    run: one\n    run: two\n  - name: b\n    run: three\n"
    check('two `run:` in one step are found', duplicate_keys(bad) == [(4, 'run')])
    good = "on:\n  push:\n    branches: [main]\njobs:\n  a:\n    steps:\n      - uses: x\n      - name: n\n        run: r\n      - name: m\n        run: r\n"
    check('one key per step is clean', duplicate_keys(good) == [])

    print('workflow files')
    names = sorted(n for n in os.listdir(WORKFLOWS) if n.endswith(('.yml', '.yaml')))
    check('there is a workflow', bool(names))
    for name in names:
        with io.open(os.path.join(WORKFLOWS, name), encoding='utf-8') as f:
            text = f.read()
        found = duplicate_keys(text)
        check('{}: no key twice in one mapping'.format(name), not found,
              ', '.join('line {} `{}`'.format(n, k) for n, k in found))
        check('{}: no tab characters'.format(name), '\t' not in text)
    with io.open(os.path.join(WORKFLOWS, 'lite-guard.yml'), encoding='utf-8') as f:
        text = f.read()
    check('lite-guard.yml runs lite_guard/run_all.py (every test, by name)',
          'python3 lite_guard/run_all.py' in text)
    check('it runs on pull requests and on main', 'pull_request' in text and 'push' in text)

    if FAILED:
        print('\n{} check(s) failed.'.format(len(FAILED)))
        return 1
    print('\nAll workflow checks passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
