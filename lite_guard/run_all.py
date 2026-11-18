# -*- coding: utf-8 -*-
"""Run the guard and every lite_guard/test_*.py; exit 1 when any fails.

The CI workflow calls only this file, so a new test is picked up by its name
and nobody edits YAML per test (2026-11-17: a duplicated `run:` key made the
workflow invalid and CI ran nothing for 37 pushes, releases included).

Run: python lite_guard/run_all.py
"""
import glob
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)


def commands():
    tests = sorted(glob.glob(os.path.join(HERE, 'test_*.py')))
    return [os.path.join(HERE, 'guard.py')] + tests


def main():
    env = dict(os.environ, PYTHONPATH='', PYTHONWARNINGS='ignore')
    failed = []
    for path in commands():
        name = os.path.relpath(path, REPO).replace(os.sep, '/')
        start = time.time()
        proc = subprocess.Popen([sys.executable, path], cwd=REPO, env=env,
                                stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
        out = proc.communicate()[0].decode('utf-8', 'replace')
        ok = proc.returncode == 0
        print('{}  {:5.1f}s  {}'.format('ok  ' if ok else 'FAIL', time.time() - start, name))
        if not ok:
            failed.append(name)
            print('\n'.join('      ' + line for line in out.splitlines()[-40:]))
    print('\n{} commands, {} failed'.format(len(commands()), len(failed)))
    for name in failed:
        print('  failed: ' + name)
    return 1 if failed else 0


if __name__ == '__main__':
    sys.exit(main())
