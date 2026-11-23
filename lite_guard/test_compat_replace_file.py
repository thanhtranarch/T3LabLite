# -*- coding: utf-8 -*-
"""Snippets._compat.replace_file: dst survives a failed move on both engines.

The IronPython branches are simulated on CPython: os.replace removed, a fake
System.IO.File (or none) in sys.modules.

Run:  python lite_guard/test_compat_replace_file.py      (no Revit)
"""
import io
import os
import shutil
import sys
import tempfile
import types

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'lib'))
from Snippets._compat import replace_file   # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


def write(path, text):
    with io.open(path, 'w', encoding='utf-8') as f:
        f.write(text)


def read(path):
    with io.open(path, encoding='utf-8') as f:
        return f.read()


class FakeFile(object):
    """System.IO.File stand-in; `fail` makes Replace / Move raise."""
    fail = False
    calls = []

    @classmethod
    def Replace(cls, src, dst, backup):
        cls.calls.append('Replace')
        if cls.fail:
            raise IOError('volume mismatch')
        shutil.move(src, dst)

    @classmethod
    def Move(cls, src, dst):
        cls.calls.append('Move')
        if cls.fail:
            raise IOError('volume mismatch')
        shutil.move(src, dst)


def ipy(net):
    """Context: no os.replace; System.IO.File = FakeFile (net=True) or missing."""
    class _Ctx(object):
        def __enter__(self):
            self.rep = os.replace
            del os.replace
            self.saved = dict((k, sys.modules.get(k)) for k in ('System', 'System.IO'))
            if net:
                system = types.ModuleType('System')
                sysio = types.ModuleType('System.IO')
                sysio.File = FakeFile
                system.IO = sysio
                sys.modules['System'], sys.modules['System.IO'] = system, sysio
            else:
                sys.modules['System'] = sys.modules['System.IO'] = None   # ImportError
            FakeFile.calls, FakeFile.fail = [], False
            return self

        def __exit__(self, *exc):
            os.replace = self.rep
            for k, v in self.saved.items():
                if v is None:
                    sys.modules.pop(k, None)
                else:
                    sys.modules[k] = v
    return _Ctx()


def case(tmp, name):
    d = os.path.join(tmp, name)
    os.makedirs(d)
    src, dst = os.path.join(d, 'new.tmp'), os.path.join(d, 'data.abs')
    write(src, 'new')
    write(dst, 'old')
    return d, src, dst


tmp = tempfile.mkdtemp(prefix='t3lab_replace_')
try:
    d, src, dst = case(tmp, 'cpython')
    replace_file(src, dst)
    check('CPython: os.replace moves src over dst', read(dst) == 'new' and not os.path.exists(src))

    d, src, dst = case(tmp, 'net_replace')
    with ipy(net=True):
        replace_file(src, dst)
        calls = list(FakeFile.calls)
    check('IronPython: File.Replace when dst exists', calls == ['Replace'] and read(dst) == 'new')

    d, src, dst = case(tmp, 'net_move')
    os.remove(dst)
    with ipy(net=True):
        replace_file(src, dst)
        calls = list(FakeFile.calls)
    check('IronPython: File.Move when dst is missing', calls == ['Move'] and read(dst) == 'new')

    d, src, dst = case(tmp, 'net_fails')
    with ipy(net=True):
        FakeFile.fail = True
        replace_file(src, dst)
    check('IronPython: .NET failure falls back to the rename path',
          read(dst) == 'new' and os.listdir(d) == ['data.abs'], os.listdir(d))

    d, src, dst = case(tmp, 'rename')
    with ipy(net=False):
        replace_file(src, dst)
    check('no .NET: rename via backup, backup removed',
          read(dst) == 'new' and os.listdir(d) == ['data.abs'], os.listdir(d))

    d, src, dst = case(tmp, 'rename_fails')
    real_rename = os.rename

    def flaky(a, b):
        if a == src:
            raise OSError('locked by antivirus')
        return real_rename(a, b)
    raised = False
    with ipy(net=False):
        os.rename = flaky
        try:
            replace_file(src, dst)
        except OSError:
            raised = True
        finally:
            os.rename = real_rename
    check('failed move raises and dst keeps its old content',
          raised and read(dst) == 'old' and sorted(os.listdir(d)) == ['data.abs', 'new.tmp'],
          (raised, sorted(os.listdir(d))))
finally:
    shutil.rmtree(tmp, ignore_errors=True)

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
