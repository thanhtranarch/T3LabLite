# -*- coding: utf-8 -*-
"""
CPython 3 test: opt-in open timing log (lib/open_timing.py).

Run:  python3 lite_guard/test_open_timing.py      Exit 0 = all pass.

Covers the switch (flag file / env), one run per click, the steps of a
window, the log line, rotation, and that the bootstrap and T3WPFWindow
hooks are in place. The WPF events themselves need Revit.
"""
import ast
import io
import os
import shutil
import sys
import tempfile

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'T3Lab.extension', 'lib')
if not os.path.isdir(LIB):
    LIB = os.path.join(REPO, 'lib')         # T3LabLite layout
sys.path.insert(0, LIB)

_TMP = tempfile.mkdtemp(prefix='t3timing_')
os.environ['APPDATA'] = _TMP
os.environ.pop('T3LAB_OPEN_TIMING', None)

import open_timing as OT   # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok   ' + name)
    else:
        print('  FAIL ' + name + (' -- ' + str(detail) if detail else ''))
        FAILURES.append(name)


def fake_click(tool):
    """Run run_started()/bootstrap_done() from a frame named like a script.py."""
    path = os.path.join(_TMP, 'Demo.tab', 'X.panel', tool + '.pushbutton', 'script.py')
    code = compile('import open_timing\nopen_timing.run_started()\n'
                   'open_timing.run_started()\nopen_timing.bootstrap_done()\n', path, 'exec')
    exec(code, {})


def test_switch():
    print('[switch]')
    check('off by default', not OT.enabled())
    check('no window timer while off', OT.window_started('A.xaml') is None)
    OT.set_enabled(True)
    check('flag file turns it on', OT.enabled() and os.path.exists(OT.flag_path()))
    OT.set_enabled(False)
    check('off again', not OT.enabled())
    os.environ['T3LAB_OPEN_TIMING'] = '1'
    check('env turns it on', OT.enabled())
    del os.environ['T3LAB_OPEN_TIMING']


def test_helpers():
    print('[helpers]')
    check('bundle name', OT.bundle_name('C:\\ext\\Demo.tab\\V.panel\\BatchOut.pushbutton\\script.py') == 'BatchOut')
    check('no script -> ?', OT.bundle_name(None) == '?')
    line = OT.format_line('2026-11-17 10:00:00', 'BatchOut', 'ExportManager.xaml', 2310.4,
                          [('bootstrap', 40), ('wpf_parse', 820.2), ('init_data', 380)],
                          3, 'CPython 3.12', 2850)
    check('slowest step marked', '*wpf_parse 820' in line and '*bootstrap' not in line, line)
    check('total + counts', 'total   2310 ms' in line and 'open #3' in line and '2850 XAML lines' in line, line)


def test_run_and_window():
    print('[run + window]')
    OT.set_enabled(True)
    fake_click('BatchOut')
    run = OT._state['run']
    check('run named after the bundle', run and run['tool'] == 'BatchOut', run)
    t0 = run['t0']
    fake_click('BatchOut')
    check('a second bootstrap of the same click keeps the run', OT._state['run']['t0'] == t0)
    timer = OT.window_started('C:\\x\\ExportManager.xaml')
    check('window timer', timer is not None and timer.xaml == 'ExportManager.xaml')
    check('first window carries bootstrap + imports',
          [s[0] for s in timer.steps] == ['bootstrap', 'imports'], timer.steps)
    for name in ('xaml_read', 'sanitize', 'wpf_parse', 'wiring', 'clips', 'load_rest',
                 'init_data', 'layout'):
        timer.step(name)
    timer.xaml_lines = 2850
    line = timer.finish()
    check('finish writes one line', line and 'BatchOut' in line and 'first_render' in line, line)
    check('finish only once', timer.finish() is None)
    with io.open(OT.log_path(), encoding='utf-8') as f:
        lines = f.read().splitlines()
    check('line in the log file', len(lines) == 1 and lines[0] == line, lines)
    second = OT.window_started('T3Dialog.xaml')
    check('a second window of the click has no bootstrap step',
          second is not None and not second.steps, second and second.steps)
    fake_click('ManaSheets')
    check('a new click starts a new run', OT._state['run']['tool'] == 'ManaSheets')
    OT.set_enabled(False)
    fake_click('ManaSheets')
    check('off: no run kept', OT._state['run'] is None)


def test_rotation():
    print('[rotation]')
    with io.open(OT.log_path(), 'w', encoding='utf-8') as f:
        f.write('x' * (OT.LOG_MAX_BYTES + 10))
    OT._append('new line')
    check('big log rotated to .1', os.path.exists(OT.log_path() + '.1'))
    with io.open(OT.log_path(), encoding='utf-8') as f:
        check('fresh log starts with the new line', f.read() == 'new line\n')


def test_hooks():
    print('[hooks]')
    boot = io.open(os.path.join(LIB, '_cpython_bootstrap.py'), encoding='utf-8').read()
    fn = [n for n in ast.parse(boot).body
          if isinstance(n, ast.FunctionDef) and n.name == 'init_cpython_paths'][0]
    calls = [ast.get_source_segment(boot, st) for st in fn.body]
    check('run_started near the top of init_cpython_paths',
          any("_timing('run_started')" in c for c in calls[:3]), calls[:3])
    check('bootstrap_done is its last statement', "_timing('bootstrap_done')" in calls[-1], calls[-1])
    wpf = io.open(os.path.join(LIB, 'GUI', 'WPF_Base.py'), encoding='utf-8').read()
    for step in ('xaml_read', 'sanitize', 'wpf_parse', 'wiring', 'clips', 'load_rest'):
        check('WPF_Base marks ' + step, "_t3_step('{}')".format(step) in wpf)
    for event in ('SourceInitialized', 'Loaded', 'ContentRendered'):
        check('window event ' + event, 'self.{} += on_'.format(event) in wpf)
    src = io.open(OT.__file__, encoding='utf-8').read()
    import re
    check('no f-strings (IronPython)', not re.search(r'(?<![\w])[fF][rR]?[\'"]', src))


if __name__ == '__main__':
    try:
        test_switch()
        test_helpers()
        test_run_and_window()
        test_rotation()
        test_hooks()
    finally:
        shutil.rmtree(_TMP, ignore_errors=True)
    print()
    if FAILURES:
        print('FAILED: {}'.format(', '.join(FAILURES)))
        sys.exit(1)
    print('All open timing tests passed.')
