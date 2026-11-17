# -*- coding: utf-8 -*-
"""Opt-in timing log for opening a tool: which step takes the time.

On while %APPDATA%\\T3LabAI\\open_timing.on exists or env T3LAB_OPEN_TIMING=1
(`python lib/open_timing.py on|off|show|clear`). Off, it costs one stat() per
click and per window. One line per window in %APPDATA%\\T3LabAI\\open_times.log:

  bootstrap   init_cpython_paths()                (first window of a click only)
  imports     bootstrap end -> window __init__    (script + dialog module imports)
  xaml_read / sanitize / wpf_parse (XamlReader) / wiring (names, events, logo)
  clips       rounded window / panel clips (logical tree walk)
  load_rest   owner, Esc, dispatcher guard
  init_data   load_xaml end -> SourceInitialized  (tool __init__: model reads)
  layout      SourceInitialized -> Loaded
  first_render  Loaded -> ContentRendered

Time from the ribbon click to the script's first line (pyRevit) is not seen.
Hooks: _cpython_bootstrap.init_cpython_paths, GUI.WPF_Base.T3WPFWindow.
Engine-agnostic (IronPython 2.7 + CPython 3). Test: lite_guard/test_open_timing.py.
"""

from __future__ import print_function, unicode_literals

import io
import os
import sys
import time

ENV = 'T3LAB_OPEN_TIMING'
FLAG_NAME = 'open_timing.on'
LOG_NAME = 'open_times.log'
LOG_MAX_BYTES = 256 * 1024          # then rotated to open_times.log.1
RUN_REUSE_S = 120                   # same script bootstrapping again = same click

_clock = getattr(time, 'perf_counter', None) or time.time
_state = {'run': None, 'opens': 0}


def _dir():
    base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI')


def flag_path():
    return os.path.join(_dir(), FLAG_NAME)


def log_path():
    return os.path.join(_dir(), LOG_NAME)


def enabled():
    if os.environ.get(ENV, '').strip().lower() in ('1', 'true', 'on', 'yes'):
        return True
    try:
        return os.path.exists(flag_path())
    except Exception:
        return False


def set_enabled(on):
    """Create / remove the flag file; takes effect on the next click."""
    path = flag_path()
    if on:
        if not os.path.isdir(_dir()):
            os.makedirs(_dir())
        with io.open(path, 'w', encoding='utf-8') as f:
            f.write('Delete this file to stop the T3Lab open timing log.\n')
    elif os.path.exists(path):
        os.remove(path)
    return bool(on)


# ─── Pure helpers ──────────────────────────────────────────────────────────────

def bundle_name(script_path):
    """'.../BatchOut.pushbutton/script.py' -> 'BatchOut'."""
    if not script_path:
        return '?'
    folder = os.path.basename(os.path.dirname(script_path.replace('\\', '/')))
    return folder.rsplit('.', 1)[0] if '.' in folder else (folder or '?')


def format_line(when, tool, xaml, total_ms, steps, opens, engine, xaml_lines):
    """One log line; `steps` = [(name, ms)] in order, the slowest marked '*'."""
    slowest = max(steps, key=lambda s: s[1])[0] if steps else None
    parts = ['{}{} {}'.format('*' if name == slowest else '', name, int(round(ms)))
             for name, ms in steps]
    return '{}  {:<20} {:<28} total {:>6} ms | {} | open #{} in this Revit session · {} · {} XAML lines'.format(
        when, tool[:20], xaml[:28], int(round(total_ms)), ' · '.join(parts),
        opens, engine, xaml_lines)


# ─── Run (one ribbon click) ────────────────────────────────────────────────────

def _caller_script():
    try:
        frame = sys._getframe(1)
    except Exception:
        return None
    while frame is not None:
        name = frame.f_code.co_filename or ''
        if os.path.basename(name).lower() == 'script.py':
            return name
        frame = frame.f_back
    return None


def run_started():
    """init_cpython_paths() entry. A second call by the same script (lib
    modules bootstrap too) keeps the run."""
    if not enabled():
        _state['run'] = None
        return
    now = _clock()
    script = _caller_script()
    run = _state['run']
    if (run is not None and run['script'] == script and not run['windows']
            and now - run['t0'] < RUN_REUSE_S):
        return
    _state['run'] = {'script': script, 'tool': bundle_name(script), 't0': now,
                     'bootstrap_end': None, 'windows': 0}


def bootstrap_done():
    run = _state['run']
    if run is not None and run['bootstrap_end'] is None:
        run['bootstrap_end'] = _clock()


# ─── Window ────────────────────────────────────────────────────────────────────

class WindowTimer(object):
    """Steps of one T3WPFWindow, from __init__ to ContentRendered."""

    def __init__(self, xaml, run=None, now=None):
        self.start = _clock() if now is None else now
        self.last = self.start
        self.xaml = xaml or '<inline XAML>'
        self.xaml_lines = 0
        self.steps = []
        self.done = False
        self.run = run
        self.first = run is not None and run['windows'] == 0
        if self.first:
            if run['bootstrap_end'] is not None:
                self.steps.append(('bootstrap', (run['bootstrap_end'] - run['t0']) * 1000.0))
                self.steps.append(('imports', (self.start - run['bootstrap_end']) * 1000.0))
            else:
                self.steps.append(('before_window', (self.start - run['t0']) * 1000.0))
        if run is not None:
            run['windows'] += 1

    def step(self, name, now=None):
        now = _clock() if now is None else now
        self.steps.append((name, (now - self.last) * 1000.0))
        self.last = now

    def total_ms(self):
        t0 = self.run['t0'] if self.first else self.start
        return (self.last - t0) * 1000.0

    def finish(self, last_step='first_render'):
        """ContentRendered: write the line once. Never raises."""
        if self.done:
            return None
        self.done = True
        try:
            if last_step:
                self.step(last_step)
            _state['opens'] += 1
            tool = self.run['tool'] if self.run else '?'
            engine = 'IronPython' if sys.platform == 'cli' else 'CPython %d.%d' % sys.version_info[:2]
            line = format_line(time.strftime('%Y-%m-%d %H:%M:%S'), tool, self.xaml,
                               self.total_ms(), self.steps, _state['opens'],
                               engine, self.xaml_lines)
            _append(line)
            return line
        except Exception:
            return None


def window_started(xaml_source, literal_string=False):
    """WindowTimer for a new T3WPFWindow, or None while the log is off."""
    if not enabled():
        return None
    name = ('<inline XAML>' if literal_string
            else '{}'.format(xaml_source).replace('\\', '/').rsplit('/', 1)[-1])
    return WindowTimer(name, _state['run'])


def _append(line):
    folder = _dir()
    if not os.path.isdir(folder):
        os.makedirs(folder)
    path = log_path()
    if os.path.exists(path) and os.path.getsize(path) > LOG_MAX_BYTES:
        try:
            os.replace(path, path + '.1')
        except Exception:
            os.remove(path)
    with io.open(path, 'a', encoding='utf-8') as f:
        f.write(line + '\n')


def _main(argv):
    cmd = (argv[1] if len(argv) > 1 else 'show').lower()
    if cmd == 'on':
        set_enabled(True)
        print('Open timing log is ON. Open the slow tools, then send:', log_path())
    elif cmd == 'off':
        set_enabled(False)
        print('Open timing log is OFF.')
    elif cmd == 'clear':
        if os.path.exists(log_path()):
            os.remove(log_path())
        print('Cleared', log_path())
    elif cmd == 'show':
        print('Open timing log is', 'ON' if enabled() else 'OFF', '-', log_path())
        if os.path.exists(log_path()):
            with io.open(log_path(), 'r', encoding='utf-8') as f:
                print(f.read())
    else:
        print('Usage: python open_timing.py on|off|show|clear')
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(_main(sys.argv))
