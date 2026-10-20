#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Tests for the automatic update schedule.

    python3 lite_guard/test_update_schedule.py

  * lib/core/update_schedule.py -- when a check is due (pure logic)
  * lib/core/updater.py          -- run_auto_update()/start_auto_update() order
                                    the stamps correctly; the .NET classes it
                                    imports are replaced by stubs

Plain CPython 3, standard library only -- no Revit, no pyRevit. CI runs it next
to guard.py (.github/workflows/lite-guard.yml).
"""
import importlib.util
import json
import os
import shutil
import sys
import tempfile
import types
import unittest
from datetime import date

sys.dont_write_bytecode = True

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_spec = importlib.util.spec_from_file_location(
    't3_update_schedule', os.path.join(ROOT, 'lib', 'core', 'update_schedule.py'))
sched = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(sched)

MON = date(2026, 10, 5)     # ISO 2026-W41
TUE = date(2026, 10, 6)
SUN = date(2026, 10, 11)    # last day of W41
NEXT_MON = date(2026, 10, 12)


class WeekId(unittest.TestCase):
    def test_week_runs_monday_to_sunday(self):
        self.assertEqual(sched.week_id(MON), '2026-W41')
        self.assertEqual(sched.week_id(SUN), '2026-W41')
        self.assertEqual(sched.week_id(NEXT_MON), '2026-W42')

    def test_iso_year_at_new_year(self):
        # 2026 has 53 ISO weeks; its last one runs into January 2027.
        self.assertEqual(sched.week_id(date(2026, 12, 31)), '2026-W53')
        self.assertEqual(sched.week_id(date(2027, 1, 1)), '2026-W53')
        self.assertEqual(sched.week_id(date(2027, 1, 3)), '2026-W53')
        self.assertEqual(sched.week_id(date(2027, 1, 4)), '2027-W01')

    def test_period_id_follows_the_interval(self):
        self.assertEqual(sched.period_id(TUE, 'weekly'), '2026-W41')
        self.assertEqual(sched.period_id(TUE, 'daily'), '2026-10-06')
        self.assertEqual(sched.period_id(TUE, None), '2026-W41')


class NormalizeInterval(unittest.TestCase):
    def test_default_is_weekly(self):
        for value in (None, '', 'weekly', 'WEEKLY', 'monthly', 'nonsense', 7, True):
            self.assertEqual(sched.normalize_interval(value), 'weekly', value)

    def test_daily_in_any_case(self):
        for value in ('daily', 'Daily', ' DAILY '):
            self.assertEqual(sched.normalize_interval(value), 'daily', value)


class StampPeriod(unittest.TestCase):
    def test_week_stamp(self):
        self.assertEqual(sched.stamp_period('2026-W41', 'weekly'), '2026-W41')

    def test_week_stamp_cannot_answer_for_a_day(self):
        self.assertEqual(sched.stamp_period('2026-W41', 'daily'), '')

    def test_old_daily_stamp_maps_to_its_week(self):
        # Installs that ran the daily check wrote a date. It counts for its week.
        self.assertEqual(sched.stamp_period('2026-10-07', 'weekly'), '2026-W41')
        self.assertEqual(sched.stamp_period('2026-10-11', 'weekly'), '2026-W41')
        self.assertEqual(sched.stamp_period('2026-10-12', 'weekly'), '2026-W42')

    def test_daily_stamp_in_daily_mode(self):
        self.assertEqual(sched.stamp_period('2026-10-07', 'daily'), '2026-10-07')

    def test_unreadable_stamps(self):
        for bad in (None, '', 'yesterday', '2026-13-40', '2026-02-30', '2026-W4', 12345, {}):
            self.assertEqual(sched.stamp_period(bad, 'weekly'), '', bad)


class IsDue(unittest.TestCase):
    def test_fresh_install_is_due(self):
        self.assertTrue(sched.is_due({}, MON))

    def test_completed_this_week_is_not_due_again(self):
        settings = {'last_update_check': '2026-W41'}
        for day in (MON, TUE, SUN):
            self.assertFalse(sched.is_due(settings, day), day)

    def test_due_again_on_the_first_start_of_the_next_week(self):
        settings = {'last_update_check': '2026-W41', 'last_update_attempt': '2026-10-09'}
        self.assertTrue(sched.is_due(settings, NEXT_MON))

    def test_not_opened_on_monday_still_checks_on_the_first_start_that_week(self):
        settings = {'last_update_check': '2026-W40'}
        self.assertTrue(sched.is_due(settings, date(2026, 10, 8)))   # Thursday

    def test_old_daily_stamp_from_this_week_is_respected(self):
        # Upgrading from the daily check on a Wednesday must not recheck on Thursday.
        settings = {'last_update_check': '2026-10-07'}
        self.assertFalse(sched.is_due(settings, date(2026, 10, 8)))
        self.assertTrue(sched.is_due(settings, NEXT_MON))

    def test_failed_attempt_is_not_repeated_the_same_day(self):
        settings = {'last_update_check': '2026-W40', 'last_update_attempt': '2026-10-05'}
        self.assertFalse(sched.is_due(settings, MON))

    def test_failed_attempt_is_retried_the_next_day(self):
        # Offline on Monday: no stamp for the week, so Tuesday tries again.
        settings = {'last_update_check': '2026-W40', 'last_update_attempt': '2026-10-05'}
        self.assertTrue(sched.is_due(settings, TUE))

    def test_unreadable_stamp_is_due(self):
        self.assertTrue(sched.is_due({'last_update_check': 'garbage'}, MON))

    def test_daily_interval(self):
        settings = {'auto_update_interval': 'daily', 'last_update_check': '2026-10-05'}
        self.assertFalse(sched.is_due(settings, MON))
        self.assertTrue(sched.is_due(settings, TUE))

    def test_switching_from_weekly_to_daily_checks_once(self):
        settings = {'auto_update_interval': 'daily', 'last_update_check': '2026-W41'}
        self.assertTrue(sched.is_due(settings, TUE))
        settings['last_update_check'] = sched.period_id(TUE, 'daily')
        self.assertFalse(sched.is_due(settings, TUE))

    def test_clock_set_back_does_not_wedge_the_check(self):
        # A stamp from a later week than today is simply a different period.
        settings = {'last_update_check': '2026-W45'}
        self.assertTrue(sched.is_due(settings, MON))


class SourceStaysIronPythonSafe(unittest.TestCase):
    """startup.py runs under IronPython 2.7, which imports this module."""

    def test_no_python3_only_syntax(self):
        path = os.path.join(ROOT, 'lib', 'core', 'update_schedule.py')
        with open(path, encoding='utf-8') as handle:
            source = handle.read()
        import ast
        ast.parse(source, feature_version=(3, 4))          # rejects f-strings, annotations
        for needle in ('nonlocal ', 'yield from', 'async ', 'await ', ' -> '):
            self.assertNotIn(needle, source)
        self.assertIn('from __future__ import unicode_literals', source)


# ── updater.py glue ──────────────────────────────────────────────────────────

def _stub_dotnet():
    """clr and the System.* classes updater.py imports at module level."""
    names = ('clr', 'System', 'System.Net', 'System.Text', 'System.Diagnostics')
    saved = dict((n, sys.modules.get(n)) for n in names)
    clr = types.ModuleType('clr')
    clr.AddReference = lambda *a, **k: None
    sys.modules['clr'] = clr
    for name, attrs in (
            ('System', {}),
            ('System.Net', ['WebClient', 'ServicePointManager',
                            'SecurityProtocolType', 'CredentialCache']),
            ('System.Text', ['Encoding']),
            ('System.Diagnostics', ['Process', 'ProcessStartInfo'])):
        mod = types.ModuleType(name)
        for attr in (attrs or []):
            setattr(mod, attr, type(str(attr), (object,), {}))
        sys.modules[name] = mod
    return saved


def _restore(saved):
    for name, mod in saved.items():
        if mod is None:
            sys.modules.pop(name, None)
        else:
            sys.modules[name] = mod


class UpdaterGlue(unittest.TestCase):
    """run_auto_update() stamps the right things in the right order."""

    @classmethod
    def setUpClass(cls):
        cls._saved_dotnet = _stub_dotnet()
        cls._saved_path = list(sys.path)
        sys.path.insert(0, os.path.join(ROOT, 'lib'))
        for name in ('core', 'core.updater', 'core.paths', 'core.update_schedule'):
            sys.modules.pop(name, None)
        cls.updater = importlib.import_module('core.updater')
        cls.paths = importlib.import_module('core.paths')

    @classmethod
    def tearDownClass(cls):
        for name in ('core', 'core.updater', 'core.paths', 'core.update_schedule'):
            sys.modules.pop(name, None)
        sys.path[:] = cls._saved_path
        _restore(cls._saved_dotnet)

    def setUp(self):
        self.appdata = tempfile.mkdtemp()
        self._env = os.environ.get('APPDATA')
        os.environ['APPDATA'] = self.appdata
        self.u = self.updater
        self.today = MON
        self.remote = '9.9.9'
        self.local = '1.0.0'
        self.apply_result = (True, 'updated via git pull')
        self.fetch_error = None
        self.fetches = 0
        self.applies = 0
        u = self.u
        self._patched = dict((n, getattr(u, n)) for n in (
            '_today', 'fetch_remote_version', 'read_local_version',
            'auto_apply_update', 'enable_tls12', 'log'))
        u._today = lambda: self.today
        u.read_local_version = lambda: self.local
        u.enable_tls12 = lambda: None
        u.log = lambda message: None

        def fetch():
            self.fetches += 1
            if self.fetch_error:
                raise self.fetch_error
            return self.remote

        def apply():
            self.applies += 1
            return self.apply_result

        u.fetch_remote_version = fetch
        u.auto_apply_update = apply

    def tearDown(self):
        for name, value in self._patched.items():
            setattr(self.u, name, value)
        if self._env is None:
            os.environ.pop('APPDATA', None)
        else:
            os.environ['APPDATA'] = self._env
        shutil.rmtree(self.appdata, ignore_errors=True)

    def settings(self):
        return self.paths.load_settings()

    def write(self, **values):
        data = self.settings()
        data.update(values)
        self.paths.save_settings(data)

    def test_first_start_of_the_week_updates_and_ticks_the_week_off(self):
        self.assertEqual(self.u.run_auto_update(), ('updated', '9.9.9'))
        self.assertEqual(self.settings()['last_update_check'], '2026-W41')
        self.assertEqual(self.settings()['last_update_attempt'], '2026-10-05')
        self.assertEqual(self.applies, 1)

    def test_later_starts_in_the_same_week_do_nothing(self):
        self.u.run_auto_update()
        for day in (TUE, SUN):
            self.today = day
            status, _ = self.u.run_auto_update()
            self.assertEqual(status, 'skipped', day)
        self.assertEqual((self.fetches, self.applies), (1, 1))

    def test_next_week_checks_again(self):
        self.u.run_auto_update()
        self.today = NEXT_MON
        self.assertEqual(self.u.run_auto_update()[0], 'updated')
        self.assertEqual(self.settings()['last_update_check'], '2026-W42')

    def test_already_current_still_completes_the_week(self):
        self.remote = self.local
        self.assertEqual(self.u.run_auto_update()[0], 'current')
        self.assertEqual(self.settings()['last_update_check'], '2026-W41')
        self.assertEqual(self.applies, 0)

    def test_offline_monday_retries_tuesday_not_the_same_day(self):
        self.fetch_error = Exception('offline')
        self.assertEqual(self.u.run_auto_update()[0], 'failed')
        self.assertNotIn('last_update_check', self.settings())
        self.assertEqual(self.u.run_auto_update()[0], 'skipped')      # same day
        self.assertEqual(self.fetches, 1)
        self.today = TUE
        self.fetch_error = None
        self.assertEqual(self.u.run_auto_update()[0], 'updated')
        self.assertEqual(self.settings()['last_update_check'], '2026-W41')

    def test_apply_failure_does_not_tick_the_week_off(self):
        self.apply_result = (False, 'git pull --ff-only failed: local edits')
        status, detail = self.u.run_auto_update()
        self.assertEqual(status, 'failed')
        self.assertIn('local edits', detail)
        self.assertNotIn('last_update_check', self.settings())
        self.today = TUE
        self.apply_result = (True, 'updated via git pull')
        self.assertEqual(self.u.run_auto_update()[0], 'updated')

    def test_opt_out_is_respected_and_kept(self):
        self.write(auto_update=False)
        self.assertEqual(self.u.run_auto_update()[0], 'disabled')
        self.assertFalse(self.u.start_auto_update(delay=0))
        self.assertIs(self.settings()['auto_update'], False)
        self.assertEqual(self.fetches, 0)

    def test_first_run_writes_the_options_into_the_settings_file(self):
        self.u.run_auto_update()
        settings = self.settings()
        self.assertIs(settings['auto_update'], True)
        self.assertEqual(settings['auto_update_interval'], 'weekly')

    def test_daily_interval_stamps_the_day(self):
        self.write(auto_update_interval='daily')
        self.u.run_auto_update()
        self.assertEqual(self.settings()['last_update_check'], '2026-10-05')
        self.today = TUE
        self.assertEqual(self.u.run_auto_update()[0], 'updated')

    def test_force_ignores_the_schedule_and_the_stamps(self):
        self.u.run_auto_update()
        self.assertEqual(self.u.run_auto_update(force=True)[0], 'updated')
        self.assertEqual(self.fetches, 2)

    def test_old_daily_stamp_from_this_week_is_not_rechecked(self):
        self.write(last_update_check='2026-10-05', auto_update=True)
        self.today = date(2026, 10, 8)
        self.assertEqual(self.u.run_auto_update()[0], 'skipped')

    def test_start_auto_update_only_spawns_a_thread_when_due(self):
        started = []

        class FakeThread(object):
            def __init__(self, target=None, args=()):
                started.append((target, args))
                self.daemon = False

            def start(self):
                started.append('started')

        real = self.u.threading.Thread
        self.u.threading.Thread = FakeThread
        try:
            self.assertTrue(self.u.start_auto_update(delay=0))
            self.assertIn('started', started)
            del started[:]
            self.write(last_update_check='2026-W41')
            self.assertFalse(self.u.start_auto_update(delay=0))
            self.assertEqual(started, [])
        finally:
            self.u.threading.Thread = real

    def test_old_function_names_still_resolve(self):
        self.assertIs(self.u.stamp_today, self.u.stamp_check)
        self.assertIs(self.u.run_daily_update, self.u.run_auto_update)
        self.assertIs(self.u.start_daily_update, self.u.start_auto_update)


if __name__ == '__main__':
    unittest.main()
