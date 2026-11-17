# -*- coding: utf-8 -*-
"""pyRevit Routes kept off the T3Lab MCP ports (lib/routes_guard.py). Plain CPython 3."""
import io
import os
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, 'lib'))
import routes_guard as RG  # noqa: E402


class FakeConfig(object):
    def __init__(self, enabled, port):
        self.routes_server, self.routes_port, self.saved = enabled, port, 0

    def save_changes(self):
        self.saved += 1


INI = (u'[core]\r\nbincache = true\r\n\r\n[routes]\r\nenabled = true\r\nport = 48884\r\n'
       u'\r\n[telemetry]\r\nport = 48884\r\n')


class RoutesGuardTests(unittest.TestCase):
    def test_new_port(self):
        self.assertEqual(RG.new_port(True, 48884), RG.ROUTES_PORT)
        self.assertEqual(RG.new_port('true', '48894'), RG.ROUTES_PORT)
        self.assertEqual(RG.new_port(True, None), RG.ROUTES_PORT)      # pyRevit default 48884
        self.assertIsNone(RG.new_port(False, 48884))                   # Routes off: leave it
        self.assertIsNone(RG.new_port(True, 48895))
        self.assertIsNone(RG.new_port(True, 48870))
        self.assertTrue(RG.ROUTES_PORT + 13 < RG.T3LAB_PORT_MIN)       # room for 14 windows

    def test_api_moves_the_port_for_this_session(self):
        cfg = FakeConfig(True, 48884)
        note = RG.apply(user_config=cfg)
        self.assertEqual((cfg.routes_port, cfg.saved), (RG.ROUTES_PORT, 1))
        self.assertIn('48884', note)
        self.assertEqual(RG.apply(user_config=cfg), '')                # second start: nothing
        off = FakeConfig(False, 48884)
        self.assertEqual(RG.apply(user_config=off), '')
        self.assertEqual((off.routes_port, off.saved), (48884, 0))

    def test_ini_fallback_keeps_the_file(self):
        path = os.path.join(tempfile.mkdtemp(), 'pyRevit_config.ini')
        with io.open(path, 'w', encoding='utf-8', newline='') as f:
            f.write(INI)
        note = RG.apply(path=path)
        with io.open(path, 'r', encoding='utf-8', newline='') as f:
            text = f.read()
        self.assertIn('takes effect at the next Revit start', note)
        self.assertIn(u'[routes]\r\nenabled = true\r\nport = {}\r\n'.format(RG.ROUTES_PORT), text)
        self.assertIn(u'[telemetry]\r\nport = 48884', text)          # other sections untouched
        self.assertEqual(RG.apply(path=path), '')

    def test_ini_cases(self):
        self.assertEqual(RG.ini_fix(u'[routes]\nenabled = false\nport = 48884'), (None, None))
        self.assertEqual(RG.ini_fix(u'[core]\nx = 1'), (None, None))
        text, old = RG.ini_fix(u'[routes]\nenabled = True\n\n[core]\nx = 1')
        self.assertEqual(old, RG.ROUTES_DEFAULT_PORT)
        self.assertEqual(text, u'[routes]\nenabled = True\nport = {}\n\n[core]\nx = 1'.format(
            RG.ROUTES_PORT))
        self.assertEqual(RG.ini_fix(u'[routes]\nenabled=true\nport=48890')[0],
                         u'[routes]\nenabled=true\nport={}'.format(RG.ROUTES_PORT))

    def test_never_raises(self):
        class Broken(object):
            routes_server = True
            routes_port = 48884

            def save_changes(self):
                raise IOError('read-only')
        self.assertIn('failed', RG.apply(user_config=Broken(), path=os.devnull + 'x') or 'failed')
        self.assertEqual(RG.apply(path=os.path.join(tempfile.mkdtemp(), 'missing.ini')), '')

    def test_startup_calls_it_before_the_mcp_server(self):
        src = io.open(os.path.join(ROOT, 'startup.py'), encoding='utf-8').read()
        call = src.index('_load_lib_module("routes_guard")')
        self.assertLess(call, src.index('MCPService.start_server()'))

    def test_ports_match_the_server(self):
        src = io.open(os.path.join(ROOT, 'lib', 'core', 'server.py'), encoding='utf-8').read()
        self.assertIn('PORT_MIN = {}'.format(RG.T3LAB_PORT_MIN), src)
        self.assertIn('PORT_MAX = {}'.format(RG.T3LAB_PORT_MAX), src)


if __name__ == '__main__':
    unittest.main()
