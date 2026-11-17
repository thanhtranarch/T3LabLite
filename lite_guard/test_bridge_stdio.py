# -*- coding: utf-8 -*-
"""
CPython 3 test: core/bridge.py as an AI app runs it -- a separate process on
stdin/stdout -- against a fake in-Revit server on 127.0.0.1. No Revit needed.

  - Vietnamese text in a tool call reaches Revit intact even when Python's
    default stdin encoding is an ANSI code page (Windows + -I / -E).
  - A proxy in the environment is never used for 127.0.0.1.
  - Junk lines (5, null, a list) are skipped; the bridge keeps answering.
  - tools/list falls back to the cached manifest when Revit answers an error.

Run:  python3 lite_guard/test_bridge_stdio.py   (exit 0 = all pass)
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import unittest
from http.server import BaseHTTPRequestHandler, HTTPServer

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
BRIDGE = os.path.join(LIB, 'core', 'bridge.py')
sys.path.insert(0, LIB)

from core import bridge  # noqa: E402

TEXT = u'Đường ống – Tầng 1 cửa'


class _Fake(BaseHTTPRequestHandler):
    received = []

    def log_message(self, *a):
        pass

    def _send(self, obj):
        body = json.dumps(obj).encode('utf-8')
        self.send_response(200)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        self._send({'status': 'ok'})

    def do_POST(self):
        req = json.loads(self.rfile.read(int(self.headers['Content-Length'])).decode('utf-8'))
        _Fake.received.append(req)
        self._send({'jsonrpc': '2.0', 'id': req.get('id'),
                    'result': {'content': [{'type': 'text', 'text': 'ok'}]}})


class StdioTests(unittest.TestCase):
    def setUp(self):
        _Fake.received = []
        self.http = HTTPServer(('127.0.0.1', 0), _Fake)
        self.port = self.http.server_address[1]
        threading.Thread(target=self.http.serve_forever, daemon=True).start()
        self.appdata = tempfile.mkdtemp()

    def tearDown(self):
        self.http.shutdown()
        self.http.server_close()
        shutil.rmtree(self.appdata, ignore_errors=True)

    def _run(self, lines):
        env = dict(os.environ)
        env.update({
            'APPDATA': self.appdata,
            'PYTHONPATH': '',
            # What Windows gives a piped stdin under -I: an ANSI code page.
            'PYTHONIOENCODING': 'cp1252',
            # A proxy that does not exist: using it would fail every call.
            'HTTP_PROXY': 'http://127.0.0.1:9', 'http_proxy': 'http://127.0.0.1:9',
            'NO_PROXY': '', 'no_proxy': '',
        })
        data = b''.join(json.dumps(l, ensure_ascii=False).encode('utf-8') + b'\n'
                        if not isinstance(l, bytes) else l for l in lines)
        out = subprocess.run([sys.executable, BRIDGE, str(self.port)], input=data,
                             env=env, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                             timeout=60)
        replies = [json.loads(l) for l in out.stdout.decode('ascii').splitlines() if l.strip()]
        return out, replies

    def test_utf8_proxy_and_junk_lines(self):
        call = {'jsonrpc': '2.0', 'id': 7, 'method': 'tools/call',
                'params': {'name': 'find_views', 'arguments': {'name': TEXT}}}
        out, replies = self._run([b'5\n', b'null\n', b'[1]\n', b'not json\n',
                                  {'jsonrpc': '2.0', 'method': 'notifications/initialized'},
                                  call])
        self.assertEqual(out.returncode, 0, out.stderr.decode('utf-8', 'replace'))
        self.assertEqual([r.get('id') for r in replies], [7], replies)
        self.assertIn('result', replies[0], replies[0])
        sent = [r for r in _Fake.received if r.get('method') == 'tools/call']
        self.assertEqual(len(sent), 1, _Fake.received)
        self.assertEqual(sent[0]['params']['arguments']['name'], TEXT)


class ToolsListFallbackTests(unittest.TestCase):
    def setUp(self):
        self._orig = (bridge._forward, bridge._load_cached_tools, bridge._save_cached_tools)

    def tearDown(self):
        bridge._forward, bridge._load_cached_tools, bridge._save_cached_tools = self._orig

    def test_error_from_revit_uses_cache(self):
        cached = [{'name': 'cached_tool', 'inputSchema': {'type': 'object'}}]
        bridge._forward = lambda *a, **k: {'jsonrpc': '2.0', 'id': 1,
                                           'error': {'code': -32603, 'message': 'reloading'}}
        bridge._load_cached_tools = lambda: list(cached)
        bridge._save_cached_tools = lambda tools: None
        resp = bridge._handle_tools_list({'jsonrpc': '2.0', 'id': 1, 'method': 'tools/list'},
                                         {'port': bridge.PORT_MIN}, '')
        names = [t['name'] for t in resp['result']['tools']]
        self.assertIn('cached_tool', names)


if __name__ == '__main__':
    unittest.main()
