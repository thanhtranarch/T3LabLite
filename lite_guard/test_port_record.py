# -*- coding: utf-8 -*-
"""One probe, not a scan, to find the MCP server of this Revit.

MCPService.start_server runs on Revit's UI thread at every Revit start and
every document open. It used to probe all ports of the range one after
another (a closed loopback port is slow to refuse on Windows, and IronPython
ignores the connect timeout), so each call could cost seconds. The server
now records its port for the whole process and the service probes only that.

Run: python lite_guard/test_port_record.py
"""
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)

from Services import mcp_service as M   # noqa: E402

FAILED = []


def check(label, ok, detail=''):
    print('  {} {}{}'.format('ok  ' if ok else 'FAIL', label, ('  ' + str(detail)) if not ok else ''))
    if not ok:
        FAILED.append(label)


def run(recorded, alive=(), platform=None):
    """(_local_servers() result, ports probed) with the record and /health faked."""
    probed = []

    def probe(port, timeout=0.35):
        probed.append(port)
        return {'status': 'ok', 'pid': os.getpid(), 'port': port} if port in alive else None

    saved = M._recorded_port, M._probe_health, sys.platform
    M._recorded_port, M._probe_health = (lambda: recorded), probe
    if platform:
        sys.platform = platform
    try:
        return M._local_servers(), probed
    finally:
        M._recorded_port, M._probe_health, sys.platform = saved


def main():
    span = M.SERVER_PORT_MAX - M.SERVER_PORT_MIN + 1

    print('a recorded port is the only one probed')
    found, probed = run(48890, alive=(48890,))
    check('the server is found', [p for p, _ in found] == [48890], found)
    check('one probe', probed == [48890], probed)

    print('recorded as stopped: nothing is probed')
    found, probed = run(0, alive=(48884,))
    check('no server, no probe', found == [] and probed == [], (found, probed))

    print('a recorded port that no longer answers')
    found, probed = run(48890)
    check('reads as stopped after one probe', found == [] and probed == [48890], (found, probed))

    print('no record')
    found, probed = run(None, alive=(48886,), platform='cli')
    check('IronPython (the host engine): not started here, no scan',
          found == [] and probed == [], (found, probed))
    found, probed = run(None, alive=(48886,), platform='win32')
    check('elsewhere: the whole range is scanned, as before',
          [p for p, _ in found] == [48886] and len(probed) == span, (found, probed))

    print('the server keeps the record')
    with io.open(os.path.join(LIB, 'core', 'server.py'), encoding='utf-8') as f:
        src = f.read()
    with io.open(os.path.join(LIB, 'Services', 'mcp_service.py'), encoding='utf-8') as f:
        svc = f.read()
    body = src.split('        def run_server():', 1)[1].split('self._server_thread = ', 1)[0]
    check('set when serving starts', '_record_port(port)' in body.split('serve_forever')[0])
    check('cleared when serving ends', '_record_port(0)' in body.split('finally:', 1)[1])
    key = src.split("_PROCESS_SINGLETON_KEY = '", 1)[1].split("'", 1)[0] + '.port'
    check('service and server use one key', "GetData('{}')".format(key) in svc
          and "_PROCESS_SINGLETON_KEY + '.port'" in src)

    if FAILED:
        print('\n{} check(s) failed.'.format(len(FAILED)))
        return 1
    print('\nAll port record checks passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
