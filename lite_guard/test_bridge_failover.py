# -*- coding: utf-8 -*-
"""The bridge never runs one tools/call in two Revit windows.

With two Revit windows open, a call that reached window A and then lost its
answer (crash, timeout) used to be sent again to window B: the same write ran
on another document, and B silently became the target (review 2026-11-18).
Only a refused connection - nothing was delivered - may fail over.

Run: python lite_guard/test_bridge_failover.py
"""
import os
import socket
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(REPO, 'lib'))

from core import bridge as B    # noqa: E402

FAILED = []


def check(label, ok, detail=''):
    print('  {} {}{}'.format('ok  ' if ok else 'FAIL', label, ('  ' + str(detail)) if not ok else ''))
    if not ok:
        FAILED.append(label)


def run(method, error, alive=(48885,), a_alive=False):
    """_forward with window A (48884) failing with `error`; (result or exception, ports posted)."""
    posted = []

    def post(port, request, token, timeout):
        posted.append(port)
        if port == 48884:
            raise error
        return {'jsonrpc': '2.0', 'id': 1, 'result': {'port': port}}

    saved = B._post_mcp, B._alive_ports, B._port_alive
    B._post_mcp = post
    B._alive_ports = lambda: list(alive)
    B._port_alive = lambda port, timeout=None: (a_alive if port == 48884 else port in alive)
    current = {'port': 48884, 'validated': True}
    try:
        try:
            out = B._forward({'jsonrpc': '2.0', 'id': 1, 'method': method}, current, 't')
        except Exception as ex:
            out = ex
        return out, posted, current
    finally:
        B._post_mcp, B._alive_ports, B._port_alive = saved


def main():
    print('a tools/call that was delivered is never sent to another Revit')
    for name, error in (('timeout', socket.timeout('timed out')),
                        ('connection reset', ConnectionResetError(104, 'reset'))):
        out, posted, current = run('tools/call', error)
        check('{}: the error reaches the client'.format(name), isinstance(out, Exception), out)
        check('{}: only window A was called'.format(name), posted == [48884], posted)
        check('{}: the target stays A, and is checked again next time'.format(name),
              current['port'] == 48884 and current['validated'] is False, current)

    print('nothing delivered: the call may go to the other window')
    out, posted, current = run('tools/call', ConnectionRefusedError(111, 'refused'))
    check('refused tools/call runs once, in B', posted == [48884, 48885]
          and isinstance(out, dict) and current['port'] == 48885, (out, posted))

    print('a read of a dead port fails over')
    out, posted, current = run('tools/list', socket.timeout('timed out'))
    check('tools/list is answered by B', isinstance(out, dict) and posted == [48884, 48885], (out, posted))
    out, posted, current = run('tools/list', socket.timeout('timed out'), a_alive=True)
    check('but not while A is still healthy', isinstance(out, Exception) and posted == [48884],
          (out, posted))

    if FAILED:
        print('\n{} check(s) failed.'.format(len(FAILED)))
        return 1
    print('\nAll bridge failover checks passed.')
    return 0


if __name__ == '__main__':
    sys.exit(main())
