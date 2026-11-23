# -*- coding: utf-8 -*-
"""Conductor: the one Revit slot of an MCP server, shared fairly by every caller.

Revit runs one /mcp request at a time. A bare Lock let whoever won the race go
next: one agent's burst starved the others, a person waited behind a worker's
queue, and a call queued behind a long write still ran after its client had
given up (a write nobody waited for). The conductor orders the queue and
refuses, before anything runs, a call that cannot start within its wait budget.

Order: lowest class first; a waiting call moves up one class every AGING_S so
nobody starves; within a class the agent served longest ago goes first (round
robin), then arrival order. Classes come from the bridge's `t3lab/lane` stamp.

Pure Python for IronPython 2.7 / 3.4 and CPython. Docs: docs/agents/conductor.md.
Test: dev/test_conductor.py.
"""
import contextlib
import threading
import time

CLASS_CONTROL, CLASS_LEAD, CLASS_TASK, CLASS_AUTO = 0, 1, 2, 3
# lead = harness lead / master, person = a client with a person watching,
# task = holds a harness task, auto = autonomous without a task.
LANE_CLASS = {'lead': CLASS_LEAD, 'person': CLASS_LEAD,
              'task': CLASS_TASK, 'auto': CLASS_AUTO}
CONTROL_METHODS = ('initialize', 'tools/list')
AGING_S = 15.0
MAX_WAIT_S = 600.0

AGENT_META = 't3lab/agent'
TASK_META = 't3lab/task'
LANE_META = 't3lab/lane'
WAIT_META = 't3lab/queue_wait_s'   # seconds a call may wait to START; absent = no limit


class Busy(Exception):
    """The call could not start in time; `info` is the refusal payload."""

    def __init__(self, info):
        Exception.__init__(self, info.get('error'))
        self.info = info


class Ticket(object):
    __slots__ = ('seq', 'method', 'agent', 'tool', 'task', 'lane', 'klass',
                 'arrived', 'deadline', 'started')

    def __init__(self, method='tools/call', agent=None, tool=None, task=None,
                 lane=None, wait_s=None):
        self.seq = 0
        self.method = method
        self.agent = agent or ''
        self.tool = tool or ''
        self.task = task or ''
        self.lane = lane or ''
        if method in CONTROL_METHODS:
            self.klass = CLASS_CONTROL
        else:
            # No lane: an older bridge or a direct caller - the worker class.
            self.klass = LANE_CLASS.get(lane, CLASS_TASK)
        self.arrived = 0.0
        self.deadline = wait_s      # relative until queued
        self.started = None

    def view(self, now):
        out = {'agent': self.agent or None, 'tool': self.tool or self.method,
               'lane': self.lane or None}
        if self.task:
            out['task'] = self.task
        if self.started is not None:
            out['running_s'] = round(now - self.started, 1)
        else:
            out['waiting_s'] = round(now - self.arrived, 1)
        return out


def ticket_for(request):
    """Ticket for one JSON-RPC request dict (meta stamped by bridge.py)."""
    method = request.get('method') or ''
    params = request.get('params') or {}
    if not isinstance(params, dict):
        params = {}
    meta = params.get('_meta')
    meta = meta if isinstance(meta, dict) else {}
    wait_s = meta.get(WAIT_META)
    try:
        wait_s = None if wait_s is None else max(0.0, min(float(wait_s), MAX_WAIT_S))
    except (TypeError, ValueError):
        wait_s = None
    tool = params.get('name') if method == 'tools/call' else None
    return Ticket(method, meta.get(AGENT_META), tool, meta.get(TASK_META),
                  meta.get(LANE_META), wait_s)


class Conductor(object):

    def __init__(self, clock=None, aging_s=AGING_S):
        self._clock = clock or time.time
        self._aging_s = float(aging_s)
        self._cond = threading.Condition(threading.Lock())
        self._waiting = []
        self._holder = None
        self._seq = 0
        self._served = {}           # agent -> grant number of its last call
        self._grants = 0
        self._avg_s = 2.0           # moving average of a call, for retry_after_s
        self._local = threading.local()

    # ── ordering ────────────────────────────────────────────────────────
    def _rank(self, ticket, now):
        aged = int((now - ticket.arrived) / self._aging_s) if self._aging_s > 0 else 0
        klass = max(CLASS_CONTROL, ticket.klass - aged)
        return (klass, self._served.get(ticket.agent, -1), ticket.seq)

    def _next(self, now):
        if not self._waiting:
            return None
        return min(self._waiting, key=lambda t: self._rank(t, now))

    def order(self):
        """Waiting tickets in the order they would start (tests, snapshot)."""
        with self._cond:
            now = self._clock()
            return sorted(self._waiting, key=lambda t: self._rank(t, now))

    # ── slot ────────────────────────────────────────────────────────────
    def acquire(self, ticket):
        """Block until `ticket` holds the slot; Busy when its budget runs out."""
        with self._cond:
            now = self._clock()
            self._seq += 1
            ticket.seq = self._seq
            ticket.arrived = now
            if ticket.deadline is not None:
                ticket.deadline = now + ticket.deadline
            self._waiting.append(ticket)
            try:
                while True:
                    now = self._clock()
                    if self._holder is None and self._next(now) is ticket:
                        break
                    if ticket.deadline is not None and now >= ticket.deadline:
                        raise Busy(self._busy_info(ticket, now))
                    # Wake now and then: aging can change the order with no release.
                    wait = 1.0
                    if ticket.deadline is not None:
                        wait = min(wait, max(0.01, ticket.deadline - now))
                    self._cond.wait(wait)
            except BaseException:
                self._waiting.remove(ticket)
                self._cond.notify_all()
                raise
            self._waiting.remove(ticket)
            self._grants += 1
            self._served[ticket.agent] = self._grants
            ticket.started = now
            self._holder = ticket
        self._local.ticket = ticket
        return ticket

    def release(self, ticket):
        with self._cond:
            if self._holder is ticket:
                if ticket.started is not None:
                    took = max(0.0, self._clock() - ticket.started)
                    self._avg_s = 0.8 * self._avg_s + 0.2 * took
                self._holder = None
                self._cond.notify_all()
        if getattr(self._local, 'ticket', None) is ticket:
            self._local.ticket = None

    @contextlib.contextmanager
    def slot(self, ticket):
        self.acquire(ticket)
        try:
            yield ticket
        finally:
            self.release(ticket)

    @contextlib.contextmanager
    def step_aside(self):
        """Lend the slot while this thread waits on a person (approval dialog).

        Comes back first in line and without a deadline: the person said yes,
        so the call must not be refused afterwards."""
        ticket = getattr(self._local, 'ticket', None)
        if ticket is None or self._holder is not ticket:
            yield
            return
        self.release(ticket)
        try:
            yield
        finally:
            back = Ticket(ticket.method, ticket.agent, ticket.tool, ticket.task,
                          ticket.lane, None)
            back.klass = CLASS_CONTROL
            self.acquire(back)
            # The handler releases the ticket it acquired; hand it the new one.
            with self._cond:
                ticket.started = back.started
                self._holder = ticket
            self._local.ticket = ticket

    # ── views ───────────────────────────────────────────────────────────
    def _busy_info(self, ticket, now):
        ahead = sum(1 for t in self._waiting
                    if t is not ticket and self._rank(t, now) < self._rank(ticket, now))
        holder = self._holder.view(now) if self._holder is not None else None
        retry = int(min(60, max(2, round(self._avg_s * (ahead + 1)))))
        return {
            'error': ('Revit is busy with another call and this one could not start '
                      'within {:.0f} s. Nothing ran; retry in about {} s.').format(
                          now - ticket.arrived, retry),
            'error_code': 'revit_busy', 'retryable': True, 'nothing_ran': True,
            'retry_after_s': retry, 'running': holder, 'ahead': ahead,
        }

    def snapshot(self):
        """{running, waiting:[...]} in start order; JSON-safe."""
        with self._cond:
            now = self._clock()
            order = sorted(self._waiting, key=lambda t: self._rank(t, now))
            return {'running': self._holder.view(now) if self._holder is not None else None,
                    'waiting': [t.view(now) for t in order]}

    def summary(self):
        """Small /health view: who runs, how long, how many wait (no task titles)."""
        snap = self.snapshot()
        running = snap['running'] or {}
        return {'busy': bool(snap['running']), 'agent': running.get('agent'),
                'tool': running.get('tool'), 'running_s': running.get('running_s'),
                'waiting': len(snap['waiting'])}


def busy_result(info):
    """MCP tools/call result for a refused call (same shape as the bridge's errors)."""
    from core import jsonsafe
    return {'content': [{'type': 'text',
                         'text': jsonsafe.dumps(info, separators=(',', ':'))}],
            'isError': True}
