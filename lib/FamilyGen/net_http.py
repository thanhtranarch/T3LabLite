# -*- coding: utf-8 -*-
"""Asynchronous HTTPS POST for FamiGen Studio, without a Python thread.

Why not a Python worker thread: under pyRevit's CPython engine Revit's main
thread holds the GIL whenever no CPython code runs (CLAUDE.md rule 8), so a
CPython thread started by a modeless window freezes as soon as the command
returns. The request therefore runs entirely in .NET (``HttpClient.SendAsync``
on the thread pool, no Python callback), and a ``DispatcherTimer`` on the UI
thread polls the Task every ``POLL_MS``. The timer tick is the only Python
code, and it runs on the thread that already holds the GIL - the same pattern
as the T3Lab Assistant pane's GIL pump.

Retries 408/409/429/5xx/529 like the official SDKs (``retry-after`` honoured,
at most ``MAX_RETRIES``). One request at a time per poster; ``cancel()``
aborts it and the callback is never called.
"""
import time

import clr

clr.AddReference('System')
clr.AddReference('WindowsBase')
try:
    clr.AddReference('System.Net.Http')          # .NET Framework 4.8 (Revit 2022-2024)
except Exception:
    pass                                          # part of the shared framework on .NET 8

from System import Convert, TimeSpan                                  # noqa: E402
from System.Net.Http import (HttpClient, HttpMethod, HttpRequestMessage,  # noqa: E402
                             StringContent)
from System.Text import Encoding                                      # noqa: E402
from System.Threading import CancellationTokenSource                  # noqa: E402
from System.Windows.Threading import DispatcherTimer                  # noqa: E402

POLL_MS = 100
TIMEOUT_S = 600            # a long family answer can stream for minutes
MAX_RETRIES = 2
RETRY_STATUSES = (408, 409, 429, 500, 502, 503, 504, 529)

_CLIENT = {'client': None}


def _tls12():
    try:
        from System.Net import ServicePointManager, SecurityProtocolType
        current = ServicePointManager.SecurityProtocol
        ServicePointManager.SecurityProtocol = current | SecurityProtocolType.Tls12
    except Exception:
        pass


def _client():
    if _CLIENT['client'] is None:
        _tls12()
        client = HttpClient()
        client.Timeout = TimeSpan.FromSeconds(TIMEOUT_S + 30)
        _CLIENT['client'] = client
    return _CLIENT['client']


def _task_error(task):
    try:
        ex = task.Exception
        inner = ex.InnerException if ex is not None else None
        while inner is not None and inner.InnerException is not None:
            inner = inner.InnerException
        return u'{}'.format((inner or ex).Message)
    except Exception:
        return u'request failed'


class DotNetPoster(object):
    """``start(url, headers, body_text, callback)``; ``callback(status, raw)``.

    ``status`` is the HTTP status, or 0 for a transport error (``raw`` is then
    the error text). Must be created and used on the UI thread.
    """

    def __init__(self, poll_ms=POLL_MS, timeout_s=TIMEOUT_S, max_retries=MAX_RETRIES):
        self._timeout_s = timeout_s
        self._max_retries = max_retries
        self._job = None
        self._timer = DispatcherTimer()
        self._timer.Interval = TimeSpan.FromMilliseconds(poll_ms)
        self._timer.Tick += self._tick

    # ── public ───────────────────────────────────────────────────────────
    def start(self, url, headers, body_text, callback):
        self.cancel()
        self._job = {'url': url, 'headers': dict(headers or {}), 'body': body_text,
                     'callback': callback, 'attempt': 0, 'stage': None, 'task': None,
                     'cts': None, 'status': 0, 'retry_at': 0.0, 'deadline': 0.0}
        self._send()
        self._timer.Start()

    def cancel(self):
        job, self._job = self._job, None
        self._timer.Stop()
        if job and job.get('cts') is not None:
            try:
                job['cts'].Cancel()
            except Exception:
                pass

    def dispose(self):
        self.cancel()
        try:
            self._timer.Tick -= self._tick
        except Exception:
            pass

    @property
    def busy(self):
        return self._job is not None

    # ── internals ────────────────────────────────────────────────────────
    def _send(self):
        job = self._job
        request = HttpRequestMessage(HttpMethod.Post, job['url'])
        request.Content = StringContent(job['body'], Encoding.UTF8, 'application/json')
        for key, value in job['headers'].items():
            if key.lower() == 'content-type':
                continue
            request.Headers.TryAddWithoutValidation(key, u'{}'.format(value))
        job['cts'] = CancellationTokenSource()
        job['task'] = _client().SendAsync(request, job['cts'].Token)
        job['stage'] = 'send'
        job['deadline'] = time.time() + self._timeout_s

    def _finish(self, status, raw):
        job, self._job = self._job, None
        self._timer.Stop()
        if job is not None:
            job['callback'](status, raw)

    def _retry_or_finish(self, status, raw, retry_after=None):
        job = self._job
        if status in RETRY_STATUSES + (0,) and job['attempt'] < self._max_retries:
            job['attempt'] += 1
            delay = retry_after if retry_after is not None else 2.0 ** job['attempt']
            job['retry_at'] = time.time() + min(max(delay, 1.0), 60.0)
            job['stage'] = 'wait'
            return
        self._finish(status, raw)

    def _tick(self, sender, args):
        job = self._job
        if job is None:
            self._timer.Stop()
            return
        try:
            if job['stage'] == 'wait':
                if time.time() >= job['retry_at']:
                    self._send()
                return
            task = job['task']
            if not task.IsCompleted:
                if time.time() > job['deadline']:
                    try:
                        job['cts'].Cancel()
                    except Exception:
                        pass
                    self._finish(0, u'no answer after {} s'.format(self._timeout_s))
                return
            if job['stage'] == 'send':
                if task.IsFaulted or task.IsCanceled:
                    self._retry_or_finish(0, _task_error(task))
                    return
                response = task.Result
                job['status'] = Convert.ToInt32(response.StatusCode)
                job['retry_after'] = None
                try:
                    delta = response.Headers.RetryAfter.Delta if response.Headers.RetryAfter else None
                    if delta is not None:
                        job['retry_after'] = float(delta.TotalSeconds)
                except Exception:
                    pass
                job['task'] = response.Content.ReadAsStringAsync()
                job['stage'] = 'read'
                return
            if job['stage'] == 'read':
                if task.IsFaulted or task.IsCanceled:
                    self._retry_or_finish(0, _task_error(task))
                    return
                raw = task.Result
                status = job['status']
                if status == 200:
                    self._finish(status, raw)
                else:
                    self._retry_or_finish(status, raw, job.get('retry_after'))
        except Exception as ex:          # never leave the timer spinning
            self._finish(0, u'{}'.format(ex))
