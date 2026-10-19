# -*- coding: utf-8 -*-
"""Upload audit-log summaries to the Tool Hub (team usage stats).

Reads the day files core/t3lab_audit.py writes (audit-YYYY-MM-DD[.N].jsonl)
from where the last upload stopped, keeps only:

    call_id, tool, source, decision, status, exec_ms,
    changes {added, modified, deleted}, doc_hash (16 hex of the path hash)

and sends them in batches of at most 500 to toolhub_log_events. Never sent:
arguments, args previews, code, element ids, document titles or paths.
Off when signed out; at most one upload a minute; never raises.
Cursor: %APPDATA%\\T3LabAI\\toolhub\\events_cursor.json {file: byte offset}.
"""
from __future__ import unicode_literals

import io
import json
import os
import re
import time

from Intelligence.toolhub import config

CURSOR_FILE = 'events_cursor.json'
BATCH = 500
MIN_INTERVAL_S = 60
_FILE_RE = re.compile(r'^audit-\d{4}-\d{2}-\d{2}(\.\d+)?\.jsonl$')
_last_upload = [0.0]


def summarize(rec):
    """One audit record -> the event row the hub stores, or None."""
    if not isinstance(rec, dict) or rec.get('event') != 'call':
        return None
    if not rec.get('call_id') or not rec.get('tool'):
        return None
    ch = rec.get('changes') or {}
    return {
        'call_id': str(rec['call_id'])[:32],
        'tool': str(rec['tool'])[:64],
        'source': (rec.get('source') or '')[:16],
        'decision': (rec.get('decision') or '')[:16],
        'status': (rec.get('status') or '')[:16],
        'exec_ms': rec.get('exec_ms'),
        'changes': {'added': int(ch.get('added') or 0),
                    'modified': int(ch.get('modified') or 0),
                    'deleted': int(ch.get('deleted') or 0)},
        'doc_hash': (rec.get('doc_path_sha') or '')[:16],
    }


def collect(folder, cursor):
    """(events, new_cursor) for everything after `cursor` in `folder`."""
    events, new = [], dict(cursor or {})
    try:
        names = sorted(n for n in os.listdir(folder) if _FILE_RE.match(n))
    except OSError:
        return [], new
    for name in names:
        p = os.path.join(folder, name)
        start = int(new.get(name) or 0)
        try:
            size = os.path.getsize(p)
            if size <= start:
                continue
            with io.open(p, 'rb') as f:
                f.seek(start)
                chunk = f.read(size - start)
        except (IOError, OSError):
            continue
        # A half-written last line waits for the next upload.
        end = chunk.rfind(b'\n') + 1
        for line in chunk[:end].splitlines():
            try:
                ev = summarize(json.loads(line.decode('utf-8')))
            except (ValueError, UnicodeDecodeError):
                continue
            if ev:
                events.append(ev)
        new[name] = start + end
    return events, new


def upload(client, folder=None, force=False, now=None):
    """Send pending events. Returns the number accepted (0 when skipped)."""
    now = time.time() if now is None else now
    if not force and now - _last_upload[0] < MIN_INTERVAL_S:
        return 0
    _last_upload[0] = now
    try:
        if not client.signed_in():
            return 0
        if folder is None:
            from core import t3lab_audit
            folder = t3lab_audit.audit_dir()
        cursor = config.read_json(CURSOR_FILE, {}) or {}
        events, new_cursor = collect(folder, cursor)
        sent = 0
        for i in range(0, len(events), BATCH):
            res = client.rpc('toolhub_log_events', {'p_events': events[i:i + BATCH]})
            sent += res if isinstance(res, int) else 0
        config.write_json(CURSOR_FILE, new_cursor)
        return sent
    except Exception:
        return 0
