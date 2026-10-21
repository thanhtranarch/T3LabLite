# -*- coding: utf-8 -*-
"""Tool Hub connection settings and local folder.

The publishable key is public by design (it ships in every Supabase web
client): what protects the data is RLS on schema `toolhub` (no grants at all)
and the role checks inside each public.toolhub_* RPC. A machine can point at
another project with %APPDATA%\\T3LabAI\\toolhub\\settings.json:

    {"url": "https://<ref>.supabase.co", "key": "sb_publishable_..."}
"""
from __future__ import unicode_literals

import io
import json
import os

DEFAULT_URL = 'https://irbfdprcoopoatiusows.supabase.co'
DEFAULT_KEY = 'sb_publishable_mXNYQq-g0xoGURMakUYeeQ_ChxU_eiz'

HTTP_TIMEOUT = 8.0


def toolhub_dir():
    """%APPDATA%\\T3LabAI\\toolhub, created on first use."""
    base = os.environ.get('APPDATA') or os.path.join(os.path.expanduser('~'), '.t3lab')
    d = os.path.join(base, 'T3LabAI', 'toolhub')
    try:
        if not os.path.isdir(d):
            os.makedirs(d)
    except OSError:
        pass
    return d


def path(name):
    return os.path.join(toolhub_dir(), name)


def read_json(name, default=None):
    p = path(name)
    if not os.path.isfile(p):
        return default
    try:
        with io.open(p, encoding='utf-8') as f:
            return json.load(f)
    except (IOError, OSError, ValueError):
        return default


def write_json(name, data):
    """Write through a temp file so a crash never leaves half a file."""
    p = path(name)
    tmp = p + '.tmp'
    text = json.dumps(data, ensure_ascii=False, indent=1, sort_keys=True)
    if isinstance(text, bytes):
        text = text.decode('utf-8')
    with io.open(tmp, 'w', encoding='utf-8') as f:
        f.write(text)
    if os.path.exists(p):
        os.remove(p)
    os.rename(tmp, p)


def endpoint():
    """(url, publishable key) — the override file wins over the defaults."""
    over = read_json('settings.json', {}) or {}
    url = (over.get('url') or DEFAULT_URL).rstrip('/')
    key = over.get('key') or DEFAULT_KEY
    return url, key


def titles_file():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tool_titles_vi.json')
