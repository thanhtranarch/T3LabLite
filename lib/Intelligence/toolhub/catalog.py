# -*- coding: utf-8 -*-
"""Local cache of the Tool Hub catalog: %APPDATA%\\T3LabAI\\toolhub\\catalog.json.

    {"rows": {slug: row}, "last_updated": "<max updated_at seen>"}

sync() asks toolhub_catalog(p_since=last_updated) for the delta and merges it
by updated_at, so a sync after the first one moves a few rows, not all 156.
local_rows() builds rows for the tools this machine already has (the MCP
registry the caller passes, plus the curated Vietnamese titles), so search
works before anyone signs in.
"""
from __future__ import unicode_literals

import io
import json

from Intelligence.toolhub import config

CACHE_FILE = 'catalog.json'


def load():
    data = config.read_json(CACHE_FILE, {}) or {}
    return {'rows': data.get('rows') or {}, 'last_updated': data.get('last_updated')}


def save(cache):
    config.write_json(CACHE_FILE, cache)


def merge(cache, rows):
    """Apply a delta. A row replaces the cached one when it is not older.
    Returns the number of rows that changed."""
    changed = 0
    for row in rows or []:
        if isinstance(row, str):
            try:
                row = json.loads(row)
            except ValueError:
                continue
        slug = row.get('slug') if isinstance(row, dict) else None
        if not slug:
            continue
        cur = cache['rows'].get(slug)
        stamp = row.get('updated_at') or ''
        if cur is None or stamp >= (cur.get('updated_at') or ''):
            if cur != row:
                changed += 1
            cache['rows'][slug] = row
        if stamp and stamp > (cache.get('last_updated') or ''):
            cache['last_updated'] = stamp
    return changed


def sync(client):
    """Pull the delta from the hub into the cache. Returns (changed, total)."""
    cache = load()
    rows = client.rpc('toolhub_catalog', {'p_since': cache.get('last_updated')})
    changed = merge(cache, rows if isinstance(rows, list) else [])
    save(cache)
    return changed, len(cache['rows'])


def titles():
    try:
        with io.open(config.titles_file(), encoding='utf-8') as f:
            data = json.load(f)
    except (IOError, OSError, ValueError):
        return {}
    return {k: v for k, v in data.items() if not k.startswith('_')}


def local_rows(tools=None):
    """Rows for tools known on this machine.

    tools: [{"name", "description"}] (Anthropic or MCP shape); the registry
    the caller already has. Curated title_vi / keywords are added from
    tool_titles_vi.json.
    """
    t = titles()
    if tools is None:
        # No registry at hand (a slash command, the MCP server before its
        # registry loads): the curated titles still name ~100 tools.
        tools = [{'name': name, 'description': ''} for name in t]
    rows = []
    for tool in tools:
        name = tool.get('name')
        if not name:
            continue
        cur = t.get(name, {})
        rows.append({
            'slug': name, 'kind': 'mcp_tool', 'name': name,
            'title_vi': cur.get('title_vi', ''),
            'description': tool.get('description') or '',
            'keywords': cur.get('keywords', []),
            'status': 'published', 'local': True,
        })
    return rows


def all_rows(tools=None):
    """Hub rows (published and own drafts) over the local ones, by slug."""
    out = {r['slug']: r for r in local_rows(tools)}
    for slug, row in load()['rows'].items():
        if row.get('status') == 'deprecated':
            out.pop(slug, None)
            continue
        out[slug] = row
    return list(out.values())
