# -*- coding: utf-8 -*-
"""
Remote MCP servers for the T3Lab Assistant (Claude only).

The Assistant can use tools from MCP servers on the internet -- a company
knowledge base, Notion, Linear, an ACC/Forma bridge -- through Claude's MCP
connector: the request names the servers (`mcp_servers`) and one
`mcp_toolset` per server, and Anthropic's side connects to them, lists their
tools and runs the calls inside the same response (mcp_tool_use /
mcp_tool_result blocks). No MCP client runs in Revit, so nothing here needs a
background thread (see CLAUDE.md rule 8).

Limits of the connector (platform.claude.com, "MCP connector"):
  - the server must be reachable from the internet over https (Streamable
    HTTP or SSE); a local stdio server cannot be used;
  - only MCP tools, no resources or prompts;
  - Claude API only: other providers ignore these servers.

The calls run on Anthropic's side, so the Assistant's "Ask before edits" gate
(agent_loop.ToolGate) never sees them -- the settings UI says so, and only
servers the user added by hand are ever sent.

Stored in settings.json under agents.remote_mcp_servers:
    [{"name": "notion", "url": "https://...", "enabled": true,
      "token": "dpapi:..."}]
The token is protected with DPAPI like the API keys (config.secret_store).

Plain Python, IronPython 2.7 + CPython 3, no Revit import.

Author: Tran Tien Thanh
"""

from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__ = "Remote MCP Servers"

import re

SETTING_KEY = 'remote_mcp_servers'
BETA_HEADER = 'mcp-client-2025-11-20'
MAX_SERVERS = 8

_NAME_RE = re.compile(r'^[A-Za-z0-9_-]{1,64}$')


def _settings():
    from config.settings import get_settings
    return get_settings()


def _secrets():
    from config import secret_store
    return secret_store


def clean_name(text):
    """'Company Wiki' -> 'company-wiki' (the API's server name rules)."""
    name = re.sub(r'[^A-Za-z0-9_-]+', '-', u'{}'.format(text or u'').strip())
    return name.strip('-')[:64].lower()


def problem(server, others=()):
    """Why `server` cannot be used, or None. `others` are the names already
    taken by the other servers."""
    name = server.get('name') or u''
    url = (server.get('url') or u'').strip()
    if not _NAME_RE.match(name):
        return u'Name: letters, digits, - and _ only (1-64 characters).'
    if name in others:
        return u'Another server already uses the name "{}".'.format(name)
    if not url.lower().startswith(u'https://') or len(url) <= len(u'https://'):
        return (u'The URL must start with https:// -- Claude connects to the '
                u'server from the internet, so a local (stdio) server cannot '
                u'be used.')
    return None


def load(settings=None):
    """The stored servers, token NOT decrypted (has_token says whether one is
    set). Never raises: a broken entry is skipped."""
    try:
        raw = (settings or _settings()).get_agent_option(SETTING_KEY, []) or []
    except Exception:
        return []
    out = []
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict):
            continue
        out.append({'name': u'{}'.format(item.get('name') or u''),
                    'url': u'{}'.format(item.get('url') or u'').strip(),
                    'enabled': item.get('enabled', True) is not False,
                    'has_token': bool(item.get('token'))})
    return out


def save(servers, settings=None):
    """Replace the stored list. Each entry: name, url, enabled and, to change
    the token, `token` (u'' removes it; a missing key keeps the stored one,
    also across a rename when `previous_name` names the old entry).
    Returns (saved list as load() shows it, error or None); nothing is
    written when any entry has a problem."""
    settings = settings or _settings()
    try:
        stored = dict((s.get('name'), s) for s in
                      (settings.get_agent_option(SETTING_KEY, []) or [])
                      if isinstance(s, dict))
    except Exception:
        stored = {}
    if len(servers) > MAX_SERVERS:
        return load(settings), u'At most {} servers.'.format(MAX_SERVERS)
    rows, names = [], []
    for server in servers:
        row = {'name': clean_name(server.get('name')),
               'url': (server.get('url') or u'').strip(),
               'enabled': server.get('enabled', True) is not False}
        err = problem(row, names)
        if err:
            return load(settings), u'{}: {}'.format(row['name'] or u'?', err)
        names.append(row['name'])
        if 'token' in server:
            token = (server.get('token') or u'').strip()
            if token:
                row['token'] = _secrets().protect(token)
        else:
            old = (stored.get(row['name'])
                   or stored.get(server.get('previous_name')))
            if old and old.get('token'):
                row['token'] = old['token']
        rows.append(row)
    settings.set_agent_option(SETTING_KEY, rows)
    return load(settings), None


def request_parts(settings=None):
    """(mcp_servers, toolsets) for a Claude Messages API request: every
    enabled, valid server, with its token decrypted. ([], []) when none."""
    try:
        raw = (settings or _settings()).get_agent_option(SETTING_KEY, []) or []
    except Exception:
        return [], []
    servers, toolsets, names = [], [], []
    for item in raw if isinstance(raw, list) else []:
        if not isinstance(item, dict) or item.get('enabled', True) is False:
            continue
        if problem(item, names):
            continue
        names.append(item['name'])
        entry = {'type': 'url', 'url': item['url'].strip(), 'name': item['name']}
        token = item.get('token')
        if token:
            try:
                secrets = _secrets()
                token = (secrets.unprotect(token) if secrets.is_protected(token)
                         else token)
            except Exception:
                token = None
            if token:
                entry['authorization_token'] = token
        servers.append(entry)
        toolsets.append({'type': 'mcp_toolset', 'mcp_server_name': item['name']})
    return servers, toolsets
