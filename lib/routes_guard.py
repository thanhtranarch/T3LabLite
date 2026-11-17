# -*- coding: utf-8 -*-
"""Keep pyRevit's Routes server off the T3Lab MCP ports (48884-48894).

Routes defaults to port 48884, and early T3Lab builds also wrote its port
into the pyRevit config. With Routes enabled, pyRevit binds 0.0.0.0:<port>
after the extensions load, while T3Lab's MCP server already holds
127.0.0.1:48884: Windows refuses with WinError 10013 and pyRevit prints
"Routes servers failed activation" (owner report, 2026-11-17).

startup.py calls apply() at every Revit start, any pyRevit version (6.x and
7.x tested against their source): when Routes is enabled with a port inside
the T3Lab range it is moved to ROUTES_PORT, through pyRevit's own config API
when present (fixes the current session too: Routes starts after startup
scripts), else by editing pyRevit_config.ini (next start). Routes disabled,
or already outside the range: nothing is touched.

Engine-agnostic (IronPython 2.7 + CPython 3). Test: lite_guard/test_routes_guard.py.
"""
from __future__ import unicode_literals

import io
import os
import re

T3LAB_PORT_MIN, T3LAB_PORT_MAX = 48884, 48894   # = core.server PORT_MIN / PORT_MAX
ROUTES_PORT = 48870         # pyRevit walks up from here per Revit window: 14 before 48884
ROUTES_DEFAULT_PORT = 48884  # pyRevit's own default when the config has no port


def _truthy(value):
    return '{}'.format(value).strip().lower() in ('true', '1', 'yes', 'on')


def new_port(enabled, port):
    """The port Routes should move to, or None when it can stay."""
    if not _truthy(enabled):
        return None
    try:
        port = int(port)
    except (TypeError, ValueError):
        port = ROUTES_DEFAULT_PORT
    return ROUTES_PORT if T3LAB_PORT_MIN <= port <= T3LAB_PORT_MAX else None


_SECTION = re.compile(r'^\s*\[([^\]]+)\]\s*$')
_KEY = re.compile(r'^(\s*)([^=:\s][^=:]*?)(\s*[=:]\s*)(.*?)(\s*)$')


def ini_fix(text):
    """(new text, old port) with [routes] port moved, or (None, None)."""
    lines = text.split('\n')
    start = end = None
    for i, line in enumerate(lines):
        m = _SECTION.match(line)
        if m:
            if start is not None:
                end = i
                break
            if m.group(1).strip().lower() == 'routes':
                start = i
    if start is None:
        return None, None
    end = len(lines) if end is None else end
    values, where = {}, {}
    for i in range(start + 1, end):
        m = _KEY.match(lines[i])
        if m and not lines[i].lstrip().startswith(('#', ';')):
            key = m.group(2).strip().lower()
            values[key] = m.group(4)
            where[key] = i
    port = values.get('port', ROUTES_DEFAULT_PORT)
    target = new_port(values.get('enabled', 'false'), port)
    if target is None:
        return None, None
    if 'port' in where:
        m = _KEY.match(lines[where['port']])
        lines[where['port']] = '{}{}{}{}'.format(m.group(1), m.group(2), m.group(3), target)
    else:
        last = start
        for i in range(start + 1, end):
            if lines[i].strip():
                last = i
        lines.insert(last + 1, 'port = {}'.format(target))
    return '\n'.join(lines), port


def _config_path(user_config=None):
    try:
        if user_config is not None and getattr(user_config, 'config_file', None):
            return user_config.config_file
    except Exception:
        pass
    root = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(root, 'pyRevit', 'pyRevit_config.ini')


def _via_api(user_config):
    """(old, new) when pyRevit's config API moved the port; None when this
    pyRevit has no such API; (None, None) when nothing had to change."""
    for name in ('routes_server', 'routes_port', 'save_changes'):
        if not hasattr(user_config, name):
            return None
    target = new_port(user_config.routes_server, user_config.routes_port)
    if target is None:
        return None, None
    old = user_config.routes_port
    user_config.routes_port = target
    user_config.save_changes()
    return old, target


def apply(user_config=None, path=None):
    """Move Routes off the T3Lab ports. Returns a log line, or '' when nothing
    changed. Never raises."""
    try:
        if user_config is None and path is None:
            try:
                from pyrevit.userconfig import user_config
            except Exception:
                user_config = None
        if user_config is not None:
            try:
                done = _via_api(user_config)
            except Exception:
                done = None
            if done is not None:
                if done[0] is None:
                    return ''
                return ('pyRevit Routes moved from port {} to {} (it collided with the '
                        'T3Lab MCP ports {}-{}).'.format(done[0], done[1], T3LAB_PORT_MIN,
                                                         T3LAB_PORT_MAX))
        path = path or _config_path(user_config)
        if not os.path.isfile(path):
            return ''
        with io.open(path, 'r', encoding='utf-8-sig', newline='') as f:
            raw = f.read()
        newline = '\r\n' if '\r\n' in raw else '\n'
        text, old = ini_fix(raw.replace('\r\n', '\n'))
        if text is None:
            return ''
        with io.open(path, 'w', encoding='utf-8', newline='') as f:
            f.write(text.replace('\n', newline))
        return ('pyRevit Routes port {} moved to {} in {} (it collided with the T3Lab MCP '
                'ports); takes effect at the next Revit start.'.format(old, ROUTES_PORT, path))
    except Exception as ex:
        return 'pyRevit Routes port check failed: {}'.format(ex)
