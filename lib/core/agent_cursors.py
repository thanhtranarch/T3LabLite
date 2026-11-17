# -*- coding: utf-8 -*-
"""Agent cursors: a live overlay of which AI agent does what, where.

Each MCP tool call moves one cursor per (agent, action) — Claude · Create,
Claude · Modify, GPT · Read … — to each element it read, edited, created or
deleted in turn (move, click, next), in every visible view, with a plain
caption ("Moving elements · 3 elements"), a badge per element ("Walls ·
1200 × 300") and an effect per action; a call with no element scans the view.
The HUD lists the last steps. Follow mode opens the views the agent works in; zoom brings off-screen
targets into view. Idea after swarm-cursors (github.com/DhvanilPatel/
swarm-cursors); drawn here in code-built WPF, no code from it.

before() / notify() run on Revit's main thread from
MCPExternalEventHandler.Execute (core/server.py): before() just ahead of the
tool (boxes of what a delete / move is about to change), notify() after the
caller has its result. Off by default; Settings > Agent Cursors or the HUD.
Engine-agnostic (IronPython 2.7 + CPython 3); the pure helpers are tested by
lite_guard/test_agent_cursors.py.
"""

from __future__ import unicode_literals

import numbers
import os
import re

SETTING_KEY = 'agent_cursors_enabled'
MAX_TARGETS = 12            # bounding boxes per call: keeps a bulk edit cheap
MAX_FOLLOW_VIEWS = 4        # views follow mode may open per call
IDLE_HIDE_S = 12            # overlay fades out after this long without a call
BOX_HOLD_S = 3.5            # element boxes stay this long, then fade
MOVE_MS = 650               # cursor glide (one stop; shorter per stop on a tour)
MAX_STOPS = 6               # targets the cursor visits one by one per call
DWELL_MS = 220              # pause + click at each stop
FEED_SIZE = 5               # last steps listed on the HUD
REPLAY_SIZE = 20            # steps the HUD can replay
SUMMARY_S = 20              # idle summary card stays this long
NARRATE_MAX = 90            # goal banner text

# name -> (mcp_paths.json key, default). The HUD toggles them live.
OPTIONS = (
    ('follow', 'agent_cursors_follow_views', True),
    ('zoom', 'agent_cursors_zoom', True),
    ('select', 'agent_cursors_select', False),
    ('tile', 'agent_cursors_tile_views', False),
    ('reduce', 'agent_cursors_reduce_motion', False),
    ('glow', 'agent_cursors_glow', False),
)
# Motion speed (mcp_paths.json agent_cursors_speed): name -> time factor.
SPEED_KEY = 'agent_cursors_speed'
SPEEDS = (('calm', 1.6), ('normal', 1.0), ('fast', 0.55))
OPTION_LABELS = {
    'follow': ('Follow', 'Open and switch to the views the agent works in'),
    'zoom': ('Zoom', 'Zoom the view to targets that are off screen'),
    'select': ('Select', "Select the agent's targets (replaces your selection)"),
    'tile': ('Tile', 'Tile the open views when the agent opens another one'),
    'reduce': ('Less motion', 'No glides, ripples or scans: boxes and captions only'),
    'glow': ('Glow', 'Experimental: outline touched elements inside the Revit view '
                     '(DirectContext3D). Turn off if Revit misbehaves.'),
}

# agent id (bridge `t3lab/agent`) -> (tag, colour). Unknown agents get a
# stable colour from _SPARE by name.
AGENTS = {
    'claude-desktop':  ('Claude', '#D97757'),
    'claude-code':     ('Claude Code', '#B4532A'),
    'codex':           ('GPT', '#10A37F'),
    'chatgpt':         ('GPT', '#10A37F'),
    'antigravity':     ('Gemini', '#4285F4'),
    'cursor':          ('Cursor', '#7C3AED'),
    'cli':             ('CLI', '#64748B'),
    't3lab-assistant': ('T3Lab Assistant', '#EA680C'),
}
# MCP clientInfo names -> AGENTS key.
_ALIASES = {
    'claude-ai': 'claude-desktop', 'claude': 'claude-desktop',
    'codex-mcp-client': 'codex', 'openai-codex': 'codex',
    'gemini-cli-mcp-client': 'antigravity', 'gemini': 'antigravity',
    'cursor-vscode': 'cursor',
}
_SPARE = ('#E11D48', '#0EA5E9', '#22C55E', '#F59E0B', '#8B5CF6', '#14B8A6')
_FT_TO_MM = 304.8
# Views with no model geometry to box.
_NO_BOXES = ('Schedule', 'ColumnSchedule', 'PanelSchedule', 'Report', 'CostReport',
             'LoadsReport', 'PresureLossReport', 'PressureLossReport', 'Internal',
             'ProjectBrowser', 'SystemBrowser', 'Undefined')

# action -> (label, Segoe MDL2 glyph, colour). Order = HUD order.
ACTIONS = (
    ('read',   'Read',   '', '#64748B'),
    ('create', 'Create', '', '#16A34A'),
    ('modify', 'Modify', '', '#2563EB'),
    ('delete', 'Delete', '', '#DC2626'),
    ('view',   'View',   '', '#7C3AED'),
    ('export', 'Export', '', '#0891B2'),
    ('code',   'Code',   '', '#D97706'),
)
_ACTION = dict((a[0], a) for a in ACTIONS)

PROS = (
    "See at a glance which AI app is working, what it does and on which elements.",
    "Plain-English captions, the agent's goal (t3lab_narrate) and a feed of the last steps on the bar.",
    "The cursor moves to each target and clicks it; values it changed show as \"Mark: A-101 → A-102\".",
    "Long calls show \"Working 8 s\"; a call waiting for your OK pulses amber.",
    "Two agents editing the same element flash red; targets off screen get an edge arrow.",
    "When the agents go idle: a summary, Select touched and Replay.",
)
CONS = (
    "Each tool call does a little extra work (bounding boxes of up to {} elements, twice for a delete or move).".format(MAX_TARGETS),
    "Follow switches the active view, so a later tool that works on 'the active view' works on the view the agent opened.",
    "Zoom and Select change your camera and selection while an agent works (Select is off by default).",
    "Boxes are a snapshot: pan or zoom while an agent works and they sit off until its next call; Replay uses the old positions.",
    "Glow (off by default) is experimental: it draws inside Revit's 3D views and is not tested on every Revit version.",
    "The overlay shows up in screenshots and recordings.",
)


def pros_cons_text():
    """Benefits and drawbacks, as shown when the feature is toggled."""
    lines = ['Benefits']
    lines += ['  + ' + p for p in PROS]
    lines += ['', 'Drawbacks']
    lines += ['  - ' + c for c in CONS]
    return '\n'.join(lines)


# ─── Settings ──────────────────────────────────────────────────────────────────

_cache = {'mtime': None, 'data': {}}


def _paths():
    from core import paths
    return paths


def _settings():
    """mcp_paths.json, re-read only when the file changes."""
    try:
        paths = _paths()
        mtime = os.path.getmtime(paths.settings_file())
    except Exception:
        return {}
    if mtime != _cache['mtime']:
        try:
            _cache['data'] = dict(paths.load_settings() or {})
        except Exception:
            _cache['data'] = {}
        _cache['mtime'] = mtime
    return _cache['data']


def is_enabled():
    return bool(_settings().get(SETTING_KEY, False))


def set_enabled(on):
    """Persist the toggle; the next tool call picks it up."""
    _paths().set_setting(SETTING_KEY, bool(on))
    _cache['mtime'] = None
    if not on:
        hide()
        _session['touched'].clear()     # the next session starts clean
        _session['goals'].clear()
    return bool(on)


def options():
    """{'follow': bool, 'zoom': bool, 'select': bool, 'tile': bool}."""
    data = _settings()
    return dict((name, bool(data.get(key, default))) for name, key, default in OPTIONS)


def set_option(name, value):
    for opt, key, _default in OPTIONS:
        if opt == name:
            _paths().set_setting(key, bool(value))
            _cache['mtime'] = None
            return bool(value)
    raise KeyError(name)


def speed():
    """'calm' | 'normal' | 'fast'."""
    name = _settings().get(SPEED_KEY, 'normal')
    return name if name in dict(SPEEDS) else 'normal'


def speed_factor(name=None):
    return dict(SPEEDS)[name or speed()]


def next_speed(name):
    names = [n for n, _f in SPEEDS]
    return names[(names.index(name) + 1) % len(names)] if name in names else 'normal'


def set_speed(name):
    _paths().set_setting(SPEED_KEY, name if name in dict(SPEEDS) else 'normal')
    _cache['mtime'] = None
    return speed()


def options_text():
    """Follow on · Zoom on · Select off · Tile off."""
    opts = options()
    return ' · '.join('{} {}'.format(OPTION_LABELS[n][0], 'on' if opts[n] else 'off')
                      for n, _k, _d in OPTIONS)


# ─── Pure helpers ──────────────────────────────────────────────────────────────

def agent_key(agent):
    key = (agent or '').strip().lower()
    return _ALIASES.get(key, key)


def agent_style(agent):
    """(tag, '#RRGGBB') for a bridge agent id or MCP client name."""
    key = agent_key(agent)
    if key in AGENTS:
        return AGENTS[key]
    if not key:
        return AGENTS['claude-desktop']
    total = 0
    for ch in key:
        total = (total * 31 + ord(ch)) % 1000003
    return (agent.strip()[:18], _SPARE[total % len(_SPARE)])


def action_style(action):
    """(label, glyph, colour) of an ACTIONS key; unknown -> modify."""
    return _ACTION.get(action, _ACTION['modify'])[1:]


# First word of the tool name or its `operation` -> action; first hit wins.
_VERBS = (
    ('code', ('send_code', 'execute', 'run', 'eval', 'script')),
    ('delete', ('delete', 'remove', 'purge', 'clear', 'unload', 'erase')),
    ('export', ('export', 'print', 'save', 'sync', 'publish', 'synchronize',
                'batch')),
    ('view', ('set_active_view', 'open', 'activate', 'zoom', 'show', 'select',
              'isolate', 'hide', 'unhide', 'switch', 'navigate', 'reset_temporary',
              'halftone', 'unhalftone', 'transparency', 'override', 'color',
              'colour', 'reset_color', 'focus', 'close')),
    ('create', ('create', 'place', 'add', 'new', 'duplicate', 'copy', 'draw',
                'generate', 'load', 'tag', 'make', 'insert', 'room_to_floor',
                'build')),
    ('read', ('get', 'list', 'find', 'query', 'search', 'check', 'analyze',
              'analyse', 'count', 'describe', 'audit', 'detect', 'read', 'inspect',
              'say_hello', 'propose', 'preview', 'measure', 'status', 'wiki',
              'memory', 'capabilities', 'help', 'ai_element', 'file_watcher')),
    ('modify', ('set', 'update', 'modify', 'move', 'rotate', 'rename', 'change',
                'apply', 'edit', 'align', 'join', 'split', 'mirror', 'bulk', 'pin',
                'unpin', 'group', 'ungroup', 'flip', 'offset', 'replace', 'swap',
                'attach', 'detach', 'renumber', 'resize', 'store')),
)
_PREFIXES = ('revit_', 't3lab_', 'famigen_', 'mcp_')


def _verb_action(text):
    text = (text or '').strip().lower()
    for prefix in _PREFIXES:
        if text.startswith(prefix):
            text = text[len(prefix):]
    if not text:
        return None
    for action, verbs in _VERBS:
        for verb in verbs:
            if text == verb or text.startswith(verb + '_') or (
                    '_' in verb and text.startswith(verb)):
                return action
    return None


def classify_action(tool_name, arguments=None, result=None):
    """ACTIONS key for one tool call: its `operation` / `action` argument
    first (manage_view set_scale, operate_element hide …), then the tool name."""
    name = '{}'.format(tool_name or '').strip().lower()
    if name == 'send_code_to_revit':
        return 'code'
    if isinstance(arguments, dict):
        for key in ('operation', 'action', 'mode', 'op'):
            value = arguments.get(key)
            if isinstance(value, (str, type(''))) and value:
                action = _verb_action(value)
                if action:
                    return action
    action = _verb_action(name)
    if action:
        return action
    if isinstance(result, dict) and (result.get('created_ids') or result.get('new_element_id')):
        return 'create'
    try:
        from Intelligence.tool_schema import READ_ONLY_TOOL_NAMES
        if name in READ_ONLY_TOOL_NAMES:
            return 'read'
    except Exception:
        pass
    return 'modify'


# element_id, element_ids, ids, id, created_ids, new_element_id, host_id ...
_ID_KEY = re.compile(r'(^|_)(element_?)?ids?$|^(elementid|elementids)$', re.I)
# Ids of things that are not model geometry worth boxing.
_NOT_ELEMENT = re.compile(r'^(new_|created_|target_|active_|source_)?(view|sheet|type|level|family|symbol|workset|category|phase|'
                          r'param|doc|session|approval|job|request|tool|template|material|'
                          r'schedule|revision|link)', re.I)
# view_id, view_ids, sheet_id, new_view_id, created_view_ids, target_view_id ...
_VIEW_KEY = re.compile(r'^(new_|created_|target_|active_|source_)?(view|sheet)_?ids?$', re.I)


def _as_int(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, numbers.Integral):
        return int(value)
    try:
        if isinstance(value, float) and value == int(value):
            return int(value)
        text = '{}'.format(value).strip()
        if re.match(r'^-?\d+$', text):
            return int(text)
    except Exception:
        pass
    return None


def _walk_ids(payloads, match, limit, max_depth, bare_id_lists):
    out = []

    def add(v):
        i = _as_int(v)
        if i is not None and i > 0 and i not in out:
            out.append(i)

    def walk(node, depth, in_list):
        if len(out) >= limit or depth > max_depth:
            return
        if isinstance(node, dict):
            for key, value in node.items():
                k = '{}'.format(key)
                bare = k.lower() == 'id'
                if (bare and in_list) or (not bare and match(k)):
                    if isinstance(value, (list, tuple)):
                        for v in value:
                            add(v)
                    else:
                        add(value)
                elif isinstance(value, (dict, list, tuple)):
                    walk(value, depth + 1, k.lower() in bare_id_lists)
        elif isinstance(node, (list, tuple)):
            for item in node:
                walk(item, depth + 1, in_list)

    for payload in payloads:
        walk(payload, 0, False)
    return out[:limit]


def extract_ids(*payloads, **kwargs):
    """Element ids named in tool arguments / results, in order, unique.

    Reads keys like element_id / element_ids / created_ids at any depth (3),
    plus `id` inside an `elements` list. `limit` caps the count.
    """
    return _walk_ids(
        payloads, lambda k: bool(_ID_KEY.search(k)) and not _NOT_ELEMENT.match(k),
        kwargs.get('limit', MAX_TARGETS), 3, ('elements', 'created', 'items'))


def extract_view_ids(arguments, result=None, action=None, limit=MAX_FOLLOW_VIEWS):
    """View / sheet ids the call works in: from its arguments always, from its
    result only for a create / view call (a list_views result is not 'work')."""
    payloads = [arguments]
    if action in ('create', 'view'):
        payloads.append(result)
    return _walk_ids(payloads, lambda k: bool(_VIEW_KEY.match(k)), limit, 2,
                     ('views', 'sheets'))


def project(point, origin, right, up, model_w, model_h, rect):
    """Model point -> screen (x, y) for an orthographic view.

    point/origin/right/up are (x, y, z); origin is the zoom-corner minimum,
    model_w/model_h the visible extents along right/up, rect = (l, t, r, b).
    """
    d = (point[0] - origin[0], point[1] - origin[1], point[2] - origin[2])
    u = d[0] * right[0] + d[1] * right[1] + d[2] * right[2]
    v = d[0] * up[0] + d[1] * up[1] + d[2] * up[2]
    left, top, r, bottom = rect
    x = left + (u / model_w) * (r - left)
    y = bottom - (v / model_h) * (bottom - top)
    return x, y


def screen_box(corners, origin, right, up, model_w, model_h, rect):
    """Screen bounds of 3D box corners, clipped to rect; None when outside."""
    pts = [project(c, origin, right, up, model_w, model_h, rect) for c in corners]
    x0 = min(p[0] for p in pts)
    x1 = max(p[0] for p in pts)
    y0 = min(p[1] for p in pts)
    y1 = max(p[1] for p in pts)
    left, top, r, bottom = rect
    if x1 < left or x0 > r or y1 < top or y0 > bottom:
        return None
    return (max(x0, left), max(y0, top), min(x1, r), min(y1, bottom))


def size_label(width_ft, height_ft):
    """'1200 × 300' in mm, as the box badge."""
    return '{} × {}'.format(int(round(width_ft * _FT_TO_MM)),
                                 int(round(height_ft * _FT_TO_MM)))


def _overlaps(a, b):
    return a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]


def visible_views(rects, active):
    """Indexes of the open views a user can see: the active one plus every
    view whose rectangle does not overlap it (tiled). Tabbed views share the
    active view's rectangle, so they are left out."""
    if active is None or not (0 <= active < len(rects)):
        return []
    out = [active]
    for i, rect in enumerate(rects):
        if i == active or rect[2] - rect[0] < 40 or rect[3] - rect[1] < 40:
            continue
        if not _overlaps(rect, rects[active]) and all(
                not _overlaps(rect, rects[j]) for j in out):
            out.append(i)
    return out


def union_rect(rects):
    return (min(r[0] for r in rects), min(r[1] for r in rects),
            max(r[2] for r in rects), max(r[3] for r in rects))


def moved(old, new, tolerance=3.0):
    """True when a screen box moved or resized by more than `tolerance` px."""
    return any(abs(a - b) > tolerance for a, b in zip(old, new))


# First word of a tool / operation -> what the caption says the agent is doing.
_GERUNDS = {
    'get': 'Reading', 'read': 'Reading', 'list': 'Listing', 'find': 'Finding',
    'query': 'Querying', 'search': 'Searching', 'check': 'Checking',
    'analyze': 'Analyzing', 'analyse': 'Analyzing', 'count': 'Counting',
    'describe': 'Describing', 'audit': 'Auditing', 'detect': 'Detecting',
    'inspect': 'Inspecting', 'propose': 'Proposing', 'preview': 'Previewing',
    'create': 'Creating', 'place': 'Placing', 'add': 'Adding', 'new': 'Creating',
    'duplicate': 'Duplicating', 'copy': 'Copying', 'draw': 'Drawing',
    'generate': 'Generating', 'load': 'Loading', 'tag': 'Tagging', 'make': 'Making',
    'insert': 'Inserting', 'build': 'Building', 'set': 'Setting', 'update': 'Updating',
    'modify': 'Modifying', 'move': 'Moving', 'rotate': 'Rotating', 'rename': 'Renaming',
    'change': 'Changing', 'apply': 'Applying', 'edit': 'Editing', 'align': 'Aligning',
    'join': 'Joining', 'split': 'Splitting', 'mirror': 'Mirroring', 'pin': 'Pinning',
    'unpin': 'Unpinning', 'group': 'Grouping', 'ungroup': 'Ungrouping',
    'store': 'Storing', 'delete': 'Deleting', 'remove': 'Removing', 'purge': 'Purging',
    'clear': 'Clearing', 'export': 'Exporting', 'print': 'Printing', 'save': 'Saving',
    'sync': 'Syncing', 'open': 'Opening', 'close': 'Closing', 'activate': 'Activating',
    'zoom': 'Zooming to', 'show': 'Showing', 'select': 'Selecting', 'isolate': 'Isolating',
    'hide': 'Hiding', 'unhide': 'Unhiding', 'switch': 'Switching', 'color': 'Coloring',
    'colour': 'Colouring', 'override': 'Overriding', 'reset': 'Resetting',
    'execute': 'Running', 'run': 'Running', 'say': 'Saying', 'manage': 'Managing',
    'operate': 'Changing', 'bulk': 'Bulk', 'halftone': 'Halftoning',
}
_CAPTION_MAX = 46


def describe_call(tool_name, arguments=None, n_targets=0, failed=False):
    """'Listing open documents', 'Moving elements · 3 elements', 'Failed: …'."""
    name = '{}'.format(tool_name or '').strip().lower()
    if name == 'send_code_to_revit':
        text = 'Running code in Revit'
    else:
        op = arguments.get('operation') if isinstance(arguments, dict) else None
        words = []
        if isinstance(op, (str, type(''))) and op.strip():
            words = op.strip().lower().replace('-', '_').split('_') + ['elements']
        else:
            for prefix in _PREFIXES:
                if name.startswith(prefix):
                    name = name[len(prefix):]
            words = [w for w in name.split('_') if w]
        if not words:
            return 'Working'
        head = words[0]
        verb = _GERUNDS.get(head)
        if verb is None:
            verb = (head[:-1] if head.endswith('e') and not head.endswith('ee') else head)
            verb = verb.capitalize() + 'ing'
        rest = ' '.join(w for w in words[1:] if w not in ('revit',))
        text = (verb + ' ' + rest).strip()
    if n_targets:
        text = '{} · {} element{}'.format(text, n_targets, '' if n_targets == 1 else 's')
    if failed:
        text = 'Failed: ' + text
    return text if len(text) <= _CAPTION_MAX else text[:_CAPTION_MAX - 1] + '…'


def badge_text(category, size):
    """'Walls · 1200 × 300' under a box; the size alone without a category."""
    category = (category or '').strip()
    return '{} · {}'.format(category, size) if category else size


def hop_ms(stops):
    """Glide time per stop: shorter when the cursor has more stops to visit."""
    return max(260, MOVE_MS - 70 * max(0, stops - 1))


def tour_schedule(points, hop, dwell):
    """[(x, y, arrive_ms, leave_ms)]: glide `hop` ms to each point, stay `dwell`."""
    out, t = [], 0
    for x, y in points:
        arrive = t + hop
        out.append((x, y, arrive, arrive + dwell))
        t = arrive + dwell
    return out


def review_points(rect, seed, n=3):
    """`n` spread-out points in the middle of rect (l, t, r, b) for a call with
    no element to point at; the same seed gives the same path."""
    left, top, right, bottom = rect
    w, h = right - left, bottom - top
    state = (seed * 7919 + 17) % 2147483647 or 1
    pts = []
    for i in range(n):
        state = (state * 48271) % 2147483647
        fx = 0.22 + 0.56 * ((i + (state % 1000) / 1000.0) / float(n))
        state = (state * 48271) % 2147483647
        fy = 0.25 + 0.5 * ((state % 1000) / 1000.0)
        pts.append((left + w * fx, top + h * fy))
    return pts


def counts_text(counts):
    """{'create': 3, 'read': 1} -> 'Read 1 · Create 3' (ACTIONS order)."""
    parts = []
    for key, label, _g, _c in ACTIONS:
        if counts.get(key):
            parts.append('{} {}'.format(label, counts[key]))
    return ' · '.join(parts)


def clean_goal(text):
    """One line, at most NARRATE_MAX characters, or ''."""
    text = ' '.join('{}'.format(text or '').split())
    return text if len(text) <= NARRATE_MAX else text[:NARRATE_MAX - 1] + '…'


def _short(value, limit=24):
    text = '{}'.format(value if value is not None else '')
    text = '""' if text == '' else text
    return text if len(text) <= limit else text[:limit - 1] + '…'


def change_text(param, old, new):
    """'Comments: "" → Checked' for a parameter the call changed."""
    return '{}: {} → {}'.format(_short(param, 20), _short(old), _short(new))


EDIT_ACTIONS = ('create', 'modify', 'delete')


def track_touches(touched, agent, ids, action):
    """Remember which agent edited which element; return [(id, [other agents])]
    for ids another agent edited before. `touched` = {id: set(agents)}."""
    if action not in EDIT_ACTIONS:
        return []
    clashes = []
    for i in ids:
        others = touched.setdefault(i, set())
        before = sorted(a for a in others if a != agent)
        if before:
            clashes.append((i, before))
        others.add(agent)
    return clashes


def edge_point(rect, point, margin=18.0):
    """Where the ray from rect's centre to `point` leaves rect (inset by
    margin): (x, y, angle in degrees, 0 = right, 90 = down)."""
    import math
    left, top, right, bottom = rect
    cx, cy = (left + right) / 2.0, (top + bottom) / 2.0
    dx, dy = point[0] - cx, point[1] - cy
    if abs(dx) < 1e-9 and abs(dy) < 1e-9:
        dx = 1.0
    hw, hh = max((right - left) / 2.0 - margin, 1.0), max((bottom - top) / 2.0 - margin, 1.0)
    t = min(hw / abs(dx) if abs(dx) > 1e-9 else 1e18, hh / abs(dy) if abs(dy) > 1e-9 else 1e18)
    return cx + dx * t, cy + dy * t, math.degrees(math.atan2(dy, dx))


def offscreen_markers(rect, points, margin=18.0):
    """[(x, y, angle, count)]: one edge arrow per 45° sector holding targets
    that sit outside rect."""
    import math
    left, top, right, bottom = rect
    cx, cy = (left + right) / 2.0, (top + bottom) / 2.0
    sectors = {}
    for px, py in points:
        if left <= px <= right and top <= py <= bottom:
            continue
        ang = math.atan2(py - cy, px - cx)
        key = int(round(math.degrees(ang) / 45.0)) % 8
        acc = sectors.setdefault(key, [0.0, 0.0, 0])
        acc[0] += math.cos(ang)
        acc[1] += math.sin(ang)
        acc[2] += 1
    out = []
    for key in sorted(sectors):
        sx, sy, n = sectors[key]
        x, y, ang = edge_point(rect, (cx + sx * 1000.0, cy + sy * 1000.0), margin)
        out.append((x, y, ang, n))
    return out


def summary_lines(counts_by_agent, touched_count):
    """Idle card: one line per agent plus the elements changed."""
    lines = []
    for agent in sorted(counts_by_agent):
        text = counts_text(counts_by_agent[agent])
        if text:
            lines.append('{}: {}'.format(agent_style(agent)[0], text))
    if touched_count:
        lines.append('{} element{} changed'.format(touched_count, '' if touched_count == 1 else 's'))
    return lines


def _cam(point, eye, right, up, forward):
    d = (point[0] - eye[0], point[1] - eye[1], point[2] - eye[2])
    z = d[0] * forward[0] + d[1] * forward[1] + d[2] * forward[2]
    if z <= 1e-6:
        return None             # behind the camera
    x = d[0] * right[0] + d[1] * right[1] + d[2] * right[2]
    y = d[0] * up[0] + d[1] * up[1] + d[2] * up[2]
    return x / z, y / z


def project_perspective(point, cam, rect):
    """Model point -> screen (x, y) in a perspective 3D view, or None behind
    the eye. cam: eye/right/up/forward + c0/c1, the zoom corners, which map
    to the view's screen corners."""
    p = _cam(point, cam['eye'], cam['right'], cam['up'], cam['forward'])
    a = _cam(cam['c0'], cam['eye'], cam['right'], cam['up'], cam['forward'])
    b = _cam(cam['c1'], cam['eye'], cam['right'], cam['up'], cam['forward'])
    if p is None or a is None or b is None:
        return None
    u0, u1 = min(a[0], b[0]), max(a[0], b[0])
    v0, v1 = min(a[1], b[1]), max(a[1], b[1])
    if u1 - u0 < 1e-12 or v1 - v0 < 1e-12:
        return None
    left, top, right, bottom = rect
    return (left + (p[0] - u0) / (u1 - u0) * (right - left),
            bottom - (p[1] - v0) / (v1 - v0) * (bottom - top))


def perspective_box(corners, cam, rect):
    """Screen bounds of 3D box corners in a perspective view, clipped; None
    when any corner is behind the eye or the box is off screen."""
    pts = [project_perspective(c, cam, rect) for c in corners]
    if any(p is None for p in pts):
        return None
    x0, x1 = min(p[0] for p in pts), max(p[0] for p in pts)
    y0, y1 = min(p[1] for p in pts), max(p[1] for p in pts)
    left, top, r, bottom = rect
    if x1 < left or x0 > r or y1 < top or y0 > bottom:
        return None
    return (max(x0, left), max(y0, top), min(x1, r), min(y1, bottom))


# ─── Revit snapshot (main thread) ──────────────────────────────────────────────

def _xyz(p):
    return (p.X, p.Y, p.Z)


def _view_geometry(uiview, view):
    """Projection of one open view, or None when it is too small to draw in."""
    from Snippets._compat import eid_value
    r = uiview.GetWindowRectangle()
    rect = (float(r.Left), float(r.Top), float(r.Right), float(r.Bottom))
    if rect[2] - rect[0] < 40 or rect[3] - rect[1] < 40:
        return None
    geo = {'rect': rect, 'id': eid_value(view.Id), 'name': view.Name,
           'view': view, 'boxes': [], 'gone': [], 'ghosts': [], 'ortho': False}
    try:
        if '{}'.format(view.ViewType) in _NO_BOXES:
            return geo
        if bool(view.IsPerspective):
            ori = view.GetOrientation()
            corners = uiview.GetZoomCorners()
            geo['persp'] = {'eye': _xyz(ori.EyePosition), 'forward': _xyz(ori.ForwardDirection),
                            'up': _xyz(ori.UpDirection), 'right': _xyz(view.RightDirection),
                            'c0': _xyz(corners[0]), 'c1': _xyz(corners[1])}
            return geo
    except Exception:
        return geo
    try:
        corners = uiview.GetZoomCorners()
        c0, c1 = corners[0], corners[1]
        right, up = _xyz(view.RightDirection), _xyz(view.UpDirection)
    except Exception:
        return geo
    diag = (c1.X - c0.X, c1.Y - c0.Y, c1.Z - c0.Z)
    model_w = sum(diag[i] * right[i] for i in range(3))
    model_h = sum(diag[i] * up[i] for i in range(3))
    if abs(model_w) < 1e-9 or abs(model_h) < 1e-9:
        return geo
    geo.update({'origin': _xyz(c0), 'right': right, 'up': up,
                'w': model_w, 'h': model_h, 'ortho': True})
    return geo


def _shown_in(doc, view, ids):
    """Subset of ids visible in view (hidden / other level / cropped out drop)."""
    if not ids:
        return set()
    try:
        from Autodesk.Revit.DB import ElementId, ElementIdSetFilter, FilteredElementCollector
        from System.Collections.Generic import List
        from Snippets._compat import eid_value, make_eid
        lst = List[ElementId]()
        for i in ids:
            lst.Add(make_eid(i))
        col = FilteredElementCollector(doc, view.Id).WherePasses(ElementIdSetFilter(lst))
        return set(eid_value(e) for e in col.ToElementIds())
    except Exception:
        return set(ids)


def _boxes(doc, geo, ids):
    """[{'id', 'box', 'size'}] of the ids shown in one orthographic view."""
    out = []
    persp = geo.get('persp')
    if not (geo.get('ortho') or persp) or not ids:
        return out
    from Snippets._compat import make_eid
    view = geo['view']
    shown = _shown_in(doc, view, ids)
    right, up = (persp['right'], persp['up']) if persp else (geo['right'], geo['up'])
    geo.setdefault('offscreen', [])
    for i in ids:
        if i not in shown:
            continue
        try:
            elem = doc.GetElement(make_eid(i))
            if elem is None:
                continue
            bb = elem.get_BoundingBox(view) or elem.get_BoundingBox(None)
            if bb is None:
                continue
            lo, hi = bb.Min, bb.Max
            pts = [(x, y, z) for x in (lo.X, hi.X) for y in (lo.Y, hi.Y) for z in (lo.Z, hi.Z)]
            if bb.Transform is not None and not bb.Transform.IsIdentity:
                from Autodesk.Revit.DB import XYZ
                pts = [_xyz(bb.Transform.OfPoint(XYZ(*p))) for p in pts]
            if persp:
                box = perspective_box(pts, persp, geo['rect'])
            else:
                box = screen_box(pts, geo['origin'], right, up, geo['w'], geo['h'], geo['rect'])
            if box is None:
                # In the view but off screen: an edge arrow points at it.
                mid = tuple((lo_ + hi_) / 2.0 for lo_, hi_ in zip(pts[0], pts[-1]))
                centre = (project_perspective(mid, persp, geo['rect']) if persp else
                          project(mid, geo['origin'], right, up, geo['w'], geo['h'], geo['rect']))
                if centre is not None:
                    geo['offscreen'].append(centre)
                continue
            ext = (hi.X - lo.X, hi.Y - lo.Y, hi.Z - lo.Z)
            w = abs(sum(ext[k] * abs(right[k]) for k in range(3)))
            h = abs(sum(ext[k] * abs(up[k]) for k in range(3)))
            try:
                cat = elem.Category.Name if elem.Category is not None else ''
            except Exception:
                cat = ''
            out.append({'id': i, 'box': box, 'size': size_label(w, h),
                        'label': badge_text(cat, size_label(w, h))})
        except Exception:
            continue
    return out


def _open_views(app):
    """(doc, [geo], index of the active view) for the active document."""
    uidoc = app.ActiveUIDocument
    if uidoc is None:
        return None, [], None
    doc = uidoc.Document
    from Snippets._compat import eid_value
    active_id = eid_value(doc.ActiveView.Id)
    geos, active = [], None
    for uv in uidoc.GetOpenUIViews():
        try:
            view = doc.GetElement(uv.ViewId)
            geo = _view_geometry(uv, view) if view is not None else None
        except Exception:
            geo = None
        if geo is None:
            continue
        if geo['id'] == active_id:
            active = len(geos)
        geos.append(geo)
    return doc, geos, active


def _snapshot(app, ids, pre=None):
    """Geometry of every visible view with the targets boxed, or None."""
    doc, geos, active = _open_views(app)
    if doc is None or active is None:
        return None
    shown = [geos[i] for i in visible_views([g['rect'] for g in geos], active)]
    for geo in shown:
        geo['boxes'] = _boxes(doc, geo, ids)
        before = (pre or {}).get(geo['id'])
        if not before or before.get('rect') != geo['rect'] or before.get('origin') != geo.get('origin'):
            continue        # view moved / zoomed in between: old boxes are stale
        now = dict((b['id'], b) for b in geo['boxes'])
        for old in before['boxes']:
            new = now.get(old['id'])
            if new is None:
                geo['gone'].append(old)
            elif moved(old['box'], new['box']):
                geo['ghosts'].append((old, new))
    snap = {'rect': union_rect([g['rect'] for g in shown]), 'views': shown,
            'active': shown[0], 'view': shown[0]['name']}
    snap['any_box'] = any(g['boxes'] for g in shown)
    snap['offscreen'] = list(shown[0].get('offscreen') or [])
    return snap


def _pre_snapshot(app, ids):
    """{view id: {'rect', 'origin', 'boxes'}} before a delete / move."""
    doc, geos, active = _open_views(app)
    if doc is None or active is None:
        return None
    out = {}
    for i in visible_views([g['rect'] for g in geos], active):
        geo = geos[i]
        out[geo['id']] = {'rect': geo['rect'], 'origin': geo.get('origin'),
                          'boxes': _boxes(doc, geo, ids)}
    return out


# ─── Model interaction (main thread) ──────────────────────────────────────────

_NOT_OPENABLE = ('Internal', 'ProjectBrowser', 'SystemBrowser', 'Undefined')


def _openable(view):
    try:
        from Autodesk.Revit.DB import View
        if not isinstance(view, View) or view.IsTemplate:
            return False
        return '{}'.format(view.ViewType) not in _NOT_OPENABLE
    except Exception:
        return False


def _follow_views(uidoc, task, action, ids):
    """Open every view the call works in and end on the first one.

    Views come from the call (view_id / sheet_id / a created view) plus the
    owner view of view-specific targets. Returns [(view id, name, newly opened)].
    """
    from Snippets._compat import eid_value, make_eid
    doc = uidoc.Document
    wanted = extract_view_ids(getattr(task, 'arguments', None),
                              getattr(task, 'result', None), action)
    for i in ids:
        if len(wanted) >= MAX_FOLLOW_VIEWS:
            break
        try:
            elem = doc.GetElement(make_eid(i))
            if elem is not None and elem.ViewSpecific:
                owner = eid_value(elem.OwnerViewId)
                if owner > 0 and owner not in wanted:
                    wanted.append(owner)
        except Exception:
            continue
    views = []
    for vid in wanted[:MAX_FOLLOW_VIEWS]:
        try:
            view = doc.GetElement(make_eid(vid))
        except Exception:
            view = None
        if view is not None and _openable(view):
            views.append(view)
    if not views:
        return []
    open_ids = set(eid_value(uv.ViewId) for uv in uidoc.GetOpenUIViews())
    out = []
    # Open the extra views first, so the first (main) one ends up active.
    for view in reversed(views):
        vid = eid_value(view.Id)
        is_new = vid not in open_ids
        try:
            if is_new or (view is views[0] and eid_value(uidoc.ActiveView.Id) != vid):
                uidoc.ActiveView = view
        except Exception:
            continue        # transaction / edit mode: leave the view as it is
        out.insert(0, (vid, view.Name, is_new))
    return out


def _zoom_to(uidoc, ids):
    """Zoom the active view to targets it shows but has off screen."""
    try:
        from Autodesk.Revit.DB import XYZ
        from Snippets._compat import eid_value, make_eid
        doc = uidoc.Document
        view = doc.ActiveView
        shown = _shown_in(doc, view, ids)
        lo = hi = None
        for i in ids:
            if i not in shown:
                continue
            elem = doc.GetElement(make_eid(i))
            bb = elem.get_BoundingBox(view) if elem is not None else None
            if bb is None:
                continue
            lo = bb.Min if lo is None else XYZ(min(lo.X, bb.Min.X), min(lo.Y, bb.Min.Y), min(lo.Z, bb.Min.Z))
            hi = bb.Max if hi is None else XYZ(max(hi.X, bb.Max.X), max(hi.Y, bb.Max.Y), max(hi.Z, bb.Max.Z))
        if lo is None:
            return False
        pad = max(2.0, 0.25 * max(hi.X - lo.X, hi.Y - lo.Y, hi.Z - lo.Z))
        lo = XYZ(lo.X - pad, lo.Y - pad, lo.Z - pad)
        hi = XYZ(hi.X + pad, hi.Y + pad, hi.Z + pad)
        for uv in uidoc.GetOpenUIViews():
            if eid_value(uv.ViewId) == eid_value(view.Id):
                uv.ZoomAndCenterRectangle(lo, hi)
                return True
    except Exception:
        pass
    return False


def _select(uidoc, ids):
    try:
        from Autodesk.Revit.DB import ElementId
        from System.Collections.Generic import List
        from Snippets._compat import make_eid
        lst = List[ElementId]()
        for i in ids:
            if uidoc.Document.GetElement(make_eid(i)) is not None:
                lst.Add(make_eid(i))
        if lst.Count:
            uidoc.Selection.SetElementIds(lst)
    except Exception:
        pass


def _tile(app):
    """Post View > Tile Views; Revit runs it after this external event."""
    try:
        from Snippets._compat import postable_command_id
        cmd = postable_command_id(('TileViews', 'TileWindows', 'Tile'))
        if cmd is not None and app.CanPostCommand(cmd):
            app.PostCommand(cmd)
            return True
    except Exception:
        pass
    return False


# ─── Entry points (main thread, inside Execute) ───────────────────────────────

def _agent_of(task):
    agent = getattr(task, 'agent', None)
    audit = getattr(task, 'audit', None)
    agent = agent or getattr(audit, 'agent', None)
    if agent:
        return agent
    source = getattr(task, 'source', None) or getattr(audit, 'source', '')
    return 't3lab-assistant' if source == 'assistant' else 'claude-desktop'


def _internal(task):
    return '{}'.format(getattr(task, 'tool_name', '')).startswith('__')


# set_parameter / bulk_set_parameter / rename_element: (name arg, value arg).
_PARAM_TOOLS = {
    'set_parameter': (('parameter_name', 'name'), 'value'),
    'bulk_set_parameter': (('parameter_name', 'name'), 'value'),
    'rename_element': ((), 'new_name'),
}
# Per Revit session: who edited what, the agents' goals, the select event.
_session = {'touched': {}, 'goals': {}, 'select_event': None, 'select_cls': None}


def _param_value(elem, name):
    """Displayed value of a parameter (or the element's Name), or None."""
    try:
        if not name:
            return elem.Name
        prm = elem.LookupParameter(name)
        if prm is None:
            return None
        return prm.AsValueString() or prm.AsString() or (
            '{}'.format(prm.AsDouble()) if '{}'.format(prm.StorageType) == 'Double'
            else '{}'.format(prm.AsInteger()) if '{}'.format(prm.StorageType) == 'Integer' else '')
    except Exception:
        return None


def _param_name(tool, args):
    spec = _PARAM_TOOLS.get(tool)
    if spec is None or not isinstance(args, dict):
        return None, False
    for key in spec[0]:
        if args.get(key):
            return '{}'.format(args.get(key)), True
    return ('Name', True) if not spec[0] else (None, False)


def _read_values(app, ids, name):
    from Snippets._compat import make_eid
    doc = app.ActiveUIDocument.Document
    out = {}
    for i in ids[:MAX_STOPS]:
        elem = doc.GetElement(make_eid(i))
        if elem is not None:
            out[i] = _param_value(elem, None if name == 'Name' else name)
    return out


def _hwnd(app):
    try:
        return app.MainWindowHandle
    except Exception:
        return None


def before(app, task):
    """Ahead of the tool: show the agent working, remember the boxes a delete /
    move will change and the values an edit will overwrite. Never raises;
    costs one mtime check while the feature is off."""
    try:
        if not is_enabled() or _internal(task):
            return
        tool = getattr(task, 'tool_name', '')
        args = getattr(task, 'arguments', None)
        action = classify_action(tool, args)
        ids = extract_ids(args)
        try:
            doc, geos, active = _open_views(app)
            if active is not None:
                _overlay(_hwnd(app)).call(
                    'working', _agent_of(task), action,
                    describe_call(tool, args, len(ids)), tool, geos[active]['rect'])
        except Exception:
            pass
        name, ok = _param_name(tool, args)
        if ok and ids:
            task._cursor_old = (name, _read_values(app, ids, name))
        if action in ('delete', 'modify') and ids:
            task._cursor_pre = _pre_snapshot(app, ids)
    except Exception:
        pass


def _changes(app, task, ids):
    """['Comments: "" → Checked', ...] for the values the call changed."""
    old = getattr(task, '_cursor_old', None)
    if not old:
        return []
    name, before_values = old
    after_values = _read_values(app, ids, name)
    out = []
    for i in ids:
        if i in before_values and before_values[i] != after_values.get(i):
            out.append((i, change_text(name, before_values[i], after_values.get(i))))
    return out


def notify(app, task):
    """One finished tool call. Never raises."""
    try:
        if not is_enabled():
            hide()              # switched off from another engine (Settings)
            return
        if _internal(task):
            return                  # internal calls (status probes)
        tool = getattr(task, 'tool_name', '')
        args = getattr(task, 'arguments', None)
        result = getattr(task, 'result', None)
        failed = bool(getattr(task, 'exception', None)) or (
            isinstance(result, dict) and bool(result.get('error')))
        action = classify_action(tool, args, result)
        ids = [] if failed else extract_ids(args, result)
        agent = _agent_of(task)
        opts = options()
        uidoc = app.ActiveUIDocument
        followed = []
        pre = getattr(task, '_cursor_pre', None)
        if uidoc is not None and not failed:
            if opts['follow']:
                followed = _follow_views(uidoc, task, action, ids)
                if followed:
                    pre = None      # different view now: old boxes do not apply
            if opts['select'] and ids and action != 'delete':
                _select(uidoc, ids)
        snap = _snapshot(app, ids, pre)
        if snap is None:
            return
        if (opts['zoom'] and ids and not snap['any_box'] and action != 'delete'
                and uidoc is not None and _zoom_to(uidoc, ids)):
            snap = _snapshot(app, ids) or snap
        if opts['tile'] and any(new for _v, _n, new in followed):
            if len(list(uidoc.GetOpenUIViews())) >= 2:
                _tile(app)
        changes = [] if failed else _changes(app, task, ids)
        clashes = [] if failed else track_touches(_session['touched'], agent, ids, action)
        _ensure_select_event()
        try:
            from core import agent_cursors_glow
            if opts['glow']:
                agent_cursors_glow.update(app, sorted(_session['touched']))
            else:
                agent_cursors_glow.stop(app)        # no-op when never started
        except Exception:
            pass
        _overlay(_hwnd(app)).call(
            'show', agent, action, tool, failed, snap, followed, args=args,
            n_targets=len(ids), changes=changes, clashes=clashes,
            touched=len(_session['touched']),
            goal=_session['goals'].get(agent_key(agent) or 'claude-desktop', ''))
    except Exception:
        pass


# ─── Narration / approval wait (any thread) ────────────────────────────────────

NARRATE_TOOL = 't3lab_narrate'
NARRATE_SCHEMA = {
    'name': NARRATE_TOOL,
    'description': ('Say in one short line what you are about to do in Revit, e.g. '
                    '"Renumber the doors on Level 2". Shown to the user as a goal banner '
                    'over the model while Agent Cursors are on; changes nothing. Call it '
                    'before a multi-step task and when the goal changes.'),
    'inputSchema': {
        'type': 'object',
        'properties': {
            'message': {'type': 'string',
                        'description': 'The goal, one line (max {} characters).'.format(NARRATE_MAX)},
        },
        'required': ['message'],
    },
}


def narrate(agent, message):
    """t3lab_narrate: set the agent's goal banner. Any thread; never raises.
    Returns the tool result."""
    goal = clean_goal(message)
    _session['goals'][agent_key(agent) or 'claude-desktop'] = goal
    try:
        if is_enabled() and _state.get('overlay') is not None:
            _state['overlay'].call('narrate', agent or 'claude-desktop', goal)
    except Exception:
        pass
    return {'success': True, 'goal': goal,
            'shown': bool(goal) and is_enabled(),
            'note': ('Shown over the model.' if is_enabled() else
                     'Agent Cursors are off in Revit; the goal is kept for when they are on.')}


def waiting(agent, tool, arguments=None, on=True):
    """A call waits for the user's OK (approval dialog): pulse the cursor.
    Any thread; never raises."""
    try:
        if is_enabled() and _state.get('overlay') is not None:
            _state['overlay'].call('waiting', agent or 'claude-desktop', tool,
                                   describe_call(tool, arguments), bool(on))
    except Exception:
        pass


# ─── Select touched (HUD button -> ExternalEvent -> main thread) ──────────────

def touched_ids():
    return sorted(_session['touched'])


def _ensure_select_event():
    """ExternalEvent the HUD's "Select touched" raises. Created here, in an
    API context (notify runs inside Execute); one per Revit session."""
    if _session['select_event'] is not None:
        return _session['select_event']
    try:
        from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler
        cls = _session['select_cls']
        if cls is None:
            class _SelectTouched(IExternalEventHandler):
                __namespace__ = 'T3Lab.AgentCursorsSelect'

                def Execute(self, uiapp):
                    try:
                        _select(uiapp.ActiveUIDocument, touched_ids())
                    except Exception:
                        pass

                def GetName(self):
                    return 'T3Lab Agent Cursors - select touched'
            cls = _session['select_cls'] = _SelectTouched
        _session['select_event'] = ExternalEvent.Create(cls())
    except Exception:
        _session['select_event'] = None
    return _session['select_event']


def select_touched():
    """HUD button: select every element the agents changed. Any thread."""
    event = _session.get('select_event')
    if event is not None:
        try:
            event.Raise()
            return True
        except Exception:
            pass
    return False


def hide():
    """Hide the overlay now (no-op when it was never shown)."""
    ov = _state.get('overlay')
    if ov is not None:
        try:
            ov.call('hide')
        except Exception:
            pass


# ─── Overlay (core/agent_cursors_overlay.py) ──────────────────────────────────

_state = {'overlay': None}


def _overlay(hwnd):
    """The overlay host (its own UI thread under IronPython, see
    agent_cursors_overlay.OverlayHost); .call(method, ...) posts to it."""
    ov = _state.get('overlay')
    if ov is None or not ov.alive:
        from core.agent_cursors_overlay import OverlayHost
        ov = OverlayHost(hwnd, {'select_touched': select_touched})
        _state['overlay'] = ov
    return ov
