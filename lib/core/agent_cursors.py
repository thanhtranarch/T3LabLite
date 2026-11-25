# -*- coding: utf-8 -*-
"""Agent cursors: a live overlay of which AI agent does what, where.

Each MCP tool call moves the agent's one cursor (the action is the icon on
its tag) to each element it read, edited, created or deleted in turn (move,
click, next), in every visible view, with a plain caption ("Moving elements ·
3 elements"), a badge per element ("Walls · 1200 × 300") and an effect per
action. Elements a call names are framed while it runs; what it created or
changed is taken from Revit's DocumentChanged too (code returns no ids) and
stays highlighted with a chip ("Walls ×3 created"). v3 tags, draw-along,
thinking, parking and Crew: t3lab-revit-api dev/plan/agent-cursors-v3-study.md.
Glides arc like a hand on a mouse; ui_route walks the Revit UI first
(Project Browser > Sheets, Properties) - NEEDS VERIFICATION in Revit.
The HUD shows what runs now as a block of the action's colour (band_rows),
lists the last steps and each agent's plan (t3lab_narrate tasks / step):
while a step runs the cursor names that task, the call goes underneath; a
send_code_to_revit call is named by its `purpose` or its code (summarize_code).
Keep bar (on by default) keeps the bar, folded, between calls; the cursors
leave when the agents go idle.
Follow mode opens the views the agent works in; zoom brings off-screen
targets into view. Idea after swarm-cursors (github.com/DhvanilPatel/
swarm-cursors); drawn here in code-built WPF, no code from it.

before() / notify() run on Revit's main thread from
MCPExternalEventHandler.Execute (core/server.py): before() just ahead of the
tool (boxes of what a delete / move is about to change), notify() after the
caller has its result. Off by default; Settings > Agent Cursors or the HUD.
Engine-agnostic (IronPython 2.7 + CPython 3); the pure helpers are tested by
lite_guard/test_agent_cursors.py, the bar on a fake WPF by
lite_guard/test_agent_cursors_overlay.py.
"""

from __future__ import unicode_literals

import io
import json
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
BEND = 0.12                 # a hop arcs sideways by this share of its length...
BEND_MAX = 48.0             # ...at most this many px (glide_path)
GLIDE_SAMPLES = 10          # keyframes per hop along the arc
PANE_CACHE_S = 4.0          # pane rects (Project Browser, Properties) reused this long
FEED_SIZE = 5               # last steps listed on the HUD
REPLAY_SIZE = 20            # steps the HUD can replay
SUMMARY_S = 20              # idle summary card stays this long
COLLAPSE_S = 10             # idle this long: the bar folds to one line (mouse over opens it)
FEED_SHOWN = 3              # feed rows on the bar (the rest wait for Replay)
HUD_POS_KEY = 'agent_cursors_hud_pos'   # "dx,dy" DIP from the default spot, per user
NARRATE_MAX = 90            # goal banner text
PLAN_MAX = 8                # tasks in an agent's plan (t3lab_narrate tasks=[...])
TASK_MAX = 60               # characters per task
BOARD_FILE = 'board.json'   # = core/harness_store.py BOARD_FILE (docs/harness.md)
BOARD_ROWS_MAX = 6          # harness tasks listed on the HUD
NAME_MAX = 16               # characters of an unknown agent's name on its tag
TAG_MAX = 36                # characters of "what it does" on a cursor tag
MAX_CAPTURE = 60            # added / modified ids kept per call (DocumentChanged)
FRESH_HOLD_S = 7.0          # created / changed elements stay highlighted this long
SCAN_RADIUS = 70.0          # px round the cursor a call with no element scans

# name -> (mcp_paths.json key, default). The HUD toggles them live; a saved
# choice wins. 'model' is the master switch, OFF by default (owner,
# 2026-11-18): only the bar shows; no cursor, boxes or effects on the views,
# and Follow / Zoom / Select / Tile / Glow do nothing. Turned on, the others
THINK_AFTER_S = 4           # after a call the cursor shows "thinking" (= BUSY_GAP_S)
PARK_DIM = 0.55             # opacity of a parked cursor (agents idle)
PARK_S = 20                 # parked cursors stay this long beside their work, then leave
CREW_MAX = 3                # helper cursors one call fans out into (Crew)
SPREAD_GAP = 22.0           # px two cursors keep between them on one target
# apply as saved (all on by default, owner 2026-11-17).
OPTIONS = (
    ('model', 'agent_cursors_on_model', False),
    ('follow', 'agent_cursors_follow_views', True),
    ('zoom', 'agent_cursors_zoom', True),
    ('select', 'agent_cursors_select', True),
    ('tile', 'agent_cursors_tile_views', True),
    ('glow', 'agent_cursors_glow', True),
    ('trail', 'agent_cursors_trail', True),
    ('ripple', 'agent_cursors_ripple', True),
    ('labels', 'agent_cursors_labels', True),
    ('crew', 'agent_cursors_crew', True),
    ('keep', 'agent_cursors_keep_bar', True),
)
# Animation (mcp_paths.json agent_cursors_speed): name -> time factor. 'off' =
# no glides, ripples or sweeps, boxes and captions only. Motion is part of the
# feature: no control on the bar or in Settings (owner, 2026-11-19), the key
# stays for a machine that needs it off. It replaced the "Less motion" switch,
# which shipped on and was saved for everyone by the Settings dialog, so a
# saved value never meant a choice.
SPEED_KEY = 'agent_cursors_speed'
SPEEDS = (('off', 1.0), ('calm', 1.6), ('normal', 1.0), ('fast', 0.55))
SPEED_DEFAULT = 'calm'      # owner, 2026-11-19
OPTION_LABELS = {
    'model': ('On model', 'Draw the cursor, element boxes and effects on the views and let the '
                          'options below act. Off: the bar only, nothing on the model'),
    'follow':('Follow', 'Open and switch to the views the agent works in'),
    'zoom': ('Zoom', 'Zoom the view to targets that are off screen'),
    'select': ('Select', "Select the agent's targets (replaces your selection)"),
    'tile': ('Tile', 'Tile the open views when the agent opens another one'),
    'glow': ('Glow', 'Experimental: outline touched elements inside the Revit view '
                     '(DirectContext3D). Turn off if Revit misbehaves.'),
    'trail': ('Trail', 'Draw the path the cursor takes to its next target, and join the '
                       'elements a call created'),
    'ripple': ('Ripple', 'A ring where the cursor lands on a target, and the click of '
                         'the pointer'),
    'labels': ('Labels', 'Name each target (category and size), say what a call created '
                         'and show the values it changed'),
    'crew': ('Crew', 'A call that works on several categories sends one helper cursor per '
                     'category, named after the work ("Claude \u00b7 doors"), not a person'),
    'keep': ('Keep bar', 'Keep the bar (folded) on screen between calls instead of '
                         'fading out (the x on the bar turns Agent Cursors off)'),
}

# agent id (bridge `t3lab/agent`) -> (tag, colour). Unknown agents get a
# stable colour from _SPARE by name.
AGENTS = {
    'claude-desktop':  ('Claude', '#D97757'),
    'claude-code':     ('Claude Code', '#B4532A'),
    'codex':           ('GPT', '#10A37F'),
    'chatgpt':         ('GPT', '#10A37F'),
    'antigravity':     ('Antigravity', '#4285F4'),
    'cursor':          ('Cursor', '#7C3AED'),
    'github-copilot':  ('Copilot', '#24292F'),
    'windows-copilot': ('Copilot', '#0078D4'),
    'cli':             ('CLI', '#64748B'),
    't3lab-assistant': ('T3Lab Assistant', '#EA680C'),
    'mcp-client':      ('AI app', '#64748B'),      # no name sent at all
}
# MCP clientInfo names -> AGENTS key.
_ALIASES = {
    'claude-ai': 'claude-desktop', 'claude': 'claude-desktop',
    'codex-mcp-client': 'codex', 'openai-codex': 'codex',
    'gemini-cli-mcp-client': 'antigravity', 'gemini': 'antigravity',
    'cursor-vscode': 'cursor',
}
# Client names that carry a suffix: first matching prefix wins. Never map a
# name on a guess (2026-11-17: 'local-agent-mode-...' was mapped to Claude,
# then to Antigravity, both unverified). The bridge names the app from
# --agent or from the process that launched it (bridge.host_agent).
_ALIAS_PREFIXES = (
    ('claude-code', 'claude-code'),
    ('claude', 'claude-desktop'), ('codex', 'codex'), ('chatgpt', 'chatgpt'),
    ('openai', 'chatgpt'), ('gemini', 'antigravity'), ('antigravity', 'antigravity'),
    ('cursor', 'cursor'),
)
_AGENT_NOISE = re.compile(r'([-_ ./]+(mcp|client|server|app|bridge|stdio))+$', re.I)
_AGENT_VERSION = re.compile(r'[-_ /]+v?\d+(\.\d+)*$')
_ACRONYMS = ('ai', 'api', 'bim', 'cli', 'gpt', 'ide', 'llm', 'mcp', 'qa')
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
    "The bar shows each step as a block of its action's colour with a plain-English caption, plus the agent's goal (t3lab_narrate) and a feed of the last steps.",
    "One cursor per AI app moves to each target and clicks it; new or changed elements stay highlighted (\"Walls ×3 created\"), changed values show as \"Mark: A-101 → A-102\".",
    "Long calls show \"Working 8 s\"; a call waiting for your OK pulses amber.",
    "Two agents editing the same element flash red; targets off screen get an edge arrow.",
    "When the agents go idle: a summary, Select touched and Replay.",
)
CONS = (
    "By default only the bar shows. Turn on \"On model\" for the cursor, boxes and the options below; then each tool call does a little extra work (bounding boxes of up to {} elements, twice when the call names its elements).".format(MAX_TARGETS),
    "Follow switches the active view, so a later tool that works on 'the active view' works on the view the agent opened.",
    "Zoom and Select change your camera and selection while an agent works (with On model).",
    "Boxes are a snapshot: pan or zoom while an agent works and they sit off until its next call; Replay uses the old positions.",
    "Glow (on by default) is experimental: it draws inside Revit's 3D views and is not tested on every Revit version.",
    "The overlay shows up in screenshots and recordings.",
)


def pros_cons_sections():
    """[(heading, [line], glyph)] for the Settings dialog (T3Dialog sections)."""
    return [('Benefits', list(PROS), '\uE73E'), ('Drawbacks', list(CONS), '\uE7BA')]


def pros_cons_text():
    """The same as plain text (TaskDialog fallback)."""
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
        _session['plans'].clear()
    return bool(on)


def options():
    """{name: bool} for every OPTIONS entry."""
    data = _settings()
    return dict((name, bool(data.get(key, default))) for name, key, default in OPTIONS)


def effective(opts):
    """What actually acts: with 'model' off (the default) only 'keep' stays,
    every option that draws on or moves the model is off."""
    opts = dict(opts)
    if not opts.get('model'):
        for name in opts:
            if name not in ('model', 'keep'):
                opts[name] = False
    return opts


def on_model():
    return bool(options().get('model'))


def set_option(name, value):
    for opt, key, _default in OPTIONS:
        if opt == name:
            _paths().set_setting(key, bool(value))
            _cache['mtime'] = None
            return bool(value)
    raise KeyError(name)


def speed():
    """'off' | 'calm' | 'normal' | 'fast'."""
    name = _settings().get(SPEED_KEY, SPEED_DEFAULT)
    return name if name in dict(SPEEDS) else SPEED_DEFAULT


def motion_off():
    """Animation is Off: nothing glides, ripples or sweeps."""
    return speed() == 'off'


def speed_factor(name=None):
    return dict(SPEEDS)[name or speed()]


def next_speed(name):
    names = [n for n, _f in SPEEDS]
    return names[(names.index(name) + 1) % len(names)] if name in names else SPEED_DEFAULT


def set_speed(name):
    _paths().set_setting(SPEED_KEY, name if name in dict(SPEEDS) else SPEED_DEFAULT)
    _cache['mtime'] = None
    return speed()


def hud_offset():
    """(dx, dy) the user dragged the bar from its default spot, in DIP."""
    return parse_offset(_settings().get(HUD_POS_KEY, ''))


def set_hud_offset(dx, dy):
    _paths().set_setting(HUD_POS_KEY, '' if (dx, dy) == (0, 0) else '{},{}'.format(int(dx), int(dy)))
    _cache['mtime'] = None


# Pause (HUD): the user holds model changes; read-only tools still run. Lives in
# memory only, so a Revit restart never starts paused.
_pause = {'on': False}


def is_paused():
    return bool(_pause['on'])


def set_paused(on):
    _pause['on'] = bool(on)
    try:
        if _state.get('overlay') is not None:
            _state['overlay'].call('set_paused', bool(on))
    except Exception:
        pass
    return bool(on)


def paused_refusal(tool_name, is_write):
    """The tool result for a write refused while paused, else None."""
    if not (_pause['on'] and is_write) or tool_name == NARRATE_TOOL:
        return None
    return {'success': False, 'error_code': 'paused', 'nothing_ran': True,
            'error': ('Paused by the user in Revit (Agent Cursors bar): model changes wait '
                      'until they click Resume. Read-only tools still work; ask the user '
                      'or try again later.')}


def options_text():
    """Follow on · Zoom on · Select off · Tile off."""
    opts = options()
    return ' · '.join('{} {}'.format(OPTION_LABELS[n][0], 'on' if opts[n] else 'off')
                      for n, _k, _d in OPTIONS)


# ─── Pure helpers ──────────────────────────────────────────────────────────────

def agent_key(agent):
    key = (agent or '').strip().lower()
    if key in _ALIASES:
        return _ALIASES[key]
    if key in AGENTS or '/' in key:
        return key
    for prefix, name in _ALIAS_PREFIXES:
        if key == prefix or key.startswith(prefix + '-') or key.startswith(prefix + '_'):
            return name
    return key


def pretty_agent(name):
    """A raw MCP client id as a tag: 'acme-bot-mcp-client' -> 'Acme Bot'
    (noise and version dropped, title case, at most NAME_MAX characters)."""
    text = ' '.join('{}'.format(name or '').split())
    for _ in range(3):
        text = _AGENT_NOISE.sub('', _AGENT_VERSION.sub('', text))
    words = [w for w in re.split(r'[-_./\s]+', text) if w]
    if not words:
        return 'Agent'
    text = ' '.join(w.upper() if w.lower() in _ACRONYMS else w[0].upper() + w[1:]
                    for w in words)
    return text if len(text) <= NAME_MAX else text[:NAME_MAX - 1].rstrip() + '…'


def agent_style(agent):
    """(tag, '#RRGGBB') for a bridge agent id, MCP client name or harness
    member. A member "claude-desktop/arch" (one chat of a client) is
    "Claude · arch" in a colour of its own."""
    key = agent_key(agent)
    if key in AGENTS:
        return AGENTS[key]
    if not key:
        return AGENTS['mcp-client']
    total = 0
    for ch in key:
        total = (total * 31 + ord(ch)) % 1000003
    color = _SPARE[total % len(_SPARE)]
    if '/' in key:
        base, alias = key.split('/', 1)
        base = agent_key(base)
        tag = AGENTS[base][0] if base in AGENTS else pretty_agent(base)
        return (u'{} · {}'.format(tag, alias[:10]), color)
    return (pretty_agent(agent), color)


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
        kwargs.get('limit', MAX_TARGETS), 3,
        ('elements', 'created', 'items', 'created_elements', 'new_elements', 'placed',
         'instances', 'results'))


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


def _round_all(value, digits):
    if value is None:
        return None
    return tuple(round(float(v), digits) for v in value)


def layout_key(geos, active):
    """What decides where a box lands on screen: each open view's id, window
    rectangle and zoom (origin + extents), plus which one is active. Two equal
    keys = Revit finished opening / tiling / zooming views."""
    return (active, tuple((g.get('id'), _round_all(g.get('rect'), 0),
                           _round_all(g.get('origin'), 4),
                           round(float(g.get('w') or 0.0), 4),
                           round(float(g.get('h') or 0.0), 4)) for g in geos))


def overlap_px(a, b):
    """(width, height) px of the overlap of two (l, t, r, b) rectangles;
    (0, 0) when apart."""
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return (w, h) if w > 0 and h > 0 else (0, 0)


def front_hides(front_is_revit, revit_minimized, front_rect, overlay_rect, min_px=16):
    """Whether the topmost overlay must hide. It used to hide whenever Revit
    was not the foreground app, so while the user typed in the AI app (the
    normal MCP case) the cursors vanished and flashed back on every call.
    Now it hides only when Revit is minimized or the app in front actually
    covers the views the overlay draws over."""
    if revit_minimized:
        return True
    if front_is_revit or front_rect is None or overlay_rect is None:
        return False
    w, h = overlap_px(front_rect, overlay_rect)
    return w >= min_px and h >= min_px     # a window edge touching is not covering


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


_CODE_COMMENT = re.compile(r'^\s*#\s*(.+?)\s*$')
_CODE_DOC = re.compile(r'^\s*[uUrR]?("""|\'\'\')\s*(.+?)\s*("""|\'\'\')?\s*$')
_CODE_TRANSACTION = re.compile(r'Transaction(?:Group)?\s*\(\s*\w+\s*,\s*[uUrR]?([\'"])(.+?)\1')
_CODE_CATEGORY = re.compile(r'\bOST_([A-Za-z]+)')
_CODE_CLASS = re.compile(r'OfClass\s*\(\s*(?:\w+\.)*([A-Z][A-Za-z]+)\s*\)')
_CODE_EDIT = re.compile(r'\.Start\s*\(|\bSet\s*\(|\.Delete\s*\(|\.Create\w*\s*\(|\.NewFamilyInstance\s*\(')
_CAMEL = re.compile(r'(?<=[a-z])(?=[A-Z])')


def _humanize(name):
    """'StairsRailing' -> 'stairs railing'."""
    return _CAMEL.sub(' ', name).lower()


def summarize_code(code):
    """What a send_code_to_revit snippet does, in a few words: its first
    comment or docstring line, else its Transaction name, else the categories
    or classes it collects ('Editing stairs, railings'), else ''."""
    code = '{}'.format(code or '')
    for line in code.splitlines()[:6]:
        if not line.strip():
            continue
        found = _CODE_COMMENT.match(line) or _CODE_DOC.match(line)
        if found:
            text = found.group(found.lastindex if found.re is _CODE_COMMENT else 2).strip(' .:-')
            if len(text) >= 4 and not text.startswith('!') and 'coding' not in text:
                return text[0].upper() + text[1:]
        if not line.lstrip().startswith(('#', 'import', 'from', '"""', "'''")):
            break
    found = _CODE_TRANSACTION.search(code)
    if found and found.group(2).strip():
        return found.group(2).strip()
    names = []
    found = [_humanize(c) for c in _CODE_CATEGORY.findall(code)]
    found += [_humanize(c) + ('' if c.endswith('s') else 's') for c in _CODE_CLASS.findall(code)]
    for name in found:
        if name not in names and name not in ('elements', 'element types'):
            names.append(name)
    if names:
        verb = 'Editing' if _CODE_EDIT.search(code) else 'Reading'
        return '{} {}'.format(verb, ', '.join(names[:3]))
    return ''


def describe_call(tool_name, arguments=None, n_targets=0, failed=False):
    """'Listing open documents', 'Moving elements · 3 elements', 'Failed: …'.
    send_code_to_revit: its `purpose` (Lite: `description`), else what the
    code itself says."""
    name = '{}'.format(tool_name or '').strip().lower()
    if name == 'send_code_to_revit':
        args = arguments if isinstance(arguments, dict) else {}
        purpose = ' '.join('{}'.format(args.get('purpose') or args.get('description')
                                       or '').split())
        text = purpose or summarize_code(args.get('code')) or 'Running code in Revit'
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


def ease_human(t):
    """Minimum-jerk profile (10t^3 - 15t^4 + 6t^5): how a hand moves a mouse,
    slow off the mark, fast mid-way, settling on the target."""
    t = min(max(t, 0.0), 1.0)
    return t * t * t * (10.0 - 15.0 * t + 6.0 * t * t)


def glide_path(p0, p1, bend=BEND, samples=GLIDE_SAMPLES):
    """[(x, y, tau)] from p0 to p1 on a gentle arc, tau = share of the hop's
    time (0..1], spaced by ease_human. bend = sideways offset of the arc's
    middle as a share of the distance (sign = side), at most BEND_MAX px."""
    import math
    (x0, y0), (x1, y1) = p0, p1
    dx, dy = x1 - x0, y1 - y0
    dist = math.hypot(dx, dy)
    if dist < 2.0:
        return [(x1, y1, 1.0)]
    off = max(-BEND_MAX, min(BEND_MAX, dist * bend))
    cx, cy = (x0 + x1) / 2.0 - dy / dist * off, (y0 + y1) / 2.0 + dx / dist * off
    out = []
    for k in range(1, samples + 1):
        tau = k / float(samples)
        u = ease_human(tau)
        a, b, c = (1 - u) * (1 - u), 2 * (1 - u) * u, u * u
        out.append((a * x0 + b * cx + c * x1, a * y0 + b * cy + c * y1, tau))
    return out


def motion_keyframes(start, sched, seed=0, straight=()):
    """[(x, y, ms)] for a tour sched [(x, y, arrive_ms, leave_ms)] from
    `start`: each hop an arc (glide_path, the side alternating), each stop
    held until it leaves. Linear keyframes between them read as one smooth
    human motion. Hops into the indexes in `straight` are straight lines
    (the cursor drawing a wall along its length)."""
    frames = [(start[0], start[1], 0)]
    pos, t = start, 0
    for k, (x, y, arrive, leave) in enumerate(sched):
        bend = 0.0 if k in straight else BEND if (seed + k) % 2 == 0 else -BEND
        for px, py, tau in glide_path(pos, (x, y), bend):
            frames.append((px, py, int(round(t + (arrive - t) * tau))))
        if leave > arrive:
            frames.append((x, y, leave))
        pos, t = (x, y), leave
    return frames


# ─── Revit UI on the way (Project Browser, Properties) ───────────────────────
# A person creating sheets goes to Project Browser > Sheets first; the cursor
# does too, then into the model / the opened sheet. Pane rects come from
# Revit's WPF tree or, failing that, an estimate (_locate_panes).
PANE_TITLES = (('browser', 'Project Browser'), ('properties', 'Properties'))
_PANE_TITLE = dict(PANE_TITLES)
# (token, browser node) in priority order: "place views on sheets" is Sheets.
_BROWSER_NODES = (
    ('sheet', 'Sheets'), ('schedule', 'Schedules/Quantities'), ('legend', 'Legends'),
    ('view', 'Views'), ('family', 'Families'), ('group', 'Groups'), ('link', 'Revit Links'),
)
_TOKEN_ALIASES = {
    'sheets': 'sheet', 'schedules': 'schedule', 'quantities': 'schedule',
    'takeoff': 'schedule', 'legends': 'legend', 'views': 'view', 'families': 'family',
    'famigen': 'family', 'groups': 'group', 'links': 'link', 'parameters': 'parameter',
    'param': 'parameter', 'params': 'parameter', 'properties': 'parameter',
}
# Revit API names in send_code_to_revit code -> token.
_CODE_TOKENS = (
    (re.compile(r'\bViewSheet\b|OST_Sheets'), 'sheet'),
    (re.compile(r'\bViewSchedule\b|OST_Schedules'), 'schedule'),
    (re.compile(r'\b(ViewPlan|ViewSection|View3D|ViewDrafting)\b|\.Duplicate\s*\(\s*ViewDuplicateOption'),
     'view'),
    (re.compile(r'\bLoadFamily\w*\s*\(|\bFamilyManager\b'), 'family'),
    (re.compile(r'\bLookupParameter\s*\(|\bget_Parameter\s*\(|\.Set\s*\('), 'parameter'),
)
# A token here means the call is about elements in a view, not the views
# themselves (get_current_view_elements, create_view_filter).
_NOT_VIEWS = ('current', 'filter', 'filters', 'element', 'elements', 'crop', 'override')
# Read-only AI-side FamiGen calls never touch Revit's UI.
_NOT_FAMILY = ('schema', 'templates', 'validate', 'propose', 'studio')


def _tokens(*texts):
    out = []
    for text in texts:
        for word in re.split(r'[^a-z0-9]+', '{}'.format(text or '').lower()):
            if word:
                out.append(_TOKEN_ALIASES.get(word, word))
    return out


def ui_route(tool_name, arguments=None, action=None, categories=()):
    """The Revit UI a person would use for this call, as stops for the cursor:
    [{'pane', 'title', 'node', 'label', 'when'}] (empty when it is pure
    model work). when = 'before' the model (Project Browser: pick the sheet /
    view first) or 'after' it (Properties: select, then edit). categories =
    what DocumentChanged saw the call add or change ('Sheets'), for code."""
    args = arguments if isinstance(arguments, dict) else {}
    name = '{}'.format(tool_name or '').lower()
    words = _tokens(name, args.get('purpose'), args.get('operation'), *list(categories or ()))
    code = args.get('code') if name == 'send_code_to_revit' else None
    if isinstance(code, (str, type(''))):
        words += [tok for rx, tok in _CODE_TOKENS if rx.search(code)]
    if 'view' in words and any(w in words for w in _NOT_VIEWS):
        words = [w for w in words if w != 'view']
    if 'family' in words and any(w in words for w in _NOT_FAMILY):
        words = [w for w in words if w != 'family']
    if 'parameter' in words and 'project' in words:
        words = [w for w in words if w != 'parameter']     # Manage > Project Parameters
    for token, node in _BROWSER_NODES:
        if token in words:
            return [{'pane': 'browser', 'title': _PANE_TITLE['browser'], 'node': node,
                     'label': u'{} \u203a {}'.format(_PANE_TITLE['browser'], node),
                     'when': 'before'}]
    if 'parameter' in words or 'rename' in words:
        param = args.get('parameter_name') or args.get('param_name') or args.get('parameter')
        node = _short(param, 20) if isinstance(param, (str, type(''))) and param else None
        label = _PANE_TITLE['properties'] + (u' \u203a {}'.format(node) if node else u'')
        return [{'pane': 'properties', 'title': _PANE_TITLE['properties'], 'node': node,
                 'label': label, 'when': 'after'}]
    return []


# Where a browser node usually sits, as a share of the pane height, when its
# row cannot be found (collapsed tree, a non-WPF browser).
_NODE_SHARE = {'Views': 0.14, 'Legends': 0.24, 'Schedules/Quantities': 0.32,
               'Sheets': 0.40, 'Families': 0.50, 'Groups': 0.60, 'Revit Links': 0.68}


def pane_point(pane_rect, node=None, node_rect=None):
    """Where the cursor clicks in a pane: the node's row when it was found,
    else the node's usual height (browser) or the upper part (Properties)."""
    left, top, right, bottom = pane_rect
    w, h = right - left, bottom - top
    if node_rect is not None:
        return (min(node_rect[0] + 24.0, node_rect[2]), (node_rect[1] + node_rect[3]) / 2.0)
    share = _NODE_SHARE.get(node, 0.30)
    return (left + min(70.0, w * 0.35), top + h * share)


def dock_estimate(main_rect, mdi_rect, shown, min_gap=120.0):
    """{pane: rect} guessed from where Revit docks panes by default (left of
    the views: Properties above, Project Browser below), for the panes Revit
    says are `shown`; {} when there is no dock-sized gap on either side."""
    if not main_rect or not mdi_rect or not shown:
        return {}
    ml, mt, mr, mb = main_rect
    vl, vt, vr, vb = mdi_rect
    if vl - ml >= min_gap:
        x0, x1 = ml + 4.0, vl - 4.0
    elif mr - vr >= min_gap:
        x0, x1 = vr + 4.0, mr - 4.0
    else:
        return {}
    order = [p for p in ('properties', 'browser') if p in shown]
    if not order:
        return {}
    step = (vb - vt) / float(len(order))
    return dict((pane, (x0, vt + step * k, x1, vt + step * (k + 1)))
                for k, pane in enumerate(order))


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


def scan_points(anchor, rect, seed, n=3, radius=SCAN_RADIUS):
    """`n` points within `radius` px of anchor (x, y), kept inside rect
    (l, t, r, b): a call with no element scans next to where the cursor last
    worked instead of crossing the view. anchor None = the middle of rect;
    n = 1 is the anchor itself; the same seed gives the same path."""
    import math
    left, top, right, bottom = rect
    pad = min(24.0, max(0.0, (right - left) / 4.0), max(0.0, (bottom - top) / 4.0))

    def clamp(x, y):
        return (min(max(x, left + pad), right - pad), min(max(y, top + pad), bottom - pad))
    if anchor is None:
        anchor = ((left + right) / 2.0, (top + bottom) / 2.0)
    ax, ay = clamp(anchor[0], anchor[1])
    if n <= 1:
        return [(ax, ay)]
    state = (seed * 7919 + 17) % 2147483647 or 1
    pts = []
    for i in range(n):
        state = (state * 48271) % 2147483647
        ang = 2 * math.pi * ((i + (state % 1000) / 1000.0) / float(n))
        state = (state * 48271) % 2147483647
        dist = radius * (0.55 + 0.45 * (state % 1000) / 1000.0)
        pts.append(clamp(ax + dist * math.cos(ang), ay + dist * math.sin(ang)))
    return pts


def place_tags(items, rect=None, gap=12.0, margin=4.0, avoid=()):
    """Where each cursor's tag goes: [(dx, dy)] from its pointer tip, for
    items [(x, y, tag_w, tag_h)] in priority order. Right-below by default;
    flipped left / above when it would leave rect (l, t, r, b), and moved to
    the next free spot when it would cover a tag placed before it or one of
    the `avoid` boxes (the bar)."""
    placed, out = [tuple(a) for a in avoid or ()], []
    for x, y, w, h in items:
        spots = [(gap, gap), (-w - margin, gap), (gap, -h - margin), (-w - margin, -h - margin)]
        for k in (1, 2, 3):
            step = k * (h + 3.0)
            spots += [(gap, gap + step), (-w - margin, gap + step),
                      (gap, -h - margin - step), (-w - margin, -h - margin - step)]
        inside = [(dx, dy) for dx, dy in spots if rect is None or (
            x + dx >= rect[0] + margin and x + dx + w <= rect[2] - margin
            and y + dy >= rect[1] + margin and y + dy + h <= rect[3] - margin)]
        free = [(dx, dy) for dx, dy in inside if not any(
            _overlaps((x + dx, y + dy, x + dx + w, y + dy + h), p) for p in placed)]
        dx, dy = (free or inside or spots)[0]
        placed.append((x + dx, y + dy, x + dx + w, y + dy + h))
        out.append((dx, dy))
    return out


_WORK_VERBS = {'create': 'created', 'modify': 'changed', 'code': 'changed',
               'delete': 'deleted'}


def work_label(categories, action):
    """'Walls ×3 created' / 'Walls ×3 · Doors ×2 · +4 more changed' for the
    category names of the elements a call created or changed; '' for none."""
    order, counts = [], {}
    for name in categories or ():
        name = '{}'.format(name or '').strip() or 'Elements'
        if name not in counts:
            order.append(name)
        counts[name] = counts.get(name, 0) + 1
    if not order:
        return ''
    order.sort(key=lambda n: -counts[n])        # stable: ties keep call order
    parts = ['{} ×{}'.format(n, counts[n]) for n in order[:2]]
    rest = sum(counts[n] for n in order[2:])
    if rest:
        parts.append('+{} more'.format(rest))
    return '{} {}'.format(' · '.join(parts), _WORK_VERBS.get(action, 'touched'))


def merge_targets(named, added=(), modified=(), limit=MAX_TARGETS):
    """Element ids a call points at: the ones it names, then the ones Revit
    says it added; what it merely modified only when there is nothing else."""
    out = []
    for i in list(named) + list(added):
        if i not in out:
            out.append(i)
    if not out:
        out = list(modified)
    return out[:limit]


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


def clean_tasks(tasks):
    """A plan: up to PLAN_MAX one-line tasks of at most TASK_MAX characters."""
    if not isinstance(tasks, (list, tuple)):
        return []
    out = []
    for task in tasks:
        text = ' '.join('{}'.format(task or '').split())
        if text:
            out.append(text if len(text) <= TASK_MAX else text[:TASK_MAX - 1] + '…')
    return out[:PLAN_MAX]


def plan_rows(plan):
    """[(state, text)] with state done | current | todo, for the HUD."""
    tasks, step = plan.get('tasks') or [], plan.get('step') or 0
    return [('done' if i + 1 < step else 'current' if i + 1 == step else 'todo', t)
            for i, t in enumerate(tasks)]


def current_task(plan):
    """'Step 2/4 · Number the risers' while a step of the plan is running, else ''."""
    tasks, step = plan.get('tasks') or [], plan.get('step') or 0
    if not tasks or not 1 <= step <= len(tasks):
        return ''
    return 'Step {}/{} · {}'.format(step, len(tasks), tasks[step - 1])


def harness_task_text(ref, title):
    """'Task a · Check doors' for a harness task ('<session>/<key>'), else ''."""
    ref = u'{}'.format(ref or '')
    if not ref:
        return ''
    key = ref.split('/', 1)[-1]
    title = ' '.join(u'{}'.format(title or '').split())
    if len(title) > TASK_MAX:
        title = title[:TASK_MAX - 1] + u'…'
    return u'Task {} · {}'.format(key, title) if title else u'Task {}'.format(key)


def harness_dir():
    """= core/harness_store.harness_dir (that module is CPython-only)."""
    override = os.environ.get('T3LAB_HARNESS_DIR')
    if override:
        return override
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI', 'harness')


_board_cache = {'stamp': None, 'board': {}}


def read_board(path=None):
    """The harness board snapshot the bridge writes ({} when there is none).
    One stat per call; the file is parsed again only when it changed."""
    path = path or os.path.join(harness_dir(), BOARD_FILE)
    try:
        stamp = (path, os.path.getmtime(path), os.path.getsize(path))
    except (OSError, IOError):
        _board_cache.update(stamp=None, board={})
        return {}
    if stamp != _board_cache['stamp']:
        try:
            with io.open(path, encoding='utf-8') as f:
                board = json.load(f)
        except (OSError, IOError, ValueError):
            return _board_cache['board']        # mid-write: keep the last good one
        _board_cache.update(stamp=stamp, board=board if isinstance(board, dict) else {})
    return _board_cache['board']


_BOARD_STATE = {'claimed': 'current', 'review': 'review', 'failed': 'failed'}


def board_rows(board, limit=BOARD_ROWS_MAX):
    """Open harness sessions for the HUD: [{'session', 'goal', 'document',
    'summary', 'rows': [(state, text, document)], 'more'}]. rows = the tasks
    being worked on, waiting for review or failed; `more` = those not shown."""
    out = []
    for session in (board or {}).get('sessions') or []:
        tasks = session.get('tasks') or []
        if not tasks:
            continue
        counts = session.get('counts') or {}
        parts = []
        for status, label in (('claimed', 'working'), ('review', 'in review'),
                              ('failed', 'failed')):
            if counts.get(status):
                parts.append(u'{} {}'.format(counts[status], label))
        parts.append(u'{}/{} done'.format(counts.get('done', 0), len(tasks)))
        sdoc = session.get('document') or ''
        rows = []
        for task in tasks:
            state = _BOARD_STATE.get(task.get('status'))
            if not state:
                continue
            title = ' '.join(u'{}'.format(task.get('title') or '').split())
            if len(title) > TASK_MAX:
                title = title[:TASK_MAX - 1] + u'…'
            doc = task.get('document') or ''
            rows.append((state, u'{} · {}'.format(agent_style(task.get('owner') or '?')[0],
                                                    title),
                         doc if doc and doc.lower() != sdoc.lower() else ''))
        rows.sort(key=lambda r: ('current', 'review', 'failed').index(r[0]))
        out.append({'session': session.get('session') or '',
                    'goal': clean_goal(session.get('goal')), 'document': sdoc,
                    'summary': u' · '.join(parts), 'rows': rows[:limit],
                    'more': max(0, len(rows) - limit)})
    return out


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


# ─── Cursor tags, parking, crew (v3; plan in t3lab-revit-api dev/plan/) ─────

_TAG_STATE = {'waiting': '#F59E0B', 'failed': '#DC2626'}
_TAG_GROUND = {'light': '#FFFFFF', 'dark': '#1E2430'}
_TAG_INK = {'light': '#1F2937', 'dark': '#F8FAFC'}
TAG_WASH = {'light': 46, 'dark': 110}       # >= 4.5:1 with the ink for every colour


def tag_style(color, state='idle', theme='light'):
    """(fill, border, ink) of a cursor tag: the agent colour as a pastel wash,
    a 1 px border in the full colour and dark ink (light ink on a dark
    theme). White text on the solid colour was 2.2-3.6:1 for 9 of 16 colours
    (2026-11-25); this is 10:1 light, 6:1 dark. waiting = amber, failed = red."""
    theme = 'dark' if theme == 'dark' else 'light'
    base = _TAG_STATE.get(state, color)
    return blend(base, _TAG_GROUND[theme], TAG_WASH[theme]), base, _TAG_INK[theme]


_SINGULAR = {'Furniture': 'Furniture', 'Casework': 'Casework', 'Topography': 'Topography',
             'Toposolid': 'Toposolid', 'Mass': 'Mass', 'Parking': 'Parking',
             'Railings': 'Railing', 'Ceilings': 'Ceiling'}


def singular(category):
    """'Walls' -> 'Wall', 'Structural Columns' -> 'Structural Column',
    'Assemblies' -> 'Assembly'; the last word only."""
    words = '{}'.format(category or '').split()
    if not words:
        return ''
    last = words[-1]
    if last in _SINGULAR:
        last = _SINGULAR[last]
    elif last.endswith('ies') and len(last) > 4:
        last = last[:-3] + 'y'
    elif last.endswith('s') and not last.endswith('ss') and len(last) > 2:
        last = last[:-1]
    return ' '.join(words[:-1] + [last])


def object_caption(items):
    """What a call works on, as Mila's tags say it: 'Wall 3450 × 200' for one
    element, 'Walls ×3' for several of one category, '' for a mix (the call
    caption stays). items = [(category, size)]."""
    items = [(c or '', s or '') for c, s in items or ()]
    cats = set(c for c, _s in items)
    if not items or len(cats) != 1 or not items[0][0]:
        return ''
    cat = items[0][0]
    if len(items) == 1:
        return '{} {}'.format(singular(cat), items[0][1]).strip()
    return '{} \u00d7{}'.format(cat, len(items))


def home_slot(rect, slot, gap=44.0, margin=28.0):
    """Where a new cursor waits: down the left edge of rect (l, t, r, b), one
    slot per agent, the next column when the edge is full. Mila's cursors
    wait over empty paper; ours used to be born mid-view on top of the work."""
    left, top, right, bottom = rect
    rows = max(1, int((bottom - top - 2 * margin - 40) // gap))
    col, row = divmod(max(0, slot), rows)
    x = min(left + margin + col * 140.0, right - margin)
    return (x, top + margin + 40 + row * gap)


def spread_point(point, others, gap=SPREAD_GAP):
    """`point`, or the nearest spot round it at least `gap` px from every
    point in `others`: two agents on one element sit side by side."""
    import math
    others = list(others or ())

    def clear(p):
        return all(math.hypot(p[0] - o[0], p[1] - o[1]) >= gap for o in others)
    if clear(point):
        return point
    for ring in (1, 2, 3):
        for k in range(8):
            ang = math.pi / 4.0 * k + (math.pi / 8.0 if ring % 2 == 0 else 0.0)
            cand = (point[0] + ring * gap * math.cos(ang), point[1] + ring * gap * math.sin(ang))
            if clear(cand):
                return cand
    return point


def crew_groups(items, limit=CREW_MAX + 1):
    """[(category, [item])] of box items (dicts with 'category'), in the
    order the categories first appear; past `limit` groups the rest join the
    first (the agent's own cursor). One group = no crew."""
    order, groups = [], {}
    for item in items or ():
        cat = item.get('category') or ''
        if cat not in groups:
            order.append(cat)
            groups[cat] = []
        groups[cat].append(item)
    out = [(c, groups[c]) for c in order]
    if len(out) > limit:
        head, rest = out[:limit], out[limit:]
        head[0] = (head[0][0], head[0][1] + [i for _c, its in rest for i in its])
        out = head
    return out


def crew_key(agent, category):
    return '{}~crew~{}'.format(agent, category)


def is_crew(key):
    return '~crew~' in '{}'.format(key)


def crew_style(agent, category, k):
    """(tag, colour) of the k-th helper cursor (k >= 1): named after the work,
    never a made-up person ("Claude · doors"), in a shade of the agent colour."""
    tag, color = agent_style(agent)
    shade = (blend(color, '#000000', 190) if k % 2 else blend(color, '#FFFFFF', 200))
    work = '{}'.format(category or 'elements').strip().lower() or 'elements'
    return ('{} \u00b7 {}'.format(tag, work), shade)


# ─── HUD helpers (core/agent_cursors_overlay.py) ─────────────────────────────

# state -> (label, palette key of its colour)
STATES = {
    'working': ('Working', 'warn'),
    'waiting': ('Waiting for your OK', 'warn'),
    'paused': ('Paused', 'info'),
    'thinking': ('Waiting for the agent', 'info'),
    'idle': ('Idle', 'ok'),
}

# Bar colours per Revit theme. One orange accent (T3Lab / t3lab.space).
HUD_PALETTES = {
    'dark': {'bg': '#1E2430', 'card': '#273042', 'line': '#334155', 'chip': '#334155',
             'chip_hot': '#475569', 'text': '#F8FAFC', 'text2': '#CBD5E1', 'muted': '#94A3B8',
             'faint': '#64748B', 'accent': '#F27522', 'on_accent': '#FFFFFF', 'ok': '#4ADE80',
             'bad': '#F87171', 'warn': '#FBBF24', 'info': '#93C5FD', 'track': '#475569'},
    'light': {'bg': '#FFFFFF', 'card': '#F3F4F6', 'line': '#E5E7EB', 'chip': '#E5E7EB',
              'chip_hot': '#D1D5DB', 'text': '#111827', 'text2': '#374151', 'muted': '#6B7280',
              'faint': '#9CA3AF', 'accent': '#EA680C', 'on_accent': '#FFFFFF', 'ok': '#16A34A',
              'bad': '#DC2626', 'warn': '#B45309', 'info': '#2563EB', 'track': '#D1D5DB'},
}

# Options drawer (⚙): group -> option names, in OPTIONS terms.
OPTION_GROUPS = (
    ('Show', ('model',)),
    ('View', ('follow', 'zoom', 'tile')),
    ('Selection', ('select',)),
    ('Effects', ('glow', 'trail', 'ripple', 'labels', 'crew')),
    ('Bar', ('keep',)),
)


def hud_palette(theme):
    return HUD_PALETTES['light' if theme == 'light' else 'dark']


# ─── Action bands (the bar's colour blocks) ──────────────────────────────────
# A step shows as a block of its action's colour: a solid stripe, a wash of the
# same colour and a glyph. Text on it stays the bar's own text colour: green
# on white is 3.3:1, enough for a stripe, too weak for 12.5 px text.

BAR_W = 420.0           # width of the bar's content; fills are measured on it
BAND_STRIPE = 5.0       # px of solid colour on a band's left edge
BAND_WASH = {'dark': 56, 'light': 36}   # alpha of the colour behind a band
BAND_MAX = 3            # bands at once (agents working together), then "+n more"
SWEEP_W = 96.0          # block that crosses a band while its call runs
COUNT_BAR_W = 72.0      # an agent's counts as one stacked bar
_FAILED_GLYPH = ''
_WAITING_GLYPH = ''
_tones = {}


def _rgb(color):
    h = color.lstrip('#')
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _luminance(color):
    lin = []
    for v in _rgb(color):
        v = v / 255.0
        lin.append(v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4)
    return 0.2126 * lin[0] + 0.7152 * lin[1] + 0.0722 * lin[2]


def contrast(a, b):
    """WCAG contrast ratio of two '#RRGGBB' colours, 1.0 to 21.0."""
    la, lb = _luminance(a), _luminance(b)
    return (max(la, lb) + 0.05) / (min(la, lb) + 0.05)


def blend(color, onto, alpha):
    """`color` at `alpha` (0-255) over `onto`, as one opaque '#RRGGBB'."""
    t = max(0, min(255, alpha)) / 255.0
    return '#{:02X}{:02X}{:02X}'.format(*[
        int(round(c * t + o * (1.0 - t))) for c, o in zip(_rgb(color), _rgb(onto))])


def band_wash(theme):
    return BAND_WASH['light' if theme == 'light' else 'dark']


def bar_tone(color, theme, target=3.0):
    """`color` as the bar draws it: moved towards white (dark bar) or black
    (light bar) just far enough to read at `target`:1 on its own wash. Purple
    and blue are 2.5:1 on the dark bar as they are. The model overlay keeps
    the raw colour."""
    theme = 'light' if theme == 'light' else 'dark'
    key = (color, theme, target)
    if key not in _tones:
        bg = HUD_PALETTES[theme]['bg']
        pole = '#000000' if theme == 'light' else '#FFFFFF'
        tone = color
        for step in range(21):
            tone = blend(pole, color, int(round(255 * step / 20.0)))
            if contrast(tone, blend(tone, bg, BAND_WASH[theme])) >= target:
                break
        _tones[key] = tone
    return _tones[key]


def band_style(action, theme, state='done', failed=False):
    """(label, glyph, colour) of a band on the `theme` bar: the action's own
    label and colour; a failed step turns red with a cross, a call waiting
    for the user's OK turns amber. The label always stays the action's."""
    label, glyph, color = action_style(action)
    pal = hud_palette(theme)
    if failed:
        return label, _FAILED_GLYPH, bar_tone(pal['bad'], theme)
    if state == 'waiting':
        return label, _WAITING_GLYPH, bar_tone(pal['warn'], theme)
    return label, glyph, bar_tone(color, theme)


def band_rows(working=(), waiting=(), last=None, replay=None, limit=BAND_MAX):
    """What the bar's bands show, top to bottom: ([row], agents left out).
    working / waiting = [(agent, action, caption)] of the calls running now /
    waiting for the user's OK; last = the newest step {'agent', 'action',
    'caption', 'failed'}; replay = (k, n) when `last` is step k of a replay.
    Live calls win, one band per agent in agent order (a band never jumps
    when another agent starts); else the one last step.
    row = {'agent', 'action', 'caption', 'state', 'failed', 'note',
    'progress'}; state: running | waiting | replay | done; note = the text on
    the band's right; progress = (k, n) or None."""
    live = {}
    for state, calls in (('running', working), ('waiting', waiting)):
        for agent, action, caption in calls or ():
            live[agent] = {'agent': agent, 'action': action, 'caption': caption or '',
                           'state': state, 'failed': False, 'progress': None,
                           'note': 'Waiting' if state == 'waiting' else action_style(action)[0]}
    if live:
        rows = [live[agent] for agent in sorted(live)]
        return rows[:limit], max(0, len(rows) - limit)
    if not last:
        return [], 0
    row = {'agent': last.get('agent') or '', 'action': last.get('action'),
           'caption': last.get('caption') or '', 'failed': bool(last.get('failed')),
           'state': 'done', 'progress': None, 'note': action_style(last.get('action'))[0]}
    if replay:
        k, n = replay
        row.update(state='replay', progress=(k, n), note='Replay {} / {}'.format(k, n))
    return [row], 0


def band_motion(rows, seen):
    """([motion], signatures for the next draw): how each band_rows row comes
    on when the bar is drawn again. 'enter' = its agent had no band, 'shake' =
    its step failed, 'pop' = another step or state, None = nothing changed
    (the bar is redrawn on every hover: a plain redraw must not move).
    `seen` = {agent: signature} of the last draw."""
    motions, now = [], {}
    for row in rows:
        sig = (row['state'], row['action'], row['caption'], row['failed'], row['progress'])
        old = seen.get(row['agent'])
        if old is None:
            motions.append('enter')
        elif old == sig:
            motions.append(None)
        else:
            motions.append('shake' if row['failed'] else 'pop')
        now[row['agent']] = sig
    return motions, now


def fill_width(done, total, width=BAR_W, minimum=0.0):
    """px of `width` that `done` of `total` covers, within [minimum, width]."""
    try:
        part = float(done) / float(total)
    except (TypeError, ValueError, ZeroDivisionError):
        part = 0.0
    return min(float(width), max(float(minimum), width * min(1.0, max(0.0, part))))


def count_bar(counts, width=COUNT_BAR_W, gap=1.0, minimum=3.0):
    """[(action, px)] in ACTIONS order: an agent's counts as one stacked bar.
    Widths follow the counts, none under `minimum` (1 call in 30 still
    shows), and with the gaps between them they add up to `width`."""
    items = count_items(counts or {})
    if not items:
        return []
    room = width - gap * (len(items) - 1)
    if room <= minimum * len(items):
        return [(action, room / len(items)) for action, _n in items]
    pinned = set()
    while True:
        rest = float(sum(n for action, n in items if action not in pinned))
        free = room - minimum * len(pinned)
        widths = [(action, minimum if action in pinned else free * n / rest)
                  for action, n in items]
        small = [action for action, w in widths if w < minimum]
        if not small:
            return widths
        pinned.update(small)


def hud_state(working, waiting, asks, paused, thinking=False):
    """working | waiting | paused | thinking | idle — what the bar's status
    pill says. thinking = between two calls of a session: the agent's model
    is deciding its next step (Mila's "waiting for the model")."""
    if asks or waiting:
        return 'waiting'
    if paused:
        return 'paused'
    if working:
        return 'working'
    return 'thinking' if thinking else 'idle'


def last_step_text(seconds):
    """'just now' · '7 s ago' · '25 s ago' (5 s steps from 10 s: the bar
    redraws only when the text changes) · '3 min ago'."""
    s = max(0, int(seconds))
    if s < 2:
        return 'just now'
    if 10 <= s < 60:
        s = s // 5 * 5
    return '{} ago'.format(ago_text(s))


def status_text(state, seconds=None):
    """'Working 8 s' / 'Paused' / ..."""
    label = STATES[state][0]
    if state == 'working' and seconds is not None and seconds >= 1:
        return '{} {} s'.format(label, int(seconds))
    return label


def plan_progress(plan):
    """(done, total) of a t3lab_narrate plan, or None without tasks."""
    tasks, step = (plan or {}).get('tasks') or [], (plan or {}).get('step') or 0
    if not tasks:
        return None
    return max(0, min(step - 1, len(tasks))), len(tasks)


def ago_text(seconds):
    """'now' · '9 s' · '3 min' · '2 h'."""
    s = max(0, int(seconds))
    if s < 2:
        return 'now'
    if s < 60:
        return '{} s'.format(s)
    if s < 3600:
        return '{} min'.format(s // 60)
    return '{} h'.format(s // 3600)


def short_text(text, limit=44):
    text = ' '.join('{}'.format(text or '').split())
    return text if len(text) <= limit else text[:limit - 1] + '…'


def compact_line(caption, progress=None):
    """One-line bar text: 'Moving 12 walls · 3/5'."""
    text = short_text(caption, 36)
    if progress:
        text += ' · {}/{}'.format(min(progress[0] + 1, progress[1]), progress[1])
    return text


def count_items(counts):
    """[(action, n)] in ACTIONS order, zeros left out."""
    return [(key, counts[key]) for key, _l, _g, _c in ACTIONS if counts.get(key)]


def footer_text(total, failed, views):
    """'12 actions · 1 failed · 2 views'."""
    parts = ['{} action{}'.format(total, '' if total == 1 else 's')]
    if failed:
        parts.append('{} failed'.format(failed))
    if views:
        parts.append('{} view{}'.format(views, '' if views == 1 else 's'))
    return ' · '.join(parts)


def should_collapse(user_folded, idle_s, hovering, asks):
    """Fold the bar to one line: the user asked, or nothing happened for
    COLLAPSE_S; never while the mouse is on it. A question opens it unless the
    user folded it (the one-line bar carries Allow / Deny then)."""
    if hovering:
        return False
    if user_folded:
        return True
    return not asks and idle_s >= COLLAPSE_S


def parse_offset(text):
    try:
        dx, dy = ('{}'.format(text or '')).split(',')
        return float(dx), float(dy)
    except (TypeError, ValueError):
        return 0.0, 0.0


def hud_place(view, size, offset=(0.0, 0.0), margin=16.0):
    """(left, top) DIP of the bar: centred at the bottom of `view` (DIP
    left, top, right, bottom), moved by the user's offset, kept inside."""
    left, top, right, bottom = view
    w, h = size
    x = left + (right - left - w) / 2.0 + offset[0]
    y = bottom - h - margin + offset[1]
    x = min(max(x, left + 4), max(right - w - 4, left + 4))
    y = min(max(y, top + 4), max(bottom - h - 4, top + 4))
    return x, y


def union_box(boxes):
    """The box round [(x0, y0, x1, y1)]."""
    return (min(b[0] for b in boxes), min(b[1] for b in boxes),
            max(b[2] for b in boxes), max(b[3] for b in boxes))


def cluster_stops(n_boxes):
    """More boxes than the cursor visits: draw one frame round them all."""
    return n_boxes > MAX_STOPS


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
        # IsPerspective is a View3D property: a plan / section has none.
        if bool(getattr(view, 'IsPerspective', False)):
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


def _line_ends(elem, geo, persp, right, up):
    """Screen px (start, end) of a line-based element (wall, beam, grid,
    detail line) for the draw-along, else None."""
    try:
        curve = getattr(elem, 'Curve', None)        # grids, detail / model lines
        if curve is None:
            curve = getattr(elem.Location, 'Curve', None)
        if curve is None or not curve.IsBound:
            return None
        ends = []
        for k in (0, 1):
            p = _xyz(curve.GetEndPoint(k))
            ends.append(project_perspective(p, persp, geo['rect']) if persp else
                        project(p, geo['origin'], right, up, geo['w'], geo['h'], geo['rect']))
        if None in ends:
            return None
        return tuple(ends)
    except Exception:
        return None


def _dimension_value(elem):
    """'320' for a one-segment dimension (Mila: "cota 320"), else ''."""
    try:
        return '{}'.format(elem.ValueString or '') if elem.NumberOfSegments == 0 else ''
    except Exception:
        return ''


def _boxes(doc, geo, ids):
    """[{'id', 'box', 'size', 'label', 'category', 'line'}] of the ids shown
    in one view (orthographic or perspective), at most MAX_TARGETS. line =
    screen (start, end) of a line-based element, else None."""
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
        if len(out) >= MAX_TARGETS:
            break
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
            size = (_dimension_value(elem) if cat == 'Dimensions' else '') or size_label(w, h)
            out.append({'id': i, 'box': box, 'size': size, 'label': badge_text(cat, size),
                        'category': cat, 'line': _line_ends(elem, geo, persp, right, up)})
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


def _pre_snapshot(app, ids, views=None):
    """{view id: {'rect', 'origin', 'boxes'}} before a delete / move.
    views = an _open_views() result to reuse."""
    doc, geos, active = views or _open_views(app)
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
    # A harness member (one chat of a client) gets a cursor of its own.
    agent = getattr(task, 'member', None) or getattr(task, 'agent', None)
    audit = getattr(task, 'audit', None)
    agent = agent or getattr(audit, 'agent', None)
    if agent:
        return agent
    source = getattr(task, 'source', None) or getattr(audit, 'source', '')
    # No name at all: say "AI app", never guess one (2026-11-17: Antigravity
    # was shown as Claude).
    return 't3lab-assistant' if source == 'assistant' else 'mcp-client'


BRIDGE_AGENT = 't3lab-bridge'      # bridge.py's own probes (list_open_documents fan-out)


def _internal(task):
    """Calls that are plumbing, not a step of an AI app: no cursor, no count."""
    if getattr(task, 'agent', None) == BRIDGE_AGENT:
        return True
    return '{}'.format(getattr(task, 'tool_name', '')).startswith('__')


# set_parameter / bulk_set_parameter / rename_element: (name arg, value arg).
_PARAM_TOOLS = {
    'set_parameter': (('parameter_name', 'name'), 'value'),
    'bulk_set_parameter': (('parameter_name', 'name'), 'value'),
    'rename_element': ((), 'new_name'),
}
# Per Revit session: who edited what, the agents' goals, the select event.
_session = {'touched': {}, 'goals': {}, 'plans': {}, 'select_event': None, 'select_cls': None,
            'call': None, 'watch': None}
_WATCH_KEY = 'T3Lab.AgentCursors.DocumentChanged'


def _on_changed(sender, args):
    """DocumentChanged: ids the running call added / modified. Never raises."""
    call = _session.get('call')
    if call is None:
        return
    try:
        from Snippets._compat import eid_value
        for kind, ids in (('added', args.GetAddedElementIds()),
                          ('modified', args.GetModifiedElementIds())):
            keep = call[kind]
            for eid in ids:
                if len(keep) >= MAX_CAPTURE:
                    break
                value = eid_value(eid)
                if value not in keep:
                    keep.append(value)
    except Exception:
        pass


def _watch_changes(app):
    """Hook Application.DocumentChanged once per session (main thread). The
    delegate is kept on the AppDomain so a pyRevit reload detaches the old one."""
    if _session.get('watch') is not None:
        return
    _session['watch'] = False           # tried: never again this session
    try:
        from System import AppDomain, EventHandler
        from Autodesk.Revit.DB.Events import DocumentChangedEventArgs
        application = app.Application
        domain = AppDomain.CurrentDomain
        old = domain.GetData(_WATCH_KEY)
        if old is not None:
            try:
                application.DocumentChanged -= old
            except Exception:
                pass
        delegate = EventHandler[DocumentChangedEventArgs](_on_changed)
        application.DocumentChanged += delegate
        domain.SetData(_WATCH_KEY, delegate)
        _session['watch'] = delegate
    except Exception:
        pass


def _real_elements(doc, ids):
    """[(id, category name)] of the ids that are model / annotation elements
    with a box: sketches, types and views a call adds on the side drop out."""
    from Snippets._compat import make_eid
    out = []
    for i in ids:
        try:
            elem = doc.GetElement(make_eid(i))
            cat = elem.Category if elem is not None else None
            if cat is None or cat.Name.startswith('<'):
                continue
            kind = '{}'.format(getattr(cat, 'CategoryType', 'Model'))
            if kind not in ('Model', 'Annotation'):
                continue
            if elem.get_BoundingBox(None) is None and elem.get_BoundingBox(doc.ActiveView) is None:
                continue
            out.append((i, cat.Name))
        except Exception:
            continue
    return out


def _plan_of(agent):
    return _session['plans'].get(agent_key(agent) or 'mcp-client') or {}


def _task_text(task, agent):
    """What the cursor names: the agent's own plan step (t3lab_narrate), else
    the harness task the bridge stamped on the call."""
    return (current_task(_plan_of(agent))
            or harness_task_text(getattr(task, 'task_ref', None),
                                 getattr(task, 'task_title', None)))


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


def _window_rect(hwnd):
    """Screen rect (px) of a window handle, or None."""
    try:
        import ctypes
        from ctypes import wintypes
        r = wintypes.RECT()
        h = hwnd.ToInt64() if hasattr(hwnd, 'ToInt64') else int(hwnd)
        if ctypes.windll.user32.GetWindowRect(ctypes.c_void_p(h), ctypes.byref(r)):
            return (float(r.left), float(r.top), float(r.right), float(r.bottom))
    except Exception:
        pass
    return None


def _element_rect(el):
    from System.Windows import Point
    a = el.PointToScreen(Point(0, 0))
    b = el.PointToScreen(Point(el.ActualWidth, el.ActualHeight))
    return (min(a.X, b.X), min(a.Y, b.Y), max(a.X, b.X), max(a.Y, b.Y))


def _pane_container(title_el, max_w):
    """The pane round its title text: the first ancestor big enough to hold
    content (a ribbon "Properties" button never is: too low / too wide)."""
    from System.Windows import FrameworkElement
    from System.Windows.Media import VisualTreeHelper
    el = title_el
    for _ in range(14):
        el = VisualTreeHelper.GetParent(el)
        if el is None:
            return None
        if isinstance(el, FrameworkElement) and el.ActualHeight >= 150:
            return el if 120 <= el.ActualWidth <= max_w else None
    return None


def _wpf_panes(wanted, node, max_w, limit=9000):
    """{pane: {'rect', 'node_rect'}} found by their title in Revit's WPF trees
    (docked or floating). Main thread only; best effort."""
    from System.Windows import PresentationSource, FrameworkElement, Visibility
    from System.Windows.Controls import TextBlock
    from System.Windows.Media import VisualTreeHelper, Visual
    found, titles = {}, [(p, t) for p, t in PANE_TITLES if p in wanted]
    for src in list(PresentationSource.CurrentSources):
        try:
            if not src.Dispatcher.CheckAccess():
                continue            # the overlay's own windows live on its thread
            stack = [src.RootVisual]
        except Exception:
            continue
        seen = 0
        while stack and seen < limit and len(found) < len(titles):
            el = stack.pop()
            seen += 1
            if el is None or (isinstance(el, FrameworkElement)
                              and el.Visibility != Visibility.Visible):
                continue
            if isinstance(el, TextBlock) and el.IsVisible:
                text = el.Text or ''
                for pane, title in titles:
                    if pane in found or not (text == title or text.startswith(title + ' - ')):
                        continue
                    box = _pane_container(el, max_w)
                    if box is not None:
                        found[pane] = {'rect': _element_rect(box), 'el': box}
            try:
                for k in range(VisualTreeHelper.GetChildrenCount(el)):
                    child = VisualTreeHelper.GetChild(el, k)
                    if isinstance(child, Visual):
                        stack.append(child)
            except Exception:
                continue
    browser = found.get('browser')
    if browser is not None and node:
        browser['node_rect'] = _node_row(browser.pop('el'), node)
    for item in found.values():
        item.pop('el', None)
    return found


def _node_row(container, node, limit=4000):
    """Screen rect of the visible tree row whose text starts with `node`."""
    from System.Windows import FrameworkElement, Visibility
    from System.Windows.Controls import TextBlock
    from System.Windows.Media import VisualTreeHelper, Visual
    stack, seen, best = [container], 0, None
    while stack and seen < limit:
        el = stack.pop()
        seen += 1
        if isinstance(el, FrameworkElement) and el.Visibility != Visibility.Visible:
            continue
        if isinstance(el, TextBlock) and el.IsVisible and (el.Text or '').startswith(node):
            rect = _element_rect(el)
            if best is None or rect[1] < best[1]:
                best = rect
        try:
            for k in range(VisualTreeHelper.GetChildrenCount(el)):
                child = VisualTreeHelper.GetChild(el, k)
                if isinstance(child, Visual):
                    stack.append(child)
        except Exception:
            continue
    return best


def _shown_panes(app, wanted):
    """Panes Revit says are open (DockablePane.IsShown), or all when unknown."""
    try:
        from Autodesk.Revit.UI import DockablePanes
        ids = {'browser': DockablePanes.BuiltInDockablePanes.ProjectBrowser,
               'properties': DockablePanes.BuiltInDockablePanes.PropertiesPalette}
        return [p for p in wanted if app.GetDockablePane(ids[p]).IsShown()]
    except Exception:
        return list(wanted)


_PANES = {'at': 0.0, 'key': None, 'value': None}


def _locate_panes(app, route, mdi_rect):
    """The route's stops with screen rects and click points (px), cached
    PANE_CACHE_S; a stop whose pane is closed is dropped. Main thread."""
    wanted = sorted(set(st['pane'] for st in route))
    node = next((st['node'] for st in route if st['pane'] == 'browser'), None)
    main = _window_rect(_hwnd(app))
    key = (tuple(wanted), node, main, mdi_rect)
    now = time.time()
    if _PANES['key'] == key and now - _PANES['at'] < PANE_CACHE_S:
        panes = _PANES['value']
    else:
        shown = _shown_panes(app, wanted)
        max_w = (main[2] - main[0]) * 0.6 if main else 900.0
        try:
            panes = _wpf_panes(shown, node, max_w) if shown else {}
        except Exception:
            panes = {}
        guess = dock_estimate(main, mdi_rect, [p for p in shown if p not in panes])
        for pane, rect in guess.items():
            panes[pane] = {'rect': rect, 'estimated': True}
        _PANES.update(at=now, key=key, value=panes)
    out = []
    for st in route:
        found = panes.get(st['pane'])
        if found is None:
            continue
        stop = dict(st, rect=found['rect'], estimated=bool(found.get('estimated')))
        stop['point'] = pane_point(found['rect'], st.get('node') if st['pane'] == 'browser'
                                   else None, found.get('node_rect'))
        out.append(stop)
    return out


def _ui_stops(app, tool, args, action, shown_rects, categories=()):
    """ui_route located on screen; [] when there is nothing to walk through."""
    try:
        route = ui_route(tool, args, action, categories)
        if not route or not shown_rects:
            return []
        return _locate_panes(app, route, union_rect(shown_rects))
    except Exception:
        return []


def before(app, task):
    """Ahead of the tool: show the agent working, remember the boxes a delete /
    move will change and the values an edit will overwrite. Never raises;
    costs one mtime check while the feature is off."""
    try:
        if not is_enabled() or _internal(task):
            return
        _flush_pending()        # the previous step's show comes first
        tool = getattr(task, 'tool_name', '')
        args = getattr(task, 'arguments', None)
        action = classify_action(tool, args)
        ids = extract_ids(args)
        _watch_changes(app)
        _session['call'] = {'added': [], 'modified': []}
        try:
            views = _open_views(app)
            doc, geos, active = views
            if active is not None:
                # Elements the call names: framed while it runs, cursor on them.
                # Bar only (the default): no boxes to measure.
                boxes = []
                model = on_model()
                if model and ids and action in ('delete', 'modify'):
                    task._cursor_pre = _pre_snapshot(app, ids, views)
                    boxes = task._cursor_pre[geos[active]['id']]['boxes']
                elif model and ids:
                    boxes = _boxes(doc, geos[active], ids)
                shown = [geos[i]['rect'] for i in visible_views([g['rect'] for g in geos], active)]
                # Project Browser / Properties on the way, like a person would
                # (on the model only: the bar draws no cursor).
                ui = task._cursor_ui = _ui_stops(app, tool, args, action, shown) if model else []
                agent = _agent_of(task)
                _overlay(_hwnd(app)).call(
                    'working', agent, action, describe_call(tool, args, len(ids)), tool,
                    geos[active]['rect'], task_text=_task_text(task, agent),
                    targets=[b['box'] for b in boxes[:MAX_STOPS]],
                    canvas=union_rect(shown + [st['rect'] for st in ui]), ui=ui,
                    subject=object_caption([(b['category'], b['size']) for b in boxes]))
            elif _state.get('overlay') is not None:
                # No view to draw in (no document, a view window too small):
                # the bar, once it is up, still shows the call.
                agent = _agent_of(task)
                _state['overlay'].call(
                    'working', agent, action, describe_call(tool, args, len(ids)), tool, None,
                    task_text=_task_text(task, agent))
        except Exception:
            pass
        name, ok = _param_name(tool, args)
        if ok and ids:
            task._cursor_old = (name, _read_values(app, ids, name))
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


def _done(task, agent, action, tool, failed, args):
    """End a call on the bar when notify() had nothing to draw or failed.
    Only when the bar exists: a call never brings it up without a view."""
    try:
        ov = _state.get('overlay')
        if ov is not None and ov.alive:
            ov.call('done', agent, action, tool, failed, args=args,
                    goal=_session['goals'].get(agent_key(agent) or 'mcp-client', ''),
                    task_text=_task_text(task, agent))
    except Exception:
        pass


def notify(app, task):
    """One finished tool call. Never raises. Whatever happens after before()
    said Working, the bar is told the call is over: show(), or done() when
    there is nothing to draw or this function failed half way (2026-11-19:
    such a call left the bar on "Working" for ever). finish() takes that
    over once it is handed the call, also when it runs later (_when_settled)."""
    agent = action = tool = args = None
    failed, started, handed = False, False, False
    try:
        _flush_pending()        # keep the steps in order
        if not is_enabled():
            hide()              # switched off from another engine (Settings)
            return
        if _internal(task):
            return                  # internal calls (status probes)
        tool = getattr(task, 'tool_name', '')
        args = getattr(task, 'arguments', None)
        result = getattr(task, 'result', None)
        if isinstance(result, dict) and result.get('error_code') == 'paused':
            _overlay(_hwnd(app)).call('held', _agent_of(task), describe_call(tool, args))
            return
        failed = bool(getattr(task, 'exception', None)) or (
            isinstance(result, dict) and bool(result.get('error')))
        action = classify_action(tool, args, result)
        ids = [] if failed else extract_ids(args, result)
        agent = _agent_of(task)
        started = True
        opts = effective(options())     # bar only: nothing below touches the model
        uidoc = app.ActiveUIDocument
        # What Revit says the call added / modified (code returns no ids).
        call, _session['call'] = _session.get('call'), None
        fresh, label, touch, made = [], '', action, []
        if call and not failed and uidoc is not None:
            added = _real_elements(uidoc.Document, call['added'])
            changed = [] if ids or added else _real_elements(uidoc.Document, call['modified'])
            if added or changed:
                touch = 'create' if added else 'modify'
                made = [c for _i, c in added or changed]
                label = work_label(made, touch)
                ids = merge_targets(ids, [i for i, _c in added], [i for i, _c in changed])
                fresh = [i for i, _c in added or changed if i in ids]
            elif action == 'create' and ids:
                named = _real_elements(uidoc.Document, ids)
                label = work_label([c for _i, c in named], 'create')
                fresh = [i for i, _c in named]
        followed = []
        pre = getattr(task, '_cursor_pre', None)
        if uidoc is not None and not failed:
            if opts['follow']:
                followed = _follow_views(uidoc, task, action, ids)
                if followed:
                    pre = None      # different view now: old boxes do not apply
            if opts['select'] and ids and action != 'delete':
                _select(uidoc, ids)
        if uidoc is None:
            # No document: nothing to draw (finally: done). Before Glow, which
            # cannot draw without one and would switch itself off for good.
            return
        tiled = False
        if opts['tile'] and any(new for _v, _n, new in followed):
            if len(list(uidoc.GetOpenUIViews())) >= 2:
                tiled = _tile(app)
        changes = [] if failed else _changes(app, task, ids)
        clashes = [] if failed else track_touches(_session['touched'], agent, ids, touch)
        _ensure_select_event()
        try:
            from core import agent_cursors_glow
            if opts['glow']:
                if not agent_cursors_glow.update(app, sorted(_session['touched'])):
                    # Revit refused the DirectContext3D server: switch Glow off
                    # and say so, instead of failing on every call.
                    set_option('glow', False)
                    agent_cursors_glow.stop(app)
                    _overlay(_hwnd(app)).call(
                        'notice', 'Glow turned off: this Revit could not draw it.')
            else:
                agent_cursors_glow.stop(app)        # no-op when never started
        except Exception:
            pass
        _sync_theme()
        touched = len(_session['touched'])
        goal = _session['goals'].get(agent_key(agent) or 'mcp-client', '')
        task_text = _task_text(task, agent)

        def finish():
            ended = False
            try:
                if not is_enabled():
                    ended = True        # switched off meanwhile: the bar is gone
                    return
                # Bar only: the views place the bar; no element boxes to measure.
                snap = _snapshot(app, ids if opts['model'] else [], pre)
                if snap is None:
                    return
                if (opts['zoom'] and ids and not snap['any_box'] and action != 'delete'
                        and uidoc is not None and _zoom_to(uidoc, ids)):
                    snap = _snapshot(app, ids) or snap
                ui = getattr(task, '_cursor_ui', None)
                if not ui and opts['model']:
                    # Code names no tool: what it made says where a person would go.
                    ui = _ui_stops(app, tool, args, action,
                                   [g['rect'] for g in snap['views']], sorted(set(made)))
                _overlay(_hwnd(app)).call(
                    'show', agent, action, tool, failed, snap, followed, args=args,
                    n_targets=len(ids), changes=changes, clashes=clashes,
                    touched=touched, goal=goal, task_text=task_text,
                    created=label, fresh=fresh, ui=ui or [])
                ended = True
            except Exception:
                pass
            finally:
                if not ended:
                    _done(task, agent, action, tool, failed, args)

        # Opening views and the posted Tile land AFTER this event returns:
        # measured now, every box and the cursor sat where the views used to
        # be. Measure once Revit has laid the views out.
        handed = True
        if (followed or tiled) and _when_settled(app, finish):
            return
        finish()
    except Exception:
        pass
    finally:
        if started and not handed:
            _done(task, agent, action, tool, failed, args)


# ─── Deferred show: wait for Revit to finish laying views out ────────────────

SETTLE_MAX_TICKS = 8        # Idling passes before showing anyway
_pending = {'finish': None, 'handler': None, 'app': None}


def _flush_pending():
    """Run a deferred show now (a new call starts, so steps stay in order).
    Never raises."""
    finish, handler, app = _pending['finish'], _pending['handler'], _pending['app']
    _pending.update(finish=None, handler=None, app=None)
    if handler is not None and app is not None:
        try:
            app.Idling -= handler
        except Exception:
            pass
    if finish is not None:
        try:
            finish()
        except Exception:
            pass


def _when_settled(app, finish):
    """Run finish() on a later Idling, once two passes in a row see the same
    view layout (layout_key), or after SETTLE_MAX_TICKS. False when Idling
    could not be hooked (the caller shows at once)."""
    _flush_pending()
    state = {'key': None, 'ticks': 0}

    def on_idling(sender, args):
        if _pending['handler'] is not on_idling:
            return              # already flushed by a newer call
        state['ticks'] += 1
        try:
            _doc, geos, active = _open_views(app)
            key = layout_key(geos, active)
        except Exception:
            key = None
        settled = key is not None and key == state['key']
        state['key'] = key
        if settled or state['ticks'] >= SETTLE_MAX_TICKS:
            _flush_pending()
            return
        try:
            args.SetRaiseWithoutDelay()     # next pass now, not in a second
        except Exception:
            pass

    try:
        _pending.update(finish=finish, handler=on_idling, app=app)
        app.Idling += on_idling
        return True
    except Exception:
        _pending.update(finish=None, handler=None, app=None)
        return False


# ─── Narration / approval wait (any thread) ────────────────────────────────────

NARRATE_TOOL = 't3lab_narrate'
NARRATE_SCHEMA = {
    'name': NARRATE_TOOL,
    'description': ('Show the user what you are doing in Revit: a goal and the list of tasks, '
                    'over the model while Agent Cursors are on; changes nothing. Before a '
                    'multi-step job send message + tasks (short steps in plain words); then '
                    'step=N as you start task N, and done=true at the end. The cursor shows '
                    'the current task instead of the tool name.'),
    'inputSchema': {
        'type': 'object',
        'properties': {
            'message': {'type': 'string',
                        'description': 'The goal, one line (max {} characters).'.format(NARRATE_MAX)},
            'tasks': {'type': 'array', 'items': {'type': 'string'},
                      'description': 'The plan: up to {} short tasks, in order.'.format(PLAN_MAX)},
            'step': {'type': 'integer', 'description': 'The task you start now (1 = first).'},
            'done': {'type': 'boolean', 'description': 'Every task is finished.'},
        },
        'required': [],
    },
}


def narrate(agent, message=None, tasks=None, step=None, done=False):
    """t3lab_narrate: set the agent's goal and plan. Any thread; never raises.
    Returns the tool result."""
    key = agent_key(agent) or 'mcp-client'
    if message is not None:
        _session['goals'][key] = clean_goal(message)
    plan = dict(_session['plans'].get(key) or {})
    new_tasks = clean_tasks(tasks)
    if new_tasks:
        plan = {'tasks': new_tasks, 'step': 1}
    if step is not None:
        try:
            plan['step'] = max(0, int(step))
        except (TypeError, ValueError):
            pass
    if done and plan.get('tasks'):
        plan['step'] = len(plan['tasks']) + 1
    if plan.get('tasks'):
        _session['plans'][key] = plan
    goal = _session['goals'].get(key, '')
    try:
        if is_enabled() and _state.get('overlay') is not None:
            _state['overlay'].call('narrate', agent or 'mcp-client', goal, dict(plan))
    except Exception:
        pass
    if not goal and not plan.get('tasks'):
        return {'success': False, 'error': 'Send message (the goal) and/or tasks.',
                'error_code': 'bad_arguments', 'nothing_ran': True}
    out = {'success': True, 'goal': goal, 'shown': is_enabled(),
           'note': ('Shown over the model.' if is_enabled() else
                    'Agent Cursors are off in Revit; the goal is kept for when they are on.')}
    if plan.get('tasks'):
        out['tasks'] = plan['tasks']
        out['step'] = plan['step']
        out['current'] = current_task(plan) or ('all done' if plan['step'] > len(plan['tasks'])
                                                else '')
    return out


def waiting(agent, tool, arguments=None, on=True):
    """A call waits for the user's OK (approval dialog): pulse the cursor.
    Any thread; never raises."""
    try:
        if is_enabled() and _state.get('overlay') is not None:
            _state['overlay'].call('waiting', agent or 'mcp-client', tool,
                                   describe_call(tool, arguments), bool(on))
    except Exception:
        pass


# ─── Select touched (HUD button -> ExternalEvent -> main thread) ──────────────

def touched_ids():
    return sorted(_session['touched'])


from Snippets._compat import clr_namespace   # per-load .NET type name (CLAUDE.md rule 5)
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
                __namespace__ = clr_namespace('AgentCursorsSelect')

                def Execute(self, uiapp):
                    actions, _session['ui_actions'] = _session.get('ui_actions') or [], []
                    for action in actions or ['select']:
                        try:
                            if action == 'idle':        # agents went quiet: Glow off
                                from core import agent_cursors_glow
                                agent_cursors_glow.stop(uiapp)
                            else:
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


def _raise_ui(action):
    """Run `action` ('select' | 'idle') on Revit's main thread. Any thread."""
    event = _session.get('select_event')
    if event is not None:
        try:
            _session.setdefault('ui_actions', []).append(action)
            event.Raise()
            return True
        except Exception:
            pass
    return False


def select_touched():
    """HUD button: select every element the agents changed. Any thread."""
    return _raise_ui('select')


def agents_idle():
    """Overlay: the agents went quiet. Glow is taken off the model (it came
    back on the next call); owner, 2026-11-17: it stayed after the agent stopped."""
    return _raise_ui('idle')


def hide():
    """Hide the overlay now (no-op when it was never shown)."""
    ov = _state.get('overlay')
    if ov is not None:
        try:
            ov.call('hide')
        except Exception:
            pass


# ─── Consent on the bar + the offer to turn cursors on (T3Lab Lite) ──────────

OFFER_KEY = 'agent_cursors_offer'       # 'never' = do not offer again
_asks = {}                              # request id -> on_answer(approved, always)
_ask_seq = [0]


def ask(app, agent, tool, arguments, main, content, expanded='', always_text='',
        on_answer=None):
    """Main thread: put a consent question on the bar (Allow / Deny) instead of
    a modal Revit dialog, so Revit stays usable. on_answer(approved, always)
    runs on the overlay thread. The request id when shown (cancel_ask takes
    it); '' = the caller asks its own way (cursors off, no view, overlay down)."""
    if on_answer is None or not is_enabled():
        return ''
    rid = None
    try:
        _doc, geos, active = _open_views(app)
        if active is None:
            return ''
        ov = _overlay(_hwnd(app))
        if not ov.alive:
            return ''
        _ask_seq[0] += 1
        rid = 'ask{}'.format(_ask_seq[0])
        _asks[rid] = on_answer
        ov.call('ask', rid, agent or 'mcp-client', tool,
                describe_call(tool, arguments), '{}'.format(main or ''),
                '{}'.format(content or ''), '{}'.format(expanded or ''),
                '{}'.format(always_text or ''), geos[active]['rect'])
        return rid
    except Exception:
        if rid is not None:
            _asks.pop(rid, None)
        return ''


def cancel_ask(rid):
    """The caller stopped waiting: take the card off the bar, no answer."""
    if _asks.pop(rid, None) is not None and _state.get('overlay') is not None:
        try:
            _state['overlay'].call('cancel_ask', rid)
        except Exception:
            pass


def answer(rid, approved, always=False):
    """Overlay thread: Allow / Deny clicked. Never raises."""
    fn = _asks.pop(rid, None)
    if fn is not None:
        try:
            fn(bool(approved), bool(always))
        except Exception:
            pass


def pending_asks():
    return len(_asks)


def offer_text(agent):
    """(main instruction, content) of the one-time offer."""
    tag = agent_style(agent)[0]
    return (u'Follow {} in Revit?'.format(tag),
            u'{} is working in this model through MCP. Agent Cursors shows each step on '
            u'the model with its task list, and asks for your OK on its bar at the bottom '
            u'of the view instead of in a dialog, so Revit stays free while the AI '
            u'works.'.format(tag))


def offer(app, agent):
    """Main thread, an MCP call from `agent` with cursors off: offer once per
    Revit session (never again after "Don't ask again"). True when the user
    turned them on. Never raises."""
    if is_enabled() or _state.get('offered') or not agent or agent == BRIDGE_AGENT:
        return False
    _state['offered'] = True
    try:
        if '{}'.format(_settings().get(OFFER_KEY, '')).lower() == 'never':
            return False
        from Autodesk.Revit.UI import (TaskDialog, TaskDialogCommandLinkId,
                                       TaskDialogResult)
        main, content = offer_text(agent)
        dlg = TaskDialog('T3Lab Agent Cursors')
        dlg.TitleAutoPrefix = False
        dlg.MainInstruction = main
        dlg.MainContent = content
        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, 'Turn on Agent Cursors',
                           'You can turn them off with the x on the bar.')
        dlg.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, 'Not now',
                           'Confirmations stay in Revit dialogs.')
        dlg.VerificationText = "Don't ask again"
        turned_on = dlg.Show() == TaskDialogResult.CommandLink1
        if dlg.WasVerificationChecked() and not turned_on:
            _paths().set_setting(OFFER_KEY, 'never')
            _cache['mtime'] = None
        if turned_on:
            set_enabled(True)
        return turned_on
    except Exception:
        return False


# ─── Overlay (core/agent_cursors_overlay.py) ──────────────────────────────────

_state = {'overlay': None, 'offered': False, 'theme': None}


def revit_theme():
    """'dark' | 'light' as Revit shows it (UIThemeManager, Revit 2024+;
    earlier Revit has only light). Main thread."""
    try:
        import Autodesk.Revit.UI as RUI
        mgr = getattr(RUI, 'UIThemeManager', None)
        if mgr is not None and 'dark' in '{}'.format(mgr.CurrentTheme).lower():
            return 'dark'
    except Exception:
        pass
    return 'light'


def _sync_theme():
    theme = revit_theme()
    if theme != _state.get('theme') and _state.get('overlay') is not None:
        _state['theme'] = theme
        _state['overlay'].call('set_theme', theme)


def _overlay(hwnd):
    """The overlay host (its own UI thread under IronPython, see
    agent_cursors_overlay.OverlayHost); .call(method, ...) posts to it."""
    ov = _state.get('overlay')
    if ov is None or not ov.alive:
        from core.agent_cursors_overlay import OverlayHost
        ov = OverlayHost(hwnd, {'select_touched': select_touched, 'answer': answer,
                                'idle': agents_idle,
                                'pause': set_paused})
        _state['overlay'] = ov
        _state['theme'] = None
        _sync_theme()
    return ov
