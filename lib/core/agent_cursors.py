# -*- coding: utf-8 -*-
"""Agent cursors: a live overlay of which AI agent does what, where.

Each MCP tool call moves one cursor per (agent, action) — Claude · Create,
Claude · Modify, GPT · Read … — onto the elements it read, edited, created or
deleted, in every visible view (tiled views included), with an effect per
action. Follow mode opens the views the agent works in; zoom brings off-screen
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
import time

SETTING_KEY = 'agent_cursors_enabled'
MAX_TARGETS = 12            # bounding boxes per call: keeps a bulk edit cheap
MAX_FOLLOW_VIEWS = 4        # views follow mode may open per call
IDLE_HIDE_S = 12            # overlay fades out after this long without a call
BOX_HOLD_S = 3.5            # element boxes stay this long, then fade
MOVE_MS = 650               # cursor glide

# name -> (mcp_paths.json key, default). The HUD toggles them live.
OPTIONS = (
    ('follow', 'agent_cursors_follow_views', True),
    ('zoom', 'agent_cursors_zoom', True),
    ('select', 'agent_cursors_select', False),
    ('tile', 'agent_cursors_tile_views', False),
)
OPTION_LABELS = {
    'follow': ('Follow', 'Open and switch to the views the agent works in'),
    'zoom': ('Zoom', 'Zoom the view to targets that are off screen'),
    'select': ('Select', "Select the agent's targets (replaces your selection)"),
    'tile': ('Tile', 'Tile the open views when the agent opens another one'),
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
    "One cursor per action (Read, Create, Modify, Delete, View, Export, Code), each with its own effect.",
    "Follow mode opens the views the agent works in; boxes show in every tiled view.",
    "Spot a wrong target early and stop the agent before it goes further.",
    "Several agents at once each get their own colour and name tag.",
)
CONS = (
    "Each tool call does a little extra work (bounding boxes of up to {} elements, twice for a delete or move).".format(MAX_TARGETS),
    "Follow switches the active view, so a later tool that works on 'the active view' works on the view the agent opened.",
    "Zoom and Select change your camera and selection while an agent works (Select is off by default).",
    "Boxes are a snapshot: pan or zoom while an agent works and they sit off until its next call.",
    "No boxes in perspective 3D views; the overlay shows up in screenshots and recordings.",
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


def counts_text(counts):
    """{'create': 3, 'read': 1} -> 'Read 1 · Create 3' (ACTIONS order)."""
    parts = []
    for key, label, _g, _c in ACTIONS:
        if counts.get(key):
            parts.append('{} {}'.format(label, counts[key]))
    return ' · '.join(parts)


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
        if bool(view.IsPerspective) or '{}'.format(view.ViewType) in _NO_BOXES:
            return geo
    except Exception:
        pass
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
    if not geo.get('ortho') or not ids:
        return out
    from Snippets._compat import make_eid
    view = geo['view']
    shown = _shown_in(doc, view, ids)
    right, up = geo['right'], geo['up']
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
            box = screen_box(pts, geo['origin'], right, up, geo['w'], geo['h'], geo['rect'])
            if box is None:
                continue
            ext = (hi.X - lo.X, hi.Y - lo.Y, hi.Z - lo.Z)
            w = abs(sum(ext[k] * abs(right[k]) for k in range(3)))
            h = abs(sum(ext[k] * abs(up[k]) for k in range(3)))
            out.append({'id': i, 'box': box, 'size': size_label(w, h)})
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


def before(app, task):
    """Ahead of the tool: remember the boxes a delete / move will change.
    Never raises; costs one mtime check while the feature is off."""
    try:
        if not is_enabled() or _internal(task):
            return
        args = getattr(task, 'arguments', None)
        action = classify_action(getattr(task, 'tool_name', ''), args)
        if action not in ('delete', 'modify'):
            return
        ids = extract_ids(args)
        if ids:
            task._cursor_pre = _pre_snapshot(app, ids)
    except Exception:
        pass


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
        hwnd = None
        try:
            hwnd = app.MainWindowHandle
        except Exception:
            pass
        _overlay(hwnd).show(_agent_of(task), action, tool, failed, snap, followed)
    except Exception:
        pass


def hide():
    """Hide the overlay now (no-op when it was never shown)."""
    ov = _state.get('overlay')
    if ov is not None:
        try:
            ov.hide()
        except Exception:
            pass


# ─── WPF overlay ───────────────────────────────────────────────────────────────

_state = {'overlay': None}


def _overlay(hwnd):
    ov = _state.get('overlay')
    if ov is None or not ov.alive:
        ov = _Overlay(hwnd)
        _state['overlay'] = ov
    return ov


def _NO_STYLE(enum):
    # `WindowStyle.None` is a syntax error in Python.
    return getattr(enum, 'None')


_GWL_EXSTYLE = -20
_WS_EX_TRANSPARENT = 0x20
_WS_EX_TOOLWINDOW = 0x80
_WS_EX_NOACTIVATE = 0x08000000
_MAX_TRANSIENT = 400        # animation handlers kept alive (see _hook)


def _set_exstyle(window, bits):
    """Add extended window styles (click-through / no focus); best effort."""
    try:
        import ctypes
        from System.Windows.Interop import WindowInteropHelper
        hwnd = WindowInteropHelper(window).Handle.ToInt64()
        user32 = ctypes.windll.user32
        style = user32.GetWindowLongW(ctypes.c_void_p(hwnd), _GWL_EXSTYLE)
        user32.SetWindowLongW(ctypes.c_void_p(hwnd), _GWL_EXSTYLE, style | bits)
    except Exception:
        pass


def _opacity_property():
    from System.Windows import UIElement
    return UIElement.OpacityProperty


class _Overlay(object):
    """Two owned, non-activating windows: a click-through canvas over the
    visible views for cursors, boxes and effects, and a small HUD with the
    agents, their action counts, the views they opened and live toggles."""

    def __init__(self, owner_hwnd):
        import clr
        for ref in ('PresentationCore', 'PresentationFramework', 'WindowsBase'):
            try:
                clr.AddReference(ref)
            except Exception:
                pass
        from System.Windows import Window, WindowStyle, ResizeMode, SizeToContent
        from System.Windows.Controls import Canvas
        from System.Windows.Media import Brushes

        self.alive = True
        self._owner = owner_hwnd
        self._cursors = {}          # 'agent|action' -> dict(group, tx, ...)
        self._counts = {}           # agent -> {action: n}
        self._views = []            # [(view id, name)] the agents worked in
        self._actions = 0
        self._last = 0.0
        self._scale = None
        self._origin = (0.0, 0.0)   # screen px of the canvas top-left
        self._handlers = []         # strong refs: delegates must outlive the GC
        self._transient = []

        self.win = Window()
        self.win.WindowStyle = _NO_STYLE(WindowStyle)
        self.win.AllowsTransparency = True
        self.win.Background = Brushes.Transparent
        self.win.ResizeMode = ResizeMode.NoResize
        self.win.ShowInTaskbar = False
        self.win.ShowActivated = False
        self.win.Focusable = False
        self.win.IsHitTestVisible = False
        self.win.Title = 'T3Lab Agent Cursors'
        self.canvas = Canvas()
        self.canvas.IsHitTestVisible = False
        self.win.Content = self.canvas
        self._own(self.win)
        self._hook(self.win, 'SourceInitialized',
                   lambda s, e: _set_exstyle(self.win, _WS_EX_TRANSPARENT |
                                             _WS_EX_TOOLWINDOW | _WS_EX_NOACTIVATE))
        self._hook(self.win, 'Closed', lambda s, e: setattr(self, 'alive', False))

        self.hud = Window()
        self.hud.WindowStyle = _NO_STYLE(WindowStyle)
        self.hud.AllowsTransparency = True
        self.hud.Background = Brushes.Transparent
        self.hud.ResizeMode = ResizeMode.NoResize
        self.hud.ShowInTaskbar = False
        self.hud.ShowActivated = False
        self.hud.SizeToContent = SizeToContent.WidthAndHeight
        self.hud.Title = 'T3Lab Agent Cursors'
        self._own(self.hud)
        self._hook(self.hud, 'SourceInitialized',
                   lambda s, e: _set_exstyle(self.hud, _WS_EX_TOOLWINDOW | _WS_EX_NOACTIVATE))
        self._build_hud()

        from System import TimeSpan
        from System.Windows.Threading import DispatcherTimer
        self._timer = DispatcherTimer()
        self._timer.Interval = TimeSpan.FromSeconds(1)
        self._hook(self._timer, 'Tick', self._on_tick)

    # ── plumbing ──────────────────────────────────────────────────────────────

    def _hook(self, source, event, fn, transient=False):
        keep = self._transient if transient else self._handlers
        keep.append(fn)
        if transient and len(keep) > _MAX_TRANSIENT:
            del keep[:len(keep) - _MAX_TRANSIENT]
        handler = getattr(source, event)
        handler += fn
        setattr(source, event, handler)

    def _own(self, window):
        if not self._owner:
            return
        try:
            from System.Windows.Interop import WindowInteropHelper
            WindowInteropHelper(window).Owner = self._owner
        except Exception:
            pass

    @staticmethod
    def _brush(hex_color, alpha=255):
        from System.Windows.Media import SolidColorBrush, Color
        h = hex_color.lstrip('#')
        return SolidColorBrush(Color.FromArgb(alpha, int(h[0:2], 16),
                                              int(h[2:4], 16), int(h[4:6], 16)))

    def _dpi(self):
        if self._scale:
            return self._scale
        try:
            from System.Windows import PresentationSource
            src = PresentationSource.FromVisual(self.win)
            self._scale = float(src.CompositionTarget.TransformToDevice.M11) or 1.0
        except Exception:
            return 1.0
        return self._scale

    def _local(self, box):
        """Screen px box -> canvas DIP (x0, y0, x1, y1)."""
        s = self._dpi()
        ox, oy = self._origin
        return ((box[0] - ox) / s, (box[1] - oy) / s, (box[2] - ox) / s, (box[3] - oy) / s)

    @staticmethod
    def _text(text, size=11, bold=True, color='#FFFFFF', glyph=False):
        from System.Windows import FontWeights
        from System.Windows.Controls import TextBlock
        from System.Windows.Media import FontFamily
        tb = TextBlock()
        tb.Text = text
        tb.FontFamily = FontFamily('Segoe MDL2 Assets' if glyph else 'Segoe UI')
        tb.FontSize = size
        tb.FontWeight = FontWeights.SemiBold if bold and not glyph else FontWeights.Normal
        tb.Foreground = _Overlay._brush(color)
        return tb

    @staticmethod
    def _pill(child, color, radius=4, alpha=255):
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border
        b = Border()
        b.Background = _Overlay._brush(color, alpha)
        b.CornerRadius = CornerRadius(radius)
        b.Padding = Thickness(6, 2, 6, 2)
        b.Child = child
        return b

    def _row(self, *children):
        from System.Windows import VerticalAlignment
        from System.Windows.Controls import StackPanel, Orientation
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        for child in children:
            child.VerticalAlignment = VerticalAlignment.Center
            row.Children.Add(child)
        return row

    def _glyph_label(self, glyph, label, color, size=10):
        from System.Windows import Thickness
        icon = self._text(glyph, size, color=color, glyph=True)
        icon.Margin = Thickness(0, 0, 4, 0)
        return self._row(icon, self._text(label, size, color=color))

    def _anim(self, frames, done=None):
        """DoubleAnimationUsingKeyFrames from [(value, ms)]."""
        from System import TimeSpan
        from System.Windows.Media.Animation import (
            DoubleAnimationUsingKeyFrames, LinearDoubleKeyFrame, KeyTime)
        anim = DoubleAnimationUsingKeyFrames()
        for value, ms in frames:
            anim.KeyFrames.Add(LinearDoubleKeyFrame(
                value, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(ms))))
        if done is not None:
            self._hook(anim, 'Completed', done, transient=True)
        return anim

    def _pulse(self, element, delay_ms, hold_ms, done=None):
        """Fade in at delay, hold, fade out — one animation: a second
        BeginAnimation on Opacity would replace the first."""
        element.BeginAnimation(_opacity_property(), self._anim(
            ((0.0, delay_ms), (1.0, delay_ms + 180), (1.0, delay_ms + 180 + hold_ms),
             (0.0, delay_ms + 780 + hold_ms)), done))

    def _remover(self, element):
        def remove(s, e):
            try:
                self.canvas.Children.Remove(element)
            except Exception:
                pass
        return remove

    def _scale_anim(self, element, cx, cy, frames):
        """Scale `element` around its own (cx, cy)."""
        from System.Windows.Media import ScaleTransform
        st = ScaleTransform()
        st.CenterX, st.CenterY = cx, cy
        element.RenderTransform = st
        for prop in (ScaleTransform.ScaleXProperty, ScaleTransform.ScaleYProperty):
            st.BeginAnimation(prop, self._anim(frames))

    # ── HUD ───────────────────────────────────────────────────────────────────

    def _build_hud(self):
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border, StackPanel, Orientation
        from System.Windows.Input import Cursors

        bar = Border()
        bar.Background = self._brush('#1F2937', 235)
        bar.CornerRadius = CornerRadius(8)
        bar.Padding = Thickness(12, 8, 8, 8)
        col = StackPanel()
        bar.Child = col

        top = StackPanel()
        top.Orientation = Orientation.Horizontal
        col.Children.Add(top)
        self._agents_panel = StackPanel()
        top.Children.Add(self._agents_panel)

        self._toggles = StackPanel()
        self._toggles.Orientation = Orientation.Horizontal
        self._toggles.Margin = Thickness(16, 0, 0, 0)
        top.Children.Add(self._toggles)

        close = self._text('', 11, color='#CBD5E1', glyph=True)
        close.Margin = Thickness(8, 2, 4, 0)
        close.Cursor = Cursors.Hand
        close.ToolTip = ('Turn agent cursors off. Settings > Agent Cursors '
                         'turns them back on.')
        self._hook(close, 'MouseLeftButtonUp', lambda s, e: set_enabled(False))
        top.Children.Add(close)

        self._views_text = self._text('', 11, bold=False, color='#CBD5E1')
        self._views_text.Margin = Thickness(0, 4, 0, 0)
        col.Children.Add(self._views_text)
        self.hud.Content = bar

    def _toggle(self, name, on):
        from System.Windows import Thickness
        from System.Windows.Input import Cursors
        label, tip = OPTION_LABELS[name]
        glyph = '' if on else ''        # CheckMark / Checkbox
        chip = self._pill(self._glyph_label(glyph, label, '#FFFFFF' if on else '#94A3B8'),
                          '#374151' if on else '#1F2937')
        chip.Margin = Thickness(4, 0, 0, 0)
        chip.Cursor = Cursors.Hand
        chip.ToolTip = tip

        def flip(s, e):
            try:
                set_option(name, not options()[name])
                self._refresh_hud()
            except Exception:
                pass
        self._hook(chip, 'MouseLeftButtonUp', flip, transient=True)
        return chip

    def _refresh_hud(self):
        from System.Windows import Thickness, CornerRadius, HorizontalAlignment
        from System.Windows.Controls import Border
        self._agents_panel.Children.Clear()
        for agent in sorted(self._counts):
            tag, color = agent_style(agent)
            dot = Border()
            dot.Width = dot.Height = 18
            dot.CornerRadius = CornerRadius(9)
            dot.Background = self._brush(color)
            dot.Margin = Thickness(0, 0, 8, 0)
            letter = self._text(tag[:1].upper(), 9)
            letter.HorizontalAlignment = HorizontalAlignment.Center
            dot.Child = letter
            parts = [dot, self._text(tag, 11.5)]
            counts = self._counts[agent]
            for key, label, glyph, acolor in ACTIONS:
                if not counts.get(key):
                    continue
                chip = self._pill(self._glyph_label(glyph, '{}'.format(counts[key]), '#FFFFFF'),
                                  acolor, radius=2)
                chip.Margin = Thickness(6, 0, 0, 0)
                chip.ToolTip = '{} · {} {}'.format(tag, label, counts[key])
                parts.append(chip)
            row = self._row(*parts)
            row.Margin = Thickness(0, 1, 0, 1)
            row.ToolTip = '{}: {}'.format(tag, counts_text(counts))
            self._agents_panel.Children.Add(row)
        self._toggles.Children.Clear()
        opts = options()
        for name, _k, _d in OPTIONS:
            self._toggles.Children.Add(self._toggle(name, opts[name]))
        n = self._actions
        names = [v[1] for v in self._views[-3:]]
        more = len(self._views) - len(names)
        views = ' · '.join(names) + (' · +{}'.format(more) if more > 0 else '')
        self._views_text.Text = '{} action{}{}'.format(
            n, '' if n == 1 else 's', ('  ·  Views: ' + views) if views else '')

    # ── cursors ──────────────────────────────────────────────────────────────

    def _cursor(self, agent, action):
        key = '{}|{}'.format(agent, action)
        cur = self._cursors.get(key)
        if cur is not None:
            return cur
        from System.Windows import Thickness
        from System.Windows.Controls import Canvas, StackPanel
        from System.Windows.Media import TranslateTransform, Geometry
        from System.Windows.Shapes import Path
        tag, color = agent_style(agent)
        label, glyph, acolor = action_style(action)

        group = Canvas()
        arrow = Path()
        arrow.Data = Geometry.Parse('M0,0 L0,17 L4.6,12.8 L7.6,19.6 L10.4,18.4 L7.4,11.6 L13.4,11.6 Z')
        arrow.Fill = self._brush(color)
        arrow.Stroke = self._brush('#FFFFFF')
        arrow.StrokeThickness = 1.2
        group.Children.Add(arrow)

        tags = StackPanel()
        Canvas.SetLeft(tags, 12)
        Canvas.SetTop(tags, 18)
        name = self._pill(self._row(self._text(tag), self._action_chip(glyph, label, acolor)),
                          color)
        tags.Children.Add(name)
        tool = self._text('', 10, bold=False, color='#1F2937')
        tool_pill = self._pill(tool, '#FFFFFF', alpha=230)
        tool_pill.Margin = Thickness(0, 2, 0, 0)
        tool_pill.BorderBrush = self._brush(acolor)
        tool_pill.BorderThickness = Thickness(1)
        tags.Children.Add(tool_pill)
        group.Children.Add(tags)

        tx = TranslateTransform()
        group.RenderTransform = tx
        self.canvas.Children.Add(group)
        # Cursors start at different spots so a swarm never stacks up.
        slot = len(self._cursors)
        cur = {'group': group, 'tx': tx, 'tool': tool, 'tool_pill': tool_pill,
               'color': acolor, 'x': 40.0 + 36 * slot, 'y': 40.0 + 24 * slot}
        tx.X, tx.Y = cur['x'], cur['y']
        self._cursors[key] = cur
        return cur

    def _action_chip(self, glyph, label, color):
        from System.Windows import Thickness
        chip = self._pill(self._glyph_label(glyph, label, '#FFFFFF', 9.5), color, radius=2)
        chip.Padding = Thickness(4, 0, 4, 0)
        chip.Margin = Thickness(6, 0, 0, 0)
        return chip

    def _glide(self, cur, x, y):
        from System import TimeSpan
        from System.Windows import Duration
        from System.Windows.Media import TranslateTransform
        from System.Windows.Media.Animation import DoubleAnimation, CubicEase, EasingMode
        ease = CubicEase()
        ease.EasingMode = EasingMode.EaseInOut
        for prop, start, end in ((TranslateTransform.XProperty, cur['x'], x),
                                 (TranslateTransform.YProperty, cur['y'], y)):
            anim = DoubleAnimation()
            anim.From = start
            anim.To = end
            anim.Duration = Duration(TimeSpan.FromMilliseconds(MOVE_MS))
            anim.EasingFunction = ease
            cur['tx'].BeginAnimation(prop, anim)
        cur['x'], cur['y'] = x, y

    def _shake(self, cur):
        from System.Windows.Media import TranslateTransform
        x = cur['x']
        cur['tx'].BeginAnimation(TranslateTransform.XProperty, self._anim(
            ((x, MOVE_MS), (x - 6, MOVE_MS + 60), (x + 6, MOVE_MS + 120),
             (x - 4, MOVE_MS + 180), (x, MOVE_MS + 240))))

    # ── effects ──────────────────────────────────────────────────────────────

    def _trail(self, x0, y0, x1, y1, color):
        """Dashed path from where the cursor was to where it goes."""
        from System.Windows.Media import DoubleCollection
        from System.Windows.Shapes import Line
        if abs(x1 - x0) + abs(y1 - y0) < 24:
            return
        line = Line()
        line.X1, line.Y1, line.X2, line.Y2 = x0, y0, x1, y1
        line.Stroke = self._brush(color, 170)
        line.StrokeThickness = 1.5
        dash = DoubleCollection()
        dash.Add(3)
        dash.Add(3)
        line.StrokeDashArray = dash
        line.Opacity = 0
        self.canvas.Children.Insert(0, line)
        line.BeginAnimation(_opacity_property(), self._anim(
            ((0.0, 0), (0.9, 120), (0.9, MOVE_MS), (0.0, MOVE_MS + 700)),
            self._remover(line)))

    def _ripple(self, x, y, color, delay_ms, size=36):
        """Click ring where the cursor lands."""
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Ellipse
        ring = Ellipse()
        ring.Width = ring.Height = size
        ring.Stroke = self._brush(color)
        ring.StrokeThickness = 2
        ring.Fill = self._brush(color, 40)
        ring.Opacity = 0
        Canvas.SetLeft(ring, x - size / 2.0)
        Canvas.SetTop(ring, y - size / 2.0)
        self.canvas.Children.Add(ring)
        self._scale_anim(ring, size / 2.0, size / 2.0,
                         ((0.1, delay_ms), (0.1, delay_ms), (1.0, delay_ms + 520)))
        ring.BeginAnimation(_opacity_property(), self._anim(
            ((0.0, delay_ms), (0.9, delay_ms + 40), (0.0, delay_ms + 560)),
            self._remover(ring)))

    def _frame(self, x0, y0, x1, y1, color, dashed=False, thickness=1.0, fill_alpha=18):
        from System.Windows.Controls import Canvas
        from System.Windows.Media import DoubleCollection
        from System.Windows.Shapes import Rectangle
        frame = Rectangle()
        frame.Width = max(x1 - x0, 2)
        frame.Height = max(y1 - y0, 2)
        frame.Stroke = self._brush(color)
        frame.StrokeThickness = thickness
        frame.Fill = self._brush(color, fill_alpha)
        if dashed:
            dash = DoubleCollection()
            dash.Add(4)
            dash.Add(3)
            frame.StrokeDashArray = dash
        Canvas.SetLeft(frame, x0)
        Canvas.SetTop(frame, y0)
        return frame

    def _draw_box(self, x0, y0, x1, y1, size, action, delay_ms):
        """One target box with the action's effect."""
        from System.Windows import Thickness
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Rectangle, Shape
        from System.Windows.Media import TranslateTransform
        label, glyph, color = action_style(action)
        group = Canvas()
        group.Opacity = 0
        frame = self._frame(x0, y0, x1, y1, color, dashed=(action == 'code'))
        group.Children.Add(frame)
        for hx, hy in ((x0, y0), (x1, y0), (x0, y1), (x1, y1)):
            handle = Rectangle()
            handle.Width = handle.Height = 5
            handle.Fill = self._brush('#FFFFFF')
            handle.Stroke = self._brush(color)
            handle.StrokeThickness = 1
            Canvas.SetLeft(handle, hx - 2.5)
            Canvas.SetTop(handle, hy - 2.5)
            group.Children.Add(handle)
        badge = self._pill(self._glyph_label(glyph, size, '#FFFFFF'), color, radius=2)
        badge.Padding = Thickness(4, 1, 4, 1)
        Canvas.SetLeft(badge, (x0 + x1) / 2.0 - 34)
        Canvas.SetTop(badge, y1 + 4)
        group.Children.Add(badge)
        self.canvas.Children.Insert(0, group)    # under the cursors
        hold = int(BOX_HOLD_S * 1000)
        start = delay_ms + 180

        if action == 'create':
            # Grows out of its centre.
            self._scale_anim(group, (x0 + x1) / 2.0, (y0 + y1) / 2.0,
                             ((0.4, delay_ms), (0.4, delay_ms),
                              (1.08, start + 160), (1.0, start + 260)))
        elif action == 'modify':
            frame.BeginAnimation(Shape.StrokeThicknessProperty, self._anim(
                ((1.0, start), (3.0, start + 200), (1.0, start + 400),
                 (3.0, start + 600), (1.0, start + 800))))
        elif action == 'read':
            line = Rectangle()
            line.Width = max(x1 - x0, 2)
            line.Height = 2
            line.Fill = self._brush(color, 200)
            Canvas.SetLeft(line, x0)
            Canvas.SetTop(line, y0)
            tt = TranslateTransform()
            line.RenderTransform = tt
            group.Children.Add(line)
            h = max(y1 - y0 - 2, 0)
            tt.BeginAnimation(TranslateTransform.YProperty, self._anim(
                ((0.0, start), (h, start + 700), (0.0, start + 1400))))
        elif action == 'code':
            frame.BeginAnimation(Shape.StrokeDashOffsetProperty, self._anim(
                ((0.0, start), (-42.0, start + hold))))
        elif action == 'export':
            up = self._pill(self._text('', 11, glyph=True), color, radius=9)
            Canvas.SetLeft(up, x1 - 10)
            Canvas.SetTop(up, y0 - 10)
            tt = TranslateTransform()
            up.RenderTransform = tt
            group.Children.Add(up)
            tt.BeginAnimation(TranslateTransform.YProperty, self._anim(
                ((0.0, start), (-28.0, start + 900))))
        self._pulse(group, delay_ms, hold, self._remover(group))

    def _draw_gone(self, x0, y0, x1, y1, delay_ms):
        """A deleted element: red frame with a cross that shrinks away."""
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Line
        _l, _g, color = action_style('delete')
        group = Canvas()
        group.Opacity = 0
        group.Children.Add(self._frame(x0, y0, x1, y1, color, dashed=True, fill_alpha=40))
        for ax, ay, bx, by in ((x0, y0, x1, y1), (x0, y1, x1, y0)):
            line = Line()
            line.X1, line.Y1, line.X2, line.Y2 = ax, ay, bx, by
            line.Stroke = self._brush(color, 200)
            line.StrokeThickness = 1.5
            group.Children.Add(line)
        self.canvas.Children.Insert(0, group)
        hold = 900
        self._scale_anim(group, (x0 + x1) / 2.0, (y0 + y1) / 2.0,
                         ((1.0, delay_ms + 180 + hold), (0.15, delay_ms + 780 + hold)))
        self._pulse(group, delay_ms, hold, self._remover(group))

    def _draw_ghost(self, old, new, color, delay_ms):
        """Where a moved element was, with an arrow to where it is now."""
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Line
        group = Canvas()
        group.Opacity = 0
        group.Children.Add(self._frame(old[0], old[1], old[2], old[3], color,
                                       dashed=True, fill_alpha=0))
        line = Line()
        line.X1, line.Y1 = (old[0] + old[2]) / 2.0, (old[1] + old[3]) / 2.0
        line.X2, line.Y2 = (new[0] + new[2]) / 2.0, (new[1] + new[3]) / 2.0
        line.Stroke = self._brush(color, 190)
        line.StrokeThickness = 1.5
        group.Children.Add(line)
        head = self._ripple_dot(line.X2, line.Y2, color)
        group.Children.Add(head)
        self.canvas.Children.Insert(0, group)
        self._pulse(group, delay_ms, int(BOX_HOLD_S * 1000), self._remover(group))

    def _ripple_dot(self, x, y, color):
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Ellipse
        dot = Ellipse()
        dot.Width = dot.Height = 7
        dot.Fill = self._brush(color)
        Canvas.SetLeft(dot, x - 3.5)
        Canvas.SetTop(dot, y - 3.5)
        return dot

    def _flash_view(self, geo, text, color):
        """Frame the whole view and name it: the agent opened / switched to it."""
        from System.Windows import Thickness
        from System.Windows.Controls import Canvas
        x0, y0, x1, y1 = self._local(geo['rect'])
        group = Canvas()
        group.Opacity = 0
        group.Children.Add(self._frame(x0 + 2, y0 + 2, x1 - 2, y1 - 2, color,
                                       thickness=3, fill_alpha=0))
        chip = self._pill(self._glyph_label('', text, '#FFFFFF', 11), color)
        chip.Padding = Thickness(8, 4, 8, 4)
        Canvas.SetLeft(chip, x0 + 12)
        Canvas.SetTop(chip, y0 + 12)
        group.Children.Add(chip)
        self.canvas.Children.Insert(0, group)
        self._pulse(group, 0, 1600, self._remover(group))

    # ── public ────────────────────────────────────────────────────────────────

    def show(self, agent, action, tool_name, failed, snap, followed=()):
        left, top, right, bottom = snap['rect']
        if not self.win.IsVisible:
            self.win.Width = self.win.Height = 1
            self.win.Show()
        scale = self._dpi()
        if self._origin != (left, top):
            # Canvas moved: older cursors keep their canvas spot, so shift them.
            dx = (self._origin[0] - left) / scale
            dy = (self._origin[1] - top) / scale
            for cur in self._cursors.values():
                cur['x'] += dx
                cur['y'] += dy
                cur['tx'].BeginAnimation(_translate_x(), None)
                cur['tx'].BeginAnimation(_translate_y(), None)
                cur['tx'].X, cur['tx'].Y = cur['x'], cur['y']
        self._origin = (left, top)
        self.win.Left, self.win.Top = left / scale, top / scale
        self.win.Width = (right - left) / scale
        self.win.Height = (bottom - top) / scale

        label, _glyph, acolor = action_style(action)
        cur = self._cursor(agent, action)
        cur['tool'].Text = '{} · failed'.format(tool_name) if failed else tool_name
        cur['tool_pill'].BorderBrush = self._brush('#DC2626' if failed else acolor)

        for vid, name, is_new in followed:
            if vid not in [v[0] for v in self._views]:
                self._views.append((vid, name))
        if followed:
            first = followed[0]
            self._flash_view(snap['active'], '{} · {}'.format(
                'Opened' if first[2] else 'Switched to', first[1]), action_style('view')[2])

        target, n = None, 0
        for geo in snap['views']:
            for item in geo['boxes']:
                x0, y0, x1, y1 = self._local(item['box'])
                self._draw_box(x0, y0, x1, y1, item['size'], action, MOVE_MS + 60 * n)
                n += 1
                if target is None:
                    target = (x1 - 4, y0 + 4)       # tags hang right of the box
            for item in geo['gone']:
                x0, y0, x1, y1 = self._local(item['box'])
                self._draw_gone(x0, y0, x1, y1, MOVE_MS + 60 * n)
                n += 1
                if target is None:
                    target = ((x0 + x1) / 2.0, (y0 + y1) / 2.0)
            for old, _new in geo['ghosts']:
                self._draw_ghost(self._local(old['box']), self._local(_new['box']),
                                 acolor, MOVE_MS)
        if target is None:
            # No element to point at: drift in the active view, one slot per
            # (agent, action) wide enough for its tags, so cursors never overlap.
            ax0, ay0, ax1, ay1 = self._local(snap['active']['rect'])
            slot = sorted(self._cursors).index('{}|{}'.format(agent, action))
            target = (ax0 + (ax1 - ax0) * 0.25 + 210 * (slot % 3) + (self._actions % 3) * 10,
                      ay0 + (ay1 - ay0) * 0.25 + 64 * (slot // 3) + (self._actions % 2) * 8)
        self._trail(cur['x'], cur['y'], target[0], target[1], acolor)
        self._glide(cur, target[0], target[1])
        self._ripple(target[0], target[1], '#DC2626' if failed else acolor, MOVE_MS)
        if failed:
            self._shake(cur)
        cur['group'].BeginAnimation(_opacity_property(), None)
        cur['group'].Opacity = 1.0

        counts = self._counts.setdefault(agent, {})
        counts[action] = counts.get(action, 0) + 1
        self._actions += 1
        self._last = time.time()
        self._refresh_hud()
        if not self.hud.IsVisible:
            self.hud.Show()
        self.hud.UpdateLayout()
        aleft, _atop, aright, abottom = snap['active']['rect']
        hud_w = self.hud.ActualWidth or 320
        self.hud.Left = aleft / scale + ((aright - aleft) / scale - hud_w) / 2.0
        self.hud.Top = abottom / scale - (self.hud.ActualHeight or 48) - 16
        self._timer.Start()

    def hide(self):
        self._timer.Stop()
        for cur in self._cursors.values():
            cur['group'].BeginAnimation(_opacity_property(), None)
        self.canvas.Children.Clear()
        self._cursors = {}
        self._counts = {}
        self._views = []
        self._actions = 0
        self._transient = []
        self.win.Hide()
        self.hud.Hide()

    def _on_tick(self, sender, e):
        try:
            if self.hud.IsMouseOver:
                self._last = time.time()        # reading the HUD / toggling
            elif time.time() - self._last >= IDLE_HIDE_S:
                self.hide()
        except Exception:
            pass


def _translate_x():
    from System.Windows.Media import TranslateTransform
    return TranslateTransform.XProperty


def _translate_y():
    from System.Windows.Media import TranslateTransform
    return TranslateTransform.YProperty
