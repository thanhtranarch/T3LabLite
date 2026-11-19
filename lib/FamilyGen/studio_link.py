# -*- coding: utf-8 -*-
"""FamiGen Studio <-> MCP link: small JSON files, no sockets, no shared memory.

FamiGen Studio runs in pyRevit's CPython engine; the MCP server runs in the
IronPython host (CLAUDE.md rule 8). They cannot call each other, so they meet
in ``%APPDATA%\\T3LabAI\\famigen_studio`` (docs/famigen-studio-spec.md §7):

| File | Written by | Read by | Meaning |
|---|---|---|---|
| ``presence.json`` | Studio, every 2 s | server | Studio is open (fresh < 10 s) |
| ``mcp_activity.json`` | server, each tool call (throttled) | Studio | an AI app is connected (fresh < 10 min) |
| ``state.json`` | Studio, on change | server | current family, category, selected part |
| ``inbox.json`` | Studio (MCP mode) | server | messages typed in Studio for the AI app |
| ``listener.json`` | server, famigen_wait_for_message | Studio | an AI app is listening (fresh < 2 min) |

Proposals themselves travel through ``FamilyGen.proposals.ProposalStore``.
Writes are atomic (temp file + replace). Pure Python, IronPython 3 safe.
"""
import io
import json
import os
import time
import uuid

PRESENCE_MAX_AGE = 10.0          # s: Studio heartbeat is every 2 s
ACTIVITY_MAX_AGE = 600.0         # s: an AI app counts as connected for 10 min
MAX_INBOX = 50


def folder():
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI', 'famigen_studio')


def _path(name, root=None):
    return os.path.join(root or folder(), name)


def _write(name, data, root=None):
    path = _path(name, root)
    directory = os.path.dirname(path)
    if not os.path.isdir(directory):
        os.makedirs(directory)
    tmp = '{}.{}.tmp'.format(path, uuid.uuid4().hex[:6])
    with io.open(tmp, 'w', encoding='utf-8') as handle:
        handle.write(json.dumps(data, ensure_ascii=False))
    try:
        os.replace(tmp, path)
    except AttributeError:            # very old runtimes without os.replace
        if os.path.exists(path):
            os.remove(path)
        os.rename(tmp, path)


def _read(name, root=None):
    try:
        with io.open(_path(name, root), 'r', encoding='utf-8') as handle:
            return json.loads(handle.read())
    except (IOError, OSError, ValueError):
        return None


def _remove(name, root=None):
    try:
        os.remove(_path(name, root))
    except OSError:
        pass


# ── presence (Studio -> server) ──────────────────────────────────────────────
def beat(info=None, root=None, clock=time.time):
    data = dict(info or {})
    data.update({'pid': os.getpid(), 'updated': clock()})
    _write('presence.json', data, root)


def leave(root=None):
    _remove('presence.json', root)


def studio_open(root=None, clock=time.time, max_age=PRESENCE_MAX_AGE):
    data = _read('presence.json', root)
    return bool(isinstance(data, dict) and clock() - float(data.get('updated') or 0) <= max_age)


# ── AI app activity (server -> Studio) ───────────────────────────────────────
_LAST_NOTE = {'time': 0.0, 'client': None}
NOTE_EVERY = 5.0                 # s: other tools are noted at most this often


def note_activity(client, tool, root=None, clock=time.time):
    """An AI app called a T3Lab tool. famigen_* calls are always written; any other
    tool at most every NOTE_EVERY s - enough for the Studio's Auto to put MCP
    first while an app is connected (owner request, 2026-11-19)."""
    now = clock()
    name = u'{}'.format(client or 'AI app')
    key = (_path('mcp_activity.json', root), name)
    if (not u'{}'.format(tool or '').startswith('famigen_')
            and _LAST_NOTE['client'] == key and now - _LAST_NOTE['time'] < NOTE_EVERY):
        return
    _LAST_NOTE.update(time=now, client=key)
    _write('mcp_activity.json', {'client': name, 'tool': tool, 'time': now}, root)


def active_client(root=None, clock=time.time, max_age=ACTIVITY_MAX_AGE):
    """Name of the AI app that used a famigen_* tool recently, else None."""
    data = _read('mcp_activity.json', root)
    if not isinstance(data, dict):
        return None
    if clock() - float(data.get('time') or 0) > max_age:
        return None
    return friendly_client(data.get('client'))


def friendly_client(raw):
    """'claude-ai' -> 'Claude Desktop' etc.; unknown names stay as they are."""
    text = u'{}'.format(raw or '').strip()
    low = text.lower()
    for key, name in (('claude-code', 'Claude Code'), ('claude', 'Claude Desktop'),
                      ('codex', 'ChatGPT (Codex)'), ('openai', 'ChatGPT (Codex)'),
                      ('antigravity', 'Antigravity'), ('cursor', 'Cursor'),
                      ('mcp-client', 'AI app')):
        if key in low:
            return name
    return text or 'AI app'


# ── Studio state (Studio -> server) ──────────────────────────────────────────
def write_state(state, root=None, clock=time.time):
    data = dict(state or {})
    data['updated'] = clock()
    _write('state.json', data, root)


def read_state(root=None):
    data = _read('state.json', root)
    return data if isinstance(data, dict) else {}


# ── inbox (Studio -> AI app) ─────────────────────────────────────────────────
def post_message(text, files=(), root=None, clock=time.time):
    items = _read('inbox.json', root)
    items = items if isinstance(items, list) else []
    item = {'id': uuid.uuid4().hex[:8], 'time': clock(), 'text': u'{}'.format(text),
            'files': list(files or []), 'read': False}
    items = (items + [item])[-MAX_INBOX:]
    _write('inbox.json', items, root)
    return item


def take_unread(root=None, mark_read=True):
    items = _read('inbox.json', root)
    items = items if isinstance(items, list) else []
    unread = [dict(i) for i in items if isinstance(i, dict) and not i.get('read')]
    if unread and mark_read:
        for item in items:
            if isinstance(item, dict):
                item['read'] = True
        _write('inbox.json', items, root)
    for item in unread:
        item.pop('read', None)
    return unread


# ── listening AI app (server -> Studio) ─────────────────────────────────────
LISTEN_MAX_AGE = 120.0           # s: a listener calls again within one wait (<= 55 s)
WAIT_MAX = 55.0                  # s: under the 60 s tool timeout of most MCP clients
# Pasted once in the AI app: an exact tool name, so the app does not read
# "continue in FamiGen" as "search past chats" (2026-11-19).
LISTEN_COMMAND = (u'FamiGen Studio in Revit sends you requests through the T3Lab MCP server. '
                  u'Do not search past chats. Call the T3Lab tool famigen_wait_for_message '
                  u'now (through t3lab_call if it is not listed). For every message it returns, '
                  u'work on the family it includes with famigen_patch_family or '
                  u'famigen_propose_family, then call famigen_wait_for_message again. Keep '
                  u'listening until I say stop.')


def note_listening(client, root=None, clock=time.time):
    _write('listener.json', {'client': friendly_client(client), 'time': clock()}, root)


def listener(root=None, clock=time.time, max_age=LISTEN_MAX_AGE):
    """Name of the AI app waiting on famigen_wait_for_message, else None."""
    data = _read('listener.json', root)
    if not isinstance(data, dict) or clock() - float(data.get('time') or 0) > max_age:
        return None
    return data.get('client') or u'AI app'


def wait_unread(timeout=45.0, root=None, clock=time.time, sleep=time.sleep, poll=0.5):
    """Block until the Studio has unread messages (returned, marked read) or ``timeout`` s.

    Files only, no Revit: the server runs it on its HTTP thread, never on
    Revit's (core/server.py, famigen_wait_for_message)."""
    end = clock() + max(0.0, min(float(timeout), WAIT_MAX))
    while True:
        items = take_unread(root)
        if items or clock() >= end:
            return items
        sleep(poll)


def withdraw(message_id, root=None):
    """Take an unread message back out of the inbox (sent to the built-in Claude instead)."""
    items = _read('inbox.json', root)
    items = [i for i in items or [] if isinstance(i, dict) and i.get('id') != message_id]
    _write('inbox.json', items, root)


def read_ids(ids, root=None):
    """The ids among ``ids`` the AI app has read (gone from the inbox counts as read)."""
    items = _read('inbox.json', root)
    known = dict((i.get('id'), bool(i.get('read'))) for i in items or [] if isinstance(i, dict))
    return set(i for i in ids if known.get(i, True))


def last_activity(root=None):
    """{'client', 'tool', 'time'} of the AI app's last famigen_* call, or None."""
    data = _read('mcp_activity.json', root)
    if not isinstance(data, dict):
        return None
    return {'client': friendly_client(data.get('client')), 'tool': data.get('tool') or u'',
            'time': float(data.get('time') or 0)}


def unread_count(root=None):
    items = _read('inbox.json', root)
    return sum(1 for i in items or [] if isinstance(i, dict) and not i.get('read'))


# ── part-level edits for famigen_patch_family (token saver) ──────────────────
def patch_family(schema, set_fields=None, parts=None):
    """Apply a small edit to a family without resending all of it.

    ``set_fields``: top-level fields to replace (family_name, materials,
    parameters, types, reference_planes, dimensions ...); ``None`` deletes one.
    ``parts``: [{"id": ..., "remove": true} | {"id": ..., <fields to merge>}]
    keyed by geometry ``id``; an unknown id with a ``type`` is appended.
    Returns (new_schema, notes). Never mutates the input.
    """
    if not isinstance(schema, dict):
        raise ValueError('There is no family to patch.')
    out = json.loads(json.dumps(schema))
    notes = []
    for key, value in (set_fields or {}).items():
        if key == 'geometry':
            raise ValueError('Edit geometry through parts, or propose the whole family.')
        if value is None:
            out.pop(key, None)
            notes.append(u'removed {}'.format(key))
        else:
            out[key] = value
            notes.append(u'set {}'.format(key))
    geometry = out.get('geometry') if isinstance(out.get('geometry'), list) else []
    for edit in parts or []:
        if not isinstance(edit, dict) or not isinstance(edit.get('id'), str):
            raise ValueError('Each parts[] entry needs a string id.')
        pid = edit['id']
        index = next((i for i, g in enumerate(geometry)
                      if isinstance(g, dict) and g.get('id') == pid), None)
        if edit.get('remove'):
            if index is None:
                raise ValueError(u'No part with id {!r} to remove.'.format(pid))
            geometry.pop(index)
            notes.append(u'removed part {}'.format(pid))
            continue
        fields = dict((k, v) for k, v in edit.items() if k != 'remove')
        if index is None:
            if 'type' not in fields:
                raise ValueError(u'No part with id {!r}; give its type to add it.'.format(pid))
            geometry.append(fields)
            notes.append(u'added part {}'.format(pid))
        else:
            merged = dict(geometry[index])
            for key, value in fields.items():
                if value is None:
                    merged.pop(key, None)
                else:
                    merged[key] = value
            geometry[index] = merged
            notes.append(u'changed part {}'.format(pid))
    out['geometry'] = geometry
    return out, notes
