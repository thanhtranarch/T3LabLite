# -*- coding: utf-8 -*-
"""
Session context — memory for ONE chat, and memory of the chats before it.

Claude Projects gives every conversation three layers of context: the
project's instructions and knowledge, a memory of earlier chats in the
project, and the conversation itself. The assistant had the first and the
last; between them sat only hand-saved facts (assistant_memory), and a new
chat knew nothing about what the previous ten had done.

This module adds the two missing layers. It owns no storage of its own:
everything lives in the chat JSON files the pane already writes.

  Session notes   facts that hold for THIS conversation only — "we only touch
                  level 3", "use titleblock A1 for this batch". Saved with the
                  chat (active history file + its archive), restored when the
                  chat is resumed, never shown to any other chat. Prompt
                  section: "## Notes for this conversation".
  Past chats      a digest of every archived conversation (what was asked,
                  what came out). A new chat starts with the newest digests of
                  its project (or, without a project, of the same Revit model):
                  "## Earlier conversations". Anything older is one
                  search_past_chats call away — the model can search every
                  archived chat in scope and open any of them.

Durable preferences and project conventions stay in assistant_memory
(global / project facts); session notes are the third, narrowest scope.

Pure Python — no Revit, no WPF, no network. CPython 3 / IronPython 2.7.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Session Context"

import re
from datetime import datetime

SESSION_SCOPE = "session"

MAX_NOTE_CHARS = 1000
MIN_NOTE_CHARS = 4

# One archived chat's digest. Long enough for the ask and its key results.
DIGEST_CHARS = 700

# search_past_chats: transcript budget when the model opens one chat, and
# the longest single message inside it.
READ_CHARS = 24000
_READ_MSG_CHARS = 2000
_SNIPPET_CHARS = 280
MAX_RESULTS = 10

_HEAD_NOTES_EN = u"## Notes for this conversation"
_HEAD_NOTES_VI = u"## Ghi chú của cuộc hội thoại này"
_HEAD_PAST_EN = u"## Earlier conversations"
_HEAD_PAST_VI = u"## Các cuộc hội thoại trước"


def _one_line(text):
    return u' '.join(u'{}'.format(text or u'').split())


def _norm(text):
    return _one_line(text).lower()


def _fold(text):
    text = (text or u'').lower()
    try:
        from Intelligence.knowledge.vi_text import fold_diacritics
        return fold_diacritics(text)
    except Exception:
        return text


def _content_text(content):
    if isinstance(content, (list, tuple)):
        return u' '.join(u'{}'.format(b.get('text') or u'')
                         for b in content
                         if isinstance(b, dict) and b.get('type') == 'text')
    return u'{}'.format(content or u'')


# ─── Session notes ────────────────────────────────────────────────────────────
# A list of {'text', 'created'} dicts — the same shape as assistant_memory's
# facts. Every function returns a NEW list; the caller stores it.

def clean_note(text):
    """One line, clipped; '' when too short to mean anything."""
    text = _one_line(text)
    if len(text) > MAX_NOTE_CHARS:
        text = text[:MAX_NOTE_CHARS].rstrip()
    return text if len(text) >= MIN_NOTE_CHARS else u''


def normalize_notes(raw):
    """Notes as loaded from JSON (dicts or plain strings) → clean dict list."""
    out = []
    seen = set()
    for item in raw or []:
        if isinstance(item, dict):
            text, created = clean_note(item.get('text')), item.get('created')
        else:
            text, created = clean_note(item), None
        if not text or text.lower() in seen:
            continue
        seen.add(text.lower())
        out.append({'text': text, 'created': u'{}'.format(created or u'')})
    return out


def add_note(notes, text):
    """Add one note. Returns (notes, ok, user_message)."""
    notes = normalize_notes(notes)
    text = clean_note(text)
    if not text:
        return notes, False, u'Nothing concrete to note in that.'
    if any(n['text'].lower() == text.lower() for n in notes):
        return notes, True, u'Already noted for this conversation.'
    notes.append({'text': text,
                  'created': datetime.now().strftime('%Y-%m-%d %H:%M')})
    return notes, True, u'Noted for this conversation: {}'.format(text)


def _note_index(notes, ref):
    """Index of the note `ref` names: "S2" / "2" (1-based) or its content."""
    ref = _one_line(ref)
    m = re.match(r'^[sS]?(\d+)$', ref)
    if m:
        idx = int(m.group(1)) - 1
        return idx if 0 <= idx < len(notes) else None
    q = ref.lower()
    if not q:
        return None
    for i, n in enumerate(notes):
        if n['text'].lower() == q:
            return i
    for i, n in enumerate(notes):
        t = n['text'].lower()
        if q in t or t in q:
            return i
    return None


def forget_note(notes, ref):
    """Remove the note `ref` names. Returns (notes, removed_text_or_None)."""
    notes = normalize_notes(notes)
    idx = _note_index(notes, ref)
    if idx is None:
        return notes, None
    removed = notes.pop(idx)
    return notes, removed['text']


def update_note(notes, old, new):
    """Replace a note's text; adds `new` when `old` matches nothing.

    Returns (notes, ok, user_message)."""
    notes = normalize_notes(notes)
    text = clean_note(new)
    if not text:
        return notes, False, u'Nothing concrete to note in that.'
    idx = _note_index(notes, old) if old else None
    if idx is None:
        return add_note(notes, text)
    notes[idx] = {'text': text, 'created': notes[idx].get('created') or u''}
    return notes, True, u'Updated for this conversation: {}'.format(text)


def notes_block(notes, viet=False, tool_hint=True):
    """Prompt section for the session notes ('' when there are none)."""
    notes = normalize_notes(notes)
    if not notes:
        return u''
    lines = [_HEAD_NOTES_VI if viet else _HEAD_NOTES_EN,
             (u'Áp dụng cho MỌI lượt của cuộc hội thoại này (không áp dụng '
              u'cho cuộc hội thoại khác).' if viet else
              u'Apply these to EVERY turn of this conversation (they do not '
              u'apply to other conversations).')]
    for i, n in enumerate(notes):
        lines.append(u'S{}. {}'.format(i + 1, n['text']))
    if tool_hint:
        lines.append(u'To change one, call `remember_fact` with scope '
                     u'"session" and action "update" or "forget".')
    return u'\n'.join(lines)


def notes_report_lines(notes, viet=False):
    """Markdown lines for /memory: the session notes, numbered S1, S2…"""
    notes = normalize_notes(notes)
    if not notes:
        return []
    out = [u'',
           (u'**Ghi chú của cuộc hội thoại này** ({} mục):' if viet else
            u'**Notes for this conversation** ({} notes):').format(len(notes)),
           u'']
    for i, n in enumerate(notes):
        out.append(u'- **S{}** {}'.format(i + 1, n['text'].replace(u'|', u'\\|')))
    return out


# ─── Digests of archived chats ────────────────────────────────────────────────

def make_digest(messages, summary=u'', title=u'', max_chars=DIGEST_CHARS):
    """A few lines that say what one conversation was about and produced.

    Extractive, like the rolling summary it builds on: every user request,
    and assistant turns that report something concrete. The FIRST line (what
    the chat set out to do) is always kept; the rest is the newest that fits.
    """
    try:
        from Intelligence.conversation import summarize_turns
        turn_lines = summarize_turns(messages or [])
    except Exception:
        turn_lines = []
    lines = [l.strip() for l in (summary or u'').split(u'\n') if l.strip()]
    lines.extend(turn_lines)
    if not lines and title:
        lines = [u'User asked: {}'.format(_one_line(title))]
    if not lines:
        return u''
    first, rest = lines[0], lines[1:]
    kept = []
    used = len(first)
    for line in reversed(rest):
        if used + len(line) + 1 > max_chars:
            break
        kept.insert(0, line)
        used += len(line) + 1
    if len(kept) < len(rest):
        first = first + u'\n…'
    return u'\n'.join([first] + kept)[:max_chars + 2]


def past_chats_block(sessions, max_chars, viet=False, project=True,
                     tool_hint=True):
    """Prompt section with the digests of earlier chats ('' when none).

    `sessions` is chat_sessions.list_sessions() output (newest first), each
    carrying 'digest'. Digests are added until `max_chars` is reached.
    """
    sessions = [s for s in (sessions or []) if (s.get('digest') or u'').strip()]
    if not sessions:
        return u''
    head = _HEAD_PAST_VI if viet else _HEAD_PAST_EN
    if viet:
        where = u'trong dự án này' if project else u'về model Revit này'
        note = (u'Tóm tắt các cuộc hội thoại trước của người dùng {} (mới '
                u'nhất trước). Chỉ là bối cảnh — KHÔNG phải việc cần làm '
                u'lại.'.format(where))
    else:
        where = u'in this project' if project else u'about this Revit model'
        note = (u"Digests of the user's previous conversations {} (newest "
                u"first). Background only — NOT tasks to redo.".format(where))
    if tool_hint:
        note += (u' Gọi `search_past_chats` để tìm chi tiết hoặc mở lại '
                 u'một cuộc hội thoại.' if viet else
                 u' Call `search_past_chats` to look up details or open one.')
    parts = [head, note]
    used = len(head) + len(note) + 2
    shown = 0
    for s in sessions:
        entry = u'### {} — {} (id {})\n{}'.format(
            s.get('created') or u'?', _one_line(s.get('title')) or u'Conversation',
            s.get('id') or u'?', s.get('digest').strip())
        if used + len(entry) + 2 > max_chars and shown:
            break
        if used + len(entry) + 2 > max_chars:
            entry = entry[:max(0, max_chars - used - 3)] + u'…'
        parts.append(entry)
        used += len(entry) + 2
        shown += 1
    rest = len(sessions) - shown
    if rest > 0:
        parts.append((u'{} cuộc hội thoại cũ hơn không hiển thị.' if viet else
                      u'{} older conversation(s) not shown.').format(rest))
    return u'\n\n'.join(parts)


# ─── search_past_chats ────────────────────────────────────────────────────────

def _snippet(messages, query, width=_SNIPPET_CHARS):
    """Text around the first message that contains every query term."""
    terms = _fold(query).split()
    for m in messages or []:
        if not isinstance(m, dict):
            continue
        text = _one_line(_content_text(m.get('content')))
        folded = _fold(text)
        if terms and all(t in folded for t in terms):
            pos = folded.find(terms[0])
            start = max(0, pos - width // 3)
            clip = text[start:start + width]
            who = u'User' if m.get('role') == 'user' else u'Assistant'
            return u'{}: {}{}{}'.format(who, u'…' if start else u'', clip,
                                        u'…' if start + width < len(text) else u'')
    return u''


def _transcript(messages, max_chars=READ_CHARS):
    """'User: …' / 'Assistant: …' lines; head + newest when it is too long.

    Returns (text, omitted_message_count).
    """
    lines = []
    for m in messages or []:
        if not isinstance(m, dict) or m.get('role') not in ('user', 'assistant'):
            continue
        text = _content_text(m.get('content')).strip()
        if not text:
            continue
        if len(text) > _READ_MSG_CHARS:
            text = text[:_READ_MSG_CHARS].rstrip() + u' …[cut]'
        who = u'User' if m.get('role') == 'user' else u'Assistant'
        lines.append(u'{}: {}'.format(who, text))
    full = u'\n\n'.join(lines)
    if len(full) <= max_chars:
        return full, 0
    head = lines[:2]
    used = sum(len(l) + 2 for l in head)
    tail = []
    for line in reversed(lines[2:]):
        if used + len(line) + 2 > max_chars:
            break
        tail.insert(0, line)
        used += len(line) + 2
    omitted = len(lines) - len(head) - len(tail)
    gap = [u'[… {} message(s) omitted …]'.format(omitted)] if omitted else []
    return u'\n\n'.join(head + gap + tail), omitted


def _public(meta):
    return {'id': meta.get('id'), 'title': meta.get('title'),
            'created': meta.get('created'),
            'messages': meta.get('message_count')}


def search_past_chats(pid=None, query=None, session_id=None, limit=5,
                      doc_key=None, exclude_id=None, read_chars=READ_CHARS):
    """The search_past_chats tool. Scope = the active project's archived
    chats, or with no project the "No project" archive.

        session_id  → that chat's transcript (head + newest within read_chars)
        query       → chats whose title/summary/messages contain every term,
                      newest first, each with a snippet around the match
        neither     → the most recent chats with their digests

    `doc_key` narrows to one Revit model; `exclude_id` drops the chat that is
    open right now (its turns are already in the context).
    """
    from config import chat_sessions
    scope = u'project' if pid else u'no project'
    try:
        limit = max(1, min(MAX_RESULTS, int(limit or 5)))
    except (TypeError, ValueError):
        limit = 5

    def _visible(m):
        return not (exclude_id and m.get('id') == exclude_id)

    if session_id:
        sid = _one_line(session_id)
        for meta in chat_sessions.list_sessions(pid=pid, doc_key=doc_key):
            if meta.get('id') == sid or meta.get('path', u'').endswith(
                    u'_{}.json'.format(sid)):
                data = chat_sessions.read_session(meta['path']) or {}
                text, omitted = _transcript(data.get('messages') or [],
                                            max_chars=read_chars)
                out = {'scope': scope, 'session': _public(meta),
                       'transcript': text}
                if omitted:
                    out['omitted_messages'] = omitted
                notes = normalize_notes(data.get('session_notes'))
                if notes:
                    out['session_notes'] = [n['text'] for n in notes]
                return out
        return {'error': u'No archived conversation with id {} in this scope '
                         u'({}).'.format(sid, scope)}

    if query and _one_line(query):
        hits = [m for m in chat_sessions.list_sessions(
            pid=pid, doc_key=doc_key, query=query) if _visible(m)]
        matches = []
        for meta in hits[:limit]:
            data = chat_sessions.read_session(meta['path']) or {}
            item = _public(meta)
            item['snippet'] = (_snippet(data.get('messages') or [], query)
                               or (meta.get('digest') or u'')[:_SNIPPET_CHARS])
            matches.append(item)
        out = {'scope': scope, 'query': _one_line(query), 'total': len(hits),
               'matches': matches}
        if not hits:
            out['note'] = (u'No archived conversation mentions that. Only '
                           u'finished chats are searchable; the current one '
                           u'is already in your context.')
        elif len(hits) > limit:
            out['note'] = (u'Showing {} of {}; narrow the query or open one '
                           u'with session_id.'.format(limit, len(hits)))
        return out

    recent = []
    for meta in chat_sessions.list_sessions(pid=pid, doc_key=doc_key):
        if not _visible(meta):
            continue
        item = _public(meta)
        item['digest'] = meta.get('digest') or u''
        recent.append(item)
        if len(recent) >= limit:
            break
    return {'scope': scope, 'recent': recent}
