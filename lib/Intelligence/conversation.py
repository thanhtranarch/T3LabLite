# -*- coding: utf-8 -*-
"""
Conversation window + rolling summary.

The assistant keeps as many turns as the active model's context budget allows
(Intelligence.context_budget) and sends them to the model each turn. Older
turns were simply dropped, so a long session silently
forgot how it started — the user re-states a constraint from turn 2 and by
turn 20 the model has never seen it.

Instead of dropping, fold. Turns leaving the window are compressed into a
running summary that keeps the parts a later turn is most likely to need:
what the user asked for, and any concrete result (counts, element ids, names,
file paths) the assistant reported.

The compression is EXTRACTIVE and deterministic — no model call, so it adds
nothing to turn latency, works offline, and cannot hallucinate a fact into the
history. An LLM refinement could sit on top later; it must never be on the
critical path.

Pure Python — no Revit, no WPF, no network. CPython-3 testable.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Conversation"

import re

from Intelligence import context_budget

# The verbatim window is a TOKEN budget, not a message count.
#
# It was a fixed 24 messages (12 exchanges) for every model — a 1M-token
# Claude forgot the start of a session as fast as an 8k local model. The
# budget now comes from the active model's context window
# (context_budget.history_budget); callers pass it in. This default is what
# a caller that knows nothing about the model gets: the smallest window the
# assistant supports, so it can never overflow anything.
DEFAULT_HISTORY_TOKENS = context_budget.history_budget(
    context_budget.UNKNOWN_WINDOW)

# Default summary budget — the floor. Callers that know the model pass
# context_budget.summary_budget_chars(window), which grows with the window.
MAX_SUMMARY_CHARS = 1400

# One folded turn, before the whole summary is trimmed.
_MAX_LINE_CHARS = 180

_HEADING_VI = u"## Trước đó trong hội thoại này"
_HEADING_EN = u"## Earlier in this conversation"

# A fact worth carrying usually has a number, an id, or a quoted name in it.
_CONCRETE = re.compile(
    r'\d|"|“|‘|`|'
    r'\b(id|ids|element|elements|sheet|view|level|wall|floor|room|param)\b',
    re.IGNORECASE)

# Boilerplate the assistant emits that carries nothing forward.
_NOISE_PREFIXES = (
    u'──', u'●', u'đang ', u'working', u'let me', u'để tôi', u'ok,', u'okay,',
)


def _flatten(content):
    """Turn any stored content shape into plain text."""
    if isinstance(content, list):
        return u' '.join(
            b.get('text', '') for b in content
            if isinstance(b, dict) and b.get('type') == 'text').strip()
    return u'{}'.format(content or u'').strip()


def _one_line(text, limit=_MAX_LINE_CHARS):
    line = u' '.join((text or u'').split())
    if len(line) > limit:
        line = line[:limit].rstrip() + u'…'
    return line


def _is_noise(line):
    low = line.lower()
    return (not line) or any(low.startswith(p) for p in _NOISE_PREFIXES)


def summarize_turns(turns, viet=False):
    """Compress dropped turns into summary lines. Returns a list of strings.

    Every user request is kept — it is the record of what was asked. An
    assistant turn is kept only when it reports something concrete; narration
    ("Đang kiểm tra…", "Let me look at that") carries nothing forward.
    """
    lines = []
    for turn in turns or []:
        if not isinstance(turn, dict):
            continue
        role = turn.get('role')
        text = _one_line(_flatten(turn.get('content')))
        if _is_noise(text):
            continue
        if role == 'user':
            lines.append((u'Người dùng hỏi: ' if viet else u'User asked: ')
                         + text)
        elif role == 'assistant' and _CONCRETE.search(text):
            lines.append((u'Kết quả: ' if viet else u'Result: ') + text)
    return lines


def fold_summary(existing, turns, viet=False, max_chars=MAX_SUMMARY_CHARS):
    """Merge `turns` into an existing summary, oldest first, newest kept.

    When the budget is exceeded the OLDEST lines are dropped: a summary that
    forgets the start of a long session is still better than one that forgets
    what just happened.
    """
    lines = [l for l in (existing or u'').split(u'\n') if l.strip()]
    lines.extend(summarize_turns(turns, viet=viet))

    while lines and len(u'\n'.join(lines)) > max_chars:
        lines.pop(0)
    return u'\n'.join(lines)


def trim_history(history, summary=u'', budget_tokens=None, viet=False,
                 max_summary_chars=None, limit=None):
    """Enforce the window, folding whatever falls out into `summary`.

    budget_tokens   verbatim history allowed, in estimated tokens (default
                    DEFAULT_HISTORY_TOKENS); the newest messages that fit
                    are kept
    limit           an explicit message count instead of a token budget
    max_summary_chars  summary size (default MAX_SUMMARY_CHARS)

    Returns (kept_history, new_summary). When everything fits nothing changes
    and the summary is returned untouched.
    """
    hist = list(history or [])
    if limit:
        if len(hist) <= limit:
            return hist, (summary or u'')
        kept, dropped = hist[-limit:], hist[:len(hist) - limit]
    else:
        kept, dropped = context_budget.fit_messages(
            hist, budget_tokens or DEFAULT_HISTORY_TOKENS)
    if not dropped:
        return hist, (summary or u'')
    return kept, fold_summary(summary, dropped, viet=viet,
                              max_chars=max_summary_chars or MAX_SUMMARY_CHARS)


def summary_block(summary, viet=False):
    """The prompt section for a rolling summary, or '' when there is none.

    Belongs at the END of the STATIC half of the system prompt: it changes
    only when turns fall out of the window (every few turns), so the prompt
    cache still hits in between — unlike live Revit state, which changes every
    turn and therefore travels with the user message instead.
    """
    text = (summary or u'').strip()
    if not text:
        return u''
    heading = _HEADING_VI if viet else _HEADING_EN
    note = (u"Tóm tắt các lượt đã trôi khỏi cửa sổ hội thoại. Dùng làm bối "
            u"cảnh; đây KHÔNG phải việc cần làm lại."
            if viet else
            u"Condensed record of turns that have scrolled out of the "
            u"conversation window. Background only — these are NOT tasks to "
            u"redo.")
    return u"{}\n{}\n{}".format(heading, note, text)
