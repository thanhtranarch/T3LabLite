# -*- coding: utf-8 -*-
"""Auto model switch for Claude — a fast pre-read picks the model per turn.

Before the real turn, the fastest model (Haiku) reads the request plus a
little context and grades it:

    fast      greetings, short questions, one simple lookup or command → Haiku
    standard  normal work: a few tool calls, edits, explanations     → Sonnet
    deep      multi-step plans, code, audits, costly-if-wrong work   → Opus

It also writes a one-line brief (what the user wants + the key context) that
rides the user turn when Sonnet / Opus take over, so they start from the
pre-read instead of from zero.

Tiers map to provider ROLES (claude_provider.pick_tier_model), so model ids
still come from the live /v1/models list and a role pinned in settings wins.

Switch: agents.auto_model (default on). Off, or with Maximum quality on, the
turn runs on the saved / default model exactly as before.
"""
from __future__ import unicode_literals

import json
import re
import threading
import time

TIERS = ('fast', 'standard', 'deep')
_RANK = {t: i for i, t in enumerate(TIERS)}

# Triage must stay cheap next to the turn it routes: a short cap, no thinking,
# and a hard timeout — when it is slow or fails the turn simply runs on the
# middle tier.
_TRIAGE_MAX_TOKENS = 120
_TRIAGE_TIMEOUT_MS = 6000
_FALLBACK_TIER = 'standard'

# The same words re-asked within a few minutes ("thử lại", a resend) get the
# same grade without another call.
_CACHE_TTL = 600.0
_CACHE_MAX = 64
_cache = {}
_cache_lock = threading.Lock()

# Specialists whose work is never a one-liner, whatever the wording.
_STANDARD_FLOOR_SPECIALISTS = ('modeling', 'multi_doc', 'comment', 'qa_check')

_SYSTEM = (
    "You route requests for a Revit / BIM assistant to a model tier. "
    "Pick the CHEAPEST tier that will still answer the request well.\n"
    "fast - greetings, thanks, short factual or how-to questions, one simple "
    "lookup or one simple command (open a tool, count / list / select "
    "elements, rename one item).\n"
    "standard - normal work: a few tool calls, tagging or editing a set of "
    "elements, explaining a concept, summarising a document, questions that "
    "need care.\n"
    "deep - multi-step plans, writing or fixing code (IronPython, Dynamo, "
    "C#), audits or analysis across many elements, sheets or documents, "
    "design decisions, ambiguous requests, anything costly to get wrong.\n"
    "The request may be in Vietnamese or English.\n"
    'Reply ONLY with JSON: {"tier": "fast|standard|deep", "brief": "<one '
    'line, max 25 words, English: what the user wants and the key '
    'context>"}'
)


def is_enabled(provider_name):
    """True when this turn should be auto-routed."""
    if provider_name != 'claude':
        return False
    try:
        from config.settings import get_settings
        s = get_settings()
        if s.is_quality_mode_enabled():
            return False              # Maximum quality = always the top model
        return bool(s.is_auto_model_enabled())
    except Exception:
        return False


def floor_tier(specialist=None, has_images=False, text=u''):
    """Lowest tier the turn may get, from signals that need no model."""
    floor = 'fast'
    if has_images or specialist in _STANDARD_FLOOR_SPECIALISTS:
        floor = 'standard'
    if text and (u'```' in text or len(text) > 2500):
        floor = 'deep' if u'```' in text else 'standard'
    return floor


def raise_to(tier, floor):
    return tier if _RANK.get(tier, 1) >= _RANK.get(floor, 0) else floor


def parse(raw):
    """{"tier", "brief"} from the triage reply, or None."""
    if not raw:
        return None
    m = re.search(r'\{.*\}', raw, re.DOTALL)
    if m:
        try:
            data = json.loads(m.group(0))
        except Exception:
            data = None
        if isinstance(data, dict):
            tier = u'{}'.format(data.get('tier') or u'').strip().lower()
            if tier in TIERS:
                brief = u' '.join(u'{}'.format(data.get('brief') or u'').split())
                return {'tier': tier, 'brief': brief[:240]}
    # A bare word is a usable answer too.
    word = re.search(r'\b(fast|standard|deep)\b', raw.lower())
    if word:
        return {'tier': word.group(1), 'brief': u''}
    return None


def _cache_key(text, context):
    return u'{}\x00{}'.format(u' '.join((text or u'').lower().split()),
                              (context or u'')[-200:])


def _cache_get(key):
    with _cache_lock:
        hit = _cache.get(key)
        if hit and (time.time() - hit[0]) < _CACHE_TTL:
            return dict(hit[1])
        return None


def _cache_put(key, value):
    with _cache_lock:
        if len(_cache) >= _CACHE_MAX:
            oldest = min(_cache, key=lambda k: _cache[k][0])
            _cache.pop(oldest, None)
        _cache[key] = (time.time(), dict(value))


def _context_text(history, revit_context):
    """Last exchange + the Revit state line — enough to grade a follow-up
    ("làm tiếp cho tầng 2") without shipping the whole transcript."""
    parts = []
    for turn in (history or [])[-2:]:
        role = turn.get('role')
        content = turn.get('content')
        if role not in ('user', 'assistant') or not isinstance(
                content, (type(u''), str)):
            continue
        parts.append(u'{}: {}'.format(role, u' '.join(content.split())[:300]))
    if revit_context:
        parts.append(u'Revit: {}'.format(u' '.join(revit_context.split())[:300]))
    return u'\n'.join(parts)


def triage(provider, text, history=None, revit_context=u'', specialist=None,
           has_images=False):
    """Grade the turn and pick its model. WORKER THREAD (one HTTP call).

    Returns {"tier", "brief", "model", "source"} — source is "triage",
    "cache" or "fallback" (triage failed: middle tier) — or None when the
    provider cannot name a model for any tier (no key / no live list).
    """
    text = (text or u'').strip()
    floor = floor_tier(specialist, has_images, text)
    context = _context_text(history, revit_context)
    key = _cache_key(text, context)

    result = _cache_get(key)
    source = 'cache'
    if result is None:
        source = 'triage'
        fast = None
        try:
            fast = provider.pick_tier_model('fast')
        except Exception:
            fast = None
        raw = None
        if fast and text:
            user = text[:2000]
            if context:
                user = u'Context:\n{}\n\nRequest:\n{}'.format(context, user)
            if specialist and specialist != 'general':
                user += u'\n(Routed to the {} specialist.)'.format(specialist)
            try:
                raw = provider.chat([], _SYSTEM, user,
                                    max_tokens=_TRIAGE_MAX_TOKENS,
                                    model_override=fast, thinking_off=True,
                                    timeout_ms=_TRIAGE_TIMEOUT_MS)
            except Exception:
                raw = None
        result = parse(raw)
        if result is None:
            result = {'tier': _FALLBACK_TIER, 'brief': u''}
            source = 'fallback'
        else:
            _cache_put(key, result)

    tier = raise_to(result['tier'], floor)
    model = None
    try:
        model = provider.pick_tier_model(tier)
    except Exception:
        model = None
    if not model:
        return None
    return {'tier': tier, 'brief': result.get('brief') or u'',
            'model': model, 'source': source}


def brief_note(decision, viet=False):
    """The pre-read handed to Sonnet / Opus with the user turn, or u''.

    Not for the fast tier: there the model that wrote the brief also answers.
    """
    if not decision or decision.get('tier') == 'fast':
        return u''
    brief = (decision.get('brief') or u'').strip()
    if not brief:
        return u''
    return (u'[Pre-read by a fast model (may be imperfect - trust the user\'s '
            u'own words over it): {}]'.format(brief))


def short_name(model):
    """'claude-sonnet-5-5' -> 'Sonnet 5.5' for the activity line / chip."""
    m = re.search(r'(haiku|sonnet|opus|fable|mythos)-(\d+)(?:-(\d+))?',
                  u'{}'.format(model or u''))
    if not m:
        return u'{}'.format(model or u'')
    ver = m.group(2) + (u'.' + m.group(3) if m.group(3) else u'')
    return u'{} {}'.format(m.group(1).capitalize(), ver)
