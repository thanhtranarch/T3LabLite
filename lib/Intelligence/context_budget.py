# -*- coding: utf-8 -*-
"""
Context budget — how much conversation the active model can actually take.

The assistant used to cap its context with fixed numbers picked for the
smallest model it supports: 24 messages of verbatim history, a 1400-char
summary, 30 remembered facts, 60 messages on disk, the last 8 or 16 messages
on the legacy paths. A 1M-token Claude got the same 12 exchanges as an 8k
Ollama model, so a long working session forgot its own start long before the
model was anywhere near full.

Now every one of those numbers comes from the model's real context window:

    context_window(provider, model)   input tokens the model accepts
    history_budget(window, overhead)  tokens of verbatim history per turn
    summary_budget_chars(window)      size of the rolling summary of older turns
    memory_budget_chars(window)       remembered facts shown in the prompt
    past_chats_budget_chars(window)   digests of earlier chats in the project

The history budget is what the window has LEFT, not a share of it:

    history = window − overhead − turn reserve

    overhead      the fixed part of the request: system prompt + tool schemas.
                  Measured on every agent turn (request_overhead), because it
                  varies a lot — on a local model the tool subset alone is
                  ~6k tokens for revit_data and ~18k for the general catalog.
                  Until a turn has been measured, a conservative default.
    turn reserve  room one turn needs beyond its input: the reply plus the
                  tool results an agent turn reads.

A fixed share could not be right for both ends: 15% of a 32k local window
left almost nothing free on a general turn and overflowed it, while 15% of a
64k window wasted ~20k free tokens. Window minus overhead adapts to both.

History is never thrown away: whatever no longer fits the budget is folded
into the rolling summary (Intelligence.conversation), the same idea as
Claude's own auto-compaction. The verbatim part is capped at
COMPACT_AT_TOKENS (150k, the default trigger of Anthropic's server-side
compaction) because re-sending a 400k-token transcript every turn costs real
money for no gain — older turns survive in the summary and through
search_past_chats. Overrides per machine:

    T3LAB_CONTEXT_WINDOW   force the window (tokens), e.g. for a local model
                           loaded with a larger context than the default
    T3LAB_HISTORY_BUDGET   force the verbatim-history budget (tokens);
                           0 or "none" = everything the window has left
                           (no 150k compaction ceiling)

Pure Python — no Revit, no WPF, no network. CPython 3 / IronPython 2.7.

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Context Budget"

import os
import re

# ─── Windows (input tokens) ───────────────────────────────────────────────────
# Conservative on purpose: a window set too high makes the API reject the
# request, one set too low only folds a few more turns into the summary.

CLAUDE_LARGE_WINDOW = 1000000   # Opus/Sonnet 4.6+, every 5.x model, Fable, Mythos
CLAUDE_WINDOW       = 200000    # Haiku 4.5, Claude 3.x/4.0–4.5, unknown ids
OPENAI_WINDOW       = 128000    # gpt-4o / gpt-4-turbo and unknown ids
DEEPSEEK_WINDOW     = 64000
GEMINI_WINDOW       = 1000000   # every Gemini 2.x / 3.x model takes ~1M
LMSTUDIO_WINDOW     = 8192      # what a model loads with unless told otherwise
OLLAMA_WINDOW       = 32768     # OllamaProvider.NUM_CTX_MAX default
UNKNOWN_WINDOW      = 32768

# OpenAI: first substring that matches wins (most specific first).
_OPENAI_WINDOWS = (
    ('gpt-4.1', 1000000),
    ('gpt-5', 400000),
    ('o4', 200000),
    ('o3', 200000),
    ('o1', 200000),
    ('gpt-4o', 128000),
    ('gpt-4-turbo', 128000),
    ('gpt-3.5', 16000),
)

_CLAUDE_ID = re.compile(
    r'claude-(opus|sonnet|haiku|fable|mythos)-(\d+)(?:-(\d+))?', re.I)

# ─── Fixed part of a request (overhead) ───────────────────────────────────────
# Defaults for a provider no turn has been measured on yet, from the shipped
# prompts (2026-10-16, chars/3): local = system prompt ~5.4k + the 80
# essential tool schemas ~18.4k + memory/summary/digest floors ~2k; cloud =
# the full ~106-tool catalog ~23.5k + a longer prompt.
DEFAULT_OVERHEAD_TOKENS       = 30000
DEFAULT_OVERHEAD_TOKENS_LOCAL = 26000

# The dialog folds history with the LARGEST overhead of the provider's last
# few turns, so a turn with a big tool catalog never receives more history
# than fits — anything an agent turn had to drop would be in no summary.
_OVERHEAD_SAMPLES = 8
_MEASURED = {}            # provider name -> [overhead tokens, newest last]

# Turn reserve: the reply (max_tokens) + tool results read inside one agent
# turn (each capped at 4000 chars by AgentLoop, old ones folded).
TURN_RESERVE_SHARE       = 0.15
TURN_RESERVE_SHARE_LOCAL = 0.12
TURN_RESERVE_MIN         = 16000
TURN_RESERVE_MIN_LOCAL   = 4096
DEFAULT_REPLY_TOKENS     = 1500

# ─── Shares of the window for the text blocks ────────────────────────────────
SUMMARY_SHARE        = 0.03
MEMORY_SHARE         = 0.03
PAST_CHATS_SHARE     = 0.03

# Verbatim history above this folds into the summary (see module doc).
COMPACT_AT_TOKENS    = 150000
# Always offered, even when the window is "full": fit_messages keeps the
# newest message regardless, and the summary carries the rest.
MIN_HISTORY_TOKENS   = 1500

# Character floors/ceilings for the text blocks. The floors are the sizes the
# old fixed caps had, so a small local model behaves exactly as before.
_MIN_SUMMARY_CHARS    = 1400
_MAX_SUMMARY_CHARS    = 60000
_MIN_MEMORY_CHARS     = 3000
_MAX_MEMORY_CHARS     = 60000
_MIN_PAST_CHATS_CHARS = 1200
_MAX_PAST_CHATS_CHARS = 30000

LOCAL_PROVIDERS = ('ollama', 'lmstudio')

# chars per token. Tokenizers give ~3.5–4 for English and fewer for
# Vietnamese with diacritics; 3 over-counts both, which is the safe side
# (OllamaProvider._num_ctx_for uses the same divisor).
_CHARS_PER_TOKEN = 3.0
_MESSAGE_OVERHEAD_TOKENS = 4


# ─── Overrides ────────────────────────────────────────────────────────────────

def _env_int(name):
    """Positive int from the environment, 0 for "none"/"0", None when unset."""
    raw = (os.environ.get(name) or u'').strip().lower()
    if not raw:
        return None
    if raw in (u'none', u'off', u'unlimited', u'0'):
        return 0
    try:
        val = int(float(raw))
    except (TypeError, ValueError):
        return None
    return val if val > 0 else 0


# ─── Window lookup ────────────────────────────────────────────────────────────

def _claude_window(model):
    m = _CLAUDE_ID.search(model or u'')
    if not m:
        return CLAUDE_WINDOW
    family = m.group(1).lower()
    major = int(m.group(2))
    minor = m.group(3)
    # A dated snapshot (claude-sonnet-4-20250514) has no minor version.
    minor = int(minor) if minor and len(minor) <= 2 else 0
    if family in ('fable', 'mythos'):
        return CLAUDE_LARGE_WINDOW
    if family == 'haiku':
        return CLAUDE_LARGE_WINDOW if major >= 5 else CLAUDE_WINDOW
    return (CLAUDE_LARGE_WINDOW if (major, minor) >= (4, 6)
            else CLAUDE_WINDOW)


def _openai_window(model):
    low = (model or u'').lower()
    for key, window in _OPENAI_WINDOWS:
        if key in low:
            return window
    return OPENAI_WINDOW


def _custom_window(model):
    """A user-configured OpenAI-compatible server can serve anything: use a
    known family's window when the model id names one, else the cautious
    UNKNOWN_WINDOW (an Azure deployment name says nothing about the model)."""
    low = (model or u'').lower().split('/')[-1]
    if low.startswith('claude'):
        return _claude_window(low)
    if 'gemini' in low:
        return GEMINI_WINDOW
    if low.startswith('gpt-') or re.match(r'o\d', low):
        return _openai_window(low)
    return UNKNOWN_WINDOW


def context_window(provider_name, model=None, local_max=None):
    """Input tokens the model accepts. Never raises.

    `local_max` is the provider's own ceiling for local models
    (OllamaProvider.NUM_CTX_MAX), used when given.
    """
    forced = _env_int('T3LAB_CONTEXT_WINDOW')
    if forced:
        return forced
    name = (provider_name or u'').lower()
    if name == 'claude':
        return _claude_window(model)
    if name == 'openai':
        return _openai_window(model)
    if name == 'gemini':
        return GEMINI_WINDOW
    if name == 'deepseek':
        return DEEPSEEK_WINDOW
    if name == 'custom':
        return _custom_window(model)
    if name == 'ollama':
        return int(local_max) if local_max else OLLAMA_WINDOW
    if name == 'lmstudio':
        return int(local_max) if local_max else LMSTUDIO_WINDOW
    return UNKNOWN_WINDOW


def provider_window(provider):
    """context_window() for a provider object (anything with NAME)."""
    if provider is None:
        return UNKNOWN_WINDOW
    name = getattr(provider, 'NAME', u'') or u''
    model = None
    if not is_local(name):
        # Cloud providers answer from their cached model list (no HTTP).
        # Local ones are skipped: their get_active_model() can probe the
        # server, and their window does not depend on the model anyway.
        try:
            model = provider.get_active_model()
        except Exception:
            model = getattr(provider, '_model', None)
    local_max = getattr(provider, 'NUM_CTX_MAX', None)
    return context_window(name, model, local_max=local_max)


def active_window():
    """Window of the provider the router has active right now. Never raises."""
    try:
        from Intelligence.llm_router import LLMRouter
        return provider_window(LLMRouter().get_active_provider())
    except Exception:
        return UNKNOWN_WINDOW


def is_local(provider_name):
    return (provider_name or u'').lower() in LOCAL_PROVIDERS


# ─── Budgets ──────────────────────────────────────────────────────────────────

def _clamp(val, lo, hi):
    return max(lo, min(hi, val))


def default_overhead(local=False):
    return DEFAULT_OVERHEAD_TOKENS_LOCAL if local else DEFAULT_OVERHEAD_TOKENS


def request_overhead(system_prompt=u'', tools=None, extra=None):
    """Estimated tokens of the fixed part of one request: the system prompt,
    the tool schemas exactly as sent, and any `extra` content (str or
    content blocks — e.g. the user turn)."""
    total = estimate_tokens(system_prompt or u'')
    if tools:
        # Claude's deferred tools (defer_loading) stay out of the context
        # until tool search loads one — counting them would only shrink the
        # history for nothing.
        sent = [t for t in tools
                if not (isinstance(t, dict) and t.get('defer_loading'))]
        try:
            import json
            total += estimate_tokens(json.dumps(sent, ensure_ascii=False))
        except Exception:
            total += estimate_tokens(u'{}'.format(sent))
    if extra:
        total += message_tokens({'content': extra})
    return total


def note_overhead(provider_name, tokens):
    """Record the measured overhead of a request this provider just got."""
    try:
        tokens = int(tokens)
    except (TypeError, ValueError):
        return
    if tokens <= 0:
        return
    key = (provider_name or u'').lower()
    samples = list(_MEASURED.get(key) or [])
    samples.append(tokens)
    _MEASURED[key] = samples[-_OVERHEAD_SAMPLES:]


def measured_overhead(provider_name):
    """Largest overhead among the provider's last few measured requests, or
    the default when none was measured yet."""
    samples = _MEASURED.get((provider_name or u'').lower())
    if samples:
        return max(samples)
    return default_overhead(is_local(provider_name))


def turn_reserve(window, local=False, reply_tokens=None):
    """Room one turn needs beyond its input: reply + in-turn tool results."""
    share = TURN_RESERVE_SHARE_LOCAL if local else TURN_RESERVE_SHARE
    floor = TURN_RESERVE_MIN_LOCAL if local else TURN_RESERVE_MIN
    reply = int(reply_tokens or DEFAULT_REPLY_TOKENS)
    return max(int((window or UNKNOWN_WINDOW) * share), floor) + reply


def history_budget(window, local=False, overhead=None, reply_tokens=None):
    """Tokens of verbatim history to send with each turn: what the window
    has left after the fixed part of the request and the turn reserve."""
    forced = _env_int('T3LAB_HISTORY_BUDGET')
    if forced:
        return forced
    if overhead is None:
        overhead = default_overhead(local)
    window = window or UNKNOWN_WINDOW
    budget = int(window - overhead - turn_reserve(window, local, reply_tokens))
    if forced is None:            # 0 / "none" lifts the compaction ceiling
        budget = min(budget, COMPACT_AT_TOKENS)
    return max(budget, MIN_HISTORY_TOKENS)


def history_budget_for(provider, overhead=None, reply_tokens=None):
    """history_budget() for a provider object. Without `overhead`, the
    provider's measured overhead (measured_overhead) is used."""
    name = getattr(provider, 'NAME', u'') if provider is not None else u''
    if overhead is None:
        overhead = measured_overhead(name)
    return history_budget(provider_window(provider), local=is_local(name),
                          overhead=overhead, reply_tokens=reply_tokens)


def summary_budget_chars(window):
    return int(_clamp((window or UNKNOWN_WINDOW) * SUMMARY_SHARE
                      * _CHARS_PER_TOKEN,
                      _MIN_SUMMARY_CHARS, _MAX_SUMMARY_CHARS))


def memory_budget_chars(window):
    return int(_clamp((window or UNKNOWN_WINDOW) * MEMORY_SHARE
                      * _CHARS_PER_TOKEN,
                      _MIN_MEMORY_CHARS, _MAX_MEMORY_CHARS))


def past_chats_budget_chars(window):
    return int(_clamp((window or UNKNOWN_WINDOW) * PAST_CHATS_SHARE
                      * _CHARS_PER_TOKEN,
                      _MIN_PAST_CHATS_CHARS, _MAX_PAST_CHATS_CHARS))


# ─── Counting & fitting ───────────────────────────────────────────────────────

_IMAGE_TOKENS = 1600


def _flatten(content):
    """(text, image_count) of any stored content shape."""
    if isinstance(content, (list, tuple)):
        parts, images = [], 0
        for b in content:
            if isinstance(b, dict):
                if b.get('type') == 'image':
                    images += 1
                else:
                    parts.append(u'{}'.format(b.get('text') or u''))
            else:
                parts.append(u'{}'.format(b))
        return u' '.join(parts), images
    return u'{}'.format(content or u''), 0


def estimate_tokens(text):
    """Rough token count of a string (over-counts on purpose)."""
    if not text:
        return 0
    return int(len(text) / _CHARS_PER_TOKEN) + 1


def message_tokens(msg):
    if not isinstance(msg, dict):
        return 0
    text, images = _flatten(msg.get('content'))
    return (estimate_tokens(text) + images * _IMAGE_TOKENS
            + _MESSAGE_OVERHEAD_TOKENS)


def history_tokens(messages):
    return sum(message_tokens(m) for m in (messages or []))


def fit_messages(messages, budget_tokens):
    """Newest messages that fit `budget_tokens`. Returns (kept, dropped).

    The newest message is always kept, even alone over budget — a turn with
    no context beats a turn with none of the user's last words. Order is
    preserved; nothing is reworded. budget_tokens None/0 keeps everything.
    """
    msgs = list(messages or [])
    if not budget_tokens or not msgs:
        return msgs, []
    total = 0
    cut = len(msgs)
    for i in range(len(msgs) - 1, -1, -1):
        cost = message_tokens(msgs[i])
        if total + cost > budget_tokens and cut < len(msgs):
            break
        total += cost
        cut = i
    return msgs[cut:], msgs[:cut]


def fit_history(messages, provider, system_prompt=None, tools=None,
                user=None, reply_tokens=None):
    """fit_messages() against a provider's history budget; kept list only.

    With `system_prompt` (and optionally `tools` / `user`) the overhead of
    THIS request is measured; otherwise the provider's measured overhead is
    used (measured_overhead).
    """
    overhead = None
    if system_prompt is not None or tools or user:
        overhead = request_overhead(system_prompt, tools, user)
    kept, _dropped = fit_messages(
        messages, history_budget_for(provider, overhead=overhead,
                                     reply_tokens=reply_tokens))
    return kept
