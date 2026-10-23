# -*- coding: utf-8 -*-
"""
Assistant constants — chat icon glyphs and colours, plan-card answers and
timeouts, progress-step looks, and the explicit "remember ..." trigger.

Split out of GUI/T3LabAssistantDialog.py, which re-imports every name
here: `from GUI.T3LabAssistantDialog import <name>` keeps working.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Constants"

import re

from GUI.assistant.theme import _trgb


# ─── Minimal icon vocabulary for bot chat messages ─────────────────────────────
# Verified Segoe MDL2 glyphs only; chat icons are monochrome, never colour emoji.
_ICON_INFO    = u""   # Info — neutral notices ("no AI configured", RAG note)
_ICON_SEARCH  = u""   # Zoom — tool discovery
_ICON_WARNING = u""   # Warning triangle
_ICON_SUCCESS = u""   # CheckMark
_ICON_SYNC    = u""   # Sync — "in progress" (matches tool-card running glyph)
_ICON_STOP    = u""   # Stop
_ICON_REFRESH = u""   # Refresh (chat cleared)
_ICON_ATTACH  = u""   # Attach (paperclip)
_ICON_ANALYZE = u""   # Analyze — fast-context "instant DB answer" badge
_ICON_LIST    = u""   # List/reference — stats & selection section headers

# Inline message-icon colours, read from the theme at import (dark-mode safe).
_ICON_BLUE  = _trgb('Blue')      # info / discovery
_ICON_AMBER = _trgb('Amber')     # warning
_ICON_GREEN = _trgb('Success')   # success
_ICON_RED   = _trgb('Danger')    # "no / not available" markers
_ICON_SLATE = _trgb('Muted')     # neutral / muted

# Plan/progress glyphs (dev/icons/mdl2_codepoints.tsv); every state also has a word.
_ICON_PENDING = u""   # CircleRing — step not started
_ICON_FAILED  = u""   # Cancel — step failed
_ICON_EDIT    = u""   # Edit — "changes the model" tag


# ─── Plan → approve → run (multi-step plans that change the model) ─────────────
# Plan-card answers to the waiting worker; STOPPED / EXPIRED: agent_loop.PendingDecision.
_PLAN_RUN        = 'run'
_PLAN_EDIT       = 'edit'
_PLAN_CANCEL     = 'cancel'
_PLAN_SUPERSEDED = 'superseded'    # the user sent a new message instead

#: Plan-card wait; long on purpose (the user is reading). Stop / new message end it.
_PLAN_APPROVAL_TIMEOUT = 600

#: Per-call confirmation card (destructive tools, or any edit under
#: "Ask before edits").
_CONFIRM_TIMEOUT = 120

# Step states on the progress card: (glyph, word, theme token).
_STEP_LOOK = {
    'pending':   (_ICON_PENDING, u"Pending",  'Muted'),
    'running':   (_ICON_SYNC,    u"Running",  'Accent'),
    'done':      (_ICON_SUCCESS, u"Done",     'Success'),
    'failed':    (_ICON_FAILED,  u"Failed",   'Danger'),
    'skipped':   (_ICON_STOP,    u"Skipped",  'Muted'),
    'cancelled': (_ICON_STOP,    u"Stopped",  'Muted'),
}


# ─── Persistent memory: explicit "remember ..." save trigger ───────────────────
# Only unambiguous phrasings; bare "nhớ"/"remember to" stay with the LLM
# (remember_fact tool). Diacritic + folded forms keep positions valid on the ORIGINAL text.
_MEM_SAVE_RX = re.compile(
    r'^\s*(?:please\s+|h[aã]y\s+)?'
    r'(?:remember|ghi\s*nh[oớ]|nh[oớ]\s+r[aằ]ng|nh[oớ]\s+l[aà]'
    r'|t[uừ]\s+nay(?:\s+tr[oở]\s+[dđ]i)?|from\s+now\s+on)'
    r'\s*[:,]?\s+(.+)$',
    re.IGNORECASE | re.DOTALL | re.UNICODE)
