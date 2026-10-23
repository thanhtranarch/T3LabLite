# -*- coding: utf-8 -*-
"""
Assistant chat history — per-document transcript persistence (save / load /
clear) and the context-budget helpers that decide how much of it reaches the
model.

Split out of GUI/T3LabAssistantDialog.py, which re-imports every name
here: `from GUI.T3LabAssistantDialog import <name>` keeps working.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant History"

import io
import json
import os
import re

from GUI.assistant._log import logger
from GUI.assistant.text import _exc_text

# Same folder the dialog called `lib_dir` (<extension>/lib).
lib_dir = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))))


# ─── Chat history persistence ─────────────────────────────────────────────────

LEGACY_DOC_KEY = "default"

# Transcripts are personal: per user in %APPDATA%\T3LabAI\chat_history, never in
# the shared git clone. The old lib/config/chat_history/ is imported once.
_LEGACY_CHAT_HISTORY_DIR = os.path.join(lib_dir, 'config', 'chat_history')

# Archived conversations (History tab) kept on disk, newest first.
MAX_ARCHIVED_SESSIONS = 100


def _chat_history_dir(sub=None):
    """%APPDATA%\\T3LabAI\\chat_history (or its `sub` folder), created on
    demand. The matching legacy folder is imported once per destination."""
    from core.paths import settings_dir
    from core import housekeeping
    base = os.path.join(settings_dir(), 'chat_history')
    folder = os.path.join(base, sub) if sub else base
    if not os.path.isdir(folder):
        try:
            os.makedirs(folder)
        except OSError:
            pass
    legacy = (os.path.join(_LEGACY_CHAT_HISTORY_DIR, sub) if sub
              else _LEGACY_CHAT_HISTORY_DIR)
    housekeeping.run_once(u'chat_history_legacy:' + folder,
                          housekeeping.import_legacy_folder,
                          legacy, folder, ('.json',))
    return folder


def _prune_archived_sessions(sessions_dir, keep=MAX_ARCHIVED_SESSIONS):
    """Keep the newest `keep` archived sessions. Returns how many went."""
    try:
        from core import housekeeping
        return housekeeping.prune_files(sessions_dir, keep_newest=keep,
                                        suffixes=('.json',), min_age_s=0)
    except Exception:
        return 0


def _active_doc():
    """The document Revit has active right now, or None. Never raises.

    Uses Snippets._host.resolve_doc(): `revit.doc` is None in the docked pane.
    """
    try:
        from Snippets._host import resolve_doc
        doc, _err = resolve_doc()
        return doc
    except Exception:
        return None


def _get_doc_key():
    """Return a filesystem-safe key for the current Revit document.

    Via Snippets._host.resolve_doc(), NOT `revit.doc` (None in this modeless
    pane, which made every model share one "default" history).
    """
    title = None
    try:
        from Snippets._host import resolve_doc
        doc, _err = resolve_doc()
        if doc is not None:
            title = doc.Title
    except Exception:
        pass
    if not title:
        return LEGACY_DOC_KEY
    safe = re.sub(r'[\\/:*?"<>|]', '_', title)
    return safe[:80] or LEGACY_DOC_KEY   # cap at 80 chars


def _history_file(doc_key):
    """Path to doc_key's JSON history: projects/<pid>/chats/ when a project is
    active, else the per-document chat_history folder."""
    try:
        from config.project_store import ProjectStore
        _ps = ProjectStore()
        _pid = _ps.get_active_project_id()
        if _pid:
            return _ps.history_path(_pid, doc_key)
    except Exception:
        pass
    return os.path.join(_chat_history_dir(), '{}.json'.format(doc_key))


def save_chat_history(doc_key, messages, summary=u"", session=None):
    """Persist the whole conversation to disk for this document.

    messages: [{role, content, ts}]; summary: rolling summary of turns out of
    the window; session: {'session_id', 'session_notes', 'session_path'}
    (Intelligence.session_context). Every message is kept; only what reaches
    the model is bounded (context budget).
    """
    try:
        from core import jsonsafe
        path = _history_file(doc_key)
        to_save = list(messages or [])
        # Serialize FIRST, then write in one shot (no truncated file). MUST be
        # jsonsafe, not json.dumps(ensure_ascii=True): IronPython 2.7's escaper
        # raises on non-ASCII and Vietnamese chats were silently never saved.
        payload = {"doc_key": doc_key, "messages": to_save,
                   "summary": summary or u""}
        for key, val in (session or {}).items():
            if val:
                payload[key] = val
        data = jsonsafe.dumps(payload, indent=2)
        with io.open(path, 'w', encoding='utf-8') as f:
            f.write(data)
    except Exception as ex:
        logger.debug(u"Could not save chat history: {}".format(_exc_text(ex)))


def _is_stale_tool_call_blob(role, content):
    """True for a saved assistant turn that is only a tool call written as chat
    text (e.g. {"name": "apply_playbook", "parameters": {...}}).

    Old transcripts carry them; replayed, later turns copy the fake call.
    Dropped on load.
    """
    if role != "assistant":
        return False
    try:
        from Intelligence.agent_loop import is_bare_tool_call_text
        return is_bare_tool_call_text(content)
    except Exception:
        return False


def _history_path_for_read(doc_key):
    """Path to read this document's history from, or None.

    Falls back (read-only) to the legacy shared "default" file so pre-doc-key
    conversations are adopted instead of vanishing.
    """
    path = _history_file(doc_key)
    if os.path.exists(path):
        return path
    if doc_key != LEGACY_DOC_KEY:
        legacy = _history_file(LEGACY_DOC_KEY)
        if os.path.exists(legacy):
            return legacy
    return None


def load_chat_history(doc_key):
    """Load saved messages for doc_key.  Returns [] if none / error."""
    try:
        path = _history_path_for_read(doc_key)
        if not path:
            return []
        with io.open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return [m for m in data.get("messages", [])
                if not _is_stale_tool_call_blob(m.get("role", ""),
                                                m.get("content", ""))]
    except Exception as ex:
        logger.debug(u"Could not load chat history: {}".format(_exc_text(ex)))
        return []


def load_chat_summary(doc_key):
    """Rolling summary saved alongside the messages, or u''."""
    try:
        path = _history_path_for_read(doc_key)
        if not path:
            return u""
        with io.open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data.get("summary") or u""
    except Exception as ex:
        logger.debug(u"Could not load chat summary: {}".format(_exc_text(ex)))
        return u""


def load_chat_session(doc_key):
    """The open chat's identity and notes saved with its history:
    {'session_id', 'session_notes', 'session_path'} ({} when none)."""
    try:
        path = _history_path_for_read(doc_key)
        if not path:
            return {}
        with io.open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return {'session_id': data.get('session_id') or None,
                'session_notes': data.get('session_notes') or [],
                'session_path': data.get('session_path') or None}
    except Exception as ex:
        logger.debug(u"Could not load chat session: {}".format(_exc_text(ex)))
        return {}


def clear_chat_history(doc_key):
    """Delete the saved history file for doc_key."""
    try:
        path = _history_file(doc_key)
        if os.path.exists(path):
            os.remove(path)
    except Exception as ex:
        logger.debug(u"Could not clear chat history: {}".format(_exc_text(ex)))


# ─── Context budget ───────────────────────────────────────────────────────────
# Intelligence.context_budget derives every number from the ACTIVE model's real
# context window (fixed caps were the same for 1M Claude and an 8k local model).

# Bubbles re-rendered on restore/resume; only rendering is bounded here.
MAX_REPLAY_BUBBLES = 120


def _active_budgets():
    """{'window', 'history', 'summary_chars', 'memory_chars',
    'past_chats_chars'} for the active provider. Never raises."""
    try:
        from Intelligence import context_budget as _cb
    except Exception:
        return {'window': 32768, 'history': 1500, 'summary_chars': 1400,
                'memory_chars': 3000, 'past_chats_chars': 1200}
    provider = None
    try:
        from Intelligence.llm_router import LLMRouter
        provider = LLMRouter().get_active_provider()
    except Exception:
        provider = None
    window = _cb.provider_window(provider)
    return {'window': window,
            'history': _cb.history_budget_for(provider),
            'summary_chars': _cb.summary_budget_chars(window),
            'memory_chars': _cb.memory_budget_chars(window),
            'past_chats_chars': _cb.past_chats_budget_chars(window)}


def _budget_history(history, provider=None, system_prompt=None, query=None,
                    max_tokens=None):
    """The newest history that fits `provider`'s window (default: active)
    after THIS request's system prompt and query."""
    try:
        from Intelligence.context_budget import fit_history
        if provider is None:
            from Intelligence.llm_router import LLMRouter
            provider = LLMRouter().get_active_provider()
        return fit_history(history, provider, system_prompt=system_prompt,
                           user=query, reply_tokens=max_tokens)
    except Exception:
        return list(history or [])[-16:]


def _fit_restored(messages, summary, viet=False):
    """(context_window_messages, summary) for a restored / resumed chat.

    The newest messages that fit the budget go to the model; the overflow is
    folded into a summary only when none was saved (no double summarising).
    """
    plain = [{"role": m.get("role", ""), "content": m.get("content", "")}
             for m in (messages or [])
             if m.get("role") in ("user", "assistant") and m.get("content")]
    budgets = _active_budgets()
    try:
        from Intelligence import context_budget as _cb
        kept, dropped = _cb.fit_messages(plain, budgets['history'])
    except Exception:
        kept, dropped = plain[-24:], plain[:-24]
    if dropped and not (summary or u"").strip():
        try:
            from Intelligence.conversation import fold_summary
            summary = fold_summary(u"", dropped, viet=viet,
                                   max_chars=budgets['summary_chars'])
        except Exception:
            pass
    return kept, summary or u""
