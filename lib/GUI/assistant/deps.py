# -*- coding: utf-8 -*-
"""
Assistant dependencies — every optional Intelligence / Services import the
chat window uses, each with the degradation stub it had in the dialog.

A failed import logs a warning and leaves a stub (or, for a few names, nothing
— exactly as before the split). Also: the provider / host probes.

Split out of GUI/T3LabAssistantDialog.py, which re-imports every name
here: `from GUI.T3LabAssistantDialog import <name>` keeps working.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Dependencies"

from GUI.assistant._log import logger
from GUI.assistant.text import _exc_text

# ─── NLP module ───────────────────────────────────────────────────────────────
try:
    from Intelligence.t3lab_assistant import (has_api_key, keyword_parse,
                                              learn_pattern, find_learned_match,
                                              has_local_llm,
                                              get_local_model_name, parse_command_nlu,
                                              inject_discovered_tools,
                                              get_active_provider_name,
                                              get_provider_display_label,
                                              get_setup_guidance_message,
                                              _build_system_prompt,
                                              _RAG_SYSTEM_PREFIX)
    HAS_NLP = True
except Exception as e:
    logger.warning("Could not import t3lab_assistant: {}".format(e))
    HAS_NLP = False
    # The offline / "no answer" paths call keyword_parse unconditionally: a
    # missing name there was a NameError bubble on a broken install. Without
    # the NLP layer nothing parses, so the caller shows the guidance below.
    def keyword_parse(*a, **kw): return None
    def has_api_key(*a, **kw): return False
    def learn_pattern(*a, **kw): pass
    def find_learned_match(*a, **kw): return None
    def has_local_llm(*a, **kw): return False
    def get_local_model_name(*a, **kw): return None
    def parse_command_nlu(*a, **kw): return None
    def inject_discovered_tools(*a, **kw): pass
    def get_active_provider_name(*a, **kw): return "claude"
    def get_provider_display_label(*a, **kw): return "AI"
    def get_setup_guidance_message(viet=True):
        return (u"Chưa hiểu yêu cầu — bạn mô tả cụ thể hơn nhé." if viet
                else "I didn't understand — could you describe it more specifically?")
    def _build_system_prompt(revit_context=u""):
        from Intelligence.t3lab_agent import build_system_prompt
        return build_system_prompt(revit_context=revit_context)
    _RAG_SYSTEM_PREFIX = u""  # fallback: no RAG prefix if NLP module unavailable

# ─── Tool discovery module ────────────────────────────────────────────────────
try:
    from Services.tool_discovery import (discover_new_tools, get_registered_tools,
                                         make_generic_launcher, make_launcher_for)
    HAS_DISCOVERY = True
except Exception as e:
    logger.warning("Could not import tool_discovery: {}".format(e))
    HAS_DISCOVERY = False
    def discover_new_tools(): return []
    def get_registered_tools(): return []
    def make_generic_launcher(script_path, title, url=None): return lambda: False
    def make_launcher_for(entry): return lambda: False

# ─── Revit API context runner ─────────────────────────────────────────────────
# WPF callbacks run while Revit is idle = NOT an API context; tool scripts need
# one (Services/revit_context.py).
try:
    from Services.revit_context import (ensure_api_context, run_in_api_context)
    HAS_API_CONTEXT = True
except Exception as e:
    logger.warning("Could not import revit_context: {}".format(e))
    HAS_API_CONTEXT = False
    def ensure_api_context(): return False, u"revit_context unavailable"
    def run_in_api_context(func, on_done=None):
        try:
            res = func()
            if res is None:
                ok, err = True, u""
            elif isinstance(res, tuple):
                ok = bool(res[0]) if res else False
                err = res[1] if len(res) > 1 else u""
            else:
                ok, err = bool(res), u""
        except Exception as ex:
            ok, err = False, _exc_text(ex)
        if on_done:
            on_done(ok, err or u"")
        return 'inline'

# ─── Context Scout (BIM Context) ────────────────────────────────────────────────
try:
    from Selection.scout import ContextScout
    HAS_SCOUT = True
except Exception as e:
    logger.warning("Could not import ContextScout: {}".format(e))
    HAS_SCOUT = False
    class ContextScout:
        @staticmethod
        def get_context_summary_for_ai(): return ""

# ─── BatchOut executor (configure + direct export) ────────────────────────────
try:
    from Services.batchout_executor import configure_batchout_window, direct_export
    HAS_EXECUTOR = True
except Exception as e:
    logger.warning("Could not import batchout_executor: {}".format(e))
    HAS_EXECUTOR = False

# ─── RAG processor (PDF / image attachments) ──────────────────────────────────
try:
    from Intelligence.rag_processor import (is_supported, is_image, is_pdf,
                                           build_text_context, build_vision_content_blocks,
                                           has_images, summarize_attachments, SUPPORTED_EXTS)
    HAS_RAG = True
except Exception as e:
    logger.warning("Could not import rag_processor: {}".format(e))
    HAS_RAG = False
    def is_supported(p): return False
    def is_image(p): return False
    def is_pdf(p): return False
    def build_text_context(files): return ''
    def build_vision_content_blocks(text, files): return [{"type": "text", "text": text}]
    def has_images(files): return False
    def summarize_attachments(files): return ''
    SUPPORTED_EXTS = set()

# ─── Link reader (URLs / UNC / local paths pasted into the message) ───────────
try:
    from Intelligence import link_reader
    HAS_LINKS = True
except Exception as e:
    logger.warning("Could not import link_reader: {}".format(e))
    link_reader = None
    HAS_LINKS = False

# ─── Knowledge stack (RAG v2: index + retrieval) ──────────────────────────────
try:
    from Intelligence.knowledge.knowledge_store import (get_global_store,
                                                        get_active_store)
    HAS_KNOWLEDGE = True
except Exception as e:
    logger.warning("Could not import knowledge_store: {}".format(e))
    HAS_KNOWLEDGE = False
    def get_global_store(): return None
    def get_active_store(): return None

# ─── Routing decisions (pure, testable — see lib/Intelligence/routing.py) ─────
try:
    from Intelligence import routing
except Exception as e:
    logger.warning("Could not import routing: {}".format(e))

    class routing(object):
        """Degradation stub: never carry over; learned patterns win."""
        @staticmethod
        def is_continuation(*a, **kw): return False

        @staticmethod
        def learned_pattern_wins(learned, nlu_result, raw=None): return bool(learned)

        @staticmethod
        def wants_spellcheck_fix(args): return False

# ─── Bilingual language analysis (VI/EN) ──────────────────────────────────────
# Degrading to the diacritic check is safe: every consumer takes a `default`.
try:
    from Intelligence.language import analyzer as _lang_analyzer
    from Intelligence.language import lang_detect as _lang_detect
    HAS_LANGUAGE = True
except Exception as e:
    logger.warning("Could not import language layer: {}".format(e))
    HAS_LANGUAGE = False
    _lang_analyzer = None
    _lang_detect = None

# ─── Graph agent layer (multi-goal turns) ─────────────────────────────────────
try:
    from Intelligence.graph.orchestrator import GraphOrchestrator
    HAS_GRAPH = True
except Exception as e:
    logger.warning("Could not import graph layer: {}".format(e))
    HAS_GRAPH = False
    GraphOrchestrator = None

# ─── Multi-agent dispatcher (specialist routing) ──────────────────────────────
try:
    from Intelligence.agents.dispatcher import AgentDispatcher
    HAS_AGENTS = True
except Exception as e:
    logger.warning("Could not import AgentDispatcher: {}".format(e))
    HAS_AGENTS = False
    class AgentDispatcher(object):
        def classify(self, *a, **kw):
            return {'specialist': 'general', 'skill': None,
                    'source': 'default', 'confidence': 0.0}

# ─── "Narrated instead of acted" guard (shared with the native AgentLoop) ─────
try:
    from Intelligence.agent_loop import (_announces_work,
                                         _ANNOUNCE_FIXUP)
except Exception as e:
    logger.warning("Could not import the announce guard: {}".format(e))
    _ANNOUNCE_FIXUP = u"Call the tool that answers the request now."

    def _announces_work(_text):
        return False

# ─── Specialist specs (per-agent prompt/tools/budget) ─────────────────────────
try:
    from Intelligence.agents.specialists import get_spec, build_specialist_prompt
    HAS_SPECIALISTS = True
except Exception as e:
    logger.warning("Could not import specialists: {}".format(e))
    HAS_SPECIALISTS = False
    def get_spec(name): return None
    def build_specialist_prompt(spec, revit_context='', **kw):
        from Intelligence.agent_loop import build_agent_system_prompt
        return build_agent_system_prompt(local=kw.get('local', False),
                                         lang=kw.get('lang', 'auto'))

# ─── Turn telemetry (latency + token/cache accounting) ────────────────────────
try:
    from Intelligence import telemetry
    HAS_TELEMETRY = True
except Exception as e:
    logger.warning("Could not import telemetry: {}".format(e))
    HAS_TELEMETRY = False

    class telemetry(object):
        """Degradation stub: measuring nothing must never break a turn."""
        @staticmethod
        def TurnTimer(*a, **kw): return None

        @staticmethod
        def record(turn, path=None): return False

        @staticmethod
        def prune(*a, **kw): return 0

# ─── Streaming message extractor (live token rendering) ───────────────────────
try:
    from Intelligence.llm_provider import StreamingJSONExtractor
except Exception as e:
    logger.warning("Could not import StreamingJSONExtractor: {}".format(e))
    class StreamingJSONExtractor(object):
        """Fallback: show whatever raw text streams in (no JSON unwrapping)."""
        def display(self, raw):
            return (raw or u"").strip()


# ─── Provider / host probes ───────────────────────────────────────────────────

def _provider_usable(provider):
    """True when `provider` is worth calling for a real chat turn.

    Prefers is_configured() (credential check) over check_health(), which
    hits the vendor's /models endpoint.
    """
    if provider is None:
        return False
    try:
        probe = getattr(provider, "is_configured", None) or provider.check_health
        return bool(probe())
    except Exception:
        return False


def _get_uiapp():
    """The live UIApplication, or None.

    `HOST_APP.uiapp`/`uidoc` are None on this pyRevit build (`__revit__` is a
    DB.Application), so wrap it with UIApplication's public constructor.
    """
    try:
        from pyrevit import HOST_APP
        uiapp = getattr(HOST_APP, 'uiapp', None)
        if uiapp is not None:
            return uiapp
    except Exception:
        pass
    try:
        import builtins as __builtin__
        from Autodesk.Revit.UI import UIApplication
        rvt = getattr(__builtin__, '__revit__', None)
        if rvt is None:
            return None
        if isinstance(rvt, UIApplication):
            return rvt
        return UIApplication(rvt)
    except Exception as ex:
        logger.debug("_get_uiapp failed: {}".format(_exc_text(ex)))
        return None
