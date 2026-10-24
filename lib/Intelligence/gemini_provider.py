# -*- coding: utf-8 -*-
"""
Gemini Provider

Google Gemini adapter for the T3Lab LLM router, through Google's
OpenAI-compatible endpoint:

    https://generativelanguage.googleapis.com/v1beta/openai/
      POST chat/completions   (tools, streaming, image_url vision)
      GET  models             (live model list)
    Authorization: Bearer <GEMINI_API_KEY>   (key from aistudio.google.com)

See https://ai.google.dev/gemini-api/docs/openai. Models are resolved live
from /models — the names below are preference substrings only.

Thinking: Gemini 2.5 / 3.x reason before answering and count those tokens in
max_tokens; reasoning_effort (low / medium / high) maps onto their thinking
level. Gemini 3 also returns a thought signature on each tool call
(tool_calls[].extra_content) that must be sent back on the next turn — the
shared OpenAI agent helpers keep it on the assistant message.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Gemini Provider"

import re

from Intelligence.openai_compat_provider import OpenAICompatProvider


GEMINI_BASE_URL = "https://generativelanguage.googleapis.com/v1beta/openai"

# Families that are not chat models (embeddings, image / video / speech
# generation, Live API, robotics, computer use) — excluded from the list.
_EXCLUDE_SUBSTRINGS = (
    "embedding", "image", "imagen", "veo", "tts", "audio", "live",
    "robotics", "computer-use", "aqa",
)

# Gemini 1.x / 2.0 do not think; everything newer does.
_NON_THINKING_RE = re.compile(r"gemini-(1\.|2\.0)")


class GeminiProvider(OpenAICompatProvider):
    """Adapter for the Gemini API (OpenAI-compatible Chat Completions)."""

    NAME                  = "gemini"
    DISPLAY_NAME          = "Gemini (Google)"
    SUPPORTS_VISION       = True
    SUPPORTS_NATIVE_TOOLS = True

    KEY_NAME           = "Gemini"
    BASE_URL           = GEMINI_BASE_URL
    EXCLUDE_SUBSTRINGS = _EXCLUDE_SUBSTRINGS
    # Flash first: fast, cheap, vision + tools — the right default for a chat
    # window; stable ids before previews / experiments / lite.
    PREF_MODELS        = ("flash", "pro")
    AVOID_DEFAULT      = ("lite", "preview", "exp")
    FAST_HINTS         = ("flash-lite", "flash")
    SORT_MODELS_DESC   = True        # newest version first (3.x before 2.5)
    TOKEN_PARAM        = "max_tokens"

    @staticmethod
    def _normalize_model_id(model_id):
        """The compat /models list can return "models/gemini-..." ids, while
        chat/completions takes the bare id — store the bare one."""
        mid = u"{}".format(model_id or u"").strip()
        if mid.startswith("models/"):
            mid = mid[len("models/"):]
        return mid

    def _is_chat_model(self, model_id):
        low = model_id.lower()
        if "gemini" not in low:
            return False        # Gemma / Imagen / embeddings share the list
        return not any(x in low for x in self.EXCLUDE_SUBSTRINGS)

    def _is_reasoning(self, model):
        low = u"{}".format(model or u"").lower()
        return "gemini" in low and not _NON_THINKING_RE.search(low)
