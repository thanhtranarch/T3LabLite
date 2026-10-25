# -*- coding: utf-8 -*-
"""
OpenAI Provider

OpenAI GPT API adapter for the T3Lab LLM router (Chat Completions).
Models are discovered live from /v1/models after the key verifies; vision input
is supported. The request / streaming / tool-calling plumbing is shared with
Gemini and Custom through OpenAICompatProvider.

Reasoning models (o-series, gpt-5 and newer):
  - the output cap is sent as max_completion_tokens (max_tokens is deprecated
    and rejected by the o-series; every current chat model accepts the new name)
  - no temperature is sent (reasoning models only accept the default)
  - reasoning_effort: "high" when Deep reasoning is on, "low" for tiny utility
    calls, otherwise the model's own default

Chat Completions stays the wire format on purpose: the agent loop, DeepSeek,
LM Studio, Gemini and Custom all share its message/tool-call shape. The
Responses API would need its own transcript conversion for no gain here.

Vision format conversion:
  Claude block  → {"type":"image","source":{"type":"base64","media_type":"...","data":"..."}}
  OpenAI block  → {"type":"image_url","image_url":{"url":"data:<mime>;base64,..."}}

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "OpenAI Provider"

import os
import sys

from Intelligence.openai_compat_provider import (OpenAICompatProvider,
                                                 is_openai_reasoning_model)


# ─── Constants ─────────────────────────────────────────────────────────────────

OPENAI_BASE_URL   = "https://api.openai.com/v1"
OPENAI_CHAT_URL   = OPENAI_BASE_URL + "/chat/completions"
OPENAI_MODELS_URL = OPENAI_BASE_URL + "/models"

# NO hardcoded default model: what the account can use comes exclusively from
# the live /v1/models endpoint after the key is verified. The names below are
# only capability/preference HINTS matched against that live list.

# Substrings that identify models Chat Completions cannot serve (embeddings,
# audio, images, video, moderation) or that are Responses-API-only (codex,
# computer-use, deep-research, *-pro).
_EXCLUDE_SUBSTRINGS = (
    "embedding", "whisper", "tts", "dall-e", "davinci",
    "babbage", "ada", "curie", "moderation", "realtime",
    "audio", "instruct", "search", "similarity", "edit",
    "image", "transcribe", "codex", "sora", "computer-use",
    "deep-research", "-pro",
)

# Substring hints (capable-first, newest family first) used ONLY to auto-pick
# a default from the live list — never an exact-name whitelist.
_PREF_VISION = ("gpt-6", "gpt-5", "gpt-4.1", "gpt-4o", "gpt-4-turbo", "chatgpt")


# ─── Provider ──────────────────────────────────────────────────────────────────

class OpenAIProvider(OpenAICompatProvider):
    """Adapter for the OpenAI Chat Completions API."""

    NAME                  = "openai"
    DISPLAY_NAME          = "GPT (OpenAI)"
    SUPPORTS_VISION       = True
    SUPPORTS_NATIVE_TOOLS = True
    STREAM_USAGE          = True   # OpenAI sends the usage chunk on request

    KEY_NAME           = "OpenAI"
    BASE_URL           = OPENAI_BASE_URL
    EXCLUDE_SUBSTRINGS = _EXCLUDE_SUBSTRINGS
    PREF_MODELS        = _PREF_VISION
    AVOID_DEFAULT      = ("nano", "chat-latest")
    FAST_HINTS         = ("nano", "mini")
    SORT_MODELS_DESC   = True       # newest ids first (gpt-5.5 before gpt-5)
    TOKEN_PARAM        = "max_completion_tokens"

    def _pick_model(self, models, prefer_vision=False):
        """Pick a model from a LIVE list only — never invent a name.

        Plain chat defaults to the newest family's "mini" (cheap, fast, still
        vision + tools); images / agent turns with no choice made take the
        newest full model. o-series minis are not a plain-chat default: they
        are slower and the newest GPT mini beats them on everyday turns.
        """
        if not models:
            return None
        if not prefer_vision:
            for hint in self.PREF_MODELS:
                for m in models:
                    low = m.lower()
                    if hint in low and "mini" in low:
                        return m
        return OpenAICompatProvider._pick_model(self, models, prefer_vision)

    def _is_reasoning(self, model):
        return is_openai_reasoning_model(model)


# ─── Helpers ───────────────────────────────────────────────────────────────────

def _is_chat_model(model_id):
    """Return True if the model ID looks like a chat-capable model."""
    lower = model_id.lower()
    return not any(x in lower for x in _EXCLUDE_SUBSTRINGS)


def _ensure_lib_in_path():
    here    = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.dirname(here)
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)
