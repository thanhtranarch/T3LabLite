# -*- coding: utf-8 -*-
"""
Custom (OpenAI-compatible) Provider

Any server that speaks OpenAI's Chat Completions API, at a Base URL the user
types in LLMs Setting → Models:

    OpenRouter     https://openrouter.ai/api/v1
    Azure OpenAI   https://<resource>.openai.azure.com/openai/v1/   (api-key header)
    Groq           https://api.groq.com/openai/v1
    Mistral        https://api.mistral.ai/v1
    xAI            https://api.x.ai/v1
    vLLM / LiteLLM http://<company-server>:<port>/v1

Settings: endpoints.custom = {"base_url", "model", "azure"} (plain text, not
secret) and the key in api_keys["Custom"] (DPAPI, like every other key; may
be empty for an unauthenticated in-house server).

Auth header: Azure OpenAI takes the key in an `api-key` header — sent when the
"Azure" box is ticked OR the host ends in .azure.com; every other server gets
`Authorization: Bearer <key>`.

Model: the user's pick from the live /models list; when the server has no
/models (or, like Azure, lists base models rather than deployment names) the
"Model name" field is used instead.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "Custom (OpenAI-compatible) Provider"

from Intelligence.openai_compat_provider import (OpenAICompatProvider,
                                                 is_openai_reasoning_model,
                                                 _settings)

# Model-name substrings of families that accept image input. The provider can
# serve anything, so vision is decided per model instead of per provider.
_VISION_HINTS = (
    "gpt-4o", "gpt-4.1", "gpt-5", "gpt-6", "gemini", "claude", "vision",
    "-vl", "pixtral", "llava", "llama-4", "gemma-3", "grok-4",
)

# Path tails people paste by mistake — we append these ourselves.
_ENDPOINT_TAILS = ("/chat/completions", "/models")


def normalize_base_url(url):
    """Server base URL as stored: trimmed, no trailing slash, without a
    pasted "/chat/completions" or "/models" tail."""
    url = (url or u"").strip()
    while url.endswith("/"):
        url = url[:-1]
    low = url.lower()
    for tail in _ENDPOINT_TAILS:
        if low.endswith(tail):
            url = url[:-len(tail)].rstrip("/")
            break
    return url


def is_azure_url(url):
    """True for an Azure OpenAI / Foundry host (*.azure.com)."""
    try:
        rest = (url or u"").split("://", 1)[-1]
        host = rest.split("/", 1)[0].split(":", 1)[0].lower()
    except Exception:
        return False
    return host.endswith(".azure.com")


class CustomProvider(OpenAICompatProvider):
    """Adapter for a user-configured OpenAI-compatible server."""

    NAME                  = "custom"
    DISPLAY_NAME          = "Custom (OpenAI-compatible)"
    SUPPORTS_NATIVE_TOOLS = True

    KEY_NAME      = "Custom"
    REQUIRES_KEY  = False          # in-house vLLM / LiteLLM may need none
    # Non-chat families a gateway (OpenRouter, LiteLLM) may list.
    EXCLUDE_SUBSTRINGS = ("embedding", "embed", "whisper", "tts", "dall-e",
                          "moderation", "rerank", "transcribe")

    # ── Endpoint ──────────────────────────────────────────────────────────────

    def _endpoint(self):
        try:
            return _settings().get_provider_endpoint(self.NAME)
        except Exception:
            return {"base_url": u"", "model": u"", "azure": False}

    def _base_url(self):
        return normalize_base_url(self._endpoint().get("base_url"))

    def _configured_model(self):
        return (self._endpoint().get("model") or u"").strip() or None

    def uses_azure_header(self):
        ep = self._endpoint()
        return bool(ep.get("azure")) or is_azure_url(ep.get("base_url"))

    def _headers(self, api_key):
        """api-key header for Azure, Bearer for everything else. Never logged."""
        if not api_key:
            return {}
        if self.uses_azure_header():
            return {"api-key": api_key}
        return {"Authorization": "Bearer " + api_key}

    def is_configured(self):
        """Configured = a Base URL is saved (the key is optional). No HTTP."""
        try:
            return bool(self._base_url())
        except Exception:
            return False

    # ── Models ────────────────────────────────────────────────────────────────

    def get_models(self):
        """Live /models, with the typed "Model name" put first when the list
        lacks it (Azure lists base models, not deployment names) or stands in
        for it when the server has no /models at all."""
        if self._cached_models is not None:
            return list(self._cached_models)
        if not self._base_url():
            return []
        ids = self._fetch_models()
        configured = self._configured_model()
        if configured and configured not in ids:
            ids.insert(0, configured)
        if ids:
            self._cached_models = ids
        return list(ids)

    def get_active_model(self):
        if self._model:
            return self._model
        return (self._configured_model()
                or self._pick_model(self._cached_models or []))

    def _resolve_model(self, has_vision=False):
        return (self._model or self._configured_model()
                or self._pick_model(self.get_models(), prefer_vision=has_vision))

    # ── Capabilities ──────────────────────────────────────────────────────────

    @property
    def SUPPORTS_VISION(self):
        """Per model: a known vision family gets images, anything else gets
        the "this model cannot see images" text path (never a 400)."""
        try:
            model = u"{}".format(self.get_active_model() or u"").lower()
        except Exception:
            return False
        return any(h in model for h in _VISION_HINTS)

    def _is_reasoning(self, model):
        # Only OpenAI reasoning models (served by Azure / OpenRouter / a
        # LiteLLM proxy) get the reasoning knobs; other vendors' thinking
        # models take plain Chat Completions parameters.
        return is_openai_reasoning_model(model)

    def _default_token_param(self, model):
        return ("max_completion_tokens" if is_openai_reasoning_model(model)
                else "max_tokens")
