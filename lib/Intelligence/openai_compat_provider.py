# -*- coding: utf-8 -*-
"""
OpenAI-compatible Provider

Shared base for the cloud providers that speak OpenAI's Chat Completions wire
format: OpenAI itself, Google Gemini (its OpenAI-compatible endpoint) and the
user-configured "Custom (OpenAI-compatible)" server (OpenRouter, Azure OpenAI
v1, vLLM / LiteLLM company servers, Groq, Mistral, xAI, ...).

A subclass only says WHERE the server is (base URL, key slot, auth header) and
HOW to pick a model from the live /models list (preference substrings — never
an exact-name whitelist). Requests, SSE streaming, native tool calling, image
conversion and the "the server rejected a parameter" retries live here once.

DeepSeek and LM Studio keep their own adapters on purpose: DeepSeek reads
`reasoning_content`, LM Studio probes hosts and uses local sampling. Moving
them onto this base would change working behaviour for no gain.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "OpenAI-compatible Provider"

import json
import os
import re
import sys

from Intelligence.llm_provider import (BaseLLMProvider, http_post, http_post_stream,
                                       http_get_auth, parse_openai_stream_line,
                                       openai_chat_agent, openai_chat_agent_stream,
                                       openai_agent_tool_results, HAS_HTTP)


_THINK_RE = re.compile(r"<think>[\s\S]*?</think>")

# Calls whose reply budget is this small are utility calls (classification,
# one-label routing, "Connected OK"): reasoning models get low effort there.
UTILITY_MAX_TOKENS = 256

# Output-cap floors for reasoning models (thinking tokens count against it).
REASONING_UTILITY_TOKENS = 2048
REASONING_MIN_TOKENS     = 8192
REASONING_HIGH_TOKENS    = 32000

# OpenAI reasoning families: o1 / o3 / o4-mini ..., gpt-5 and newer. The
# "*-chat-latest" snapshots are the non-reasoning ChatGPT models.
_OPENAI_REASONING_RE = re.compile(r"^(o\d|gpt-([5-9]|\d{2,}))")


def _ensure_lib_in_path():
    here    = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.dirname(here)
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)


def _settings():
    _ensure_lib_in_path()
    from config.settings import T3LabAISettings
    return T3LabAISettings()


def deep_reasoning_enabled():
    """The "Deep reasoning" switch (agents.extended_thinking), or Maximum
    quality, which implies it — the same rule ClaudeProvider uses."""
    try:
        s = _settings()
        if s.is_quality_mode_enabled():
            return True
        return bool(s.get_agent_option("extended_thinking", False))
    except Exception:
        return False


def is_openai_reasoning_model(model):
    """True for OpenAI reasoning models (o-series, gpt-5+), also when routed
    through a gateway that prefixes the vendor ("openai/gpt-5")."""
    if not model:
        return False
    low = u"{}".format(model).lower().split("/")[-1]
    if "chat" in low:
        return False
    return bool(_OPENAI_REASONING_RE.match(low))


def join_url(base, path):
    """`base` + "/" + `path` with exactly one slash between them, whether or
    not the user typed a trailing slash on the base URL."""
    base = (base or u"").strip().rstrip("/")
    return base + "/" + (path or u"").lstrip("/")


class OpenAICompatProvider(BaseLLMProvider):
    """Base adapter for an OpenAI-compatible /chat/completions server."""

    NAME                  = "openai_compat"
    DISPLAY_NAME          = "OpenAI-compatible"
    SUPPORTS_VISION       = True
    SUPPORTS_NATIVE_TOOLS = True

    KEY_NAME      = None     # api_keys slot in settings.json ("OpenAI", ...)
    REQUIRES_KEY  = True
    BASE_URL      = None     # fixed base; Custom reads it from settings

    # Live model list handling (preference HINTS only, matched in order).
    EXCLUDE_SUBSTRINGS = ()
    PREF_MODELS        = ()   # default when the user picked nothing
    FAST_HINTS         = ()   # pick_fast_model()
    AVOID_DEFAULT      = ()   # skipped by the first default pass
    SORT_MODELS_DESC   = False

    # Default output-cap parameter name ("max_tokens" / "max_completion_tokens").
    TOKEN_PARAM = "max_tokens"

    CHAT_TIMEOUT_MS  = 60000
    AGENT_TIMEOUT_MS = 180000

    def __init__(self):
        self._model         = None   # None → resolve from the live model list
        self._cached_models = None   # filled on first successful /models fetch
        # Per-model verdicts learned from API rejections, so a parameter the
        # server refused once is not sent again this session.
        self._param_override = {}    # model -> token parameter that works
        self._effort_off     = set() # models that refused reasoning_effort
        self._temp_off       = set() # models that refused temperature
        self._budget_cap     = {}    # model -> largest output cap accepted

    # ── Endpoint & credentials ────────────────────────────────────────────────

    def _base_url(self):
        return self.BASE_URL

    def _url(self, path):
        return join_url(self._base_url(), path)

    def _get_api_key(self):
        if not self.KEY_NAME:
            return None
        try:
            return _settings().get_api_key(self.KEY_NAME)
        except Exception:
            return None

    def _headers(self, api_key):
        """Auth headers. Never logged."""
        if not api_key:
            return {}
        return {"Authorization": "Bearer " + api_key}

    def reload_credentials(self):
        """Clear model cache so next get_models() re-fetches live data."""
        self._cached_models = None

    def invalidate_models_cache(self):
        self._cached_models = None

    def _precheck(self, where):
        """None when a request can be made, else the recorded failure (None
        return value of _fail*) wrapped as a tuple so callers can tell."""
        if not HAS_HTTP:
            return (self._fail_no_http(where),)
        if not self._base_url():
            return (self._fail(u"{}: no server URL set for {} — add it in "
                               u"Settings → LLMs Setting".format(
                                   where, self.DISPLAY_NAME)),)
        if self.REQUIRES_KEY and not self._get_api_key():
            return (self._fail_no_key(where),)
        return None

    # ── Health & model discovery ─────────────────────────────────────────────

    def check_health(self):
        """True if the provider is configured and the live list is non-empty."""
        if not HAS_HTTP or not self._base_url():
            return False
        if self.REQUIRES_KEY and not self._get_api_key():
            return False
        return len(self.get_models()) > 0

    @staticmethod
    def _normalize_model_id(model_id):
        return u"{}".format(model_id or u"").strip()

    def _is_chat_model(self, model_id):
        low = model_id.lower()
        return not any(x in low for x in self.EXCLUDE_SUBSTRINGS)

    def _filter_models(self, data):
        """Model ids out of a /models response, chat models only."""
        items = data.get("data") if isinstance(data, dict) else data
        if items is None and isinstance(data, dict):
            items = data.get("models")     # a few gateways use "models"
        ids = []
        for m in (items or []):
            raw = m.get("id") if isinstance(m, dict) else m
            mid = self._normalize_model_id(raw)
            if mid and mid not in ids and self._is_chat_model(mid):
                ids.append(mid)
        if self.SORT_MODELS_DESC:
            ids.sort(reverse=True)
        return ids

    def _fetch_models(self):
        """Live ids from GET {base}/models, or [] (never raises)."""
        if not self._base_url():
            return []
        api_key = self._get_api_key()
        if self.REQUIRES_KEY and not api_key:
            return []
        try:
            text = http_get_auth(self._url("models"), self._headers(api_key))
            if text:
                return self._filter_models(json.loads(text))
        except Exception as ex:
            self._debug_log(u"models fetch failed: {}".format(ex))
        return []

    def get_models(self):
        """Live model list (cached). [] when unconfigured or the fetch failed,
        so callers can treat a non-empty result as a verified connection."""
        if self._cached_models is not None:
            return list(self._cached_models)
        ids = self._fetch_models()
        if ids:
            self._cached_models = ids
        return list(ids)

    def get_active_model(self):
        """User's saved choice, else the default from the CACHED live list.

        Deliberately does no HTTP (badge/sidebar call this) — returns None
        until a live model list has been fetched at least once.
        """
        if self._model:
            return self._model
        return self._pick_model(self._cached_models or [], prefer_vision=False)

    def set_model(self, model_name):
        self._model = model_name or None
        return True

    def _pick_model(self, models, prefer_vision=False):
        """Pick a model from a LIVE list only — never invent a name."""
        if not models:
            return None
        for avoid in (self.AVOID_DEFAULT, ()):
            for hint in self.PREF_MODELS:
                for m in models:
                    low = m.lower()
                    if hint in low and not any(a in low for a in avoid):
                        return m
        return models[0]

    def _resolve_model(self, has_vision=False):
        """The user's explicit choice ALWAYS wins (it came from the vendor's
        own live list); otherwise auto-pick from the LIVE list (may fetch)."""
        if self._model:
            return self._model
        return self._pick_model(self.get_models(), prefer_vision=has_vision)

    def pick_fast_model(self):
        """Fastest model from the CACHED live list (classification calls)."""
        for hint in self.FAST_HINTS:
            for m in (self._cached_models or []):
                if hint in m.lower():
                    return m
        return None

    # ── Request knobs (overridable) ──────────────────────────────────────────

    def _is_reasoning(self, model):
        """True when `model` thinks before answering (budget/effort apply)."""
        return False

    def _default_token_param(self, model):
        return self.TOKEN_PARAM

    def _reasoning_effort(self, model, max_tokens):
        """reasoning_effort to send, or None for the model's own default.

        Utility calls stay quick ("low" — the one level every reasoning
        family accepts); the "Deep reasoning" switch asks for "high"; all
        other turns leave the vendor's default (medium on current models).
        """
        if not self._is_reasoning(model):
            return None
        if int(max_tokens) <= UTILITY_MAX_TOKENS:
            return "low"
        if deep_reasoning_enabled():
            return "high"
        return None

    def _token_budget(self, model, max_tokens):
        """Output cap actually sent. Reasoning models spend part of it on
        thinking (OpenAI max_completion_tokens and Gemini max_tokens both
        count it), so a small visible-reply budget is raised to a floor —
        otherwise a 120-token "Connected OK" comes back empty."""
        mt = int(max_tokens)
        if not self._is_reasoning(model):
            return mt
        if mt <= UTILITY_MAX_TOKENS:
            return max(mt, REASONING_UTILITY_TOKENS)
        if deep_reasoning_enabled():
            return max(mt, REASONING_HIGH_TOKENS)
        return max(mt, REASONING_MIN_TOKENS)

    def _temperature(self, model):
        """Sampling temperature to send, or None (server default)."""
        return None

    def _knobs(self, model, max_tokens):
        """{"param", "budget", "extra"} for one request to `model`."""
        param = (self._param_override.get(model)
                 or self._default_token_param(model))
        extra = {}
        effort = None
        if model not in self._effort_off:
            effort = self._reasoning_effort(model, max_tokens)
        if effort:
            extra["reasoning_effort"] = effort
        temp = None if model in self._temp_off else self._temperature(model)
        if temp is not None:
            extra["temperature"] = temp
        budget = self._token_budget(model, max_tokens)
        cap = self._budget_cap.get(model)
        if cap:
            budget = min(budget, cap)
        return {"param": param, "budget": budget, "extra": extra}

    def _adapt_to_rejection(self, model, knobs, ex):
        """Learn from a 400 that names a parameter. True → retry is useful.

        OpenAI's o-series reject max_tokens, older Azure deployments reject
        max_completion_tokens, Gemini 2.0 and many gateways reject
        reasoning_effort, reasoning models reject temperature. The API body
        (llm_provider._raise_with_detail) names the parameter, so the fix is
        mechanical — and remembered for the model.
        """
        msg = u"{}".format(ex).lower()
        if knobs["extra"].get("reasoning_effort") and "reasoning" in msg:
            self._effort_off.add(model)
            return True
        if "temperature" in knobs["extra"] and "temperature" in msg:
            self._temp_off.add(model)
            return True
        if knobs["param"] == "max_tokens" and "max_tokens" in msg \
                and ("max_completion_tokens" in msg or "unsupported" in msg
                     or "not supported" in msg):
            self._param_override[model] = "max_completion_tokens"
            return True
        if knobs["param"] == "max_completion_tokens" \
                and "max_completion_tokens" in msg \
                and ("unrecognized" in msg or "unsupported" in msg
                     or "not supported" in msg or "unknown" in msg):
            self._param_override[model] = "max_tokens"
            return True
        # Cap above what the model / server allows (DeepSeek-class 8k output,
        # a vLLM server with a short max_model_len): retry once with the
        # limit the error names ("supports at most 32768"), else 4096.
        if knobs["budget"] > 1024 and (
                knobs["param"] in msg or "context length" in msg) and any(
                w in msg for w in ("less than", "too large", "maximum",
                                   "exceed", "at most", "range")):
            cap = min(4096, knobs["budget"] // 2)
            found = re.search(r"at most (\d+)", msg)
            if found and 0 < int(found.group(1)) < knobs["budget"]:
                cap = int(found.group(1))
            if cap < 256:
                return False
            self._budget_cap[model] = cap
            return True
        return False

    def _with_retry(self, model, max_tokens, send, can_retry=None):
        """send(knobs) with up to two parameter-rejection retries."""
        attempt = 0
        while True:
            knobs = self._knobs(model, max_tokens)
            try:
                return send(knobs)
            except Exception as ex:
                attempt += 1
                if attempt > 2 or (can_retry is not None and not can_retry()) \
                        or not self._adapt_to_rejection(model, knobs, ex):
                    raise
                self._debug_log(u"retrying {} without a rejected parameter"
                                .format(model))

    # ── Message building ──────────────────────────────────────────────────────

    def _vision_on(self):
        try:
            return bool(self.SUPPORTS_VISION)
        except Exception:
            return False

    def _history_msgs(self, messages, system_prompt, text, max_tokens):
        """Prior turns as plain-text user/assistant messages, trimmed to the
        model's context budget."""
        hist = messages or []
        try:
            from Intelligence.context_budget import fit_history
            hist = fit_history(hist, self, system_prompt=system_prompt,
                               user=text, reply_tokens=max_tokens)
        except Exception:
            pass
        out = []
        for h in hist:
            role    = h.get("role", "user")
            content = h.get("content", "")
            if role not in ("user", "assistant"):
                continue
            if isinstance(content, list):
                content = self.blocks_to_text(content)
            if content:
                out.append({"role": role, "content": content})
        return out

    def _chat_messages(self, messages, system_prompt, user_content, max_tokens):
        if self._vision_on():
            content = self._to_openai_content(user_content)
        else:
            content = self.blocks_to_text(user_content)
        msgs = []
        if system_prompt:
            msgs.append({"role": "system", "content": system_prompt})
        msgs.extend(self._history_msgs(messages, system_prompt,
                                       self.blocks_to_text(user_content),
                                       max_tokens))
        msgs.append({"role": "user", "content": content})
        return msgs

    @staticmethod
    def _to_openai_content(user_content):
        """Convert Claude-format content blocks to OpenAI format.

        Claude block → {"type":"image","source":{"type":"base64",...}}
        OpenAI block → {"type":"image_url","image_url":{"url":"data:...;base64,..."}}
        """
        if not isinstance(user_content, list):
            return user_content
        converted = []
        for block in user_content:
            if not isinstance(block, dict):
                converted.append(block)
                continue
            if block.get("type") == "image":
                source = block.get("source", {}) or {}
                if source.get("type") == "base64":
                    mime = source.get("media_type", "image/jpeg")
                    data = source.get("data", "")
                    converted.append({
                        "type":      "image_url",
                        "image_url": {"url": "data:{};base64,{}".format(mime, data)},
                    })
            else:
                converted.append(block)
        return converted

    @staticmethod
    def _without_images(content):
        """Text-only copy of a content-block list (non-vision models)."""
        if not isinstance(content, list):
            return content
        kept = [b for b in content
                if not (isinstance(b, dict) and b.get("type") in ("image", "image_url"))]
        return kept or u""

    def _agent_messages(self, messages):
        """Copy an agent-loop message list with Claude image blocks converted
        (or dropped for a non-vision model).

        Only the `content` of list-content turns is touched; tool_calls and
        role:"tool" results pass through untouched, and the caller's list is
        never mutated (the agent loop keeps appending to it across iterations).
        Idempotent — an already-converted image_url block is left alone.
        """
        vision = self._vision_on()
        out = []
        for m in (messages or []):
            if not isinstance(m, dict) or not isinstance(m.get("content"), list):
                out.append(m)
                continue
            converted = dict(m)
            if vision:
                converted["content"] = self._to_openai_content(m["content"])
            else:
                converted["content"] = self._without_images(m["content"])
            out.append(converted)
        return out

    @staticmethod
    def _clean(text):
        return _THINK_RE.sub(u"", text or u"").strip()

    # ── Chat ──────────────────────────────────────────────────────────────────

    def _post_chat(self, msgs, model, headers, max_tokens, timeout_ms):
        """One blocking /chat/completions call → raw message content."""
        def _send(knobs):
            payload = {"model": model, "messages": msgs,
                       knobs["param"]: knobs["budget"]}
            payload.update(knobs["extra"])
            resp_text = http_post(self._url("chat/completions"), payload,
                                  headers, timeout_ms=timeout_ms)
            data = json.loads(resp_text)
            choice = (data.get("choices") or [{}])[0]
            return (choice.get("message") or {}).get("content") or u""
        return self._with_retry(model, max_tokens, _send)

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        """Blocking chat. Image blocks are converted to image_url for a vision
        model; with no model picked, a default comes from the LIVE list."""
        self._clear_error()
        bad = self._precheck("chat()")
        if bad:
            return bad[0]
        has_vision = self.has_image_blocks(user_content)
        model = kwargs.get("model_override") or self._resolve_model(has_vision)
        if not model:
            return self._fail_no_model("chat()")

        msgs    = self._chat_messages(messages, system_prompt, user_content, max_tokens)
        headers = self._headers(self._get_api_key())
        try:
            text = self._clean(self._post_chat(
                msgs, model, headers, max_tokens,
                int(kwargs.get("timeout_ms") or self.CHAT_TIMEOUT_MS)))
            if not text:
                return self._fail(u"chat(): {} returned an empty reply (the "
                                  u"output budget may have gone to reasoning)"
                                  .format(model))
            return text
        except Exception as ex:
            self._record_error(u"chat() failed: {}".format(ex))
            return None

    def chat_stream(self, messages, system_prompt, user_content,
                    on_delta=None, max_tokens=400, **kwargs):
        """Stream a reply token-by-token via the Chat Completions SSE API."""
        self._clear_error()
        bad = self._precheck("chat_stream()")
        if bad:
            return bad[0]
        has_vision = self.has_image_blocks(user_content)
        model = kwargs.get("model_override") or self._resolve_model(has_vision)
        if not model:
            return self._fail_no_model("chat_stream()")

        msgs    = self._chat_messages(messages, system_prompt, user_content, max_tokens)
        headers = self._headers(self._get_api_key())
        timeout = int(kwargs.get("timeout_ms") or self.CHAT_TIMEOUT_MS)
        chunks  = []

        def _on_line(line):
            delta = parse_openai_stream_line(line)
            if delta:
                chunks.append(delta)
                if on_delta:
                    try:
                        on_delta(delta)
                    except Exception:
                        pass

        def _send(knobs):
            payload = {"model": model, "messages": msgs,
                       knobs["param"]: knobs["budget"], "stream": True}
            payload.update(knobs["extra"])
            http_post_stream(self._url("chat/completions"), payload, headers,
                             _on_line)

        try:
            self._with_retry(model, max_tokens, _send,
                             can_retry=lambda: not chunks)
            full = self._clean(u"".join(chunks))
            return full if full else None
        except Exception as ex:
            # Record the reason first — chat() clears the error state on
            # entry, so otherwise the window could only say "no response".
            stream_err = u"chat_stream() failed: {}".format(ex)
            self._record_error(stream_err)
            partial = u"".join(chunks)
            if not partial.strip():
                # Nothing reached the live bubble yet, so a blocking retry
                # cannot visibly swap an answer — keep the safety net.
                result = self.chat(messages, system_prompt, user_content,
                                   max_tokens, **kwargs)
                if result is None:
                    self._record_error(u"{} | blocking retry: {}".format(
                        stream_err, self.get_last_error() or u"no response"))
                return result
            # Partial text ALREADY streamed into the bubble: let the model
            # CONTINUE the same reply (prefilled assistant turn) instead of
            # regenerating a different one. Any failure keeps the partial.
            prefill = partial.rstrip()
            cont = self._continue_after_drop(msgs, model, headers, max_tokens,
                                             prefill, on_delta, timeout)
            if cont:
                return self._clean(prefill + cont)
            return self._clean(prefill)

    def _continue_after_drop(self, msgs, model, headers, max_tokens, prefill,
                             on_delta, timeout_ms):
        """Finish a stream that dropped mid-answer with ONE blocking call.
        Best-effort: any failure returns None. Never raises."""
        try:
            cont_msgs = list(msgs)
            cont_msgs.append({"role": "assistant", "content": prefill})
            cont = self._post_chat(cont_msgs, model, headers, max_tokens,
                                   timeout_ms)
            if cont and on_delta:
                try:
                    on_delta(cont)
                except Exception:
                    pass
            return cont
        except Exception as ex:
            self._record_error(u"stream continuation failed: {}".format(ex))
            return None

    # ── Agentic chat (native tool calling) ────────────────────────────────────

    def chat_agent(self, system_prompt, messages, tools,
                   on_delta=None, max_tokens=1500, **kwargs):
        """One agentic turn via the `tools` parameter.

        Streams first (text reaches on_delta live, tool-call fragments are
        reassembled); a transport/parse failure before any text arrived falls
        back to one blocking call. Returns the uniform chat_agent dict (plus
        "model"), or None with get_last_error() set.
        """
        self._clear_error()
        bad = self._precheck("chat_agent()")
        if bad:
            return bad[0]
        # Agent turns can carry Claude-format image blocks (attached images,
        # active-view snapshots): convert them, and auto-pick a model that
        # can see them.
        agent_msgs = self._agent_messages(messages)
        has_vision = any(self.has_image_blocks(m.get("content"))
                         for m in (messages or []) if isinstance(m, dict))
        model = self._resolve_model(has_vision)
        if not model:
            return self._fail_no_model("chat_agent()")

        url     = self._url("chat/completions")
        headers = self._headers(self._get_api_key())
        state   = {"streamed": False}

        def _delta(chunk):
            state["streamed"] = True
            if on_delta:
                on_delta(chunk)

        def _stream(knobs):
            return openai_chat_agent_stream(
                url, headers, model, system_prompt, agent_msgs, tools,
                knobs["budget"], timeout_ms=self.AGENT_TIMEOUT_MS,
                on_delta=_delta, extra_payload=knobs["extra"],
                token_param=knobs["param"])

        def _blocking(knobs):
            return openai_chat_agent(
                url, headers, model, system_prompt, agent_msgs, tools,
                knobs["budget"], timeout_ms=self.AGENT_TIMEOUT_MS,
                extra_payload=knobs["extra"], token_param=knobs["param"])

        try:
            result = self._with_retry(model, max_tokens, _stream,
                                      can_retry=lambda: not state["streamed"])
            result["model"] = model
            return result
        except Exception as ex:
            self._debug_log(u"chat_agent stream failed, blocking: {}".format(ex))

        try:
            result = self._with_retry(model, max_tokens, _blocking)
            result["model"] = model
            return result
        except Exception as ex:
            self._record_error(u"chat_agent() failed: {}".format(ex))
            return None

    agent_tool_results = staticmethod(openai_agent_tool_results)
