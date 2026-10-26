# -*- coding: utf-8 -*-
"""
LLM Provider

Abstract base class and shared HTTP helper for all LLM provider adapters.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "LLM Provider"

import json
import random
import re
import threading
import time

# ─── Shared HTTP backend ───────────────────────────────────────────────────────
# Revit (pythonnet / IronPython): .NET HttpWebRequest.
# CPython (testing): fall back to urllib.

_USE_NET = False
try:
    import clr
    clr.AddReference('System.Net')
    from System.Net import HttpWebRequest
    from System.Text import Encoding as _NetEncoding
    _USE_NET = True
except Exception:
    pass

_HAS_URLLIB = False
if not _USE_NET:
    try:
        from urllib.request import urlopen, Request  # Python 3 / CPython
        _HAS_URLLIB = True
    except ImportError:
        try:
            from urllib2 import urlopen, Request     # Python 2 / IronPython fallback
            _HAS_URLLIB = True
        except Exception:
            pass

HAS_HTTP = _USE_NET or _HAS_URLLIB


def _http_error_detail(ex):
    """Pull the response body out of a failed HTTP call.

    Vendors put the ACTUAL reason in the error body ("Model Not Exist",
    "Insufficient Balance", "does not support function calling", ...) —
    without this, every API rejection surfaces as a bare WebException/
    HTTPError and the UI can only say "the model didn't respond".
    """
    # .NET WebException carries the response object
    try:
        resp = getattr(ex, "Response", None)
        if resp is not None:
            from System.IO import StreamReader
            reader = StreamReader(resp.GetResponseStream(), _NetEncoding.UTF8)
            try:
                body = reader.ReadToEnd()
            finally:
                reader.Close()
            if body:
                return body[:400]
    except Exception:
        pass
    # urllib2.HTTPError is itself file-like
    try:
        read = getattr(ex, "read", None)
        if callable(read):
            body = read()
            if isinstance(body, bytes):
                body = body.decode("utf-8", "replace")
            if body:
                return body[:400]
    except Exception:
        pass
    return None


def _raise_with_detail(ex):
    """Re-raise a transport error, upgrading it to carry the API error body."""
    detail = _http_error_detail(ex)
    if detail:
        raise RuntimeError(u"API error: {}".format(detail))
    raise ex


# ─── Cancel + retry ────────────────────────────────────────────────────────────

class RequestCancelled(Exception):
    """The run was cancelled (user Stop or time budget) while a request was
    in flight. Providers re-raise it instead of falling back to a retry."""


class CancelToken(object):
    """Cancel flag shared by an agent run and its in-flight HTTP request.

    cancel() also aborts the request (HttpWebRequest.Abort), so a read with a
    600 s timeout stops at once. `reason` is "cancelled" or "timeout".
    """

    def __init__(self):
        self._evt    = threading.Event()
        self._lock   = threading.Lock()
        self._aborts = []
        self.reason  = None

    def cancel(self, reason="cancelled"):
        with self._lock:
            if self._evt.is_set():
                return
            self.reason = reason
            self._evt.set()
            aborts = list(self._aborts)
        for fn in aborts:
            try:
                fn()
            except Exception:
                pass

    def is_cancelled(self):
        return self._evt.is_set()

    def wait(self, seconds):
        """Sleep up to `seconds`; True if cancelled meanwhile."""
        return self._evt.wait(seconds)

    def check(self):
        if self._evt.is_set():
            raise RequestCancelled(self.reason or "cancelled")

    def on_cancel(self, fn):
        """Call fn on cancel() (now, if already cancelled). Returns an
        unregister callable."""
        with self._lock:
            if not self._evt.is_set():
                self._aborts.append(fn)

                def _off():
                    with self._lock:
                        try:
                            self._aborts.remove(fn)
                        except ValueError:
                            pass
                return _off
        try:
            fn()
        except Exception:
            pass
        return lambda: None


# 408/409/429/5xx + Anthropic's 529 overloaded — the statuses the official SDKs
# retry. A request is retried only before any response line reached a caller.
RETRY_STATUSES  = (408, 409, 429, 500, 502, 503, 504, 529)
MAX_RETRIES     = 2
_BACKOFF_BASE_S = 1.0
_BACKOFF_CAP_S  = 20.0


def _http_status(ex):
    """HTTP status of a failed request (.NET WebException or urllib HTTPError)."""
    resp = getattr(ex, "Response", None)
    if resp is not None:
        try:
            return int(resp.StatusCode)
        except Exception:
            try:
                from System import Convert
                return int(Convert.ToInt32(resp.StatusCode))
            except Exception:
                return None
    code = getattr(ex, "code", None)
    try:
        return int(code) if code is not None else None
    except Exception:
        return None


def _retry_after_s(ex):
    """`retry-after` in seconds, or None (absent, or the HTTP-date form)."""
    val = None
    resp = getattr(ex, "Response", None)
    if resp is not None:
        try:
            val = resp.Headers["retry-after"]
        except Exception:
            val = None
    if val is None:
        hdrs = getattr(ex, "headers", None)
        try:
            val = hdrs.get("retry-after") if hdrs is not None else None
        except Exception:
            val = None
    try:
        return max(0.0, float(val))
    except Exception:
        return None


def retry_delay_s(attempt, ex):
    """retry-after when given (capped), else exponential backoff with jitter."""
    after = _retry_after_s(ex)
    if after is not None:
        return min(after, _BACKOFF_CAP_S)
    base = min(_BACKOFF_CAP_S, _BACKOFF_BASE_S * (2 ** attempt))
    return base * (0.5 + random.random() / 2.0)


def _should_retry(ex, attempt):
    return attempt < MAX_RETRIES and _http_status(ex) in RETRY_STATUSES


def _backoff(attempt, ex, cancel):
    resp = getattr(ex, "Response", None)
    if resp is not None:
        try:
            resp.Close()          # free the pooled connection before retrying
        except Exception:
            pass
    delay = retry_delay_s(attempt, ex)
    if cancel is not None:
        if cancel.wait(delay):
            raise RequestCancelled(cancel.reason or "cancelled")
    elif delay > 0:
        time.sleep(delay)


def _send(once, cancel, retry_if=None):
    """Run once() with retry on RETRY_STATUSES; transport errors go through
    _raise_with_detail. retry_if() False → no retry (e.g. data streamed)."""
    attempt = 0
    while True:
        if cancel is not None:
            cancel.check()
        try:
            return once()
        except RequestCancelled:
            raise
        except Exception as ex:
            if cancel is not None and cancel.is_cancelled():
                raise RequestCancelled(cancel.reason or "cancelled")
            if (retry_if is None or retry_if()) and _should_retry(ex, attempt):
                _backoff(attempt, ex, cancel)
                attempt += 1
                continue
            _raise_with_detail(ex)


def _net_request(method, url, headers, timeout_ms, body_bytes=None,
                 content_type=None):
    from System.Net import HttpWebRequest
    req = HttpWebRequest.Create(url)
    req.Method           = method
    req.Timeout          = int(timeout_ms)
    req.ReadWriteTimeout = int(timeout_ms)
    if content_type:
        req.ContentType = content_type
    if headers:
        for k, v in headers.items():
            req.Headers.Add(k, v)
    if body_bytes is not None:
        req.ContentLength = body_bytes.Length
        rs = req.GetRequestStream()
        try:
            rs.Write(body_bytes, 0, body_bytes.Length)
        finally:
            rs.Close()
    return req


def _read_all(resp):
    from System.IO import StreamReader
    try:
        reader = StreamReader(resp.GetResponseStream(), _NetEncoding.UTF8)
        try:
            return reader.ReadToEnd()
        finally:
            reader.Close()
    finally:
        resp.Close()


def _get(url, headers, timeout_ms):
    """One GET with retry; raises on failure."""
    def once():
        if _USE_NET:
            return _read_all(_net_request("GET", url, headers, timeout_ms)
                             .GetResponse())
        if _HAS_URLLIB:
            req = Request(url)
            for k, v in (headers or {}).items():
                req.add_header(k, v)
            raw = urlopen(req, timeout=float(timeout_ms) / 1000.0).read()
            return raw.decode("utf-8") if isinstance(raw, bytes) else raw
        raise RuntimeError("No HTTP client available")
    return _send(once, None)


def http_get_auth(url, headers=None, timeout_ms=8000):
    """
    Authenticated GET request with optional headers.
    Returns response text string, or None on error.
    HttpWebRequest, not WebClient: WebClient has no timeout (~100 s default).
    """
    try:
        return _get(url, headers, timeout_ms)
    except Exception:
        return None


def http_post(url, payload, headers=None, timeout_ms=60000, cancel=None):
    """
    POST a JSON-serialisable payload and return the response string.

    timeout_ms: Timeout + ReadWriteTimeout (WebClient has none; local CPU
    inference needs far more than the 60 s cloud default — callers pass it).
    cancel: optional CancelToken; cancel() aborts the request.
    Retries 429/5xx/529 (MAX_RETRIES, backoff, retry-after).

    Returns:
        str: response body, or raises RuntimeError on failure.
    """
    body = json.dumps(payload, ensure_ascii=False)

    def once():
        if _USE_NET:
            data = _NetEncoding.UTF8.GetBytes(body)
            req = _net_request("POST", url, headers, timeout_ms, data,
                               "application/json; charset=utf-8")
            off = cancel.on_cancel(req.Abort) if cancel is not None else None
            try:
                return _read_all(req.GetResponse())
            finally:
                if off:
                    off()
        if _HAS_URLLIB:
            body_bytes = body.encode("utf-8") if isinstance(body, type(u"")) else body
            req_headers = {"Content-Type": "application/json; charset=utf-8"}
            if headers:
                req_headers.update(headers)
            resp = urlopen(Request(url, body_bytes, req_headers),
                           timeout=float(timeout_ms) / 1000.0)
            raw = resp.read()
            return raw.decode("utf-8") if isinstance(raw, bytes) else raw
        raise RuntimeError("No HTTP client available")

    return _send(once, cancel)


def http_get(url, timeout_ms=4000):
    """GET url; return response string or None. Times out after timeout_ms."""
    try:
        return _get(url, None, timeout_ms)
    except Exception:
        return None


# ─── Streaming (Server-Sent Events) backend ────────────────────────────────────

def http_post_stream(url, payload, headers=None, on_line=None, timeout_ms=120000,
                     cancel=None):
    """
    POST a JSON payload and stream the response back line-by-line.

    Each decoded text line goes to on_line(line) as it arrives (SSE / NDJSON).
    429/5xx are retried only before the first line; once a line reached the
    caller nothing is retried. cancel (CancelToken) is checked between lines
    and aborts the request.

    Returns:
        bool: True when the stream completes. Raises on transport failure so the
        caller can fall back to a blocking request; RequestCancelled on cancel.
    """
    body = json.dumps(payload, ensure_ascii=False)
    delivered = [False]

    def _emit(line):
        if cancel is not None and cancel.is_cancelled():
            raise RequestCancelled(cancel.reason or "cancelled")
        delivered[0] = True
        if on_line is not None:
            on_line(line)

    def once():
        if _USE_NET:
            from System.IO import StreamReader
            data = _NetEncoding.UTF8.GetBytes(body)
            req = _net_request("POST", url, headers, timeout_ms, data,
                               "application/json; charset=utf-8")
            off = cancel.on_cancel(req.Abort) if cancel is not None else None
            try:
                resp = req.GetResponse()
                try:
                    reader = StreamReader(resp.GetResponseStream(), _NetEncoding.UTF8)
                    try:
                        while True:
                            line = reader.ReadLine()
                            if line is None:
                                break
                            _emit(line)
                    finally:
                        reader.Close()
                finally:
                    resp.Close()
            finally:
                if off:
                    off()
            return True
        if _HAS_URLLIB:
            body_bytes = body.encode("utf-8") if isinstance(body, type(u"")) else body
            req_headers = {"Content-Type": "application/json; charset=utf-8"}
            if headers:
                req_headers.update(headers)
            resp = urlopen(Request(url, body_bytes, req_headers),
                           timeout=float(timeout_ms) / 1000.0)
            off = cancel.on_cancel(resp.close) if cancel is not None else None
            try:
                for raw_line in resp:
                    line = (raw_line.decode("utf-8") if isinstance(raw_line, bytes)
                            else raw_line)
                    _emit(line.rstrip("\n"))
            finally:
                if off:
                    off()
            return True
        raise RuntimeError("No HTTP client available")

    return _send(once, cancel, retry_if=lambda: not delivered[0])


def parse_anthropic_stream_line(line):
    """Return the text delta carried by one Anthropic SSE line, or None."""
    if not line:
        return None
    line = line.strip()
    if not line.startswith("data:"):
        return None
    data = line[5:].strip()
    if not data or data == "[DONE]":
        return None
    try:
        obj = json.loads(data)
    except Exception:
        return None
    if obj.get("type") == "content_block_delta":
        delta = obj.get("delta", {}) or {}
        if delta.get("type") in ("text_delta", None):
            return delta.get("text")
    return None


def parse_openai_stream_line(line):
    """Return the text delta carried by one OpenAI-format SSE line, or None.

    Shared by OpenAI and DeepSeek (both OpenAI-compatible). Reasoning models
    stream their chain-of-thought under `reasoning_content`, which is ignored —
    only the user-facing `content` delta is surfaced.
    """
    if not line:
        return None
    line = line.strip()
    if not line.startswith("data:"):
        return None
    data = line[5:].strip()
    if not data or data == "[DONE]":
        return None
    try:
        obj = json.loads(data)
    except Exception:
        return None
    try:
        choices = obj.get("choices") or []
        if not choices:
            return None
        delta = choices[0].get("delta") or {}
        return delta.get("content")
    except Exception:
        return None


# ─── Native tool calling (OpenAI wire format — shared by OpenAI/DeepSeek) ──────

# ─── Local reasoning-model sampling ────────────────────────────────────────────
# Substrings that mark a hybrid/reasoning local model (Qwen3, QwQ, DeepSeek-R1,
# Magistral, etc.). These models are trained with sampled decoding and DEGRADE
# under greedy (temperature 0): Qwen's own guidance is explicit that greedy
# decoding in thinking mode causes endless repetition and quality drops. The
# whole codebase historically pinned temperature 0.0 for determinism of tool
# JSON — correct for instruct models, actively harmful for these.
_REASONING_MODEL_HINTS = (
    "qwen3", "qwq", "deepseek-r1", "-r1", "r1-", "magistral",
    "reasoning", "thinker", "marco-o1", "openthinker", "phi-4-reasoning",
)


def is_reasoning_model(model_name):
    """True when the model name looks like a hybrid/reasoning local model."""
    if not model_name:
        return False
    low = u"{}".format(model_name).lower()
    return any(h in low for h in _REASONING_MODEL_HINTS)


def local_sampling_params(model_name):
    """Recommended sampling options for a local model, by family.

    Reasoning models (Qwen3 thinking, DeepSeek-R1, ...) get the vendor-
    recommended non-greedy profile (temp 0.6 / top_p 0.95 / top_k 20 / min_p 0)
    so thinking mode doesn't collapse into repetition. Plain instruct models
    keep the deterministic low-temperature profile that makes tool-call JSON
    stable. Returns a dict of raw option names (temperature/top_p/top_k/min_p)
    — each provider maps them onto its own payload shape.
    """
    if is_reasoning_model(model_name):
        return {"temperature": 0.6, "top_p": 0.95, "top_k": 20, "min_p": 0.0}
    return {"temperature": 0.0, "top_p": 0.9}


def openai_usage(usage):
    """OpenAI-format usage → the Anthropic shape telemetry.add_usage reads
    (input_tokens excludes cache hits). {} when the server sent none."""
    if not isinstance(usage, dict):
        return {}
    try:
        prompt = int(usage.get("prompt_tokens") or 0)
        cached = int((usage.get("prompt_tokens_details") or {}).get("cached_tokens")
                     or usage.get("prompt_cache_hit_tokens") or 0)   # DeepSeek
        return {"input_tokens":            max(0, prompt - cached),
                "output_tokens":           int(usage.get("completion_tokens") or 0),
                "cache_read_input_tokens": cached}
    except Exception:
        return {}


def ollama_usage(obj):
    """Ollama's final chunk / response counts → telemetry's Anthropic shape."""
    try:
        return {"input_tokens":  int(obj.get("prompt_eval_count") or 0),
                "output_tokens": int(obj.get("eval_count") or 0)}
    except Exception:
        return {}


def _openai_payload(model, system_prompt, messages, tools, max_tokens,
                    token_param, extra_payload):
    msgs = []
    if system_prompt:
        msgs.append({"role": "system", "content": system_prompt})
    msgs.extend(list(messages or []))
    # token_param: OpenAI reasoning / GPT-5 models reject max_tokens.
    payload = {"model": model, "messages": msgs, token_param: max_tokens}
    if tools:
        payload["tools"] = tools
    # extra_payload carries provider-specific knobs (sampling for local
    # reasoning models, tool_choice, ...) without changing OpenAI's defaults.
    if extra_payload:
        payload.update(extra_payload)
    return payload


def _openai_assistant_msg(content, raw_calls, reasoning, echo_reasoning):
    msg = {"role": "assistant", "content": content or None}
    if raw_calls:
        msg["tool_calls"] = raw_calls
    # DeepSeek thinking mode + tools: reasoning_content must go back on every
    # follow-up request or the API 400s (api-docs.deepseek.com, thinking mode).
    # OpenAI rejects the unknown field, so only providers that opt in echo it.
    if echo_reasoning and reasoning:
        msg["reasoning_content"] = reasoning
    return msg


def openai_chat_agent(url, headers, model, system_prompt, messages, tools,
                      max_tokens=1500, timeout_ms=180000, extra_payload=None,
                      token_param="max_tokens", cancel=None, echo_reasoning=False):
    """One blocking agentic turn against an OpenAI-compatible /chat/completions.

    `messages` must be OpenAI-native (may contain assistant tool_calls and
    role:"tool" results from earlier iterations) and already end with the
    latest user / tool turn. Raises on transport failure — callers wrap.

    Returns the uniform chat_agent dict:
        {"text", "tool_calls":[{"id","name","args"}], "assistant_msg",
         "stop_reason", "usage"}
    """
    payload = _openai_payload(model, system_prompt, messages, tools, max_tokens,
                              token_param, extra_payload)
    resp_text = http_post(url, payload, headers, timeout_ms=timeout_ms,
                          cancel=cancel)
    data   = json.loads(resp_text)
    choice = (data.get("choices") or [{}])[0] or {}
    msg    = choice.get("message", {}) or {}

    text = msg.get("content") or u""
    # Reasoning models may in-line their chain of thought — never show it.
    text = re.sub(r"<think>[\s\S]*?</think>", "", text).strip()

    raw_calls  = msg.get("tool_calls") or []
    tool_calls = []
    for c in raw_calls:
        fn = c.get("function", {}) or {}
        raw_args = fn.get("arguments")
        if isinstance(raw_args, dict):          # some servers send an object
            args = raw_args
        else:
            try:
                args = json.loads(raw_args) if raw_args else {}
            except Exception:
                args = {}
        tool_calls.append({"id": c.get("id", ""),
                           "name": fn.get("name", ""), "args": args})

    return {
        "text":          text,
        "tool_calls":    tool_calls,
        "assistant_msg": _openai_assistant_msg(
            msg.get("content"), raw_calls, msg.get("reasoning_content"),
            echo_reasoning),
        "stop_reason":   _agent_stop_reason(choice.get("finish_reason"), tool_calls),
        "usage":         openai_usage(data.get("usage")),
    }


def openai_chat_agent_stream(url, headers, model, system_prompt, messages, tools,
                             max_tokens=1500, timeout_ms=180000, on_delta=None,
                             extra_payload=None, token_param="max_tokens",
                             cancel=None, echo_reasoning=False,
                             include_usage=False):
    """Streaming variant of openai_chat_agent for OpenAI-compatible servers.

    Streams visible text through on_delta as it arrives AND accumulates the
    tool_call deltas (which arrive fragmented by index, with `arguments`
    streamed as partial JSON strings) into whole calls. Returns the same
    uniform chat_agent dict. Raises on transport failure so callers can fall
    back to the blocking openai_chat_agent. include_usage asks for the final
    usage chunk (stream_options); a server that sends one anyway is read too.
    """
    payload = _openai_payload(model, system_prompt, messages, tools, max_tokens,
                              token_param, extra_payload)
    payload["stream"] = True
    if include_usage:
        payload["stream_options"] = {"include_usage": True}

    state = {"text": [], "reasoning": [], "tool": {}, "order": [],
             "finish": None, "usage": None}

    def _on_line(line):
        if not line:
            return
        line = line.strip()
        if not line.startswith("data:"):
            return
        data = line[5:].strip()
        if not data or data == "[DONE]":
            return
        try:
            obj = json.loads(data)
        except Exception:
            return
        if obj.get("usage"):
            state["usage"] = obj["usage"]
        choices = obj.get("choices") or []
        if not choices:
            return
        ch    = choices[0]
        delta = ch.get("delta") or {}

        c = delta.get("content")
        if c:
            state["text"].append(c)
            if on_delta:
                try:
                    on_delta(c)
                except Exception:
                    pass
        if delta.get("reasoning_content"):
            state["reasoning"].append(delta["reasoning_content"])

        for tc in (delta.get("tool_calls") or []):
            idx  = tc.get("index", 0)
            slot = state["tool"].get(idx)
            if slot is None:
                slot = {"id": tc.get("id", ""), "name": u"", "args": []}
                state["tool"][idx] = slot
                state["order"].append(idx)
            if tc.get("id"):
                slot["id"] = tc["id"]
            fn = tc.get("function") or {}
            if fn.get("name"):
                slot["name"] = fn["name"]
            if fn.get("arguments"):
                slot["args"].append(fn["arguments"])
            # Gemini 3 (OpenAI-compat layer) carries the call's thought
            # signature here and rejects the next turn when it is not sent
            # back — keep it so the rebuilt call round-trips intact.
            if tc.get("extra_content"):
                slot["extra"] = tc["extra_content"]

        if ch.get("finish_reason"):
            state["finish"] = ch["finish_reason"]

    http_post_stream(url, payload, headers, _on_line, timeout_ms=timeout_ms,
                     cancel=cancel)

    raw_text = u"".join(state["text"])
    text     = re.sub(r"<think>[\s\S]*?</think>", "", raw_text).strip()

    tool_calls = []
    raw_calls  = []
    for idx in state["order"]:
        slot     = state["tool"][idx]
        args_str = u"".join(slot["args"])
        try:
            args = json.loads(args_str) if args_str.strip() else {}
        except Exception:
            args = {}
        tool_calls.append({"id": slot["id"], "name": slot["name"], "args": args})
        raw_call = {"id": slot["id"], "type": "function",
                    "function": {"name": slot["name"],
                                 "arguments": args_str or "{}"}}
        if slot.get("extra"):
            raw_call["extra_content"] = slot["extra"]
        raw_calls.append(raw_call)

    return {
        "text":          text,
        "tool_calls":    tool_calls,
        "assistant_msg": _openai_assistant_msg(
            raw_text, raw_calls, u"".join(state["reasoning"]), echo_reasoning),
        "stop_reason":   _agent_stop_reason(state["finish"], tool_calls),
        "usage":         openai_usage(state["usage"]),
    }


def _agent_stop_reason(finish_reason, tool_calls):
    """OpenAI finish_reason → uniform stop_reason. "length" is max_tokens even
    with tool calls: their arguments may be cut off (agent_loop won't run them)."""
    if finish_reason == "length":
        return "max_tokens"
    if tool_calls:
        return "tool_use"
    if not finish_reason or finish_reason == "stop":
        return "end_turn"
    return finish_reason


def openai_agent_tool_results(tool_calls, result_strs):
    """OpenAI format: one role:"tool" message per call, matched by id.

    `result_strs` are pre-serialized JSON strings (agent_loop truncates them);
    missing entries (cancelled run) are padded to keep the transcript valid.
    """
    out = []
    for i, tc in enumerate(tool_calls):
        res = result_strs[i] if i < len(result_strs) else u'{"cancelled": true}'
        out.append({"role": "tool",
                    "tool_call_id": tc.get("id", ""),
                    "content": res})
    return out


class StreamingJSONExtractor(object):
    """
    Incrementally surface the human-readable `message` value out of a streaming
    JSON response of the form:

        {"intent": "...", "params": {...}, "message": "<reply>"}

    The T3Lab system prompt asks models to answer in JSON, but during streaming
    the user should only ever see the `message` text — never the raw braces.
    Feed the full accumulated raw text on each delta; ``display`` returns the
    best human-readable string so far (partial messages included). If the model
    replies in plain prose instead of JSON, the prose is returned verbatim.
    """

    _ESCAPES = {"n": "\n", "t": "\t", "r": "\r", '"': '"', "\\": "\\", "/": "/"}

    @staticmethod
    def _strip_fences(text):
        t = text.lstrip()
        if t.startswith("```"):
            nl = t.find("\n")
            if nl != -1:
                t = t[nl + 1:]
            stripped = t.rstrip()
            if stripped.endswith("```"):
                t = stripped[:-3]
        return t

    def display(self, raw):
        if not raw:
            return u""
        t = self._strip_fences(raw)
        head = t.lstrip()

        # Plain prose (model ignored the JSON instruction) → show as-is.
        if not head.startswith("{"):
            return raw.strip()

        key_idx = t.find('"message"')
        if key_idx == -1:
            # JSON object opened but the message field hasn't streamed yet.
            return u""

        colon = t.find(":", key_idx)
        if colon == -1:
            return u""
        q = t.find('"', colon)
        if q == -1:
            return u""

        out = []
        i = q + 1
        n = len(t)
        while i < n:
            c = t[i]
            if c == "\\" and i + 1 < n:
                out.append(self._ESCAPES.get(t[i + 1], t[i + 1]))
                i += 2
                continue
            if c == '"':          # unescaped closing quote → end of message
                break
            out.append(c)
            i += 1
        return u"".join(out)


def text_history(provider, messages):
    """Plain-text chat history: user/assistant turns only, content blocks
    flattened (provider.blocks_to_text), empty turns dropped. One copy of the
    loop every chat()/chat_stream() used to repeat."""
    out = []
    for h in (messages or []):
        role    = h.get("role", "user")
        content = h.get("content", "")
        if role not in ("user", "assistant"):
            continue
        if isinstance(content, list):
            content = provider.blocks_to_text(content)
        if content:
            out.append({"role": role, "content": content})
    return out


# ─── Abstract base provider ────────────────────────────────────────────────────

class BaseLLMProvider(object):
    """
    Abstract base class for all LLM provider adapters.

    Subclasses must implement:
        chat(messages, system_prompt, user_content, max_tokens) → str | None
        check_health()                                          → bool
    """

    NAME         = "base"
    DISPLAY_NAME = "Base Provider"

    # True if this provider can handle image content blocks
    SUPPORTS_VISION = False

    # chat_agent(..., cancel=CancelToken) aborts the in-flight request on
    # cancel(). agent_loop passes `cancel` only when this is True.
    SUPPORTS_CANCEL = True

    # True: chat_agent()/agent_tool_results() below are implemented and the
    # Assistant routes tool turns here (T3LabAssistantDialog checks the flag).
    SUPPORTS_NATIVE_TOOLS = False

    def chat_agent(self, system_prompt, messages, tools, on_delta=None,
                   max_tokens=1500, **kwargs):
        """One agentic turn. Returns {"text", "tool_calls": [{"id", "name",
        "args"}], "assistant_msg", "stop_reason", "usage"} or None (reason in
        get_last_error). kwargs: cancel (CancelToken), provider extras."""
        return self._fail(u"chat_agent(): {} has no native tool calling"
                          .format(self.DISPLAY_NAME))

    @staticmethod
    def agent_tool_results(tool_calls, result_strs):
        """Messages answering every tool call of the last turn, in this
        provider's wire format. Only called when SUPPORTS_NATIVE_TOOLS."""
        raise NotImplementedError("agent_tool_results needs SUPPORTS_NATIVE_TOOLS")

    def _record_error(self, msg):
        """Remember the most recent failure AND debug-log it.

        chat()/chat_agent() return None on any failure, which the UI can only
        render as a generic "the model didn't respond". The chat window reads
        get_last_error() to show the user the API's real reason instead.
        """
        try:
            self._last_error = u"{}".format(msg)[:300]
        except Exception:
            self._last_error = u"unknown error"
        self._debug_log(msg)

    def get_last_error(self):
        """Most recent failure message, or None. Cleared on each new call."""
        return getattr(self, "_last_error", None)

    def validate_key(self, api_key):
        """Live model list for an UNSAVED key, or [] (so a bad key is never
        saved). Leaves the saved key and the model cache as they were."""
        saved_cache = getattr(self, "_cached_models", None)
        self._get_api_key = lambda: api_key     # shadows the method
        self._cached_models = None
        try:
            return list(self.get_models() or [])
        except Exception as ex:
            self._record_error(u"key check failed: {}".format(ex))
            return []
        finally:
            del self._get_api_key
            self._cached_models = saved_cache

    # ── Silent-failure guards ─────────────────────────────────────────────────
    # chat()/chat_stream()/chat_agent() all bail out with a bare `return None`
    # when there is no key or no resolvable model. Those two paths are the ONLY
    # ways the providers can fail without touching the network, and because
    # they recorded nothing, the chat window's "the model didn't respond"
    # branch had no reason to show — the user got a generic failure with no
    # Details line and no way to tell "your key is missing" from "the vendor
    # timed out". Route every such bail-out through these instead.

    def _fail(self, reason):
        """Record `reason` as the last error and return None (never raises)."""
        self._record_error(reason)
        return None

    def _fail_no_key(self, where="chat"):
        return self._fail(
            u"{}: no API key saved for {} — add it in Settings → LLMs Setting"
            .format(where, self.DISPLAY_NAME))

    def _fail_no_model(self, where="chat"):
        return self._fail(
            u"{}: no usable model for {} — the live model list came back empty "
            u"(key not verified, no credit, or the /models request was blocked "
            u"by the network/proxy)".format(where, self.DISPLAY_NAME))

    def _fail_no_http(self, where="chat"):
        return self._fail(
            u"{}: HTTP transport unavailable in this engine".format(where))

    def is_configured(self):
        """True when this provider has credentials — WITHOUT any network call.

        Deliberately distinct from check_health(), which for the cloud
        providers does a live GET /models on every call. Behind a corporate
        proxy that probe fails intermittently, and callers that gated the whole
        LLM path on it silently fell back to offline canned replies: the
        assistant told the user to go connect an AI they had already connected.
        Gate on THIS, attempt the real call, and let the API's own error reach
        the user when it genuinely fails.

        Local providers (no API key) override it — for them reachability is the
        only meaningful signal.
        """
        try:
            return bool(self._get_api_key())
        except Exception:
            return False

    def _clear_error(self):
        self._last_error = None

    def _debug_log(self, msg):
        """Best-effort debug log via pyRevit's logger; never raises.

        chat()/check_health() failures here are usually swallowed and
        returned as None/False, which looks identical to "not configured" —
        use this in the except-blocks that wrap the actual network call so a
        real API error (bad key, malformed response, rate limit) leaves a
        trace instead of vanishing silently.
        """
        try:
            from pyrevit import script
            script.get_logger().debug(u"{}: {}".format(self.NAME, msg))
        except Exception:
            pass

    @staticmethod
    def _wants_json(response_format):
        """True when a caller asked for a JSON-only reply.

        Accepts every shape used across the codebase: the OpenAI-style
        {"type": "json_object"} dict that the assistant's tool loop sends, and
        the bare "json" string. Providers that constrain decoding server-side
        (Ollama's `format`) use this so JSON mode is opt-in per call instead of
        forced on paths that need prose.
        """
        if not response_format:
            return False
        if isinstance(response_format, dict):
            return u"json" in u"{}".format(
                response_format.get("type", "")).lower()
        return u"json" in u"{}".format(response_format).lower()

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        """
        Send a chat request and return the raw response text.

        Args:
            messages (list): prior [{role, content}] dicts — conversation history.
                             Content may be a string or a list of content blocks.
            system_prompt (str): system instruction string.
            user_content (str|list): current user input — plain string OR a list
                                     of Claude-format content blocks (text/image).
            max_tokens (int): maximum tokens in the response.

        Returns:
            str | None: raw response text, or None on failure.
        """
        raise NotImplementedError

    def chat_stream(self, messages, system_prompt, user_content,
                    on_delta=None, max_tokens=400, **kwargs):
        """
        Streaming variant of chat(). Calls on_delta(text_chunk) for each piece of
        text as it arrives and returns the full concatenated response.

        The base implementation has no real token streaming: it performs a normal
        blocking chat() and emits the whole reply as a single delta. Providers
        that support Server-Sent Events override this for true incremental output.

        Returns:
            str | None: full response text, or None on failure.
        """
        text = self.chat(messages, system_prompt, user_content, max_tokens, **kwargs)
        if text and on_delta:
            try:
                on_delta(text)
            except Exception:
                pass
        return text

    def check_health(self):
        """Return True if the provider is reachable and has credentials."""
        return False

    def supports_vision(self):
        return self.SUPPORTS_VISION

    def get_models(self):
        """Return a list of model name strings available for this provider."""
        return []

    def get_active_model(self):
        """Return the model name currently in use, or None."""
        return None

    def pick_fast_model(self):
        """Fastest/cheapest model for tiny utility calls (classification).

        Providers override this from their CACHED live model list; None means
        "no faster option known — use the active model".
        """
        return None

    def set_model(self, model_name):
        """
        Set the model to use for future requests.

        Returns:
            bool: True if the model was accepted.
        """
        return False

    # ── Shared utilities ───────────────────────────────────────────────────────


    @staticmethod
    def blocks_to_text(user_content):
        """
        Flatten a list of Claude-format content blocks to a plain text string.
        Used by providers that do not support vision.
        """
        if isinstance(user_content, list):
            parts = []
            for block in user_content:
                if isinstance(block, dict) and block.get("type") == "text":
                    parts.append(block.get("text", ""))
            return "\n".join(parts)
        return user_content or ""

    @staticmethod
    def has_image_blocks(user_content):
        """Return True if user_content contains at least one image block."""
        if not isinstance(user_content, list):
            return False
        for block in user_content:
            if isinstance(block, dict) and block.get("type") == "image":
                return True
        return False
