# -*- coding: utf-8 -*-
"""
Local LLM

Ollama integration for local LLM inference without API keys.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Local LLM"

"""Quick start
-----------
1. Install Ollama:  https://ollama.ai
2. Pull a model. For the AGENTIC assistant (tool-calling) pick Qwen3 by VRAM —
   avoid <4B for multi-tool work, they misfire tools:
       ollama pull qwen3:4b          # ~2.7 GB  6-8 GB VRAM  (floor)
       ollama pull qwen3:8b          # ~5.2 GB  8-12 GB VRAM (balanced)
       ollama pull qwen3:14b         # ~9.3 GB  12-16 GB VRAM ★ sweet spot
       ollama pull qwen3:30b-a3b     # ~19 GB   24-32 GB VRAM (MoE: fast + smart)
   Tiny models (qwen3:1.7b, qwen2.5:0.5b) are fine only for the lightweight
   NLU/intent path, not for reliable agentic tool-calling.
   Measure on YOUR hardware:  python3 dev/bench_local_models.py
3. Start Ollama (it auto-starts on most systems after install).
4. Open T3Lab Assistant — the "LOCAL" badge will appear in the header.
   Turn on Settings → "Maximum quality" to auto-pick the strongest installed
   model and enable deep thinking.

Environment variable
--------------------
  OLLAMA_HOST  override the server address (default http://localhost:11434)
"""

import json
import os

# ─── HTTP back-end selection ───────────────────────────────────────────────────
# Try .NET HttpWebRequest first (available in IronPython / pyRevit)
# then fall back to standard-library urllib.
# HttpWebRequest, not WebClient: WebClient cannot express a timeout.

_USE_NET = False
try:
    import clr
    clr.AddReference('System.Net')
    from System.Net import (WebException, ServicePointManager,
                            HttpWebRequest)
    from System.IO import StreamReader
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
            from urllib2 import urlopen, Request     # Python 2 / IronPython
            _HAS_URLLIB = True
        except Exception:
            pass


# ─── Configuration ─────────────────────────────────────────────────────────────

OLLAMA_HOST    = os.environ.get("OLLAMA_HOST", "http://localhost:11434")
TIMEOUT_GEN    = 60   # seconds — generation call
TIMEOUT_PROBE  = 3    # seconds — availability ping

# Auto-pick order for the DEFAULT (non-quality) path: smallest/fastest first,
# used mainly by the lightweight NLU/intent fallback. The agentic assistant
# prefers an explicit user choice, and quality mode uses capability scoring
# (get_best_model(prefer_capable=True), which ranks reasoning family + size),
# so this list only needs to name-recognize the common local models —
# including Qwen3, the recommended family for tool-calling.
PREFERRED_MODELS = [
    "qwen2.5:0.5b",
    "qwen3:0.6b",
    "qwen2.5:1.5b",
    "qwen3:1.7b",
    "llama3.2:1b",
    "phi3:mini",
    "gemma2:2b",
    "qwen2.5:3b",
    "qwen3:4b",
    "llama3.2:3b",
    "mistral:7b",
    "qwen2.5:7b",
    "qwen3:8b",
    "llama3:8b",
    "qwen3:14b",
    "qwen3:30b-a3b",
    "qwen3:32b",
]

# ─── HTTP helpers ──────────────────────────────────────────────────────────────

def _post_json(url, payload, timeout=TIMEOUT_GEN):
    """POST a JSON-serialisable payload; return response string.

    Same WebClient trap as _get_text: WebClient has no timeout knob, so the
    `timeout` argument was silently ignored on the .NET path and the call
    inherited ~100s regardless of what the caller asked for.
    """
    body = json.dumps(payload, ensure_ascii=False)
    if isinstance(body, type(u"")):
        body_bytes = body.encode("utf-8")
    else:
        body_bytes = body

    if _USE_NET:
        req = HttpWebRequest.Create(url)
        req.Method = "POST"
        req.ContentType = "application/json; charset=utf-8"
        req.Timeout = int(timeout * 1000)
        req.ReadWriteTimeout = int(timeout * 1000)
        req.ContentLength = len(body_bytes)
        stream = req.GetRequestStream()
        try:
            stream.Write(body_bytes, 0, len(body_bytes))
        finally:
            stream.Close()
        resp = req.GetResponse()
        try:
            reader = StreamReader(resp.GetResponseStream(), _NetEncoding.UTF8)
            try:
                return reader.ReadToEnd()
            finally:
                reader.Close()
        finally:
            resp.Close()

    if _HAS_URLLIB:
        req = Request(url, body_bytes,
                      {"Content-Type": "application/json; charset=utf-8"})
        resp = urlopen(req, timeout=timeout)
        raw = resp.read()
        return raw.decode("utf-8") if isinstance(raw, bytes) else raw

    raise RuntimeError("No HTTP client available")


def _get_text(url, timeout=TIMEOUT_PROBE):
    """GET url; return response string or None.

    The .NET branch uses HttpWebRequest, NOT WebClient: WebClient exposes no
    timeout and inherits .NET's ~100s default, so this function ignored its own
    `timeout` argument entirely. A firewalled (DROP, not REJECT) Ollama port
    then froze the caller for a minute and a half — and because
    get_status_instant() reaches here from the UI thread, that froze Revit.
    """
    try:
        if _USE_NET:
            req = HttpWebRequest.Create(url)
            req.Method = "GET"
            req.Timeout = int(timeout * 1000)
            req.ReadWriteTimeout = int(timeout * 1000)
            resp = req.GetResponse()
            try:
                reader = StreamReader(resp.GetResponseStream(), _NetEncoding.UTF8)
                try:
                    return reader.ReadToEnd()
                finally:
                    reader.Close()
            finally:
                resp.Close()
        if _HAS_URLLIB:
            resp = urlopen(url, timeout=timeout)
            raw = resp.read()
            return raw.decode("utf-8") if isinstance(raw, bytes) else raw
    except Exception:
        return None


# ─── Public API ───────────────────────────────────────────────────────────────

def get_host():
    """Resolve the Ollama base URL: saved setting → env/default constant.

    Module-level OLLAMA_HOST is read from os.environ at import time, so it can
    never see a host the user configured in LLMs Setting. Every public function
    here goes through this instead.
    """
    try:
        from config.settings import T3LabAISettings
        host = T3LabAISettings().get_api_key("Ollama_Host")
        if host:
            return host.rstrip("/")
    except Exception:
        pass
    return OLLAMA_HOST


def is_running():
    """Return True if Ollama server is reachable."""
    text = _get_text(get_host() + "/api/tags")
    return text is not None


def list_models(host=None):
    """Return list of installed model name strings (or empty list)."""
    text = _get_text((host or get_host()) + "/api/tags")
    if not text:
        return []
    try:
        data = json.loads(text)
        return [m.get("name", "") for m in data.get("models", [])]
    except Exception:
        return []


def _param_billions(name):
    """Parse the parameter count from a model tag, e.g. 'qwen3:14b' → 14.0.
    Returns 0.0 when no size is present (unknown → treated as small)."""
    import re
    m = re.search(r'(\d+(?:\.\d+)?)\s*b\b', (name or "").lower())
    try:
        return float(m.group(1)) if m else 0.0
    except Exception:
        return 0.0


# Below this parameter count a model misfires multi-tool agentic work (see the
# module header: "avoid <4B for multi-tool work"). Used only to WARN the user
# and recommend a bigger model — never to block a choice.
TOOL_CALLING_MIN_B = 4.0


def is_tool_capable_size(name):
    """False when `name` has a KNOWN parameter count below the tool-calling floor.

    Unknown size (a custom tag with no "<n>b") returns True — we never nag on a
    model we can't measure. Use for a soft warning, not a hard gate.
    """
    b = _param_billions(name)
    return b == 0.0 or b >= TOOL_CALLING_MIN_B


def recommended_tool_model(installed=None):
    """A reasonable tool-calling model to suggest — the best installed one at or
    above the floor, else the documented sweet-spot tag to pull."""
    for n in (installed or []):
        if _param_billions(n) >= TOOL_CALLING_MIN_B:
            return n
    return "qwen3:14b"


def pick_best(installed, prefer_capable=False):
    """Rank an ALREADY-FETCHED model list and return the best name, or None.

    Pure — no HTTP. Split out of get_best_model so a caller that already probed
    the right host (OllamaProvider, which honours a configured remote URL) can
    rank those names instead of triggering a second discovery against the
    module-level default host.
    """
    if not installed:
        return None
    if prefer_capable:
        try:
            from Intelligence.llm_provider import is_reasoning_model
        except Exception:
            is_reasoning_model = lambda _n: False
        def _score(n):
            return (1 if is_reasoning_model(n) else 0, _param_billions(n))
        return sorted(installed, key=_score, reverse=True)[0]
    # Try preferred list first
    for pref in PREFERRED_MODELS:
        pref_base = pref.split(":")[0]
        for inst in installed:
            if inst == pref or inst.startswith(pref_base + ":"):
                return inst
    return installed[0]


# Preferred agentic (tool-calling) models, sweet-spot FIRST. Used by the
# assistant's main path (OllamaProvider.get_active_model), and deliberately
# distinct from PREFERRED_MODELS above — that list is smallest-first, right for
# the lightweight NLU/intent fallback but wrong for reliable multi-tool work
# (see the module header: "avoid <4B for multi-tool work"). qwen3:14b is the
# documented sweet spot; larger MoE/dense variants and 8b/4b follow.
AGENTIC_PREFERRED = [
    "qwen3:14b",
    "qwen3:30b-a3b",
    "qwen3:32b",
    "qwen3:8b",
    "qwen2.5:14b",
    "qwen2.5:7b",
    "qwen3:4b",
    "qwen2.5:3b",
]


def pick_tool_capable(installed):
    """Best tool-calling model for the agentic assistant, or None.

    Pure — no HTTP. Ranks an ALREADY-FETCHED list:
      1) the documented sweet-spot Qwen tiers, in order (14b → … → 4b);
      2) else any installed model at/above the tool-calling floor, Qwen and
         reasoning families first, then largest parameter count;
      3) else fall back to pick_best (fastest sensible pick) so a box with only
         tiny models still gets an answer — just with a soft warning elsewhere.
    """
    if not installed:
        return None
    for pref in AGENTIC_PREFERRED:
        if pref in installed:
            return pref
    capable = [n for n in installed if _param_billions(n) >= TOOL_CALLING_MIN_B]
    if capable:
        try:
            from Intelligence.llm_provider import is_reasoning_model
        except Exception:
            is_reasoning_model = lambda _n: False

        def _score(n):
            low = (n or "").lower()
            return (1 if "qwen" in low else 0,
                    1 if is_reasoning_model(n) else 0,
                    _param_billions(n))
        return sorted(capable, key=_score, reverse=True)[0]
    return pick_best(installed)


def get_best_model(prefer_capable=False, host=None):
    """Return the best installed model name, or None if none installed.

    Default (prefer_capable=False): smallest/fastest first via PREFERRED_MODELS
    — right for the lightweight NLU/intent path.

    prefer_capable=True (quality mode): pick the strongest installed model —
    reasoning models (Qwen3, ...) first, then largest parameter count. This
    generalizes to whatever Qwen3 variant the user actually installed instead
    of relying on a hardcoded whitelist.
    """
    return pick_best(list_models(host=host), prefer_capable=prefer_capable)
