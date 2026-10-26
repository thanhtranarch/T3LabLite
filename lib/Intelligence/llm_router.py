# -*- coding: utf-8 -*-
"""
LLM Router

Singleton that manages LLM provider selection, hot-swap, and opt-in fallback.

Usage:
    from Intelligence.llm_router import LLMRouter

    router = LLMRouter()
    router.switch_provider("openai")              # hot-swap
    router.set_model("openai", "gpt-4o")          # change model
    text = router.chat(history, sys_prompt, user_content)

    status = router.get_status()
    # {"claude": {"available": True, "model": "claude-haiku-...", ...}, ...}

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__  = "LLM Router"

import os
import sys


def _debug_log(msg):
    """Best-effort debug log via pyRevit's logger; never raises.

    Several except-blocks in this module used to swallow failures with zero
    trace (e.g. a provider silently failing to import looked identical to
    "not installed"). Lazy-imported so this module stays importable outside
    a pyRevit/Revit process.
    """
    try:
        from pyrevit import script
        script.get_logger().debug(msg)
    except Exception:
        pass


# ─── Router ────────────────────────────────────────────────────────────────────

class LLMRouter(object):
    """
    Singleton router for multi-provider LLM access.

    Provider display / fallback order:
        claude → openai → gemini → deepseek → custom → ollama → lmstudio

    Fallback is OFF by default (set_fallback): the prompt carries Revit
    context, so it must never reach a vendor the user did not pick. Even when
    on, a local provider never falls back to a cloud one.

    Switching is hot (no session restart required).
    The active provider and per-provider model are persisted in settings.json.
    """

    _instance = None

    # Canonical fallback order — also used as the display order in UI
    FALLBACK_CHAIN = ["claude", "openai", "gemini", "deepseek", "custom",
                      "ollama", "lmstudio"]
    # "custom" is any OpenAI-compatible URL, so it is not counted as local.
    LOCAL_PROVIDERS = ("ollama", "lmstudio")

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(LLMRouter, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        import threading
        self._providers     = {}
        self._active_name   = _default_provider()
        self._fallback_on   = False
        self._last_error    = None
        self._status_cache  = None   # cached result of get_status()
        self._status_ts     = 0.0    # epoch time of last cache fill
        # Guards reads/writes of _status_cache/_status_ts. script.py fires
        # get_status()/probe_provider() from several independent background
        # threads (startup probe, sidebar open, provider switch) that can
        # overlap; without a lock, one thread's cache write can race another
        # thread's cache read/iteration over the same dict.
        self._status_lock   = threading.Lock()
        # Bumped on every provider/model change. The router is shared by
        # every window in the Revit session (the docked Assistant pane, the
        # ribbon's LLMs Setting dialog, ...), so a change made in one must be
        # noticed by the others — see get_state_version().
        self._state_version = 0
        self._initialized   = True

        self._load_providers()
        self._restore_settings()

    # ── Initialisation helpers ─────────────────────────────────────────────────

    def _load_providers(self):
        """Instantiate all provider adapters (failures are silenced)."""
        adapters = [
            ("claude",    "Intelligence.claude_provider",    "ClaudeProvider"),
            ("openai",    "Intelligence.openai_provider",    "OpenAIProvider"),
            ("gemini",    "Intelligence.gemini_provider",    "GeminiProvider"),
            ("deepseek",  "Intelligence.deepseek_provider",  "DeepSeekProvider"),
            ("custom",    "Intelligence.custom_provider",    "CustomProvider"),
            ("ollama",    "Intelligence.ollama_provider",    "OllamaProvider"),
            ("lmstudio",  "Intelligence.lmstudio_provider",  "LMStudioProvider"),
        ]
        for name, module_path, class_name in adapters:
            try:
                parts  = module_path.split(".")
                mod    = __import__(module_path, fromlist=[parts[-1]])
                cls_   = getattr(mod, class_name)
                self._providers[name] = cls_()
            except Exception as ex:
                # A real bug here (typo, missing dependency) looks identical
                # to "provider not installed" without this — log it so the
                # difference is visible when debugging why a provider never
                # shows up.
                _debug_log("LLMRouter: failed to load provider '{}': {}".format(name, ex))

    def _restore_settings(self):
        """Load the last-used provider and per-provider model from settings."""
        try:
            _ensure_lib_in_path()
            from config.settings import T3LabAISettings
            s = T3LabAISettings()

            saved_provider = s.get_active_provider()
            if saved_provider and saved_provider in self._providers:
                self._active_name = saved_provider

            for name, provider in self._providers.items():
                saved_model = s.get_provider_model(name)
                if saved_model:
                    try:
                        provider.set_model(saved_model)
                    except Exception:
                        pass
        except Exception as ex:
            _debug_log("LLMRouter: failed to restore settings: {}".format(ex))

    # ── Provider management ────────────────────────────────────────────────────

    def get_active_name(self):
        """Return the name of the currently active provider."""
        return self._active_name

    def get_active_provider(self):
        """Return the active provider instance, or None."""
        return self._providers.get(self._active_name)

    def switch_provider(self, name, model=None, persist=True):
        """
        Hot-swap to a different provider.

        Args:
            name (str): provider name — one of FALLBACK_CHAIN.
            model (str|None): optional model to activate on the new provider.
            persist (bool): write the choice to settings.json as the user's
                global default. Pass False for a SCOPED switch (e.g. a project
                workspace's provider override) — otherwise activating a project
                permanently rewrites the user's default and deactivating it
                never restores what they had.

        Returns:
            bool: True on success, False if provider is unknown.
        """
        if name not in self._providers:
            return False

        self._active_name = name
        # Note: do NOT invalidate the status cache here. The only thing a switch
        # changes is the "active" flag, which get_status() applies live to the
        # cached snapshot. Invalidating would force a slow full re-probe of every
        # provider on the next status call (e.g. opening the badge menu).

        from config.settings import T3LabAISettings
        s = T3LabAISettings()

        if model is None:
            model = s.get_provider_model(name)

        if model:
            self._providers[name].set_model(model)
        else:
            try:
                model = self._providers[name].get_active_model()
            except Exception:
                pass

        # Persist provider selection (skipped for scoped switches)
        if persist:
            try:
                s.set_active_provider(name)
            except Exception:
                pass

        # A switch changes which provider is "active"; the cached snapshot
        # tracks that live, but the model may have changed too.
        with self._status_lock:
            if self._status_cache and name in self._status_cache and model:
                self._status_cache[name]["model"] = model

        self._bump_state()
        return True

    def set_model(self, provider_name, model_name):
        """
        Set the model for a specific provider and persist the choice.

        Returns:
            bool: True if the provider exists and accepted the model.
        """
        provider = self._providers.get(provider_name)
        if not provider:
            return False

        ok = provider.set_model(model_name)
        if ok:
            try:
                _ensure_lib_in_path()
                from config.settings import T3LabAISettings
                s = T3LabAISettings()
                s.set_provider_model(provider_name, model_name)
            except Exception:
                pass
            # Keep the cached snapshot honest. get_status_instant() /
            # get_status_cached() are what the assistant's composer chip and the
            # settings dialog render from, and nothing re-probes on a timer — so
            # without this the UI keeps showing the PREVIOUS model indefinitely.
            with self._status_lock:
                if self._status_cache and provider_name in self._status_cache:
                    self._status_cache[provider_name]["model"] = model_name
            self._bump_state()

        return ok

    def _bump_state(self):
        with self._status_lock:
            self._state_version = getattr(self, '_state_version', 0) + 1

    def get_state_version(self):
        """Counter that changes whenever the active provider or a provider's
        model changes. Cheap — UI code compares it to the value it last
        rendered to know when its provider chip has gone stale (e.g. the
        provider was switched from the ribbon's LLMs Setting while the docked
        Assistant pane stayed open)."""
        return getattr(self, '_state_version', 0)


    def get_provider(self, name):
        """Return a provider instance by name, or None."""
        return self._providers.get(name)

    def get_provider_names(self):
        """Return list of loaded provider names in FALLBACK_CHAIN order."""
        return [n for n in self.FALLBACK_CHAIN if n in self._providers]


    # ── Fallback ──────────────────────────────────────────────────────────────

    def set_fallback(self, enabled):
        """Opt in/out of retrying a failed request on other providers."""
        self._fallback_on = bool(enabled)

    def is_fallback_enabled(self):
        return self._fallback_on

    def get_last_error(self):
        """Why the last chat()/chat_stream() returned None, or None."""
        return self._last_error

    def get_fallback_candidates(self):
        """Providers a failed request may move to: none unless opted in, and
        never local → cloud."""
        if not self._fallback_on:
            return []
        local_only = self._active_name in self.LOCAL_PROVIDERS
        out = []
        for name in self.FALLBACK_CHAIN:
            if name == self._active_name or name not in self._providers:
                continue
            if local_only and name not in self.LOCAL_PROVIDERS:
                continue
            out.append(name)
        return out

    def _active_failed(self, active, ex=None):
        """Record the active provider's failure on it (get_last_error) and here."""
        if active is None:
            self._last_error = u"no active LLM provider '{}'".format(self._active_name)
            return
        if ex is not None:
            try:
                active._record_error(u"{} failed: {}".format(active.NAME, ex))
            except Exception:
                pass
        err = None
        try:
            err = active.get_last_error()
        except Exception:
            pass
        self._last_error = err or u"{} returned no response".format(
            getattr(active, "DISPLAY_NAME", self._active_name))

    def _route(self, call, kwargs):
        """Run `call(provider, kwargs)` on the active provider, then (opt-in)
        on the fallback candidates. model_override names the active vendor's
        model, so it is never forwarded to another vendor."""
        self._last_error = None
        active = self._providers.get(self._active_name)
        if active:
            try:
                result = call(active, kwargs)
                if result is not None:
                    return result
                self._active_failed(active)
            except Exception as ex:
                self._active_failed(active, ex)
        else:
            self._active_failed(None)

        other_kwargs = dict(kwargs)
        other_kwargs.pop("model_override", None)
        for name in self.get_fallback_candidates():
            provider = self._providers[name]
            try:
                if not provider.check_health():
                    continue
                result = call(provider, other_kwargs)
                if result is not None:
                    _debug_log(u"LLMRouter: fell back from {} to {} ({})".format(
                        self._active_name, name, self._last_error))
                    return result
            except Exception as ex:
                _debug_log(u"LLMRouter: fallback {} failed: {}".format(name, ex))
        return None

    # ── Chat ──────────────────────────────────────────────────────────────────

    def chat(self, messages, system_prompt, user_content, max_tokens=400, **kwargs):
        """
        Route a chat request to the active provider and return raw response text.

        On failure the reason is kept in get_last_error() (router and active
        provider); other providers are tried only after set_fallback(True).

        Returns:
            str | None: raw response text, or None on failure.
        """
        return self._route(
            lambda p, kw: p.chat(messages, system_prompt, user_content,
                                 max_tokens, **kw),
            kwargs)

    def chat_stream(self, messages, system_prompt, user_content,
                    on_delta=None, max_tokens=400, **kwargs):
        """Streaming chat(); on_delta(text_chunk) gets each streamed piece."""
        return self._route(
            lambda p, kw: p.chat_stream(messages, system_prompt, user_content,
                                        on_delta, max_tokens, **kw),
            kwargs)

    # ── Status ────────────────────────────────────────────────────────────────

    def get_status(self, use_cache=True):
        """
        Return a health snapshot of every loaded provider.

        Results are cached for 30 seconds so repeated calls (e.g. sidebar
        re-renders) are instant. Pass use_cache=False to force a fresh probe.

        Returns:
            dict: {
                "claude": {"available": bool, "model": str|None,
                           "display_name": str, "active": bool,
                           "supports_vision": bool},
                ...
            }
        """
        import time
        now = time.time()
        with self._status_lock:
            if use_cache and self._status_cache is not None and (now - self._status_ts) < 30:
                # Return cached snapshot but update the "active" flag live
                snap = {}
                for name, info in self._status_cache.items():
                    snap[name] = dict(info)
                    snap[name]["active"] = (name == self._active_name)
                return snap

        import threading
        status = {}
        threads = []
        # Local lock, deliberately NOT self._status_lock: a worker calling
        # check_health() must never contend with (or re-enter) the router-level
        # lock the block below acquires. IronPython has no GIL, so five threads
        # inserting into a plain dict can collide during a resize.
        status_lock = threading.Lock()

        def probe_worker(name):
            provider = self._providers[name]
            try:
                # Every provider gets a REAL health check. This used to report a
                # non-active remote provider "available" from mere API-key
                # presence, while probe_provider() used check_health() — both
                # wrote the same cache, so the settings dialog could show a
                # revoked key as red and then flip it green two seconds later.
                # The workers run in parallel, so wall-clock is bounded by the
                # slowest provider, not the sum.
                available = provider.check_health()
            except Exception:
                available = False
            try:
                model = provider.get_active_model()
            except Exception:
                model = None
            info = {
                "available":       available,
                "model":           model,
                "display_name":    provider.DISPLAY_NAME,
                "active":          (name == self._active_name),
                "supports_vision": provider.SUPPORTS_VISION,
                "probed":          True,
            }
            with status_lock:
                status[name] = info

        probe_order = self.get_local_provider_names() + self.get_remote_provider_names()
        for name in probe_order:
            t = threading.Thread(target=probe_worker, args=(name,))
            t.daemon = True
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        with self._status_lock:
            self._status_cache = status
            self._status_ts    = now
        return status

    def get_status_cached(self):
        """
        Return the cached status snapshot WITHOUT ever probing — instant, safe
        to call on the UI thread. Returns None when nothing has been probed yet.

        get_status(use_cache=True) still re-probes when the 30s TTL lapsed,
        which is a multi-second, network-bound operation — UI render paths
        (sidebar refresh, badge menu) must use this instead and leave probing
        to background threads.
        """
        with self._status_lock:
            if not self._status_cache:
                return None
            snap = {}
            for name, info in self._status_cache.items():
                snap[name] = dict(info)
                snap[name]["active"] = (name == self._active_name)
            return snap

    def get_status_instant(self):
        """
        Zero-HTTP status snapshot for instant UI: cached data when present,
        else availability approximated by key-presence (remote) / False
        (local). Never blocks on the network.
        """
        snap = self.get_status_cached() or {}
        for name, provider in self._providers.items():
            if name in snap:
                continue
            try:
                if name in ("ollama", "lmstudio"):
                    available = False          # needs a port probe → unknown yet
                elif hasattr(provider, "is_configured"):
                    # Key presence (Custom: a saved Base URL) — no HTTP.
                    available = bool(provider.is_configured())
                else:
                    available = False
            except Exception:
                available = False
            # NEVER call provider.get_active_model() here. For Ollama/LM Studio
            # it does real HTTP (tag probes, and local_llm's .NET GET used to
            # inherit a ~100s timeout), and this function is called from the UI
            # thread — that combination froze Revit while opening LLMs Setting.
            # The saved preference is the only zero-HTTP source of truth.
            try:
                _ensure_lib_in_path()
                from config.settings import T3LabAISettings
                model = T3LabAISettings().get_provider_model(name)
            except Exception:
                model = None
            snap[name] = {
                "available":       available,
                "model":           model,
                "display_name":    provider.DISPLAY_NAME,
                "active":          (name == self._active_name),
                "supports_vision": provider.SUPPORTS_VISION,
                "probed":          False,   # guessed, not health-checked
            }
        return snap

    def probe_provider(self, name):
        """
        Probe a SINGLE provider and merge its result into the status cache.

        Much faster than get_status() when you only need to refresh the active
        provider (e.g. right after a hot-swap). Returns the provider's status
        dict or None if the provider is unknown.
        """
        provider = self._providers.get(name)
        if not provider:
            return None
        try:
            available = provider.check_health()
        except Exception:
            available = False
        try:
            model = provider.get_active_model()
        except Exception:
            model = None
        info = {
            "available":       available,
            "model":           model,
            "display_name":    provider.DISPLAY_NAME,
            "active":          (name == self._active_name),
            "supports_vision": provider.SUPPORTS_VISION,
            "probed":          True,
        }
        # Merge into cache so subsequent get_status(use_cache=True) sees it.
        # Arm the 30s TTL only once the cache covers EVERY loaded provider —
        # arming it after a single-provider probe would serve a partial
        # snapshot for 30s and blank the other providers' status dots.
        with self._status_lock:
            if self._status_cache is None:
                self._status_cache = {}
            self._status_cache[name] = dict(info)
            if set(self._status_cache) >= set(self._providers):
                import time as _t
                self._status_ts = _t.time()
        return info

    def get_local_provider_names(self):
        """Return loaded local (no-API-key) provider names — probed first for speed."""
        return [n for n in self.LOCAL_PROVIDERS if n in self._providers]

    def get_remote_provider_names(self):
        """Return loaded remote (API-key) provider names."""
        return [n for n in self.FALLBACK_CHAIN
                if n in self._providers and n not in self.LOCAL_PROVIDERS]


    def get_display_label(self):
        """
        Return a short label for the UI badge (e.g. "GPT-4o-mini", "LOCAL").
        Mirrors the badge logic that was previously in script.py.
        """
        provider = self._providers.get(self._active_name)
        if not provider:
            return "OFFLINE"

        name = self._active_name
        try:
            model = provider.get_active_model() or ""
        except Exception:
            model = ""

        if name == "ollama":
            return "LOCAL"
        if name == "lmstudio":
            short = model.split("/")[-1] if "/" in model else model
            return short[:14] if short else "LM Studio"
        if name == "claude":
            # Auto model switch: the model is picked per turn, so naming one
            # here would be wrong most of the time.
            try:
                from Intelligence import model_tier
                if model_tier.is_enabled(name):
                    return "Claude · Auto"
            except Exception:
                pass
            # Shorten the model name for display
            if "haiku" in model:
                return "Haiku"
            if "sonnet" in model:
                return "Sonnet"
            if "opus" in model:
                return "Opus"
            return "Claude"
        if name == "openai":
            return _gpt_label(model)
        if name == "gemini":
            low = model.lower()
            if "flash" in low:
                return "Gemini Flash"
            if "pro" in low:
                return "Gemini Pro"
            return "Gemini"
        if name == "custom":
            short = model.split("/")[-1] if "/" in model else model
            return short[:14] if short else "Custom"
        if name == "deepseek":
            if "reasoner" in model:
                return "DS Reasoner"
            return "DeepSeek"

        return name.upper()


def _gpt_label(model):
    """Badge text for an OpenAI model id: "gpt-5.5" → "GPT-5.5",
    "gpt-5-mini" → "GPT-5 mini", "o4-mini" → "o4-mini"."""
    low = (model or "").lower()
    if low.startswith("gpt-"):
        parts = low[4:].split("-")
        label = "GPT-" + parts[0]
        for size in ("mini", "nano"):
            if size in parts:
                return label + " " + size
        return label
    if low[:1] == "o" and low[1:2].isdigit():
        return low[:14]
    return "GPT"


# ─── Module-level singleton accessor ──────────────────────────────────────────

def get_router():
    """Return the global LLMRouter singleton."""
    return LLMRouter()


# ─── Pane chat helpers ──────────────────────────────────────────────────────────
# Used by AssistantPaneControl (DockablePane) for lightweight chat without the
# Assistant window's agent / JSON-intent pipeline.

_PANE_SYSTEM_PROMPT = (
    "You are T3Lab Assistant, an AI helper integrated directly into Autodesk Revit. "
    "You help architects, engineers, and BIM managers with Revit workflows, BIM coordination, "
    "and T3Lab tool usage. "
    "Be concise and practical — aim for under 120 words unless the user asks for detail. "
    "Use plain text (no Markdown). "
    "Respond in the SAME language the user wrote in."
)


def _default_provider():
    try:
        _ensure_lib_in_path()
        from config.settings import DEFAULT_PROVIDER
        return DEFAULT_PROVIDER
    except Exception:
        return "ollama"


# ─── Path helper ───────────────────────────────────────────────────────────────

def _ensure_lib_in_path():
    here    = os.path.dirname(os.path.abspath(__file__))
    lib_dir = os.path.dirname(here)
    if lib_dir not in sys.path:
        sys.path.insert(0, lib_dir)
