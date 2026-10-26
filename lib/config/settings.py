# -*- coding: utf-8 -*-
"""
Settings

Configuration settings manager for T3Lab AI integration.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Settings"

import io
import os
import json

from config import secret_store as secrets


# Personal instructions (Claude's "personal preferences"): one free-text block
# the user writes once in LLMs Setting → General, injected into EVERY chat and
# EVERY project ahead of the project's own instructions. Capped so a pasted
# document can never eat the system-prompt budget of every turn.
MAX_USER_INSTRUCTIONS_CHARS = 2000

# Values `agents.reply_language` accepts (get/set_reply_language below).
REPLY_LANGUAGES = ('auto', 'en', 'vi')

# First-run provider (local: privacy + zero API cost) and the fallback when
# settings.json has no active_provider — one value for both.
DEFAULT_PROVIDER = 'ollama'

# Read-only fallback when no key is saved; never written to settings.json.
API_KEY_ENV_VARS = {'Claude': 'ANTHROPIC_API_KEY', 'OpenAI': 'OPENAI_API_KEY',
                    'DeepSeek': 'DEEPSEEK_API_KEY'}


def _clean_instructions(text):
    """Normalise personal-instruction text: str, LF newlines, trimmed, capped.

    Interior blank lines are kept (they are how people group preferences);
    only the outer whitespace and runs of 3+ newlines are folded.
    """
    if text is None:
        return u''
    if not isinstance(text, type(u'')):
        try:
            text = text.decode('utf-8')          # py2 str / bytes
        except Exception:
            text = u'{}'.format(text)
    text = text.replace(u'\r\n', u'\n').replace(u'\r', u'\n').strip()
    while u'\n\n\n' in text:
        text = text.replace(u'\n\n\n', u'\n\n')
    if len(text) > MAX_USER_INSTRUCTIONS_CHARS:
        text = text[:MAX_USER_INSTRUCTIONS_CHARS].rstrip()
    return text


class T3LabAISettings(object):
    """Settings manager for T3LabAI"""

    _instance = None

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(T3LabAISettings, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return

        self._settings_file = self._get_settings_path()
        self._load_ok       = True    # False → the file exists but we couldn't read it
        self._load_error    = None
        self._settings = self._load_settings()
        self._initialized = True
        self.migrate_plaintext_keys()

    def _get_settings_path(self):
        """Get the path to settings file"""
        app_data = os.environ.get('APPDATA', '')
        settings_dir = os.path.join(app_data, 'T3LabAI')
        if not os.path.exists(settings_dir):
            os.makedirs(settings_dir)
        return os.path.join(settings_dir, 'settings.json')

    def _load_settings(self):
        """Load settings from file, distinguishing the three failure modes.

        This used to `except Exception: pass` straight into the defaults, which
        made a truncated settings.json catastrophic: every setter reloads before
        saving, so ONE toggle in LLMs Setting rewrote the defaults over the file
        and silently destroyed every API key, model preference and knowledge dir.

        Now:
          * file absent (first run) → defaults, writes allowed
          * file parses            → parsed, writes allowed
          * file exists but is not valid JSON → quarantine it as
            settings.corrupt-<stamp>.json, then behave like "absent". The bytes
            survive, so keys can be recovered by hand, and the app stays usable.
          * file exists but cannot be read (locked by the other Revit session,
            permissions) → do NOT quarantine (the file is probably fine).
            Serve defaults for READS only and make save_settings() refuse, so a
            transient lock can never clobber good data. It self-heals on the
            next attempt because every setter reloads first.
        """
        self._load_ok    = True
        self._load_error = None

        if not os.path.exists(self._settings_file):
            return self._get_default_settings()

        try:
            with io.open(self._settings_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except ValueError as ex:
            # Malformed / truncated JSON. json.JSONDecodeError subclasses
            # ValueError on py3 and does not exist on py2 — ValueError is the
            # portable catch for both runtimes.
            self._quarantine_corrupt(ex)
            return self._get_default_settings()
        except Exception as ex:
            self._load_ok    = False
            self._load_error = u"{}".format(ex)
            return self._get_default_settings()

    def _quarantine_corrupt(self, reason):
        """Move an unparseable settings.json aside so it is never overwritten."""
        try:
            import datetime
            stamp = datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
            dest = os.path.join(
                os.path.dirname(self._settings_file),
                'settings.corrupt-{}.json'.format(stamp))
            os.rename(self._settings_file, dest)
            self._load_error = u"settings.json was not valid JSON ({}); " \
                               u"kept a copy at {}".format(reason, dest)
        except Exception as ex:
            # Could not move it — refuse to write rather than destroy it.
            self._load_ok    = False
            self._load_error = u"settings.json is corrupt and could not be " \
                               u"quarantined: {}".format(ex)

    def is_healthy(self):
        """False when the settings file exists but could not be read."""
        return bool(self._load_ok)


    def _update(self, mutator):
        """Reload from disk → apply `mutator(settings)` → save.

        The single write path for every setter. Reloading first is what stops
        two Revit sessions from overwriting each other: without it, a session
        holding a stale in-memory dict wipes whatever the other one saved.
        """
        self._settings = self._load_settings()
        mutator(self._settings)
        return self.save_settings()

    def _get_default_settings(self):
        """Get default settings"""
        return {
            'server': {
                'port': 8080,
                'host': 'localhost'
            },
            'providers': [],
            'api_keys': {},
            # Default to the local Ollama+Qwen runtime (privacy + zero API cost).
            # This is only the FIRST-RUN seed — a saved choice always wins on
            # restore. A local provider never falls back to a cloud one
            # (LLMRouter.get_fallback_candidates).
            'active_provider': DEFAULT_PROVIDER,
            'model_preferences': {},
            # Per-role model pins, validated against the provider's live
            # list: {"claude": {"fast": "", "agent": "", "quality": ""}}.
            # Empty = the provider's preference order (newest Haiku for
            # classification, newest Sonnet for agent turns, newest Opus for
            # quality mode / the teacher).
            'model_roles': {},
            # Server endpoint of a provider the user points at themselves —
            # today only "custom" (any OpenAI-compatible server). NOT secret,
            # so plain text; its API key lives under api_keys["Custom"]
            # (DPAPI) like every other key. See get/set_provider_endpoint.
            'endpoints': {
                'custom': {'base_url': '', 'model': '', 'azure': False},
            },
            'username': 'Thạnh',
            'window_state': {
                'left':         None,
                'top':          None,
                'width':        560,
                'height':       720,
                'sidebar_open': False,
            },
            'active_project': None,
            # Personal instructions — see MAX_USER_INSTRUCTIONS_CHARS.
            'profile': {
                'instructions': '',
            },
            'knowledge': {
                'dirs':               [],
                'embeddings_enabled': True,
                'embed_model':        'nomic-embed-text',
            },
            'agents': {
                'multi_agent':    True,
                'llm_classify':   True,
                'quality_mode':   False,
                # Claude: Haiku pre-reads each turn and picks Haiku / Sonnet /
                # Opus for it (Intelligence/model_tier.py).
                'auto_model':     True,
                'graph_mode':     True,
                'show_tool_calls': True,
                'show_thinking':   True,
                'self_study':      True,
                'parallel_tasks':  True,
                'opus_teacher':    False,
                'self_train_auto': False,
                # Claude: server-side tool search + defer_loading (the ~90
                # non-essential tool schemas leave the prompt until searched).
                'tool_search':     True,
                # USD per day across turns before the composer chip warns.
                # 0 = no warning. Estimates only (Intelligence/pricing.py).
                'daily_budget_usd': 0,
                # Fold tool results older than the last 6 within one turn
                # (Claude: server-side context editing; others: client stub).
                'compact_tool_results': True,
            },
            'skills': {
                'disabled': [],
            },
            'ai_mode': {
                'enabled': True,
                'fast_provider': 'auto',
                'reasoning_provider': 'auto',
                'tool_toggles': {},
            },
        }

    def get_window_state(self):
        """Return the last-saved window state dict with safe boundary clamping."""
        defaults = {'left': None, 'top': None,
                    'width': 560, 'height': 720, 'sidebar_open': False}
        saved = self._settings.get('window_state', {})
        defaults.update(saved)
        try:
            w = defaults.get('width')
            if w is not None:
                w_f = float(w)
                if w_f < 380 or w_f > 1600:
                    defaults['width'] = 560
            h = defaults.get('height')
            if h is not None:
                h_f = float(h)
                if h_f < 400 or h_f > 1400:
                    defaults['height'] = 720
        except Exception:
            defaults['width'] = 560
            defaults['height'] = 720
        return defaults

    def save_window_state(self, left, top, width, height, sidebar_open=False):
        """Persist window geometry and sidebar visibility with boundary clamping."""
        def _m(s):
            w = 560
            h = 720
            try:
                if width is not None:
                    w = max(380, min(int(float(width)), 1600))
                if height is not None:
                    h = max(400, min(int(float(height)), 1400))
            except Exception:
                w, h = 560, 720

            s['window_state'] = {
                'left':         left,
                'top':          top,
                'width':        w,
                'height':       h,
                'sidebar_open': sidebar_open,
            }
        return self._update(_m)

    def save_settings(self):
        """Save settings to file.

        Uses ensure_ascii=True + io.open(utf-8) so non-ASCII values
        (Vietnamese usernames, unicode paths) never break the dump
        under IronPython 2.7.

        Refuses when the last load failed on an existing file — writing then
        would replace real data with defaults.
        """
        if not self._load_ok:
            return False
        try:
            payload = json.dumps(self._settings, indent=2, ensure_ascii=True)
            if isinstance(payload, bytes):
                payload = payload.decode('ascii')
            with io.open(self._settings_file, 'w', encoding='utf-8') as f:
                f.write(payload)
            return True
        except Exception:
            return False


    def get_api_key(self, provider_name):
        """API key for a provider, or None. Read only: never writes the file.

        No saved key → the provider's env var (API_KEY_ENV_VARS), if set. A
        DPAPI blob this Windows account cannot open reads as None
        (api_key_source() tells the UI why).
        """
        stored = self._settings.get('api_keys', {}).get(provider_name)
        if not stored:
            return self._env_api_key(provider_name) or stored
        if secrets.is_protected(stored):
            return secrets.unprotect(stored)
        return stored

    @staticmethod
    def _env_api_key(provider_name):
        var = API_KEY_ENV_VARS.get(provider_name)
        return (os.environ.get(var) or u'').strip() or None if var else None

    def api_key_source(self, provider_name):
        """'saved' | 'env' | 'unreadable' (DPAPI blob from another account or
        PC) | None — what get_api_key() would return a key from."""
        stored = self._settings.get('api_keys', {}).get(provider_name)
        if stored:
            if secrets.is_protected(stored) and secrets.unprotect(stored) is None:
                return 'unreadable'
            return 'saved'
        return 'env' if self._env_api_key(provider_name) else None

    def migrate_plaintext_keys(self):
        """Encrypt plaintext keys left by a pre-DPAPI build, once per session
        (not in the getter: a read must never rewrite settings.json)."""
        if not self._load_ok or not secrets.available():
            return False
        keys = self._settings.get('api_keys', {}) or {}
        plain = [k for k, v in keys.items() if v and not secrets.is_protected(v)]
        if not plain:
            return False

        def _m(s):
            for k in plain:
                v = s.setdefault('api_keys', {}).get(k)
                if v and not secrets.is_protected(v):
                    s['api_keys'][k] = secrets.protect(v)
        return self._update(_m)

    def set_api_key(self, provider_name, api_key):
        """Set API key for a provider.

        Reloads the file from disk first so that keys saved by other sessions
        (or other providers) are not accidentally overwritten by stale
        in-memory data.
        """
        def _m(s):
            s.setdefault('api_keys', {})[provider_name] = secrets.protect(api_key)
        return self._update(_m)

    def get_active_provider(self):
        """Name of the last-selected LLM provider (DEFAULT_PROVIDER if unset)."""
        return self._settings.get('active_provider') or DEFAULT_PROVIDER

    def set_active_provider(self, name):
        """Persist the active provider name."""
        def _m(s):
            s['active_provider'] = name
        return self._update(_m)

    def get_role_model(self, provider_name, role):
        """Pinned model id for a role ('fast' | 'agent' | 'quality') of a
        provider, or '' when the provider's preference order decides."""
        roles = self._settings.get('model_roles', {}) or {}
        prov = roles.get(provider_name) or {}
        return u'{}'.format(prov.get(role) or u'') if isinstance(prov, dict) else u''

    def set_role_model(self, provider_name, role, model_name):
        def _m(s):
            roles = s.setdefault('model_roles', {})
            roles.setdefault(provider_name, {})[role] = model_name or ''
        return self._update(_m)

    def get_daily_budget_usd(self):
        try:
            return float(self._settings.get('agents', {}).get('daily_budget_usd', 0) or 0)
        except Exception:
            return 0.0

    def get_provider_model(self, provider_name):
        """Return the saved model name for a provider, or None."""
        return self._settings.get('model_preferences', {}).get(provider_name)

    def set_provider_model(self, provider_name, model_name):
        """Persist the preferred model name for a provider."""
        def _m(s):
            s.setdefault('model_preferences', {})[provider_name] = model_name
        return self._update(_m)

    def get_provider_endpoint(self, provider_name):
        """Endpoint config of a user-configured provider, always complete:
        {"base_url": str, "model": str, "azure": bool}. A file written before
        this existed has no "endpoints" block and reads as empty values."""
        out = {'base_url': u'', 'model': u'', 'azure': False}
        block = (self._settings.get('endpoints') or {}).get(provider_name)
        if isinstance(block, dict):
            out['base_url'] = u'{}'.format(block.get('base_url') or u'').strip()
            out['model'] = u'{}'.format(block.get('model') or u'').strip()
            out['azure'] = bool(block.get('azure', False))
        return out

    def set_provider_endpoint(self, provider_name, base_url=None, model=None,
                              azure=None):
        """Persist any of base_url / model / azure (None = leave as is)."""
        def _m(s):
            eps = s.get('endpoints')
            if not isinstance(eps, dict):
                eps = {}
                s['endpoints'] = eps
            block = eps.get(provider_name)
            if not isinstance(block, dict):
                block = {'base_url': '', 'model': '', 'azure': False}
                eps[provider_name] = block
            if base_url is not None:
                block['base_url'] = u'{}'.format(base_url).strip()
            if model is not None:
                block['model'] = u'{}'.format(model).strip()
            if azure is not None:
                block['azure'] = bool(azure)
        return self._update(_m)

    def get_username(self):
        """Return the saved user name, or default 'Thạnh'."""
        return self._settings.get('username', 'Thạnh')

    def set_username(self, username):
        """Persist the user name."""
        def _m(s):
            s['username'] = username
        return self._update(_m)

    # ------------------------------------------------------------------
    # Knowledge directories (RAG sources)
    # ------------------------------------------------------------------

    def get_knowledge_dirs(self):
        """Return the list of user-added knowledge directories."""
        return list(self._settings.get('knowledge', {}).get('dirs', []))

    def add_knowledge_dir(self, path):
        """Add a knowledge directory (reload-merge-save, like set_api_key)."""
        def _m(s):
            dirs = s.setdefault('knowledge', {}).setdefault('dirs', [])
            if path not in dirs:
                dirs.append(path)
        return self._update(_m)

    def remove_knowledge_dir(self, path):
        """Remove a knowledge directory."""
        def _m(s):
            dirs = s.setdefault('knowledge', {}).setdefault('dirs', [])
            if path in dirs:
                dirs.remove(path)
        return self._update(_m)

    def get_knowledge_option(self, key, default=None):
        """Read a scalar option from the knowledge block."""
        return self._settings.get('knowledge', {}).get(key, default)

    def set_knowledge_option(self, key, value):
        """Persist a scalar option in the knowledge block."""
        def _m(s):
            s.setdefault('knowledge', {})[key] = value
        return self._update(_m)

    # ------------------------------------------------------------------
    # Multi-agent switches
    # ------------------------------------------------------------------

    def is_multi_agent_enabled(self):
        """Kill switch for the specialist dispatcher (default on)."""
        return bool(self._settings.get('agents', {}).get('multi_agent', True))

    def is_llm_classify_enabled(self):
        """Whether the dispatcher may use one small LLM call to classify."""
        return bool(self._settings.get('agents', {}).get('llm_classify', True))

    def is_graph_mode_enabled(self):
        """Kill switch for the graph agent layer (default on).

        Only affects messages the PLANNER finds more than one goal in — a
        single-goal turn never reaches the graph layer at all, so turning this
        off changes nothing for the overwhelming majority of chats. It is here
        because a multi-step turn runs several agent turns back to back, and a
        user who wants one answer per message should be able to say so.
        """
        return bool(self._settings.get('agents', {}).get('graph_mode', True))

    def is_quality_mode_enabled(self):
        """Opus-parity switch (default off — trades cost/latency for depth).

        When on, the Claude provider prefers the most capable model for the
        default (Opus > Sonnet > Haiku), turns extended thinking on for agent
        turns, and raises the agentic token ceiling. Tiny utility calls
        (classification) pin a fast model via model_override and are unaffected.
        """
        return bool(self._settings.get('agents', {}).get('quality_mode', False))

    def is_auto_model_enabled(self):
        """Auto model switch (default on, Claude only).

        Each turn is graded by a fast triage call and runs on Haiku (simple),
        Sonnet (normal) or Opus (complex) instead of one fixed model. Maximum
        quality, when on, takes precedence: every turn runs on the top model.
        """
        return bool(self._settings.get('agents', {}).get('auto_model', True))

    # ------------------------------------------------------------------
    # Chat verbosity
    #
    # Both default ON: the assistant showing its work live is what stops a
    # long turn from looking like a hang. OFF does NOT go silent — the cards
    # collapse into one clickable summary row and the running tool still
    # shows in the typing indicator. Hiding progress entirely would trade one
    # confusion for a worse one.
    # ------------------------------------------------------------------

    def is_show_tool_calls_enabled(self):
        """Full tool cards in the chat, vs one collapsed summary row."""
        return bool(self._settings.get('agents', {}).get(
            'show_tool_calls', True))

    def set_show_tool_calls(self, enabled):
        return self.set_agent_option('show_tool_calls', bool(enabled))

    def is_show_thinking_enabled(self):
        """Interim narration between tool calls ("Đang tạo dim cho...").

        Only the INTERIM turns are affected. The final answer of a turn is
        never suppressed — a turn that renders nothing at all is the bug this
        setting exists to avoid, not a mode it should offer.
        """
        return bool(self._settings.get('agents', {}).get(
            'show_thinking', True))

    def set_show_thinking(self, enabled):
        return self.set_agent_option('show_thinking', bool(enabled))

    def is_sync_with_central_allowed(self):
        """Whether the assistant may synchronise with central (default OFF).

        Deliberately opt-in: a sync publishes the user's work to everyone on
        the project, can run for minutes, and relinquishes worksets. Every
        other model edit is local and undoable; this one is neither, so it
        needs a decision the user made once, on purpose, outside the chat.
        """
        return bool(self._settings.get('agents', {}).get(
            'allow_sync_with_central', False))


    def is_tool_search_enabled(self):
        """Claude only: send the registry with `defer_loading` + the
        server-side BM25 tool search instead of loading ~98 schemas up front.
        Default on; the provider falls back to the plain list by itself when
        the API rejects the feature."""
        return bool(self._settings.get('agents', {}).get('tool_search', True))

    def set_tool_search_enabled(self, enabled):
        return self.set_agent_option('tool_search', bool(enabled))

    def is_compact_tool_results_enabled(self):
        """Within one agent turn, old tool results are folded away so a
        20-iteration turn does not carry twenty element dumps. Claude does
        it server-side (context_management clear_tool_uses, append-only, so
        preserved thinking is untouched); other providers get a client-side
        stub. Default on."""
        return bool(self._settings.get('agents', {}).get('compact_tool_results', True))

    def get_action_mode(self):
        """Harness action mode for model-editing tools.

        'auto'    = agent executes edits immediately (legacy behavior).
        'confirm' = agent must reply with a plan and wait for the user's OK
                    before any model-modifying tool call.
        """
        mode = self._settings.get('agents', {}).get('action_mode', 'auto')
        return mode if mode in ('auto', 'confirm') else 'auto'

    def get_reply_language(self):
        """Language the assistant answers in.

        'auto' = follow the language the user wrote in (default), 'vi' and
        'en' pin it. Read by the Assistant for its own strings and threaded
        into the LLM system prompts so the model agrees with the UI.
        """
        lang = self._settings.get('agents', {}).get('reply_language', 'auto')
        return lang if lang in REPLY_LANGUAGES else 'auto'

    def set_reply_language(self, lang):
        """Persist the reply-language preference ('auto' | 'vi' | 'en')."""
        if lang not in REPLY_LANGUAGES:
            lang = 'auto'
        return self.set_agent_option('reply_language', lang)

    # ------------------------------------------------------------------
    # Personal instructions (user level, every chat + every project)
    # ------------------------------------------------------------------

    def get_user_instructions(self):
        """The user's personal instructions ('' when none).

        Stored as settings.json → profile.instructions. A file written before
        this existed has no "profile" block and reads as ''. Always returns
        cleaned text within MAX_USER_INSTRUCTIONS_CHARS, even if the file was
        hand-edited past the cap.
        """
        prof = self._settings.get('profile')
        if not isinstance(prof, dict):
            return u''
        return _clean_instructions(prof.get('instructions'))

    def set_user_instructions(self, text):
        """Persist personal instructions (cleaned + capped). Returns bool.

        Text over MAX_USER_INSTRUCTIONS_CHARS is clipped rather than refused,
        so a save never silently loses the whole edit; the settings window
        caps the box at the same length, so clipping only happens to callers
        that skip the UI.
        """
        value = _clean_instructions(text)

        def _m(s):
            prof = s.get('profile')
            if not isinstance(prof, dict):
                prof = {}
                s['profile'] = prof
            prof['instructions'] = value
        return self._update(_m)

    def build_user_instructions_block(self):
        """System-prompt section for the personal instructions ('' when none).

        One shared wording, so every prompt path (native tool calling, legacy
        JSON intent, specialists) frames the block the same way. Project
        instructions come AFTER this block and refine it for that project.
        """
        text = self.get_user_instructions()
        if not text:
            return u''
        return (u'## Personal instructions\n'
                u"(The user's own standing preferences. They apply to every "
                u'chat and every project; project instructions, when present, '
                u'refine them for that project. A request made in the current '
                u'message overrides both.)\n' + text)

    def get_agent_option(self, key, default=None):
        """Read a scalar switch from the agents block."""
        return self._settings.get('agents', {}).get(key, default)

    def set_agent_option(self, key, value):
        """Persist a switch in the agents block."""
        def _m(s):
            s.setdefault('agents', {})[key] = value
        return self._update(_m)

    # ------------------------------------------------------------------
    # Skills
    # ------------------------------------------------------------------

    def get_disabled_skills(self):
        """Return the list of skill ids the user switched off."""
        return list(self._settings.get('skills', {}).get('disabled', []))

    def set_skill_disabled(self, skill_id, disabled):
        """Toggle a skill on/off (persisted globally)."""
        def _m(s):
            items = s.setdefault('skills', {}).setdefault('disabled', [])
            if disabled and skill_id not in items:
                items.append(skill_id)
            elif not disabled and skill_id in items:
                items.remove(skill_id)
        return self._update(_m)

    # ------------------------------------------------------------------
    # Active project
    # ------------------------------------------------------------------

    def get_active_project(self):
        """Return the active project id, or None."""
        return self._settings.get('active_project')

    def set_active_project(self, project_id):
        """Persist the active project id (None = no project)."""
        def _m(s):
            s['active_project'] = project_id
        return self._update(_m)

    # ------------------------------------------------------------------
    # AI Mode (Tool Copilot & Augmented Features)
    # ------------------------------------------------------------------

    def is_ai_mode_enabled(self):
        """Return True if global AI Mode is enabled."""
        return bool(self._settings.get('ai_mode', {}).get('enabled', True))

    def set_ai_mode_enabled(self, enabled):
        """Enable or disable global AI Mode."""
        def _m(s):
            s.setdefault('ai_mode', {})['enabled'] = bool(enabled)
        return self._update(_m)

    def is_tool_ai_enabled(self, tool_name=None):
        """Return True if AI mode is enabled for a specific tool.
        If tool not explicitly configured, defaults to the global AI Mode state.
        """
        if not self.is_ai_mode_enabled():
            return False
        if not tool_name:
            return True
        toggles = self._settings.get('ai_mode', {}).get('tool_toggles', {})
        return bool(toggles.get(tool_name, True))


def get_settings():
    """Get the singleton settings instance"""
    return T3LabAISettings()
