# -*- coding: utf-8 -*-
"""T3Lab Assistant Dialog — Multimodal AI Assistant for Autodesk Revit.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__  = "Tran Tien Thanh"
__title__   = "T3Lab Assistant Dialog"

# IMPORT LIBRARIES
import io
import os
import sys
# ─── CPython 3 & lib bootstrap (engine paths: lib/_cpython_bootstrap.py) ───

_cur = os.path.dirname(os.path.abspath(__file__))
while _cur and not os.path.exists(os.path.join(_cur, 'lib')):
    _parent = os.path.dirname(_cur)
    if _parent == _cur:
        break
    _cur = _parent
_lib_dir = os.path.join(_cur, 'lib')
if os.path.exists(_lib_dir) and _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

try:
    import _cpython_bootstrap
    _cpython_bootstrap.init_cpython_paths()
except Exception:
    pass
# ──────────────────────────────────────────────────────────────────────────────
import clr
import json
import re
import datetime

clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
import System.Windows
from System.Windows import Visibility, WindowState, GridLength
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind, Action
from System.Threading import Thread, ThreadStart, ApartmentState

from pyrevit import revit, forms, script
from GUI.WPF_Base import T3WPFWindow
from Autodesk.Revit import DB

# DEFINE VARIABLES
logger = script.get_logger()
# NO script.get_output(): it materialises an output window; from this persistent
# engine's worker threads, pyRevit Reload closing it throws
# InvalidOperationException (cross-thread WPF window) and aborts Reload.

try:
    REVIT_VERSION = int(revit.doc.Application.VersionNumber)
except Exception:
    try:
        from pyrevit import HOST_APP
        REVIT_VERSION = int(HOST_APP.version)
    except Exception:
        REVIT_VERSION = 2023  # safe fallback

# ─── Lib path setup ───────────────────────────────────────────────────────────
_gui_dir = os.path.dirname(__file__)
_lib_dir = os.path.dirname(_gui_dir)
extension_dir = os.path.dirname(_lib_dir)
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.insert(0, lib_dir)

# ─── Theme: follow Revit's own UI theme (GUI/assistant/theme.py) ────────────
# All chat-surface colours come from there (no hex at call sites) so the
# window re-skins on Revit Light/Dark (2024+).
from GUI.assistant.theme import (_theme, _tb, _trgb, _theme_color, _bind_theme,
                                 _bind_fg, _bind_bg, _bind_border,
                                 _bind_stroke)
try:
    from GUI.assistant.theme import _LIGHT_FALLBACK   # only without RevitTheme
except ImportError:
    pass


# ─── Shared with the LLMs Setting dialog ──────────────────────────────────────
# One copy of the openers: private ones lacked the .NET 8 UseShellExecute
# fallback and died silently on Revit 2025+.
try:
    from GUI.AssistantShared import (PROVIDER_COLORS as _SHARED_PROVIDER_COLORS,
                                     PROVIDER_GRAY as _SHARED_PROVIDER_GRAY,
                                     open_in_explorer as _open_in_explorer,
                                     open_file_with_shell as _open_file_shell,
                                     project_scope_lines as _project_scope_lines)
except Exception:
    _SHARED_PROVIDER_COLORS = {
        "claude": (217, 119, 87), "openai": (16, 163, 127),
        "deepseek": (37, 99, 235), "ollama": (59, 130, 246),
        "lmstudio": (124, 58, 237), "gemini": (66, 133, 244),
        "custom": (100, 116, 139),
    }
    _SHARED_PROVIDER_GRAY = (161, 161, 170)

    def _open_in_explorer(path, create=True):
        return False

    def _open_file_shell(path):
        return False

    def _project_scope_lines(meta, doc_counts=None):
        return []

# Pure Python: a failure here is a broken install, so let it surface.
from GUI import AssistantGreetings as _greetings

# ─── Split-out module-level helpers (GUI/assistant/) ─────────────────────────
# Optional Intelligence/Services imports + stubs, text / language helpers,
# tool launchers, chat-history persistence, icon & plan constants. Every name
# is re-imported so `from GUI.T3LabAssistantDialog import X` keeps working.
from GUI.assistant.deps import (learn_pattern, find_learned_match,
                                has_local_llm, parse_command_local,
                                get_local_model_name, parse_command_nlu,
                                inject_discovered_tools,
                                get_active_provider_name,
                                get_provider_display_label,
                                get_setup_guidance_message,
                                _build_system_prompt, _RAG_SYSTEM_PREFIX,
                                HAS_NLP, discover_new_tools,
                                get_registered_tools, make_generic_launcher,
                                make_launcher_for, HAS_DISCOVERY,
                                ensure_api_context, run_in_api_context,
                                HAS_API_CONTEXT, ContextScout, HAS_SCOUT,
                                HAS_EXECUTOR, is_supported, is_image, is_pdf,
                                build_text_context,
                                build_vision_content_blocks, has_images,
                                summarize_attachments, SUPPORTED_EXTS, HAS_RAG,
                                link_reader, HAS_LINKS, get_global_store,
                                get_active_store, HAS_KNOWLEDGE, routing,
                                _lang_analyzer, _lang_detect, HAS_LANGUAGE,
                                GraphOrchestrator, HAS_GRAPH, AgentDispatcher,
                                HAS_AGENTS, _announces_work, _ANNOUNCE_FIXUP,
                                get_spec, build_specialist_prompt,
                                HAS_SPECIALISTS, telemetry, HAS_TELEMETRY,
                                StreamingJSONExtractor, _provider_usable,
                                _get_uiapp)
# Defined only when their optional import worked (NLP layer / BatchOut
# executor), exactly as before the split.
try:
    from GUI.assistant.deps import parse_command, has_api_key, keyword_parse
except ImportError:
    pass
try:
    from GUI.assistant.deps import configure_batchout_window, direct_export
except ImportError:
    pass
from GUI.assistant.text import (_THINK_BLOCK_RE, _THINK_OPEN_RE,
                                _hide_reasoning, _exc_text, _json_safe,
                                _json_text, _call_summary)
from GUI.assistant.lang import (_VIET_CHARS, _reply_language, _is_viet_text,
                                _detect_viet, _note_user_language,
                                _relative_time, _time_greeting, _ui_viet)
from GUI.assistant.launchers import (_get_tool_script_dir, _load_script,
                                     _load_batchout_mod, launch_batchout,
                                     launch_batchout_configured,
                                     launch_loadfamily, _SPECIAL_LAUNCHERS,
                                     TOOL_LAUNCHERS, TOOL_TITLES,
                                     DROPPED_TOOLS,
                                     _register_discovered_launchers,
                                     get_tool_title)
from GUI.assistant.history import (LEGACY_DOC_KEY, _LEGACY_CHAT_HISTORY_DIR,
                                   MAX_ARCHIVED_SESSIONS, _chat_history_dir,
                                   _prune_archived_sessions, _active_doc,
                                   _get_doc_key, _history_file,
                                   save_chat_history, _is_stale_tool_call_blob,
                                   _history_path_for_read, load_chat_history,
                                   load_chat_summary, load_chat_session,
                                   clear_chat_history, MAX_REPLAY_BUBBLES,
                                   _active_budgets, _budget_history,
                                   _fit_restored)
from GUI.assistant.constants import (_ICON_INFO, _ICON_SEARCH, _ICON_WARNING,
                                     _ICON_SUCCESS, _ICON_SYNC, _ICON_STOP,
                                     _ICON_REFRESH, _ICON_ATTACH,
                                     _ICON_ANALYZE, _ICON_LIST, _ICON_BLUE,
                                     _ICON_AMBER, _ICON_GREEN, _ICON_RED,
                                     _ICON_SLATE, _ICON_PENDING, _ICON_FAILED,
                                     _ICON_EDIT, _PLAN_RUN, _PLAN_EDIT,
                                     _PLAN_CANCEL, _PLAN_SUPERSEDED,
                                     _PLAN_APPROVAL_TIMEOUT, _CONFIRM_TIMEOUT,
                                     _STEP_LOOK, _MEM_SAVE_RX)
from GUI.assistant import lang as _lang


def __getattr__(name):
    """Live values for module state that moved out (PEP 562).

    `_LAST_USER_VIET` is rebound inside GUI.assistant.lang on every turn; a
    `from ... import` copy here would freeze at None.
    """
    if name == '_LAST_USER_VIET':
        return _lang._LAST_USER_VIET
    raise AttributeError(
        "module {!r} has no attribute {!r}".format(__name__, name))


# Window helpers: groups of T3LabAssistantWindow methods, by composition
# (a CLR-derived window cannot take a Python mixin — CLAUDE.md rule 6).
from GUI.assistant.sessions import AssistantSessions
from GUI.assistant.projects import AssistantProjects
from GUI.assistant.commands import AssistantCommands
from GUI.assistant.documents import AssistantDocuments
from GUI.assistant.plans import AssistantPlans
from GUI.assistant.spellcheck import AssistantSpellcheck
from GUI.assistant.tool_cards import AssistantToolCards


# CLASS/FUNCTIONS

class T3LabAssistantWindow(T3WPFWindow):
    """Standalone T3Lab Assistant chatbox window."""

    _DYNAMIC_BTNS = []   # Button objects added by _bootstrap_discovered_tools

    def __init__(self, is_docked=False):
        self.is_docked = is_docked
        # Method groups split out of this class (GUI/assistant/): plain
        # classes holding a back-reference, built FIRST so XAML handlers
        # and the setup below can reach them. All state stays on self.
        self._sessions = AssistantSessions(self)
        self._projects = AssistantProjects(self)
        self._commands = AssistantCommands(self)
        self._docs = AssistantDocuments(self)
        self._plans = AssistantPlans(self)
        self._spell = AssistantSpellcheck(self)
        self._tool_cards = AssistantToolCards(self)
        try:
            xaml_path = os.path.join(_gui_dir, 'Tools', 'T3LabAssistant.xaml')
            T3WPFWindow.__init__(self, xaml_path)
        except Exception as ex:
            logger.error(u"Could not load T3LabAssistant XAML: {}".format(_exc_text(ex)))
            raise

        # Skin BEFORE anything renders, else one frame of the wrong theme flashes.
        try:
            self._apply_revit_skin()
        except Exception as ex:
            logger.debug(u"_apply_revit_skin error: {}".format(_exc_text(ex)))

        self.doc = revit.doc

        # ── Session state ─────────────────────────────────────────────────────
        self._busy             = False          # concurrency guard
        self._switching_provider = False        # guard: _switch_provider bg probe in flight
        self._warming_up       = False          # guard: local-model warm-up in flight
        self._warmed_provider  = None           # last local provider we preloaded
        self._warmed_at        = 0.0            # timestamp of that warm-up
        self._snapshot_building = False         # guard: model-snapshot build in flight
        self._snapshot_built_for = None         # doc_key whose snapshot we ensured
        self._typing_row       = None           # reference to typing indicator element
        self._conversation_history = []         # [{role, content}, ...] multi-turn context
        self._history_summary  = u''            # turns folded out of the window
        self._last_raw         = ''             # last user input (for learning)
        self._last_result      = None           # last parsed result (for votes)
        self._doc_key          = _get_doc_key() # document identifier for history
        self._persisted_msgs   = []             # flat list with timestamps, for save/load
        # This chat's identity + its own memory (Intelligence.session_context).
        self._session_id       = None           # stable across restore / resume
        self._session_path     = None           # its archive, when resumed from History
        self._session_notes    = []             # facts for THIS conversation only
        self._past_chats_cache = None           # (key, block) — digests of earlier chats
        self._attached_files   = []             # list of file paths (images / PDFs)
        # {provider_name: [model_list]} — ONLY from live probes, never hardcoded;
        # the MODEL combo stays disabled until the provider reports its models.
        self._models_cache     = {}
        
        # ── History & Typing Animation State ──────────────────────────────────
        self._input_history    = []
        self._history_index    = -1
        self._current_input_temp = ""
        self._typing_timer     = None
        self._typing_elapsed   = 0

        # ── Live streaming bubble state ───────────────────────────────────────
        self._stream_row       = None           # Grid row of the live reply
        self._stream_tb        = None           # TextBlock being filled token-by-token
        self._stream_host      = None           # StackPanel holding body + action row

        # ── Claude-style message actions ──────────────────────────────────────
        # During "Try again": the answer becomes another version of that row
        # and the user message is not echoed again.
        self._retry_target_row   = None
        self._suppress_user_echo = False

        # ── Native agent loop state ───────────────────────────────────────────
        self._agent_loop        = None    # running AgentLoop (native tools path)
        self._cancel_requested  = False   # Stop pressed before the loop existed

        # ── Turn bookkeeping ──────────────────────────────────────────────────
        # _request_id is bumped by _set_busy(True); terminal handlers no-op on a
        # mismatch, so a late worker and the Stop watchdog never BOTH release
        # the UI (that drained two queued messages).
        self._request_id  = 0
        self._replied     = False   # a terminal message was already posted
        self._tool_runs   = 0       # tools executed in the current turn

        # ── Easy-to-use input: multi-line with Shift+Enter ────────────────────
        try:
            from System.Windows.Controls import ScrollBarVisibility
            from System.Windows import TextWrapping
            self.chat_input.AcceptsReturn = False
            self.chat_input.TextWrapping  = TextWrapping.Wrap
            self.chat_input.MaxHeight     = 120
            self.chat_input.VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            self.chat_input.ToolTip = (u"Type your request (English or Vietnamese)  •  "
                                       u"Enter to send, Shift+Enter for a new line")
        except Exception:
            pass

        # ── Claude-style composer state (slash-skills popup + chips) ──────────
        self._slash_open      = False   # skills popup visible?
        self._slash_items     = []      # filtered catalog metas shown in popup
        self._slash_rows      = []      # Border rows (for highlight swap)
        self._slash_sel       = 0       # highlighted index
        self._forced_skill_id = None    # skill forced via /slash for THIS message
        self._forced_skill_args = u""   # user text after the /slash id (pre-boilerplate)
        self._tool_group      = None    # collapsed tool-card group for THIS turn

        # Provider before a project override, restored on leaving the project.
        self._pre_project_provider = None

        # ── Usage-flow state: message queue + quick-reply chips ───────────────
        # Input stays enabled while busy: Enter queues, auto-sent on release.
        # One quick-reply chip row at a time, replaced on every send.
        self._queued_inputs  = []      # [{'text': unicode, 'row': element}]
        self._quickreply_row = None    # active chip row (or None)

        # Chips AFTER first render: LLMRouter + project JSON are off the startup path.
        def _init_chips():
            try:
                self._update_composer_chips()
                self._update_action_mode_chip()
            except Exception:
                pass
        try:
            from System.Windows.Threading import DispatcherPriority
            self.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                                        Action(_init_chips))
        except Exception:
            _init_chips()
        # Popups must never linger when focus leaves the input / the window
        try:
            def _input_lost_focus(s, ev):
                self._close_skills_popup()
            self.chat_input.LostKeyboardFocus += _input_lost_focus

            def _win_deactivated(s, ev):
                self._close_skills_popup()
                try:
                    self.project_popup.IsOpen = False
                    self.model_popup.IsOpen = False
                except Exception:
                    pass
            self.Deactivated += _win_deactivated
        except Exception:
            pass

        # Per-project scheduled prompts: 30s UI-thread timer while the window is open.
        try:
            self._projects._start_schedule_timer()
        except Exception:
            pass

        # ── Paste (Ctrl+V) + drag-drop files straight into the chat ──────────
        # Wired in code: pyRevit's XAML wiring of Window-level events can break loading.
        try:
            self.AllowDrop = True
            self.PreviewDragOver += self._file_drag_over
            self.PreviewDrop += self._file_drop
            self.chat_input.PreviewKeyDown += self._input_preview_keydown
        except Exception:
            pass


        # ── Restore conversation from previous session ─────────────────────────
        self._sessions._restore_history()
        self._update_welcome_greeting()

        # ── Tool discovery (background, then inject chips into UI) ─────────────
        def _discover_and_update():
            # Guarded: closing the window during the sleep threw on a .NET thread.
            try:
                import time; time.sleep(0.3)
                self.Dispatcher.Invoke(Action(self._bootstrap_discovered_tools))
            except Exception as ex:
                logger.debug("tool discovery skipped: {}".format(_exc_text(ex)))

        _dt = Thread(ThreadStart(_discover_and_update))
        _dt.IsBackground = True
        _dt.SetApartmentState(ApartmentState.STA)
        _dt.Start()

        # Restore window geometry and sidebar state from last session
        self._restore_window_state()

        # ── Register the MCP ExternalEvent on the UI thread ────────────────────
        # MUST be on Revit's main thread: ExternalEvent.Create throws from the
        # background probe, leaving every model-editing MCP tool without a transaction.
        try:
            from Services.mcp_service import MCPService as _MCPService
            _ok, _ee_err = _MCPService.ensure_external_event()
            if not _ok:
                logger.debug(u"MCP ExternalEvent init failed: {}".format(_exc_text(_ee_err)))
        except Exception as _ex:
            logger.debug(u"MCP ExternalEvent init error: {}".format(_exc_text(_ex)))

        # ── Register the tool-launch ExternalEvent, same constraint ───────────
        # Else launched tools run in the IDLE context and throw "Attempting to
        # create an ExternalEvent outside of a standard API execution".
        try:
            _ok, _ctx_err = ensure_api_context()
            if not _ok:
                logger.debug(u"API context init failed: {}".format(_exc_text(_ctx_err)))
        except Exception as _ex:
            logger.debug(u"API context init error: {}".format(_exc_text(_ex)))

        # Update AI badge, pre-load models cache, and warm up router status in background
        def _bg_startup_probe():
            try:
                import time
                time.sleep(0.5)   # let window render first

                # ─── 1. Auto-start Revit MCP Server & File Watcher ───
                # In-process, runs unattended; a failure is logged to activity.
                try:
                    from Services.mcp_service import MCPService
                    srv_status = MCPService.server_status()
                    if not srv_status.get('running'):
                        MCPService.start_server()
                    wat_status = MCPService.watcher_status()
                    if not wat_status.get('running'):
                        # Opt-in only (MCP Control > Advanced).
                        MCPService.auto_start_watcher()
                except Exception as ex:
                    logger.debug(u"Auto-start MCP/watcher failed: {}".format(_exc_text(ex)))
                    self._log_activity(
                        u"MCP server/watcher did not start: {}".format(
                            _exc_text(ex)))

                # ─── 2. Local engine: never started or downloaded here ─────
                # (it once spawned `ollama serve` + pulled ~1 GB without
                # consent). A running Ollama / LM Studio shows up in the model
                # popup by itself (_refresh_model_popup_async), and onboarding
                # covers "no provider yet".

                # ─── 3.4 Skills registry scan ───
                try:
                    from Intelligence.skills_engine import SkillsEngine
                    _n_skills = SkillsEngine().scan()
                    logger.debug("Skills scanned: {}".format(_n_skills))
                except Exception as ex:
                    logger.debug(u"Skills scan failed: {}".format(_exc_text(ex)))
                    self._log_activity(
                        u"Skills scan failed: {}".format(_exc_text(ex)))

                # ─── 3.5 Knowledge index: incremental scan + vectors ───
                # Incremental; embeddings only when the model is ALREADY
                # installed (the ~270 MB pull is never silent).
                try:
                    if HAS_KNOWLEDGE:
                        _store = get_active_store()
                        if _store is not None:
                            _scan_res = _store.scan()
                            if _scan_res.get('added') or _scan_res.get('updated'):
                                logger.debug(u"Knowledge scan: {}".format(_exc_text(_scan_res)))
                            try:
                                from Intelligence.knowledge.embeddings import (
                                    get_default_embedder)
                                _emb = get_default_embedder()
                                if _emb is not None and _emb.is_available():
                                    _store.embed_pending(_emb, budget_sec=90)
                            except Exception:
                                pass
                            self.Dispatcher.Invoke(
                                Action(self._update_knowledge_status))
                except Exception as ex:
                    logger.debug(u"Knowledge scan failed: {}".format(_exc_text(ex)))
                    self._log_activity(
                        u"Knowledge scan failed: {}".format(_exc_text(ex)))

                from Intelligence.llm_router import LLMRouter
                router = LLMRouter()
                active = router.get_active_name()

                self.Dispatcher.Invoke(Action(self._update_ai_badge))

                # Step 1: Probe the active provider first to fill cache ASAP
                try:
                    router.probe_provider(active)
                    provider = router.get_provider(active)
                    if provider:
                        self._models_cache[active] = provider.get_models()
                except Exception:
                    pass

                # Step 2: Pre-load/cache models for all other providers in the background
                for name in router.get_provider_names():
                    if name == active:
                        continue
                    try:
                        p = router.get_provider(name)
                        if p and p.check_health():
                            self._models_cache[name] = p.get_models()
                    except Exception:
                        pass

                # Step 3: Warm up LLMRouter status cache
                try:
                    router.get_status(use_cache=False)
                except Exception:
                    pass

                # Step 4: setup nudge — fresh chat, onboarding already done, no provider.
                try:
                    from config.user_profile import UserProfile
                    already_onboarded = not UserProfile().is_first_run()
                    fresh_chat = not self._persisted_msgs
                    no_provider = not has_api_key() and not has_local_llm()
                    if already_onboarded and fresh_chat and no_provider:
                        def _nudge():
                            _guide = get_setup_guidance_message(_ui_viet())
                            self._append_bot_message(_guide, icon=_ICON_INFO,
                                                     icon_color=_ICON_SLATE)
                            self._add_to_history("assistant", _guide)
                        self.Dispatcher.Invoke(Action(_nudge))
                except Exception:
                    pass
            except Exception:
                pass

        _t = Thread(ThreadStart(_bg_startup_probe))
        _t.IsBackground = True
        _t.SetApartmentState(ApartmentState.STA)
        _t.Start()

        # ── First-run onboarding (new installs only) ──────────────────────────
        try:
            from config.user_profile import UserProfile
            if UserProfile().is_first_run():
                self._show_onboarding()
        except Exception as ex:
            logger.debug(u"onboarding check error: {}".format(_exc_text(ex)))

        # Live Revit context strip (view + selection); last, so a refused read costs nothing.
        try:
            self._update_revit_context()
            self._start_context_timer()
        except Exception as ex:
            logger.debug(u"revit context strip unavailable: {}".format(_exc_text(ex)))

        # Persist window geometry/sidebar on close
        try:
            self.Closing += self._on_closing
        except Exception:
            pass

    def setup_icon(self):
        """Override pyRevit's setup_icon to remove the window icon from the title bar."""
        pass

    # ═══ REVIT INTEGRATION — chrome, theme, live context (reads as a Revit panel; survives a 320px dock) ═══

    def _apply_revit_skin(self):
        """Make the window sit inside Revit instead of on top of it."""
        self._current_skin = None
        self._sync_theme(force=True)

        if self.is_docked:
            self._flatten_chrome_for_pane()

        # The theme can change while open (Options, BG Theme tool): re-read on activate.
        try:
            self.Activated += self._on_window_activated
        except Exception:
            pass

        try:
            self.SizeChanged += self._on_size_changed
            if hasattr(self, 'root_chrome') and self.root_chrome is not None:
                self.root_chrome.SizeChanged += self._on_size_changed
        except Exception:
            pass
        self._compact = None
        self._narrow_state = None
        try:
            _initial_w = self.root_chrome.ActualWidth if (hasattr(self, 'root_chrome') and self.root_chrome.ActualWidth > 0) else (self.ActualWidth or self.Width)
            self._apply_compact_layout(_initial_w)
            self._apply_narrow_layout(_initial_w)
        except Exception:
            pass

    def _flatten_chrome_for_pane(self):
        """Strip the floating-card chrome (rounded corners showed the host through)."""
        try:
            from System.Windows import CornerRadius, Thickness
            self.root_chrome.CornerRadius    = CornerRadius(0)
            self.root_chrome.BorderThickness = Thickness(0)
        except Exception as ex:
            logger.debug(u"_flatten_chrome_for_pane: {}".format(_exc_text(ex)))

        # Revit provides the pane's own close/undock buttons.
        try:
            self.float_ctrls_panel.Visibility = Visibility.Collapsed
        except Exception:
            pass

    def _theme_scopes(self):
        """Every element that must carry its own copy of the T3Theme* brushes.

        Docked, AssistantPaneControl detaches win.Content from the Window, so
        Window.Resources stops being an ancestor scope: every
        {DynamicResource T3Theme*} resolved to nothing and the pane went
        transparent over Revit's black HwndSource ("pane went dark"). Tokens
        on the content root travel with it; floating, root_chrome just wins.
        """
        scopes = []
        root = getattr(self, 'root_chrome', None)
        if root is not None:
            scopes.append(root)
        scopes.append(self)     # the Window's own Background binds too
        return scopes

    def _sync_theme(self, force=False):
        """Repaint the shell in Revit's current theme (Light / Dark).

        DynamicResource bindings follow immediately; the transcript is not
        rebuilt (a stream may be writing into a row) — see _bind_theme.
        """
        if _theme is None:
            return None
        try:
            theme = _theme.current_theme()
            if theme == getattr(self, '_current_skin', None) and not force:
                return theme
            self._current_skin = theme
            for scope in self._theme_scopes():
                try:
                    _theme.apply(scope, theme)
                except Exception:
                    continue
            return theme
        except Exception as ex:
            logger.debug(u"_sync_theme error: {}".format(_exc_text(ex)))
            return None

    def _on_window_activated(self, sender, e):
        self._sync_theme()
        self._sync_model_chip_if_stale()
        self._update_revit_context()
        self._maybe_warm_up_local()
        self._maybe_refresh_model_snapshot()

    def _maybe_warm_up_local(self, force=False):
        """Preload the active LOCAL model so the first message replies fast.

        Ollama/LM Studio load on first use. No-op for cloud; throttled
        (keep_alive holds ~15m); background thread.
        """
        try:
            import time as _time
            name = get_active_provider_name()
            if name not in ("ollama", "lmstudio"):
                return
            now = _time.time()
            if not force and self._warmed_provider == name \
                    and (now - self._warmed_at) < 600:
                return
            if self._warming_up:
                return
            self._warming_up = True

            def _bg():
                try:
                    from Intelligence.llm_router import LLMRouter
                    provider = LLMRouter().get_active_provider()
                    if provider is not None and hasattr(provider, "warm_up"):
                        provider.warm_up()
                        self._warmed_provider = name
                        self._warmed_at = _time.time()
                except Exception:
                    pass
                finally:
                    self._warming_up = False

            t = Thread(ThreadStart(_bg))
            t.IsBackground = True
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        except Exception:
            self._warming_up = False

    def _model_snapshot_grounding(self):
        """Cached digest of the OPEN model as a short grounding block for the turn
        (saves a local model read round-trips). Empty until a snapshot exists."""
        try:
            from Intelligence.learning.enrichers import model_snapshot
            digest = model_snapshot.load_digest(self._doc_key)
            return model_snapshot.digest_to_grounding_text(digest) or u""
        except Exception:
            return u""

    def _maybe_refresh_model_snapshot(self):
        """Build the open-model digest in the background if this document has
        none. Once per document per session; the enricher marshals to Revit."""
        try:
            key = self._doc_key
            if self._snapshot_built_for == key or self._snapshot_building:
                return
            from Intelligence.learning.enrichers import model_snapshot
            if model_snapshot.load_digest(key).get('element_counts'):
                self._snapshot_built_for = key   # a usable snapshot already exists
                return
            self._snapshot_building = True

            def _bg():
                try:
                    model_snapshot.run(doc_key=key)
                    self._snapshot_built_for = key
                except Exception:
                    pass
                finally:
                    self._snapshot_building = False

            t = Thread(ThreadStart(_bg))
            t.IsBackground = True
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        except Exception:
            self._snapshot_building = False

    # ─── Compact layout for narrow docks ─────────────────────────────────────
    # Widths are DIPs (display scaling already folded in).

    COMPACT_WIDTH = 400

    #: Below this, secondary labels collapse to icons (_apply_narrow_layout).
    NARROW_WIDTH = 340

    #: Floating min / max / close cluster (3 x 28 + 4); collapsed when docked.
    _FLOAT_CTRLS_WIDTH = 88

    _PLACEHOLDER_FULL = u"Type a request — press / for skills"
    _PLACEHOLDER_SHORT = u"Ask, or type / for skills"

    def _on_size_changed(self, sender, e):
        try:
            if e.WidthChanged:
                self._apply_compact_layout(e.NewSize.Width)
                self._apply_narrow_layout(e.NewSize.Width)
                # Always: a resize that never crosses the threshold must still refit.
                self._fit_popups_to_width(e.NewSize.Width)
        except Exception:
            pass

    def _apply_narrow_layout(self, width):
        """Icon-only secondary buttons and a shorter composer hint. UI THREAD.

        Each label collapses on its own threshold; icon and tooltip stay.
        Name trimming is pure XAML.
        """
        try:
            w = float(width or 0)
        except Exception:
            return
        if w <= 0:
            return
        floating = False
        try:
            floating = self.float_ctrls_panel.Visibility == Visibility.Visible
        except Exception:
            pass
        top_narrow = w < self.NARROW_WIDTH + (
            self._FLOAT_CTRLS_WIDTH if floating else 0)
        narrow = w < self.NARROW_WIDTH
        state = (top_narrow, narrow)
        if state == getattr(self, '_narrow_state', None):
            return
        self._narrow_state = state

        def _show(name, visible):
            try:
                el = getattr(self, name, None)
                if el is not None:
                    el.Visibility = (Visibility.Visible if visible
                                     else Visibility.Collapsed)
            except Exception:
                pass

        _show('saved_prompts_label', not top_narrow)
        for name in ('action_mode_text', 'undo_label', 'activity_log_label'):
            _show(name, not narrow)
        try:
            self.composer_placeholder.Text = (
                self._PLACEHOLDER_SHORT if narrow else self._PLACEHOLDER_FULL)
        except Exception:
            pass

    def _apply_compact_layout(self, width):
        """Tighten padding and drop optional footer text on a narrow dock.

        The copyright stays (required); the disclaimer goes.
        """
        try:
            compact = bool(width) and width < self.COMPACT_WIDTH
        except Exception:
            return
        if compact == getattr(self, '_compact', None):
            return
        self._compact = compact

        from System.Windows import Thickness
        try:
            if compact:
                self.composer_bar.Padding = Thickness(9, 6, 9, 8)
                self.chat_scroll.Padding  = Thickness(10, 10, 5, 6)
            else:
                self.composer_bar.Padding = Thickness(16, 8, 16, 10)
                self.chat_scroll.Padding  = Thickness(16, 12, 10, 8)
        except Exception:
            pass
        # First-run card chrome left its two buttons ~200 DIP on a 280 dock.
        try:
            self.onboarding_card.Margin  = Thickness(8 if compact else 16)
            self.onboarding_card.Padding = Thickness(16 if compact else 24)
        except Exception:
            pass

        vis = Visibility.Collapsed if compact else Visibility.Visible
        try:
            self.footer_disclaimer.Visibility = vis
        except Exception:
            pass

    # Docked panes are often narrower than the popups' authored MinWidth.
    _POPUP_PANELS = ('skills_popup_panel', 'project_popup_panel',
                     'model_popup_panel', 'saved_prompts_list_panel')
    _POPUP_GUTTER = 24    # popup border, margin and shadow
    # Widest authored MinWidth (skills_popup); below it popups pin to the pane.
    _POPUP_MIN_AUTHORED = 340

    def _fit_popups_to_width(self, width):
        """Clamp popup content to the pane. UI THREAD. Never raises.

        Pinned on the Border that carries MinWidth/MaxWidth; cleared again
        when the pane is wide enough.
        """
        from System.Windows import FrameworkElement
        from System.Windows.Controls import Border
        try:
            avail = float(width or 0) - self._POPUP_GUTTER
        except Exception:
            return
        if avail <= 0:
            return
        for name in self._POPUP_PANELS:
            try:
                node = getattr(self, name, None)
                # Walk up (past a ScrollViewer) to the Border.
                while node is not None and not isinstance(node, Border):
                    node = getattr(node, 'Parent', None)
                if node is None:
                    continue
                if avail < self._POPUP_MIN_AUTHORED:
                    node.MinWidth = avail
                    node.MaxWidth = avail
                else:
                    node.ClearValue(FrameworkElement.MinWidthProperty)
                    node.ClearValue(FrameworkElement.MaxWidthProperty)
            except Exception:
                continue

    #: Space kept between a popup and the pane's left / right edge.
    _POPUP_EDGE = 8.0

    def _keep_popup_in_pane(self, popup):
        """Shift a popup left so it does not hang past the pane. UI THREAD.

        Placement Top/Bottom aligns LEFT edges, so popups of right-side targets
        overhung the pane. Call right before opening; offsets by the overhang,
        never past the left edge. Never raises.
        """
        try:
            from System.Windows import Point, Size
            root = getattr(self, 'root_chrome', None)
            target = popup.PlacementTarget
            child = popup.Child
            if root is None or target is None or child is None:
                return
            pane_w = float(root.ActualWidth or 0)
            if pane_w <= 0:
                return
            child.Measure(Size(float('inf'), float('inf')))
            pop_w = float(child.DesiredSize.Width or 0)
            left = float(target.TranslatePoint(Point(0, 0), root).X)
            overhang = left + pop_w - (pane_w - self._POPUP_EDGE)
            if overhang <= 0:
                popup.HorizontalOffset = 0.0
            else:
                popup.HorizontalOffset = -min(
                    overhang, max(0.0, left - self._POPUP_EDGE))
        except Exception as ex:
            logger.debug(u"_keep_popup_in_pane: {}".format(_exc_text(ex)))

    # ─── Live Revit context ──────────────────────────────────────────────────

    def _start_context_timer(self):
        """Poll the active view / selection so the strip stays live.

        Reads only, legal outside an API context. Three consecutive failures
        stop the timer (floating window only — see _on_context_tick).
        """
        try:
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan
            self._ctx_failures = 0
            self._ctx_timer = DispatcherTimer()
            self._ctx_timer.Interval = TimeSpan.FromSeconds(2)
            self._ctx_timer.Tick += self._on_context_tick
            self._ctx_timer.Start()
        except Exception as ex:
            logger.debug(u"context timer unavailable: {}".format(_exc_text(ex)))

    def _on_context_tick(self, sender, e):
        # Ask the element on screen: docked, the Window itself is never shown.
        probe = getattr(self, 'root_chrome', None)
        if probe is None:
            probe = self
        try:
            if not probe.IsVisible:
                return          # docked pane hidden — nothing to refresh
        except Exception:
            pass
        # A docked pane raises no Window.Activated; this tick notices theme /
        # provider (LLMs Setting) changes. Both no-op when unchanged.
        self._sync_theme()
        self._sync_model_chip_if_stale()
        if not self._update_revit_context():
            self._ctx_failures = getattr(self, '_ctx_failures', 0) + 1
            # Docked, no open document is normal; stopping would freeze later docs.
            if self._ctx_failures >= 3 and not self.is_docked:
                try:
                    self._ctx_timer.Stop()
                except Exception:
                    pass
        else:
            self._ctx_failures = 0
            self._warm_live_context()

    def _warm_live_context(self):
        """Keep ContextScout's levels/grids digest warm. UI THREAD (tick).

        Cached per document (usually a dict lookup); a rebuild runs here while
        Revit is idle, not on the routing worker mid-turn.
        """
        if not HAS_SCOUT:
            return
        try:
            from Selection import scout as _scout
            uiapp = _get_uiapp()
            uidoc = getattr(uiapp, 'ActiveUIDocument', None) if uiapp else None
            doc = getattr(uidoc, 'Document', None) if uidoc else None
            if doc is not None:
                _scout.get_levels_and_grids(doc)
        except Exception:
            pass

    def _read_revit_context(self):
        """(view_name, selection_count) for the active document, or None."""
        uiapp = _get_uiapp()
        if uiapp is None:
            return None
        uidoc = getattr(uiapp, 'ActiveUIDocument', None)
        if uidoc is None:
            return None
        doc = getattr(uidoc, 'Document', None)
        if doc is None:
            return None
        # The docked pane outlives its document; follow the active one.
        if self.is_docked:
            self.doc = doc
        view_name = u""
        try:
            view = doc.ActiveView
            if view is not None:
                view_name = view.Name
        except Exception:
            pass
        count = 0
        try:
            count = uidoc.Selection.GetElementIds().Count
        except Exception:
            pass
        return (view_name, count)

    def _update_revit_context(self):
        """Refresh the context strip. Returns False if Revit could not be read."""
        try:
            ctx = self._read_revit_context()
        except Exception as ex:
            logger.debug(u"_read_revit_context: {}".format(_exc_text(ex)))
            ctx = None
        if ctx is None:
            try:
                self.revit_ctx_btn.Visibility = Visibility.Collapsed
            except Exception:
                pass
            return False

        view_name, count = ctx
        # Plain-string copy for worker threads (auto model triage).
        self._ctx_line = u"active view: {}; {} element(s) selected".format(
            view_name or u"none", count)
        try:
            self.revit_ctx_view.Text = view_name or u"No active view"
            if count:
                self.revit_ctx_sel.Text = (u"· 1 selected" if count == 1
                                           else u"· {} selected".format(count))
            else:
                self.revit_ctx_sel.Text = u""
            self.revit_ctx_btn.Visibility = Visibility.Visible
        except Exception:
            return False
        return True

    def revit_ctx_clicked(self, sender, e):
        """Re-read Revit immediately instead of waiting for the next tick."""
        self._update_revit_context()

    # ─── Window state persistence ─────────────────────────────────────────────

    def _restore_window_state(self):
        """Restore window position and size from settings with strict boundary clamping."""
        if getattr(self, 'is_docked', False):
            return

        try:
            from config.settings import T3LabAISettings
            ws = T3LabAISettings().get_window_state()

            # Strict clamp: never opens full screen
            w = ws.get('width')
            h = ws.get('height')
            if w and h:
                w_f = float(w)
                h_f = float(h)
                self.Width = w_f if 480.0 <= w_f <= 720.0 else 560.0
                self.Height = h_f if 580.0 <= h_f <= 950.0 else 720.0
            else:
                self.Width  = 560.0
                self.Height = 720.0

            # Must still be on-screen; default to screen right
            left = ws.get('left')
            top  = ws.get('top')
            try:
                import System.Windows
                sw = System.Windows.SystemParameters.PrimaryScreenWidth
                sh = System.Windows.SystemParameters.PrimaryScreenHeight
                if left is not None and top is not None:
                    left_f = float(left)
                    top_f  = float(top)
                    if 0 <= left_f <= sw - 100 and 0 <= top_f <= sh - 60:
                        self.Left = left_f
                        self.Top  = top_f
                    else:
                        self.Left = max(40.0, sw - self.Width - 40.0)
                        self.Top  = max(40.0, (sh - self.Height) / 2.0)
                else:
                    self.Left = max(40.0, sw - self.Width - 40.0)
                    self.Top  = max(40.0, (sh - self.Height) / 2.0)
            except Exception:
                pass

        except Exception as ex:
            logger.debug(u"_restore_window_state error: {}".format(_exc_text(ex)))

    def _save_window_state(self):
        """Persist current window geometry to settings (skipped when docked or maximized)."""
        if getattr(self, 'is_docked', False):
            return

        try:
            from System.Windows import WindowState
            if getattr(self, 'WindowState', None) != WindowState.Normal:
                return

            from config.settings import T3LabAISettings
            w = max(480.0, min(float(self.Width), 720.0))
            h = max(580.0, min(float(self.Height), 950.0))
            T3LabAISettings().save_window_state(
                float(self.Left), float(self.Top),
                w, h,
                False,
            )
        except Exception as ex:
            logger.debug(u"_save_window_state error: {}".format(_exc_text(ex)))

    # ─── Window controls ──────────────────────────────────────────────────────

    def _on_closing(self, sender, e):
        """Persist window state on close (Alt+F4 / pane close) and stop timers."""
        self._save_window_state()
        # DispatcherTimers keep ticking after close and hold the window alive:
        # a due scheduled task would run an invisible LLM request, and each
        # reopen leaked another window tree.
        try:
            self._sched_timer.Stop()
        except Exception:
            pass
        try:
            self._typing_timer.Stop()
        except Exception:
            pass
        try:
            self._ctx_timer.Stop()
        except Exception:
            pass
        self._stop_stop_watchdog()

    def minimize_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_clicked(self, sender, e):
        """Toggle between standard width (560px) and wide width (720px)."""
        try:
            if self.Width > 640.0:
                self.Width = 560.0
            else:
                self.Width = 720.0
        except Exception:
            pass

    def close_clicked(self, sender, e):
        self.Close()

    def undo_clicked(self, sender, e):
        """Undo the last Revit transaction.

        DB.Document has no Undo(): it is a postable command, so PostCommand
        with a RevitCommandId. UIApplication from _get_uiapp(), NOT
        `HOST_APP.uiapp` (None here).
        """
        try:
            from Autodesk.Revit.UI import RevitCommandId, PostableCommand
            uiapp = _get_uiapp()
            if uiapp is None:
                raise AttributeError("no UIApplication available")
            cid = RevitCommandId.LookupPostableCommandId(PostableCommand.Undo)
            # False while a modal dialog or another command owns the UI.
            if not uiapp.CanPostCommand(cid):
                self._append_bot_message(
                    u"Revit is busy (a dialog or command is open) — "
                    u"close it and press Undo again.",
                    icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                return
            uiapp.PostCommand(cid)
            # PostCommand is async and an empty undo stack is a silent no-op,
            # so the wording stays neutral.
            self._append_bot_message(u"Undo sent to Revit.")
        except Exception as ex:
            # Not _report_error: claiming the turn would suppress a running reply.
            logger.error(u"undo: {}".format(_exc_text(ex)))
            self._append_bot_message(
                u"Undo is unavailable right now — use Ctrl+Z in Revit.",
                icon=_ICON_WARNING, icon_color=_ICON_AMBER)

    # ─── Tool discovery bootstrap ──────────────────────────────────────────────

    def _bootstrap_discovered_tools(self):
        """Discover tools, register launchers + NLP prompt, announce NEW tools. UI THREAD."""
        try:
            if not HAS_DISCOVERY:
                return

            # ── Discover (writes registry, drops vanished buttons) ───────────
            new_tools = discover_new_tools()

            # ── Register launchers + inject into NLP (ONE pass, full registry) ──
            _register_discovered_launchers(get_registered_tools())

            if DROPPED_TOOLS:
                self._log_activity(
                    u"Tool catalog: skipped {} intent(s) with no script on "
                    u"disk — {}".format(
                        len(DROPPED_TOOLS),
                        u", ".join(i for i, _p in DROPPED_TOOLS[:8])))

            # ── Chat notification for NEW tools only ──────────────────────────
            if new_tools:
                names = u', '.join(t['title'] for t in new_tools[:5])
                if len(new_tools) > 5:
                    names += u'...'
                self._append_bot_message(
                    u"Discovered {} new tools: {}.\n"
                    u"I learned them and can open them from natural-language commands.".format(
                        len(new_tools), names),
                    icon=_ICON_SEARCH, icon_color=_ICON_BLUE
                )
        except Exception as ex:
            logger.debug(u"_bootstrap_discovered_tools error: {}".format(_exc_text(ex)))

    # ─── Old-context repeat guard ──────────────────────────────────────────────
    # Every (intent, args) COMPLETED in this conversation, newest last: small
    # local models re-issue earlier turns' work (one "tô đỏ tường" also
    # recoloured Floors/Doors from older turns). Bounded.
    _MAX_REMEMBERED_CALLS = 60

    def _earlier_turn_calls(self):
        """Calls completed in earlier turns of this conversation."""
        return set(getattr(self, '_prev_turn_calls', None) or set())

    def _remember_turn_calls(self, calls):
        """Fold this turn's executed calls into the conversation-wide set."""
        try:
            known = getattr(self, '_prev_turn_calls', None)
            if not isinstance(known, set):
                known = set()
            order = getattr(self, '_prev_turn_call_order', None)
            if not isinstance(order, list):
                order = [k for k in known]
            for key in calls:
                if key not in known:
                    known.add(key)
                    order.append(key)
            # Drop the oldest: an old command may legitimately be wanted again.
            while len(order) > self._MAX_REMEMBERED_CALLS:
                known.discard(order.pop(0))
            self._prev_turn_calls = known
            self._prev_turn_call_order = order
        except Exception as ex:
            logger.debug(u"repeat-guard bookkeeping: {}".format(_exc_text(ex)))

    # ─── Assistant Sessions: GUI/assistant/sessions.py (self._sessions) ───

    def _session_prompt_block(self, viet=False, tool_hint=True):
        """Looked up by name: AssistantSessions._session_prompt_block."""
        return self._sessions._session_prompt_block(viet, tool_hint)

    def new_chat_clicked(self, sender=None, e=None):
        """XAML handler: AssistantSessions.new_chat_clicked."""
        return self._sessions.new_chat_clicked(sender, e)

    def tab_chat_clicked(self, sender=None, e=None):
        """XAML handler: AssistantSessions.tab_chat_clicked."""
        return self._sessions.tab_chat_clicked(sender, e)

    def tab_history_clicked(self, sender=None, e=None):
        """XAML handler: AssistantSessions.tab_history_clicked."""
        return self._sessions.tab_history_clicked(sender, e)

    def suggestion_card_clicked(self, sender=None, e=None):
        """Handle click on one of the suggestion prompt cards."""
        try:
            tag = getattr(sender, 'Tag', None)
            if not tag:
                return
            prompt_text = str(tag).strip()
            if prompt_text:
                self.chat_input.Text = prompt_text
                self._process_input()
        except Exception as ex:
            logger.debug(u"suggestion_card_clicked error: {}".format(_exc_text(ex)))

    # ─── History view: per-project list, search, scope, rename ───────────────

    #: Search debounce: each render re-reads the archive metadata.
    _HISTORY_SEARCH_DELAY_MS = 300

    # ─── Assistant Sessions: GUI/assistant/sessions.py (self._sessions) ───

    def history_search_changed(self, sender=None, e=None):
        """XAML handler: AssistantSessions.history_search_changed."""
        return self._sessions.history_search_changed(sender, e)

    def history_scope_changed(self, sender=None, e=None):
        """XAML handler: AssistantSessions.history_scope_changed."""
        return self._sessions.history_scope_changed(sender, e)

    def clear_all_history_clicked(self, sender=None, e=None):
        """XAML handler: AssistantSessions.clear_all_history_clicked."""
        return self._sessions.clear_all_history_clicked(sender, e)

    # ─── Saved Prompts & Skills Management ────────────────────────────────────

    def saved_prompts_clicked(self, sender=None, e=None):
        """Toggle the Saved Prompts & Skills popup."""
        try:
            is_open = getattr(self.saved_prompts_popup, 'IsOpen', False)
            if is_open:
                self.saved_prompts_popup.IsOpen = False
            else:
                self._populate_saved_prompts_list("")
                self._keep_popup_in_pane(self.saved_prompts_popup)
                self.saved_prompts_popup.IsOpen = True
                try:
                    self.saved_prompts_search.Text = ""
                    self.saved_prompts_search.Focus()
                except Exception:
                    pass
        except Exception as ex:
            logger.debug(u"saved_prompts_clicked error: {}".format(_exc_text(ex)))

    def close_saved_prompts_clicked(self, sender=None, e=None):
        """Close the Saved Prompts & Skills popup."""
        try:
            self.saved_prompts_popup.IsOpen = False
        except Exception:
            pass

    def saved_prompts_search_changed(self, sender=None, e=None):
        """Filter prompts/skills list as the user types."""
        try:
            q = (self.saved_prompts_search.Text or u"").strip().lower()
            self._populate_saved_prompts_list(q)
        except Exception as ex:
            logger.debug(u"saved_prompts_search_changed error: {}".format(_exc_text(ex)))

    def open_skills_folder_clicked(self, sender=None, e=None):
        """Open the user skills folder (custom prompt .md files) in Explorer."""
        try:
            from Intelligence.skills_engine import _user_skills_dir
            udir = _user_skills_dir()
            if not os.path.exists(udir):
                try:
                    os.makedirs(udir)
                except Exception:
                    pass
            _open_in_explorer(udir)
        except Exception as ex:
            logger.debug(u"open_skills_folder_clicked error: {}".format(_exc_text(ex)))

    def manage_skills_clicked(self, sender=None, e=None):
        """Open the Skills management tab in the LLMs Setting dialog."""
        try:
            self.saved_prompts_popup.IsOpen = False
        except Exception:
            pass
        self._open_llm_settings(tab='skills')

    def _populate_saved_prompts_list(self, query=u""):
        """Build clickable rows for each saved prompt / skill. UI THREAD."""
        try:
            from System.Windows import Thickness, CornerRadius, GridUnitType, GridLength
            from System.Windows.Controls import Border, Grid, StackPanel, TextBlock, ColumnDefinition
            from System.Windows.Input import Cursors
            from Intelligence.skills_engine import get_skills_engine

            self.saved_prompts_list_panel.Children.Clear()
            engine = get_skills_engine()
            skills = engine.all_skills() or []

            if query:
                q = query.lower()
                filtered = []
                for s in skills:
                    name = (s.get('name') or u'').lower()
                    desc = (s.get('description') or u'').lower()
                    sid = (s.get('id') or u'').lower()
                    trigs = u" ".join(s.get('triggers') or []).lower()
                    if q in name or q in desc or q in sid or q in trigs:
                        filtered.append(s)
                skills = filtered

            if not skills:
                tb_none = TextBlock()
                tb_none.Text = u"No matching prompts or skills found."
                tb_none.Style = self.FindResource("T3.Caption")
                tb_none.Margin = Thickness(12, 16, 12, 16)
                tb_none.HorizontalAlignment = System.Windows.HorizontalAlignment.Center
                self.saved_prompts_list_panel.Children.Add(tb_none)
                return

            bg_normal = self.FindResource("T3.Surface")
            bg_hover = self.FindResource("T3.SurfaceSunken")
            br_normal = self.FindResource("T3.Border")
            br_hover = self.FindResource("T3.BorderStrong")

            for s in skills:
                sid = s.get('id') or u''
                name = s.get('name') or sid
                desc = s.get('description') or u''

                card = Border()
                card.Background = bg_normal
                card.BorderBrush = br_normal
                card.BorderThickness = Thickness(1)
                card.CornerRadius = CornerRadius(6)
                card.Margin = Thickness(0, 0, 0, 4)
                card.Padding = Thickness(10, 7, 10, 7)
                card.Cursor = Cursors.Hand
                card.ToolTip = u"Click to insert: /{}\n{}".format(sid, desc)

                def _make_hover(target_card):
                    def _enter(sender, ev):
                        target_card.Background = bg_hover
                        target_card.BorderBrush = br_hover
                    def _leave(sender, ev):
                        target_card.Background = bg_normal
                        target_card.BorderBrush = br_normal
                    return _enter, _leave

                _ent, _lev = _make_hover(card)
                card.MouseEnter += _ent
                card.MouseLeave += _lev

                grid = Grid()
                col0 = ColumnDefinition()
                col0.Width = GridLength(1.0, GridUnitType.Star)
                col1 = ColumnDefinition()
                col1.Width = GridLength.Auto
                grid.ColumnDefinitions.Add(col0)
                grid.ColumnDefinitions.Add(col1)

                info_stack = StackPanel()
                info_stack.VerticalAlignment = System.Windows.VerticalAlignment.Center
                Grid.SetColumn(info_stack, 0)

                tb_name = TextBlock()
                tb_name.Text = name
                tb_name.Style = self.FindResource("T3.BodyStrong")
                tb_name.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
                info_stack.Children.Add(tb_name)

                if desc:
                    tb_desc = TextBlock()
                    tb_desc.Text = desc
                    tb_desc.Style = self.FindResource("T3.Caption")
                    tb_desc.Foreground = self.FindResource("T3.TextSecondary")
                    tb_desc.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
                    tb_desc.Margin = Thickness(0, 2, 0, 0)
                    info_stack.Children.Add(tb_desc)

                grid.Children.Add(info_stack)

                badge = Border()
                badge.Background = self.FindResource("T3.SurfaceSunken")
                badge.BorderBrush = self.FindResource("T3.Border")
                badge.BorderThickness = Thickness(1)
                badge.CornerRadius = CornerRadius(4)
                badge.Padding = Thickness(6, 2, 6, 2)
                badge.Margin = Thickness(8, 0, 0, 0)
                badge.VerticalAlignment = System.Windows.VerticalAlignment.Center
                Grid.SetColumn(badge, 1)

                tb_action = TextBlock()
                tb_action.Text = u"/{}".format(sid)
                tb_action.Style = self.FindResource("T3.Mono")
                tb_action.FontSize = 10.5
                tb_action.Foreground = self.FindResource("T3.TextMuted")
                tb_action.VerticalAlignment = System.Windows.VerticalAlignment.Center
                badge.Child = tb_action

                grid.Children.Add(badge)
                card.Child = grid

                def _make_pick_handler(skill_id):
                    def _h(sender_btn, ev):
                        try:
                            self.saved_prompts_popup.IsOpen = False
                            self.tab_chat_clicked()
                            self.chat_input.Text = u"/{} ".format(skill_id)
                            self.chat_input.CaretIndex = len(self.chat_input.Text)
                            self.chat_input.Focus()
                        except Exception as ex:
                            logger.debug(u"pick skill error: {}".format(_exc_text(ex)))
                    return _h

                card.MouseLeftButtonUp += _make_pick_handler(sid)
                self.saved_prompts_list_panel.Children.Add(card)
        except Exception as ex:
            logger.debug(u"_populate_saved_prompts_list error: {}".format(_exc_text(ex)))

    def _reset_session_state(self, clear_transcript=True):
        """Drop everything tied to the OLD conversation scope. UI THREAD.

        Shared by reset_chat_clicked / _activate_project / _create_new_project.
        A stale `_stream_row` / `_typing_row` strands the answer in an orphaned
        TextBlock and stalls `_show_typing_indicator`.
        """
        try:
            if clear_transcript:
                # Keep the welcome greeting panel (first child)
                while self.chat_history_panel.Children.Count > 1:
                    self.chat_history_panel.Children.RemoveAt(1)
        except Exception:
            pass
        try:
            self._clear_stream_refs()
        except Exception:
            pass
        self._typing_row        = None
        self._typing_text_block = None
        self._quickreply_row    = None
        self._queued_inputs     = []      # their chips were just removed
        self._conversation_history = []
        self._history_summary      = u''
        self._persisted_msgs       = []
        # New id on first message; earlier-chats digest rebuilt (old chat is one now).
        self._session_id           = None
        self._session_path         = None
        self._session_notes        = []
        self._past_chats_cache     = None
        self._prev_agent_decision  = None  # specialist choice must not carry over
        self._prev_turn_calls      = set()  # repeat guard is per-conversation
        self._prev_turn_call_order = []     # insertion order, for the cap
        # Attachments / forced /skill carried into a new project mis-file documents.
        try:
            self._attached_files = []
            self._refresh_attachment_panel()
        except Exception:
            pass
        self._forced_skill_id   = None
        self._forced_skill_args = u""

    # ─── AI badge & provider switcher ────────────────────────────────────────

    # Provider brand colors, shared with the LLMs Setting dialog.
    _BADGE_COLORS = _SHARED_PROVIDER_COLORS
    _BADGE_GRAY = _SHARED_PROVIDER_GRAY   # #A1A1AA — no provider / offline

    def _render_greeting(self, name, fresh=False):
        """Set the welcome greeting for a name (no settings read).

        Phrase (GUI/AssistantGreetings) redrawn only on fresh=True, first
        render, or a period / language change, so it does not reshuffle.
        """
        try:
            viet = bool(_ui_viet())
            key = (_greetings.period_for_hour(), viet)
            tpl = getattr(self, '_greeting_tpl', None)
            if fresh or tpl is None or key != getattr(self, '_greeting_key', None):
                tpl = _greetings.pick_template(viet=viet)
                self._greeting_tpl = tpl
                self._greeting_key = key
            self.welcome_greeting_text.Text = _greetings.fill(
                tpl, name or u"Thạnh")
        except Exception:
            pass

    def _update_welcome_greeting(self, fresh=False):
        """Refresh the greeting from settings; fresh=True only on New chat / Reset."""
        try:
            from config.user_profile import UserProfile
            name = UserProfile().get_name() or u"Thạnh"
            self._render_greeting(name, fresh=fresh)

            if self._persisted_msgs:
                self.welcome_greeting_panel.Visibility = Visibility.Collapsed
            else:
                self.welcome_greeting_panel.Visibility = Visibility.Visible
        except Exception:
            pass

    # ─── First-run onboarding ─────────────────────────────────────────────────

    def _selected_onboarding_provider(self):
        """Return the provider tag selected in the onboarding combo."""
        try:
            item = self.onboarding_provider_combo.SelectedItem
            if item is not None and item.Tag:
                return str(item.Tag)
        except Exception:
            pass
        return "claude"

    def _sync_onboarding_key_panel(self):
        """Hide the API-key field for local providers; relabel for remote ones."""
        try:
            prov = self._selected_onboarding_provider()
            is_local = prov in ("ollama", "lmstudio")
            self.onboarding_key_panel.Visibility = (
                Visibility.Collapsed if is_local else Visibility.Visible)
            if not is_local:
                labels = {
                    "claude":   u"ANTHROPIC API KEY",
                    "openai":   u"OPENAI API KEY",
                    "deepseek": u"DEEPSEEK API KEY",
                    "gemini":   u"GEMINI API KEY",
                }
                self.onboarding_key_label.Text = labels.get(prov, u"API KEY")
        except Exception:
            pass

    def _show_onboarding(self):
        """Display the first-run onboarding card."""
        try:
            from config.user_profile import UserProfile
            prof = UserProfile()
            try:
                nm = prof.get_name(fallback=False)
                if nm:
                    self.onboarding_name_box.Text = nm
            except Exception:
                pass

            try:
                self.onboarding_greeting.Text = u"{}!".format(_time_greeting())
            except Exception:
                pass

            self._sync_onboarding_key_panel()
            self.onboarding_overlay.Visibility = Visibility.Visible
            try:
                self.onboarding_name_box.Focus()
            except Exception:
                pass
        except Exception as ex:
            logger.debug(u"_show_onboarding error: {}".format(_exc_text(ex)))

    def _hide_onboarding(self):
        try:
            self.onboarding_overlay.Visibility = Visibility.Collapsed
        except Exception:
            pass

    def onboarding_provider_changed(self, sender, e):
        self._sync_onboarding_key_panel()

    def onboarding_skip_clicked(self, sender, e):
        """Dismiss onboarding without saving — won't show again."""
        try:
            from config.user_profile import UserProfile
            UserProfile().mark_setup_completed()
        except Exception:
            pass
        self._hide_onboarding()

    def onboarding_save_clicked(self, sender, e):
        """Persist the new user's profile + model setup, then dismiss onboarding."""
        try:
            from System.Windows.Media import SolidColorBrush, Color
            name = (self.onboarding_name_box.Text or u"").strip()
            if not name:
                try:
                    self.onboarding_name_box.BorderBrush = SolidColorBrush(
                        Color.FromRgb(239, 68, 68))      # rose
                    self.onboarding_name_box.Focus()
                except Exception:
                    pass
                return

            from config.user_profile import UserProfile
            prof     = UserProfile()
            provider = self._selected_onboarding_provider()

            prof.set_name(name)
            prof.set_model_setup(provider)

            key = u""
            try:
                key = (self.onboarding_key_box.Text or u"").strip()
            except Exception:
                pass
            if key and provider in ("claude", "openai", "deepseek", "gemini"):
                key_map = {"claude": "Claude", "openai": "OpenAI", "deepseek": "DeepSeek",
                           "gemini": "Gemini"}
                try:
                    from config.settings import T3LabAISettings
                    T3LabAISettings().set_api_key(key_map[provider], key)
                except Exception:
                    pass

            prof.mark_setup_completed()

            try:
                from Intelligence.llm_router import LLMRouter
                router = LLMRouter()
                router.switch_provider(provider)
                p = router.get_active_provider()
                if p and hasattr(p, "reload_credentials"):
                    p.reload_credentials()
            except Exception:
                pass

            self._hide_onboarding()

            self._render_greeting(name)
            try:
                self._update_welcome_greeting()
            except Exception:
                pass
            try:
                self._update_ai_badge()
            except Exception:
                pass

            self._append_bot_message(
                u"Nice to meet you, {}!\n"
                u"Your profile is saved. Try 'open batchout' or ask me anything about Revit.".format(name),
                icon=_ICON_SUCCESS, icon_color=_ICON_GREEN)
        except Exception as ex:
            logger.debug(u"onboarding_save_clicked error: {}".format(_exc_text(ex)))

    def _update_ai_badge(self):
        """Refresh the composer model chip (the only provider indicator)."""
        try:
            self._update_composer_chips()
        except Exception:
            pass

    def _switch_provider(self, name):
        """Hot-swap the active LLM provider — instant UI, network probes in background."""
        try:
            from Intelligence.llm_router import LLMRouter
            router = LLMRouter()
            ok = router.switch_provider(name)
            if not ok:
                return

            # Chip renders from cached data (no network)
            self.Dispatcher.Invoke(Action(self._update_ai_badge))

            # Background probe + model list; guarded so rapid switches don't pile up threads.
            if self._switching_provider:
                return
            self._switching_provider = True

            def _bg():
                try:
                    provider = router.get_active_provider()
                    router.probe_provider(name)
                    if provider:
                        try:
                            self._models_cache[name] = provider.get_models()
                        except Exception:
                            pass
                        # Local provider: preload so the first message isn't a cold VRAM load.
                        if name in ("ollama", "lmstudio") and \
                                hasattr(provider, "warm_up"):
                            try:
                                import time as _time
                                provider.warm_up()
                                self._warmed_provider = name
                                self._warmed_at = _time.time()
                            except Exception:
                                pass
                except Exception:
                    pass
                finally:
                    self._switching_provider = False

            t = Thread(ThreadStart(_bg))
            t.IsBackground = True
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
        except Exception as ex:
            logger.debug(u"_switch_provider error: {}".format(_exc_text(ex)))

    # ─── Settings (LLMs Setting hub dialog) ──────────────────────────────────

    def settings_btn_clicked(self, sender, e):
        """Open the LLMs Setting hub (provider, model, key, projects, knowledge, skills)."""
        self._open_llm_settings()

    def _open_llm_settings(self, tab=None, select_pid=None):
        """Show the LLMs Setting dialog modally, then refresh what it may have changed.

        tab: 'general'/'provider'/'projects'/'knowledge'/'skills' to open on.
        select_pid preselects a project in Projects (after New project).
        """
        try:
            from GUI.LLMSettingDialog import LLMSettingWindow
            dlg = LLMSettingWindow()
            try:
                if select_pid:
                    dlg._load_projects_tab(select_pid=select_pid)
                if tab:
                    _rb = getattr(dlg, 'tab_' + tab, None)
                    if _rb is not None:
                        _rb.IsChecked = True
                if tab == 'projects':
                    def _focus_name(s, ev):
                        try:
                            dlg.project_name_box.Focus()
                            dlg.project_name_box.SelectAll()
                        except Exception:
                            pass
                    dlg.Loaded += _focus_name
            except Exception:
                pass
            _scope_before = self._projects._active_pid()
            dlg.ShowDialog()
        except Exception as ex:
            logger.debug(u"_open_llm_settings error: {}".format(_exc_text(ex)))
            _scope_before = self._projects._active_pid()
        self._refresh_after_settings(scope_before=_scope_before)

    # ─── Assistant Projects: GUI/assistant/projects.py (self._projects) ───

    def _refresh_after_settings(self, scope_before=None):
        """Sync the chat window with state edited in the settings dialog."""
        # A project DELETED in the dialog clears active_project; re-scope before
        # _persist_message writes the old transcript into the global history.
        try:
            _pid_now = self._projects._active_pid()
            if _pid_now != scope_before:
                self._reset_session_state(clear_transcript=True)
                self._sessions._restore_history()
                self._kick_knowledge_scan()
            # The provider/model override may have been edited; re-activation
            # is blocked for an unchanged pid, so apply it here.
            self._projects._apply_project_provider(_pid_now)
        except Exception as ex:
            logger.debug(u"scope re-sync error: {}".format(_exc_text(ex)))
        try:
            self._update_composer_chips()
        except Exception:
            pass
        try:
            self._update_action_mode_chip()
        except Exception:
            pass
        try:
            self._update_welcome_greeting()
        except Exception:
            pass
        # An open project panel was rendered from pre-edit data.
        try:
            if (getattr(self, 'project_panel_overlay', None) is not None
                    and self.project_panel_overlay.Visibility == Visibility.Visible):
                _pid = self._projects._active_pid()
                if _pid:
                    self._projects._build_project_panel(_pid)
                else:
                    self.project_panel_overlay.Visibility = Visibility.Collapsed
        except Exception as ex:
            logger.debug(u"project panel refresh error: {}".format(_exc_text(ex)))

    # ─── Assistant Projects: GUI/assistant/projects.py (self._projects) ───

    #: First line of assistant_memory.build_memory_block (detects it in a prompt).
    _MEMORY_BLOCK_HEAD = u"## Persistent memory"

    # ─── Knowledge (RAG v2 index) ─────────────────────────────────────────────

    def _update_knowledge_status(self):
        """No-op hook: knowledge UI lives in the LLMs Setting dialog."""
        pass

    def _kick_knowledge_scan(self):
        """(Re)scan the active knowledge store on a background thread."""
        if not HAS_KNOWLEDGE or getattr(self, '_kn_scan_busy', False):
            return
        self._kn_scan_busy = True

        def _scan():
            try:
                store = get_active_store()
                if store is not None:
                    store.scan()
                    try:
                        from Intelligence.knowledge.embeddings import (
                            get_default_embedder)
                        emb = get_default_embedder()
                        if emb is not None and emb.is_available():
                            store.embed_pending(emb, budget_sec=120)
                    except Exception:
                        pass
            except Exception as ex:
                logger.debug(u"knowledge scan error: {}".format(_exc_text(ex)))
            finally:
                self._kn_scan_busy = False
        _kt = Thread(ThreadStart(_scan))
        _kt.IsBackground = True
        _kt.SetApartmentState(ApartmentState.STA)
        _kt.Start()

    # ─── Skills chat chips ────────────────────────────────────────────────────

    def _append_skill_chips(self, skill_ids):
        """Small 'skill activated' chip row in the chat. UI THREAD."""
        try:
            from System.Windows.Controls import Border, TextBlock, StackPanel, Orientation
            from System.Windows import Thickness, CornerRadius
            from System.Windows.Media import SolidColorBrush, Color
            from Intelligence.skills_engine import get_skills_engine

            engine = get_skills_engine()
            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Margin = Thickness(0, 0, 8, 8)
            added = False
            for sid in skill_ids:
                meta = None
                try:
                    meta = engine._skills.get(sid)
                except Exception:
                    pass
                name = (meta or {}).get('name', sid)
                chip = Border()
                _bind_bg(chip, 'SelectedBg')
                _bind_border(chip, 'CardBorder')
                chip.BorderThickness = Thickness(1)
                chip.CornerRadius = CornerRadius(4)
                chip.Padding = Thickness(8, 2, 8, 3)
                chip.Margin = Thickness(0, 0, 4, 0)
                # MDL2 glyph run, not a ⚡ emoji (full colour, ignores the theme).
                tb = TextBlock()
                tb.FontSize = 10
                tb.FontFamily = System.Windows.Media.FontFamily("Segoe UI")
                _bind_fg(tb, 'Muted')
                from System.Windows.Documents import Run as _WpfRun
                _gr = _WpfRun()
                _gr.Text = _ICON_ANALYZE + u"  "
                _gr.FontFamily = System.Windows.Media.FontFamily(u"Segoe MDL2 Assets")
                tb.Inlines.Add(_gr)
                _nr = _WpfRun()
                _nr.Text = name
                tb.Inlines.Add(_nr)
                chip.Child = tb
                chip.ToolTip = u"Skill applied to this reply"
                row.Children.Add(chip)
                added = True
            if added:
                self.chat_history_panel.Children.Add(row)
                self._scroll_to_bottom()
        except Exception as ex:
            logger.debug(u"_append_skill_chips error: {}".format(_exc_text(ex)))

    # ─── Quick-reply chips (confirm / retry / continue) ───────────────────────

    def _append_quick_replies(self, labels, send_texts=None):
        """Clickable reply chips under the last bot message. UI THREAD.

        send_texts: what each click sends (default: the label). Only ONE row
        exists; removed on a click or any other send (_remove_quick_replies).
        """
        try:
            self._remove_quick_replies()
            from System.Windows.Controls import Border, TextBlock, StackPanel, Orientation
            from System.Windows import Thickness, CornerRadius
            from System.Windows.Media import SolidColorBrush, Color
            from System.Windows.Input import Cursors

            texts = list(send_texts) if send_texts else list(labels)
            row = StackPanel()
            row.Orientation = Orientation.Horizontal
            row.Margin = Thickness(0, 2, 8, 12)

            _bg      = _theme_color('AccentSoft')
            _bg_hov  = _theme_color('IconHoverBg')
            for i, label in enumerate(labels):
                chip = Border()
                chip.Background = SolidColorBrush(_bg)
                _bind_border(chip, 'AccentSoft')
                chip.BorderThickness = Thickness(1)
                chip.CornerRadius = CornerRadius(4)
                chip.Padding = Thickness(11, 4, 11, 5)
                chip.Margin = Thickness(0, 0, 6, 0)
                chip.Cursor = Cursors.Hand
                tb = TextBlock()
                tb.Text = label
                tb.FontSize = 11.5
                tb.FontFamily = System.Windows.Media.FontFamily("Segoe UI")
                tb.FontWeight = System.Windows.FontWeights.SemiBold
                _bind_fg(tb, 'Accent')
                chip.Child = tb

                def _click(s, ev, _t=texts[i] if i < len(texts) else label):
                    try:
                        self._remove_quick_replies()
                        self.chat_input.Text = _t
                        self._process_input()
                    except Exception:
                        pass
                    ev.Handled = True
                chip.MouseLeftButtonUp += _click

                def _enter(s, ev, _c=chip):
                    _c.Background = SolidColorBrush(_bg_hov)

                def _leave(s, ev, _c=chip):
                    _c.Background = SolidColorBrush(_bg)
                chip.MouseEnter += _enter
                chip.MouseLeave += _leave
                row.Children.Add(chip)

            self._quickreply_row = row
            self.chat_history_panel.Children.Add(row)
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug(u"_append_quick_replies error: {}".format(_exc_text(ex)))

    def _remove_quick_replies(self):
        """Drop the active quick-reply row (if any). UI THREAD."""
        try:
            if self._quickreply_row is not None:
                self.chat_history_panel.Children.Remove(self._quickreply_row)
        except Exception:
            pass
        self._quickreply_row = None

    # ─── Turn termination (cancel / error) ────────────────────────────────────

    def _ui_invoke(self, fn, wait=True):
        """Run fn on the UI thread, swallowing a shut-down dispatcher.

        (A throw on a .NET background thread can take the process down.)
        wait=False posts (BeginInvoke) so a pure UI update cannot deadlock a
        worker the UI thread waits on. Same FIFO order either way.
        """
        try:
            if wait:
                self.Dispatcher.Invoke(Action(fn))
            else:
                self.Dispatcher.BeginInvoke(Action(fn))
        except Exception:
            pass

    def _claim_turn(self, rid=None):
        """True if this caller owns the turn's terminal message (worker vs Stop watchdog)."""
        if rid is not None and rid != getattr(self, '_request_id', 0):
            return False
        if getattr(self, '_replied', False):
            return False
        self._replied = True
        return True

    def _cancelled(self):
        """True once Stop has been pressed for the running turn."""
        return bool(getattr(self, '_cancel_requested', False))

    def _finish_cancelled(self, rid=None, note=None):
        """End the turn because the user pressed Stop. Any thread.

        Checked BETWEEN steps: a Revit Transaction in flight is never aborted.
        """
        if rid is not None and rid != getattr(self, '_request_id', 0):
            return
        if getattr(self, '_batchout_request_id', None) == getattr(self, '_request_id', 0):
            if getattr(self, '_batchout_running', False):
                # The running export must report its retained-file count.
                return
            # Queued: nothing touched yet; its callback sees the stale id.
            self._batchout_request_id = None
        if not self._claim_turn(rid):
            return

        def _ui():
            try:
                ran = getattr(self, '_tool_runs', 0)
                if ran:
                    msg = (u"Stopped. {} step(s) had already run and were "
                           u"not undone.".format(ran))
                else:
                    msg = u"Stopped."
                if note:
                    msg += u" " + note
                try:
                    self._remove_stream_bubble()
                    self._clear_stream_refs()
                except Exception:
                    pass
                self._hide_typing_indicator()
                self._append_bot_message(msg, icon=_ICON_WARNING,
                                         icon_color=_ICON_SLATE)
            except Exception:
                pass
            finally:
                self._set_busy(False)
        self._ui_invoke(_ui)

    def _report_error(self, where, exc=None, rid=None, hint=None):
        """End the turn with a VISIBLE message. Any thread.

        logger.error alone goes to a pyRevit console that is not open.
        """
        try:
            logger.error(u"{}: {}".format(where, _exc_text(exc)))
        except Exception:
            pass
        # Grab the failing frame NOW: sys.exc_info() is only valid inside the
        # caller's except block; the message is built later on the UI thread.
        _frame = u""
        try:
            import traceback as _tbmod
            _frames = _tbmod.extract_tb(sys.exc_info()[2])
            if _frames:
                _f = _frames[-1]
                _frame = u"{}:{} in {}()".format(
                    os.path.basename(u"{}".format(_f[0])), _f[1], _f[2])
        except Exception:
            _frame = u""
        if not self._claim_turn(rid):
            return

        def _ui():
            try:
                msg = (u"Something went wrong while handling that request — "
                       u"nothing else was changed in the model. Please try "
                       u"again; the details are in the pyRevit console.")
                # Show the failure inline: in a docked pane the console is closed.
                _detail = u""
                try:
                    _detail = u"{}".format(_exc_text(exc) or u"").strip()
                    _detail = _detail.splitlines()[-1] if _detail else u""
                except Exception:
                    _detail = u""
                if _detail:
                    msg += u"\n\n`{}: {}`".format(where, _detail[:300])
                if _frame:
                    msg += u"\n`at {}`".format(_frame)
                if hint:
                    msg = hint + u"\n\n" + msg
                try:
                    self._remove_stream_bubble()
                    self._clear_stream_refs()
                except Exception:
                    pass
                self._hide_typing_indicator()
                self._append_bot_message(msg, icon=_ICON_WARNING,
                                         icon_color=_ICON_AMBER)
            except Exception:
                pass
            finally:
                self._set_busy(False)
        self._ui_invoke(_ui)

    def _stop_stop_watchdog(self):
        """Cancel the Stop watchdog timer, if one is armed."""
        try:
            t = getattr(self, '_stop_timer', None)
            if t is not None:
                t.Stop()
        except Exception:
            pass
        self._stop_timer = None

    # ─── Turn telemetry ───────────────────────────────────────────────────────
    # Started / flushed in _set_busy: every routing path and Stop pass through it.

    def _begin_turn_timer(self):
        """Start measuring the turn about to run. UI thread. Never raises.

        Everything read here must be free — NOT has_local_llm() (an HTTP
        check_health probe); the local flag comes from the provider name.
        """
        self._turn_timer = None
        if not HAS_TELEMETRY:
            return
        try:
            name = get_active_provider_name()
            self._turn_timer = telemetry.TurnTimer(
                provider=name,
                model=self._active_model_name(),
                local=(name in ("ollama", "lmstudio")))
        except Exception:
            self._turn_timer = None

    def _mark_first_token(self):
        """First streamed delta reached the UI. Any thread; idempotent."""
        t = getattr(self, '_turn_timer', None)
        if t is not None:
            try:
                t.mark_first_token()
            except Exception:
                pass

    def _note_turn_route(self, specialist=None, skills=None):
        """Record which specialist/skills handled the turn. Any thread."""
        t = getattr(self, '_turn_timer', None)
        if t is None:
            return
        try:
            if specialist:
                t.specialist = specialist
            if skills:
                t.skills = list(skills)
        except Exception:
            pass

    def _end_turn_timer(self):
        """Flush the finished turn to telemetry. UI thread; write on a pool thread."""
        t = getattr(self, '_turn_timer', None)
        self._turn_timer = None
        if t is None:
            return
        try:
            t.finish(u'cancelled' if self._cancelled() else u'done')
        except Exception:
            return
        try:
            self._note_turn_cost(t)
        except Exception:
            pass
        try:
            from System.Threading import ThreadPool, WaitCallback

            def _write(_state, _t=t):
                try:
                    telemetry.record(_t)
                except Exception:
                    pass
                try:
                    from Intelligence import result_store
                    result_store.prune_once()
                except Exception:
                    pass
            ThreadPool.QueueUserWorkItem(WaitCallback(_write))
        except Exception:
            try:
                telemetry.record(t)
            except Exception:
                pass

    def _note_turn_cost(self, t):
        """Turn USD + today's total on the model chip tooltip. UI thread.

        Chip text warns past agents.daily_budget_usd (0 = off). The just-logged
        turn is added in memory so the tooltip never lags a turn."""
        from Intelligence import pricing
        usd = t.usd
        if usd is None:
            self._cost_hint = u""
            return
        today = telemetry.spend_today() + usd
        budget = 0.0
        try:
            from config.settings import get_settings as _gs
            budget = _gs().get_daily_budget_usd()
        except Exception:
            budget = 0.0
        hint = u"Last turn {} \u00b7 today {}".format(
            pricing.format_usd(usd), pricing.format_usd(today))
        over = budget > 0 and today > budget
        if budget > 0:
            hint += u" / budget {}".format(pricing.format_usd(budget))
        self._cost_hint = (u"\u26a0 over budget" if over
                           else pricing.format_usd(today) + u" today")
        try:
            self.model_chip_btn.ToolTip = hint
        except Exception:
            pass
        self._update_composer_chips()

    _PROVIDER_LABELS = {
        'claude': u'Claude (Anthropic)', 'openai': u'OpenAI',
        'deepseek': u'DeepSeek', 'ollama': u'Ollama (local)',
        'lmstudio': u'LM Studio (local)', 'gemini': u'Gemini (Google)',
        'custom': u'Custom (OpenAI-compatible)',
    }

    def _try_which_model(self, raw):
        """Answer "which model/LLM are you running?" from fact. WORKER THREAD.

        True when handled. Known exactly, so it must never reach the LLM
        (see the call site in _route_input).
        """
        try:
            if not routing.asks_which_llm(raw):
                return False
        except Exception:
            return False

        try:
            return self._answer_which_model(raw)
        except Exception as ex:
            logger.debug(u'_try_which_model error: {}'.format(_exc_text(ex)))
            return False

    def _answer_which_model(self, raw):
        """Post the provider/model answer; a failure falls back to the normal path."""
        viet = _is_viet_text(raw)
        try:
            prov = get_active_provider_name() or u''
        except Exception:
            prov = u''
        label = self._PROVIDER_LABELS.get(prov, prov or (
            u'chưa cấu hình' if viet else u'not configured'))
        model = self._active_model_name()

        if not prov:
            msg = (u"Hiện chưa có AI provider nào được cấu hình. "
                   u"Mở **LLMs Setting** để kết nối một provider."
                   if viet else
                   u"No AI provider is configured yet. Open **LLMs Setting** "
                   u"to connect one.")
        elif model:
            msg = (u"Mình đang chạy trên **{}**, model `{}`.\n"
                   u"Đổi trong **LLMs Setting**.".format(label, model)
                   if viet else
                   u"I'm running on **{}**, model `{}`.\n"
                   u"You can change it in **LLMs Setting**.".format(label, model))
        else:
            # Model id not resolved yet (cache-only); do not invent one.
            msg = (u"Mình đang chạy trên **{}**. Tên model cụ thể chưa được "
                   u"nạp — mở **LLMs Setting** để xem và đổi.".format(label)
                   if viet else
                   u"I'm running on **{}**. The exact model id isn't loaded "
                   u"yet — open **LLMs Setting** to see or change it."
                   .format(label))

        def _show():
            self._hide_typing_indicator()
            self._append_bot_message(msg, icon=_ICON_INFO,
                                     icon_color=_ICON_BLUE)
            self._add_to_history("assistant", msg)
            self._set_busy(False)

        self.Dispatcher.Invoke(Action(_show))
        return True

    def _active_model_name(self):
        """Model id of the active provider, or ''. Cache-only (no HTTP); never raises."""
        try:
            from Intelligence.llm_router import get_router
            prov = get_router().get_active_provider()
            return u'{}'.format(prov.get_active_model() or u'')
        except Exception:
            return u''

    # ─── Session guard & UI state ─────────────────────────────────────────────

    def _set_busy(self, busy):
        """Lock/unlock the input area. Call from UI thread only.

        While busy the send button becomes Stop and the input stays enabled:
        Enter queues the next message, drained when busy releases.
        """
        if (not busy and getattr(self, '_batchout_request_id', None)
                == getattr(self, '_request_id', 0)):
            # BatchOut's completion callback owns this turn's release.
            return
        self._busy = busy
        if busy:
            self._cancel_requested = False
            self._request_id = getattr(self, '_request_id', 0) + 1
            self._replied    = False
            self._tool_runs  = 0
            # Plan approval covers ONE plan's edits, never the next request.
            self._plan_preapproved = False
            self._begin_turn_timer()
            # The self-study loop backs off while active (Intelligence/learning).
            try:
                from Intelligence.learning import activity as _study_activity
                _study_activity.note_active()
            except Exception:
                pass
        else:
            self._stop_stop_watchdog()
            self._end_turn_timer()
        try:
            self.chat_input.IsEnabled = True
            self._render_send_button(busy)
            self.btn_attach.IsEnabled = not busy
        except Exception:
            pass
        for btn in self._DYNAMIC_BTNS:
            try:
                btn.IsEnabled = not busy
            except Exception:
                pass

        try:
            from System.Windows import Visibility
            if busy:
                self.top_loading_bar.Visibility = Visibility.Visible
            else:
                self.top_loading_bar.Visibility = Visibility.Collapsed
        except Exception:
            pass

        if busy:
            self._show_typing_indicator()
        else:
            self._hide_typing_indicator()
            # Rotation means "thinking"; a finished turn stands still.
            self._stop_avatar_spins()
            # An armed "Try again" target would file the next unrelated answer
            # as a version of an old message.
            self._retry_target_row = None

        # BeginInvoke lets the finish-handler stack unwind before the next send.
        if not busy and self._queued_inputs:
            try:
                self.Dispatcher.BeginInvoke(Action(self._drain_queued_input))
            except Exception:
                pass

    def _drain_queued_input(self):
        """Send the oldest queued message. UI THREAD (via BeginInvoke)."""
        try:
            if self._busy or not self._queued_inputs:
                return
            item = self._queued_inputs.pop(0)
            try:
                if item.get('row') is not None:
                    self.chat_history_panel.Children.Remove(item['row'])
            except Exception:
                pass
            # Keep the user's draft across the send.
            draft = self.chat_input.Text
            self.chat_input.Text = item['text']
            self._process_input()
            try:
                self.chat_input.Text = draft
                self.chat_input.CaretIndex = len(draft or u"")
            except Exception:
                pass
        except Exception as ex:
            logger.debug(u"_drain_queued_input error: {}".format(_exc_text(ex)))

    def _queue_input(self, raw):
        """Queue a message typed while busy, as a 'Queued' chip (click cancels). UI THREAD."""
        try:
            if len(self._queued_inputs) >= 3:
                self._append_bot_message(
                    u"Queue is full (3 messages) — please wait for the "
                    u"current request to finish.",
                    icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                return
            from System.Windows.Controls import Border, TextBlock
            from System.Windows import Thickness, CornerRadius, TextTrimming
            from System.Windows.Media import SolidColorBrush, Color
            from System.Windows.Input import Cursors

            row = Border()
            _bind_bg(row, 'SelectedBg')
            _bind_border(row, 'CardBorder')
            row.BorderThickness = Thickness(1)
            row.CornerRadius = CornerRadius(4)
            row.Padding = Thickness(10, 4, 10, 5)
            row.Margin = Thickness(40, 0, 0, 8)
            row.HorizontalAlignment = System.Windows.HorizontalAlignment.Right
            row.Cursor = Cursors.Hand
            row.ToolTip = (u"Will send automatically when the current "
                           u"request finishes — click to cancel")
            tb = TextBlock()
            _short = raw if len(raw) <= 70 else raw[:70] + u"…"
            tb.Text = u"Queued: {}".format(_short)
            tb.FontSize = 11
            tb.FontFamily = System.Windows.Media.FontFamily("Segoe UI")
            _bind_fg(tb, 'Muted')
            tb.TextTrimming = TextTrimming.CharacterEllipsis
            row.Child = tb
            item = {'text': raw, 'row': row}

            def _cancel_queued(s, ev, _item=item):
                try:
                    if _item in self._queued_inputs:
                        self._queued_inputs.remove(_item)
                    self.chat_history_panel.Children.Remove(_item['row'])
                except Exception:
                    pass
                ev.Handled = True
            row.MouseLeftButtonUp += _cancel_queued

            self._queued_inputs.append(item)
            self.chat_history_panel.Children.Add(row)
            self.chat_input.Text = u""
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug(u"_queue_input error: {}".format(_exc_text(ex)))

    def _render_send_button(self, busy):
        """Swap the send button's Path between arrow-up and Stop. UI thread only."""
        try:
            from System.Windows.Media import Geometry
            btn = self.send_button
            if not busy:
                self.send_icon.Data = Geometry.Parse(u"M12 19V5 M5 12l7-7 7 7")
                btn.ToolTip = u"Send (Enter)"
            else:
                self.send_icon.Data = Geometry.Parse(u"M8 8 h8 v8 h-8 Z")
                btn.ToolTip = u"Stop the running task"
            btn.IsEnabled = True
        except Exception as ex:
            logger.debug(u"_render_send_button error: {}".format(_exc_text(ex)))

    def _request_stop(self):
        """Stop button pressed while a request is running (UI thread)."""
        try:
            self._cancel_requested = True
            loop = self._agent_loop
            if loop is not None:
                loop.cancel()
            try:
                self.send_button.IsEnabled = False   # re-enabled by _set_busy(False)
            except Exception:
                pass
            self._safe_update_typing_text(u"● ● ●  Stopping after the current step…")
            self._arm_stop_watchdog()
        except Exception as ex:
            logger.debug(u"_request_stop error: {}".format(_exc_text(ex)))

    def _arm_stop_watchdog(self, seconds=20.0):
        """Release the UI even if the worker never reaches a cancel checkpoint.

        E.g. a hung provider call (server killed mid-generation). If the worker
        finishes first, the _request_id guard makes this a no-op.
        """
        self._stop_stop_watchdog()
        try:
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan
            rid = getattr(self, '_request_id', 0)
            timer = DispatcherTimer()
            timer.Interval = TimeSpan.FromSeconds(seconds)

            def _tick(s, ev):
                try:
                    timer.Stop()
                    if self._busy and self._cancelled():
                        self._finish_cancelled(
                            rid,
                            note=u"A background step is still finishing; "
                                 u"its result will be discarded.")
                except Exception:
                    pass
            timer.Tick += _tick
            timer.Start()
            self._stop_timer = timer
        except Exception:
            self._stop_timer = None

    @staticmethod
    def _quality_mode_on():
        """True when the Opus-parity / maximum-quality switch is enabled."""
        try:
            from config.settings import get_settings
            return bool(get_settings().is_quality_mode_enabled())
        except Exception:
            return False

    def _safe_update_typing_text(self, text):
        """Thread-safe update of the typing indicator text."""
        def action():
            try:
                if hasattr(self, "_typing_text_block") and self._typing_text_block is not None:
                    self._typing_text_block.Text = text
            except Exception:
                pass
        try:
            self.Dispatcher.Invoke(Action(action))
        except Exception:
            pass

    def _show_typing_indicator(self):
        """Add an animated '● ● ●' bubble to the chat."""
        try:
            if self._typing_row is not None:
                return  # already shown
            self._typing_row = self._make_typing_row()
            self.chat_history_panel.Children.Add(self._typing_row)
            self._scroll_to_bottom()

            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan
            self._typing_elapsed = 0
            self._update_typing_text()
            
            self._typing_timer = DispatcherTimer()
            self._typing_timer.Interval = TimeSpan.FromSeconds(1)
            self._typing_timer.Tick += self._on_typing_timer_tick
            self._typing_timer.Start()
        except Exception:
            pass

    def _hide_typing_indicator(self):
        """Remove the typing indicator bubble."""
        try:
            if self._typing_timer is not None:
                self._typing_timer.Stop()
                self._typing_timer = None
            if self._typing_row is not None:
                self.chat_history_panel.Children.Remove(self._typing_row)
                self._typing_row = None
                self._typing_text_block = None
        except Exception:
            pass

    def _on_typing_timer_tick(self, sender, e):
        self._typing_elapsed += 1
        self._update_typing_text()

    def _update_typing_text(self):
        try:
            if not hasattr(self, "_typing_text_block") or self._typing_text_block is None:
                return
            
            is_vn = _is_viet_text(self._last_raw)

            # Quality mode = extended reasoning; say so, so latency isn't read as a stall.
            if self._quality_mode_on():
                if self._typing_elapsed < 2:
                    txt = u"● ● ●  Đang suy luận sâu..." if is_vn else "● ● ●  Reasoning deeply..."
                else:
                    txt = u"● ● ●  Đang suy luận & xử lý..." if is_vn else "● ● ●  Thinking it through..."
                self._typing_text_block.Text = txt
                return

            if self._typing_elapsed < 1:
                txt = u"● ● ●  Đang đọc dữ liệu Revit..." if is_vn else "● ● ●  Reading Revit data..."
            elif self._typing_elapsed < 3:
                txt = u"● ● ●  Đang phân tích yêu cầu..." if is_vn else "● ● ●  Formulating response..."
            else:
                txt = u"● ● ●  Đang phản hồi..." if is_vn else "● ● ●  Responding..."

            self._typing_text_block.Text = txt
        except Exception:
            pass

    def _make_typing_row(self):
        """Typing indicator: spinning brand mark + muted status text.

        Replies carry no avatar (_append_bot_message): this is the ONE
        "working" mark per turn, where the answer will appear.
        """
        from System.Windows.Controls import TextBlock, Grid, ColumnDefinition
        from System.Windows import Thickness, GridLength, VerticalAlignment

        row = Grid()
        row.Margin = Thickness(0, 0, 8, 12)

        col_av = ColumnDefinition()
        col_av.Width = GridLength.Auto
        col_msg = ColumnDefinition()
        col_msg.Width = GridLength(1, System.Windows.GridUnitType.Star)
        row.ColumnDefinitions.Add(col_av)
        row.ColumnDefinitions.Add(col_msg)

        av = self._make_avatar("T3", size=24)
        av.Margin = Thickness(0, 0, 9, 0)
        av.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(av, 0)
        row.Children.Add(av)

        dots = TextBlock()
        dots.Text       = u"● ● ●"
        dots.FontSize   = 12.5
        dots.FontFamily = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
        _bind_fg(dots, 'Muted')
        dots.VerticalAlignment = VerticalAlignment.Center

        self._typing_text_block = dots

        Grid.SetColumn(dots, 1)
        row.Children.Add(dots)
        return row


    def _safe_append_bot(self, msg, icon=None, icon_color=None):
        """Thread-safe bot message append (can be called from background threads)."""
        try:
            self.Dispatcher.Invoke(Action(
                lambda: self._append_bot_message(msg, icon=icon, icon_color=icon_color)))
        except Exception:
            pass

    def _add_to_history(self, role, content):
        """Add a message to conversation history and persist to disk.

        Turns leaving the window fold into a rolling summary; the window is the
        ACTIVE model's token budget (Intelligence.context_budget), not a count.
        """
        self._conversation_history.append({"role": role, "content": content})
        try:
            from Intelligence import conversation as _conv
            _b = _active_budgets()
            self._conversation_history, self._history_summary = _conv.trim_history(
                self._conversation_history,
                getattr(self, '_history_summary', u''),
                budget_tokens=_b['history'], viet=_ui_viet(),
                max_summary_chars=_b['summary_chars'])
        except Exception:
            # Safety net only.
            if len(self._conversation_history) > 200:
                self._conversation_history = self._conversation_history[-200:]
        self._sessions._persist_message(role, content)
        # Activity journal: assistant side only (user side: _process_input).
        if role == "assistant" and content:
            one_line = u" ".join((content or u"").split())
            if len(one_line) > 400:
                one_line = one_line[:400] + u"…"
            self._log_activity(u"Assistant: {}".format(one_line))

    # ─── File attachment ──────────────────────────────────────────────────────

    def attach_clicked(self, sender, e):
        """Open a file picker and add selected file to attachment list."""
        try:
            import clr
            clr.AddReference('System.Windows.Forms')
            from System.Windows.Forms import OpenFileDialog, DialogResult

            _docs  = "*.pdf;*.txt;*.md;*.json;*.csv;*.tsv;*.log;*.xml;*.html;*.htm;*.yaml;*.yml;*.ini"
            _imgs  = "*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.webp"
            exts = ("Documents & Images|{};{}|"
                    "Documents (PDF, text, JSON...)|{}|"
                    "Images|{}|All files|*.*").format(_docs, _imgs, _docs, _imgs)
            dlg = OpenFileDialog()
            dlg.Title  = u"Choose documents or images to attach"
            dlg.Filter = exts
            dlg.Multiselect = True

            if dlg.ShowDialog() == DialogResult.OK:
                self._attach_paths(list(dlg.FileNames))
        except Exception as ex:
            logger.error(u"attach_clicked error: {}".format(_exc_text(ex)))

    def _attach_paths(self, paths):
        """Shared intake for picker / Ctrl+V / drag-drop. Returns count added. UI THREAD.

        With RAG present only supported formats (PDF + images) pass."""
        if self._busy:
            self._append_bot_message(
                u"A request is running — attach more files once it finishes.",
                icon=_ICON_SYNC, icon_color=_ICON_SLATE)
            return 0
        added, skipped = 0, []
        for path in (paths or []):
            try:
                if not path or not os.path.isfile(path):
                    continue
                if HAS_RAG and not is_supported(path):
                    skipped.append(os.path.basename(path))
                    continue
                if path not in self._attached_files:
                    self._attached_files.append(path)
                    self._add_attachment_chip(path)
                    added += 1
            except Exception:
                pass
        if added:
            self._refresh_attachment_panel()
        if skipped:
            self._append_bot_message(
                u"Skipped {} unsupported file(s) ({}) — only PDF and "
                u"images are accepted.".format(len(skipped), u", ".join(skipped[:5])),
                icon=_ICON_SYNC, icon_color=_ICON_SLATE)
        return added

    def _input_preview_keydown(self, sender, e):
        """Ctrl+V with files/image on the clipboard → attach thay vì paste text."""
        try:
            from System.Windows.Input import Key, Keyboard, ModifierKeys
            if e.Key != Key.V:
                return
            if (Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control:
                return
            if self._paste_from_clipboard():
                e.Handled = True
        except Exception as ex:
            logger.debug(u"_input_preview_keydown error: {}".format(_exc_text(ex)))

    def _paste_from_clipboard(self):
        """Attach clipboard files (copy từ Explorer) hoặc ảnh clipboard. UI THREAD.

        False for a text paste, so the TextBox keeps its normal behaviour.
        """
        from System.Windows import Clipboard
        try:
            if Clipboard.ContainsFileDropList():
                files = [p for p in Clipboard.GetFileDropList()]
                return self._attach_paths(files) > 0
        except Exception:
            pass
        # Raw bitmap (PrtScn / Snipping Tool)
        try:
            if Clipboard.ContainsImage():
                img = Clipboard.GetImage()
                if img is not None:
                    path = self._save_clipboard_image(img)
                    if path:
                        return self._attach_paths([path]) > 0
        except Exception as ex:
            logger.debug(u"clipboard image paste error: {}".format(_exc_text(ex)))
        return False

    def _save_clipboard_image(self, bmp_source):
        """Write a clipboard BitmapSource as PNG vào folder attachments hôm nay.

        Saved in the dated archive folder so archive_attachments sees src == dst.
        """
        try:
            import time as _time
            from System.Windows.Media.Imaging import PngBitmapEncoder, BitmapFrame
            from System.IO import FileStream, FileMode
            from config.project_store import ProjectStore
            ps = ProjectStore()
            dest = ps.attachments_dir(ps.get_active_project_id())
            path = os.path.join(
                dest, u"pasted_{}.png".format(_time.strftime('%H%M%S')))
            enc = PngBitmapEncoder()
            enc.Frames.Add(BitmapFrame.Create(bmp_source))
            fs = FileStream(path, FileMode.Create)
            try:
                enc.Save(fs)
            finally:
                fs.Close()
            return path
        except Exception as ex:
            logger.debug(u"_save_clipboard_image error: {}".format(_exc_text(ex)))
            return None

    def _file_drag_over(self, sender, e):
        """Show the Copy cursor for file drags anywhere over the window."""
        try:
            from System.Windows import DragDropEffects, DataFormats
            if e.Data.GetDataPresent(DataFormats.FileDrop):
                e.Effects = DragDropEffects.Copy
                e.Handled = True
        except Exception:
            pass

    def _file_drop(self, sender, e):
        """Drop files anywhere on the window → attach."""
        try:
            from System.Windows import DataFormats
            if not e.Data.GetDataPresent(DataFormats.FileDrop):
                return
            files = e.Data.GetData(DataFormats.FileDrop)
            if self._attach_paths(list(files)):
                e.Handled = True
                try:
                    self.chat_input.Focus()
                except Exception:
                    pass
        except Exception as ex:
            logger.debug(u"_file_drop error: {}".format(_exc_text(ex)))

    def clear_attachments_clicked(self, sender, e):
        """Remove all attachments."""
        self._attached_files = []
        try:
            self.attachment_chips_panel.Children.Clear()
            self._refresh_attachment_panel()
        except Exception:
            pass

    def _add_attachment_chip(self, file_path):
        """Add a small chip label for an attached file."""
        try:
            from System.Windows.Controls import Button, StackPanel, TextBlock
            from System.Windows import Thickness
            import os as _os

            name = _os.path.basename(file_path)

            btn = Button()
            try:
                btn.Style = self.FindResource('AttachChipBtn')
            except Exception:
                pass
            btn.Margin = Thickness(0, 0, 4, 4)

            sp = StackPanel()
            sp.Orientation = System.Windows.Controls.Orientation.Horizontal

            icon_lbl = TextBlock()
            icon_lbl.Text = _ICON_ATTACH + u" "
            icon_lbl.FontFamily = System.Windows.Media.FontFamily(u"Segoe MDL2 Assets")
            icon_lbl.FontSize = 11
            icon_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
            sp.Children.Add(icon_lbl)

            name_lbl = TextBlock()
            name_lbl.Text = name if len(name) <= 22 else name[:19] + u"..."
            name_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
            sp.Children.Add(name_lbl)

            x_lbl = TextBlock()
            x_lbl.Text = u"  ✕"
            x_lbl.FontSize = 9
            x_lbl.Foreground = System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(150, 150, 150))
            x_lbl.VerticalAlignment = System.Windows.VerticalAlignment.Center
            sp.Children.Add(x_lbl)

            btn.Content = sp
            btn.ToolTip = file_path

            _path = file_path

            def _on_remove(s, ev, p=_path):
                if p in self._attached_files:
                    self._attached_files.remove(p)
                try:
                    self.attachment_chips_panel.Children.Remove(s)
                except Exception:
                    pass
                self._refresh_attachment_panel()

            btn.Click += _on_remove
            self.attachment_chips_panel.Children.Add(btn)
        except Exception as ex:
            logger.debug(u"_add_attachment_chip error: {}".format(_exc_text(ex)))

    def _refresh_attachment_panel(self):
        """Show or hide the attachment preview border depending on file list."""
        try:
            self.attachment_preview_border.Visibility = (
                Visibility.Visible if self._attached_files else Visibility.Collapsed
            )
        except Exception:
            pass

    # ─── Chat input ───────────────────────────────────────────────────────────

    def send_clicked(self, sender, e):
        if self._busy:
            # Stop; the legacy path still finishes its current step.
            self._request_stop()
            return
        self._process_input()

    def input_keydown(self, sender, e):
        from System.Windows.Input import Key, Keyboard, ModifierKeys

        if self._slash_open:
            if e.Key == Key.Up:
                self._slash_move(-1)
                e.Handled = True
                return
            if e.Key == Key.Down:
                self._slash_move(1)
                e.Handled = True
                return
            if e.Key in (Key.Return, Key.Enter, Key.Tab):
                self._slash_accept()
                e.Handled = True
                return
            if e.Key == Key.Escape:
                self._close_skills_popup()
                e.Handled = True
                return

        if (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control and e.Key == Key.N:
            self.new_chat_clicked()
            e.Handled = True
            return

        if e.Key == Key.Return or e.Key == Key.Enter:
            if (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift:
                caret = self.chat_input.CaretIndex
                text = self.chat_input.Text or ""
                self.chat_input.Text = text[:caret] + "\n" + text[caret:]
                self.chat_input.CaretIndex = caret + 1
                e.Handled = True
                return

            self._process_input()
            e.Handled = True
        elif e.Key == Key.Up:
            # Don't hijack Up while editing a multi-line draft.
            if u"\n" in (self.chat_input.Text or u""):
                return
            if self._input_history:
                if self._history_index == -1:
                    self._current_input_temp = self.chat_input.Text
                    self._history_index = len(self._input_history) - 1
                else:
                    self._history_index = max(0, self._history_index - 1)
                self.chat_input.Text = self._input_history[self._history_index]
                self.chat_input.CaretIndex = len(self.chat_input.Text)
                e.Handled = True
        elif e.Key == Key.Down:
            if self._input_history and self._history_index != -1:
                if self._history_index == len(self._input_history) - 1:
                    self.chat_input.Text = self._current_input_temp
                    self._history_index = -1
                else:
                    self._history_index += 1
                    self.chat_input.Text = self._input_history[self._history_index]
                self.chat_input.CaretIndex = len(self.chat_input.Text)
                e.Handled = True

    # ─── Claude-style composer: slash-skills popup + project/model chips ──────

    def input_text_changed(self, sender, e):
        """Placeholder visibility + slash-skills popup filtering. UI THREAD."""
        try:
            text = self.chat_input.Text or u""
            self.composer_placeholder.Visibility = (
                Visibility.Collapsed if text else Visibility.Visible)
            m = re.match(r'^/([\w\-]*)$', text)
            if m is not None:
                self._refresh_skills_popup(m.group(1))
            elif self._slash_open:
                self._close_skills_popup()
        except Exception as ex:
            logger.debug(u"input_text_changed error: {}".format(_exc_text(ex)))

    def _refresh_skills_popup(self, query):
        """(Re)build the skills popup filtered by `query`. UI THREAD."""
        try:
            from Intelligence.skills_engine import get_skills_engine
            skills = [s for s in get_skills_engine().all_skills()
                      if s.get('enabled')]
        except Exception:
            skills = []
        q = (query or u"").lower()
        items = [s for s in skills
                 if q in s['id'].lower() or q in (s['name'] or u"").lower()]
        # Built-in commands (not skills), listed for discoverability: /memory
        # routes to _try_memory_command. Prefix test, not "q in word" (that
        # surfaced Memory for /e, /r and a bare /).
        if u'memory'.startswith(q):
            items.append({'id': u'memory', 'name': u'Memory',
                          'description': (u'View or manage what the '
                                          u'assistant remembers'),
                          'source': None, 'enabled': True})
        # /skills → _try_skills_command.
        if u'skills'.startswith(q):
            items.append({'id': u'skills', 'name': u'Skills',
                          'description': (u'Install or update skills from a '
                                          u'GitHub repo'),
                          'source': None, 'enabled': True})
        # Tool Hub opens LLMs Setting on its tab instead of inserting a command.
        if u'toolhub'.startswith(q):
            items.append({'id': u'toolhub', 'name': u'Tool Hub',
                          'description': (u'Open Tool Hub: sign in, sync, '
                                          u'your tools, search'),
                          'source': None, 'enabled': True,
                          'opens': u'toolhub'})
        if u'save-tool'.startswith(q) and q:
            items.append({'id': u'save-tool', 'name': u'Save tool',
                          'description': (u'Save the last turn\'s tool calls '
                                          u'as a Tool Hub recipe draft'),
                          'source': None, 'enabled': True})
        if not items:
            self._close_skills_popup()
            return
        self._slash_items = items
        self._slash_rows = []
        self._slash_sel = 0
        panel = self.skills_popup_panel
        panel.Children.Clear()
        for i, meta in enumerate(items):
            row = self._make_skill_row(i, meta)
            self._slash_rows.append(row)
            panel.Children.Add(row)
        self._slash_highlight()
        self.skills_popup.IsOpen = True
        self._slash_open = True

    def _make_skill_row(self, idx, meta):
        """One row of the slash-skills popup. UI THREAD."""
        subtitle = meta.get('description') or u""
        src = {'builtin': u"built-in", 'user': u"yours",
               'project': u"project"}.get(meta.get('source'))
        if src:
            subtitle = (subtitle + u"  ·  " + src) if subtitle else src

        def _click(s, ev, _i=idx):
            self._slash_sel = _i
            self._slash_accept()
            ev.Handled = True

        row = self._popup_row(u"/" + meta['id'], subtitle or None,
                              hover=False, handler=_click)

        def _enter(s, ev, _i=idx):
            self._slash_sel = _i
            self._slash_highlight()
        row.MouseEnter += _enter
        return row

    def _slash_highlight(self):
        """Paint the highlighted popup row. UI THREAD."""
        from System.Windows.Media import Brushes, SolidColorBrush, Color
        for i, row in enumerate(self._slash_rows):
            if i == self._slash_sel:
                _bind_bg(row, 'SelectedBg')
            else:
                row.Background = Brushes.Transparent

    def _slash_move(self, delta):
        """Move popup selection with Up/Down (wraps). UI THREAD."""
        if not self._slash_items:
            return
        self._slash_sel = (self._slash_sel + delta) % len(self._slash_items)
        self._slash_highlight()
        try:
            self._slash_rows[self._slash_sel].BringIntoView()
        except Exception:
            pass

    def _slash_accept(self):
        """Insert the highlighted skill as '/id ' into the input. UI THREAD."""
        try:
            meta = self._slash_items[self._slash_sel]
        except Exception:
            self._close_skills_popup()
            return
        self._close_skills_popup()
        if meta.get('opens'):
            self.chat_input.Text = u""
            self._open_llm_settings(tab=meta['opens'])
            return
        self.chat_input.Text = u"/" + meta['id'] + u" "
        self.chat_input.CaretIndex = len(self.chat_input.Text)
        self.chat_input.Focus()

    def _close_skills_popup(self):
        self._slash_open = False
        self._slash_items = []
        self._slash_rows = []
        try:
            self.skills_popup.IsOpen = False
        except Exception:
            pass

    def _apply_forced_skill(self):
        """Inject the /slash-forced skill into the dispatcher decision.

        Idempotent. skill_forced=True bypasses filter_for_specialist. The
        specialist follows the skill's `agents:` list: a bare "/skill-id" routes
        boilerplate whose "MODIFY" matched the action-verb table → revit_action.
        """
        sid = getattr(self, '_forced_skill_id', None)
        if not sid:
            return
        if not self._agent_decision:
            self._agent_decision = {'specialist': 'general',
                                    'source': 'slash', 'confidence': 1.0}
        self._agent_decision['skill'] = sid
        self._agent_decision['skill_forced'] = True
        try:
            from Intelligence.skills_engine import get_skills_engine
            _cur = self._agent_decision.get('specialist')
            _spec = get_skills_engine().specialist_for(sid, _cur)
            if _spec and _spec != _cur:
                logger.debug("slash /{}: specialist {} -> {} "
                             "(skill frontmatter)".format(sid, _cur, _spec))
                self._agent_decision['specialist'] = _spec
                self._agent_decision['source'] = 'slash'
        except Exception as _sp_ex:
            logger.debug(u"specialist_for error: {}".format(_exc_text(_sp_ex)))

    def _popup_row(self, title, subtitle=None, icon=None, active=False,
                   dot=None, hover=True, handler=None):
        """Claude-style popup menu row (used by all composer popups). UI THREAD."""
        from System.Windows.Controls import (Grid, ColumnDefinition,
                                             StackPanel, TextBlock, Border)
        from System.Windows.Shapes import Ellipse
        from System.Windows import (Thickness, CornerRadius, GridLength,
                                    GridUnitType, TextTrimming,
                                    VerticalAlignment)
        from System.Windows.Media import (Brushes, SolidColorBrush, Color,
                                          FontFamily)
        from System.Windows.Input import Cursors

        row = Border()
        row.CornerRadius = CornerRadius(8)
        row.Padding = Thickness(10, 7, 10, 7)
        row.Background = Brushes.Transparent
        row.Cursor = Cursors.Hand

        g = Grid()
        c0 = ColumnDefinition(); c0.Width = GridLength.Auto
        c1 = ColumnDefinition(); c1.Width = GridLength(1, GridUnitType.Star)
        c2 = ColumnDefinition(); c2.Width = GridLength.Auto
        g.ColumnDefinitions.Add(c0)
        g.ColumnDefinitions.Add(c1)
        g.ColumnDefinitions.Add(c2)

        lead = None
        if dot is not None:
            lead = Ellipse()
            lead.Width = 8
            lead.Height = 8
            lead.Fill = SolidColorBrush(Color.FromRgb(dot[0], dot[1], dot[2]))
            lead.VerticalAlignment = VerticalAlignment.Center
            lead.Margin = Thickness(0, 0, 8, 0)
        elif icon:
            lead = TextBlock()
            lead.Text = icon
            lead.FontFamily = FontFamily(u"Segoe MDL2 Assets")
            lead.FontSize = 11
            _bind_fg(lead, 'Muted')
            lead.VerticalAlignment = VerticalAlignment.Center
            lead.Margin = Thickness(0, 1, 8, 0)
        if lead is not None:
            Grid.SetColumn(lead, 0)
            g.Children.Add(lead)

        body = StackPanel()
        t = TextBlock()
        t.Text = title
        t.FontSize = 12
        t.FontFamily = FontFamily(u"Hanken Grotesk")
        t.FontWeight = System.Windows.FontWeights.SemiBold
        _bind_fg(t, 'Ink')
        t.TextTrimming = TextTrimming.CharacterEllipsis
        body.Children.Add(t)
        if subtitle:
            st = TextBlock()
            st.Text = subtitle
            st.FontSize = 10.5
            st.FontFamily = FontFamily(u"Hanken Grotesk")
            _bind_fg(st, 'Muted')
            st.TextTrimming = TextTrimming.CharacterEllipsis
            st.Margin = Thickness(0, 1, 0, 0)
            body.Children.Add(st)
        Grid.SetColumn(body, 1)
        g.Children.Add(body)

        if active:
            chk = TextBlock()
            chk.Text = u""
            chk.FontFamily = FontFamily(u"Segoe MDL2 Assets")
            chk.FontSize = 11
            _bind_fg(chk, 'Blue')
            chk.VerticalAlignment = VerticalAlignment.Center
            chk.Margin = Thickness(10, 0, 0, 0)
            Grid.SetColumn(chk, 2)
            g.Children.Add(chk)

        row.Child = g

        if hover:
            def _enter(s, ev):
                _bind_bg(row, 'IconHoverBg')

            def _leave(s, ev):
                row.Background = Brushes.Transparent
            row.MouseEnter += _enter
            row.MouseLeave += _leave
        if handler is not None:
            row.PreviewMouseLeftButtonDown += handler
        return row

    def _popup_separator(self):
        from System.Windows.Controls import Border
        from System.Windows import Thickness
        from System.Windows.Media import SolidColorBrush, Color
        sep = Border()
        sep.Height = 1
        _bind_bg(sep, 'Divider')
        sep.Margin = Thickness(6, 5, 6, 5)
        return sep

    def _update_composer_chips(self):
        """Refresh the project + model chips of the composer. UI THREAD."""
        from System.Windows.Media import SolidColorBrush, Color
        try:
            from config.project_store import ProjectStore
            ps = ProjectStore()
            pid = ps.get_active_project_id()
            meta = ps.get_project(pid) if pid else None
            if meta:
                self.project_chip_text.Text = meta.get('name') or u"Project"
                self.project_chip_text.Foreground = SolidColorBrush(
                    _theme_color('Accent'))
            else:
                self.project_chip_text.Text = u"Project"
                self.project_chip_text.Foreground = SolidColorBrush(
                    _theme_color('Muted'))
        except Exception as ex:
            logger.debug(u"_update_composer_chips project error: {}".format(_exc_text(ex)))
        try:
            from Intelligence.llm_router import LLMRouter
            router = LLMRouter()
            self._chip_router_ver = router.get_state_version()
            _label = router.get_display_label()
            _cost = getattr(self, '_cost_hint', u"")
            if _cost:
                _label = u"{} \u00b7 {}".format(_label, _cost)
            self.model_chip_text.Text = _label
            rgb = self._BADGE_COLORS.get(router.get_active_name(),
                                         self._BADGE_GRAY)
            self.model_chip_dot.Fill = SolidColorBrush(
                Color.FromRgb(rgb[0], rgb[1], rgb[2]))
        except Exception as ex:
            logger.debug(u"_update_composer_chips model error: {}".format(_exc_text(ex)))

    def _sync_model_chip_if_stale(self):
        """Re-render the chips when the provider/model changed elsewhere. UI THREAD.

        The router is session-wide; LLMs Setting can switch provider while the
        docked pane stays open (chip said "LOCAL" while the popup ticked Claude).
        One int compare when nothing changed.
        """
        try:
            from Intelligence.llm_router import LLMRouter
            ver = LLMRouter().get_state_version()
            if ver != getattr(self, '_chip_router_ver', None):
                self._update_composer_chips()
        except Exception as ex:
            logger.debug(u"_sync_model_chip_if_stale error: {}".format(_exc_text(ex)))

    def project_chip_clicked(self, sender, e):
        """XAML handler: AssistantProjects.project_chip_clicked."""
        return self._projects.project_chip_clicked(sender, e)

    # ─── Assistant Projects: GUI/assistant/projects.py (self._projects) ───

    def project_panel_close_clicked(self, sender, e):
        """XAML handler: AssistantProjects.project_panel_close_clicked."""
        return self._projects.project_panel_close_clicked(sender, e)
    def _open_active_project_folder(self):
        """Open the active project's knowledge folder (files/) in Explorer."""
        d = u""
        try:
            from config.project_store import ProjectStore
            ps = ProjectStore()
            pid = ps.get_active_project_id()
            if not pid:
                self._append_bot_message(
                    u"No active project — pick one from the project chip first.",
                    icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                return
            d = os.path.join(ps.project_dir(pid), 'files')
            # shared opener: carries the .NET 8 UseShellExecute fallback
            if _open_in_explorer(d):
                return
        except Exception as ex:
            logger.debug(u"_open_active_project_folder error: {}".format(_exc_text(ex)))
        try:
            self._append_bot_message(
                u"Couldn't open the knowledge folder:\n`{}`".format(d),
                icon=_ICON_WARNING, icon_color=_ICON_AMBER)
        except Exception:
            pass

    # ─── Assistant Projects: GUI/assistant/projects.py (self._projects) ───

    def model_chip_clicked(self, sender, e):
        """Open the Claude-style provider/model picker popup."""
        try:
            self._sync_model_chip_if_stale()
            self._build_model_popup()
            self._keep_popup_in_pane(self.model_popup)
            self.model_popup.IsOpen = True
        except Exception as ex:
            logger.debug(u"model_chip_clicked error: {}".format(_exc_text(ex)))

    def _build_model_popup(self, _skip_refresh=False):
        """Fill the model popup from the instant router snapshot. UI THREAD.

        Lists only READY providers (configured remote / reachable local) plus
        the active one; a background probe updates Ollama / LM Studio, which
        the instant snapshot reports unavailable until probed.
        """
        from Intelligence.llm_router import LLMRouter
        router = LLMRouter()
        status = router.get_status_instant()
        panel = self.model_popup_panel
        panel.Children.Clear()
        shown = 0
        for name in router.get_provider_names():
            info = status.get(name, {})
            if not info.get('available') and not info.get('active'):
                continue                       # hide "Not set up" providers
            model = info.get('model') or u""
            subtitle = model if model else (
                None if info.get('available')
                else u"Not configured — open Settings")
            # flag a local model too small for reliable tool-calling
            if model and name in ("ollama", "lmstudio"):
                try:
                    from Intelligence import local_llm as _ll
                    if not _ll.is_tool_capable_size(model):
                        subtitle = u"{} · small, may misfire tools".format(model)
                except Exception:
                    pass
            panel.Children.Add(self._popup_row(
                info.get('display_name', name), subtitle,
                active=info.get('active', False),
                dot=self._BADGE_COLORS.get(name, self._BADGE_GRAY),
                handler=(lambda s, ev, _n=name:
                         self._model_popup_select(_n))))
            shown += 1
        if shown == 0:
            panel.Children.Add(self._popup_row(
                u"No providers ready - open Settings to configure",
                None, icon=u"", hover=False,
                handler=lambda s, ev: self._model_popup_settings()))
        if not _skip_refresh:
            self._refresh_model_popup_async(status)
        panel.Children.Add(self._popup_separator())
        panel.Children.Add(self._popup_row(
            u"Configure API key & model…", None, icon=u"",
            handler=lambda s, ev: self._model_popup_settings()))

    def _refresh_model_popup_async(self, prev_status):
        """Probe local providers off-thread; rebuild the popup once if their
        Ready status changed (else a just-started engine stays hidden)."""
        if getattr(self, '_model_popup_refreshing', False):
            return
        self._model_popup_refreshing = True

        def _work():
            changed = False
            try:
                from Intelligence.llm_router import LLMRouter
                router = LLMRouter()
                names = router.get_provider_names()
                for name in ('ollama', 'lmstudio'):
                    if name not in names:
                        continue
                    before = bool((prev_status.get(name) or {}).get('available'))
                    try:
                        info = router.probe_provider(name)
                    except Exception:
                        info = None
                    if bool((info or {}).get('available')) != before:
                        changed = True
            except Exception:
                changed = False

            def _rebuild():
                self._model_popup_refreshing = False
                try:
                    if (changed and getattr(self, 'model_popup', None) is not None
                            and self.model_popup.IsOpen):
                        self._build_model_popup(_skip_refresh=True)
                except Exception:
                    pass
            try:
                self.Dispatcher.Invoke(Action(_rebuild))
            except Exception:
                self._model_popup_refreshing = False

        t = Thread(ThreadStart(_work))
        t.IsBackground = True
        t.Start()

    def _model_popup_select(self, name):
        try:
            self.model_popup.IsOpen = False
        except Exception:
            pass
        try:
            from Intelligence.llm_router import LLMRouter
            if LLMRouter().get_active_name() == name:
                return
            self._switch_provider(name)
        except Exception as ex:
            logger.debug(u"_model_popup_select error: {}".format(_exc_text(ex)))

    def _model_popup_settings(self):
        try:
            self.model_popup.IsOpen = False
        except Exception:
            pass
        self._open_llm_settings()

    # ─── Harness: action mode (auto / confirm-first) + activity log ──

    @staticmethod
    def _action_mode():
        """'auto' | 'confirm' (settings agents.action_mode). Never raises."""
        try:
            from config.settings import get_settings
            mode = get_settings().get_action_mode()
        except Exception:
            mode = 'auto'
        return mode if mode in ('auto', 'confirm') else 'auto'

    # Chip tooltips. Enforced in code (ToolGate in Intelligence/agent_loop.py):
    # a promise the pane keeps, not a hint to the model.
    _MODE_TIP_AUTO = (
        u"Auto — the assistant applies model edits right away. Destructive "
        u"actions (delete, purge, ungroup, sync) still ask first, and a "
        u"multi-step plan that edits the model is shown for approval before "
        u"it runs.\nClick to switch to Ask before edits.")
    _MODE_TIP_CONFIRM = (
        u"Ask before edits — every call that would change the model waits "
        u"for your Confirm on a card in the chat. Reading the model and "
        u"exports (PDF, DWG, images) never ask.\nClick to switch to Auto.")

    def action_mode_clicked(self, sender, e):
        """Toggle between 'auto' (act immediately) and 'confirm' (ask first)."""
        try:
            from config.settings import get_settings
            settings = get_settings()
            new_mode = ('confirm' if self._action_mode() == 'auto'
                        else 'auto')
            settings.set_agent_option('action_mode', new_mode)
            # the chip shows the new state; no chat bubble (it piled up)
            self._update_action_mode_chip()
        except Exception as ex:
            logger.debug(u"action_mode_clicked error: {}".format(_exc_text(ex)))

    def _update_action_mode_chip(self):
        """Render the action-mode chip (shield = ask, zap = auto). UI THREAD."""
        try:
            from System.Windows.Media import Geometry
            confirm = (self._action_mode() == 'confirm')
            _SHIELD = (u"M20 13c0 5-3.5 7.5-7.66 8.95a1 1 0 0 1-.67-.01"
                       u"C7.5 20.5 4 18 4 13V6a1 1 0 0 1 1-1c2 0 4.5-1.2 "
                       u"6.24-2.72a1 1 0 0 1 1.52 0C14.5 3.8 17 5 19 5a1 1 0 0 1 1 1z")
            _ZAP = (u"M4 14a1 1 0 0 1-.78-1.63l9.9-10.2a.5.5 0 0 1 .86.46"
                    u"l-1.92 6.02A1 1 0 0 0 13 10h7a1 1 0 0 1 .78 1.63"
                    u"l-9.9 10.2a.5.5 0 0 1-.86-.46l1.92-6.02A1 1 0 0 0 11 14z")
            # label collapses on a narrow pane (_apply_narrow_layout); tooltip stays
            try:
                self.action_mode_btn.ToolTip = (
                    self._MODE_TIP_CONFIRM if confirm else self._MODE_TIP_AUTO)
            except Exception:
                pass
            # bound, not copied: follows a Revit Light/Dark switch
            if confirm:
                self.action_mode_text.Text = u"Ask before edits"
                _bind_fg(self.action_mode_text, 'Accent')
                self.action_mode_icon.Data = Geometry.Parse(_SHIELD)
                _bind_stroke(self.action_mode_icon, 'Blue')
            else:
                self.action_mode_text.Text = u"Auto"
                _bind_fg(self.action_mode_text, 'Muted')
                self.action_mode_icon.Data = Geometry.Parse(_ZAP)
                _bind_stroke(self.action_mode_icon, 'Muted')
        except Exception as ex:
            logger.debug(u"_update_action_mode_chip error: {}".format(_exc_text(ex)))

    def activity_log_clicked(self, sender, e):
        """Open today's activity log (notepad) or the logs folder (Explorer).

        Never Process.Start the .md directly: without a markdown association it
        throws Win32Exception (button looked dead).
        """
        try:
            from config.project_store import ProjectStore
            ps = ProjectStore()
            pid = ps.get_active_project_id()
            path = ps.activity_log_path(pid)
            folder = os.path.dirname(path)
            if os.path.isfile(path):
                if _open_file_shell(path):
                    return
            elif _open_in_explorer(folder):
                return
            raise RuntimeError(u"shell refused {}".format(path))
        except Exception as ex:
            logger.error("activity_log_clicked error: {}".format(_exc_text(ex)))
            try:
                self._append_bot_message(
                    u"Could not open the activity log.",
                    icon=_ICON_WARNING, icon_color=_ICON_AMBER)
            except Exception:
                pass

    def _log_activity(self, text):
        """Append to today's activity log on a pool thread (never the UI
        thread: slow disk / AV scan would freeze it). Never raises."""
        try:
            from System.Threading import ThreadPool, WaitCallback

            def _write(_state):
                try:
                    from config.project_store import ProjectStore
                    ps = ProjectStore()
                    ps.append_activity(text, ps.get_active_project_id())
                except Exception:
                    pass
            ThreadPool.QueueUserWorkItem(WaitCallback(_write))
        except Exception:
            pass

    # ─── Assistant Commands: GUI/assistant/commands.py (self._commands) ───

    def _post_bot(self, msg, icon, color):
        """Finish a deterministic turn: show `msg`, record it, release busy."""
        def _show():
            self._hide_typing_indicator()
            self._append_bot_message(msg, icon=icon, icon_color=color)
            self._add_to_history('assistant', msg)
            self._set_busy(False)
        self.Dispatcher.Invoke(Action(_show))

    def _process_input(self):
        """Read input (+ any attachments), dispatch to NLP or keyword fallback."""
        try:
            raw = self.chat_input.Text.strip()
            attached = list(self._attached_files)   # snapshot

            # Read-and-clear: "Try again" echo suppression is ONE-SHOT (left set,
            # an early return below would swallow the next real message).
            suppress_echo = self._suppress_user_echo
            self._suppress_user_echo = False

            if not raw and not attached:
                self._retry_target_row = None
                return

            # this turn's language steers every string emitted (cards, notices)
            _note_user_language(raw)

            # ── Concurrency: queue instead of reject ──
            # BEFORE the slash parse: a queued "/skill" must not overwrite
            # _forced_skill_id while the current request is still routing.
            if self._busy:
                # A plan card awaiting Run/Edit/Cancel has run nothing: a new
                # message supersedes it and goes through once the worker lets go.
                _plan_wait = getattr(self, '_pending_plan', None)
                if _plan_wait is not None and (raw or attached):
                    _plan_wait.decide(_PLAN_SUPERSEDED)
                if attached:
                    self._append_bot_message(
                        u"Still working on the previous request — attachments "
                        u"can be sent again once it finishes.",
                        icon=_ICON_SYNC, icon_color=_ICON_SLATE)
                elif raw:
                    self._queue_input(raw)
                return

            self._remove_quick_replies()
            # fresh tool group: this turn's cards never fold into the last one's
            self._tool_cards._reset_tool_group()

            # ── Slash-skill: "/skill-id [rest]" (unknown ids pass through) ──
            self._close_skills_popup()
            self._forced_skill_id = None
            self._forced_skill_args = u""
            route_text = raw
            m = re.match(r'^/([\w\-]+)(?:\s+(.*))?$', raw, re.S) if raw else None
            if m:
                try:
                    from Intelligence.skills_engine import get_skills_engine
                    _cat = get_skills_engine().get_catalog(enabled_only=False)
                    if any(s['id'] == m.group(1) for s in _cat):
                        self._forced_skill_id = m.group(1)
                        route_text = (m.group(2) or u"").strip()
                        # user's own words, so the router tells scan from fix
                        self._forced_skill_args = route_text
                        if not route_text:
                            # Bare "/skill-id" reads as gibberish (the model
                            # guesses a target), so send an explicit
                            # instruction. "Follow the workflow in your system
                            # prompt": native and JSON-intent paths both inject
                            # the skill body there ("Apply the playbook" gave
                            # {"intent": "apply_playbook"} → "Tool does not exist").
                            # Three modes (a bare invocation has no target):
                            #  * reference (no `tools:`): summarise + ask; told
                            #    to scan, a small model replayed last turn's calls.
                            #  * modifying: read-only diagnosis, then a plan
                            #    (/annotation-standard once dimensioned every
                            #    plan unprompted — declared tools say what a
                            #    skill CAN reach, never that a bare call fires them).
                            #  * read-only: scan the whole project now.
                            _eng = get_skills_engine()
                            _sid = m.group(1)
                            if _eng.is_reference_skill(_sid):
                                route_text = (
                                    u"Answer from the '{}' reference (the "
                                    u"Active skill in your system prompt) "
                                    u"for THIS message only. Summarise the "
                                    u"rules it defines and ask me for the "
                                    u"input it needs (file/sheet list, "
                                    u"project code...). Do NOT call any "
                                    u"Revit tool and do NOT repeat or "
                                    u"re-run any earlier request in this "
                                    u"conversation."
                                ).format(_sid)
                            elif _eng.modifies_model(_sid):
                                route_text = (
                                    u"Follow the '{}' workflow (the Active "
                                    u"skill in your system prompt) for THIS "
                                    u"message. I gave you NO target, so run "
                                    u"its READ/diagnostic steps only — over "
                                    u"the ENTIRE project unless the workflow "
                                    u"says otherwise — then STOP and present "
                                    u"the plan: what you would change, on "
                                    u"which targets, how many, and anything "
                                    u"you still need me to decide. Do NOT "
                                    u"create, modify, delete, rename, tag, "
                                    u"dimension, override or export anything "
                                    u"in this turn; wait for my confirmation. "
                                    u"Ignore every earlier request in this "
                                    u"conversation — do not repeat it."
                                ).format(_sid)
                            else:
                                route_text = (
                                    u"Follow the '{}' workflow (the Active "
                                    u"skill in your system prompt) NOW, "
                                    u"using its DEFAULT scope. This workflow "
                                    u"only reads the model, so start "
                                    u"scanning the ENTIRE project "
                                    u"immediately — do NOT ask about scope. "
                                    u"Ignore every earlier request in this "
                                    u"conversation — do not repeat it."
                                ).format(_sid)
                except Exception:
                    pass

            if raw:
                if not self._input_history or self._input_history[-1] != raw:
                    self._input_history.append(raw)
                    if len(self._input_history) > 50:
                        self._input_history.pop(0)
                self._history_index = -1
                self._current_input_temp = ""

            self.chat_input.Text = ""
            self._last_raw = raw or u"[attached documents]"

            # ── Show user message (not on "Try again": already on screen) ──
            display_text     = raw
            attachment_note  = summarize_attachments(attached) if attached else None
            if not suppress_echo:
                # drop a stale retry target (it'd capture an unrelated answer)
                self._retry_target_row = None
                self._append_user_message(display_text,
                                          attachment_note=attachment_note)
            # A retry records the prompt again on purpose: roles must alternate
            # (Anthropic rejects two assistant turns in a row).
            history_text = (u"{}\n[attached: {}]".format(display_text, attachment_note)
                            if attachment_note else display_text)
            self._add_to_history("user", history_text)

            if self._forced_skill_id:
                try:
                    self._append_skill_chips([self._forced_skill_id])
                except Exception:
                    pass

            if attached:
                self._attached_files = []
                try:
                    self.attachment_chips_panel.Children.Clear()
                    self._refresh_attachment_panel()
                except Exception:
                    pass

            # Lock UI + typing indicator IMMEDIATELY; routing (PDF extraction,
            # Ollama probe, NLU scoring: ms to SECONDS) runs on a worker thread.
            self._set_busy(True)

            _rid = getattr(self, '_request_id', 0)

            def _route(_rt=route_text):
                try:
                    # archive attachments + journal off the UI thread (big PDFs)
                    _files = attached
                    if _files:
                        try:
                            from config.project_store import ProjectStore
                            _ps = ProjectStore()
                            _pid = _ps.get_active_project_id()
                            _files = _ps.archive_attachments(_files, _pid)
                            _ps.append_activity(
                                u"Attached ({}): {}".format(
                                    len(_files),
                                    u", ".join(os.path.basename(x)
                                               for x in _files)), _pid)
                        except Exception:
                            _files = attached
                    _req = _rt if _rt else u"[attached documents]"
                    _skill_note = (u"  _(skill: /{})_".format(
                        self._forced_skill_id)
                        if self._forced_skill_id else u"")
                    self._log_activity(u"User: {}{}".format(
                        _req, _skill_note))
                    self._route_input(_rt, _files)
                except Exception as ex:
                    # always say something (it used to fail silently)
                    if self._cancelled():
                        self._finish_cancelled(_rid)
                    else:
                        self._report_error(u"_route_input", ex, _rid)
                finally:
                    # Belt and braces if a path leaked busy. Does NOT claim the
                    # turn (that would suppress a legitimate finisher's message).
                    if self._busy and _rid == getattr(self, '_request_id', 0):
                        self._ui_invoke(lambda: self._set_busy(False))
                    self._clear_turn_model()

            rt = Thread(ThreadStart(_route))
            rt.IsBackground = True
            rt.SetApartmentState(ApartmentState.STA)
            rt.Start()

        except Exception as ex:
            # UI thread: Dispatcher.Invoke runs inline, no deadlock
            self._report_error(u"_process_input", ex)

    def _auto_route_model(self, raw, history, attached, continuation):
        """Pick Haiku / Sonnet / Opus for this turn. WORKER THREAD.

        Intelligence.model_tier grades the request with one fast call; unpinned
        calls of the turn use the pick. A clarifying answer keeps the previous
        turn's model. Never raises; on failure the saved/default model runs.
        """
        self._turn_triage = None
        try:
            from Intelligence.llm_router import LLMRouter
            from Intelligence import model_tier
            router = LLMRouter()
            if not model_tier.is_enabled(router.get_active_name()):
                return
            provider = router.get_active_provider()
            if provider is None or not hasattr(provider, 'set_turn_model'):
                return
            prev = getattr(self, '_prev_turn_triage', None)
            if continuation and prev:
                decision = dict(prev)
                decision['source'] = 'carryover'
            else:
                self._safe_update_typing_text(
                    u"● ● ●  Đang phân tích yêu cầu…" if _is_viet_text(raw)
                    else u"● ● ●  Reading your request…")
                images = [p for p in (attached or []) if is_image(p)]
                decision = model_tier.triage(
                    provider, raw, history=history,
                    revit_context=getattr(self, '_ctx_line', u''),
                    specialist=(self._agent_decision or {}).get('specialist'),
                    has_images=bool(images))
            if not decision:
                return
            self._turn_model_token = provider.set_turn_model(decision['model'])
            self._turn_triage = decision
            self._prev_turn_triage = decision
            name = model_tier.short_name(decision['model'])
            self._log_activity(u"Auto model: {} ({}, {}){}".format(
                name, decision['tier'], decision['source'],
                u" — " + decision['brief'] if decision.get('brief') else u""))
            self._ui_invoke(lambda: self._show_turn_model(name), wait=False)
        except Exception as ex:
            logger.debug(u"_auto_route_model error: {}".format(_exc_text(ex)))

    def _clear_turn_model(self):
        """End of turn: drop the auto-picked model, put the chip back."""
        token = getattr(self, '_turn_model_token', None)
        if token is None:
            return
        self._turn_model_token = None
        try:
            from Intelligence.llm_router import LLMRouter
            provider = LLMRouter().get_provider('claude')
            if provider is not None:
                provider.clear_turn_model(token)
        except Exception:
            pass
        self._ui_invoke(self._update_composer_chips, wait=False)

    def _show_turn_model(self, name):
        """Chip reads 'Auto · Sonnet 5.5' while the turn runs. UI THREAD."""
        try:
            self.model_chip_text.Text = u"Auto · {}".format(name)
        except Exception:
            pass

    def _provider_unreachable_note(self, viet=True):
        """Message for 'a provider is selected but its health probe failed'.

        u"" when no provider is configured (stock setup guidance applies).
        WORKER THREAD.
        """
        try:
            from Intelligence.llm_router import LLMRouter
            provider = LLMRouter().get_active_provider()
            if provider is None:
                return u""
            label = (getattr(provider, 'DISPLAY_NAME', None)
                     or getattr(provider, 'NAME', u'AI'))
            err = u""
            try:
                err = provider.get_last_error() or u""
            except Exception:
                err = u""
            if viet:
                msg = (u"**{}** đang được chọn nhưng không kết nối được, nên "
                       u"tôi phải trả lời ở chế độ offline (chỉ hiểu lệnh cụ "
                       u"thể).\n\nKiểm tra: API key còn hạn/còn credit, mạng "
                       u"công ty hoặc proxy có chặn không, rồi bấm **Cài đặt → "
                       u"LLMs Setting** để thử lại kết nối.").format(label)
            else:
                msg = (u"**{}** is selected but unreachable, so I answered in "
                       u"offline mode (specific commands only).\n\nCheck the "
                       u"API key (valid, still has credit) and whether a "
                       u"corporate proxy is blocking it, then re-test the "
                       u"connection in **Settings → LLMs Setting**.").format(label)
            if err:
                msg += u"\n\n`{}`".format(err[:200])
            return msg
        except Exception as ex:
            logger.debug(u"_provider_unreachable_note error: {}".format(_exc_text(ex)))
            return u""

    # ─── Assistant Documents: GUI/assistant/documents.py (self._docs) ───

    def _route_input(self, raw, attached):
        """Classify + dispatch one user request. WORKER THREAD.

        Routing ladder (RAG context → fast Revit-DB answer → learned patterns
        → offline NLU → LLM turn) runs off the UI thread; UI feedback and
        _execute_result marshal back through the Dispatcher.
        """
        try:
            attached = list(attached or [])
            has_attach = bool(attached) and HAS_RAG

            # Deterministic commands (no LLM), before learned patterns / NLU
            # can hijack the wording: /memory + "remember that ...", /train.
            if not attached and raw and self._commands._try_memory_command(raw):
                return

            if not attached and raw and self._commands._try_train_command(raw):
                return

            # "which model are you using?" — left to the LLM, "model bạn đang
            # dùng là gì" was read as a question about the .rvt.
            if not attached and raw and self._try_which_model(raw):
                return

            # /skills + "cài skill từ <repo link>" (never re-read as a Revit
            # command), /toolhub (team catalog on Supabase, network only).
            if not attached and raw and self._commands._try_skills_command(raw):
                return

            if not attached and raw and self._commands._try_toolhub_command(raw):
                return

            # Pasted URL / UNC / local path → fetched and pushed onto
            # `attached`, so downstream treats it like a dropped file. Runs
            # AFTER the skills command, which owns its own GitHub URL.
            _link_ctx = u""
            if HAS_LINKS and raw:
                _link_ctx = self._docs._resolve_message_links(raw, attached)
                has_attach = bool(attached) and HAS_RAG

            # PDF markup comments (R4): CommentAgent when the user asks about
            # comments with a PDF attached, or drops an annotated PDF with no
            # text. No annotations → normal document analysis.
            if HAS_AGENTS and has_attach:
                try:
                    _pdfs = [p for p in attached if is_pdf(p)]
                    if _pdfs:
                        from Intelligence.comments import pdf_annots as _pa
                        _annotated = [p for p in _pdfs
                                      if _pa.has_annotations(p)]
                        _kw_dec = AgentDispatcher().classify(
                            raw, allow_llm=False)
                        _wants = (_kw_dec.get('specialist') == 'comment'
                                  or (_annotated and not (raw or u'').strip()))
                        if _annotated and _wants:
                            if self._docs._run_comment_agent(
                                    _annotated[0],
                                    list(self._conversation_history[:-1])):
                                return
                except Exception as _cm_ex:
                    logger.debug(u"comment route error: {}".format(_exc_text(_cm_ex)))

            if has_attach and not raw:
                raw = u"Analyze and summarize the attached document."

            # RAG v2: large PDFs are indexed and only top-k excerpts (with page
            # citations) injected; small files/images keep full-text/vision.
            rag_context = ''
            if has_attach:
                rag_context = self._docs._build_attachment_context(raw, attached)
            # Link notes (folder listings, unreadable links) have no file.
            if _link_ctx:
                rag_context = (u"{}\n\n{}".format(_link_ctx, rag_context)
                               if rag_context else _link_ctx)

            # NLP routing uses ONLY the raw text: prepending the ContextScout
            # summary leaked "sheet"/"view"/"wall" into intent triggers, and
            # do_nlp() already sends it in the system prompt.
            captured = raw

            # One language analysis per turn, shared by dispatcher, graph
            # planner and prompt hint so every stage agrees.
            _utterance = None
            if HAS_LANGUAGE and raw:
                try:
                    _utterance = _lang_analyzer.analyze(raw)
                except Exception as _ln_ex:
                    logger.debug("language analysis error: {}".format(_ln_ex))

            history  = list(self._conversation_history[:-1])

            use_local        = HAS_NLP and has_local_llm()
            use_claude       = HAS_NLP and has_api_key()   # True for any configured provider
            _active_provider = get_active_provider_name()  # "claude" | "openai" | "ollama"

            # A /slash skill must reach the agent with its playbook — learned
            # patterns / NLU would hijack it (e.g. "/english-spellcheck").
            _skill_forced = bool(getattr(self, '_forced_skill_id', None))

            # EXCEPT /english-spellcheck SCANS: the agent sees one
            # ai_element_filter page (50 notes) yet reported "4495 scanned".
            # Scans go to _run_spellcheck (reads EVERY string); fix/apply
            # wording still reaches the agent (set_parameter / rename_element).
            if _skill_forced and self._forced_skill_id == 'english-spellcheck':
                _args = getattr(self, '_forced_skill_args', u'') or u''
                _fix_req = routing.wants_spellcheck_fix(_args)
                if not _fix_req:
                    def _run_spell():
                        self._execute_result({'intent': 'check_spelling',
                                              'params': {}})
                    self.Dispatcher.Invoke(Action(_run_spell))
                    return

            # Carryover: a short reply ("1", "ok") to a clarifying question
            # continues the same skill/specialist instead of re-routing.
            _prev_dec = getattr(self, '_prev_agent_decision', None)
            _continuation = False
            if not _skill_forced and not has_attach and _prev_dec:
                try:
                    _last_bot = u''
                    for _h in reversed(history):
                        if _h.get('role') == 'assistant':
                            _last_bot = u'{}'.format(_h.get('content') or u'')
                            break
                    # A keyword command hit is a NEW command even after a
                    # clarifying question ("tô vàng sàn" after "tô đỏ tường"
                    # replayed the finished red-walls task).
                    _fresh_kw = False
                    if HAS_AGENTS:
                        _fresh = AgentDispatcher().classify(
                            raw, allow_llm=False)
                        _fresh_kw = bool(_fresh
                                         and _fresh.get('source') == 'keyword')
                    _continuation = routing.is_continuation(
                        _last_bot, raw, _prev_dec, fresh_keyword_hit=_fresh_kw)
                except Exception:
                    _continuation = False
            if _continuation:
                logger.debug("carryover: continuing {} / {}".format(
                    _prev_dec.get('specialist'), _prev_dec.get('skill')))

            # ── 1/2. Deterministic: NLU, then learned patterns ──
            # Learned mappings go stale (renamed tools kept winning on a
            # Jaccard >= 0.8 match), so routing.learned_pattern_wins() arbitrates.
            nlu_result = None
            learned = None
            _deterministic = (HAS_NLP and not has_attach and not _skill_forced
                              and not _continuation)
            if _deterministic:
                nlu_result = parse_command_nlu(captured, history)
                learned = find_learned_match(raw)

            if learned and routing.learned_pattern_wins(learned, nlu_result,
                                                        raw=raw):
                def _run_learned(_r=learned):
                    self._execute_result(_r)
                self.Dispatcher.Invoke(Action(_run_learned))
                return

            if _deterministic and nlu_result \
                    and nlu_result.get("intent") not in (None, "unknown"):
                # _authoritative = answered from the real tool catalog; the LLM
                # must not override it. _instant = pure small talk, canned even
                # with a provider online (biggest latency win on a slow local
                # model); emotional chat is not _instant and reaches the LLM.
                if nlu_result["intent"] not in ("chat", "help") \
                        or nlu_result.get("_authoritative") \
                        or nlu_result.get("_instant") \
                        or not (use_local or use_claude):
                    # Configured provider that failed its health probe: the
                    # canned "connect an AI in Settings" is wrong — say why.
                    if nlu_result["intent"] in ("chat", "help") \
                            and not (use_local or use_claude):
                        _note = self._provider_unreachable_note(
                            viet=_is_viet_text(raw))
                        if _note:
                            nlu_result = dict(nlu_result)
                            nlu_result["message"] = _note
                    def _run_nlu(_r=nlu_result):
                        self._execute_result(_r)
                    self.Dispatcher.Invoke(Action(_run_nlu))
                    return

            # ── 2.5 Specialist dispatch (multi-agent layer) ──
            # Keywords always; one tiny LLM classify only for non-chat messages
            # keywords couldn't place. Kill switch: agents.multi_agent.
            self._agent_decision = None
            if HAS_AGENTS and not has_attach and (use_local or use_claude):
                try:
                    from config.settings import get_settings as _get_settings
                    _settings = _get_settings()
                    _multi_on = _settings.is_multi_agent_enabled()
                    _llm_clf_on = _settings.is_llm_classify_enabled()
                except Exception:
                    _multi_on, _llm_clf_on = True, True
                if _multi_on:
                    # Classify the USER's words: a slash skill routes its args,
                    # never the synthetic boilerplate (that put every bare
                    # "/skill-id" on the action specialist); empty text →
                    # specialist from skill frontmatter (_apply_forced_skill).
                    _clf_text = raw
                    if _skill_forced:
                        _clf_text = (getattr(self, '_forced_skill_args', u'')
                                     or u'').strip()
                    _clf_provider = None
                    _is_chat_msg = bool(
                        nlu_result and nlu_result.get("intent") == "chat")
                    if _llm_clf_on and not _is_chat_msg and _clf_text:
                        try:
                            from Intelligence.llm_router import LLMRouter as _LLMR
                            _clf_provider = _LLMR().get_active_provider()
                        except Exception:
                            _clf_provider = None
                    _skills_eng = None
                    try:
                        from Intelligence.skills_engine import SkillsEngine
                        _skills_eng = SkillsEngine()
                    except Exception:
                        pass
                    try:
                        self._agent_decision = AgentDispatcher().classify(
                            _clf_text, provider=_clf_provider,
                            skills_engine=_skills_eng,
                            allow_llm=bool(_clf_provider),
                            # only when classifying the user's own words
                            analysis=(_utterance if _clf_text == raw
                                      else None))
                    except Exception as _disp_ex:
                        logger.debug(u"dispatcher error: {}".format(_exc_text(_disp_ex)))
                    if self._agent_decision:
                        logger.debug("dispatcher: {} ({}, {:.2f})".format(
                            self._agent_decision.get('specialist'),
                            self._agent_decision.get('source'),
                            self._agent_decision.get('confidence', 0.0)))
                    # /slash skill beats whatever the dispatcher matched
                    self._apply_forced_skill()
                    if HAS_KNOWLEDGE and self._agent_decision and \
                            self._agent_decision.get('specialist') == 'knowledge':
                        if self._docs._run_knowledge_agent(raw, history):
                            return
                        # no hits / LLM mute → continue down the normal path
                    if self._agent_decision and \
                            self._agent_decision.get('specialist') == 'comment':
                        _guide = (u"To process drawing comments: attach the "
                                  u"marked-up PDF (attach button) and resend. "
                                  u"I will read each comment, trace which "
                                  u"model its sheet belongs to, propose a "
                                  u"fix, and execute it once you confirm.")

                        def _need_pdf(_g=_guide):
                            self._hide_typing_indicator()
                            self._append_bot_message(
                                _g, icon=_ICON_SYNC, icon_color=_ICON_SLATE)
                            self._add_to_history("assistant", _g)
                            self._set_busy(False)
                        self.Dispatcher.Invoke(Action(_need_pdf))
                        return

            # Also when the dispatcher never ran (kill switch, attachments,
            # no provider).
            self._apply_forced_skill()

            # Carryover: the short reply matched nothing — reuse last decision.
            if _continuation and not (self._agent_decision
                                      and self._agent_decision.get('skill')):
                self._agent_decision = dict(_prev_dec)
                self._agent_decision['source'] = 'carryover'

            # For the next turn's carryover check.
            if self._agent_decision and (
                    self._agent_decision.get('skill')
                    or self._agent_decision.get('specialist')
                    not in (None, 'general')):
                self._prev_agent_decision = dict(self._agent_decision)

            # ── 2.6 Auto model switch (Claude) ──
            # Haiku picks Haiku/Sonnet/Opus for the rest of the turn; cleared
            # in _process_input's finally.
            if use_claude:
                self._auto_route_model(raw, history, attached, _continuation)

            # ── 2.7 Graph path (multi-goal turns) ──
            # "thống kê tường rồi xuất pdf sheet A-101" is TWO requests; the
            # single-specialist path dropped the second. The planner declines
            # what it cannot split confidently, so one-goal turns are unchanged.
            # Kill switch: agents.graph_mode.
            if (not has_attach and not _skill_forced and not _continuation
                    and (use_local or use_claude)):
                try:
                    if self._plans._run_graph_plan(raw, history,
                                            utterance=_utterance):
                        return
                except Exception as _gp_ex:
                    logger.debug("graph path error: {}".format(_gp_ex))

            if use_local or use_claude or has_attach:
                # ── 3/4. LLM path (typing indicator already showing) ──
                nlu_hint = nlu_result if (HAS_NLP and not has_attach) else None

                def do_nlp():
                    result = None
                    # A later turn bumps _request_id and invalidates stale finishers.
                    rid = getattr(self, '_request_id', 0)
                    # "model returned nothing" (timeout/empty) vs "unrecognised
                    # intent" — otherwise both showed the same offline fallback.
                    llm_call_failed = False
                    from Intelligence.llm_router import LLMRouter
                    import json as _json

                    _router = LLMRouter()
                    _provider = _router.get_active_provider()

                    # ── Native function-calling agent path ──
                    # SUPPORTS_NATIVE_TOOLS providers: schemas via the API
                    # `tools` param, plain-text replies. Returns False if it
                    # can't start (no registry, provider mute on turn 1, e.g.
                    # Ollama without tools) → legacy JSON-intent loop below.
                    # Compact RAG excerpts stay native; only bulky stuffing goes
                    # legacy. Images go native on a vision provider — the legacy
                    # path sends only the "[Hình ảnh đính kèm: x.png]" string.
                    _img_files = [p for p in attached if is_image(p)]
                    _can_see = bool(_img_files) and bool(
                        getattr(_provider, 'SUPPORTS_VISION', False))
                    _compact_rag = bool(
                        rag_context and len(rag_context) < 4000
                        and (not _img_files or _can_see))

                    # Text-only provider (Ollama / LM Studio / DeepSeek): tell
                    # the model why it can't see the image and what to suggest.
                    _rag_query = rag_context
                    if _img_files and not _can_see:
                        _note = (
                            u"[System: the user attached {} image(s): {}. The "
                            u"active model ({}) has no vision capability, so "
                            u"you cannot see them. Say this plainly in the "
                            u"user's language, tell them to switch to Claude "
                            u"or OpenAI in the model picker to analyse "
                            u"images, then answer whatever else you can.]"
                        ).format(
                            len(_img_files),
                            u", ".join(os.path.basename(p) for p in _img_files),
                            getattr(_provider, 'DISPLAY_NAME', None)
                            or getattr(_provider, 'NAME', u'local'))
                        _rag_query = (u"{}\n\n{}".format(_note, rag_context)
                                      if rag_context else _note)

                    if (not has_attach and not rag_context) or _compact_rag:
                        _handled = False
                        # None = legacy behaviour.
                        _spec = None
                        _skill_ids = None
                        _dec = getattr(self, '_agent_decision', None)
                        if HAS_SPECIALISTS and _dec:
                            # AgentLoop specialists ('knowledge'/'comment' have
                            # own pipelines above; 'general' = _spec None).
                            if _dec.get('specialist') in (
                                    'revit_data', 'revit_action',
                                    'multi_doc', 'modeling',
                                    'qa_check', 'export'):
                                _spec = get_spec(_dec['specialist'])
                            if _dec.get('skill'):
                                # build_skills_block injects up to two: send
                                # the top-2 ranked; a forced /slash skill
                                # stays alone and skips the specialist filter.
                                _skill_ids = [_dec['skill']]
                                if not _dec.get('skill_forced'):
                                    _skill_ids = self._ranked_skill_ids(
                                        captured, _dec['skill'])
                                if not _dec.get('skill_forced'):
                                    try:
                                        from Intelligence.skills_engine import (
                                            get_skills_engine)
                                        _spec_name = (_spec.name if _spec
                                                      else 'general')
                                        _skill_ids = get_skills_engine() \
                                            .filter_for_specialist(_skill_ids,
                                                                   _spec_name)
                                    except Exception:
                                        pass
                        self._note_turn_route(
                            specialist=(_spec.name if _spec else 'general'),
                            skills=_skill_ids)
                        try:
                            _handled = self._run_native_agent(
                                _provider, list(history), captured,
                                spec=_spec, skill_ids=_skill_ids,
                                rag_context=(rag_context if _compact_rag
                                             else None),
                                image_files=(_img_files if _can_see else None),
                                analysis=_utterance)
                        except Exception as _na_ex:
                            # Also to the activity log: pyRevit's debug log is
                            # off by default, so a crash looked like the legacy
                            # path had simply been chosen.
                            logger.debug(u"native agent path error: {}".format(
                                _exc_text(_na_ex)))
                            self._log_activity(
                                u"native agent path FAILED, falling back to the "
                                u"legacy JSON path — {}: {}".format(
                                    type(_na_ex).__name__,
                                    _exc_text(_na_ex))[:600])
                        if _handled:
                            return

                    # Skill playbook for the LEGACY path too — without it the
                    # model hallucinated intents like "apply_playbook".
                    _legacy_skills_block = u""
                    try:
                        _dec_lg = getattr(self, '_agent_decision', None) or {}
                        _sk_lg = _dec_lg.get('skill') or getattr(
                            self, '_forced_skill_id', None)
                        if _sk_lg:
                            from Intelligence.skills_engine import (
                                build_skills_block as _bsb)
                            # Same top-2 as the native path.
                            _ids_lg = ([_sk_lg]
                                       if getattr(self, '_forced_skill_id', None)
                                       else self._ranked_skill_ids(captured, _sk_lg))
                            _legacy_skills_block = _bsb(_ids_lg)
                    except Exception:
                        _legacy_skills_block = u""

                    # Step budget: the specialist's own (native path gives
                    # 6–14); the old 5 made an ISO 19650 name check give up.
                    max_iterations = 10
                    try:
                        _spec_budget = getattr(_spec, 'max_iterations', None)
                        if _spec_budget:
                            max_iterations = int(_spec_budget)
                    except Exception:
                        pass
                    current_iteration = 0
                    current_history = list(history)
                    # Repeat guards: small local models re-emit calls in the
                    # same turn ("tô xanh sàn" ran 3x) or replay a previous
                    # turn's ("tô tím cửa"); prompt rules alone don't stop it.
                    _turn_calls = set()          # (intent, args) run THIS turn
                    _prev_calls = self._earlier_turn_calls()
                    _blocked_repeats = 0
                    _announce_nudges = 0         # "I'm about to…" with no call
                    # ONE ToolGate per request, as on the native path
                    # (destructive confirm, "Ask before edits", first-purge dry
                    # run). Built on the first tool call (needs the server).
                    _legacy_gate = {"g": None}

                    # _rag_query == rag_context unless a text-only provider got images.
                    current_query = (_rag_query + u"\n\n" + captured) if _rag_query else captured

                    while current_iteration < max_iterations:
                        current_iteration += 1

                        # Honour Stop before another LLM turn (it once kept
                        # running write tools after Stop).
                        if self._cancelled():
                            self._finish_cancelled(rid)
                            return

                        _ctx_block = u""
                        if HAS_SCOUT:
                            try:
                                _ctx_block = ContextScout.get_context_summary_for_ai()
                            except Exception:
                                pass

                        # Compact catalog (C4): name + one-line purpose; full
                        # schemas (~75 tools) crushed small models, so the model
                        # asks for one via the describe_tool meta-intent.
                        server_tools_str = u""
                        tools_list = []
                        try:
                            from core.server import get_t3labai_server
                            srv = get_t3labai_server()
                            tools_list = srv._handle_tools_list().get('tools', [])
                            # Local models: essential subset only (as native path).
                            if tools_list and getattr(_provider, 'NAME', '') in ("ollama", "lmstudio"):
                                from Intelligence.tool_schema import ESSENTIAL_TOOL_NAMES
                                _ess = [t for t in tools_list
                                        if t.get('name') in ESSENTIAL_TOOL_NAMES]
                                tools_list = _ess or tools_list
                            if tools_list:
                                server_tools_str = u"\n\nLocal MCP Server Tools (name: purpose):\n"
                                for tool in tools_list:
                                    desc = (tool.get('description') or u'').strip()
                                    desc = desc.splitlines()[0][:110] if desc else tool['name']
                                    server_tools_str += u"- `{}`: {}\n".format(
                                        tool['name'], desc)
                                server_tools_str += (
                                    u"\nTool calls need correct parameters. If you are "
                                    u"unsure of a tool's parameters, FIRST reply exactly "
                                    u'{"intent": "describe_tool", "params": {"name": "<tool_name>"}} '
                                    u"to receive its full JSON schema, then call the tool "
                                    u"in your next reply.\n")
                        except Exception as tool_err:
                            logger.debug(u"Failed to list server tools: {}".format(_exc_text(tool_err)))

                        system_prompt = _build_system_prompt(revit_context=_ctx_block)
                        # Project instructions + remembered facts (this path
                        # serves attachments, where conventions matter most).
                        system_prompt = self._projects._apply_project_blocks(system_prompt)
                        if server_tools_str:
                            system_prompt += server_tools_str
                        if _legacy_skills_block:
                            system_prompt += (
                                u"\n\n" + _legacy_skills_block +
                                u"\n\nIMPORTANT: the Active skill above is a "
                                u"set of INSTRUCTIONS for you — 'playbook', "
                                u"'skill' and 'apply_playbook' are NOT intent "
                                u"or tool names. Execute its steps now using "
                                u"ONLY the intents and MCP tools listed in "
                                u"this prompt.")

                        if has_attach or rag_context:
                            system_prompt = _RAG_SYSTEM_PREFIX + system_prompt

                        # Iteration 1 streams into a bubble; later internal
                        # tool-loop turns use a blocking call.
                        _resp = None
                        try:
                            if current_iteration == 1:
                                _resp = self._stream_llm_turn(
                                    _provider, _router, current_history,
                                    system_prompt, current_query, max_tokens=1200,
                                    response_format={"type": "json_object"}
                                )
                            elif _provider and _provider_usable(_provider):
                                _resp = _provider.chat(_budget_history(current_history, _provider, system_prompt, current_query, 1200), system_prompt, current_query, max_tokens=1200, response_format={"type": "json_object"})
                            else:
                                _resp = _router.chat(_budget_history(current_history, None, system_prompt, current_query, 1200), system_prompt, current_query, max_tokens=1200, response_format={"type": "json_object"})
                        except Exception as chat_ex:
                            logger.debug(u"Router chat error: {}".format(_exc_text(chat_ex)))

                        if not _resp or not _resp.strip():
                            llm_call_failed = True
                            break

                        _cleaned = self._clean_bot_response(_resp)
                        _parsed = None
                        try:
                            import re as _re
                            # The model may wrap the JSON in extra text.
                            _m = _re.search(r'\{[\s\S]*\}', _resp)
                            if _m:
                                _parsed = _json.loads(_m.group())
                            else:
                                _parsed = _json.loads(_resp)
                                
                            if "intent" not in _parsed:
                                _parsed["intent"] = "unknown"
                        except Exception:
                            _parsed = {
                                "intent": "unknown",
                                "message": u"Could not read data from the model. Please try again."
                            }

                        if _parsed and _parsed.get("intent"):
                            intent = _parsed.get("intent")
                            params = _parsed.get("params", {}) or {}
                            message = _parsed.get("message", "")

                            # C4 meta-intent: feed one tool's full schema back;
                            # the real call comes next iteration.
                            if intent == "describe_tool":
                                self.Dispatcher.Invoke(Action(self._remove_stream_bubble))
                                tname = u"{}".format(
                                    (params or {}).get("name")
                                    or (params or {}).get("tool") or u"").strip()
                                match = None
                                for _t in (tools_list or []):
                                    if _t.get('name') == tname:
                                        match = _t
                                        break
                                if match:
                                    info = _json.dumps(
                                        {"name": match['name'],
                                         "description": match.get('description', ''),
                                         "inputSchema": match.get('inputSchema', {})},
                                        ensure_ascii=False)
                                else:
                                    info = _json.dumps(
                                        {"error": "unknown tool", "name": tname},
                                        ensure_ascii=False)
                                current_history.append(
                                    {"role": "assistant",
                                     "content": _json_text(_parsed)})
                                current_history.append(
                                    {"role": "user",
                                     "content": u"Tool schema: {}".format(info)})
                                current_query = (
                                    u"Schema: {}. Now respond with the actual tool call as "
                                    u'{{"intent": "<tool_name>", "params": {{...}}, '
                                    u'"message": "..."}}.'.format(info))
                                continue

                            # Check the real registry, not a "revit_" prefix
                            # (place_wall, create_grid don't have one).
                            is_local_tool = False
                            try:
                                from core.server import get_t3labai_server
                                srv = get_t3labai_server()
                                if intent in srv._tools:
                                    is_local_tool = True
                            except Exception:
                                pass

                            if is_local_tool:
                                _call_key = (intent, _json_text(
                                    params or {}, sort_keys=True))
                                _dup_now = _call_key in _turn_calls
                                # Previous turn's call AFTER this turn did other
                                # work = drift (as a first call it's allowed:
                                # the user may repeat a command).
                                _old_ctx = (not _dup_now and bool(_turn_calls)
                                            and _call_key in _prev_calls)
                                if _dup_now or _old_ctx:
                                    _blocked_repeats += 1
                                    if _blocked_repeats >= 2:
                                        # Model is looping — conclude for it.
                                        result = {"intent": "chat",
                                                  "message": (
                                            u"Đã thực hiện xong yêu cầu của bạn."
                                            if _is_viet_text(captured) else
                                            u"Done — the request has been completed.")}
                                        break
                                    if _dup_now:
                                        _guard_msg = (
                                            u"Tool `{}` with IDENTICAL arguments already ran "
                                            u"successfully in THIS turn — never repeat a "
                                            u"completed call. If the request \"{}\" is "
                                            u"fulfilled, reply with the final JSON now: "
                                            u'{{"intent": "chat", "message": "<short summary>"}}.'
                                            .format(intent, captured))
                                    else:
                                        _guard_msg = (
                                            u"Tool `{}` with these arguments already ran in a "
                                            u"PREVIOUS completed turn — that work is DONE. Act "
                                            u"ONLY on the current request: \"{}\". Reply with "
                                            u"the next NEW tool call for it, or the final "
                                            u'summary JSON {{"intent": "chat", "message": '
                                            u'"..."}} if it is fulfilled.'.format(intent, captured))
                                    current_history.append(
                                        {"role": "assistant",
                                         "content": _json_text(_parsed)})
                                    current_history.append(
                                        {"role": "user", "content": _guard_msg})
                                    current_query = _guard_msg
                                    continue
                                self.Dispatcher.Invoke(Action(self._remove_stream_bubble))
                                is_vn = _is_viet_text(captured)
                                status_msg = u"● ● ●  Đang chạy công cụ `{}`...".format(intent) if is_vn else "● ● ●  Executing tool `{}`...".format(intent)
                                self._safe_update_typing_text(status_msg)

                                # Model's "what I'm doing" line; name/args/result
                                # are on the card below.
                                if message:
                                    self._safe_append_bot(
                                        message, icon=_ICON_SYNC,
                                        icon_color=_ICON_BLUE)

                                # Tool card (as on the native path) shows the
                                # result whatever the model says after — "list
                                # levels" once answered without the levels.
                                _card = {"cur": None}

                                def _open_card(_n=intent, _a=params):
                                    _card["cur"] = self._tool_cards._append_tool_card(_n, _a)

                                try:
                                    self.Dispatcher.Invoke(Action(_open_card))
                                except Exception:
                                    pass

                                import time as _time
                                _t0 = _time.time()
                                tool_result = None
                                try:
                                    from core.server import get_t3labai_server
                                    srv = get_t3labai_server()
                                    if _legacy_gate["g"] is None:
                                        _legacy_gate["g"] = self._tool_cards._new_tool_gate(srv)
                                    tool_result = self._tool_cards._run_gated_tool(
                                        srv, _legacy_gate["g"], intent,
                                        params, _is_viet_text(captured))
                                except Exception as execute_err:
                                    # _exc_text, never str(): non-ASCII Revit
                                    # errors raised UnicodeEncodeError here.
                                    tool_result = {"error": _exc_text(execute_err),
                                                   "tool": intent}

                                _ok = not (isinstance(tool_result, dict)
                                           and tool_result.get('error'))
                                _secs = _time.time() - _t0

                                def _close_card(_ok=_ok, _s=_secs,
                                                _r=tool_result):
                                    self._tool_cards._update_tool_card(_card["cur"], _ok,
                                                           _s, _r)

                                try:
                                    self.Dispatcher.Invoke(Action(_close_card))
                                except Exception:
                                    pass

                                # Errors stay unregistered (rule 4: ONE retry).
                                # A DECLINED call is registered (no re-asking)
                                # but doesn't count as a run step.
                                if not (isinstance(tool_result, dict)
                                        and tool_result.get('error')):
                                    _turn_calls.add(_call_key)
                                    if not (isinstance(tool_result, dict)
                                            and tool_result.get('cancelled')):
                                        self._tool_runs += 1

                                # Checkpoint that makes Stop feel immediate;
                                # returning skips finish() (no double release).
                                if self._cancelled():
                                    self._finish_cancelled(rid)
                                    return

                                _result_text = _json_text(tool_result)
                                current_history.append({"role": "assistant", "content": _json_text(_parsed)})
                                current_history.append({"role": "user", "content": u"Tool `{}` returned: {}".format(intent, _result_text)})

                                # Restate WHICH request: a bare "next step" let
                                # small models replay the previous turn's command.
                                current_query = (
                                    u"Tool `{}` returned: {}. Continue with the "
                                    u"NEXT step of the user's LATEST request "
                                    u"ONLY: \"{}\". Earlier conversation turns "
                                    u"are already completed — never repeat "
                                    u"them. Reply with the next tool call, or "
                                    u"a final summary if this request is "
                                    u"done.".format(
                                        intent,
                                        _result_text,
                                        captured))
                                continue
                            else:
                                # "unknown" with a message (common on small
                                # models): show the answer, not the fallback.
                                if intent == "unknown" and message.strip():
                                    _parsed["intent"] = "chat"
                                # Narrated instead of acted ("Đang liệt kê các
                                # mức…" with no tool run): nudge once.
                                if (self._tool_runs == 0 and not _turn_calls
                                        and _announce_nudges < 1
                                        and _announces_work(
                                            message or _cleaned)):
                                    _announce_nudges += 1
                                    self.Dispatcher.Invoke(Action(
                                        self._remove_stream_bubble))
                                    current_history.append(
                                        {"role": "assistant",
                                         "content": _json_text(_parsed)})
                                    current_history.append(
                                        {"role": "user",
                                         "content": _ANNOUNCE_FIXUP})
                                    current_query = (
                                        u"{}\n\nThe request is: \"{}\"".format(
                                            _ANNOUNCE_FIXUP, captured))
                                    continue
                                result = _parsed
                                break
                        else:
                            result = {"intent": "chat", "message": _cleaned}
                            break

                    # For later turns' old-context guard.
                    if _turn_calls:
                        self._remember_turn_calls(_turn_calls)

                    # Loop exhausted leaves result=None (tool branches
                    # `continue`), which showed "I didn't understand" after
                    # five successful edits. Synthesise only when tools ran.
                    if result is None and self._tool_runs:
                        result = {
                            "intent": "chat",
                            "message": (
                                u"Completed {} step(s), then stopped — this "
                                u"request needs more steps than one turn "
                                u"allows. Tell me what to do next, or split it "
                                u"into smaller requests.".format(self._tool_runs)),
                            "params": {},
                            "_max_steps": True,
                        }

                    def finish():
                        try:
                            if self._cancelled():
                                self._finish_cancelled(rid)
                                return
                            if not self._claim_turn(rid):
                                return
                            has_stream = (self._stream_tb is not None)
                            r_intent   = result.get("intent") if result else None
                            _conv      = ("chat", "help", "greet")

                            # Budget exhausted: one-click Continue, as native.
                            if result and result.get("_max_steps"):
                                if has_stream:
                                    self._remove_stream_bubble()
                                    self._clear_stream_refs()
                                self._hide_typing_indicator()
                                msg = result.get("message", u"")
                                self._append_bot_message(
                                    msg, icon=_ICON_WARNING,
                                    icon_color=_ICON_AMBER)
                                self._add_to_history("assistant", msg)
                                try:
                                    self._append_quick_replies(
                                        [u"Continue"],
                                        [u"continue where you stopped"])
                                except Exception:
                                    pass
                                self._set_busy(False)
                                return

                            # Streamed chat reply → keep the bubble, apply markdown.
                            if has_stream and r_intent in _conv:
                                if r_intent == "help":
                                    msg = result.get("params", {}).get(
                                        "answer", result.get("message", ""))
                                else:
                                    msg = result.get("message", "")
                                msg = self._clean_bot_response(msg) if msg else u""
                                if not msg:
                                    try:
                                        msg = self._stream_tb.Text or u""
                                    except Exception:
                                        msg = u""
                                if not msg:
                                    msg = u"Anything else I can help with?"
                                self._finalize_stream_bubble(msg)
                                self._add_to_history("assistant", msg)
                                self._clear_stream_refs()
                                self._set_busy(False)
                                return

                            # Action result → drop the preview; the executor renders.
                            if has_stream:
                                self._remove_stream_bubble()
                            self._hide_typing_indicator()
                            # The NLU's generic fallback only when the LLM did
                            # answer; if it never did, llm_call_failed says why.
                            nlu_hint_usable = (
                                nlu_hint and nlu_hint.get("intent") not in (None, "unknown")
                                and not (llm_call_failed and nlu_hint.get("_generic_fallback"))
                            )
                            if result and result.get("intent") not in (None, "unknown"):
                                self._execute_result(result)
                            elif nlu_hint_usable:
                                self._execute_result(nlu_hint)
                            else:
                                if has_attach and not use_claude and not use_local:
                                    if rag_context:
                                        self._append_bot_message(
                                            u"Document content:\n\n" + rag_context[:2000],
                                            icon=_ICON_ATTACH, icon_color=_ICON_SLATE
                                        )
                                    else:
                                        self._append_bot_message(
                                            u"Could not extract text from the document. The PDF may be scanned."
                                        )
                                    self._set_busy(False)
                                    return
                                fb = keyword_parse(captured)
                                if fb:
                                    self._execute_result(fb)
                                elif llm_call_failed:
                                    # Never answered: say so, with the provider's
                                    # reason (get_last_error) when recorded.
                                    label = get_provider_display_label()
                                    detail = u""
                                    try:
                                        _le = (_provider.get_last_error()
                                               if _provider else None)
                                        if _le:
                                            detail = (u"\nDetails: {}".format(_le)
                                                      if _is_viet_text(captured)
                                                      else u"\nDetail: {}".format(_le))
                                    except Exception:
                                        pass
                                    msg = (u"Model AI ({}) không phản hồi (model quá nặng, "
                                           u"mất kết nối hoặc API báo lỗi). Thử lại hoặc chọn "
                                           u"model khác trong Cài đặt.{}".format(label, detail)
                                           if _is_viet_text(captured) else
                                           u"The AI model ({}) didn't respond (too heavy, "
                                           u"disconnected, or the API returned an error). Try again "
                                           u"or pick another model in Settings.{}".format(label, detail))
                                    self._append_bot_message(msg, icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                                    self._set_busy(False)
                                else:
                                    msg = (u"I did not understand this request — could you describe it in more detail?"
                                           if _is_viet_text(captured) else
                                           "I didn't understand this request — could you describe it more specifically?")
                                    self._append_bot_message(msg)
                                    self._set_busy(False)
                        except Exception as finish_ex:
                            self._clear_stream_refs()
                            # Release _claim_turn so the error can be shown.
                            self._replied = False
                            self._report_error(u"finish", finish_ex, rid)

                    self.Dispatcher.Invoke(Action(finish))

                # Already on the worker thread — run inline.
                do_nlp()
            else:
                # ── 5. No provider configured — keyword fallback ──
                fb = keyword_parse(raw)

                def _finish_offline(_fb=fb):
                    if _fb:
                        self._execute_result(_fb)
                    else:
                        self._append_bot_message(
                            get_setup_guidance_message(_is_viet_text(raw)),
                            icon=_ICON_INFO, icon_color=_ICON_SLATE)
                        self._set_busy(False)
                self.Dispatcher.Invoke(Action(_finish_offline))
        except Exception:
            # To _process_input's handler (hides typing, releases busy on UI).
            raise

    # ─── Execute intent ───

    def _execute_result(self, result):
        """Execute a parsed result dict: show + record the reply, learn the
        pattern, release busy (also after background exports)."""
        intent  = result.get("intent", "unknown")
        message = result.get("message", "")
        params  = result.get("params", {})
        raw     = self._last_raw
        # For 👍/👎 on the reply (_vote_decision).
        self._last_result = dict(result or {})

        def _bot(msg):
            self._append_bot_message(msg)
            self._add_to_history("assistant", msg)

        def _learn(msg=''):
            learn_pattern(raw, intent, params, msg)

        # Last line of defence: small talk never launches a tool, whatever
        # layer (mis-learned pattern, LLM hallucination) produced the intent.
        if intent not in ("help", "chat", "greet", "unknown") and HAS_NLP:
            try:
                from Intelligence.nlu_engine import is_conversational
                _is_smalltalk = is_conversational(raw)
            except Exception:
                _is_smalltalk = False
            if _is_smalltalk:
                conv = parse_command_nlu(raw) or {}
                if conv.get("intent") in ("greet", "chat", "help") and conv.get("message"):
                    reply = conv["message"]
                elif _is_viet_text(raw):
                    reply = (u"Xin chào! Tôi là T3Lab Assistant.\n"
                             u"Hôm nay bạn muốn làm gì?")
                else:
                    reply = (u"Hello! I'm T3Lab Assistant.\n"
                             u"What would you like to do today?")
                _bot(reply)
                self._set_busy(False)
                return

        if intent in ("help", "chat", "greet"):
            reply = params.get("answer", message) if intent == "help" else message
            reply = self._clean_bot_response(reply) if reply else u""
            _bot(reply or u"Anything else I can help with?")
            self._set_busy(False)
            return

        # Both paths construct Revit-bound windows and must run in API context.
        if intent in ("export_direct", "open_batchout_configured"):
            request_id = getattr(self, '_request_id', 0)
            if getattr(self, '_batchout_request_id', None) == request_id:
                return
            # finish() already claimed the turn; hand it to the API callback.
            self._replied = False
            _bot("Export queued. Waiting for Revit..." if intent == "export_direct"
                 else "BatchOut launch queued. Waiting for Revit...")
            expected_doc = self.doc
            self._batchout_request_id = request_id
            self._batchout_running = False

            def _cancel_batchout():
                return (self._cancelled() or request_id != getattr(self, '_request_id', 0))

            def _run_batchout():
                if _cancel_batchout():
                    return False, "BatchOut request stopped before execution. No files were created."
                # resolve_doc(), not revit.doc: docked, revit.doc is None here
                # (see _get_doc_key), so this refused every BatchOut request.
                if (expected_doc is None or not expected_doc.IsValidObject
                        or _active_doc() != expected_doc):
                    return False, "The active document changed or closed. Refresh the Assistant context before retrying."
                self._batchout_running = True
                if intent == "export_direct":
                    if not HAS_EXECUTOR:
                        return False, "The BatchOut executor is unavailable. Reload pyRevit before retrying."
                    ok, count, details = direct_export(_load_batchout_mod(), params,
                                                       cancel_check=_cancel_batchout)
                    return ok, details
                notices = []
                ok = launch_batchout_configured(params, notices.append)
                return ok, ("BatchOut was opened with the requested settings."
                            if ok else (notices[-1] if notices else "Could not open configured BatchOut."))

            def _finished_batchout(ok, details):
                def _finish_ui():
                    if getattr(self, '_batchout_request_id', None) == request_id:
                        self._batchout_request_id = None
                        self._batchout_running = False
                    if not self._claim_turn(request_id):
                        return
                    try:
                        _bot(details or "BatchOut did not return a result. Check Revit before retrying.")
                        if ok:
                            _learn(details)
                    finally:
                        self._set_busy(False)
                self._ui_invoke(_finish_ui)

            if not HAS_API_CONTEXT:
                _finished_batchout(False, "Revit API event support is unavailable. Reload pyRevit and reopen the Assistant.")
                return
            try:
                run_in_api_context(_run_batchout, _finished_batchout, require_api_context=True)
            except Exception as ex:
                _finished_batchout(False, _exc_text(ex))
            return

        # Spell-check: deterministic scan HERE (small models replied "please
        # send me the text notes"), then LLM proofread.
        if intent == "check_spelling":
            confirm = message or (
                u"Đang quét Text Note trong model để kiểm tra chính tả tiếng Anh..."
                if _is_viet_text(raw) else
                u"Scanning model Text Notes for English spelling errors...")
            _bot(confirm)
            _learn(confirm)
            self._spell._run_spellcheck(raw)
            return

        if intent in TOOL_LAUNCHERS:
            confirm = message or u"Opening {}...".format(get_tool_title(intent))
            _bot(confirm)
            _learn(confirm)
            self._launch_tool(intent, TOOL_LAUNCHERS[intent])
            return

        # MCP Revit intents
        try:
            from Intelligence.t3lab_agent import is_mcp_intent, get_intent_info
            if is_mcp_intent(intent):
                cat, desc = get_intent_info(intent)
                param_txt = u""
                if params:
                    param_txt = u"\n**Parameters:** " + u", ".join(
                        u"{}={}".format(k, v) for k, v in params.items() if v is not None
                    )
                reply = (
                    message + u"\n\n"
                    if message else u""
                ) + (
                    u"**Revit API:** `{}`\n"
                    u"**Action:** {}{}\n\n"
                    u"*Requires the [revit-mcp](https://github.com/mcp-servers-for-revit/revit-mcp) "
                    u"server running alongside Revit.*"
                ).format(intent, desc, param_txt)
                _learn(reply)
                _bot(reply)
                self._set_busy(False)
                return
        except Exception as _mcp_ex:
            logger.debug(u"MCP intent handler error: {}".format(_exc_text(_mcp_ex)))

        # Near-miss open_* names ("open_mcp_control"): resolve the user's words
        # against the catalog; launch only on a clear winner.
        if intent.startswith("open_") and HAS_NLP:
            _match = None
            try:
                from Intelligence.nlu_engine import resolve_tool
                _match, _cands = resolve_tool(raw)
            except Exception:
                _match = None
            if _match and _match['intent'] in TOOL_LAUNCHERS:
                label = _match.get('title', _match['intent'])
                confirm = (u"Đang mở {}...".format(label) if _is_viet_text(raw)
                           else u"Opening {}...".format(label))
                _bot(confirm)
                learn_pattern(raw, _match['intent'], {}, confirm)
                self._launch_tool(_match['intent'],
                                  TOOL_LAUNCHERS[_match['intent']])
                return
            # Nothing resolved — name tools that DO exist.
            if intent not in TOOL_LAUNCHERS:
                _bot(self._unknown_tool_text(intent, raw))
                self._set_busy(False)
                return

        # Fallthrough: a named intent here is unregistered anywhere — say so,
        # never "Đã thực hiện." (nothing ran).
        if intent == "unknown":
            _bot(params.get("message", u"The request is not clear — could you describe it in more detail?"))
        elif message:
            _bot(message)
        else:
            _bot(u"Tool `{}` does not exist — check the name or describe what you need.".format(intent))
        self._set_busy(False)

    def _launch_tool(self, intent, launcher):
        """Open a ribbon tool from chat. UI THREAD.

        The launcher runs in a real Revit API context: the dispatcher thread
        while Revit is IDLE is not one, so tools creating an ExternalEvent
        (BCF Reader, ManaLoca, BatchOut) or a Transaction threw. The hop is
        ASYNC (waiting would deadlock this thread's message loop), so busy is
        released at once and the outcome reported later.
        """
        def _report(ok, err):
            # Report success too: an unresolved "Opening X..." reads as a hang.
            if ok:
                self._ui_invoke(
                    lambda: self._append_bot_message(
                        self._launch_success_text(intent),
                        icon=_ICON_SUCCESS, icon_color=_ICON_GREEN))
                return
            self._ui_invoke(
                lambda: self._append_bot_message(
                    self._launch_failure_text(intent, err)))

        try:
            run_in_api_context(launcher, _report)
        except Exception as ex:
            _report(False, _exc_text(ex))
        self._set_busy(False)

    def _launch_success_text(self, intent):
        """Terminal confirmation for a launcher that DID open its window."""
        title = get_tool_title(intent)
        if _ui_viet():
            return u"Đã mở **{}**. Cửa sổ tool đang ở trong Revit.".format(title)
        return u"Opened **{}** — the tool window is now in Revit.".format(title)

    def _launch_failure_text(self, intent, err):
        """Failure message naming the tool and the reason (not "check the console")."""
        title = get_tool_title(intent)
        if err:
            return u"Could not open **{}**: {}".format(title, err)
        return u"Could not open **{}** — the tool reported no reason.".format(title)

    def _unknown_tool_text(self, intent, raw):
        """Message for an intent with no installed tool, listing real
        alternatives from the live catalog."""
        lines = [u"There is no tool called `{}` in this extension.".format(intent)]
        names = []
        try:
            from Intelligence.nlu_engine import resolve_tool
            _m, cands = resolve_tool(raw or u"")
            for c in (cands or [])[:3]:
                t = c.get('title') or c.get('intent')
                if t and t not in names:
                    names.append(t)
        except Exception:
            pass
        if names:
            lines.append(u"Closest matches: {}.".format(
                u", ".join(u"**{}**".format(n) for n in names)))
        else:
            lines.append(u"Type `/` or open the command palette to see what is "
                         u"available.")
        return u"\n\n".join(lines)

    # ─── Assistant Spellcheck: GUI/assistant/spellcheck.py (self._spell) ───


    # ─── Native agentic loop (function calling) ───

    def _ranked_skill_ids(self, text, primary, limit=2):
        """Up to `limit` skill ids, best first; `primary` (the dispatcher's
        pick, maybe LLM-chosen with no keyword match) always leads."""
        ids = [primary] if primary else []
        try:
            from Intelligence.skills_engine import get_skills_engine
            for sid in get_skills_engine().match(text):
                if sid not in ids:
                    ids.append(sid)
                if len(ids) >= limit:
                    break
        except Exception:
            pass
        return ids[:limit]

    # ─── Assistant Documents: GUI/assistant/documents.py (self._docs) ───

    # ─── Assistant Plans: GUI/assistant/plans.py (self._plans) ───

    @staticmethod
    def _analysis_hint(text, analysis=None):
        """Per-message language hint for the volatile context block.

        u'' when unavailable or not needed (no extra tokens). Never raises.
        """
        if not HAS_LANGUAGE or not (text or u'').strip():
            return u''
        try:
            utt = analysis if analysis is not None else _lang_analyzer.analyze(
                text)
            lang = _reply_language()
            return utt.to_prompt_hint(
                lang=(lang if lang in ('vi', 'en') else None))
        except Exception:
            return u''

    def _note_native_skip(self, reason):
        """Record WHY the native path was skipped, in the activity log too
        (pyRevit's debug log is off by default)."""
        msg = u"native path skipped: {}".format(reason)
        logger.debug(msg)
        try:
            self._log_activity(msg[:600])
        except Exception:
            pass

    def _run_native_agent(self, provider, history, captured,
                          spec=None, skill_ids=None, rag_context=None,
                          image_files=None, analysis=None, hold_busy=False):
        """Run the native tool-calling agent loop. WORKER THREAD.

        spec: SpecialistSpec narrowing tools/prompt/budget (None = default).
        rag_context: retrieval excerpts prepended to the user content only.
        image_files: sent as vision blocks (caller checks SUPPORTS_VISION).
        analysis: pre-built language.Utterance for `captured`, reused.
        hold_busy: keep busy ON after this turn — the graph path runs several
        of these for ONE message and must not free the composer mid-plan.

        Returns True when fully handled; False (only if nothing was shown
        yet) lets the legacy JSON-intent path run. The answer is left on
        `self._last_agent_text` for the graph path's next node.
        """
        self._last_agent_text = u""
        if provider is None or not getattr(provider, "SUPPORTS_NATIVE_TOOLS", False):
            self._note_native_skip(
                u"provider {} has no native tools".format(
                    getattr(provider, 'NAME', provider)))
            return False

        from Intelligence.agent_loop import AgentLoop, build_agent_system_prompt
        from Intelligence import tool_schema

        try:
            from core.server import get_t3labai_server
            srv = get_t3labai_server()
        except Exception as _srv_ex:
            self._note_native_skip(u"MCP server unavailable ({})".format(
                _exc_text(_srv_ex)))
            return False

        launcher = tool_schema.make_launcher_tool(list(TOOL_LAUNCHERS.keys()))
        # Local providers get a curated subset: ~110 schemas overflow local
        # context and small models pick worse from a huge catalog.
        _is_local = provider.NAME in ("ollama", "lmstudio")
        _spec_tools = spec.tools_for(_is_local) if spec is not None else None
        _use_launcher = spec.use_launcher if spec is not None else True
        _extras = [launcher] if _use_launcher else []
        # Local tools (handled in _exec_tool, never reach the Revit server):
        # memory, past chats (read-only), Tool Hub search for tools the model
        # cannot see, and published recipes (sha256-verified) as recipe_<slug>.
        _extras.append(tool_schema.make_memory_tool())
        _extras.append(tool_schema.make_past_chats_tool())
        _extras.append(tool_schema.make_find_tools_tool())
        try:
            from Intelligence.toolhub import catalog as _hub_catalog, recipes as _hub_recipes
            _extras.extend(_hub_recipes.recipe_tools(_hub_catalog.load()['rows'].values()))
        except Exception:
            pass
        # Matched skills may declare extra tools — union into the subset.
        if _spec_tools and skill_ids:
            try:
                from Intelligence.skills_engine import get_skills_engine
                _sk_tools = set()
                for _sid in skill_ids:
                    _sk_tools.update(get_skills_engine().tools_for(_sid))
                if _sk_tools:
                    _spec_tools = frozenset(_spec_tools) | _sk_tools
            except Exception:
                pass
        # Claude: defer_loading + server-side tool search (only CLOUD_ESSENTIAL
        # + extras in the prompt); the provider strips it if the API rejects it.
        _defer = False
        if provider.NAME == "claude" and not _is_local:
            try:
                from config.settings import get_settings as _gs
                _defer = bool(_gs().is_tool_search_enabled())
            except Exception:
                _defer = True
        if _spec_tools:
            tools = tool_schema.get_tools_by_names(
                provider.NAME, _spec_tools, _extras, defer=_defer)
        else:
            tools = tool_schema.get_tools_for_provider(
                provider.NAME, _extras, essential_only=_is_local, defer=_defer)
        _deferred_on = any(isinstance(t, dict) and t.get("defer_loading")
                           for t in tools)
        if len(tools) <= len(_extras):   # registry unavailable
            self._note_native_skip(
                u"tool registry empty for provider {} (got {} tools, {} of them "
                u"extras)".format(provider.NAME, len(tools), len(_extras)))
            return False

        ctx = u""
        if HAS_SCOUT:
            try:
                ctx = ContextScout.get_context_summary_for_ai()
            except Exception:
                pass
        # Cached open-model digest: the agent starts grounded instead of
        # spending round-trips re-discovering the project.
        _snap = self._model_snapshot_grounding()
        if _snap:
            ctx = u"{}\n\n{}".format(ctx, _snap) if ctx else _snap

        _proj_instructions, _mem_block = self._projects._project_prompt_blocks()
        _skills_block = u""
        try:
            if skill_ids:
                from Intelligence.skills_engine import build_skills_block
                _skills_block = build_skills_block(skill_ids)
        except Exception:
            pass

        # Answer in the window's language (else Vietnamese cards wrap an
        # English answer).
        _lang = _reply_language()
        if _lang == 'auto':
            _lang = 'vi' if _is_viet_text(captured) else 'en'

        # ── STATIC system prompt ──────────────────────────────────────────────
        # Session-stable, behind ONE prompt-cache breakpoint. Per-turn state
        # goes in _volatile with the user turn, or the cache never hits.
        if spec is not None and HAS_SPECIALISTS:
            # Instructions + skills appended below in build_specialist_prompt's
            # order, without its own "## Project instructions" header.
            system_prompt = build_specialist_prompt(
                spec, project_instructions=u"",
                skills_block=u"", local=_is_local, lang=_lang)
        else:
            system_prompt = build_agent_system_prompt(local=_is_local,
                                                      lang=_lang)
        if _deferred_on:
            # A constant — the system block stays cacheable.
            system_prompt += u"\n\n" + tool_schema.TOOL_CATEGORIES_LINE
        if _proj_instructions:
            system_prompt += u"\n\n" + _proj_instructions
        if _skills_block:
            system_prompt += u"\n\n" + _skills_block

        if _mem_block:
            system_prompt += u"\n\n" + _mem_block

        # Rolling summary of scrolled-out turns: changes only every few turns,
        # so the cache still hits in between.
        try:
            from Intelligence.conversation import summary_block
            _sum_block = summary_block(
                getattr(self, '_history_summary', u''), viet=(_lang == 'vi'))
            if _sum_block:
                system_prompt += u"\n\n" + _sum_block
        except Exception:
            pass

        # ── VOLATILE context (travels with the user turn) ─────────────────────
        # Project-knowledge reference (BM25 on the request) so tool answers
        # follow documented standards; the header tells the model to ignore
        # stray hits. knowledge/comment specialists retrieve on their own.
        _kref = u""
        try:
            _spec_name = spec.name if spec is not None else 'general'
            if _spec_name not in ('knowledge', 'comment'):
                _kref = self._docs._build_knowledge_reference(
                    captured, local=_is_local, history=history) or u""
        except Exception:
            _kref = u""

        # 'confirm' = "Ask before edits": enforced by _run_gated_tool's card, so
        # the model must NOT also ask in text (that made a double confirmation).
        # Fixed per mode, so the system block still caches.
        if self._action_mode() == 'confirm':
            system_prompt += (
                u"\n\n## ACTION MODE: ASK BEFORE EDITS\n"
                u"The user approves every tool call that modifies the model "
                u"on a card in the chat before it runs; the pane holds the "
                u"call for you. So call tools exactly as you otherwise would "
                u"— do NOT ask for confirmation in text first. If a call "
                u"comes back cancelled (the user declined it), do not retry "
                u"it: say what was skipped and ask what to change. Read-only "
                u"tools are never held.")

        viet = _is_viet_text(captured)

        # ── C2: vision ────────────────────────────────────────────────────────
        # Attached images + (for "look at this view") an active-view snapshot,
        # as Claude-format image blocks (openai_provider converts to image_url).
        user_content = captured
        if rag_context:
            user_content = rag_context + u"\n\n" + captured
        # Auto model switch: Haiku's pre-read of this turn, for Sonnet / Opus.
        try:
            from Intelligence.model_tier import brief_note
            _brief = brief_note(getattr(self, '_turn_triage', None))
            if _brief:
                user_content = _brief + u"\n\n" + user_content
        except Exception:
            pass
        _vision_files = list(image_files or [])
        shot = None
        if self._wants_view_snapshot(captured, provider):
            self._safe_update_typing_text(
                u"● ● ●  Đang chụp active view…" if viet
                else u"● ● ●  Capturing the active view…")
            shot = self._capture_active_view(srv)
            if shot:
                _vision_files.append(shot)
        if _vision_files:
            try:
                from Intelligence.rag_processor import build_vision_content_blocks
                user_content = build_vision_content_blocks(
                    user_content, _vision_files)
            except Exception as _vx:
                logger.debug(u"vision block build error: {}".format(_exc_text(_vx)))
        if shot:
            # Now base64 in user_content; %TEMP%\T3Lab_ViewShots kept them all.
            try:
                os.remove(shot)
            except Exception:
                pass

        # Live state rides with THIS turn only; _conversation_history keeps
        # the raw user text, so the transcript does not grow.
        try:
            from Intelligence.agent_loop import (build_context_block,
                                                 apply_context_block)
            _volatile = build_context_block(revit_context=ctx,
                                            knowledge_ref=_kref,
                                            analysis_hint=self._analysis_hint(
                                                captured, analysis))
            user_content = apply_context_block(user_content, _volatile)
        except Exception as _ctx_ex:
            logger.debug(u"context block build error: {}".format(_exc_text(_ctx_ex)))

        # ── B5: tool-execution wrapper ────────────────────────────────────────
        # Every call goes through _run_gated_tool with ONE ToolGate per request
        # (same seam as the legacy loop): destructive tools and, in "Ask before
        # edits", every model-modifying call wait on a card; the first
        # purge_unused is forced to dry_run.
        #
        # B4 (one request = one TransactionGroup) is withdrawn: a group cannot
        # outlive the ExternalEvent.Execute that opened it (Revit discards it,
        # journal: "An API event handler left some transaction phases open"),
        # and its stale handle, Assimilate()d later, crashed Revit. Read the note on
        # __begin_action_group in lib/core/server.py before reinstating it;
        # a multi-step request is now N undo entries.
        gate = self._tool_cards._new_tool_gate(srv)
        # A server anchored before a pyRevit reload may still hold a group.
        try:
            srv._release_stale_action_group()
        except Exception:
            pass

        # ── D1: cancel the loop when the USER switches documents mid-request ──
        # Agent-initiated switch/open/close_document DISARMS the guard instead
        # (deliberate multi-doc work). Must stay above _exec_tool, which
        # closes over it (else NameError at the first tool call).
        doc_guard = {"key": _get_doc_key(), "armed": True}

        _loop_box = {}

        def _find_tools(args):
            """find_tools, local: Tool Hub search -> tool_reference blocks
            (Claude, deferred) or matches + add_tools (other providers)."""
            try:
                from Intelligence.toolhub import find as _hub_find
                from Intelligence.toolhub.client import ToolHubClient
                _lp = _loop_box.get('loop')
                _cur = (_lp._tools if _lp is not None else tools) or []
                _names = [t.get('name') or (t.get('function') or {}).get('name')
                          for t in _cur if isinstance(t, dict)]
                _reg = [{'slug': t['name'], 'description': t.get('description', '')}
                        for t in tool_schema.get_server_tools()]
                _client = ToolHubClient()
                if not _client.signed_in():
                    _client = None
                res, add = _hub_find.run(args.get('query'), args.get('limit'),
                                         _names, _reg, deferred=_deferred_on,
                                         client=_client)
                if add and _lp is not None:
                    _lp.add_tools(tool_schema.tools_named(provider.NAME, add))
                return res
            except Exception as ex:
                return {"error": u"find_tools failed: {}".format(_exc_text(ex))}

        def _exec_tool(name, args):
            args = dict(args or {})
            if name == tool_schema.FIND_TOOLS_NAME:
                return _find_tools(args)
            if name.startswith('recipe_'):
                # Steps re-enter _exec_tool: ToolGate + audit see each call.
                try:
                    from Intelligence.toolhub import catalog as _hc, recipes as _hr
                    _row = _hc.load()['rows'].get(name[len('recipe_'):])
                    if not _row:
                        return {"error": u"Recipe {} is not in the local catalog; "
                                         u"run /toolhub sync.".format(name)}
                    return _hr.run(_row, args, _exec_tool)
                except Exception as ex:
                    return {"error": u"Recipe failed: {}".format(_exc_text(ex))}
            if name == tool_schema.PAST_CHATS_TOOL_NAME:
                # Local and read-only: the archived chats of this scope.
                try:
                    from Intelligence import session_context as _sc
                    _pid = self._projects._active_pid()
                    return _sc.search_past_chats(
                        pid=_pid, query=args.get('query'),
                        session_id=args.get('session_id'),
                        limit=args.get('limit') or 5,
                        doc_key=None if _pid else self._doc_key,
                        exclude_id=getattr(self, '_session_id', None))
                except Exception as _pc_ex:
                    return {"error": u"search_past_chats failed: {}".format(
                        _exc_text(_pc_ex))}
            if name == tool_schema.MEMORY_TOOL_NAME:
                # Local pseudo-tool — persists a fact, never touches Revit.
                try:
                    from Intelligence import assistant_memory
                    _pid = None
                    try:
                        from config.project_store import ProjectStore
                        _pid = ProjectStore().get_active_project_id()
                    except Exception:
                        pass
                    _action = (args.get('action') or 'save').lower()
                    _fact = args.get('fact')
                    _scope = args.get('scope') or 'project'
                    _replaces = args.get('replaces')
                    if _scope == 'session':
                        return self._sessions._session_note_tool(_action, _fact, _replaces)
                    if _action == 'forget':
                        # Fall back to `fact` if the model put the gist there.
                        _ok, _removed = assistant_memory.forget_fact(
                            _replaces or _fact, project_id=_pid)
                        if _ok:
                            return {"success": True,
                                    "note": u"Forgot: {}".format(_removed)}
                        if getattr(self, '_session_notes', None):
                            # Not a long-term fact — maybe a note of this chat.
                            _res = self._sessions._session_note_tool(
                                'forget', _fact, _replaces)
                            if _res.get("success"):
                                return _res
                        return {"error": u"No matching fact to forget."}
                    if _action == 'update':
                        # update_fact saves anew if the old fact isn't found.
                        _ok, _note = assistant_memory.update_fact(
                            _replaces or _fact, _fact,
                            scope=_scope, project_id=_pid)
                        if _ok:
                            return {"success": True, "note": _note}
                        return {"error": _note}
                    _ok, _note = assistant_memory.add_fact(
                        _fact, scope=_scope, project_id=_pid)
                    if _ok:
                        return {"success": True, "note": _note}
                    return {"error": _note}
                except Exception as _mem_ex:
                    return {"error": u"{}".format(_mem_ex)}
            # Destructiveness is declared with the tools (server._DESTRUCTIVE_TOOLS
            # / _DESTRUCTIVE_OPS), read via srv.is_destructive.
            res = self._tool_cards._run_gated_tool(srv, gate, name, args, viet)
            # Deliberate doc change: disarm (activation is async, so retargeting
            # the key could race).
            if name in ('switch_active_document', 'open_document',
                        'close_document'):
                try:
                    if isinstance(res, dict) and 'error' not in res:
                        doc_guard["armed"] = False
                except Exception:
                    doc_guard["armed"] = False
            return res

        def _exec_reads_batch(batch):
            """Run a run of read-only calls in ONE crossing to Revit.

            Reads only (agent_loop.leading_read_run), but a read can still need
            a card (check_bad_geometry deep_probe): then return None and
            AgentLoop runs each call through _exec_tool. (srv.execute_tools_batch
            never existed; the swallowed AttributeError hid that.)
            `batch` is [(name, args), ...]; return a list of the same length.
            """
            from Intelligence.agent_loop import GATE_RUN
            if any(gate.verdict(_name, _args) != GATE_RUN
                   for _name, _args in batch):
                return None
            out = []
            for _name, _args in batch:
                out.append(srv._execute_tool(_name, _args))
            return out

        def _guard_check():
            try:
                if not doc_guard["armed"]:
                    return False
                cur = _get_doc_key()
                # "default" = read failed (unknown, not changed).
                return cur != "default" and cur != doc_guard["key"]
            except Exception:
                return False

        import time as _time
        stream = {"text": u"", "last": 0.0, "open": False}
        card   = {"cur": None}

        # ── Throttled live-stream rendering ───────────────────────────────────
        # ≤ ~25 updates/s via BeginInvoke (worker never blocks); only the
        # state-transition callbacks below use synchronous Invoke.
        def _push_stream(force):
            now = _time.time()
            if not force and (now - stream["last"]) < 0.04:
                return
            stream["last"] = now
            snap = _hide_reasoning(stream["text"])

            def _ui():
                try:
                    if not stream["open"]:
                        # Still "thinking": keep the typing indicator.
                        if not snap:
                            return
                        stream["open"] = True
                        self._hide_typing_indicator()
                        self._begin_stream_bubble()
                    if self._stream_tb is not None:
                        self._stream_tb.Text = snap
                        self._scroll_to_bottom()
                except Exception:
                    pass

            try:
                self.Dispatcher.BeginInvoke(Action(_ui))
            except Exception:
                pass

        def on_text_delta(chunk):
            self._mark_first_token()
            stream["text"] += chunk
            _push_stream(False)

        def on_turn_text(text, is_final):
            text = _hide_reasoning(text)
            stream["text"] = u""
            if not text:
                # Reasoning-only turn: drop any half-open bubble.
                def _drop():
                    try:
                        if stream["open"]:
                            stream["open"] = False
                            self._remove_stream_bubble()
                    except Exception:
                        pass
                self._ui_invoke(_drop, wait=False)
                return

            def _ui():
                try:
                    # "Show thinking" off hides INTERIM narration only; it still
                    # goes to history and the typing indicator.
                    if not is_final and not self._tool_cards._chat_verbosity()[1]:
                        if stream["open"]:
                            stream["open"] = False
                            self._remove_stream_bubble()
                            self._clear_stream_refs()
                        self._show_typing_indicator()
                        if getattr(self, "_typing_text_block", None) is not None:
                            _line = u" ".join(text.split())[:80]
                            self._typing_text_block.Text = u"● ● ●  {}".format(_line)
                        self._add_to_history("assistant", text)
                        return
                    if stream["open"]:
                        stream["open"] = False
                        self._finalize_stream_bubble(text)
                        self._clear_stream_refs()
                    else:
                        self._hide_typing_indicator()
                        self._append_bot_message(text)
                    self._add_to_history("assistant", text)
                except Exception:
                    pass

            try:
                self.Dispatcher.Invoke(Action(_ui))
            except Exception:
                pass

        # "/save-tool" turns this turn's successful calls into a recipe draft.
        self._last_turn_calls = []
        self._last_turn_user_text = captured
        _pending_args = {}

        def on_tool_start(name, args, iteration):
            _pending_args[name] = dict(args or {})
            def _ui():
                try:
                    self._hide_typing_indicator()
                    card["cur"] = self._tool_cards._append_tool_card(name, args)
                    self._show_typing_indicator()
                    if getattr(self, "_typing_text_block", None) is not None:
                        self._typing_text_block.Text = (
                            u"● ● ●  Running `{}`…".format(name))
                except Exception:
                    pass

            try:
                self.Dispatcher.Invoke(Action(_ui))
            except Exception:
                pass

        def on_tool_done(name, result, ok, seconds):
            try:
                self._last_turn_calls.append((name, _pending_args.pop(name, {}), bool(ok)))
            except Exception:
                pass
            def _ui():
                try:
                    self._tool_cards._update_tool_card(card["cur"], ok, seconds, result)
                except Exception:
                    pass
            # Pure UI; queue order keeps it after on_tool_start's card.
            self._ui_invoke(_ui, wait=False)

        # Skill chip (a /slash-forced skill drew it in _process_input).
        _dec_chip = getattr(self, '_agent_decision', None)
        if skill_ids and not (_dec_chip and _dec_chip.get('skill_forced')):
            def _show_chips():
                try:
                    self._append_skill_chips(skill_ids)
                except Exception:
                    pass
            self._ui_invoke(_show_chips, wait=False)

        try:
            from config.settings import get_settings as _gs2
            _compact_on = bool(_gs2().is_compact_tool_results_enabled())
        except Exception:
            _compact_on = True

        def _spill_result(name, result):
            try:
                from Intelligence import result_store
                return result_store.save_result(name, result)
            except Exception:
                return None

        loop = AgentLoop(
            provider, _exec_tool, tools,
            callbacks={
                "on_text_delta": on_text_delta,
                "on_turn_text":  on_turn_text,
                "on_tool_start": on_tool_start,
                "on_tool_done":  on_tool_done,
                "guard_check":   _guard_check,
            },
            max_iterations=(spec.max_iterations if spec is not None else 10),
            # see specialists.CLOUD_MAX_TOKENS
            max_tokens=(spec.tokens_for(_is_local) if spec is not None
                        else (1500 if _is_local else 16000)),
            # Confirm-card waits are excluded by pause()/resume().
            time_budget_sec=(240 if _is_local else 600),
            # Bigger results are shrunk; result_store keeps them in full.
            result_chars=(4000 if _is_local else 32000),
            spill_result=_spill_result,
            # Client-side folding; Claude does context editing in the API.
            compact_results=(provider.NAME != "claude" and _compact_on),
            turn_timer=getattr(self, '_turn_timer', None),
            execute_tools_batch=_exec_reads_batch,
            # NOT srv.is_write_tool — it never existed and the AttributeError
            # silently sent every request to the legacy path.
            is_write_tool=tool_schema.is_model_modifying)
        _loop_box['loop'] = loop

        self._agent_loop = loop
        if self._cancel_requested:
            loop.cancel()
        try:
            result = loop.run(history, system_prompt, user_content)
        finally:
            self._agent_loop = None
            self._commands._toolhub_upload_soon()

        # Provider never answered turn 1 and nothing reached the UI →
        # hand back to the legacy path (it has its own fallbacks).
        if (result.get("status") == "failed"
                and result.get("iterations", 0) <= 1
                and not result.get("text")
                and not result.get("tool_runs")
                and not stream["open"]):
            # Log WHY: the handover is invisible, and without the provider's
            # reason it looks like the model "chose" not to use tools.
            try:
                logger.debug(u"native path: turn 1 failed ({}) -> legacy "
                             u"fallback".format(provider.get_last_error()
                                                or u"no reason recorded"))
            except Exception:
                pass
            return False

        # EMPTY model turn (confused local model; 2026-07-22: a skill chip,
        # then silence) → legacy path, which always replies visibly.
        if (result.get("status") == "done"
                and not (result.get("text") or u"").strip()
                and not result.get("tool_runs")
                and not stream["open"]):
            logger.debug("native path: empty model turn -> legacy fallback")
            return False

        def _finish_ui():
            try:
                # A turn interrupted mid-stream leaves an open live bubble.
                if stream["open"]:
                    stream["open"] = False
                    txt = stream["text"]
                    if txt.strip():
                        self._finalize_stream_bubble(txt)
                        self._add_to_history("assistant", txt)
                    else:
                        self._remove_stream_bubble()
                    self._clear_stream_refs()
                self._hide_typing_indicator()

                st = result.get("status")
                if st == "cancelled":
                    self._append_bot_message(
                        u"Stopped as requested." if viet else u"Stopped.",
                        icon=_ICON_STOP, icon_color=_ICON_SLATE)
                elif st == "doc_changed":
                    self._append_bot_message(
                        (u"You switched to another document — the request was "
                         u"cancelled to avoid editing the wrong model.") if viet else
                        (u"The active document changed — request cancelled "
                         u"to avoid editing the wrong model."),
                        icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                elif st == "failed":
                    label = get_provider_display_label()
                    detail = u""
                    try:
                        _le = provider.get_last_error()
                        if _le:
                            detail = (u"\nDetails: {}".format(_le) if viet
                                      else u"\nDetail: {}".format(_le))
                    except Exception:
                        pass
                    self._append_bot_message(
                        (u"Model AI ({}) bị ngắt giữa chừng — kết quả có thể "
                         u"chưa đầy đủ. Vui lòng thử lại.{}".format(label, detail))
                        if viet else
                        (u"The AI model ({}) dropped mid-request — the result "
                         u"may be incomplete. Please retry.{}".format(label, detail)),
                        icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                    # One-click resend of the exact same request.
                    if self._last_raw and not self._last_raw.startswith(u"["):
                        self._append_quick_replies(
                            [u"Thử lại" if viet else u"Retry"],
                            [self._last_raw])
                elif st in ("max_iterations", "timeout"):
                    self._append_bot_message(
                        (u"The request is too long — stopped after {} steps. Break "
                         u"it into smaller requests to continue.").format(result.get("iterations"))
                        if viet else
                        (u"Request too long — stopped after {} steps. Split it "
                         u"up to continue.").format(result.get("iterations")),
                        icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                    self._append_quick_replies(
                        [u"Tiếp tục" if viet else u"Continue"],
                        [u"tiếp tục phần còn lại" if viet
                         else u"continue where you stopped"])
                elif (st == "done" and not result.get("text")
                        and result.get("tool_runs")):
                    self._append_bot_message(
                        u"Completed {} tool steps.".format(
                            result.get("tool_runs")) if viet else
                        u"Completed {} tool step(s).".format(
                            result.get("tool_runs")),
                        icon=_ICON_SUCCESS, icon_color=_ICON_GREEN)

                # Reply ends with a yes/no approval question → one-click
                # answers; the chip reply rides the carryover so the same
                # specialist/skill resumes.
                if st == "done":
                    _tail = (result.get("text") or u"")[-220:].lower()
                    if u"?" in _tail and any(k in _tail for k in (
                            u"confirm", u"proceed", u"shall i", u"should i",
                            u"go ahead", u"do you want", u"want me to",
                            u"xác nhận", u"đồng ý", u"tiến hành",
                            u"thực hiện", u"ok?", u"okay?")):
                        self._append_quick_replies(
                            [u"Đồng ý, làm đi", u"Không, hủy"] if viet else
                            [u"Yes, go ahead", u"No, cancel"])

                li = result.get("launch_intent")
                if li:
                    # _execute_result opens the window and releases busy.
                    self._execute_result({"intent": li, "message": u"",
                                          "params": {}})
                elif not hold_busy:
                    self._set_busy(False)
            except Exception as ex:
                self._report_error(u"native agent finish", ex)

        self.Dispatcher.Invoke(Action(_finish_ui))
        # For the graph path; a fully streamed turn can leave result["text"] empty.
        self._last_agent_text = ((result.get("text") or u"").strip()
                                 or (stream.get("text") or u"").strip())
        return True

    # ─── Assistant Tool Cards: GUI/assistant/tool_cards.py (self._tool_cards) ───

    # ─── Assistant Spellcheck: GUI/assistant/spellcheck.py (self._spell) ───

    # ─── Assistant Tool Cards: GUI/assistant/tool_cards.py (self._tool_cards) ───

    # ─── Vision view capture (C2) ──────────────────────────────────────────────

    def _wants_view_snapshot(self, text, provider):
        """True when the user asks to LOOK at the current view and the
        provider takes image blocks (Claude only, for now)."""
        if not text:
            return False
        if provider is None or getattr(provider, "NAME", "") != "claude":
            return False
        try:
            if not provider.supports_vision():
                return False
        except Exception:
            return False
        import re as _re
        pat = (u"(nhìn|xem|quan sát|soi|chụp|đánh giá|kiểm tra"
               u"|look|inspect|review|check|analy)"
               u"[^\n]{0,24}"
               u"(view|màn hình|bố cục|layout|screen)")
        return _re.search(pat, text, _re.IGNORECASE | _re.UNICODE) is not None

    def _capture_active_view(self, srv):
        """WORKER thread: export the active view as a PNG (~1280px) through
        the ExternalEvent write path. Returns the PNG path or None."""
        try:
            import tempfile
            folder = os.path.join(tempfile.gettempdir(), 'T3Lab_ViewShots')
            # Clears shots a crashed/cancelled turn left (sent ones are deleted).
            try:
                from core import housekeeping
                housekeeping.run_once('assistant.viewshots',
                                      housekeeping.prune_files, folder,
                                      max_age_days=1)
            except Exception:
                pass
            res = srv._execute_tool('export_image',
                                    {'width': 1280, 'output_folder': folder})
            files = (res or {}).get('files') or []
            for p in files:
                if p.lower().endswith('.png') and os.path.isfile(p):
                    return p
            return files[0] if files else None
        except Exception:
            return None

    # ─── Chat UI helpers ──────────────────────────────────────────────────────

    _avatar_bitmap = None   # lazy-loaded, cached BitmapImage of icon.png (per window)
    _avatar_bitmap_failed = False

    def _get_avatar_bitmap(self):
        """Cached BitmapImage of the ribbon icon (T3LabAssistant.pushbutton/
        icon.png) used as the chat avatar, or None."""
        if self._avatar_bitmap is not None:
            return self._avatar_bitmap
        if self._avatar_bitmap_failed:
            return None
        try:
            from System import Uri, UriKind
            from System.Windows.Media.Imaging import BitmapCacheOption
            icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icon.png')
            bmp = BitmapImage()
            bmp.BeginInit()
            bmp.UriSource   = Uri(icon_path, UriKind.Absolute)
            bmp.CacheOption = BitmapCacheOption.OnLoad
            bmp.EndInit()
            bmp.Freeze()
            self._avatar_bitmap = bmp
            return bmp
        except Exception as ex:
            logger.debug(u"_get_avatar_bitmap error: {}".format(_exc_text(ex)))
            self._avatar_bitmap_failed = True
            return None

    def _stop_avatar_spins(self):
        """Stop every live avatar spin (the 'thinking' indicator). UI thread."""
        spins = getattr(self, '_live_avatar_spins', None) or []
        if not spins:
            return
        try:
            from System.Windows.Media import RotateTransform
            for spin in spins:
                try:
                    spin.BeginAnimation(RotateTransform.AngleProperty, None)
                except Exception:
                    pass
        except Exception:
            pass
        self._live_avatar_spins = []

    def _make_avatar(self, letter, size=36):
        """Circular avatar showing icon.png; an initials badge if it can't load."""
        from System.Windows.Controls import Border, TextBlock
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Media import SolidColorBrush, Color, ImageBrush, Stretch
        from System.Windows import HorizontalAlignment, VerticalAlignment

        av = Border()
        av.Width = size
        av.Height = size
        av.CornerRadius = CornerRadius(size / 2.0)
        av.Margin = Thickness(0, 2, 10, 0)
        av.VerticalAlignment = VerticalAlignment.Top

        # Spin ONLY while busy (user request 2026-07-22); _set_busy(False)
        # stops live spins. 30fps cap keeps it cheap.
        if getattr(self, '_busy', False):
            try:
                from System.Windows.Media import RotateTransform
                from System.Windows.Media.Animation import (DoubleAnimation,
                                                            RepeatBehavior,
                                                            Timeline)
                from System.Windows import Point, Duration
                from System import TimeSpan, Nullable, Int32
                av.RenderTransformOrigin = Point(0.5, 0.5)
                spin = RotateTransform()
                av.RenderTransform = spin
                anim = DoubleAnimation(
                    0.0, 360.0, Duration(TimeSpan.FromSeconds(1.6)))
                anim.RepeatBehavior = RepeatBehavior.Forever
                Timeline.SetDesiredFrameRate(anim, Nullable[Int32](30))
                spin.BeginAnimation(RotateTransform.AngleProperty, anim)
                if not hasattr(self, '_live_avatar_spins'):
                    self._live_avatar_spins = []
                self._live_avatar_spins.append(spin)
            except Exception:
                pass

        bmp = self._get_avatar_bitmap()
        if bmp is not None:
            brush = ImageBrush(bmp)
            brush.Stretch = Stretch.UniformToFill
            av.Background = brush
            return av

        av.Background = SolidColorBrush(Color.FromRgb(52, 152, 219))   # #3498DB fallback
        lbl = TextBlock()
        lbl.Text = letter
        lbl.FontSize = 12
        lbl.FontWeight = System.Windows.FontWeights.Bold
        lbl.Foreground = SolidColorBrush(Color.FromRgb(255, 255, 255))
        lbl.HorizontalAlignment = HorizontalAlignment.Center
        lbl.VerticalAlignment = VerticalAlignment.Center
        av.Child = lbl
        return av

    @staticmethod
    def _add_icon_run(text_block, glyph, rgb=None, size=13.5):
        """Prepend a Segoe MDL2 Assets icon Run + thin space to a TextBlock.

        The body font has no glyphs at these private-use codepoints, so the
        icon needs its own FontFamily run.
        """
        from System.Windows.Documents import Run
        from System.Windows.Media import FontFamily as _WpfFontFamily, SolidColorBrush, Color

        icon_run = Run(glyph)
        icon_run.FontFamily = _WpfFontFamily(u"Segoe MDL2 Assets")
        icon_run.FontSize   = size
        if rgb:
            icon_run.Foreground = SolidColorBrush(Color.FromRgb(*rgb))
        text_block.Inlines.Add(icon_run)
        text_block.Inlines.Add(Run(u"  "))

    # ─── Claude-style message actions ────────────────────────────────────────
    # Line-art glyphs for the reply action row (same 1.5px / 24x24 language
    # as the composer icons).

    _ACT_COPY  = (u"M8 8h11a2 2 0 0 1 2 2v11a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2V10a2 2 0 0 1 2-2Z "
                  u"M4 16a2 2 0 0 1-2-2V3a2 2 0 0 1 2-2h11a2 2 0 0 1 2 2")
    _ACT_CHECK = u"M20 6 9 17l-5-5"
    _ACT_RETRY = u"M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8 M3 3v5h5"
    _ACT_UP    = (u"M7 10v12 M15 5.88 14 10h5.83a2 2 0 0 1 1.92 2.56l-2.33 8A2 2 0 0 1 17.5 22H4"
                  u"a2 2 0 0 1-2-2v-8a2 2 0 0 1 2-2h2.76a2 2 0 0 0 1.79-1.11L12 2a3.13 3.13 0 0 1 3 3.88Z")
    _ACT_DOWN  = (u"M17 14V2 M9 18.12 10 14H4.17a2 2 0 0 1-1.92-2.56l2.33-8A2 2 0 0 1 6.5 2H20"
                  u"a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2h-2.76a2 2 0 0 0-1.79 1.11L12 22a3.13 3.13 0 0 1-3-3.88Z")
    _ACT_PREV  = u"M15 18l-6-6 6-6"
    _ACT_NEXT  = u"M9 18l6-6-6-6"


    def _make_message_actions(self, row):
        """The action/feedback bar under assistant replies — disabled per user request."""
        return None


    def _refresh_reply_row(self, row):
        """Re-render a reply row's body + action bar from its Tag state."""
        from System.Windows.Controls import TextBlock
        from System.Windows import TextWrapping

        state = row.Tag or {}
        versions = state.get('versions') or [u""]
        idx = max(0, min(state.get('index', 0), len(versions) - 1))
        text = versions[idx]

        host = state.get('host')
        if host is None:
            return
        host.Children.Clear()
        try:
            body = self._render_md_blocks(text, icon=state.get('icon'),
                                          icon_color=state.get('icon_color'))
        except Exception:
            body = TextBlock()
            body.FontSize     = 13.5
            body.FontFamily   = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
            _bind_fg(body, 'BotText')
            body.TextWrapping = TextWrapping.Wrap
            body.LineHeight   = 21
            body.Text         = text
        host.Children.Add(body)
        acts = self._make_message_actions(row)
        if acts is not None:
            host.Children.Add(acts)

    # ─── Action handlers ─────────────────────────────────────────────────────


    def _vote_decision(self):
        """Snapshot of how the current turn was routed, for the vote loop.

        Taken at row build time so a late vote refers to its own turn.
        """
        dec = getattr(self, '_agent_decision', None) or {}
        last = getattr(self, '_last_result', None) or {}
        return {
            'raw':        getattr(self, '_last_raw', u"") or u"",
            'intent':     last.get('intent'),
            'params':     last.get('params') or {},
            'message':    last.get('message') or u"",
            'skill':      dec.get('skill'),
            'specialist': dec.get('specialist'),
        }


    def _adopt_retry_version(self, text, icon=None, icon_color=None):
        """File `text` as another version of the row Try again was pressed on.

        True = adopted (caller must NOT append a row). On False the reply
        lands as a new message, so a broken regenerate never swallows it.
        """
        row = getattr(self, '_retry_target_row', None)
        if row is None:
            return False
        self._retry_target_row = None
        try:
            state = row.Tag or {}
            versions = state.get('versions')
            if not isinstance(versions, list):
                return False
            versions.append(text)
            state['versions']   = versions
            state['index']      = len(versions) - 1
            state['icon']       = icon
            state['icon_color'] = icon_color
            state['vote']       = None
            row.Tag = state
            self._refresh_reply_row(row)
            self._scroll_to_bottom()
            return True
        except Exception as ex:
            logger.debug(u"_adopt_retry_version failed: {}".format(_exc_text(ex)))
            return False

    def _append_user_message(self, text, attachment_note=None):
        """Add the user's message as right-aligned text.

        No fill/border/shadow: a card reads as a second surface in a pane this
        narrow; alignment alone marks the speaker. attachment_note renders on
        its own line with an MDL2 Attach glyph.
        """
        try:
            from System.Windows.Controls import Border, TextBlock, Grid, ColumnDefinition
            from System.Windows.Documents import Run, LineBreak
            from System.Windows import Thickness, CornerRadius, TextWrapping, GridLength, HorizontalAlignment

            row = Grid()
            row.Margin = Thickness(40, 4, 0, 14)
            col0 = ColumnDefinition()
            col0.Width = GridLength(1, System.Windows.GridUnitType.Star)
            row.ColumnDefinitions.Add(col0)

            bubble = Border()
            _bind_bg(bubble, 'ChatBg')
            bubble.Padding      = Thickness(2, 0, 2, 0)
            bubble.HorizontalAlignment = HorizontalAlignment.Right

            msg_text = TextBlock()
            msg_text.FontSize    = 13.5
            msg_text.FontFamily  = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
            _bind_fg(msg_text, 'UserBubbleText')
            msg_text.LineHeight  = 21
            msg_text.TextWrapping = TextWrapping.Wrap
            if text:
                msg_text.Inlines.Add(Run(text))
            if attachment_note:
                if text:
                    msg_text.Inlines.Add(LineBreak())
                self._add_icon_run(msg_text, _ICON_ATTACH, size=12)
                msg_text.Inlines.Add(Run(attachment_note))
            bubble.Child = msg_text

            Grid.SetColumn(bubble, 0)
            row.Children.Add(bubble)
            self.chat_history_panel.Children.Add(row)
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug(u"Error adding user message: {}".format(_exc_text(ex)))

    @staticmethod
    def _clean_bot_response(text):
        """Strip chain-of-thought, meta-commentary, and excessive whitespace."""
        import re as _re
        text = _re.sub(r'<think>[\s\S]*?</think>', '', text, flags=_re.IGNORECASE)
        _skip_prefixes = (
            u"* user", u"* role", u"* tone", u"* language", u"* option",
            u"* the user", u"* since i am", u"* as an ai",
            u"- user", u"- role", u"- tone", u"- option",
            u"i am an ai", u"as an ai assistant",
            u"let me analyze", u"let me think", u"i need to consider",
            u"i'll analyze", u"i'll consider",
        )
        lines_out = []
        for line in text.splitlines():
            low = line.strip().lower()
            if any(low.startswith(p) for p in _skip_prefixes):
                continue
            lines_out.append(line)
        text = _re.sub(r'\n{3,}', u'\n\n', u'\n'.join(lines_out))
        return text.strip()

    # Leading emoji markers → MDL2 glyph + colour at render time (old saved
    # history, nlu_engine text and model-invented emoji alike).
    _MD_ICON_MARKERS = [
        (u"⚡ ",            "_ICON_ANALYZE", "_ICON_BLUE"),   # instant DB answer
        (u"\U0001f4cb ",   "_ICON_LIST",    "_ICON_SLATE"),  # stats/info section
        (u"\U0001f5bc️ ", "_ICON_SEARCH", "_ICON_SLATE"),  # view section
        (u"\U0001f3af ",   "_ICON_LIST",    "_ICON_SLATE"),  # selection section
        (u"\U0001f4ca ",   "_ICON_LIST",    "_ICON_BLUE"),   # LLM chart/stats emoji
        (u"\U0001f4c8 ",   "_ICON_LIST",    "_ICON_BLUE"),   # LLM trend emoji
        # (house rule at _ICON_INFO). A U+FE0F variation form must precede its
        # bare twin, or startswith strands the selector.
        (u"✅ ",       "_ICON_SUCCESS", "_ICON_GREEN"),  # ✅ yes / found
        (u"✔ ",       "_ICON_SUCCESS", "_ICON_GREEN"),  # ✔ completed
        (u"❌ ",       "_ICON_STOP",    "_ICON_RED"),    # ❌ none / no tool
        (u"⚠️ ", "_ICON_WARNING", "_ICON_AMBER"),  # ⚠️ warning
        (u"⚠ ",       "_ICON_WARNING", "_ICON_AMBER"),  # ⚠ warning
        (u"\U0001f50d ",   "_ICON_SEARCH",  "_ICON_BLUE"),   # 🔍 QA / lookup
        (u"\U0001f4c2 ",   "_ICON_LIST",    "_ICON_SLATE"),  # 📂 ribbon tools
        (u"\U0001f9e9 ",   "_ICON_LIST",    "_ICON_SLATE"),  # 🧩 other tools
        (u"\U0001f4d6 ",   "_ICON_LIST",    "_ICON_SLATE"),  # 📖 read / query
        (u"\U0001f4c1 ",   "_ICON_LIST",    "_ICON_SLATE"),  # 📁 view / sheet
        (u"\U0001f4e4 ",   "_ICON_LIST",    "_ICON_SLATE"),  # 📤 export
        (u"✏️ ", "_ICON_LIST",    "_ICON_SLATE"),  # ✏️ create / modify
        (u"✏ ",       "_ICON_LIST",    "_ICON_SLATE"),  # ✏ create / modify
    ]

    @staticmethod
    def _add_inline_md(text_block, line):
        """Render ONE line's inline markdown into text_block:
        **bold** and `code` spans (code = Consolas, ink color)."""
        import re as _re
        from System.Windows.Documents import Run
        from System.Windows import FontWeights
        from System.Windows.Media import FontFamily as _WpfFontFamily, SolidColorBrush, Color

        for b_idx, b_seg in enumerate(_re.split(r'\*\*', line)):
            if not b_seg:
                continue
            bold = (b_idx % 2 == 1)
            for c_idx, c_seg in enumerate(b_seg.split(u'`')):
                if not c_seg:
                    continue
                r = Run()
                r.Text = c_seg
                if bold:
                    r.FontWeight = FontWeights.SemiBold
                if c_idx % 2 == 1:              # inside `code`
                    r.FontFamily = _WpfFontFamily(u"Consolas")
                    _bind_fg(r, 'CodeFg')
                text_block.Inlines.Add(r)

    @staticmethod
    def _build_md_inlines(text_block, text):
        """Fill text_block.Inlines with paragraph-level markdown (icon markers,
        headings, quotes, lists, inline spans). Tables: _render_md_blocks."""
        from System.Windows.Documents import Run, LineBreak
        from System.Windows import FontWeights

        module_globals = globals()
        lines = text.splitlines()
        for i, line in enumerate(lines):
            if i > 0:
                text_block.Inlines.Add(LineBreak())

            stripped = line.strip()
            for marker, glyph_name, color_name in T3LabAssistantWindow._MD_ICON_MARKERS:
                if stripped.startswith(marker):
                    T3LabAssistantWindow._add_icon_run(
                        text_block, module_globals[glyph_name], module_globals[color_name])
                    line = stripped[len(marker):]
                    break

            stripped = line.strip()
            if stripped.startswith(u'#'):
                h = len(stripped) - len(stripped.lstrip(u'#'))
                if 1 <= h <= 4 and stripped[h:h + 1] == u' ':
                    r = Run()
                    r.Text       = stripped[h + 1:].strip()
                    r.FontWeight = FontWeights.SemiBold
                    r.FontSize   = 14.5 if h <= 2 else 13.5
                    text_block.Inlines.Add(r)
                    continue

            if stripped.startswith(u'> '):
                from System.Windows.Media import SolidColorBrush, Color
                bar = Run()
                bar.Text       = u'▏ '   # ▏ left one-eighth block
                _bind_fg(bar, 'Faint')
                text_block.Inlines.Add(bar)
                quote = Run()
                quote.Text       = stripped[2:]
                _bind_fg(quote, 'Muted')
                quote.FontStyle  = System.Windows.FontStyles.Italic
                text_block.Inlines.Add(quote)
                continue

            import re as _re_md
            _om = _re_md.match(r'^(\d{1,3})[.)]\s+(.*)$', stripped)
            if _om:
                num_run = Run()
                num_run.Text       = u'{}. '.format(_om.group(1))
                num_run.FontWeight = FontWeights.SemiBold
                text_block.Inlines.Add(num_run)
                T3LabAssistantWindow._add_inline_md(text_block, _om.group(2))
                continue

            if stripped.startswith(u'* ') or stripped.startswith(u'- '):
                indent = len(line) - len(line.lstrip())
                prefix_run = Run()
                prefix_run.Text = (u'    ' if indent >= 2 else u'') + u'• '
                text_block.Inlines.Add(prefix_run)
                line = stripped[2:]

            T3LabAssistantWindow._add_inline_md(text_block, line)

    def _make_md_table(self, raw_rows):
        """Render markdown pipe-rows ("| a | b |") as a bordered WPF Grid.

        None when the rows don't form a table (caller shows the raw text).
        Star-sized columns so the table always fits the bubble.
        """
        try:
            import re as _re
            from System.Windows.Controls import Grid, ColumnDefinition, RowDefinition, Border, TextBlock
            from System.Windows import Thickness, CornerRadius, TextWrapping, GridLength, GridUnitType
            from System.Windows.Media import SolidColorBrush, Color

            parsed = []
            for r in raw_rows:
                inner = r.strip()
                if inner.startswith(u"|"):
                    inner = inner[1:]
                if inner.endswith(u"|"):
                    inner = inner[:-1]
                parsed.append([c.strip() for c in inner.split(u"|")])

            def _is_sep(cells):
                return bool(cells) and all(
                    _re.match(r'^:?-{2,}:?$', c or u'') for c in cells)

            header = None
            if len(parsed) >= 2 and _is_sep(parsed[1]):
                header = parsed[0]
                body = [r for r in parsed[2:] if not _is_sep(r)]
            else:
                body = [r for r in parsed if not _is_sep(r)]
            rows = ([header] if header is not None else []) + body
            if not rows:
                return None
            ncols = max(len(r) for r in rows)
            if ncols < 2:
                return None

            _line  = _theme_color('Divider')
            _inkhd = _theme_color('Ink')
            _ink   = _theme_color('BotText')

            g = Grid()
            for _c in range(ncols):
                cd = ColumnDefinition()
                cd.Width = GridLength(1, GridUnitType.Star)
                g.ColumnDefinitions.Add(cd)

            for ri, row in enumerate(rows):
                g.RowDefinitions.Add(RowDefinition())
                is_head = (header is not None and ri == 0)
                for ci in range(ncols):
                    cell = Border()
                    cell.BorderBrush = SolidColorBrush(_line)
                    cell.BorderThickness = Thickness(
                        0, 0,
                        1 if ci < ncols - 1 else 0,
                        1 if ri < len(rows) - 1 else 0)
                    cell.Padding = Thickness(8, 4, 8, 4)
                    if is_head:
                        _bind_bg(cell, 'SelectedBg')

                    tb = TextBlock()
                    tb.FontSize     = 12 if is_head else 12.5
                    tb.FontFamily   = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
                    tb.TextWrapping = TextWrapping.Wrap
                    tb.Foreground   = SolidColorBrush(_inkhd if is_head else _ink)
                    if is_head:
                        tb.FontWeight = System.Windows.FontWeights.SemiBold
                    self._add_inline_md(tb, row[ci] if ci < len(row) else u"")
                    cell.Child = tb
                    Grid.SetRow(cell, ri)
                    Grid.SetColumn(cell, ci)
                    g.Children.Add(cell)

            outer = Border()
            outer.BorderBrush     = SolidColorBrush(_line)
            outer.BorderThickness = Thickness(1)
            outer.CornerRadius    = CornerRadius(4)
            outer.Margin          = Thickness(0, 6, 0, 6)
            outer.Child = g
            return outer
        except Exception as ex:
            logger.debug(u"_make_md_table error: {}".format(_exc_text(ex)))
            return None

    def _make_code_block(self, code_lines):
        """Fenced ``` code block as a monospace card with horizontal scroll."""
        from System.Windows.Controls import Border, TextBlock, ScrollViewer, ScrollBarVisibility
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Media import SolidColorBrush, Color, FontFamily

        tb = TextBlock()
        tb.Text       = u"\n".join(code_lines).rstrip(u"\n")
        tb.FontFamily = FontFamily(u"Consolas")
        tb.FontSize   = 12
        _bind_fg(tb, 'Ink')

        sv = ScrollViewer()
        sv.HorizontalScrollBarVisibility = ScrollBarVisibility.Auto
        sv.VerticalScrollBarVisibility   = ScrollBarVisibility.Disabled
        sv.Content = tb

        card = Border()
        _bind_bg(card, 'CodeBg')
        _bind_border(card, 'CardBorder')
        card.BorderThickness = Thickness(1)
        card.CornerRadius    = CornerRadius(4)
        card.Padding         = Thickness(10, 8, 10, 8)
        card.Margin          = Thickness(0, 4, 0, 4)
        card.Child = sv
        return card

    def _make_hr(self):
        """A thin horizontal divider for markdown '---' / '***' rules."""
        from System.Windows.Controls import Border
        from System.Windows import Thickness
        from System.Windows.Media import SolidColorBrush, Color
        hr = Border()
        hr.Height     = 1
        _bind_bg(hr, 'Divider')
        hr.Margin      = Thickness(0, 8, 0, 8)
        return hr

    def _render_md_blocks(self, text, icon=None, icon_color=None):
        """Bot reply content: a StackPanel of paragraph TextBlocks, table
        Grids, code cards and rules. icon/icon_color prefix the first paragraph.
        """
        from System.Windows.Controls import StackPanel, TextBlock
        from System.Windows import TextWrapping, Thickness
        from System.Windows.Media import SolidColorBrush, Color

        panel = StackPanel()
        state = {"icon": icon}

        def _new_tb():
            tb = TextBlock()
            tb.FontSize     = 13
            tb.FontFamily   = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
            _bind_fg(tb, 'BotText')
            tb.TextWrapping = TextWrapping.Wrap
            tb.LineHeight   = 20
            return tb

        para = []

        def _flush_para():
            if not para:
                return
            chunk = u"\n".join(para).strip(u"\n")
            del para[:]
            if not chunk.strip() and not state["icon"]:
                return
            tb = _new_tb()
            if state["icon"]:
                self._add_icon_run(tb, state["icon"], icon_color)
                state["icon"] = None
            self._build_md_inlines(tb, chunk)
            tb.Margin = Thickness(0, 0, 0, 2)
            panel.Children.Add(tb)

        lines = (text or u"").splitlines() or [u""]
        i = 0
        n = len(lines)
        while i < n:
            s = lines[i].strip()

            if s.startswith(u"```"):
                i += 1
                code_lines = []
                while i < n and not lines[i].strip().startswith(u"```"):
                    code_lines.append(lines[i])
                    i += 1
                if i < n:               # consume the closing fence
                    i += 1
                _flush_para()
                try:
                    panel.Children.Add(self._make_code_block(code_lines))
                except Exception:
                    para.extend(code_lines)
                continue

            if s in (u"---", u"***", u"___"):
                _flush_para()
                try:
                    panel.Children.Add(self._make_hr())
                except Exception:
                    pass
                i += 1
                continue

            if s.startswith(u"|") and s.count(u"|") >= 2:
                tbl = []
                while i < n:
                    s2 = lines[i].strip()
                    if s2.startswith(u"|") and s2.count(u"|") >= 2:
                        tbl.append(s2)
                        i += 1
                    else:
                        break
                _flush_para()
                table = self._make_md_table(tbl)
                if table is not None:
                    panel.Children.Add(table)
                else:
                    para.extend(tbl)     # unparseable → show as raw text
                continue
            para.append(lines[i])
            i += 1
        _flush_para()

        if panel.Children.Count == 0:
            panel.Children.Add(_new_tb())
        return panel

    def _append_bot_message(self, text, icon=None, icon_color=None, actions=True):
        """Add an assistant reply — full-width text on the pane, no bubble
        (in a ~320px docked pane a bubble + avatar gutter leaves two-word lines).

        icon/icon_color: optional MDL2 glyph + RGB tuple before the text.
        """
        # A blank row looks like a broken answer.
        if not (text or u"").strip() and icon is None:
            return

        # A pending "Try again" folds this answer into its row.
        if getattr(self, '_retry_target_row', None) is not None:
            if self._adopt_retry_version(text, icon=icon, icon_color=icon_color):
                return

        try:
            from System.Windows.Controls import TextBlock, Grid, StackPanel, ColumnDefinition
            from System.Windows import Thickness, TextWrapping, GridLength

            row = Grid()
            row.Margin = Thickness(0, 0, 8, 16)
            col_msg = ColumnDefinition()
            col_msg.Width = GridLength(1, System.Windows.GridUnitType.Star)
            row.ColumnDefinitions.Add(col_msg)

            host = StackPanel()
            Grid.SetColumn(host, 0)
            row.Children.Add(host)

            # Action-row state: versions, and the prompt (for Try again).
            row.Tag = {
                'versions':   [text],
                'index':      0,
                'icon':       icon,
                'icon_color': icon_color,
                'vote':       None,
                'host':       host,
                'prompt':     getattr(self, '_last_raw', u"") or u"",
                # So a 👍/👎 can act on the route.
                'decision':   self._vote_decision(),
            }

            try:
                content = self._render_md_blocks(text, icon=icon, icon_color=icon_color)
            except Exception:
                content = TextBlock()
                content.FontSize     = 13.5
                content.FontFamily   = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
                _bind_fg(content, 'BotText')
                content.TextWrapping = TextWrapping.Wrap
                content.LineHeight   = 21
                content.Text         = text
            host.Children.Add(content)

            if actions:
                try:
                    acts = self._make_message_actions(row)
                    if acts is not None:
                        host.Children.Add(acts)
                except Exception as ex:
                    logger.debug(u"message actions skipped: {}".format(_exc_text(ex)))

            self.chat_history_panel.Children.Add(row)
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug(u"Error adding bot message: {}".format(_exc_text(ex)))

    # ─── Live streaming bubble ────────────────────────────────────────────────

    def _begin_stream_bubble(self):
        """Empty assistant row filled token-by-token. UI thread only.

        Same shape as a finished reply, so nothing re-flows at finalize.
        Sets self._stream_row / _stream_tb / _stream_host.
        """
        try:
            from System.Windows.Controls import TextBlock, Grid, StackPanel, ColumnDefinition
            from System.Windows import Thickness, TextWrapping, GridLength

            row = Grid()
            row.Margin = Thickness(0, 0, 8, 16)
            col_msg = ColumnDefinition()
            col_msg.Width = GridLength(1, System.Windows.GridUnitType.Star)
            row.ColumnDefinitions.Add(col_msg)

            host = StackPanel()
            Grid.SetColumn(host, 0)
            row.Children.Add(host)

            tb = TextBlock()
            tb.FontSize     = 13.5
            tb.FontFamily   = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")
            _bind_fg(tb, 'BotText')
            tb.TextWrapping = TextWrapping.Wrap
            tb.LineHeight   = 21
            host.Children.Add(tb)

            self.chat_history_panel.Children.Add(row)
            self._stream_row  = row
            self._stream_tb   = tb
            self._stream_host = host
            self._scroll_to_bottom()
        except Exception as ex:
            logger.debug(u"_begin_stream_bubble error: {}".format(_exc_text(ex)))
            self._stream_row  = None
            self._stream_tb   = None
            self._stream_host = None

    def _finalize_stream_bubble(self, text):
        """Re-render the live (plain-text) row with full markdown when done."""
        try:
            tb = self._stream_tb
            if tb is None:
                return
            row  = self._stream_row
            host = getattr(self, '_stream_host', None)

            # Pending "Try again": fold into its row, drop the stream row.
            if getattr(self, '_retry_target_row', None) is not None:
                if self._adopt_retry_version(text):
                    self._remove_stream_bubble()
                    return

            if host is not None and row is not None:
                # Promote to a full reply row, as _append_bot_message builds.
                row.Tag = {
                    'versions':   [text],
                    'index':      0,
                    'icon':       None,
                    'icon_color': None,
                    'vote':       None,
                    'host':       host,
                    'prompt':     getattr(self, '_last_raw', u"") or u"",
                    'decision':   self._vote_decision(),
                }
                self._refresh_reply_row(row)
            else:
                tb.Inlines.Clear()
                try:
                    self._build_md_inlines(tb, text)
                except Exception:
                    tb.Text = text
            self._scroll_to_bottom()
        except Exception:
            # Keep the plain streamed text rather than lose the answer.
            try:
                if self._stream_tb is not None:
                    self._stream_tb.Text = text
            except Exception:
                pass

    def _remove_stream_bubble(self):
        """Discard the live row (used when an action renders its own reply)."""
        try:
            if self._stream_row is not None:
                self.chat_history_panel.Children.Remove(self._stream_row)
        except Exception:
            pass
        self._stream_row  = None
        self._stream_tb   = None
        self._stream_host = None

    def _clear_stream_refs(self):
        """Detach references so the row persists but is no longer 'live'."""
        self._stream_row  = None
        self._stream_tb   = None
        self._stream_host = None

    def _stream_llm_turn(self, provider, router, history, system_prompt,
                         query, max_tokens=1200, **kwargs):
        """Run one streaming LLM turn, filling a live bubble. Worker thread.

        The bubble appears on the first token (typing indicator until then).
        Returns the full raw response text (str) or None.
        """
        extractor = StreamingJSONExtractor()
        state = {"raw": [], "started": False}

        def _on_delta(chunk):
            if not chunk:
                return
            state["raw"].append(chunk)
            disp = _hide_reasoning(extractor.display(u"".join(state["raw"])))

            def _ui():
                if not state["started"]:
                    if not disp:
                        return
                    state["started"] = True
                    self._hide_typing_indicator()
                    self._begin_stream_bubble()
                try:
                    if self._stream_tb is not None:
                        # Live text is plain (fast); markdown applied on finalize.
                        self._stream_tb.Text = disp
                        self._scroll_to_bottom()
                except Exception:
                    pass

            try:
                self.Dispatcher.Invoke(Action(_ui))
            except Exception:
                pass

        try:
            # **kwargs MUST be forwarded (response_format=json_object on the
            # tool loop's first turn; without it prose failed the JSON parse).
            # is_configured(), NOT check_health(): that is a live GET /models
            # (8s) before every message. Let the real call's error surface —
            # as has_api_key() in Intelligence/t3lab_assistant.py.
            if provider is not None and _provider_usable(provider):
                return provider.chat_stream(
                    _budget_history(history, provider, system_prompt, query,
                                    max_tokens), system_prompt, query, _on_delta,
                    max_tokens, **kwargs)
            return router.chat_stream(
                _budget_history(history, None, system_prompt, query,
                                max_tokens), system_prompt, query, _on_delta,
                max_tokens, **kwargs)
        except Exception as ex:
            logger.debug(u"_stream_llm_turn error: {}".format(_exc_text(ex)))
            return None

    def _scroll_to_bottom(self):
        try:
            self.chat_scroll.ScrollToBottom()
        except Exception:
            pass


# ── MAIN SCRIPT ──


def show_assistant_dialog():
    """Show T3Lab Assistant docked in Revit, like the Properties palette.

    startup.py registers the pane (lib/assistant_pane.py); GUI/
    AssistantPaneControl moves this window's content into it and runs the GIL
    pump. Later clicks show / hide it with the conversation intact.

    Falls back to the modal window when the pane is unavailable (reason in
    %APPDATA%\\T3LabAI\\engine_check.log).
    """
    if not revit.doc:
        forms.alert("Please open a Revit document first.", exitscript=True)
        return None

    # API context: (re)create the ExternalEvents, else after a reload
    # re-imports Services/ every tool launch runs outside the API context.
    try:
        ensure_api_context()
    except Exception as ex:
        logger.debug(u"ensure_api_context: {}".format(_exc_text(ex)))
    try:
        from Services.mcp_service import MCPService as _MCPService
        _MCPService.ensure_external_event()
    except Exception as ex:
        logger.debug(u"MCP ExternalEvent: {}".format(_exc_text(ex)))

    try:
        from GUI.AssistantPaneControl import show_docked
        ok, detail = show_docked(lambda: T3LabAssistantWindow(is_docked=True),
                                 _get_uiapp())
    except Exception as ex:
        ok, detail = False, _exc_text(ex)
    if ok:
        return None
    logger.debug(u"Assistant dock pane unavailable ({}); opening the window.".format(detail))

    window = T3LabAssistantWindow(is_docked=False)
    window.ShowDialog()
    return window
