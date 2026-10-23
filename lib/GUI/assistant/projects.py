# -*- coding: utf-8 -*-
"""
Assistant projects — the active project (workspace): its instructions and
remembered facts in the prompt, its provider override, activation / new
project, the composer's project chip popup, the project panel, and the
per-project scheduled prompts timer.

Methods of T3LabAssistantWindow, moved here by composition (a CLR-derived
window cannot take a Python mixin): the window builds one
`self._projects = AssistantProjects(self)` and keeps every attribute;
`self.win` below IS that window. XAML handlers keep a one-line stub on
the window that delegates here.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Projects"

import os

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
import System.Windows
from System.Windows import Visibility

from GUI.assistant._log import logger
from GUI.assistant.constants import (_ICON_GREEN, _ICON_REFRESH, _ICON_SLATE,
                                     _ICON_SUCCESS, _ICON_SYNC)
from GUI.assistant.history import _active_budgets
from GUI.assistant.text import _exc_text
from GUI.assistant.theme import _bind_bg, _bind_fg


class AssistantProjects(object):
    """Project workspaces: prompt blocks, activation, chip popup, panel, schedules."""

    def __init__(self, window):
        self.win = window   # the T3LabAssistantWindow that owns this helper

    @staticmethod
    def _active_pid():
        try:
            from config.project_store import ProjectStore
            return ProjectStore().get_active_project_id()
        except Exception:
            return None

    # ─── Projects (workspaces) ────────────────────────────────────────────────

    def _project_prompt_blocks(self, memory_tool=True):
        """(project_instructions, memory_block) for the active project.

        The single source of project grounding for EVERY prompt path, legacy
        JSON-intent included (taken on attachments / big RAG context).

        memory_tool=False for turns with NO remember_fact tool: facts listed,
        no tool hint (else the model answers {"intent": "remember_fact"}).

        First value is READY TO APPEND: "## Personal instructions" then
        "## Project instructions"; cache-stable (changes only on user edits).
        Memory = facts + this chat's notes + earlier-chat digests
        (_session_prompt_block), sized to the model's budget. Never raises.
        """
        personal = u""
        project = u""
        memory = u""
        try:
            from config.settings import get_settings
            personal = get_settings().build_user_instructions_block() or u""
        except Exception:
            personal = u""
        try:
            from config.project_store import ProjectStore
            ps = ProjectStore()
            project = ps.get_active_prompt_addendum() or u""
            try:
                from Intelligence import assistant_memory
                try:
                    _mem_chars = _active_budgets()['memory_chars']
                except Exception:
                    _mem_chars = None
                memory = assistant_memory.build_memory_block(
                    ps.get_active_project_id(),
                    include_tool_hint=bool(memory_tool),
                    max_chars=_mem_chars) or u""
            except Exception:
                memory = u""
        except Exception:
            pass
        session = u""
        try:
            _sess = getattr(self.win, '_session_prompt_block', None)
            if _sess is not None:
                session = _sess(tool_hint=bool(memory_tool)) or u""
        except Exception:
            session = u""
        if session:
            memory = (memory + u"\n\n" + session) if memory else session
        sections = []
        if personal:
            sections.append(personal)
        if project:
            sections.append(u"## Project instructions\n" + project)
        return u"\n\n".join(sections), memory

    def _apply_project_blocks(self, system_prompt):
        """Append instructions + memory to a LEGACY (JSON-intent) system prompt.

        t3lab_assistant._build_system_prompt already appends the memory block
        (no tool hint), so memory is added only when the base lacks it; a base
        that has it still gets this chat's notes and earlier-chat digests.
        """
        instructions, memory = self._project_prompt_blocks(memory_tool=False)
        if instructions:
            system_prompt += u"\n\n" + instructions
        if memory and self.win._MEMORY_BLOCK_HEAD not in (system_prompt or u""):
            system_prompt += u"\n\n" + memory
        elif memory:
            session = u""
            try:
                _sess = getattr(self.win, '_session_prompt_block', None)
                if _sess is not None:
                    session = _sess(tool_hint=False) or u""
            except Exception:
                session = u""
            if session:
                system_prompt += u"\n\n" + session
        return system_prompt

    def _apply_project_provider(self, pid):
        """Apply (or undo) a project's provider/model override.

        persist=False keeps it out of settings.json; `_pre_project_provider`
        is restored when the project is left.
        """
        try:
            from Intelligence.llm_router import LLMRouter
            router = LLMRouter()
            meta = None
            try:
                from config.project_store import ProjectStore
                meta = ProjectStore().get_project(pid) if pid else None
            except Exception:
                meta = None

            want = (meta or {}).get('provider')
            if want:
                if getattr(self.win, '_pre_project_provider', None) is None:
                    self.win._pre_project_provider = router.get_active_name()
                router.switch_provider(want, (meta or {}).get('model'),
                                       persist=False)
            else:
                prev = getattr(self.win, '_pre_project_provider', None)
                if prev and prev != router.get_active_name():
                    router.switch_provider(prev, persist=False)
                self.win._pre_project_provider = None
            self.win._update_ai_badge()
        except Exception as ex:
            logger.debug(u"_apply_project_provider error: {}".format(_exc_text(ex)))

    def _project_overview_text(self, pid, header=None):
        """Markdown overview of what a project scope changes; posted on switch / create.

        Cheap: external knowledge_dirs are NOT walked (may be network shares)."""
        from config.project_store import ProjectStore
        ps = ProjectStore()
        meta = ps.get_project(pid) or {}
        lines = [header or u"**Project: {}**".format(
            meta.get('name', pid)), u""]

        instr = u" ".join(((meta.get('instructions') or u'')).split())
        if instr:
            if len(instr) > 110:
                instr = instr[:109] + u"…"
            lines.append(u"- **Instructions:** {}".format(instr))
        else:
            lines.append(u"- **Instructions:** none yet — add them via "
                         u"project popup → Customize project; they steer "
                         u"every reply in this project.")

        # One cached counter shared with popup subtitle, project panel and settings.
        try:
            n_files, n_extra, _docs, _unscanned = ps.count_documents(pid)
        except Exception:
            n_files, n_extra = 0, 0
        if n_files or n_extra:
            lines.append(u"- **Knowledge:** {} — replies quote these documents "
                         u"(project popup → Open knowledge folder to add "
                         u"more).".format(ps.describe_documents(pid)))
        else:
            lines.append(u"- **Knowledge:** empty — attach files in chat or "
                         u"drop PDF/DOCX/MD via project popup → Open "
                         u"knowledge folder.")

        try:
            from Intelligence import assistant_memory as _am
            _facts = _am.list_facts(pid)
            _n_proj = len([1 for _s, _f in _facts
                           if _s == _am.PROJECT_SCOPE])
            lines.append(u"- **Memory:** {} project fact(s) + {} global — "
                         u'say "remember ..." to add.'.format(
                             _n_proj, len(_facts) - _n_proj))
        except Exception:
            pass

        if meta.get('provider'):
            lines.append(u"- **AI provider:** {}{} (applied whenever this "
                         u"project activates).".format(
                             meta['provider'],
                             u" / " + meta['model']
                             if meta.get('model') else u""))
        else:
            lines.append(u"- **AI provider:** follows the global setting.")

        # Only CHATS are per-document (ProjectStore.history_path); the rest per DAY.
        lines.append(u"- **Chats:** stored in this project, separate per "
                     u"Revit document.")
        lines.append(u"- **Attachments & activity logs:** stored in this "
                     u"project, filed by date.")
        return u"\n".join(lines)

    def _activate_project(self, pid):
        """Switch the active project: history, knowledge scope, provider.
        UI THREAD — callers handle the busy guard themselves."""
        try:
            from config.project_store import ProjectStore
            ps = ProjectStore()
            ps.set_active_project(pid)

            try:
                self.win._reset_session_state(clear_transcript=True)
                self.win._sessions._restore_history()
            except Exception:
                pass

            self._apply_project_provider(pid)

            self.win._update_knowledge_status()
            self.win._kick_knowledge_scan()
            name = (ps.get_project(pid) or {}).get('name') if pid else None
            if pid and name:
                try:
                    _msg = self._project_overview_text(
                        pid, header=u"Switched to project **{}** — this "
                        u"scope is now active:".format(name))
                except Exception:
                    _msg = u"Switched to project **{}**.".format(name)
            else:
                _msg = (u"Project mode off — using the shared workspace: "
                        u"global knowledge & memory, per-document chat "
                        u"history, provider unchanged.")
            self.win._append_bot_message(_msg, icon=_ICON_REFRESH,
                                     icon_color=_ICON_SLATE)
            self.win._update_composer_chips()
        except Exception as ex:
            logger.debug(u"_activate_project error: {}".format(_exc_text(ex)))

    def _create_new_project(self):
        """Create a project and activate it (called from the project popup)."""
        try:
            if self.win._busy:
                return
            from config.project_store import ProjectStore
            ps = ProjectStore()
            n = len(ps.list_projects()) + 1
            meta = ps.create_project(u"Project {}".format(n))
            ps.set_active_project(meta['id'])
            # Same scope sync as _activate_project.
            try:
                self.win._reset_session_state(clear_transcript=True)
                self.win._sessions._restore_history()
                self._apply_project_provider(meta['id'])
                self.win._kick_knowledge_scan()
            except Exception:
                pass
            self.win._update_composer_chips()
            try:
                _msg = self._project_overview_text(
                    meta['id'],
                    header=u"Created & switched to project **{}**."
                    .format(meta['name']))
            except Exception:
                _msg = u"Created project **{}**.".format(meta['name'])
            self.win._append_bot_message(_msg, icon=_ICON_SUCCESS,
                                     icon_color=_ICON_GREEN)
            # Land the user straight in the rename box.
            self.win._open_llm_settings(tab='projects', select_pid=meta['id'])
        except Exception as ex:
            logger.debug(u"_create_new_project error: {}".format(_exc_text(ex)))

    def project_chip_clicked(self, sender, e):
        """Open the Claude-style project picker popup."""
        try:
            self._build_project_popup()
            self.win._keep_popup_in_pane(self.win.project_popup)
            self.win.project_popup.IsOpen = True
        except Exception as ex:
            logger.debug(u"project_chip_clicked error: {}".format(_exc_text(ex)))

    def _project_row_subtitle(self, ps, pid):
        """Content hint for a project popup row: 'N files · instructions ✓'.

        Uses the store's cached counter: runs once per project per popup open,
        on the UI thread.
        """
        try:
            meta = ps.get_project(pid) or {}
            n, n_dirs, _docs, _un = ps.count_documents(pid, cap=99)
            has_instr = bool((meta.get('instructions') or u'').strip())
            return u"{} file{}{} · {}".format(
                u"99+" if n > 99 else n, u"" if n == 1 else u"s",
                u" +{} linked".format(n_dirs) if n_dirs else u"",
                u"instructions ✓" if has_instr else u"no instructions")
        except Exception:
            return None

    def _build_project_popup(self):
        """Fill the project popup (none, projects, active-project actions). UI THREAD."""
        from config.project_store import ProjectStore
        ps = ProjectStore()
        active = ps.get_active_project_id()
        panel = self.win.project_popup_panel
        panel.Children.Clear()

        panel.Children.Add(self.win._popup_row(
            u"No project", u"Shared workspace",
            active=(not active),
            handler=lambda s, ev: self._project_popup_select(None)))
        for meta in ps.list_projects():
            panel.Children.Add(self.win._popup_row(
                meta['name'], self._project_row_subtitle(ps, meta['id']),
                active=(meta['id'] == active),
                handler=(lambda s, ev, _p=meta['id']:
                         self._project_popup_select(_p))))
        panel.Children.Add(self.win._popup_separator())
        if active:
            panel.Children.Add(self.win._popup_row(
                u"Project panel",
                u"Instructions · Memory · Context · Scheduled",
                icon=u"",
                handler=lambda s, ev: self._open_project_panel()))
            panel.Children.Add(self.win._popup_row(
                u"Open knowledge folder",
                u"Drop PDF/DOCX/MD here for replies",
                icon=u"",
                handler=lambda s, ev: self._project_popup_open_folder()))
        panel.Children.Add(self.win._popup_row(
            u"New project", None, icon=u"",
            handler=lambda s, ev: self._project_popup_new()))
        panel.Children.Add(self.win._popup_row(
            u"Customize project…", u"Name, instructions, knowledge",
            icon=u"",
            handler=lambda s, ev: self._project_popup_settings()))

    def _project_popup_select(self, pid):
        try:
            self.win.project_popup.IsOpen = False
        except Exception:
            pass
        try:
            if self.win._busy:
                self.win._append_bot_message(
                    u"A request is running — switch projects once it finishes.",
                    icon=_ICON_SYNC, icon_color=_ICON_SLATE)
                return
            from config.project_store import ProjectStore
            if ProjectStore().get_active_project_id() == pid:
                return
            self._activate_project(pid)
        except Exception as ex:
            logger.debug(u"_project_popup_select error: {}".format(_exc_text(ex)))

    def _project_popup_new(self):
        try:
            self.win.project_popup.IsOpen = False
        except Exception:
            pass
        self._create_new_project()

    # ─── Project panel ──

    def _open_project_panel(self):
        """Open the project panel overlay for the active project."""
        try:
            self.win.project_popup.IsOpen = False
        except Exception:
            pass
        try:
            from config.project_store import ProjectStore
            pid = ProjectStore().get_active_project_id()
            if not pid:
                return
            self._build_project_panel(pid)
            self.win.project_panel_overlay.Visibility = Visibility.Visible
        except Exception as ex:
            logger.debug(u"_open_project_panel error: {}".format(_exc_text(ex)))

    def project_panel_close_clicked(self, sender, e):
        try:
            self.win.project_panel_overlay.Visibility = Visibility.Collapsed
        except Exception:
            pass

    def _edit_project_in_settings(self, pid):
        """Close the panel and open LLMs Setting → Projects on this project."""
        try:
            self.win.project_panel_overlay.Visibility = Visibility.Collapsed
        except Exception:
            pass
        self.win._open_llm_settings(tab='projects', select_pid=pid)

    def _build_project_panel(self, pid):
        """Render a READ-ONLY overview of the project scope into
        project_panel_host. UI THREAD. Editing lives in LLMs Setting → Projects."""
        from System.Windows.Controls import (Border, Button, Grid as WGrid,
                                             ColumnDefinition, TextBlock)
        from System.Windows import (Thickness, TextWrapping, GridLength,
                                    TextDecorations)
        from System.Windows.Media import SolidColorBrush, Color
        from System.Windows.Input import Cursors
        from config.project_store import ProjectStore

        ps = ProjectStore()
        meta = ps.get_project(pid) or {}
        host = self.win.project_panel_host
        host.Children.Clear()
        try:
            self.win.project_panel_title.Text = meta.get('name') or u"Project"
        except Exception:
            pass

        _font  = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")

        def _txt(text, size=12, fg=None, bold=False, wrap=True, margin=None):
            # NOT named _tb: that would shadow the module-level theme helper
            # (UnboundLocalError, swallowed → panel never appeared).
            # `fg` is a theme TOKEN NAME, bound live.
            t = TextBlock()
            t.Text = text
            t.FontSize = size
            t.FontFamily = _font
            _bind_fg(t, fg or 'Ink')
            if bold:
                t.FontWeight = System.Windows.FontWeights.SemiBold
            if wrap:
                t.TextWrapping = TextWrapping.Wrap
            else:
                # section titles share a row with a button: trim, don't overlap
                t.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
            if margin is not None:
                t.Margin = margin
            return t

        def _two_col(left, right=None, top=6):
            g = WGrid()
            g.Margin = Thickness(0, top, 0, 0)
            g.ColumnDefinitions.Add(ColumnDefinition())
            c1 = ColumnDefinition()
            c1.Width = GridLength.Auto
            g.ColumnDefinitions.Add(c1)
            g.Children.Add(left)
            if right is not None:
                WGrid.SetColumn(right, 1)
                g.Children.Add(right)
            return g

        def _section(title, right_widget=None):
            host.Children.Add(_two_col(
                _txt(title, size=13.5, bold=True, wrap=False),
                right_widget, top=0))

        def _sep():
            b = Border()
            b.Height = 1
            _bind_bg(b, 'SelectedBg')
            b.Margin = Thickness(0, 14, 0, 14)
            host.Children.Add(b)

        def _small_btn(label, handler, primary=False):
            b = Button()
            b.Content = label
            # ComposerChipBtn is theme-bound; the Lumina Primary/SecondaryButton
            # are pinned to #0F172A (invisible in Revit's dark theme).
            try:
                b.Style = self.win.FindResource("ComposerChipBtn")
            except Exception:
                pass
            if primary:
                _bind_bg(b, 'AccentSoft')
                _bind_fg(b, 'Ink')
            b.FontSize = 11
            b.Padding = Thickness(12, 4, 12, 4)
            b.Margin = Thickness(8, 0, 0, 0)
            b.Click += handler
            return b

        # Read-only by design: one edit surface (LLMs Setting → Projects), so two
        # writers can't rewrite project.json in different shapes or race.

        # ── 0. Description ──
        desc = u" ".join((meta.get('description') or u'').split())
        if desc:
            host.Children.Add(_txt(desc, size=11.5, fg='Muted',
                                   margin=Thickness(0, 0, 0, 12)))
        else:
            host.Children.Add(_txt(
                u"No description yet — say what this project is for in "
                u"LLMs Setting → Projects.",
                size=11.5, fg='Faint', margin=Thickness(0, 0, 0, 12)))

        # ── 1. Instructions ──
        instr = (meta.get('instructions') or u'').strip()
        _section(u"Instructions",
                 _small_btn(u"Edit in LLMs Setting",
                            lambda s, ev: self._edit_project_in_settings(pid)))
        if instr:
            _prev = u" ".join(instr.split())
            if len(_prev) > 260:
                _prev = _prev[:259] + u"\u2026"
            host.Children.Add(_txt(_prev, size=11.5, fg='Muted',
                                  margin=Thickness(0, 6, 0, 0)))
        else:
            host.Children.Add(_txt(
                u"None yet \u2014 instructions steer every reply in this project.",
                size=11.5, fg='Faint', margin=Thickness(0, 6, 0, 0)))

        _sep()

        # ── 2. Memory (project-scope facts; forget them in the Projects tab) ──
        _section(u"Memory")
        try:
            from Intelligence import assistant_memory as _am
            _facts = _am.list_facts(pid)
            _proj = [f for _s, f in _facts if _s == _am.PROJECT_SCOPE]
            _glob = len(_facts) - len(_proj)
            if _proj:
                for _f in _proj[:8]:
                    host.Children.Add(_txt(
                        u"\u2022 " + (_f.get('text') or u''), size=11.5,
                        fg='Muted', margin=Thickness(0, 4, 0, 0)))
                if len(_proj) > 8:
                    host.Children.Add(_txt(
                        u"\u2026 and {} more".format(len(_proj) - 8),
                        size=11, fg='Faint', margin=Thickness(0, 4, 0, 0)))
            else:
                host.Children.Add(_txt(
                    u'No project facts yet \u2014 say "remember \u2026" in chat.',
                    size=11.5, fg='Faint', margin=Thickness(0, 6, 0, 0)))
            if _glob:
                host.Children.Add(_txt(
                    u"+ {} global fact(s) apply to every project.".format(_glob),
                    size=10.5, fg='Faint', margin=Thickness(0, 6, 0, 0)))
        except Exception as _mex:
            logger.debug(u"panel memory error: {}".format(_exc_text(_mex)))

        _sep()

        # ── 3. Knowledge (what the RAG index answers from) ──
        _section(u"Knowledge")
        try:
            host.Children.Add(_txt(ps.describe_documents(pid), size=11.5,
                                  fg='Muted', margin=Thickness(0, 6, 0, 0)))
        except Exception:
            pass

        # linked external folders (else a share-based project looks empty)
        try:
            for _d in (ps.get_knowledge_dirs(pid) or [])[:8]:
                _missing = not os.path.isdir(_d)
                _row = _txt(u"\U0001F517 " + (os.path.basename(_d.rstrip(u"\\/")) or _d)
                           + (u"  (not found)" if _missing else u""),
                           size=11, fg='Danger' if _missing else 'Faint',
                           margin=Thickness(0, 4, 0, 0))
                _row.ToolTip = _d
                host.Children.Add(_row)
        except Exception:
            pass

        lnk = _txt(u"Open knowledge folder", size=11.5, fg='Blue',
                  margin=Thickness(0, 8, 0, 0))
        lnk.Cursor = Cursors.Hand
        lnk.TextDecorations = TextDecorations.Underline
        lnk.MouseLeftButtonUp += (
            lambda s, ev: self.win._open_active_project_folder())
        host.Children.Add(lnk)

        _sep()

        # ── 4. Scheduled (daily prompts while the window is open) ──
        _section(u"Scheduled prompts")
        _sched = meta.get('scheduled') or []
        if _sched:
            for it in _sched:
                _pv = u" ".join((it.get('prompt') or u'').split())
                if len(_pv) > 70:
                    _pv = _pv[:69] + u"\u2026"
                _on = bool(it.get('enabled', True))
                host.Children.Add(_txt(
                    u"{} daily \u2014 {}{}".format(
                        it.get('time') or u'?', _pv,
                        u"" if _on else u"   (paused)"),
                    size=11.5, fg='Muted' if _on else 'Faint',
                    margin=Thickness(0, 4, 0, 0)))
            host.Children.Add(_txt(
                u"Runs once per day at the set time while the assistant "
                u"window is open (skipped while a request is busy).",
                size=10.5, fg='Faint', margin=Thickness(0, 8, 0, 0)))
        else:
            host.Children.Add(_txt(
                u"None yet \u2014 add recurring daily prompts in LLMs Setting "
                u"\u2192 Projects.", size=11.5, fg='Faint',
                margin=Thickness(0, 6, 0, 0)))

    # ─── Scheduled prompts runner ──

    def _start_schedule_timer(self):
        """30s UI-thread timer firing due per-project scheduled prompts."""
        try:
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan
            t = DispatcherTimer()
            t.Interval = TimeSpan.FromSeconds(30)
            t.Tick += self._schedule_tick
            t.Start()
            self.win._sched_timer = t
        except Exception as ex:
            logger.debug(u"_start_schedule_timer error: {}".format(_exc_text(ex)))

    def _schedule_tick(self, sender, e):
        """Fire at most ONE due scheduled prompt per tick.

        Due test (enabled, not run today, HH:MM passed) is
        ProjectStore.due_schedule — testable without WPF.
        """
        try:
            if self.win._busy:
                return
            # The tick drives the real composer: _process_input() would send the
            # user's staged files / forced /skill with this prompt. Wait.
            if getattr(self.win, '_attached_files', None):
                return
            if getattr(self.win, '_forced_skill_id', None):
                return
            if (self.win.chat_input.Text or u"").strip():
                return  # half-typed message: wait until the composer is idle

            from config.project_store import ProjectStore
            import time as _t
            ps = ProjectStore()
            pid = ps.get_active_project_id()
            if not pid:
                return
            items = (ps.get_project(pid) or {}).get('scheduled') or []
            if not items:
                return
            fired = ps.due_schedule(items, _t.strftime('%H:%M'),
                                    _t.strftime('%Y-%m-%d'))
            if not fired:
                return

            self.win.chat_input.Text = fired.get('prompt') or u''
            _before = self.win._request_id
            self.win._process_input()
            # Stamp last_run ONLY if the turn actually started (an early return
            # used to mark the task done without it reaching the model).
            if self.win._request_id != _before:
                ps.update_project(pid, {'scheduled': [
                    dict(t, last_run=_t.strftime('%Y-%m-%d'))
                    if t.get('id') == fired.get('id') else t
                    for t in items]})
            else:
                self.win.chat_input.Text = u""
        except Exception as ex:
            logger.debug(u"_schedule_tick error: {}".format(_exc_text(ex)))

    def _project_popup_open_folder(self):
        try:
            self.win.project_popup.IsOpen = False
        except Exception:
            pass
        self.win._open_active_project_folder()

    def _project_popup_settings(self):
        try:
            self.win.project_popup.IsOpen = False
        except Exception:
            pass
        self.win._open_llm_settings(tab='projects')
