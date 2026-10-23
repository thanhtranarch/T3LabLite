# -*- coding: utf-8 -*-
"""
Assistant sessions — the conversation's life cycle: restore on open, persist
every message, session identity + notes (Intelligence.session_context),
new chat / archive, and the History tab (list, search, scope, rename,
resume, clear).

Methods of T3LabAssistantWindow, moved here by composition (a CLR-derived
window cannot take a Python mixin): the window builds one
`self._sessions = AssistantSessions(self)` and keeps every attribute;
`self.win` below IS that window. XAML handlers keep a one-line stub on
the window that delegates here.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Sessions"

import datetime
import io
import json
import os

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
import System.Windows
from System.Windows import Visibility

from GUI.assistant._log import logger
from GUI.assistant.constants import (_ICON_AMBER, _ICON_EDIT, _ICON_INFO,
                                     _ICON_REFRESH, _ICON_SLATE, _ICON_WARNING)
from GUI.assistant.history import (MAX_REPLAY_BUBBLES, _active_budgets,
                                   _chat_history_dir, _fit_restored,
                                   _is_stale_tool_call_blob,
                                   _prune_archived_sessions,
                                   clear_chat_history, load_chat_history,
                                   load_chat_session, load_chat_summary,
                                   save_chat_history)
from GUI.assistant.lang import _note_user_language, _relative_time, _ui_viet
from GUI.assistant.text import _exc_text


class AssistantSessions(object):
    """Chat history on disk, the open session's identity/notes, and the History tab."""

    def __init__(self, window):
        self.win = window   # the T3LabAssistantWindow that owns this helper

    # ─── History persistence ───────────────────────────────────────────────────

    def _restore_history(self):
        """Load saved conversation from disk and replay bubbles + context.

        The "restored" banner goes BEFORE the replayed turns, with no action
        row (appended last it read as the latest reply, with rate buttons).
        """
        try:
            # Summary first: turns falling out of the window fold onto it.
            self.win._history_summary = load_chat_summary(self.win._doc_key)
            saved = load_chat_history(self.win._doc_key)
            if not saved:
                return
            self._set_session(load_chat_session(self.win._doc_key))

            # Only the bubbles are capped; the model gets what the budget allows.
            replay = saved[-MAX_REPLAY_BUBBLES:]
            hidden = len(saved) - len(replay)

            # Re-learn the language BEFORE rendering (_LAST_USER_VIET is process
            # state, reset by a restart); oldest first, latest decidable wins.
            for msg in replay:
                if msg.get("role") == "user" and msg.get("content"):
                    _note_user_language(msg["content"])

            viet = _ui_viet()
            self.win._append_bot_message(self._restored_banner(saved, viet,
                                                           hidden=hidden),
                                     icon=_ICON_REFRESH, icon_color=_ICON_SLATE,
                                     actions=False)

            for msg in replay:
                role    = msg.get("role", "")
                content = msg.get("content", "")
                if not content:
                    continue
                if role == "user":
                    self.win._append_user_message(content)
                elif role == "assistant":
                    self.win._append_bot_message(content)

            # Re-populate the model's context from the WHOLE saved chat.
            self.win._conversation_history, self.win._history_summary = _fit_restored(
                saved, self.win._history_summary, viet=viet)
            self.win._persisted_msgs = list(saved)
        except Exception as ex:
            logger.debug(u"Could not restore history: {}".format(_exc_text(ex)))

    def _restored_banner(self, saved, viet, hidden=0):
        """One line of context for a restored session: how much, how old.

        `hidden`: older messages not re-drawn (MAX_REPLAY_BUBBLES), said out
        loud so a short pane is not mistaken for a lost chat.
        """
        exchanges = max(1, len([m for m in saved if m.get("role") == "user"]))
        when = _relative_time(
            (saved[-1] or {}).get("ts") if saved else None, viet)
        if viet:
            head = u"Đã khôi phục hội thoại trước ({} lượt{})".format(
                exchanges, u", " + when if when else u"")
            tail = u"Bấm nút hội thoại mới dưới ô nhập để bắt đầu lại từ đầu."
            if hidden:
                tail = (u"{} tin nhắn cũ hơn không hiển thị ở đây nhưng vẫn "
                        u"nằm trong ngữ cảnh. ".format(hidden) + tail)
        else:
            head = u"Previous conversation restored ({} turn{}{})".format(
                exchanges, u"s" if exchanges != 1 else u"",
                u", " + when if when else u"")
            tail = u"Use the new-conversation button under the message box to start fresh."
            if hidden:
                tail = (u"{} older message{} not shown here, still part of "
                        u"the context. ".format(
                            hidden, u"s are" if hidden != 1 else u" is")
                        + tail)
        return head + u"\n" + tail

    # ─── Session identity + session memory ────────────────────────────────────

    def _set_session(self, meta):
        """Adopt a chat's identity and notes ({} = a fresh chat)."""
        meta = meta or {}
        self.win._session_id = meta.get('session_id') or None
        self.win._session_path = meta.get('session_path') or None
        try:
            from Intelligence.session_context import normalize_notes
            self.win._session_notes = normalize_notes(meta.get('session_notes'))
        except Exception:
            self.win._session_notes = []
        self.win._past_chats_cache = None

    def _session_state(self):
        """What save_chat_history stores about this chat besides its turns."""
        return {'session_id': getattr(self.win, '_session_id', None),
                'session_notes': list(getattr(self.win, '_session_notes', None) or []),
                'session_path': getattr(self.win, '_session_path', None)}

    def _save_session_notes(self, notes):
        """Store new session notes and write them to the history file now
        (a note saved before the chat's first reply must survive a crash)."""
        self.win._session_notes = list(notes or [])
        if not getattr(self.win, '_session_id', None):
            self.win._session_id = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        try:
            save_chat_history(self.win._doc_key, self.win._persisted_msgs,
                              summary=getattr(self.win, '_history_summary', u''),
                              session=self._session_state())
        except Exception as ex:
            logger.debug(u"Could not save session notes: {}".format(_exc_text(ex)))

    def _session_prompt_block(self, viet=False, tool_hint=True):
        """'## Notes for this conversation' + '## Earlier conversations'.

        Both change rarely, so they sit in the STATIC half of the system
        prompt (prompt cache keeps hitting).
        """
        parts = []
        try:
            from Intelligence import session_context as _sc
            notes = _sc.notes_block(getattr(self.win, '_session_notes', None),
                                    viet=viet, tool_hint=tool_hint)
            if notes:
                parts.append(notes)
            past = self._past_chats_block(viet=viet, tool_hint=tool_hint)
            if past:
                parts.append(past)
        except Exception as ex:
            logger.debug(u"session prompt block: {}".format(_exc_text(ex)))
        return u"\n\n".join(parts)

    def _past_chats_block(self, viet=False, tool_hint=True):
        """Digests of the newest earlier chats in scope (cached per chat).

        Scope: the active project's archived chats, else "No project" chats
        on the same Revit model; the open chat's own archive is excluded.
        """
        pid = self.win._projects._active_pid()
        key = (pid, self.win._doc_key, getattr(self.win, '_session_id', None),
               bool(viet), bool(tool_hint))
        cache = getattr(self.win, '_past_chats_cache', None)
        if cache and cache[0] == key:
            return cache[1]
        block = u""
        try:
            from config import chat_sessions
            from Intelligence.session_context import past_chats_block
            sessions = chat_sessions.list_sessions(
                pid=pid, doc_key=None if pid else self.win._doc_key, limit=40)
            own = getattr(self.win, '_session_id', None)
            sessions = [s for s in sessions if not (own and s.get('id') == own)]
            block = past_chats_block(
                sessions, _active_budgets()['past_chats_chars'], viet=viet,
                project=bool(pid), tool_hint=tool_hint)
        except Exception as ex:
            logger.debug(u"past chats block: {}".format(_exc_text(ex)))
        self.win._past_chats_cache = (key, block)
        return block

    def _session_note_tool(self, action, fact, replaces=None):
        """remember_fact with scope "session": save / update / forget one of
        this chat's notes. Returns the tool result dict. WORKER THREAD safe —
        it touches no WPF object."""
        try:
            from Intelligence import session_context as _sc
            notes = getattr(self.win, '_session_notes', None) or []
            action = (action or 'save').lower()
            if action == 'forget':
                notes, removed = _sc.forget_note(notes, replaces or fact)
                if removed is None:
                    return {"error": u"No matching note in this conversation."}
                self._save_session_notes(notes)
                return {"success": True,
                        "note": u"Forgot for this conversation: {}".format(removed)}
            if action == 'update':
                notes, ok, msg = _sc.update_note(notes, replaces, fact)
            else:
                notes, ok, msg = _sc.add_note(notes, fact)
            if not ok:
                return {"error": msg}
            self._save_session_notes(notes)
            return {"success": True, "note": msg}
        except Exception as ex:
            return {"error": u"{}".format(_exc_text(ex))}

    def _persist_message(self, role, content):
        """Append one message to the in-memory list and save to disk."""
        try:
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.win._persisted_msgs.append(
                {"role": role, "content": content, "ts": ts}
            )
            # Id from the first message, kept through restarts/resumes (no duplicate History rows).
            if not getattr(self.win, '_session_id', None):
                self.win._session_id = datetime.datetime.now().strftime(
                    "%Y%m%d_%H%M%S")
            save_chat_history(self.win._doc_key, self.win._persisted_msgs,
                              summary=getattr(self.win, '_history_summary', u''),
                              session=self._session_state())
        except Exception as ex:
            logger.debug(u"Could not persist message: {}".format(_exc_text(ex)))

    def reset_chat_clicked(self, sender, e):
        """Clear the chat history for this document and reset the UI."""
        try:
            # Resetting mid-request orphaned the live bubble and the stale
            # _typing_row blocked the typing indicator for the session.
            if getattr(self.win, '_busy', False):
                self.win._append_bot_message(
                    u"Still working on the current request — press Stop first, "
                    u"then start a new conversation.",
                    icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                return
            self.win._reset_session_state(clear_transcript=True)
            clear_chat_history(self.win._doc_key)
            # No "refreshed" bubble: the welcome panel returns on an empty transcript.
            self.win._update_welcome_greeting(fresh=True)
        except Exception as ex:
            logger.debug(u"reset_chat error: {}".format(_exc_text(ex)))

    def new_chat_clicked(self, sender=None, e=None):
        """Start a new chat conversation. Archive current session if it has messages."""
        try:
            if getattr(self.win, '_busy', False):
                self.win._append_bot_message(
                    u"Still working on the current request — press Stop first, "
                    u"then start a new conversation.",
                    icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                return

            if getattr(self.win, '_persisted_msgs', None):
                self._archive_current_session()

            self.win._reset_session_state(clear_transcript=True)
            clear_chat_history(self.win._doc_key)

            self.win._update_welcome_greeting(fresh=True)

            self.tab_chat_clicked()
        except Exception as ex:
            logger.debug(u"new_chat_clicked error: {}".format(_exc_text(ex)))

    def _archive_current_session(self):
        """Archive the conversation to the active project's sessions folder.

        Tagged with "pid" (config.chat_sessions.sessions_dir), so deleting the
        project takes it along. No project → the global archive.
        """
        try:
            if not getattr(self.win, '_persisted_msgs', None):
                return None
            from core import jsonsafe
            pid = self.win._projects._active_pid()
            try:
                from config import chat_sessions
                # _chat_history_dir keeps its one-time legacy import running.
                sessions_dir = (chat_sessions.sessions_dir(pid) if pid
                                else _chat_history_dir('sessions'))
            except Exception:
                sessions_dir = _chat_history_dir('sessions')

            title = u"Conversation"
            for m in self.win._persisted_msgs:
                if m.get("role") == "user" and m.get("content"):
                    c = m["content"].strip().replace('\n', ' ')
                    if len(c) > 40:
                        c = c[:37] + u"..."
                    title = c
                    break

            now = datetime.datetime.now()
            session_id = (getattr(self.win, '_session_id', None)
                          or now.strftime("%Y%m%d_%H%M%S"))
            ts = now.strftime("%Y-%m-%d %H:%M")

            # A resumed chat goes back into ITS OWN file (same id, user's title
            # kept), not a second History row — only while still in the same
            # scope (same folder, or a global file tagged with this project).
            fpath, old = None, {}
            prev = getattr(self.win, '_session_path', None)
            try:
                if prev and os.path.isfile(prev):
                    with io.open(prev, 'r', encoding='utf-8') as f:
                        data_prev = json.load(f) or {}
                    same_dir = (os.path.normcase(os.path.dirname(os.path.abspath(prev)))
                                == os.path.normcase(os.path.abspath(sessions_dir)))
                    tagged = bool(pid) and data_prev.get('pid') == pid
                    if same_dir or tagged:
                        fpath, old = prev, data_prev
            except Exception:
                fpath, old = None, {}
            if old.get('renamed') and old.get('title'):
                title = old['title']

            summary = getattr(self.win, '_history_summary', u'')
            digest = u""
            try:
                from Intelligence.session_context import make_digest
                digest = make_digest(self.win._persisted_msgs, summary, title)
            except Exception:
                digest = u""
            session_data = {
                "id": old.get('id') or session_id,
                "doc_key": self.win._doc_key,
                "title": title,
                "timestamp": ts,
                "updated": now.strftime("%Y-%m-%d %H:%M:%S"),
                "message_count": len(self.win._persisted_msgs),
                "messages": list(self.win._persisted_msgs),
                "summary": summary
            }
            if digest:
                session_data["digest"] = digest
            notes = list(getattr(self.win, '_session_notes', None) or [])
            if notes:
                session_data["session_notes"] = notes
            if old.get('renamed'):
                session_data["renamed"] = True
            if pid:
                session_data["pid"] = pid

            if fpath is None:
                fname = "{}_{}.json".format(self.win._doc_key, session_data["id"])
                fpath = os.path.join(sessions_dir, fname)
            data = jsonsafe.dumps(session_data, indent=2)
            with io.open(fpath, 'w', encoding='utf-8') as f:
                f.write(data)
            _prune_archived_sessions(sessions_dir)
            return fpath
        except Exception as ex:
            logger.debug(u"_archive_current_session error: {}".format(_exc_text(ex)))
            return None

    def tab_chat_clicked(self, sender=None, e=None):
        """Switch to Chat view."""
        try:
            from System.Windows import Visibility
            from System.Windows.Media import Brushes
            self.win.chat_view_grid.Visibility = Visibility.Visible
            self.win.history_view_grid.Visibility = Visibility.Collapsed

            try:
                self.win.tab_chat_border.Background = self.win.FindResource("T3.Surface")
                self.win.tab_chat_border.BorderBrush = self.win.FindResource("T3.Border")
                self.win.tab_chat_icon.Foreground = self.win.FindResource("T3.Ink")
                self.win.tab_history_border.Background = Brushes.Transparent
                self.win.tab_history_border.BorderBrush = Brushes.Transparent
                self.win.tab_history_icon.Foreground = self.win.FindResource("T3.TextSecondary")
            except Exception:
                pass

            try:
                self.win.chat_input.Focus()
            except Exception:
                pass
        except Exception as ex:
            logger.debug(u"tab_chat_clicked error: {}".format(_exc_text(ex)))

    def tab_history_clicked(self, sender=None, e=None):
        """Switch to History view and render past sessions."""
        try:
            from System.Windows import Visibility
            from System.Windows.Media import Brushes
            self.win.chat_view_grid.Visibility = Visibility.Collapsed
            self.win.history_view_grid.Visibility = Visibility.Visible

            try:
                self.win.tab_history_border.Background = self.win.FindResource("T3.Surface")
                self.win.tab_history_border.BorderBrush = self.win.FindResource("T3.Border")
                self.win.tab_history_icon.Foreground = self.win.FindResource("T3.Ink")
                self.win.tab_chat_border.Background = Brushes.Transparent
                self.win.tab_chat_border.BorderBrush = Brushes.Transparent
                self.win.tab_chat_icon.Foreground = self.win.FindResource("T3.TextSecondary")
            except Exception:
                pass

            self._render_history_sessions()
        except Exception as ex:
            logger.debug(u"tab_history_clicked error: {}".format(_exc_text(ex)))

    def _history_query(self):
        try:
            return (self.win.history_search_box.Text or u"").strip()
        except Exception:
            return u""

    def _history_scope_all(self):
        """True when the "All projects" chip is on."""
        try:
            return bool(self.win.history_scope_all.IsChecked)
        except Exception:
            return False

    def _get_history_sessions(self, query=None, all_projects=None):
        """Archived sessions for History, newest first (chat_sessions.list_sessions).

        "This project": active project, THIS document; "All projects": everything.
        """
        try:
            from config import chat_sessions
            # Keeps the global folder's one-time legacy import running.
            _chat_history_dir('sessions')
            if query is None:
                query = self._history_query()
            if all_projects is None:
                all_projects = self._history_scope_all()
            if all_projects:
                return chat_sessions.list_sessions(all_projects=True,
                                                   query=query or None)
            return chat_sessions.list_sessions(pid=self.win._projects._active_pid(),
                                               doc_key=self.win._doc_key,
                                               query=query or None)
        except Exception as ex:
            logger.debug(u"_get_history_sessions error: {}".format(_exc_text(ex)))
            return []

    def history_search_changed(self, sender=None, e=None):
        """Search box typed in — re-render after a short pause. UI THREAD."""
        try:
            from System.Windows.Threading import DispatcherTimer
            from System import TimeSpan
            timer = getattr(self.win, '_history_search_timer', None)
            if timer is None:
                timer = DispatcherTimer()
                timer.Interval = TimeSpan.FromMilliseconds(
                    self.win._HISTORY_SEARCH_DELAY_MS)

                def _tick(s, ev):
                    try:
                        timer.Stop()
                        self._render_history_sessions()
                    except Exception:
                        pass
                timer.Tick += _tick
                self.win._history_search_timer = timer
            timer.Stop()
            timer.Start()
        except Exception as ex:
            logger.debug(u"history_search_changed error: {}".format(_exc_text(ex)))
            self._render_history_sessions()

    def history_scope_changed(self, sender=None, e=None):
        """"This project" / "All projects" chip. UI THREAD."""
        try:
            if self.win.history_view_grid.Visibility == Visibility.Visible:
                self._render_history_sessions()
        except Exception:
            pass

    def _history_project_names(self):
        """{pid: name} for the "All projects" card labels. Never raises."""
        try:
            from config.project_store import ProjectStore
            return dict((p.get('id'), p.get('name') or p.get('id'))
                        for p in (ProjectStore().list_projects() or []))
        except Exception:
            return {}

    def _sync_history_header(self, all_projects, count, query):
        """Subtitle, scope-chip label and empty-state text. UI THREAD."""
        pid = self.win._projects._active_pid()
        try:
            self.win.history_scope_project.Content = (u"This project" if pid
                                                  else u"No project")
        except Exception:
            pass
        try:
            if all_projects:
                sub = u"Conversations from every project and document"
            elif pid:
                sub = u"Conversations of the active project, for this document"
            else:
                sub = u"Conversations outside any project, for this document"
            if query:
                sub = u"{} result{} for \u201c{}\u201d — {}".format(
                    count, u"" if count == 1 else u"s", query, sub[0].lower()
                    + sub[1:])
            self.win.history_subtitle.Text = sub
        except Exception:
            pass
        try:
            if query:
                self.win.history_empty_title.Text = u"No matching conversations"
                self.win.history_empty_hint.Text = (
                    u"Try fewer words, or switch to All projects.")
            else:
                self.win.history_empty_title.Text = u"No chat history yet"
                self.win.history_empty_hint.Text = (
                    u"Click + above to start a conversation.")
        except Exception:
            pass

    def _render_history_sessions(self):
        """Populate the history view with saved conversation cards."""
        try:
            from System.Windows import Visibility, Thickness, CornerRadius, GridUnitType, GridLength
            from System.Windows.Controls import Border, Grid, StackPanel, TextBlock, Button, ColumnDefinition, Orientation
            from System.Windows.Media import Brushes, FontFamily
            from config import chat_sessions

            self.win.history_sessions_panel.Children.Clear()
            query = self._history_query()
            all_projects = self._history_scope_all()
            sessions = self._get_history_sessions(query=query,
                                                  all_projects=all_projects)
            self._sync_history_header(all_projects, len(sessions), query)

            if not sessions:
                self.win.history_empty_state.Visibility = Visibility.Visible
                self.win.history_sessions_panel.Visibility = Visibility.Collapsed
                return

            self.win.history_empty_state.Visibility = Visibility.Collapsed
            self.win.history_sessions_panel.Visibility = Visibility.Visible
            names = self._history_project_names() if all_projects else {}

            for sess in sessions:
                fpath = sess.get('path')
                title_text = sess.get('title') or u"Conversation"
                meta_bits = []
                if all_projects:
                    _spid = sess.get('pid')
                    meta_bits.append(names.get(_spid, _spid) if _spid
                                     else chat_sessions.NO_PROJECT_LABEL)
                meta_bits.append(sess.get('created') or u"")
                meta_bits.append(u"{} messages".format(
                    sess.get('message_count') or 0))
                meta_text = u"  •  ".join(b for b in meta_bits if b)

                card = Border()
                card.Background = self.win.FindResource("T3.Surface")
                card.BorderBrush = self.win.FindResource("T3.Border")
                card.BorderThickness = Thickness(1)
                card.CornerRadius = CornerRadius(6)
                card.Margin = Thickness(0, 0, 0, 8)
                card.Padding = Thickness(12, 8, 12, 8)

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

                tb_title = TextBlock()
                tb_title.Text = title_text
                tb_title.Style = self.win.FindResource("T3.BodyStrong")
                tb_title.TextTrimming = System.Windows.TextTrimming.CharacterEllipsis
                tb_title.ToolTip = title_text

                tb_meta = TextBlock()
                tb_meta.Text = meta_text
                tb_meta.Style = self.win.FindResource("T3.Caption")
                tb_meta.Margin = Thickness(0, 4, 0, 0)
                # Narrow pane: wrap instead of running under Resume / Delete.
                tb_meta.TextWrapping = System.Windows.TextWrapping.Wrap

                info_stack.Children.Add(tb_title)
                info_stack.Children.Add(tb_meta)
                grid.Children.Add(info_stack)

                act_stack = StackPanel()
                act_stack.Orientation = Orientation.Horizontal
                act_stack.VerticalAlignment = System.Windows.VerticalAlignment.Center
                Grid.SetColumn(act_stack, 1)

                btn_open = Button()
                btn_open.Style = self.win.FindResource("T3.Button.Ghost")
                btn_open.Height = 28
                btn_open.Padding = Thickness(8, 0, 8, 0)
                btn_open.ToolTip = u"Resume this conversation"
                btn_open.Content = u"Resume"

                def _make_resume_handler(path):
                    def _handler(s, ev):
                        self._resume_history_session(path)
                    return _handler
                btn_open.Click += _make_resume_handler(fpath)
                act_stack.Children.Add(btn_open)

                def _icon_btn(glyph, tip, fg_key):
                    b = Button()
                    b.Style = self.win.FindResource("T3.Button.Ghost")
                    b.Height = 28
                    b.Width = 28
                    b.MinWidth = 0
                    b.Padding = Thickness(0)
                    b.Margin = Thickness(4, 0, 0, 0)
                    b.ToolTip = tip
                    ic = TextBlock()
                    ic.Text = glyph
                    ic.FontFamily = FontFamily("Segoe MDL2 Assets")
                    ic.FontSize = 11.5
                    ic.Foreground = self.win.FindResource(fg_key)
                    b.Content = ic
                    return b

                # Rename (inline, in place of the title)
                btn_ren = _icon_btn(_ICON_EDIT, u"Rename this conversation",
                                    "T3.TextSecondary")

                def _make_ren_handler(path, title_tb, stack):
                    def _handler(s, ev):
                        self._begin_history_rename(path, title_tb, stack)
                    return _handler
                btn_ren.Click += _make_ren_handler(fpath, tb_title, info_stack)
                act_stack.Children.Add(btn_ren)

                btn_del = _icon_btn(u"\uE74D", u"Delete this session",
                                    "T3.Danger.Text")

                def _make_del_handler(path):
                    def _handler(s, ev):
                        try:
                            chat_sessions.delete_session(path)
                            self._render_history_sessions()
                        except Exception:
                            pass
                    return _handler
                btn_del.Click += _make_del_handler(fpath)
                act_stack.Children.Add(btn_del)

                grid.Children.Add(act_stack)
                card.Child = grid
                self.win.history_sessions_panel.Children.Add(card)
        except Exception as ex:
            logger.debug(u"_render_history_sessions error: {}".format(_exc_text(ex)))

    def _begin_history_rename(self, path, title_tb, stack):
        """Swap a card's title for an edit box. Enter or leaving the box
        saves (config.chat_sessions.rename_session); Esc cancels. UI THREAD."""
        try:
            from System.Windows.Controls import TextBox
            from System.Windows.Input import Key
            old = title_tb.Text or u""
            box = TextBox()
            try:
                box.Style = self.win.FindResource("T3.TextBox")
            except Exception:
                pass
            box.Text = old
            box.ToolTip = u"Enter to save · Esc to cancel"
            state = {"done": False}

            def _finish(save):
                if state["done"]:
                    return
                state["done"] = True
                new = (box.Text or u"").strip()
                try:
                    stack.Children.Remove(box)
                    title_tb.Visibility = Visibility.Visible
                except Exception:
                    pass
                if save and new and new != old:
                    try:
                        from config import chat_sessions
                        if chat_sessions.rename_session(path, new):
                            title_tb.Text = chat_sessions.clean_title(new)
                            title_tb.ToolTip = title_tb.Text
                    except Exception as rx:
                        logger.debug(u"rename session error: {}".format(
                            _exc_text(rx)))

            def _keydown(s, ev):
                if ev.Key in (Key.Return, Key.Enter):
                    _finish(True)
                    ev.Handled = True
                elif ev.Key == Key.Escape:
                    _finish(False)
                    ev.Handled = True

            def _lost(s, ev):
                _finish(True)

            box.KeyDown += _keydown
            box.LostKeyboardFocus += _lost
            title_tb.Visibility = Visibility.Collapsed
            stack.Children.Insert(0, box)
            box.Focus()
            box.SelectAll()
        except Exception as ex:
            logger.debug(u"_begin_history_rename error: {}".format(_exc_text(ex)))

    def _resume_history_session(self, path):
        """Resume an archived session; it keeps its id, notes and file.

        Read AFTER the open chat is archived: when the open chat IS this
        session, that archive just received its newest turns.
        """
        try:
            if not path or not os.path.exists(path):
                return

            if getattr(self.win, '_persisted_msgs', None):
                self._archive_current_session()

            with io.open(path, 'r', encoding='utf-8') as f:
                sess_data = json.load(f)

            self.win._reset_session_state(clear_transcript=True)

            msgs = [m for m in (sess_data.get("messages") or [])
                    if isinstance(m, dict)
                    and not _is_stale_tool_call_blob(m.get("role", ""),
                                                     m.get("content", ""))]
            summary = sess_data.get("summary", u"")

            self.win._persisted_msgs = list(msgs)
            self._set_session({'session_id': sess_data.get("id"),
                               'session_notes': sess_data.get("session_notes"),
                               'session_path': path})

            viet = _ui_viet()
            self.win._conversation_history, self.win._history_summary = _fit_restored(
                msgs, summary, viet=viet)

            save_chat_history(self.win._doc_key, self.win._persisted_msgs,
                              summary=self.win._history_summary,
                              session=self._session_state())

            # Bubbles are capped for rendering only (see MAX_REPLAY_BUBBLES).
            replay = msgs[-MAX_REPLAY_BUBBLES:]
            hidden = len(msgs) - len(replay)
            if hidden:
                self.win._append_bot_message(
                    (u"{} tin nhắn cũ hơn không hiển thị ở đây nhưng vẫn nằm "
                     u"trong ngữ cảnh.".format(hidden)) if viet else
                    (u"{} older message{} not shown here, still part of the "
                     u"context.".format(hidden,
                                        u"s are" if hidden != 1 else u" is")),
                    icon=_ICON_INFO, icon_color=_ICON_SLATE, actions=False)
            for m in replay:
                role = m.get("role", "")
                content = m.get("content", "")
                if not content:
                    continue
                if role == "user":
                    self.win._append_user_message(content)
                elif role == "assistant":
                    self.win._append_bot_message(content)

            self.win._update_welcome_greeting()
            self.tab_chat_clicked()
        except Exception as ex:
            logger.debug(u"_resume_history_session error: {}".format(_exc_text(ex)))

    def clear_all_history_clicked(self, sender=None, e=None):
        """Delete every conversation the History list shows, after a count confirm.

        The list can span every project or a search, hence the confirm.
        """
        try:
            sessions = self._get_history_sessions()
            if not sessions:
                return
            from System.Windows import (MessageBox, MessageBoxButton,
                                        MessageBoxImage, MessageBoxResult)
            n = len(sessions)
            answer = MessageBox.Show(
                u"Delete the {} conversation{} shown in this list? This cannot "
                u"be undone.".format(n, u"" if n == 1 else u"s"),
                u"Clear chat history", MessageBoxButton.YesNo,
                MessageBoxImage.Warning, MessageBoxResult.No)
            if answer != MessageBoxResult.Yes:
                return
            from config import chat_sessions
            for sess in sessions:
                try:
                    chat_sessions.delete_session(sess.get('path'))
                except Exception:
                    pass
            self._render_history_sessions()
        except Exception as ex:
            logger.debug(u"clear_all_history_clicked error: {}".format(_exc_text(ex)))
