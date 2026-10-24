# -*- coding: utf-8 -*-
"""
Assistant commands — the deterministic, no-LLM chat commands _route_input
tries first: /memory + "remember ...", /train, /toolhub, /skills (incl.
"install skill from <repo link>").

Methods of T3LabAssistantWindow, moved here by composition (a CLR-derived
window cannot take a Python mixin): the window builds one
`self._commands = AssistantCommands(self)` and keeps every attribute;
`self.win` below IS that window. XAML handlers keep a one-line stub on
the window that delegates here.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Commands"

import re

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
from System import Action

from GUI.assistant._log import logger
from GUI.assistant.constants import (_ICON_AMBER, _ICON_GREEN, _ICON_INFO,
                                     _ICON_LIST, _ICON_REFRESH, _ICON_SLATE,
                                     _ICON_SUCCESS, _ICON_SYNC, _ICON_WARNING,
                                     _MEM_SAVE_RX)
from GUI.assistant.lang import _is_viet_text


class AssistantCommands(object):
    """Deterministic chat commands that never reach the LLM."""

    def __init__(self, window):
        self.win = window   # the T3LabAssistantWindow that owns this helper

    # NOTE (2026-07-28): the DB-only keyword fast path was deleted (it hijacked
    # unrelated queries); the agent answers from tools ("NUMBERS MUST COME FROM
    # TOOLS" in agent_loop._AGENT_PROMPT).

    # ─── Persistent memory: /memory command + "remember ..." triggers ──

    def _try_memory_command(self, raw):
        """Deterministic persistent-memory handling. WORKER THREAD.

        /memory (show / forget N / clear) and "remember that ..." / "nhớ là ..."
        without an LLM (works offline). Returns True when fully handled here
        (UI updated, busy released).
        """
        try:
            text = (raw or u'').strip()
            if not text:
                return False
            from Intelligence import assistant_memory
            pid = None
            try:
                from config.project_store import ProjectStore
                pid = ProjectStore().get_active_project_id()
            except Exception:
                pass
            viet = _is_viet_text(text)
            low = text.lower()

            msg = None
            icon, color = _ICON_SUCCESS, _ICON_GREEN

            notes = list(getattr(self.win, '_session_notes', None) or [])

            # ── /memory [show | note X | forget N | clear] ──
            m = re.match(r'^/memory(?:\s+(.*))?$', low, re.S)
            if m:
                sub = (m.group(1) or u'').strip()
                # Same tail with its original casing (a note is stored as typed).
                raw_sub = (re.match(r'^/memory(?:\s+(.*))?$', text,
                                    re.S | re.I).group(1) or u'').strip()
                if sub in (u'', u'show', u'list'):
                    msg = assistant_memory.format_memory_report(
                        pid, viet, session_notes=notes)
                    icon, color = _ICON_LIST, _ICON_SLATE
                elif re.match(u'^(?:note|session|ghi chú|ghi chu)(?:\\s|$)', sub):
                    # A note for THIS conversation only (session_context).
                    from Intelligence import session_context as _sc
                    note = re.sub(u'^(?:note|session|ghi chú|ghi chu)\\s*', u'',
                                  raw_sub, flags=re.I).strip()
                    notes, ok, note_msg = _sc.add_note(notes, note)
                    if ok:
                        self.win._sessions._save_session_notes(notes)
                        msg = (u'Đã ghi chú cho cuộc hội thoại này: {}'.format(
                            _sc.clean_note(note)) if viet else note_msg)
                    else:
                        msg = (u'Gõ `/memory note <nội dung>` — ghi chú chỉ áp '
                               u'dụng cho cuộc hội thoại này.' if viet else
                               u'Send `/memory note <text>` — the note applies '
                               u'to this conversation only.')
                        icon, color = _ICON_INFO, _ICON_SLATE
                elif sub.startswith((u'forget', u'xoa', u'xóa')):
                    # N = fact #N; "S<n>" = this chat's note n; else by content
                    # (long-term facts first, then this chat's notes).
                    arg = re.sub(u'^(?:forget|xóa|xoa)\\s*', u'', sub).strip()
                    from Intelligence import session_context as _sc
                    if re.match(r'^s\d+$', arg):
                        notes, removed = _sc.forget_note(notes, arg)
                        ok = removed is not None
                        if ok:
                            self.win._sessions._save_session_notes(notes)
                    elif re.match(r'^\d+$', arg):
                        ok, removed = assistant_memory.remove_fact(
                            int(arg), pid)
                    elif arg:
                        ok, removed = assistant_memory.forget_fact(arg, pid)
                        if not ok and notes:
                            notes, removed = _sc.forget_note(notes, arg)
                            ok = removed is not None
                            if ok:
                                self.win._sessions._save_session_notes(notes)
                    else:
                        ok, removed = False, None
                    if ok:
                        msg = (u'Đã xóa ghi nhớ: "{}"'.format(removed) if viet
                               else u'Forgot: "{}"'.format(removed))
                    else:
                        msg = (u'Không tìm thấy mục đó — gõ `/memory` để xem '
                               u'danh sách đánh số.' if viet else
                               u'No such item — type `/memory` to see the '
                               u'numbered list.')
                        icon, color = _ICON_WARNING, _ICON_AMBER
                elif sub.startswith(u'clear') and sub[len(u'clear'):].strip() in (
                        u'session', u'notes', u'note', u'ghi chú', u'ghi chu'):
                    n = len(notes)
                    self.win._sessions._save_session_notes([])
                    msg = (u'Đã xóa {} ghi chú của cuộc hội thoại này. Ghi nhớ '
                           u'lâu dài không đổi.'.format(n) if viet else
                           u"Cleared {} note{} of this conversation. Long-term "
                           u"memory is unchanged.".format(n, u'' if n == 1 else u's'))
                    icon, color = _ICON_REFRESH, _ICON_SLATE
                elif sub.startswith(u'clear'):
                    msg, icon, color = self._memory_clear_reply(
                        sub[len(u'clear'):].strip(), pid, viet)
                else:
                    msg = (u'Lệnh memory: `/memory` · '
                           u'`/memory note <nội dung>` (chỉ cuộc hội thoại này) · '
                           u'`/memory forget <số|S số|nội dung>` · '
                           u'`/memory clear` · `/memory clear global` · '
                           u'`/memory clear session`'
                           if viet else
                           u'Memory commands: `/memory` · '
                           u'`/memory note <text>` (this conversation only) · '
                           u'`/memory forget <number|S number|text>` · '
                           u'`/memory clear` (this project) · '
                           u'`/memory clear global` · `/memory clear session`')
                    icon, color = _ICON_INFO, _ICON_SLATE

            # ── "what do you remember?" ──
            if msg is None:
                _shows = (u'what do you remember', u'what have you remembered',
                          u'what do you know about me', u'ban nho gi',
                          u'bạn nhớ gì', u'ban nho nhung gi',
                          u'bạn nhớ những gì', u'ban dang nho gi',
                          u'bạn đang nhớ gì')
                if low.rstrip(u'? .').strip() in _shows:
                    msg = assistant_memory.format_memory_report(
                        pid, viet, session_notes=notes)
                    icon, color = _ICON_LIST, _ICON_SLATE

            # ── "remember that X" / "nhớ là X" / "từ nay X" ──
            if msg is None:
                m2 = _MEM_SAVE_RX.match(text)
                if m2:
                    fact = (m2.group(1) or u'').strip()
                    fact = re.sub(r'^(?:that\s+|r[aằ]ng\s+|l[aà]\s+)', u'',
                                  fact, flags=re.IGNORECASE).strip(u' .!')
                    if fact:
                        _gl = (u'always', u'luôn', u'luon', u'every project',
                               u'all projects', u'mọi dự án', u'moi du an',
                               u'from now on', u'từ nay', u'tu nay',
                               u'prefer', u'mọi project', u'moi project',
                               u'mọi model', u'moi model')
                        scope = ('global' if any(k in low for k in _gl)
                                 else 'project')
                        ok, note = assistant_memory.add_fact(
                            fact, scope=scope, project_id=pid)
                        if ok and viet:
                            msg = u'Đã ghi nhớ ({}): {}'.format(
                                u'mọi dự án' if scope == 'global'
                                else u'dự án này', fact)
                        elif ok:
                            msg = note
                        else:
                            msg = note
                            icon, color = _ICON_WARNING, _ICON_AMBER

            if msg is None:
                return False

            self.win._log_activity(u'Memory: {}'.format(text))

            def _show(_m=msg, _i=icon, _c=color):
                self.win._hide_typing_indicator()
                self.win._append_bot_message(_m, icon=_i, icon_color=_c)
                self.win._add_to_history('assistant', _m)
                self.win._set_busy(False)
            self.win.Dispatcher.Invoke(Action(_show))
            return True
        except Exception as ex:
            logger.debug('_try_memory_command error: {}'.format(ex))
            return False

    @staticmethod
    def _memory_clear_reply(arg, pid, viet=False):
        """`/memory clear [global]` — clear ONE scope. Returns (msg, icon, color).

          /memory clear          → this project's facts only; global kept
          /memory clear global   → global facts only; projects kept
          no project active      → nothing cleared; reply says how to clear global
        (It used to wipe GLOBAL memory too.) The reply names scope and count.
        """
        from Intelligence import assistant_memory as _am
        arg = (arg or u'').strip().lower()
        wants_global = arg in (u'global', u'all projects', u'all-projects',
                               u'toan cuc', u'toàn cục', u'chung')
        if wants_global:
            n = _am.clear_scope(_am.GLOBAL_SCOPE)
            kept = _am.count_scope(_am.PROJECT_SCOPE, pid) if pid else 0
            if viet:
                msg = (u'Đã xóa {} ghi nhớ chung (áp dụng cho mọi dự án).'
                       .format(n))
                if pid:
                    msg += u' Giữ nguyên {} ghi nhớ của dự án này.'.format(kept)
            else:
                msg = (u'Cleared {} global fact{} (the ones that apply to '
                       u'every project).'.format(n, u'' if n == 1 else u's'))
                if pid:
                    msg += (u" This project's {} fact{} {} kept.".format(
                        kept, u'' if kept == 1 else u's',
                        u'was' if kept == 1 else u'were'))
            return msg, _ICON_REFRESH, _ICON_SLATE
        n_global = _am.count_scope(_am.GLOBAL_SCOPE)
        if not pid:
            # never clear global implicitly — it must be asked for by name
            if viet:
                msg = (u'Chưa chọn dự án nào nên không có ghi nhớ dự án để '
                       u'xóa. Gõ `/memory clear global` để xóa {} ghi nhớ '
                       u'chung (áp dụng cho mọi dự án).'.format(n_global))
            else:
                msg = (u'No project is active, so there is no project memory '
                       u'to clear. Send `/memory clear global` to clear the {} '
                       u'global fact{} that apply to every project.'.format(
                           n_global, u'' if n_global == 1 else u's'))
            return msg, _ICON_INFO, _ICON_SLATE
        n = _am.clear_scope(_am.PROJECT_SCOPE, pid)
        if viet:
            msg = (u'Đã xóa {} ghi nhớ của dự án này. Giữ nguyên {} ghi nhớ '
                   u'chung — gõ `/memory clear global` nếu muốn xóa cả '
                   u'chúng.'.format(n, n_global))
        else:
            msg = (u"Cleared {} fact{} from this project's memory. The {} "
                   u"global fact{} (every project) {} kept — send "
                   u"`/memory clear global` to clear {} too.".format(
                       n, u'' if n == 1 else u's', n_global,
                       u'' if n_global == 1 else u's',
                       u'was' if n_global == 1 else u'were',
                       u'it' if n_global == 1 else u'them'))
        return msg, _ICON_REFRESH, _ICON_SLATE

    # ─── Self-study: /train command ──

    def _try_train_command(self, raw):
        """Deterministic `/train` handling. WORKER THREAD.

        `/train` → dataset stats; `/train now` → launch the detached LoRA
        fine-tune (runs OUTSIDE Revit, tools/train/). Returns True when handled.
        """
        try:
            text = (raw or u'').strip()
            low = text.lower()
            m = re.match(r'^/train(?:\s+(.*))?$', low, re.S)
            if not m:
                return False
            sub = (m.group(1) or u'').strip()
            viet = _is_viet_text(text)

            from Intelligence.learning import trainer as _trainer
            st = _trainer.dataset_stats()
            count = st.get('count', 0)
            last = _trainer.last_train()

            icon, color = _ICON_INFO, _ICON_SLATE
            if sub in (u'exemplars', u'exemplar', u'portable'):
                # PORTABLE few-shot layer (no GPU): teacher data → git-tracked
                # teacher_exemplars.json, usable by any machine's local model.
                try:
                    from Intelligence.learning import exemplars as _ex
                    res = _ex.promote_from_dataset()
                    n = res.get('count', 0)
                    if res.get('status') == 'ok':
                        msg = (u'Đã dựng {} ví dụ mẫu portable vào '
                               u'`teacher_exemplars.json`. Commit file này để '
                               u'mọi máy dùng model local đều trả lời theo phong '
                               u'cách đã dạy — không cần GPU/train.'.format(n)
                               if viet else
                               u'Built {} portable exemplars into '
                               u'`teacher_exemplars.json`. Commit it so every '
                               u'machine\'s local model answers in the taught '
                               u'style — no GPU/training needed.'.format(n))
                        icon, color = _ICON_SYNC, _ICON_SLATE
                    else:
                        msg = (u'Không dựng được exemplars: {}'.format(
                            res.get('status')) if viet else
                            u'Could not build exemplars: {}'.format(
                                res.get('status')))
                        icon, color = _ICON_WARNING, _ICON_AMBER
                except Exception as _ex_err:
                    msg = u'Could not build exemplars: {}'.format(_ex_err)
                    icon, color = _ICON_WARNING, _ICON_AMBER
            elif sub in (u'now', u'chay', u'chạy', u'start'):
                ok, note = _trainer.launch()
                if ok:
                    msg = (u'Đã khởi động huấn luyện nền ({} mẫu). Quá trình '
                           u'chạy ngoài Revit và có thể mất hàng giờ; xem '
                           u'`tools/train/README.md`.'.format(count) if viet else
                           u'Started background training on {} examples. It runs '
                           u'outside Revit and can take hours; see '
                           u'`tools/train/README.md`.'.format(count))
                    icon, color = _ICON_SYNC, _ICON_SLATE
                else:
                    msg = (u'Không khởi động được: {}'.format(note) if viet
                           else u'Could not launch training: {}'.format(note))
                    icon, color = _ICON_WARNING, _ICON_AMBER
            else:
                by_src = u', '.join(u'{}: {}'.format(k, v)
                                    for k, v in sorted(
                                        (st.get('by_source') or {}).items()))
                last_line = (u'Lần cuối: {} ({} mẫu)'.format(
                    last.get('trained_at'), last.get('examples'))
                    if last else (u'Chưa huấn luyện lần nào' if viet
                                  else u'Never trained yet'))
                if viet:
                    msg = (u'**Tự học — dữ liệu huấn luyện**\n\n'
                           u'- Tổng số mẫu: **{}**\n- Nguồn: {}\n- {}\n\n'
                           u'Gõ `/train now` để chạy fine-tune nền (cần GPU, '
                           u'chạy ngoài Revit).'.format(
                               count, by_src or u'(trống)', last_line))
                else:
                    msg = (u'**Self-study — training data**\n\n'
                           u'- Total examples: **{}**\n- Sources: {}\n- {}\n\n'
                           u'Type `/train now` to run a background fine-tune '
                           u'(needs a GPU, runs outside Revit).'.format(
                               count, by_src or u'(empty)', last_line))

            self.win._log_activity(u'Train: {}'.format(text))

            def _show(_m=msg, _i=icon, _c=color):
                self.win._hide_typing_indicator()
                self.win._append_bot_message(_m, icon=_i, icon_color=_c)
                self.win._add_to_history('assistant', _m)
                self.win._set_busy(False)
            self.win.Dispatcher.Invoke(Action(_show))
            return True
        except Exception as ex:
            logger.debug('_try_train_command error: {}'.format(ex))
            return False

    def _toolhub_upload_soon(self):
        """Send audit summaries to the Tool Hub after a turn: a daemon thread,
        a no-op when signed out, at most once a minute (toolhub.events)."""
        try:
            import threading

            def _go():
                try:
                    from Intelligence.toolhub import events as _ev
                    from Intelligence.toolhub.client import ToolHubClient
                    _ev.upload(ToolHubClient())
                except Exception:
                    pass
            t = threading.Thread(target=_go, name='T3LabToolHubUpload')
            t.daemon = True
            t.start()
        except Exception:
            pass

    def _try_toolhub_command(self, raw):
        """`/toolhub …` and `/save-tool …` (Intelligence.toolhub.commands). WORKER THREAD."""
        try:
            text = (raw or u'').strip()
            low = text.lower()
            from Intelligence.toolhub import commands as toolhub_commands
            if not (low.startswith(u'/toolhub') or low.startswith(u'/save-tool')):
                # bare e-mail / code typed after a /toolhub prompt
                reply = toolhub_commands.handle_followup(text)
                if reply is None:
                    return False
                self.win._post_bot(reply, _ICON_INFO, _ICON_SLATE)
                return True
            if low == u'/toolhub':
                # bare command → the Tool Hub tab (friendlier than status text)
                self.win._post_bot(u"Opening Tool Hub in LLMs Setting…",
                               _ICON_INFO, _ICON_SLATE)
                self.win._ui_invoke(
                    lambda: self.win._open_llm_settings(tab='toolhub'), wait=False)
                return True
            if low.startswith(u'/save-tool'):
                reply = toolhub_commands.handle_save(
                    text, getattr(self.win, '_last_turn_calls', None) or [],
                    getattr(self.win, '_last_turn_user_text', u'') or u'')
            else:
                reply = toolhub_commands.handle(text)
            if reply is None:
                return False
            self.win._post_bot(reply, _ICON_INFO, _ICON_SLATE)
            return True
        except Exception as ex:
            logger.debug('_try_toolhub_command error: {}'.format(ex))
            return False

    def _try_skills_command(self, raw):
        """Install / update skills from a GitHub repo. WORKER THREAD.

        `/skills [install <url> | update | remove <id>]` or the natural request
        ("cài skill từ https://github.com/owner/repo"). No LLM needed.
        Returns True when fully handled here.
        """
        try:
            text = (raw or u'').strip()
            if not text:
                return False
            from Intelligence import skill_installer as installer
            parsed = installer.parse_skills_command(text)
            if not parsed:
                return False
            action, arg = parsed
            viet = _is_viet_text(text)

            if action == 'help':
                msg = (u'Lệnh skills: `/skills` · '
                       u'`/skills install <link GitHub>` · `/skills update` · '
                       u'`/skills remove <id>`\nHoặc chỉ cần nói: "cài skill '
                       u'từ https://github.com/owner/repo".' if viet else
                       u'Skills commands: `/skills` · '
                       u'`/skills install <github link>` · `/skills update` · '
                       u'`/skills remove <id>`\nOr just say: "install skills '
                       u'from https://github.com/owner/repo".')
                self.win._post_bot(msg, _ICON_INFO, _ICON_SLATE)
                return True

            if action == 'list':
                msg = installer.format_inventory(
                    installer.list_installed(), viet)
                self.win._post_bot(msg, _ICON_LIST, _ICON_SLATE)
                return True

            if action == 'remove':
                if not arg:
                    self.win._post_bot(
                        u'Cần id skill: `/skills remove <id>`' if viet else
                        u'Missing skill id: `/skills remove <id>`',
                        _ICON_WARNING, _ICON_AMBER)
                    return True
                ok, note = installer.remove_installed(arg)
                self._rescan_skills()
                self.win._post_bot(note,
                               _ICON_SUCCESS if ok else _ICON_WARNING,
                               _ICON_GREEN if ok else _ICON_AMBER)
                return True

            # ── install / update: network work, so say so first ──
            if action == 'install':
                notice = (u'Đang tải skill từ {}…'.format(arg.get('url'))
                          if viet else
                          u'Downloading skills from {}…'.format(arg.get('url')))
            else:
                notice = (u'Đang cập nhật các skill đã cài…' if viet else
                          u'Updating installed skills…')
            self.win._log_activity(u'Skills: {}'.format(text))

            def _notice(_m=notice):
                self.win._append_bot_message(_m, icon=_ICON_SYNC,
                                         icon_color=_ICON_SLATE)
            self.win.Dispatcher.Invoke(Action(_notice))

            if action == 'install':
                report = installer.install_from_github(arg)
            else:
                report = installer.update_all()

            self._rescan_skills()
            self.win._post_bot(installer.format_report(report, viet),
                           _ICON_SUCCESS if (report.get('installed') or
                                             report.get('updated'))
                           else _ICON_WARNING,
                           _ICON_GREEN if (report.get('installed') or
                                           report.get('updated'))
                           else _ICON_AMBER)
            return True
        except Exception as ex:
            logger.debug('_try_skills_command error: {}'.format(ex))
            return False

    def _rescan_skills(self):
        """Re-read the skill folders so new skills reach /slash immediately."""
        try:
            from Intelligence.skills_engine import get_skills_engine
            get_skills_engine().scan()
        except Exception as ex:
            logger.debug('_rescan_skills error: {}'.format(ex))
