# -*- coding: utf-8 -*-
"""
Assistant tool cards — what the chat shows for each tool call (card,
collapsed group row, result summary, select-in-Revit) and the per-call
confirmation gate (ToolGate + the blocking confirm card).

Methods of T3LabAssistantWindow, moved here by composition (a CLR-derived
window cannot take a Python mixin): the window builds one
`self._tool_cards = AssistantToolCards(self)` and keeps every attribute;
`self.win` below IS that window. XAML handlers keep a one-line stub on
the window that delegates here.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Tool Cards"

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
import System.Windows
from System import Action
from System.Threading import Thread, ThreadStart
from System.Windows import Visibility

from GUI.assistant._log import logger
from GUI.assistant.constants import (_CONFIRM_TIMEOUT, _ICON_AMBER, _ICON_EDIT,
                                     _ICON_STOP, _ICON_WARNING)
from GUI.assistant.deps import ContextScout
from GUI.assistant.lang import _ui_viet
from GUI.assistant.text import _call_summary, _exc_text, _json_text
from GUI.assistant.theme import _bind_bg, _bind_border, _bind_fg, _trgb


class AssistantToolCards(object):
    """Tool-call cards, their collapsed group row, and the per-call confirm gate."""

    def __init__(self, window):
        self.win = window   # the T3LabAssistantWindow that owns this helper

    # ─── Tool-call cards ───────────────────────────────────────────────────────

    # ── Compact tool-call group ──────────────────────────────────────────────
    # "Show tool calls" off folds a turn's cards into ONE collapsible row
    # ("✓ 5 tools · 3.2s") rather than hiding them, so failures stay findable.

    def _chat_verbosity(self):
        """(show_tool_calls, show_thinking) — never raises."""
        try:
            from config.settings import get_settings
            s = get_settings()
            return (s.is_show_tool_calls_enabled(),
                    s.is_show_thinking_enabled())
        except Exception:
            return (True, True)

    def _reset_tool_group(self):
        """Start a fresh compact group (on each new user message)."""
        self.win._tool_group = None

    def _ensure_tool_group(self):
        """The collapsible container for this turn, created on first tool."""
        group = getattr(self.win, '_tool_group', None)
        if group is not None:
            return group
        from System.Windows.Controls import Border, TextBlock, StackPanel, Orientation
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Media import FontFamily
        from System.Windows.Input import Cursors

        shell = Border()
        _bind_bg(shell, 'SelectedBg')
        _bind_border(shell, 'CardBorder')
        shell.BorderThickness = Thickness(1)
        shell.CornerRadius    = CornerRadius(8)
        shell.Padding         = Thickness(12, 8, 12, 8)
        shell.Margin          = Thickness(0, 0, 8, 10)

        outer = StackPanel()

        head = StackPanel()
        head.Orientation = Orientation.Horizontal
        head.Cursor      = Cursors.Hand

        chevron = TextBlock()
        chevron.Text       = u""          # MDL2 ChevronRight — collapsed
        chevron.FontFamily = FontFamily(u"Segoe MDL2 Assets")
        chevron.FontSize   = 10
        _bind_fg(chevron, 'Faint')
        chevron.Margin     = Thickness(0, 2, 8, 0)

        label = TextBlock()
        label.FontSize = 11.5
        _bind_fg(label, 'Muted')

        head.Children.Add(chevron)
        head.Children.Add(label)
        outer.Children.Add(head)

        body = StackPanel()
        body.Margin     = Thickness(0, 8, 0, 0)
        body.Visibility = Visibility.Collapsed
        outer.Children.Add(body)

        group = {'shell': shell, 'label': label, 'chevron': chevron,
                 'body': body, 'total': 0, 'failed': 0, 'seconds': 0.0,
                 'running': u''}

        def _toggle(sender, e, _g=group):
            expanded = _g['body'].Visibility == Visibility.Visible
            _g['body'].Visibility = (Visibility.Collapsed if expanded
                                     else Visibility.Visible)
            _g['chevron'].Text = u"" if expanded else u""   # right / down
            self.win._scroll_to_bottom()

        head.MouseLeftButtonUp += _toggle

        shell.Child = outer
        self.win.chat_history_panel.Children.Add(shell)
        self.win._tool_group = group
        self._refresh_tool_group_label(group)
        return group

    def _refresh_tool_group_label(self, group):
        running = group.get('running')
        if running:
            group['label'].Text = (
                u"Đang chạy: {}… ({} tool)".format(running, group['total'])
                if _ui_viet() else
                u"Running: {}… ({} tools)".format(running, group['total']))
            return
        done = group['total'] - group['failed']
        parts = [u"{} tool".format(done) if _ui_viet()
                 else u"{} tool{}".format(done, u"" if done == 1 else u"s")]
        if group['failed']:
            parts.append(u"{} lỗi".format(group['failed']) if _ui_viet()
                         else u"{} failed".format(group['failed']))
        parts.append(u"{0:.1f}s".format(group['seconds']))
        group['label'].Text = u"{}  {}".format(
            u"" if group['failed'] else u"", u" · ".join(parts))

    def _append_tool_card(self, name, args):
        """Add a tool-call status card to the chat. UI thread only.

        Returns a handle dict for _update_tool_card, or None on failure.
        """
        try:
            from System.Windows.Controls import Border, TextBlock, StackPanel, Orientation
            from System.Windows import Thickness, CornerRadius, TextWrapping
            from System.Windows.Media import SolidColorBrush, Color, FontFamily

            card = Border()
            _bind_bg(card, 'SelectedBg')
            _bind_border(card, 'CardBorder')
            card.BorderThickness = Thickness(1)
            card.CornerRadius    = CornerRadius(8)
            card.Padding         = Thickness(12, 8, 12, 8)
            # Left margin lines the card up with bot bubbles (avatar 36 + 10).
            card.Margin          = Thickness(0, 0, 8, 10)

            panel = StackPanel()

            head = StackPanel()
            head.Orientation = Orientation.Horizontal

            status = TextBlock()
            status.Text       = u""   # MDL2 Sync — running
            status.FontFamily = FontFamily(u"Segoe MDL2 Assets")
            status.FontSize   = 12
            _bind_fg(status, 'Blue')
            status.Margin     = Thickness(0, 1, 8, 0)

            title = TextBlock()
            title.Text       = name
            title.FontFamily = FontFamily(u"Consolas")
            title.FontSize   = 12
            title.FontWeight = System.Windows.FontWeights.SemiBold
            _bind_fg(title, 'Ink')

            dur = TextBlock()
            dur.Text       = u"running…"
            dur.FontSize   = 11
            _bind_fg(dur, 'Faint')
            dur.Margin     = Thickness(8, 1, 0, 0)

            head.Children.Add(status)
            head.Children.Add(title)
            head.Children.Add(dur)
            panel.Children.Add(head)

            # What the call does ("Delete 3 elements (id …)", the purpose of
            # a code run), not its JSON; the exact request is the tooltip.
            headline, _details = _call_summary(name, args)
            args_tb = TextBlock()
            args_tb.Text         = headline
            args_tb.ToolTip      = _json_text(args)[:4000]
            args_tb.FontSize     = 11
            _bind_fg(args_tb, 'Muted')
            args_tb.TextWrapping = TextWrapping.Wrap
            args_tb.Margin       = Thickness(20, 2, 0, 0)
            panel.Children.Add(args_tb)

            result_tb = TextBlock()
            result_tb.FontSize     = 11
            _bind_fg(result_tb, 'BotText')
            result_tb.TextWrapping = TextWrapping.Wrap
            result_tb.Margin       = Thickness(20, 3, 0, 0)
            result_tb.Visibility   = Visibility.Collapsed
            panel.Children.Add(result_tb)

            card.Child = panel
            group = None
            if not self._chat_verbosity()[0]:
                try:
                    group = self._ensure_tool_group()
                except Exception as ex:
                    logger.debug(u"tool group error: {}".format(_exc_text(ex)))
                    group = None
            if group is not None:
                # Inside the group the outer chrome would double up.
                card.BorderThickness = Thickness(0)
                card.Background      = None
                card.Padding         = Thickness(0, 0, 0, 0)
                card.Margin          = Thickness(0, 0, 0, 8)
                group['body'].Children.Add(card)
                group['total']  += 1
                group['running'] = (headline if len(headline) <= 70
                                    else headline[:69] + u"…")
                self._refresh_tool_group_label(group)
            else:
                self.win.chat_history_panel.Children.Add(card)
            self.win._scroll_to_bottom()
            return {"card": card, "status": status, "dur": dur,
                    "result": result_tb, "group": group}
        except Exception as ex:
            logger.debug(u"_append_tool_card error: {}".format(_exc_text(ex)))
            return None

    def _update_tool_card(self, handle, ok, seconds, result):
        """Mark a tool card done/failed and show the result summary. UI thread."""
        if not handle:
            return
        try:
            from System.Windows.Media import SolidColorBrush, Color

            status = handle["status"]
            # Declined on its approval card: neither success nor failure.
            declined = isinstance(result, dict) and bool(result.get("cancelled"))
            if declined:
                status.Text       = _ICON_STOP
                _bind_fg(status, 'Muted')
            elif ok:
                status.Text       = u""   # MDL2 CheckMark
                _bind_fg(status, 'Success')
            else:
                status.Text       = u""   # MDL2 Cancel
                _bind_fg(status, 'Danger')

            handle["dur"].Text = (u"skipped" if declined
                                  else u"{0:.1f}s".format(seconds))

            group = handle.get("group")
            if group is not None:
                group['seconds'] += float(seconds or 0.0)
                group['running']  = u''
                if not ok:
                    group['failed'] += 1
                    # A failure never stays folded away.
                    group['body'].Visibility = Visibility.Visible
                    group['chevron'].Text    = u""
                self._refresh_tool_group_label(group)

            # _json_text, not json.dumps: Windows code-page bytes in a Revit
            # result break the encoder (and the u"{}".format fallback).
            res_s = _json_text(result)
            rt = handle["result"]
            # One-line summary; full JSON in the tooltip.
            summary = self._summarize_tool_result(result, res_s)
            rt.Text       = summary[:240] + (u"…" if len(summary) > 240 else u"")
            rt.ToolTip    = res_s[:4000]
            rt.Visibility = Visibility.Visible

            # C1: clickable element-id links → select & zoom in Revit.
            try:
                ids = self._extract_element_ids(result)
                if ok and ids:
                    from System.Windows.Controls import StackPanel, TextBlock, Orientation
                    from System.Windows import Thickness, TextDecorations
                    from System.Windows.Input import Cursors

                    links = StackPanel()
                    links.Orientation = Orientation.Horizontal
                    links.Margin = Thickness(20, 4, 0, 0)

                    def _mk_link(label, id_list):
                        tb = TextBlock()
                        tb.Text           = label
                        tb.FontSize       = 11
                        _bind_fg(tb, 'Blue')
                        tb.TextDecorations = TextDecorations.Underline
                        tb.Cursor         = Cursors.Hand
                        tb.Margin         = Thickness(0, 0, 10, 0)
                        tb.ToolTip        = u"Select & zoom in Revit"

                        def _click(s, e, _ids=list(id_list)):
                            self._select_in_revit_async(_ids)

                        tb.MouseLeftButtonUp += _click
                        return tb

                    for eid in ids[:6]:
                        links.Children.Add(_mk_link(u"#{}".format(eid), [eid]))
                    if len(ids) > 1:
                        links.Children.Add(
                            _mk_link(u"select all {}".format(len(ids)), ids))
                    handle["card"].Child.Children.Add(links)
            except Exception:
                pass

            self.win._scroll_to_bottom()
        except Exception as ex:
            logger.debug(u"_update_tool_card error: {}".format(_exc_text(ex)))

    @staticmethod
    def _summarize_tool_result(result, fallback):
        """Short human line for the card body from the tool's count/summary
        fields (total_count, element_counts, message...); else `fallback`."""
        if not isinstance(result, dict):
            return fallback
        if result.get("error"):
            return u"Error: {}".format(result.get("error"))
        if result.get("cancelled"):
            return u"Not run — declined on the approval card"
        parts = []
        msg = result.get("message") or result.get("summary")
        if msg:
            parts.append(u"{}".format(msg))
        # The count fields the agent trusts (see agent_loop rule 10).
        for key, label in ((u"total_count", u"matched"),
                            (u"count", u"count"),
                            (u"row_count", u"rows"),
                            (u"modified_count", u"modified"),
                            (u"created_count", u"created"),
                            (u"deleted_count", u"deleted")):
            if isinstance(result.get(key), int):
                parts.append(u"{} {}".format(result[key], label))
        ec = result.get("element_counts")
        if isinstance(ec, dict) and ec:
            top = sorted(ec.items(), key=lambda kv: kv[1], reverse=True)[:4]
            parts.append(u", ".join(u"{}: {}".format(k, v) for k, v in top))
        if not parts and result.get("success") is True:
            parts.append(u"Done")
        if not parts:
            return fallback
        return u"  ·  ".join(parts)

    @staticmethod
    def _extract_element_ids(result, _limit=60):
        """Collect Revit element ids out of a tool-result dict (C1).

        Knows {'id'|'element_id': n}, {'element_ids'|'created_ids'|'ids': [...]}
        and nested lists of such dicts. Ordered, deduped, capped at _limit.
        """
        out  = []
        seen = set()

        def _add(v):
            try:
                n = int(v)
            except Exception:
                return
            if n > 0 and n not in seen:
                seen.add(n)
                out.append(n)

        def _walk(node, depth):
            if depth > 4 or len(out) >= _limit:
                return
            if isinstance(node, dict):
                for k, v in node.items():
                    lk = u"{}".format(k).lower()
                    if lk in ("id", "element_id", "new_element_id", "new_id",
                              "tag_id", "wall_id", "grid_id", "level_id"):
                        _add(v)
                    elif lk in ("element_ids", "created_ids", "ids", "new_ids",
                                "tag_ids", "wall_ids", "created_element_ids"):
                        if isinstance(v, (list, tuple)):
                            for item in v:
                                _add(item)
                    elif isinstance(v, (dict, list, tuple)):
                        _walk(v, depth + 1)
            elif isinstance(node, (list, tuple)):
                for item in node:
                    _walk(item, depth + 1)

        _walk(result if isinstance(result, dict) else {}, 0)
        return out

    def _select_in_revit_async(self, element_ids):
        """Select + zoom elements from an element-link click (C1).

        WORKER thread: waiting on the ExternalEvent from the UI thread would
        deadlock (the handler needs the UI thread).
        """
        ids = [i for i in (element_ids or [])]
        if not ids:
            return

        def _work():
            res = None
            try:
                from core.server import get_t3labai_server
                srv = get_t3labai_server()
                res = srv._execute_tool('select_elements',
                                        {'element_ids': ids, 'show': True,
                                         'limit': len(ids)})
            except Exception:
                res = None
            # Say WHY the view didn't move (Revit's own "No good view" dialog
            # is suppressed upstream, so silence looks like a dead link).
            try:
                if not isinstance(res, dict):
                    return
                why = res.get('show_skipped') or res.get('show_error')
                if not why:
                    return   # success — the view visibly changed, stay quiet

                def _ui():
                    try:
                        self.win._append_bot_message(
                            u"Couldn't jump to element {}: {}".format(
                                u", ".join(str(i) for i in ids[:3]), why),
                            icon=_ICON_WARNING, icon_color=_ICON_AMBER)
                    except Exception:
                        pass
                self.win.Dispatcher.BeginInvoke(Action(_ui))
            except Exception:
                pass

        t = Thread(ThreadStart(_work))
        t.IsBackground = True
        t.Start()

    # ─── Tool approval gate (shared by the native and legacy paths) ───────────
    # WHAT needs approval: Intelligence.agent_loop.ToolGate (pure, tested);
    # this block only shows the card and blocks on it. Both paths go through
    # _run_gated_tool so their gating cannot differ.

    def _new_tool_gate(self, srv):
        """A fresh ToolGate for ONE request (it remembers the first purge)."""
        from Intelligence.agent_loop import ToolGate
        return ToolGate(
            action_mode=self.win._action_mode(),
            is_destructive=srv.is_destructive,
            preapproved_edits=bool(getattr(self.win, '_plan_preapproved', False)))

    def _run_gated_tool(self, srv, gate, name, args, viet):
        """Gate, then execute, one tool call. WORKER THREAD.

        Returns the result, or agent_loop.declined_result(name) on no / Stop /
        timeout (the model reads it as "skipped, do not retry").
        """
        from Intelligence.agent_loop import GATE_RUN, declined_result
        args, verdict = gate.prepare(name, args)
        if verdict != GATE_RUN:
            if not self._confirm_tool_blocking(name, args, viet, kind=verdict):
                return declined_result(name)
        res = srv._execute_tool(name, args)
        if gate.modifies(name, args):
            self._note_model_changed()
        return res

    @staticmethod
    def _note_model_changed():
        """Model changed: invalidate the cached levels/grids digest."""
        try:
            _inv = getattr(ContextScout, 'invalidate_datums', None)
            if _inv is not None:
                _inv()
        except Exception:
            pass

    def _confirm_tool_blocking(self, name, args, viet,
                               timeout_sec=_CONFIRM_TIMEOUT,
                               kind='destructive'):
        """WORKER thread: show a Confirm/Cancel card and block. True only on
        an explicit Confirm (timeout / Stop / error = declined).

        kind: 'destructive' (red, every mode) or 'edit' ("Ask before edits").
        """
        from Intelligence.agent_loop import PendingDecision
        pending = PendingDecision()
        handle = {"seal": None}

        def _ui():
            try:
                self.win._hide_typing_indicator()
                handle["seal"] = self._append_confirm_card(
                    name, args, pending, viet, kind=kind)
            except Exception as _cx:
                logger.debug(u"confirm card error: {}".format(_exc_text(_cx)))
                pending.decide(False)

        try:
            self.win.Dispatcher.Invoke(Action(_ui))
        except Exception:
            return False

        def _abort():
            # Native loop's flag, or the window-wide one (legacy loop).
            loop = getattr(self.win, '_agent_loop', None)
            if loop is not None and loop.is_cancelled():
                return True
            return self.win._cancelled()

        _loop = getattr(self.win, '_agent_loop', None)
        try:
            if _loop is not None:
                _loop.pause()          # user time is not budget time
            decision = pending.wait(timeout_sec, should_abort=_abort)
        finally:
            if _loop is not None:
                _loop.resume()
        if decision is not True and decision is not False:
            # Seal the card so stale buttons can't approve a finished request.
            _msg = (u"Stopped — skipped" if decision == PendingDecision.STOPPED
                    else u"Expired — skipped")

            def _expire():
                try:
                    if handle.get("seal"):
                        handle["seal"](_msg)
                except Exception:
                    pass
            try:
                self.win.Dispatcher.BeginInvoke(Action(_expire))
            except Exception:
                pass
        return decision is True

    def _make_card_button(self, label, primary=False, tooltip=None):
        """An in-chat card button that follows the Revit theme. UI THREAD.

        T3.Button.Base gives only the shape; colours are theme-token bindings,
        so the button survives a Light/Dark switch on screen.
        """
        from System.Windows.Controls import Button
        from System.Windows import Thickness
        btn = Button()
        try:
            btn.Style = self.win.FindResource("T3.Button.Base")
        except Exception:
            pass
        btn.Content = label
        btn.Margin  = Thickness(0, 0, 8, 8)
        if tooltip:
            btn.ToolTip = tooltip
        if primary:
            _bind_bg(btn, 'Accent')
            _bind_border(btn, 'Accent')
            _bind_fg(btn, 'OnAccent')

            def _enter(s, e, _b=btn):
                _bind_bg(_b, 'AccentHover')
                _bind_border(_b, 'AccentHover')

            def _leave(s, e, _b=btn):
                _bind_bg(_b, 'Accent')
                _bind_border(_b, 'Accent')
        else:
            _bind_bg(btn, 'ControlBg')
            _bind_border(btn, 'ControlBorder')
            _bind_fg(btn, 'Ink')

            def _enter(s, e, _b=btn):
                _bind_bg(_b, 'ControlHoverBg')
                _bind_border(_b, 'ControlHoverBorder')

            def _leave(s, e, _b=btn):
                _bind_bg(_b, 'ControlBg')
                _bind_border(_b, 'ControlBorder')
        btn.MouseEnter += _enter
        btn.MouseLeave += _leave
        return btn

    def _append_confirm_card(self, name, args, pending, viet,
                             kind='destructive'):
        """Confirm/Cancel card for one held tool call. UI thread only.

        Returns seal(text), which disables the card and shows `text`. A click
        decides `pending` (agent_loop.PendingDecision); first answer wins.
        """
        from System.Windows.Controls import (Border, TextBlock, StackPanel,
                                              WrapPanel)
        from System.Windows.Documents import Run
        from System.Windows import Thickness, CornerRadius, TextWrapping
        from System.Windows.Media import SolidColorBrush, Color, FontFamily
        from System.Windows.Input import Cursors

        # What the call does, then the tool name; the exact request (for
        # send_code_to_revit, the code) is the tooltip.
        headline, details = _call_summary(name, args)
        summary = headline + u"".join(u"\n• {}".format(d) for d in details[:8])
        args_tip = u"{} — {}".format(name, _json_text(args)[:4000])

        destructive = (kind == 'destructive')
        card = Border()
        card.BorderThickness = Thickness(1)
        card.CornerRadius    = CornerRadius(8)
        card.Padding         = Thickness(12, 10, 12, 10)
        card.Margin          = Thickness(0, 0, 8, 10)
        panel = StackPanel()

        head = TextBlock()
        head.FontSize     = 12
        head.FontWeight   = System.Windows.FontWeights.SemiBold
        head.TextWrapping = TextWrapping.Wrap
        body = TextBlock()
        body.FontSize     = 11.5
        body.TextWrapping = TextWrapping.Wrap
        body.Margin       = Thickness(0, 4, 0, 8)
        status_tb = TextBlock()
        status_tb.FontSize     = 11.5
        status_tb.TextWrapping = TextWrapping.Wrap
        status_tb.Visibility   = Visibility.Collapsed

        btn_row = WrapPanel()
        if destructive:
            card.Background  = SolidColorBrush(Color.FromRgb(254, 242, 242))  # #FEF2F2
            card.BorderBrush = SolidColorBrush(Color.FromRgb(252, 165, 165))  # #FCA5A5
            head.Foreground  = SolidColorBrush(Color.FromRgb(185, 28, 28))    # #B91C1C
            # MDL2 glyph in its own run ("⚠" in the body font = emoji/tofu).
            self.win._add_icon_run(head, _ICON_WARNING, (185, 28, 28), size=12)
            head.Inlines.Add(Run(u"Confirm destructive action"))
            body.Text = u"{}\n`{}`".format(summary, name)
            # Fixed ink: the card stays light red in BOTH themes on purpose,
            # and token text (BotText #E4E4E4 in dark) would vanish on it.
            body.Foreground      = SolidColorBrush(Color.FromRgb(63, 63, 63))     # #3F3F3F
            status_tb.Foreground = SolidColorBrush(Color.FromRgb(140, 106, 106))  # #8C6A6A
            status_tb.Margin     = Thickness(0, 4, 0, 0)

            def _mk_btn(label, bg, fg):
                from System.Windows.Controls import Button
                b = Button()
                b.Content         = label
                b.FontSize        = 12
                b.FontWeight      = System.Windows.FontWeights.SemiBold
                b.Padding         = Thickness(14, 5, 14, 5)
                b.Margin          = Thickness(0, 0, 8, 4)
                b.Cursor          = Cursors.Hand
                b.Background      = SolidColorBrush(bg)
                b.Foreground      = SolidColorBrush(fg)
                b.BorderThickness = Thickness(0)
                return b

            # Fixed colours too: tokens read wrong on this surface in dark
            # mode (grey on red at 2.1:1, white on light grey).
            ok_btn = _mk_btn(u"Confirm",
                             Color.FromRgb(196, 43, 28),      # #C42B1C
                             Color.FromRgb(255, 255, 255))
            no_btn = _mk_btn(u"Cancel",
                             Color.FromRgb(230, 230, 230),    # #E6E6E6
                             Color.FromRgb(63, 63, 63))       # #3F3F3F
        else:
            # "Ask before edits": a question, not an alarm — accent edge.
            _bind_bg(card, 'CardBg')
            _bind_border(card, 'Accent')
            _bind_fg(head, 'Ink')
            self.win._add_icon_run(head, _ICON_EDIT, _trgb('Accent'), size=12)
            head.Inlines.Add(Run(u"Approve this model change"))
            body.Inlines.Add(Run(summary + u"\n"))
            _tool = Run(name)
            _tool.FontFamily = FontFamily(u"Consolas")
            body.Inlines.Add(_tool)
            _bind_fg(body, 'Ink')
            note = TextBlock()
            note.Text = (u"Ask before edits is on — nothing changes until you "
                         u"confirm.")
            note.FontSize     = 11
            note.TextWrapping = TextWrapping.Wrap
            note.Margin       = Thickness(0, 0, 0, 8)
            _bind_fg(note, 'Muted')
            _bind_fg(status_tb, 'Muted')
            ok_btn = self._make_card_button(
                u"Confirm", primary=True,
                tooltip=u"Run this one call now")
            no_btn = self._make_card_button(
                u"Cancel", tooltip=u"Skip this call — the assistant is told "
                                   u"you declined")

        body.ToolTip = args_tip
        panel.Children.Add(head)
        panel.Children.Add(body)
        if not destructive:
            panel.Children.Add(note)

        def _seal(msg):
            try:
                btn_row.Visibility   = Visibility.Collapsed
                ok_btn.IsEnabled     = False
                no_btn.IsEnabled     = False
                status_tb.Text       = msg
                status_tb.Visibility = Visibility.Visible
            except Exception:
                pass

        def _on_ok(s, e):
            if pending.decide(True):
                _seal(u"Confirmed")

        def _on_cancel(s, e):
            if pending.decide(False):
                _seal(u"Cancelled — skipped")

        ok_btn.Click += _on_ok
        no_btn.Click += _on_cancel

        btn_row.Children.Add(ok_btn)
        btn_row.Children.Add(no_btn)
        panel.Children.Add(btn_row)
        panel.Children.Add(status_tb)

        card.Child = panel
        self.win.chat_history_panel.Children.Add(card)
        self.win._scroll_to_bottom()
        return _seal
