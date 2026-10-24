# -*- coding: utf-8 -*-
"""
Assistant spellcheck — the deterministic /english-spellcheck scan (reads
EVERY text string, unlike the paged agent tools) and its findings card.

Methods of T3LabAssistantWindow, moved here by composition (a CLR-derived
window cannot take a Python mixin): the window builds one
`self._spell = AssistantSpellcheck(self)` and keeps every attribute;
`self.win` below IS that window. XAML handlers keep a one-line stub on
the window that delegates here.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Spellcheck"

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
import System.Windows
from System import Action
from System.Threading import ApartmentState, Thread, ThreadStart

from GUI.assistant._log import logger
from GUI.assistant.constants import _ICON_SEARCH
from GUI.assistant.deps import get_active_provider_name, has_local_llm
from GUI.assistant.lang import _is_viet_text
from GUI.assistant.text import _exc_text
from GUI.assistant.theme import _bind_bg, _bind_border, _bind_fg
try:
    from GUI.assistant.deps import has_api_key
except ImportError:
    pass    # defined only when its optional import worked


class AssistantSpellcheck(object):
    """The /english-spellcheck scan and its findings card."""

    def __init__(self, window):
        self.win = window   # the T3LabAssistantWindow that owns this helper

    def _run_spellcheck(self, raw):
        """Proofread all model text, post one report with element IDs.

        Background thread; active-view scope when the request says so.
        """
        viet = _is_viet_text(raw)
        low = (raw or u"").lower()
        view_only = any(k in low for k in (
            u"view này", u"view nay", u"view hiện tại", u"view hien tai",
            u"trong view", u"active view", u"current view", u"in view",
            u"in this view"))

        def work():
            report = None
            ui_findings, ui_total, ui_uniq, ui_failed = [], 0, 0, 0
            try:
                from Services import spell_checker as SC
                # Collect on the Revit MAIN thread via the MCP ExternalEvent —
                # from this worker Revit rejects it ("No Text Notes found" on
                # full models). collect_spellcheck_text gathers all
                # human-authored text (notes, names, title blocks, Project
                # Information, revisions, dimension overrides, schedules...).
                notes = None
                res = None
                try:
                    from core.server import get_t3labai_server
                    srv = get_t3labai_server()
                    res = srv._execute_tool('collect_spellcheck_text',
                                            {'view_only': view_only})
                    if not isinstance(res, dict):
                        res = {}
                    if 'items' in res:
                        notes = []
                        for it in res['items']:
                            _txt = (it.get('text') or u'').strip()
                            if _txt and any(c.isalpha() for c in _txt):
                                notes.append({
                                    'id':   it.get('id'),
                                    'text': _txt,
                                    'view': it.get('view')
                                            or it.get('source') or u'',
                                })
                except Exception as _col_ex:
                    logger.debug(u"spellcheck server collect error: {}".format(_exc_text(_col_ex)))
                    notes = None
                if notes is None and isinstance(res, dict) and res.get('error'):
                    # The scan itself broke — say so, not "No Text Notes found".
                    report = (u"Không quét được model — MCP server báo lỗi: {}"
                              if viet else
                              u"Could not scan the model — MCP server error: {}"
                              ).format(res.get('error'))
                elif notes is None:
                    # Last resort: off-thread read (works only while idle).
                    notes = SC.collect_text_notes(self.win.doc, view_only=view_only)
                if report is not None:
                    pass          # scan error already reported above
                elif not notes:
                    if viet:
                        report = u"No Text Notes found{}.".format(
                            u" in the current view" if view_only else u" in the project")
                    else:
                        report = u"No Text Notes found{}.".format(
                            u" in the active view" if view_only else u" in the project")
                else:
                    uniq = SC.dedupe_notes(notes)
                    # PRIMARY: deterministic dictionary + edit distance, no AI
                    # (small local models replied NO_ERRORS on obvious typos).
                    det = SC.check_deterministic(uniq)
                    if det is not None:
                        if len(uniq) > 300:
                            self.win._safe_append_bot(
                                u"Đang phân tích {} nội dung ({} mục)…".format(
                                    len(uniq), len(notes)) if viet else
                                u"Analyzing {} unique texts ({} items)…".format(
                                    len(uniq), len(notes)))
                        findings, failed = det, 0
                        report = SC.format_report(findings, len(notes), len(uniq),
                                                  failed, viet, view_only)
                        ui_findings, ui_total, ui_uniq, ui_failed = \
                            findings, len(notes), len(uniq), failed
                    elif not (has_api_key() or has_local_llm()):
                        report = (
                            u"Đã tìm thấy **{}** mục text nhưng không tải được từ "
                            u"điển kiểm tra chính tả và chưa kết nối AI — cài lại "
                            u"extension hoặc kết nối provider trong Settings."
                            if viet else
                            u"Found **{}** text items but the spell-check "
                            u"dictionary could not be loaded and no AI provider is "
                            u"connected — reinstall the extension or connect a "
                            u"provider in Settings.").format(len(notes))
                    else:
                        # Dictionary missing → LLM batches; cloud takes larger
                        # ones than the local default (25 notes / 3.5k chars).
                        try:
                            _cloud = get_active_provider_name() in (
                                "claude", "openai")
                        except Exception:
                            _cloud = False
                        if _cloud:
                            batches = SC.build_batches(uniq, max_notes=120,
                                                       max_chars=15000)
                            _max_tok = 3000
                        else:
                            batches = SC.build_batches(uniq)
                            _max_tok = 1400
                        if len(batches) > 1:
                            self.win._safe_append_bot(
                                u"Found {} text items ({} unique texts) — checking "
                                u"in {} batches, please wait...".format(
                                    len(notes), len(uniq), len(batches)))
                        from Intelligence.llm_router import LLMRouter
                        router = LLMRouter()
                        sys_p  = SC.build_system_prompt(viet)
                        findings, failed = [], 0
                        for batch in batches:
                            if self.win._cancel_requested:
                                break
                            resp = None
                            try:
                                resp = router.chat([], sys_p,
                                                   SC.build_batch_query(batch),
                                                   max_tokens=_max_tok)
                            except Exception as ex:
                                logger.debug(u"spellcheck batch error: {}".format(_exc_text(ex)))
                            if not resp or not resp.strip():
                                failed += 1
                                continue
                            findings.extend(SC.parse_findings(resp, batch))
                        report = SC.format_report(findings, len(notes), len(uniq),
                                                  failed, viet, view_only)
                        ui_findings, ui_total, ui_uniq, ui_failed = \
                            findings, len(notes), len(uniq), failed
            except Exception as ex:
                logger.debug(u"spellcheck error: {}".format(_exc_text(ex)))
                report = (u"Spell-check failed — see the console for details."
                          if viet else u"Spell-check failed — see console for details.")
            if report:
                # Findings → interactive card (fix + select/zoom id links);
                # the text report goes to history for follow-up "fix ..."
                # turns and is the fallback if the card fails.
                def _show(_r=report, _f=ui_findings, _t=ui_total,
                          _u=ui_uniq, _fb=ui_failed):
                    shown = False
                    if _f:
                        shown = self._append_spellcheck_findings(
                            _f, viet, _t, _u)
                    if not shown:
                        self.win._append_bot_message(_r)
                    elif _fb:
                        self.win._append_bot_message(
                            u"⚠️ {} nhóm không kiểm tra được (AI không phản "
                            u"hồi) — thử lại sau.".format(_fb) if viet else
                            u"⚠️ {} batch(es) could not be checked (no AI "
                            u"response) — try again.".format(_fb))
                    self.win._add_to_history("assistant", _r)
                self.win.Dispatcher.Invoke(Action(_show))
            self.win.Dispatcher.Invoke(Action(lambda: self.win._set_busy(False)))

        t = Thread(ThreadStart(work))
        t.IsBackground = True
        t.SetApartmentState(ApartmentState.STA)
        t.Start()

    def _append_spellcheck_findings(self, findings, viet, total, uniq):
        """Spell-check result table, one row per distinct correction; element
        ids link to their view (uidoc.ShowElements). UI THREAD.

        False = not rendered; the caller falls back to the markdown report.
        """
        try:
            from System.Windows.Controls import (Border, StackPanel, TextBlock,
                                                  WrapPanel, Grid,
                                                  ColumnDefinition, RowDefinition,
                                                  Orientation)
            from System.Windows import (Thickness, CornerRadius, TextWrapping,
                                        TextDecorations, GridLength, GridUnitType,
                                        HorizontalAlignment, VerticalAlignment)
            from System.Windows.Media import SolidColorBrush, Color
            from System.Windows.Input import Cursors

            try:
                from Services import spell_checker as SC
                rows = SC.aggregate_findings(findings)
            except Exception:
                rows = []
            if not rows:
                return False

            _font = System.Windows.Media.FontFamily("Hanken Grotesk, Inter")

            # THEME TOKENS only, never hex (hardcoded Lumina broke dark theme).
            # See docs/assistant-revit-ui.md.
            card = Border()
            _bind_bg(card, 'CardBg')
            _bind_border(card, 'CardBorder')
            card.BorderThickness = Thickness(1)
            card.CornerRadius    = CornerRadius(8)
            card.Padding         = Thickness(12, 9, 12, 10)
            card.Margin          = Thickness(0, 0, 8, 12)

            panel = StackPanel()

            instances = sum(len(r["ids"]) for r in rows)
            # MDL2 glyph, not an emoji: matches the tool cards and the theme.
            head_row = StackPanel()
            head_row.Orientation = Orientation.Horizontal
            head_row.Margin      = Thickness(0, 0, 0, 8)

            hicon = TextBlock()
            hicon.Text       = _ICON_SEARCH
            hicon.FontFamily = System.Windows.Media.FontFamily(u"Segoe MDL2 Assets")
            hicon.FontSize   = 12
            hicon.Margin     = Thickness(0, 1, 7, 0)
            hicon.VerticalAlignment = VerticalAlignment.Center
            _bind_fg(hicon, 'Blue')
            head_row.Children.Add(hicon)

            hdr = TextBlock()
            hdr.Text = (u"Kiểm tra chính tả: {} lỗi ({} vị trí) trong {} nội dung"
                        .format(len(rows), instances, uniq) if viet else
                        u"Spell-check: {} issue(s) across {} instances in {} unique texts"
                        .format(len(rows), instances, uniq))
            hdr.FontSize     = 12.5
            hdr.FontFamily   = _font
            hdr.FontWeight   = System.Windows.FontWeights.SemiBold
            _bind_fg(hdr, 'Ink')
            hdr.TextWrapping = TextWrapping.Wrap
            hdr.VerticalAlignment = VerticalAlignment.Center
            head_row.Children.Add(hdr)
            panel.Children.Add(head_row)

            grid = Grid()
            for w in (GridLength.Auto, GridLength.Auto, GridLength.Auto,
                      GridLength.Auto, GridLength(1.4, GridUnitType.Star),
                      GridLength(1.0, GridUnitType.Star)):
                cd = ColumnDefinition()
                cd.Width = w
                grid.ColumnDefinitions.Add(cd)

            _MAX_ROWS = 60
            shown_rows = rows[:_MAX_ROWS]
            for _ in range(len(shown_rows) + 1):
                grid.RowDefinitions.Add(RowDefinition())

            def _cell(text, col, row, token, size=12, bold=False, wrap=True,
                      align_r=False):
                tb = TextBlock()
                tb.Text         = text
                tb.FontSize     = size
                tb.FontFamily   = _font
                _bind_fg(tb, token)
                tb.TextWrapping = TextWrapping.Wrap if wrap else TextWrapping.NoWrap
                tb.Margin       = Thickness(0, 5, 12, 5)
                tb.VerticalAlignment = VerticalAlignment.Center
                if bold:
                    tb.FontWeight = System.Windows.FontWeights.SemiBold
                if align_r:
                    tb.HorizontalAlignment = HorizontalAlignment.Right
                Grid.SetColumn(tb, col)
                Grid.SetRow(tb, row)
                grid.Children.Add(tb)
                return tb

            heads = ([u"#", u"Sai", u"Đúng", u"SL", u"Element IDs", u"View"] if viet
                     else [u"#", u"Wrong", u"Correct", u"Qty", u"Element IDs", u"View"])
            for c, h in enumerate(heads):
                _cell(h, c, 0, 'Faint', size=10.5, bold=True,
                      align_r=(c == 3)).FontWeight = System.Windows.FontWeights.Bold
            hline = Border()
            _bind_border(hline, 'CardBorder')
            hline.BorderThickness = Thickness(0, 0, 0, 1)
            Grid.SetRow(hline, 0)
            Grid.SetColumnSpan(hline, 6)
            grid.Children.Add(hline)

            for i, r in enumerate(shown_rows):
                rr = i + 1
                bg = Border()
                if i % 2 == 1:
                    _bind_bg(bg, 'CodeBg')          # subtle zebra on both themes
                _bind_border(bg, 'Divider')
                bg.BorderThickness = Thickness(0, 0, 0, 1)
                Grid.SetRow(bg, rr)
                Grid.SetColumnSpan(bg, 6)
                grid.Children.Add(bg)

                _cell(u"{}".format(rr), 0, rr, 'Muted', size=11)

                # Wrong / Correct: same plain text, colour carries the meaning
                # (Danger is error TEXT, never a fill — no red pill).
                _cell(r["wrong"], 1, rr, 'Danger', size=12, bold=True)

                right = r["right"]
                if r.get("reason") and not right:
                    right = r["reason"]
                elif r.get("reason"):
                    right = u"{} ({})".format(right, r["reason"])
                _cell(right, 2, rr, 'Success', size=12, bold=True)

                _cell(u"{}".format(len(r["ids"])), 3, rr, 'Muted', size=11,
                      align_r=True)

                ids = [x for x in r["ids"] if x is not None]
                wrap = WrapPanel()
                wrap.Orientation = Orientation.Horizontal
                wrap.Margin      = Thickness(0, 4, 12, 4)
                Grid.SetColumn(wrap, 4)
                Grid.SetRow(wrap, rr)

                def _mk_link(label, id_list, tip):
                    tb = TextBlock()
                    tb.Text            = label
                    tb.FontSize        = 11
                    tb.FontFamily      = _font
                    _bind_fg(tb, 'Blue')
                    tb.TextDecorations = TextDecorations.Underline
                    tb.Cursor          = Cursors.Hand
                    tb.Margin          = Thickness(0, 1, 10, 1)
                    tb.ToolTip         = tip

                    def _click(s, e, _ids=list(id_list)):
                        self.win._tool_cards._select_in_revit_async(_ids)

                    tb.MouseLeftButtonUp += _click
                    return tb

                _go_tip = (u"Mở view chứa element & zoom" if viet
                           else u"Open the element's view & zoom")
                for eid in ids[:8]:
                    wrap.Children.Add(_mk_link(u"#{}".format(eid), [eid], _go_tip))
                if len(ids) > 8:
                    wrap.Children.Add(_mk_link(
                        u"+{}".format(len(ids) - 8), ids,
                        (u"Chọn tất cả {}".format(len(ids)) if viet
                         else u"Select all {}".format(len(ids)))))
                elif len(ids) > 1:
                    wrap.Children.Add(_mk_link(
                        u"chọn cả {}".format(len(ids)) if viet
                        else u"all {}".format(len(ids)), ids,
                        (u"Chọn & zoom tất cả" if viet else u"Select & zoom all")))
                grid.Children.Add(wrap)

                views = [v for v in r["views"] if v]
                _cell(u" / ".join(views[:3]) + (u" …" if len(views) > 3 else u"")
                      if views else u"—", 5, rr, 'Muted', size=11)

            panel.Children.Add(grid)

            if len(rows) > _MAX_ROWS:
                more = TextBlock()
                more.Text = (u"… và {} lỗi khác — xem bảng đầy đủ trong lịch sử chat."
                             .format(len(rows) - _MAX_ROWS) if viet else
                             u"… and {} more — full table kept in the chat history."
                             .format(len(rows) - _MAX_ROWS))
                more.FontSize = 11
                more.FontFamily = _font
                _bind_fg(more, 'Muted')
                more.Margin = Thickness(0, 8, 0, 0)
                more.TextWrapping = TextWrapping.Wrap
                panel.Children.Add(more)

            hint = TextBlock()
            hint.Text = (u'Bấm ID để mở view chứa element. Nhắn "sửa 1,2..." '
                         u'hoặc "sửa tất cả" để áp dụng đề xuất.' if viet else
                         u'Click an ID to open the element\'s view. Reply '
                         u'"fix 1,2..." or "fix all" to apply the suggestions.')
            hint.FontSize     = 11
            hint.FontFamily   = _font
            _bind_fg(hint, 'Faint')
            hint.TextWrapping = TextWrapping.Wrap
            hint.Margin       = Thickness(0, 10, 0, 0)
            panel.Children.Add(hint)

            card.Child = panel
            self.win.chat_history_panel.Children.Add(card)
            self.win._scroll_to_bottom()
            return True
        except Exception as ex:
            logger.debug(u"_append_spellcheck_findings error: {}".format(_exc_text(ex)))
            return False
