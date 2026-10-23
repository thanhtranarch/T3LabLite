# -*- coding: utf-8 -*-
"""
Assistant plans — the graph path for multi-goal turns
(Intelligence.graph): orchestrator + kill switches, the plan approval card
(Run / Edit / Cancel), the step progress card, and parallel read-only task
cards (Intelligence.agents.task_manager).

Methods of T3LabAssistantWindow, moved here by composition (a CLR-derived
window cannot take a Python mixin): the window builds one
`self._plans = AssistantPlans(self)` and keeps every attribute;
`self.win` below IS that window. XAML handlers keep a one-line stub on
the window that delegates here.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Plans"

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
import System.Windows
from System.Windows import Visibility

from GUI.assistant._log import logger
from GUI.assistant.constants import (_ICON_AMBER, _ICON_EDIT, _ICON_LIST,
                                     _ICON_SLATE, _ICON_SYNC, _ICON_WARNING,
                                     _PLAN_APPROVAL_TIMEOUT, _PLAN_CANCEL,
                                     _PLAN_EDIT, _PLAN_RUN, _PLAN_SUPERSEDED,
                                     _STEP_LOOK)
from GUI.assistant.deps import (AgentDispatcher, GraphOrchestrator, HAS_AGENTS,
                                HAS_GRAPH, HAS_SPECIALISTS,
                                build_specialist_prompt, get_spec)
from GUI.assistant.lang import _is_viet_text, _reply_language
from GUI.assistant.text import _exc_text
from GUI.assistant.theme import _bind_bg, _bind_border, _bind_fg, _trgb


class AssistantPlans(object):
    """Graph path: multi-goal planning, plan approval, progress and task cards."""

    def __init__(self, window):
        self.win = window   # the T3LabAssistantWindow that owns this helper

    # ─── Graph path (multi-goal turns) ───

    def _graph_enabled(self):
        """Kill switch + prerequisites for the graph agent layer."""
        if not (HAS_GRAPH and HAS_AGENTS and HAS_SPECIALISTS):
            return False
        try:
            from config.settings import get_settings
            settings = get_settings()
            return bool(settings.is_graph_mode_enabled()
                        and settings.is_multi_agent_enabled())
        except Exception:
            return True

    def _build_orchestrator(self, provider, history, viet, progress=None):
        """A GraphOrchestrator whose runner is one `_run_native_agent` turn.

        max_parallel = 1: turns stream into one transcript (parallel is for
        headless callers, Intelligence/graph/executor.py).
        progress: optional {"card": handle}; the caller fills "card" after Run
        of an APPROVED plan — until then step marks are no-ops.
        """
        base_history = list(history)

        def _runner(node, context, state):
            if self.win._cancelled():
                # Cancel the graph state: stops between layers instead of
                # failing every remaining node.
                state.cancel()
                return None
            step = self._plan_step_of(node)
            self._mark_plan_step(progress, step, 'running')
            spec = get_spec(node.specialist) if node.specialist else None
            skill_ids = [node.skill] if node.skill else None
            # `context` = earlier nodes' output, passed like RAG excerpts.
            try:
                self.win._run_native_agent(
                    provider, list(base_history), node.goal,
                    spec=spec, skill_ids=skill_ids,
                    rag_context=(context or None), hold_busy=True)
            except Exception:
                self._mark_plan_step(progress, step, 'failed')
                raise
            text = getattr(self.win, '_last_agent_text', u'') or None
            # Provisional — _finalize_plan_progress settles it from the graph
            # state (verifier / fallback route may change it).
            self._mark_plan_step(progress, step, 'done' if text else 'failed')
            return text

        def _on_step(index, total, node):
            header = (u"**Bước {}/{}** — {}" if viet
                      else u"**Step {}/{}** — {}").format(
                          index, total, node.goal)

            def _show():
                self.win._append_bot_message(header, icon=_ICON_SYNC,
                                         icon_color=_ICON_SLATE)
                self.win._show_typing_indicator()
            self.win._ui_invoke(_show)

        skills_engine = None
        try:
            from Intelligence.skills_engine import SkillsEngine
            skills_engine = SkillsEngine()
        except Exception:
            pass

        return GraphOrchestrator(
            _runner, dispatcher=AgentDispatcher(), skills_engine=skills_engine,
            provider=provider, max_parallel=1, on_step=_on_step)

    # ─── Parallel task cards ───

    def _parallel_tasks_enabled(self):
        """Kill switch for concurrent task cards (agents.parallel_tasks).

        Default ON; eligible_for_parallel limits it to writer-FREE plans with
        >=2 read goals. Off when the graph layer is off.
        """
        if not self._graph_enabled():
            return False
        try:
            from config.settings import get_settings
            return bool(get_settings().get_agent_option("parallel_tasks", True))
        except Exception:
            return True

    def _run_agent_text(self, provider, base_history, node, cancel_check=None):
        """Run ONE read goal with NO streaming; return its text. Never raises.

        Parallel counterpart of _build_orchestrator's runner: never touches
        the streaming bubble or self._last_agent_text (concurrent tasks would
        clobber them). Read-only (writer-free plan + read-tool catalog), so no
        action group or approval card; no launcher or memory tools. Task
        worker thread.
        """
        try:
            from Intelligence.agent_loop import (AgentLoop,
                                                 build_agent_system_prompt,
                                                 read_only_tools)
            from Intelligence import tool_schema
            from core.server import get_t3labai_server
            srv = get_t3labai_server()
        except Exception as ex:
            logger.debug(u"parallel node setup failed: {}".format(_exc_text(ex)))
            return None

        spec = get_spec(node.specialist) if (node.specialist
                                             and HAS_SPECIALISTS) else None
        _is_local = provider.NAME in ("ollama", "lmstudio")

        _spec_tools = spec.tools_for(_is_local) if spec is not None else None
        if _spec_tools:
            tools = tool_schema.get_tools_by_names(provider.NAME, _spec_tools, [])
        else:
            tools = tool_schema.get_tools_for_provider(
                provider.NAME, [], essential_only=_is_local)
        # 'general' gets the FULL catalog; with no approval card, hide writers.
        tools = read_only_tools(tools)
        if not tools:
            return None

        _lang = _reply_language()
        if _lang == 'auto':
            _lang = 'vi' if _is_viet_text(node.goal) else 'en'
        _proj_instructions, _mem_block = self.win._projects._project_prompt_blocks(
            memory_tool=False)
        if spec is not None and HAS_SPECIALISTS:
            # Grounding appended below: the builder would put its own
            # "## Project instructions" header on the personal block.
            system_prompt = build_specialist_prompt(
                spec, project_instructions=u"",
                skills_block=u"", local=_is_local, lang=_lang)
        else:
            system_prompt = build_agent_system_prompt(local=_is_local, lang=_lang)
        if _proj_instructions:
            system_prompt += u"\n\n" + _proj_instructions
        if _mem_block:
            system_prompt += u"\n\n" + _mem_block

        def _exec(name, args):
            # A text-rescued call can still name any tool.
            if (tool_schema.is_model_modifying(name)
                    or srv.is_destructive(name, args or {})):
                return {"error": u"`{}` needs the user's approval; this "
                                 u"background step is read-only. Report what "
                                 u"you found and say the action needs a "
                                 u"separate request.".format(name)}
            # The ExternalEvent queue serializes main-thread work.
            return srv._execute_tool(name, dict(args or {}))

        # Window-wide cancel + this task's Cancel button.
        def _guard():
            try:
                if self.win._cancelled():
                    return True
                return bool(cancel_check and cancel_check())
            except Exception:
                return False

        loop = AgentLoop(provider, _exec, tools,
                         callbacks={"guard_check": _guard}, max_tokens=1500)
        try:
            res = loop.run(list(base_history), system_prompt, node.goal)
        except Exception as ex:
            logger.debug(u"parallel node run failed: {}".format(_exc_text(ex)))
            return None
        return ((res or {}).get("text") or u"").strip() or None

    def _run_parallel_tasks(self, plan, provider, history, viet):
        """Run an all-read multi-goal plan concurrently as task cards.

        Each answer is posted ATOMICALLY when its task finishes (never
        streamed), so texts can't interleave. WORKER THREAD (_ui_invoke).
        """
        from Intelligence.agents.task_manager import get_task_manager, DONE
        nodes = plan.agent_nodes()
        total = len(nodes)
        mgr = get_task_manager()
        base_history = list(history)

        state = {"cards": {}, "tasks": {}, "done": 0}

        # Panel FIRST so a fast task finds its card (_ui_invoke is synchronous).
        def _panel():
            self.win._hide_typing_indicator()
            state["cards"] = self._append_task_cards(nodes, state, mgr, viet)
        self.win._ui_invoke(_panel)

        def _on_done(task, _node):
            def _ui():
                self._update_task_card(state["cards"].get(_node.id), task, viet)
                if task.status == DONE and task.result:
                    self.win._append_bot_message(u"**{}**\n\n{}".format(
                        _node.goal, task.result))
                    self.win._add_to_history("assistant", task.result)
                state["done"] += 1
                if state["done"] >= total:
                    done_ok = sum(1 for t in state["tasks"].values()
                                  if t.status == DONE)
                    summary = (u"✔ Xong {}/{} tác vụ song song.".format(
                        done_ok, total) if viet else
                        u"✔ Finished {}/{} parallel tasks.".format(
                            done_ok, total))
                    self.win._append_bot_message(summary, icon=_ICON_SYNC,
                                             icon_color=_ICON_SLATE)
                    self.win._set_busy(False)
            self.win._ui_invoke(_ui)

        for node in nodes:
            def _fn(task, _n=node):
                return self._run_agent_text(provider, base_history, _n,
                                            cancel_check=task.is_cancelled)

            def _done(task, _n=node):
                _on_done(task, _n)
            t = mgr.submit(node.goal, _fn, writer=False, on_done=_done)
            state["tasks"][node.id] = t

    def _append_task_cards(self, nodes, state, mgr, viet):
        """Task-card panel; returns {node_id: handle}. UI thread only."""
        cards = {}
        try:
            from System.Windows.Controls import (Border, TextBlock, StackPanel,
                                                  Orientation, Button)
            from System.Windows import Thickness, CornerRadius, TextWrapping
            from System.Windows.Input import Cursors
            from System.Windows.Media import FontFamily

            outer = Border()
            _bind_bg(outer, 'SelectedBg')
            _bind_border(outer, 'CardBorder')
            outer.BorderThickness = Thickness(1)
            outer.CornerRadius    = CornerRadius(8)
            outer.Padding         = Thickness(12, 8, 12, 8)
            outer.Margin          = Thickness(0, 0, 8, 10)
            col = StackPanel()

            head = TextBlock()
            head.Text       = (u"Chạy {} tác vụ song song".format(len(nodes))
                               if viet else
                               u"Running {} tasks in parallel".format(len(nodes)))
            head.FontSize   = 12
            head.FontWeight = System.Windows.FontWeights.SemiBold
            _bind_fg(head, 'Ink')
            head.Margin     = Thickness(0, 0, 0, 6)
            col.Children.Add(head)

            for node in nodes:
                row = StackPanel()
                row.Orientation = Orientation.Horizontal
                row.Margin      = Thickness(0, 2, 0, 2)

                status = TextBlock()
                status.Text       = u""   # MDL2 Sync — running
                status.FontFamily = FontFamily(u"Segoe MDL2 Assets")
                status.FontSize   = 12
                _bind_fg(status, 'Blue')
                status.Margin     = Thickness(0, 1, 8, 0)

                goal = TextBlock()
                goal.Text         = node.goal
                goal.FontSize     = 12
                _bind_fg(goal, 'Ink')
                goal.TextWrapping = TextWrapping.Wrap
                goal.MaxWidth     = 320

                cancel = Button()
                cancel.Content   = (u"Hủy" if viet else u"Cancel")
                cancel.FontSize  = 10
                cancel.Margin    = Thickness(8, 0, 0, 0)
                cancel.Padding   = Thickness(6, 1, 6, 1)
                cancel.Cursor    = Cursors.Hand

                def _cancel_click(sender, args, _nid=node.id):
                    t = state["tasks"].get(_nid)
                    if t is not None:
                        mgr.cancel(t.id)
                cancel.Click += _cancel_click

                row.Children.Add(status)
                row.Children.Add(goal)
                row.Children.Add(cancel)
                col.Children.Add(row)
                cards[node.id] = {"status": status, "cancel": cancel}

            outer.Child = col
            self.win.chat_history_panel.Children.Add(outer)
            self.win._scroll_to_bottom()
        except Exception as ex:
            logger.debug(u"_append_task_cards error: {}".format(_exc_text(ex)))
        return cards

    def _update_task_card(self, handle, task, viet):
        """Flip a task row to its terminal state. UI thread only."""
        if not handle:
            return
        try:
            from Intelligence.agents.task_manager import DONE, FAILED, CANCELLED
            status = handle["status"]
            if task.status == DONE:
                status.Text = u""   # MDL2 CheckMark
                _bind_fg(status, 'Success')
            elif task.status == CANCELLED:
                status.Text = u""   # MDL2 Cancel
                _bind_fg(status, 'Muted')
            else:  # FAILED or anything non-terminal-looking
                status.Text = u""   # MDL2 Cancel
                _bind_fg(status, 'Danger')
            try:
                handle["cancel"].IsEnabled = False
                handle["cancel"].Visibility = Visibility.Collapsed
            except Exception:
                pass
        except Exception as ex:
            logger.debug(u"_update_task_card error: {}".format(_exc_text(ex)))

    def _run_graph_plan(self, raw, history, utterance=None):
        """Plan and run a MULTI-GOAL turn as a graph. WORKER THREAD.

        True = handled. False = not multi-goal / layer off; nothing was shown
        and the caller falls through to the single-specialist path.
        A plan with a model-changing step waits on a plan card for Run (Edit →
        composer, Cancel drops, Stop / new message end the wait), then runs
        with a progress card. Read-only plans run immediately.
        """
        if not self._graph_enabled():
            return False
        try:
            from Intelligence.llm_router import LLMRouter
            provider = LLMRouter().get_active_provider()
        except Exception:
            return False
        if provider is None or not getattr(provider, 'SUPPORTS_NATIVE_TOOLS',
                                           False):
            # Every node is an agent turn — needs native tools.
            return False

        viet = _is_viet_text(raw)
        rid = getattr(self.win, '_request_id', 0)
        progress = {"card": None}

        # Decide BEFORE anything is shown, so the caller can fall through.
        try:
            orchestrator = self._build_orchestrator(provider, history, viet,
                                                    progress=progress)
            plan = orchestrator.plan_for(raw, utterance=utterance)
        except Exception as ex:
            logger.debug("graph plan error: {}".format(ex))
            return False
        if plan is None or not plan.is_multi:
            return False
        logger.debug("graph plan:\n{}".format(plan.describe()))

        approved = False
        if self._plan_needs_approval(plan):
            decision = self._await_plan_approval(plan, viet)
            if decision != _PLAN_RUN:
                self._end_unapproved_plan(plan, decision, rid)
                return True
            approved = True
        else:
            # All-READ plan → concurrent cancellable cards
            # (task_manager.eligible_for_parallel); else sequential below.
            try:
                from Intelligence.agents.task_manager import eligible_for_parallel
                if eligible_for_parallel(plan.is_multi, len(plan.agent_nodes()),
                                         plan.has_writer(),
                                         self._parallel_tasks_enabled()):
                    self._run_parallel_tasks(plan, provider, history, viet)
                    return True
            except Exception as ex:
                logger.debug(u"parallel tasks skipped: {}".format(_exc_text(ex)))

        if approved:
            # Edits approved on the card: "Ask before edits" skips per-call
            # asks (destructive still confirm). Reset in finally and by the
            # next _set_busy(True).
            self.win._plan_preapproved = True

            def _show_progress():
                self._show_loading_bar(True)
                progress["card"] = self._append_plan_progress_card(plan, viet)
            self.win._ui_invoke(_show_progress)

        # From here the plan owns the turn (falling through would answer twice).
        try:
            result = orchestrator.handle(
                raw, utterance=utterance, plan=plan,
                lang=('vi' if viet else 'en'))
        except Exception as ex:
            self._finalize_plan_progress(progress, None,
                                         cancelled=self.win._cancelled())
            self.win._report_error(u"graph plan", ex, rid)
            return True
        finally:
            self.win._plan_preapproved = False

        self._finalize_plan_progress(progress, getattr(result, 'state', None),
                                     cancelled=self.win._cancelled())

        if self.win._cancelled():
            self.win._finish_cancelled(rid)
            return True

        def _finish():
            self.win._hide_typing_indicator()
            summary = self._graph_summary(result, viet)
            if summary:
                self.win._append_bot_message(summary, icon=_ICON_SYNC,
                                         icon_color=_ICON_SLATE)
                self.win._add_to_history("assistant", summary)
            self.win._set_busy(False)

        if self.win._claim_turn(rid):
            self.win._ui_invoke(_finish)
        return True

    # ─── Plan approval card ───

    @staticmethod
    def _plan_needs_approval(plan):
        """True when `plan` must be approved before any step runs. Never
        raises; an unreadable plan errs on the side of asking."""
        try:
            check = getattr(plan, 'needs_approval', None)
            if check is not None:
                return bool(check())
            return bool(plan.is_multi and plan.has_writer())
        except Exception:
            return True

    def _await_plan_approval(self, plan, viet):
        """WORKER THREAD: show the plan card and block until it is answered.

        Returns _PLAN_RUN / _PLAN_EDIT / _PLAN_CANCEL / _PLAN_SUPERSEDED, or
        PendingDecision.STOPPED / EXPIRED / 'error' (card not shown — nothing
        may run unapproved). The turn stays busy: send is Stop, and a new
        message supersedes the plan (_process_input).
        """
        from Intelligence.agent_loop import PendingDecision
        pending = PendingDecision()
        handle = {"seal": None}
        self.win._pending_plan = pending

        def _ui():
            try:
                self.win._hide_typing_indicator()
                # Waiting on the USER: no busy bar (Run brings it back).
                self._show_loading_bar(False)
                handle["seal"] = self._append_plan_card(plan, pending, viet)
            except Exception as ex:
                logger.debug(u"plan card error: {}".format(_exc_text(ex)))

        try:
            self.win._ui_invoke(_ui)
            if handle["seal"] is None:
                pending.decide('error')
            decision = pending.wait(_PLAN_APPROVAL_TIMEOUT,
                                    should_abort=self.win._cancelled)
        finally:
            self.win._pending_plan = None

        _seal_text = {
            PendingDecision.STOPPED: u"Stopped — nothing was run",
            PendingDecision.EXPIRED: u"Expired — nothing was run",
            _PLAN_SUPERSEDED: u"Set aside for your new message — nothing "
                              u"was run",
        }.get(decision)
        if _seal_text and handle["seal"] is not None:
            def _seal():
                try:
                    handle["seal"](_seal_text)
                except Exception:
                    pass
            self.win._ui_invoke(_seal)
        return decision

    def _end_unapproved_plan(self, plan, decision, rid):
        """Close a turn whose plan was not run. WORKER THREAD.

        Adds the history record (so "skip step 2" can refer to the plan) and
        a bubble only when expired / not shown.
        """
        from Intelligence.agent_loop import PendingDecision
        if decision == PendingDecision.STOPPED:
            self.win._finish_cancelled(rid, note=u"The plan was not run.")
            return
        reason = {
            _PLAN_EDIT:       u"the user chose to edit it",
            _PLAN_CANCEL:     u"the user cancelled it",
            _PLAN_SUPERSEDED: u"the user sent a new message instead",
            PendingDecision.EXPIRED: u"it got no answer",
        }.get(decision, u"it could not be shown for approval")
        notice = None
        if decision == PendingDecision.EXPIRED:
            notice = (u"The plan waited {} minutes without an answer, so "
                      u"nothing was run. Send the request again to see it "
                      u"again.".format(_PLAN_APPROVAL_TIMEOUT // 60))
        elif decision not in (_PLAN_EDIT, _PLAN_CANCEL, _PLAN_SUPERSEDED):
            notice = (u"The plan could not be shown for approval, so nothing "
                      u"was run. Please send the request again.")
        try:
            steps = len(plan.agent_nodes())
            plan_text = plan.review_text()
        except Exception:
            steps, plan_text = 0, u""
        record = u"Proposed a {}-step plan; it was not run ({}).".format(
            steps, reason)
        if plan_text:
            record += u"\n" + plan_text

        def _ui():
            try:
                self.win._hide_typing_indicator()
                if notice:
                    self.win._append_bot_message(notice, icon=_ICON_WARNING,
                                             icon_color=_ICON_AMBER)
                self.win._add_to_history("assistant", record)
            except Exception:
                pass
            finally:
                self.win._set_busy(False)

        if self.win._claim_turn(rid):
            self.win._ui_invoke(_ui)

    def _show_loading_bar(self, visible):
        """The thin busy bar under the top bar. UI THREAD. Never raises."""
        try:
            self.win.top_loading_bar.Visibility = (Visibility.Visible if visible
                                               else Visibility.Collapsed)
        except Exception:
            pass

    def _load_plan_into_composer(self, text):
        """Edit: put the numbered plan in the message box. UI THREAD."""
        try:
            self.win.chat_input.Text = text
            self.win.chat_input.Focus()
            self.win.chat_input.CaretIndex = len(text or u"")
        except Exception as ex:
            logger.debug(u"plan edit error: {}".format(_exc_text(ex)))

    def _append_plan_card(self, plan, pending, viet):
        """Plan card: numbered steps, model-changing marks, Run / Edit /
        Cancel. UI THREAD. Returns seal(text). Star grids + WrapPanel reflow
        down to the 240 DIP pane.
        """
        from System.Windows.Controls import (Border, TextBlock, StackPanel,
                                              Grid, ColumnDefinition,
                                              WrapPanel)
        from System.Windows.Documents import Run
        from System.Windows import (Thickness, CornerRadius, TextWrapping,
                                    GridLength, GridUnitType)

        steps = plan.review_steps()
        n_mod = sum(1 for st in steps if st['modifies'])
        plan_text = plan.review_text()

        card = Border()
        _bind_bg(card, 'CardBg')
        _bind_border(card, 'Accent')
        card.BorderThickness = Thickness(1)
        card.CornerRadius    = CornerRadius(8)
        card.Padding         = Thickness(12, 10, 12, 10)
        card.Margin          = Thickness(0, 0, 8, 10)
        panel = StackPanel()

        head = TextBlock()
        head.FontSize     = 12
        head.FontWeight   = System.Windows.FontWeights.SemiBold
        head.TextWrapping = TextWrapping.Wrap
        _bind_fg(head, 'Ink')
        self.win._add_icon_run(head, _ICON_LIST, _trgb('Accent'), size=12)
        head.Inlines.Add(Run(u"Review the plan before it runs"))
        panel.Children.Add(head)

        sub = TextBlock()
        sub.Text = (u"{} steps · {} change{} the model. Nothing runs until "
                    u"you click Run.".format(
                        len(steps), n_mod, u"s" if n_mod == 1 else u""))
        sub.FontSize     = 11.5
        sub.TextWrapping = TextWrapping.Wrap
        sub.Margin       = Thickness(0, 4, 0, 4)
        _bind_fg(sub, 'Muted')
        panel.Children.Add(sub)

        for st in steps:
            row = Grid()
            row.Margin = Thickness(0, 4, 0, 0)
            c0 = ColumnDefinition()
            c0.Width = GridLength.Auto
            c1 = ColumnDefinition()
            c1.Width = GridLength(1, GridUnitType.Star)
            row.ColumnDefinitions.Add(c0)
            row.ColumnDefinitions.Add(c1)

            num = TextBlock()
            num.Text     = u"{}.".format(st['index'])
            num.FontSize = 12
            num.Margin   = Thickness(0, 0, 8, 0)
            _bind_fg(num, 'Muted')
            Grid.SetColumn(num, 0)
            row.Children.Add(num)

            col = StackPanel()
            goal = TextBlock()
            goal.Text         = st['goal']
            goal.FontSize     = 12
            goal.TextWrapping = TextWrapping.Wrap
            _bind_fg(goal, 'Ink')
            col.Children.Add(goal)

            tag = TextBlock()
            tag.FontSize     = 11
            tag.TextWrapping = TextWrapping.Wrap
            if st['modifies']:
                # Glyph AND words, never colour alone.
                self.win._add_icon_run(tag, _ICON_EDIT, _trgb('Warning'), size=11)
                tag.Inlines.Add(Run(u"Changes the model · {}".format(
                    st['role'])))
                _bind_fg(tag, 'Warning')
            else:
                tag.Inlines.Add(Run(u"Read-only · {}".format(st['role'])))
                _bind_fg(tag, 'Muted')
            col.Children.Add(tag)
            Grid.SetColumn(col, 1)
            row.Children.Add(col)
            panel.Children.Add(row)

        btn_row = WrapPanel()
        btn_row.Margin = Thickness(0, 12, 0, 0)
        if self.win._action_mode() == 'confirm':
            _run_tip = (u"Run all {} steps in order. This click approves the "
                        u"plan's edits; destructive actions still ask one by "
                        u"one.".format(len(steps)))
        else:
            _run_tip = (u"Run all {} steps in order. Destructive actions "
                        u"still ask before they run.".format(len(steps)))
        run_btn = self.win._tool_cards._make_card_button(u"Run", primary=True, tooltip=_run_tip)
        edit_btn = self.win._tool_cards._make_card_button(
            u"Edit", tooltip=u"Put the steps in the message box so you can "
                             u"change them, then send")
        cancel_btn = self.win._tool_cards._make_card_button(
            u"Cancel", tooltip=u"Drop this plan — nothing runs")
        btn_row.Children.Add(run_btn)
        btn_row.Children.Add(edit_btn)
        btn_row.Children.Add(cancel_btn)
        panel.Children.Add(btn_row)

        status_tb = TextBlock()
        status_tb.FontSize     = 11.5
        status_tb.TextWrapping = TextWrapping.Wrap
        status_tb.Margin       = Thickness(0, 8, 0, 0)
        status_tb.Visibility   = Visibility.Collapsed
        _bind_fg(status_tb, 'Muted')
        panel.Children.Add(status_tb)

        def _seal(msg):
            try:
                btn_row.Visibility   = Visibility.Collapsed
                for b in (run_btn, edit_btn, cancel_btn):
                    b.IsEnabled = False
                status_tb.Text       = msg
                status_tb.Visibility = Visibility.Visible
            except Exception:
                pass

        def _on_run(s, e):
            if pending.decide(_PLAN_RUN):
                _seal(u"Approved — running {} steps".format(len(steps)))

        def _on_edit(s, e):
            if pending.decide(_PLAN_EDIT):
                self._load_plan_into_composer(plan_text)
                _seal(u"Moved to the message box — change the steps and "
                      u"send")

        def _on_cancel(s, e):
            if pending.decide(_PLAN_CANCEL):
                _seal(u"Cancelled — nothing was run")

        run_btn.Click += _on_run
        edit_btn.Click += _on_edit
        cancel_btn.Click += _on_cancel

        card.Child = panel
        self.win.chat_history_panel.Children.Add(card)
        self.win._scroll_to_bottom()
        return _seal

    # ─── Task progress card (approved plans) ───

    @staticmethod
    def _plan_step_of(node):
        """The plan step a graph node reports on: a fallback copy (n1_alt)
        ticks its original step, not a row of its own."""
        try:
            return (node.meta or {}).get('fallback_for') or node.id
        except Exception:
            return getattr(node, 'id', None)

    def _mark_plan_step(self, progress, step_id, state):
        """Any thread: tick one step on the progress card (no-op without one)."""
        handle = (progress or {}).get("card")
        if not handle or step_id is None:
            return
        self.win._ui_invoke(lambda: self._set_plan_step(handle, step_id, state))

    def _append_plan_progress_card(self, plan, viet):
        """Progress card: each step's glyph AND word (pending → running →
        done/failed), plus Stop. UI THREAD. Returns the handle or None."""
        try:
            from System.Windows.Controls import (Border, TextBlock, StackPanel,
                                                  Grid, ColumnDefinition)
            from System.Windows import (Thickness, CornerRadius, TextWrapping,
                                        GridLength, GridUnitType,
                                        VerticalAlignment)
            from System.Windows.Media import FontFamily

            steps = plan.review_steps()
            outer = Border()
            _bind_bg(outer, 'SelectedBg')
            _bind_border(outer, 'CardBorder')
            outer.BorderThickness = Thickness(1)
            outer.CornerRadius    = CornerRadius(8)
            outer.Padding         = Thickness(12, 8, 12, 8)
            outer.Margin          = Thickness(0, 0, 8, 10)
            col = StackPanel()

            top = Grid()
            t0 = ColumnDefinition()
            t0.Width = GridLength(1, GridUnitType.Star)
            t1 = ColumnDefinition()
            t1.Width = GridLength.Auto
            top.ColumnDefinitions.Add(t0)
            top.ColumnDefinitions.Add(t1)
            head = TextBlock()
            head.FontSize          = 12
            head.FontWeight        = System.Windows.FontWeights.SemiBold
            head.TextWrapping      = TextWrapping.Wrap
            head.VerticalAlignment = VerticalAlignment.Center
            _bind_fg(head, 'Ink')
            Grid.SetColumn(head, 0)
            top.Children.Add(head)
            stop_btn = self.win._tool_cards._make_card_button(
                u"Stop", tooltip=u"Stop after the current step — steps that "
                                 u"already ran are not undone")
            stop_btn.Margin = Thickness(8, 0, 0, 0)

            def _stop(s, e):
                try:
                    self.win._request_stop()
                except Exception:
                    pass
            stop_btn.Click += _stop
            Grid.SetColumn(stop_btn, 1)
            top.Children.Add(stop_btn)
            col.Children.Add(top)

            rows = {}
            for st in steps:
                row = Grid()
                row.Margin = Thickness(0, 6, 0, 0)
                c0 = ColumnDefinition()
                c0.Width = GridLength.Auto
                c1 = ColumnDefinition()
                c1.Width = GridLength(1, GridUnitType.Star)
                row.ColumnDefinitions.Add(c0)
                row.ColumnDefinitions.Add(c1)

                glyph = TextBlock()
                glyph.FontFamily = FontFamily(u"Segoe MDL2 Assets")
                glyph.FontSize   = 12
                glyph.Margin     = Thickness(0, 2, 8, 0)
                Grid.SetColumn(glyph, 0)
                row.Children.Add(glyph)

                body = StackPanel()
                goal = TextBlock()
                goal.Text         = st['goal']
                goal.FontSize     = 12
                goal.TextWrapping = TextWrapping.Wrap
                _bind_fg(goal, 'Ink')
                body.Children.Add(goal)
                line = TextBlock()
                line.FontSize     = 11
                line.TextWrapping = TextWrapping.Wrap
                body.Children.Add(line)
                Grid.SetColumn(body, 1)
                row.Children.Add(body)
                col.Children.Add(row)
                rows[st['id']] = {"glyph": glyph, "line": line,
                                  "index": st['index'],
                                  "modifies": st['modifies']}

            outer.Child = col
            handle = {"rows": rows, "head": head, "stop": stop_btn,
                      "total": len(steps), "states": {}, "phase": u"running"}
            for sid in rows:
                self._set_plan_step(handle, sid, 'pending')
            self.win.chat_history_panel.Children.Add(outer)
            self.win._scroll_to_bottom()
            return handle
        except Exception as ex:
            logger.debug(u"_append_plan_progress_card error: {}".format(
                _exc_text(ex)))
            return None

    @staticmethod
    def _plan_progress_title(phase, done, total):
        """Header of the progress card."""
        lead = {u"running": u"Running plan", u"finished": u"Plan finished",
                u"stopped": u"Plan stopped"}.get(phase, u"Plan")
        return u"{} — {} of {} steps done".format(lead, done, total)

    def _set_plan_step(self, handle, step_id, state):
        """Paint one step row. UI THREAD."""
        try:
            row = handle["rows"].get(step_id)
            if row is None:
                return
            glyph, word, token = _STEP_LOOK.get(state, _STEP_LOOK['pending'])
            row["glyph"].Text = glyph
            _bind_fg(row["glyph"], token)
            line = u"Step {} · {}".format(row["index"], word)
            if row["modifies"]:
                line += u" · changes the model"
            row["line"].Text = line
            if state == 'failed':
                _bind_fg(row["line"], 'Danger')
            else:
                _bind_fg(row["line"], 'Muted')
            handle["states"][step_id] = state
            done = sum(1 for v in handle["states"].values() if v == 'done')
            handle["head"].Text = self._plan_progress_title(
                handle.get("phase"), done, handle["total"])
        except Exception as ex:
            logger.debug(u"_set_plan_step error: {}".format(_exc_text(ex)))

    @staticmethod
    def _final_step_state(graph_status, current, cancelled):
        """Where a step ends up once the plan is over.

        graph_status (None if the run crashed) is the authority over
        `current` (what the card shows): verifier / fallback can change it.
        """
        if graph_status in ('done', 'failed', 'skipped', 'cancelled'):
            return graph_status
        if graph_status is None and current in ('done', 'failed'):
            return current
        if cancelled:
            return 'cancelled'
        return 'failed' if current == 'running' else 'skipped'

    def _finalize_plan_progress(self, progress, state, cancelled=False):
        """Any thread: settle every row from the graph state, retire Stop."""
        handle = (progress or {}).get("card")
        if not handle:
            return

        def _ui():
            try:
                handle["phase"] = u"stopped" if cancelled else u"finished"
                for sid in list(handle["rows"].keys()):
                    status = None
                    if state is not None:
                        try:
                            status = state.status_of(sid)
                        except Exception:
                            status = None
                    self._set_plan_step(handle, sid, self._final_step_state(
                        status, handle["states"].get(sid), cancelled))
                handle["stop"].Visibility = Visibility.Collapsed
                handle["stop"].IsEnabled = False
            except Exception as ex:
                logger.debug(u"_finalize_plan_progress error: {}".format(
                    _exc_text(ex)))
        self.win._ui_invoke(_ui)

    @staticmethod
    def _graph_summary(result, viet):
        """Closing line for a graph turn: a tally plus the reducer's report
        for failed steps only (each node already streamed its answer)."""
        plan, state = result.plan, result.state
        if plan is None or state is None:
            return u''
        nodes = plan.agent_nodes()
        failed = [n for n in nodes if not state.succeeded(n.id)]
        total = len(nodes)
        if not failed:
            return (u"✔ Đã hoàn thành cả {} bước.".format(total) if viet
                    else u"✔ All {} steps completed.".format(total))
        try:
            from Intelligence.graph.reducer import Reducer
            detail = Reducer().reduce(plan.graph, state,
                                      node_ids=[n.id for n in failed],
                                      lang=('vi' if viet else 'en'))
        except Exception:
            detail = u''
        head = (u"Hoàn thành {}/{} bước.".format(total - len(failed), total)
                if viet else
                u"Completed {}/{} steps.".format(total - len(failed), total))
        return u"{}\n\n{}".format(head, detail) if detail else head
