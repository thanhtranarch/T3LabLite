# -*- coding: utf-8 -*-
"""
Assistant documents — what a turn knows beyond the chat: pasted links
(fetched into attachments), attachment context (RAG v2 over the knowledge
store), the knowledge agent (cited answers from the index), the PDF
markup-comment agent, and the knowledge reference block for the tool agent.

Methods of T3LabAssistantWindow, moved here by composition (a CLR-derived
window cannot take a Python mixin): the window builds one
`self._docs = AssistantDocuments(self)` and keeps every attribute;
`self.win` below IS that window. XAML handlers keep a one-line stub on
the window that delegates here.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Documents"

import os

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('System')
from System import Action
from System.Threading import ApartmentState, Thread, ThreadStart

from GUI.assistant._log import logger
from GUI.assistant.constants import _ICON_SLATE, _ICON_SYNC
from GUI.assistant.deps import (HAS_KNOWLEDGE, HAS_SPECIALISTS,
                                build_text_context, get_active_store, get_spec,
                                is_pdf, link_reader)
from GUI.assistant.lang import _is_viet_text, _ui_viet
from GUI.assistant.text import _exc_text


class AssistantDocuments(object):
    """Links, attachment context, and the knowledge / PDF-comment agents."""

    def __init__(self, window):
        self.win = window   # the T3LabAssistantWindow that owns this helper

    def _link_cache_dir(self):
        """Downloaded-link folder: project attachments, else temp (never Documents)."""
        try:
            from config.project_store import ProjectStore
            ps = ProjectStore()
            d = ps.attachments_dir(ps.get_active_project_id())
            if d:
                return d
        except Exception:
            pass
        try:
            import tempfile
            d = os.path.join(tempfile.gettempdir(), 'T3LabAssistant', 'links')
            if not os.path.isdir(d):
                os.makedirs(d)
            try:
                from core import housekeeping
                housekeeping.run_once('assistant.link_cache',
                                      housekeeping.prune_files, d,
                                      max_age_days=7)
            except Exception:
                pass
            return d
        except Exception:
            return None

    def _resolve_message_links(self, raw, attached):
        """Fetch every link in `raw`, append the readable files to `attached`
        (in place) and return the context block for the rest. WORKER THREAD.

        u"" (and no cost) when the message has no links.
        """
        try:
            links = link_reader.find_links(raw)
            if not links:
                return u""

            viet = _is_viet_text(raw)
            self.win._safe_update_typing_text(
                u"● ● ●  Đang đọc {} liên kết…".format(len(links)) if viet
                else u"● ● ●  Reading {} link(s)…".format(len(links)))

            res = link_reader.resolve(raw, self._link_cache_dir())

            added = []
            for path in res.get('files') or []:
                if path not in attached:
                    attached.append(path)
                    added.append(path)

            if added or res.get('notes'):
                self.win._log_activity(u"Links ({}): {}".format(
                    len(res.get('links') or []),
                    u", ".join(res.get('links') or [])))

            parts = []
            header = link_reader.source_header(res.get('links') or [])
            if header:
                parts.append(header)
            parts.extend(res.get('notes') or [])
            # _build_attachment_context reads the files; name the linked ones for citing
            if added:
                parts.append(u'[Đã tải về từ liên kết: {}]'.format(
                    u', '.join(os.path.basename(p) for p in added)))
            return u"\n\n".join(p for p in parts if p)
        except Exception as ex:
            logger.debug(u"_resolve_message_links error: {}".format(_exc_text(ex)))
            return u""

    def _build_attachment_context(self, raw, attached):
        """Build the attachment context string. WORKER THREAD.

        Large PDFs are indexed into the active knowledge store and answered via
        retrieval (top-5 excerpts, file + page citations). Small PDFs and images
        keep the legacy full-text path; any failure degrades to it, never errors.
        """
        pdfs = [p for p in attached if is_pdf(p)]
        if not pdfs or not HAS_KNOWLEDGE:
            return build_text_context(attached)
        try:
            store = get_active_store()
            if store is None:
                return build_text_context(attached)
            doc_ids, total_chunks, failed_pdfs = [], 0, []
            for path in pdfs:
                entry = store.index_file(path)
                if entry and entry.get('chunks'):
                    doc_ids.append(entry['doc_id'])
                    total_chunks += entry['chunks']
                else:
                    failed_pdfs.append(path)
            # small docs fit whole (~3 chunks ≈ the old 12k-char cap)
            if not doc_ids or total_chunks <= 3:
                return build_text_context(attached)

            # embed fresh chunks ≤20s, then hybrid search (else BM25-only)
            embedder = None
            try:
                from Intelligence.knowledge.embeddings import get_default_embedder
                embedder = get_default_embedder()
                if embedder is not None and embedder.is_available():
                    store.embed_pending(embedder, budget_sec=20)
                else:
                    embedder = None
            except Exception:
                embedder = None

            query = raw or u"main content of the document"
            hits = store.search(query, top_k=5, embedder=embedder,
                                doc_ids=set(doc_ids))
            if not hits:
                return build_text_context(attached)

            parts = [u'=== Relevant excerpts from the attached documents ===']
            for i, hit in enumerate(hits):
                page_note = (u' — page {}'.format(hit['page'])
                             if hit.get('page') else u'')
                parts.append(u'[{}] {}{}:\n{}'.format(
                    i + 1, hit['file'], page_note, hit['text'][:900]))
            parts.append(
                u'=== End of excerpts (the full document is indexed; '
                u'answer from the excerpts, cite sources [n]) ===')

            # Images + unreadable PDFs still described the legacy way
            rest = [p for p in attached if (not is_pdf(p)) or p in failed_pdfs]
            if rest:
                extra = build_text_context(rest)
                if extra:
                    parts.append(extra)
            return u'\n\n'.join(parts)
        except Exception as ex:
            logger.debug(u"attachment RAG failed, legacy path: {}".format(_exc_text(ex)))
            return build_text_context(attached)

    def _run_knowledge_agent(self, raw, history):
        """Answer from the knowledge index with citations. WORKER THREAD.

        Streams via _stream_llm_turn. True = handled; False = _route_input falls
        through to the normal LLM path (no hits, or every provider mute).
        """
        try:
            from Intelligence.knowledge.knowledge_agent import (
                KnowledgeAgent, format_citation_line)
            from Intelligence.llm_router import LLMRouter

            embedder = None
            try:
                from Intelligence.knowledge.embeddings import get_default_embedder
                embedder = get_default_embedder()
                if embedder is not None and not embedder.is_available():
                    embedder = None
            except Exception:
                embedder = None

            router = LLMRouter()
            provider = router.get_active_provider()
            agent = KnowledgeAgent(embedder=embedder)

            # same grounding as every path: instructions + remembered conventions
            _p_instr, _p_mem = self.win._projects._project_prompt_blocks(memory_tool=False)
            proj_instructions = u"\n\n".join(
                [b for b in (_p_instr, _p_mem) if b])

            skills_block = u''
            try:
                _dec = getattr(self.win, '_agent_decision', None)
                if _dec and _dec.get('skill'):
                    from Intelligence.skills_engine import (
                        build_skills_block, get_skills_engine)
                    _ids = [_dec['skill']]
                    # explicit /slash invocation bypasses the agents: filter
                    if not _dec.get('skill_forced'):
                        _ids = get_skills_engine().filter_for_specialist(
                            _ids, 'knowledge')
                    skills_block = build_skills_block(_ids)
            except Exception:
                pass

            _rid = getattr(self.win, '_request_id', 0)
            if self.win._cancelled():
                self.win._finish_cancelled(_rid)
                return True          # MUST be True — see below

            def _chat(system_prompt, query):
                if self.win._cancelled():
                    return None
                return self.win._stream_llm_turn(
                    provider, router, list(history), system_prompt, query,
                    max_tokens=900)

            result = agent.answer(raw, history, _chat,
                                  project_instructions=proj_instructions,
                                  skills_block=skills_block,
                                  viet=_is_viet_text(raw))

            # Every cancel exit returns True: False would fall through and start
            # a new LLM request right after the user pressed Stop.
            if self.win._cancelled():
                try:
                    self.win._ui_invoke(self.win._remove_stream_bubble)
                except Exception:
                    pass
                self.win._finish_cancelled(_rid)
                return True

            if result.get('status') != 'done':
                # remove any half-made bubble before falling through
                if result.get('status') == 'llm_failed':
                    def _cleanup():
                        self.win._remove_stream_bubble()
                    try:
                        self.win.Dispatcher.Invoke(Action(_cleanup))
                    except Exception:
                        pass
                return False

            text = result['text']
            shown = text + format_citation_line(result.get('citations'))

            def _done():
                self.win._hide_typing_indicator()
                if self.win._stream_tb is not None:
                    self.win._finalize_stream_bubble(shown)
                    self.win._clear_stream_refs()
                else:
                    self.win._append_bot_message(shown)
                self.win._add_to_history("assistant", text)
                self.win._set_busy(False)
            self.win.Dispatcher.Invoke(Action(_done))
            return True
        except Exception as ex:
            # falling through is the DESIGNED degradation (no bubble); log at error
            logger.error("_run_knowledge_agent error: {}".format(_exc_text(ex)))
            return False

    def _run_comment_agent(self, pdf_path, history):
        """PDF markup-comment workflow. WORKER THREAD.

        Annotations → trace sheet/model qua MCP → per-item proposal → report
        card (Run/Note/Skip). False (e.g. no annotations) = analyse the PDF as
        a normal document instead.
        """
        try:
            from Intelligence.comments.comment_agent import CommentAgent
            from Intelligence.llm_router import LLMRouter
            try:
                from core.server import get_t3labai_server
                srv = get_t3labai_server()
            except Exception:
                return False

            router = LLMRouter()
            provider = router.get_active_provider()

            skills_block = u''
            try:
                from Intelligence.skills_engine import build_skills_block
                skills_block = build_skills_block(
                    ['comment-resolution-playbook'])
            except Exception:
                pass

            agent = CommentAgent()
            _rid = getattr(self.win, '_request_id', 0)
            if self.win._cancelled():
                self.win._finish_cancelled(_rid)
                return True

            # project grounding (conventions); extra_context joins its BỐI CẢNH block
            _p_instr, _p_mem = self.win._projects._project_prompt_blocks(memory_tool=False)
            _extra = u"\n\n".join([b for b in (_p_instr, _p_mem) if b])

            report = agent.analyze(
                pdf_path, srv._execute_tool, provider, router,
                progress_cb=self.win._safe_update_typing_text,
                skills_block=skills_block,
                extra_context=_extra)

            # True, not False: False would start document analysis after Stop
            if self.win._cancelled():
                self.win._finish_cancelled(_rid)
                return True

            if report.get('error') == 'no_annotations':
                return False   # analyze as a normal document instead

            self.win._comment_agent = agent
            self.win._comment_report = report

            def _done():
                try:
                    self.win._hide_typing_indicator()
                    rendered = False
                    try:
                        from GUI.AssistantCards import build_comment_report_card
                        card = build_comment_report_card(
                            report,
                            on_run=lambda item, setter:
                                self._execute_comment_item(item, setter, 'run'),
                            on_note=lambda item, setter:
                                self._execute_comment_item(item, setter, 'note'),
                            on_skip=lambda item, setter:
                                self._skip_comment_item(item, setter),
                            viet=_ui_viet())
                        self.win.chat_history_panel.Children.Add(card)
                        self.win._scroll_to_bottom()
                        rendered = True
                    except Exception as cex:
                        logger.debug(u"comment card render failed: {}".format(_exc_text(cex)))
                    md = agent.report_to_markdown(report)
                    if not rendered:
                        self.win._append_bot_message(md)
                    self.win._add_to_history("assistant", md)
                    if report.get('needs_switch'):
                        self.win._append_bot_message(
                            u"This sheet is not in the active model — "
                            u"ask me to \"switch to model ...\" first, "
                            u"then press Run.",
                            icon=_ICON_SYNC, icon_color=_ICON_SLATE)
                    self.win._set_busy(False)
                except Exception as dex:
                    self.win._report_error(
                        u"comment report render", dex,
                        hint=u"The comment report could not be displayed.")
            self.win.Dispatcher.Invoke(Action(_done))
            return True
        except Exception as ex:
            # designed degradation as in _run_knowledge_agent; log visibly
            logger.error("_run_comment_agent error: {}".format(_exc_text(ex)))
            return False

    def _skip_comment_item(self, item, setter):
        """Skip button on a comment row — mark only. UI THREAD."""
        try:
            _v = _ui_viet()
            item['status'] = 'skipped'
            setter(u"Đã bỏ qua" if _v else u"Skipped", False)
        except Exception:
            pass

    def _execute_comment_item(self, item, setter, mode):
        """Run/Note on a comment row. UI THREAD; runs revit_action on a worker."""
        try:
            _v = _ui_viet()
            if self.win._busy:
                setter(u"Đang bận — chờ yêu cầu hiện tại" if _v
                       else u"Busy — wait for the current request", False)
                return
            agent = getattr(self.win, '_comment_agent', None)
            report = getattr(self.win, '_comment_report', None)
            if agent is None or report is None:
                setter(u"Báo cáo này không còn hiệu lực" if _v
                       else u"This report is no longer valid", False)
                return
            instruction = (agent.build_run_instruction(item, report)
                           if mode == 'run'
                           else agent.build_note_instruction(item, report))
            setter(u"Đang chạy..." if _v else u"Running...", False)
            self.win._set_busy(True)

            def _work():
                handled = False
                try:
                    from Intelligence.llm_router import LLMRouter
                    provider = LLMRouter().get_active_provider()
                    spec = get_spec('revit_action') if HAS_SPECIALISTS else None
                    handled = self.win._run_native_agent(
                        provider, list(self.win._conversation_history),
                        instruction, spec=spec)
                except Exception as wex:
                    logger.debug(u"comment exec error: {}".format(_exc_text(wex)))

                def _after():
                    try:
                        if handled:
                            item['status'] = ('done' if mode == 'run'
                                              else 'noted')
                            if mode == 'run':
                                setter(u"Xong" if _v else u"Done", True)
                            else:
                                setter(u"Đã ghi chú" if _v else u"Noted", True)
                        else:
                            setter(u"Thất bại — thử lại" if _v
                                   else u"Failed — try again", False)
                            self.win._set_busy(False)
                    except Exception:
                        pass
                try:
                    self.win.Dispatcher.Invoke(Action(_after))
                except Exception:
                    pass
            _ct = Thread(ThreadStart(_work))
            _ct.IsBackground = True
            _ct.SetApartmentState(ApartmentState.STA)
            _ct.Start()
        except Exception as ex:
            logger.debug(u"_execute_comment_item error: {}".format(_exc_text(ex)))

    def _build_knowledge_reference(self, query, local=False, history=None):
        """Compact project-knowledge block for the tool agent, or ''.

        BM25-only (fast, deterministic); smaller budget for local models.
        `history` resolves follow-ups ("còn tầng 2 thì sao?").
        """
        if not HAS_KNOWLEDGE:
            return u''
        try:
            from Intelligence.knowledge.knowledge_agent import (
                KnowledgeAgent, build_reference_block)
            agent = KnowledgeAgent(embedder=None)
            top_k = 2 if local else 3
            hits = agent.retrieve(query, top_k=top_k, history=history)
            if not hits:
                return u''
            return build_reference_block(
                hits, excerpt_chars=(500 if local else 800), max_items=top_k)
        except Exception as ex:
            logger.debug(u"knowledge reference build error: {}".format(_exc_text(ex)))
            return u''
