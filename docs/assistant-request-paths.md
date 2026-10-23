# T3Lab Assistant — every path a chat message can take

What happens between the user pressing **Enter** and an answer appearing, as
of the 2026-10-23 split of `lib/GUI/T3LabAssistantDialog.py`. Use it to find
which code (and which setting) is responsible for a given answer.

## Where the code lives

| Piece | File |
|---|---|
| Window class, the turn pipeline (`_process_input`, `_route_input`, `_execute_result`, `_run_native_agent`, streaming, rendering) | `lib/GUI/T3LabAssistantDialog.py` |
| Deterministic commands (`/memory`, `/train`, `/toolhub`, `/skills`) | `lib/GUI/assistant/commands.py` — `AssistantCommands` (`self._commands`) |
| Links, attachment RAG, knowledge agent, PDF-comment agent | `lib/GUI/assistant/documents.py` — `AssistantDocuments` (`self._docs`) |
| Graph path: multi-goal plans, plan approval, progress / task cards | `lib/GUI/assistant/plans.py` — `AssistantPlans` (`self._plans`) |
| `/english-spellcheck` scan | `lib/GUI/assistant/spellcheck.py` — `AssistantSpellcheck` (`self._spell`) |
| Tool cards + per-call confirm gate | `lib/GUI/assistant/tool_cards.py` — `AssistantToolCards` (`self._tool_cards`) |
| History / sessions, projects | `lib/GUI/assistant/sessions.py`, `lib/GUI/assistant/projects.py` |
| Optional imports + stubs (`HAS_NLP`, `HAS_AGENTS`, ...) | `lib/GUI/assistant/deps.py` |

The helpers are plain classes built in the window's `__init__`
(`self._docs = AssistantDocuments(self)`); their code reaches the window as
`self.win`. No mixins — a CLR-derived window cannot take one (CLAUDE.md
rule 6). `lite_guard/test_assistant_structure.py` checks the wiring.

## Entry points

Every way a message is sent ends in `T3LabAssistantWindow._process_input()`
(UI thread), which reads `chat_input.Text`:

| Trigger | Caller |
|---|---|
| Send button (`Click="send_clicked"`) | `send_clicked` |
| Enter in the composer (`KeyDown="input_keydown"`) | `input_keydown` (Enter on an open slash popup accepts the skill instead) |
| Message typed while a turn runs | queued by `_queue_input`, sent by `_drain_queued_input` when busy clears |
| Quick-reply chip ("Continue", "Yes, go ahead") | `_append_quick_replies` → chip click |
| Suggestion card | `suggestion_card_clicked` |
| Per-project scheduled prompt | `AssistantProjects._schedule_tick` (30 s timer) |

`_process_input`:

1. Records the message language (`lang._note_user_language`).
2. **Busy?** A plan card that is waiting is superseded (`_PLAN_SUPERSEDED`);
   the text is queued (attachments are refused). Nothing else runs.
3. Parses `/skill-id [rest]` against the skills catalog
   (`Intelligence.skills_engine`); a known id sets `_forced_skill_id` and, for
   a bare invocation, rewrites the routed text into an explicit instruction
   (reference skill / model-changing skill / read-only skill).
4. Echoes the user bubble, appends to history, clears attachments, sets busy.
5. Starts a worker thread: archives attachments into the active project,
   logs the request, then calls **`_route_input(text, files)`**. Any exception
   becomes `_report_error` (or `_finish_cancelled` after Stop).

## The routing ladder — `_route_input` (worker thread)

Checked top to bottom; the first rung that handles the message ends the turn.

| # | Path | Taken when | Runs |
|---|---|---|---|
| 1 | **`/memory`, "remember that …"** | no attachments, `_MEM_SAVE_RX` or `/memory …` | `AssistantCommands._try_memory_command` → `Intelligence.assistant_memory`, `Intelligence.session_context` |
| 2 | **`/train`** | no attachments, `/train …` | `AssistantCommands._try_train_command` → `Intelligence.learning.trainer` / `exemplars` |
| 3 | **"which model are you?"** | no attachments, question about the assistant's own model | `_try_which_model` / `_answer_which_model` (window) |
| 4 | **`/skills`, "install skill from <repo link>"** | no attachments | `AssistantCommands._try_skills_command` → `Intelligence.skill_installer`, `skills_engine` |
| 5 | **`/toolhub`** | no attachments | `AssistantCommands._try_toolhub_command` → `Intelligence.toolhub.commands` |
| — | *links* | `HAS_LINKS` and the text has a URL / UNC / path | `AssistantDocuments._resolve_message_links` (`Intelligence.link_reader`) downloads readable files **into the attachments** and returns a notes block; not a terminal rung |
| 6 | **PDF markup comments** | `HAS_AGENTS`, attachments with an annotated PDF, and the keyword dispatcher says `comment` (or the message is empty) | `AssistantDocuments._run_comment_agent` → `Intelligence.comments.comment_agent.CommentAgent`; each item can run `_run_native_agent` (`_execute_comment_item`) |
| — | *attachment context* | `HAS_RAG` and attachments | `AssistantDocuments._build_attachment_context` → `Intelligence.knowledge` (large PDFs: indexed, top-5 excerpts) or `Intelligence.rag_processor.build_text_context` |
| 7 | **`/english-spellcheck` scan** | forced skill is `english-spellcheck` and `routing.wants_spellcheck_fix(args)` is False | `_execute_result({'intent': 'check_spelling'})` → `AssistantSpellcheck._run_spellcheck` → `Services.spell_checker` (reads EVERY text; the paged agent tools cannot) |
| 8 | **Learned pattern** | `HAS_NLP`, no attachments, no forced skill, not a carry-over, `find_learned_match` hits and `routing.learned_pattern_wins` agrees | `_execute_result(learned)` |
| 9 | **Rule-based NLU** | same gate as 8; `parse_command_nlu` (→ `Intelligence.nlu_engine`) returns a known intent that is not chat/help, or is `_authoritative` / `_instant`, or no provider is usable | `_execute_result(nlu_result)` (chat/help with an unreachable provider gets `_provider_unreachable_note`) |
| 10 | **Knowledge specialist** | `HAS_AGENTS`, a usable provider, no attachments, `agents.multi_agent` on, dispatcher picks `knowledge`, `HAS_KNOWLEDGE` | `AssistantDocuments._run_knowledge_agent` → `Intelligence.knowledge.knowledge_agent.KnowledgeAgent`, streamed with `_stream_llm_turn`. No hits / provider mute → falls through |
| 11 | **Comment specialist without a PDF** | as 10, dispatcher picks `comment` | canned "attach the marked-up PDF" message |
| — | *auto model* | `use_claude` | `_auto_route_model` → `Intelligence.model_tier` picks Haiku / Sonnet / Opus for the rest of the turn |
| 12 | **Graph (multi-goal) path** | no attachments, no forced skill, not a carry-over, a usable provider, `AssistantPlans._graph_enabled()` (HAS_GRAPH + HAS_AGENTS + HAS_SPECIALISTS, `agents.graph_mode` and `agents.multi_agent` on), the provider has native tools, and `GraphOrchestrator.plan_for` finds **more than one goal** | `AssistantPlans._run_graph_plan` → `Intelligence.graph` (`orchestrator`, `planner`, `executor`, `verifier`, `reducer`). Three shapes: (a) plan with a model-changing step → plan card (`_await_plan_approval`, Run / Edit / Cancel), then sequential nodes, each one `_run_native_agent(hold_busy=True)`; (b) all-read plan with `agents.parallel_tasks` on → `_run_parallel_tasks` (`Intelligence.agents.task_manager`, each node an `AgentLoop` with read-only tools); (c) otherwise sequential nodes |
| 13 | **Native agent loop** | a usable provider (or attachments), and either no attachment/RAG text or compact RAG (< 4000 chars, images only on a vision provider) | `_run_native_agent` → `Intelligence.agent_loop.AgentLoop` with `Intelligence.tool_schema` tools, specialist spec (`Intelligence.agents.specialists`) when the dispatcher chose `revit_data` / `revit_action` / `multi_doc` / `modeling` / `qa_check` / `export`, top-2 skills. Returns False (→ 14) when the provider lacks native tools, the MCP server or the tool registry is unavailable, turn 1 fails with nothing shown, or the model returns an empty turn |
| 14 | **Legacy JSON-intent loop** | rung 13 declined or raised, or the attachment context is large / has images the provider cannot see | inline loop in `_route_input` (`do_nlp`): up to the specialist's `max_iterations` (10) of `{"intent", "params", "message"}` JSON from `_build_system_prompt` (`Intelligence.t3lab_assistant`), MCP tools through `AssistantToolCards._run_gated_tool`, `describe_tool` meta-intent, repeat guards; then `_execute_result`, else the NLU hint, else `keyword_parse`, else "did not understand" / "model did not respond" |
| 15 | **No provider configured** | no usable provider and no attachments | `keyword_parse(raw)` → `_execute_result`, else `get_setup_guidance_message` |

`use_local = HAS_NLP and has_local_llm()`; `use_claude = HAS_NLP and
has_api_key()` (true for ANY configured cloud provider, not just Claude).

### `_execute_result` — where deterministic answers land

Called by rungs 7, 8, 9, 14, 15 and by `_run_native_agent` when the model asks
to open a ribbon tool (`launch_intent`). It refuses to launch tools for small
talk (`nlu_engine.is_conversational`), then: chat / help / greet → message;
`export_direct` / `open_batchout_configured` → `Services.batchout_executor`;
`check_spelling` → `AssistantSpellcheck._run_spellcheck`; an intent in
`TOOL_LAUNCHERS` (`lib/GUI/assistant/launchers.py`) → `_launch_tool` through
`Services.revit_context.run_in_api_context`; unknown → guidance. Learns the
pattern (`learn_pattern`) for non-chat intents.

## Settings that select a path

All in `%APPDATA%\T3LabAI\settings.json`, read through
`lib/config/settings.py`. `multi_agent`, `llm_classify`, `graph_mode` and
`parallel_tasks` have no UI — they are kill switches set by editing the file;
`auto_model`, `quality_mode` and `action_mode` are toggles in LLMs Setting.

| Setting | Default | Effect on the ladder |
|---|---|---|
| `agents.multi_agent` (`is_multi_agent_enabled`) | on | off: no dispatcher → rungs 10–11 and 12 never run; rung 13 runs as `general` |
| `agents.llm_classify` (`is_llm_classify_enabled`) | on | off: dispatcher uses keywords only (no tiny classify call) |
| `agents.graph_mode` (`is_graph_mode_enabled`) | on | off: rung 12 never runs |
| `agents.parallel_tasks` (`get_agent_option`) | on | off: all-read multi-goal plans run sequentially (12c) instead of as task cards (12b) |
| `agents.auto_model` (`is_auto_model_enabled`, via `model_tier.is_enabled`) | on | off (or `agents.quality_mode` on, or a non-Claude provider): no per-turn Haiku/Sonnet/Opus pick |
| `agents.action_mode` (`get_action_mode`; Action-mode chip) | auto | `confirm` ("Ask before edits"): every model-changing tool call waits on a confirm card (`AssistantToolCards._confirm_tool_blocking`) |
| `reply_language` | auto | language of every answer and card |
| active provider / API key / local engine | — | decides `use_claude` / `use_local`, i.e. rungs 13–15 |

`ai_mode` (`is_ai_mode_enabled`) is **not** read by the chat window; it gates
the AI features of other tools (`Intelligence/ai_bridge.py`).

Import-time flags (`HAS_NLP`, `HAS_AGENTS`, `HAS_GRAPH`, `HAS_SPECIALISTS`,
`HAS_KNOWLEDGE`, `HAS_RAG`, `HAS_LINKS`, `HAS_LANGUAGE` in
`lib/GUI/assistant/deps.py`) are True on a healthy install; a failed import
logs a warning and removes the rungs that need it.

## Paths that were checked for removal

Nothing was removed. Every rung above is reachable:

- Rungs 8/9 (learned + NLU) run on every plain message while `HAS_NLP`.
- Rung 12 is on by default; its three shapes depend only on the plan.
- Rung 14 is reachable with every provider: all providers declare
  `SUPPORTS_NATIVE_TOOLS = True`, but large attachment context, images on a
  text-only model, an unavailable MCP server / tool registry, a native-loop
  exception, a failed first turn or an empty model turn all send the turn
  to it.
- Rung 15 is the no-provider install.

Dead code found on the way (left in place, it is not a request path and the
same file is shared with t3lab_dev):

- `_suppress_user_echo` is initialised `False` and only ever reset to `False`
  ("Try again" was removed with the message action bar —
  `_make_message_actions` returns `None`), so the "don't echo the user
  message" branch in `_process_input` never runs. `_adopt_retry_version`,
  `_refresh_reply_row` and `_vote_decision` belong to the same disabled bar.
- `reset_chat_clicked` has no caller and no XAML wiring (New chat is
  `new_chat_clicked`).
- `parse_command` / `parse_command_local` (`Intelligence.t3lab_assistant`)
  are imported by the dialog but no chat path calls them.
- `self._offer_local_engine()` in `__init__`'s startup probe names a method
  that does not exist; the surrounding `try` logs "Local engine offer
  failed" on every open. Listed in `KNOWN_UNRESOLVED` of the structure test.
- With the NLP layer missing (`HAS_NLP` False), rung 15 calls
  `keyword_parse`, which is then undefined: the NameError surfaces as an
  error bubble. Pre-existing; only matters on a broken install.

## Proposal: one `AgentLoop` path (not done — it changes behaviour)

Rungs 13 and 14 do the same job twice: a tool-calling loop over the MCP
tools. Rung 14 exists for providers without native tool calling and for
bulky attachment context. Folding it into `AgentLoop` would:

1. **Attachments:** send large RAG context through the native loop as a
   retrieval tool (`search_attachments(query)` over the already-indexed
   knowledge store) instead of stuffing it into the prompt. Changes what the
   model sees for big PDFs.
2. **Images on text-only models:** keep the "[System: N images, no vision]"
   note, but in the native user turn. Same user-visible answer.
3. **No native tools / mute provider:** give `AgentLoop` a text-protocol
   adapter (parse `{"intent", "params"}` JSON from plain replies — the
   existing `agent_loop.is_bare_tool_call_text` rescue already does part of
   this) and drop `do_nlp`'s loop. Repeat guards, `describe_tool`, and the
   announce nudge already exist in `AgentLoop`.
4. **Deterministic rungs 8/9:** keep them (they are instant and offline), but
   only for `_authoritative` / `_instant` NLU results; let everything else
   reach the agent.

Each step changes prompts or tool exposure, so each needs a Revit session to
verify (local Ollama model + a cloud provider, with and without a large
PDF). Until then the legacy loop stays.
