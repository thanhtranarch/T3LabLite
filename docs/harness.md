# BIM Harness in T3Lab Lite — several AI agents, one task board

**Status (2026-11-23): ported from t3lab_dev and tested without Revit
(`lite_guard/test_harness*.py`, `test_conductor.py`). Not run in Revit yet**: nothing on
this page has been checked inside Revit (HUD, MCP Control tab, toast, Harness button,
`/harness` in the Assistant, three AI apps on one model). Use the checklist at the end.

The harness is a thin coordination layer: one shared task board, a lead plans, workers take
tasks, the lead (or the person at the machine) reviews. **It grants nothing.** Every Revit
call still meets Lite's own MCP policy (`lib/core/mcp_policy.py`: read-only mode, toolsets,
the "Yes" in Revit for destructive calls). The board can only *refuse* a write.

```
Claude Desktop (lead)          Codex / Antigravity (workers)         harness run (launched agents)
      open · plan · review          join · next · report                 anthropic · command · http
            └──────────── t3lab_harness (bridge.py, one per AI app) ────────────┘
                 harness.db (SQLite, bridge process only) ──► board.json (snapshot)
                                                                 │ read in Revit by:
                     Agent Cursors HUD · MCP Control > Harness · toast · /harness
```

## What was ported

| Piece | Lite file | From t3lab_dev |
|---|---|---|
| Task board, `board.json`, write check | `lib/core/harness_store.py` | same file, unchanged |
| Agents from other platforms + runner | `lib/core/harness_agents.py` | same file, unchanged |
| Harness Studio (browser canvas) | `lib/core/harness_studio.py`, `harness_studio.html` | same files, unchanged |
| Fair Revit queue at the server | `lib/core/conductor.py` | same file, unchanged |
| `t3lab_harness` tool, identity, write check, stamps, `harness` CLI | `lib/core/bridge.py` | written by hand for Lite's bridge |
| Conductor in `/mcp`, queue in `/health`, task on the cursor | `lib/core/server.py` | by hand |
| Board on the HUD, per-member cursors, task name on the cursor | `lib/core/agent_cursors.py`, `agent_cursors_overlay.py` | by hand (board parts only) |
| MCP Control > Harness tab (Accept / Send back / Open Studio) | `lib/GUI/Tools/MCPControl.xaml`, `lib/GUI/MCPControlDialog.py`, `lib/Services/harness_settings.py` | by hand |
| Toast when a task waits for review / fails / the board is done | `lib/harness_notifier.py` (+ `startup.py`) | board part of `approval_notifier.py` |
| `/harness` in the T3Lab Assistant | `lib/Intelligence/harness_client.py`, `lib/GUI/assistant/commands.py` | reduced (see below) |
| Ribbon Support > Harness (opens Studio) | `T3Lab.tab/Support.panel/Harness.pushbutton` | copied, Lite bootstrap |
| Deployed next to the bridge | `lib/bridge_deploy.py` `COMPANIONS` | by hand |

## Use it

### 1. Each AI app has its own name

MCP Control > Connect already writes `--agent=codex` for ChatGPT (Codex) and
`--agent=antigravity` for Antigravity; Claude Desktop runs as `claude-desktop`. Without
`--agent`, the bridge maps the client's `clientInfo` name (`CLIENT_AGENTS` in `bridge.py`:
`claude-ai` → `claude-desktop`, `codex-mcp-client` → `codex`, …), else the app that launched it,
else `mcp-client`. The name only labels the caller on the board and the cursors.

### 2. Optional: one agent leads (from a terminal on the Revit machine)

```powershell
python $env:APPDATA\T3LabAI\bridge.py harness setup --master=claude-desktop --write-mode=serial
python $env:APPDATA\T3LabAI\bridge.py harness setup            # show it
python $env:APPDATA\T3LabAI\bridge.py harness setup --master=  # no master
```

`master`: the only agent that may `open` sessions. `exclusive` (on with a master): every
other agent needs a task to change the session's file, joined or not (`harness_managed`).
`write_mode=serial`: one write task per file at a time. Only a terminal changes the setup
(`use_cli` from a chat).

### 3. In the lead app (Claude)

"Open a harness session on PRJ.rvt, goal …; plan: … for codex, … for antigravity." →
`t3lab_harness action=open` + `action=plan` (tasks with `assignee`, `depends_on`, `scope`).

### 4. In each worker app (Codex, Antigravity)

"Join the harness and take a task." → `action=join`, then `action=next` (`wait_s` up to 50 s
holds the call until a task is ready), work, `action=report status=done summary=…`.

### 5. Review

The lead: `action=review task=<key> verdict=accept|redo`. The person at the machine:
MCP Control > **Harness** (Accept / Send back), Harness Studio, or `/harness accept <task>` in
the Assistant — all go through `bridge.py harness review` with `owner:true`, which the bridge
accepts only from its CLI.

### Harness Studio

The same board as a page in the browser (127.0.0.1 only, token in the link): plan steps,
assign agents, review, save templates. Ribbon **Support > Harness**, MCP Control > Harness >
**Open Studio**, `/harness studio`, or `python %APPDATA%\T3LabAI\bridge.py harness studio`.
Started from Revit it stops after 10 idle minutes.

### Agents from other platforms

```powershell
python $env:APPDATA\T3LabAI\bridge.py harness agent_add --json-file=agent.json
python $env:APPDATA\T3LabAI\bridge.py harness agents
python $env:APPDATA\T3LabAI\bridge.py harness run --loop
```

Platforms `anthropic` (Console Managed Agent, its `t3lab_call` custom tool is answered
locally), `command` (`claude -p`, `codex exec` …), `http` (webhook), `mcp` (connects by
itself). Names of person-facing apps are reserved. `run` only from a terminal.

## How the write check works in Lite

`bridge.py` checks every `tools/call` it forwards, **before Revit**:

- A tool counts as a write unless the server's `tools/list` marks it `readOnlyHint` (from the
  cached manifest). The bridge's own tools and `t3lab_harness` / `t3lab_hub` never count.
- No open session on the machine → nothing is checked (as before the port).
- The file is the active document of the Revit window the call goes to (Lite has no
  `revit_document` pin per call: to work in another file, `switch_active_document`).
- `harness_store.write_refusal` then refuses (`nothing_ran: true`): a member with no task
  (`scope_denied`), a non-member in an `exclusive` session (`harness_managed`), a task with
  `scope.write=false`, another file than the task's, `element_ids` outside the task or held by
  another agent's task (`scope_claimed`). The lead is never limited. A dry run passes.
- A board that cannot be read lets the call through: Lite's policy still applies.
- While the harness is in use (a session open or a master set), forwarded calls carry
  `t3lab/task`, `t3lab/task_title`, `t3lab/member`, `t3lab/lane` and `t3lab/queue_wait_s` (45 s).
  Without the harness nothing extra is sent.

The server replaces its bare request lock with the conductor: one Revit slot, lead / person
first, round robin per agent, and a call stamped with a wait budget that cannot start in time
gets `revit_busy` (`nothing_ran`) instead of running after the client gave up. A call without
a budget (no harness, an old bridge) waits as before. `GET /health` shows `queue`.

## What differs from t3lab_dev

| t3lab_dev | T3Lab Lite |
|---|---|
| Per-agent profiles, tiers, approval queue (`pending_approval`, `t3lab_approval`) | Not ported. Lite's policy is `mcp_policy.py` and the confirmation in Revit. Review is on the board only (`review`, Accept / Send back); Studio shows no approvals. |
| Launched agents keep `read` + approval for writes | Launched agents (`harness run`) **only read**: a write through `t3lab_call` is refused (`tool_denied`). A person-facing app takes the write tasks. |
| `t3lab_capabilities` lists sessions and your task | Lite has no `t3lab_capabilities`: `t3lab_harness action=status` / `help`, and the refusal says to join. |
| `revit_document` pin per call; cloud agents pinned to their task's file | No pin in Lite: the active document decides. |
| `t3lab/task` in the MCP audit log; approvals per task | Lite has no MCP audit log: the stamps name the task on Agent Cursors only. |
| An in-Revit approval dialog lends the conductor's slot | Not in Lite: while Revit asks for a "Yes" the call keeps the slot (other calls wait, or get `revit_busy` after 45 s). |
| `note remember=true`, close → model memory | Not ported (`model_memory.py` is dev only). Notes stay on the board. |
| `/harness plan / join / next / done / failed`: the Assistant works tasks | Lite's Assistant only looks and reviews (`/harness`, `studio`, `accept`, `redo`): its edits reach Revit in-process, past the bridge's check, so it never becomes a member. |
| HUD file strip (who works in which file) | Not ported; the board section and per-member cursors are. |
| `stdio` bridge answers calls in parallel | Lite's bridge answers one request at a time: `next wait_s` holds that app's other calls for up to 50 s. |

Both extensions share `%APPDATA%\T3LabAI\harness\` (same board format) and the toast watcher
key, so a machine with both sees one board.

## Errors

`session_required`, `session_not_found`, `session_closed`, `not_a_member`, `role_denied`,
`task_not_found`, `task_started`, `not_owner`, `not_in_review`, `bad_dependency`,
`too_many_tasks`, `scope_denied`, `scope_claimed`, `harness_managed`, `revit_busy`, `use_cli`,
`reserved_name`, `sdk_missing`, `tool_denied`, `harness_unavailable`.

## Tests (no Revit)

```
python lite_guard/test_harness.py            # board, agents, bridge identity / check / stamps / CLI
python lite_guard/test_harness_studio.py     # Studio: owner seat, flows, runner, HTTP API
python lite_guard/test_conductor.py          # queue order, wait budget, server wiring
python lite_guard/test_harness_settings.py   # MCP Control rows, review via the CLI, tab wiring
python lite_guard/test_harness_notifier.py   # toasts
python lite_guard/test_harness_client.py     # /harness in the Assistant
```

## Revit test checklist (NOT RUN)

Restart Revit after updating (do not Reload). Revit 2025 or 2026, then 2022 or 2023.

- [ ] Revit starts with no new error in `%APPDATA%\T3LabAI\engine_check.log`; `bridge.py`,
      `harness_store.py`, `harness_agents.py`, `harness_studio.py`, `harness_studio.html` are in
      `%APPDATA%\T3LabAI\` (after MCP Control > Connect if another extension owns the bridge).
- [ ] Ribbon: Support shows **Harness** between the cloud links and T3Lab Assistant; its icon
      reads in light and dark theme. Click → the browser opens Harness Studio; a second click
      reopens the same page.
- [ ] MCP Control: the **Harness** chip opens the tab; with no session the empty state shows.
      Open a session from Claude, plan two tasks, have Codex `join` + `next` + `report`: the
      grid lists the task first as *Waiting for review*; **Accept** and **Send back** work and
      the status line says so with the title.
- [ ] Toast: the reported task raises "… finished … It waits for review"; a failed report and
      a finished board raise their toasts once.
- [ ] Agent Cursors on: the HUD shows "Harness · <goal>", "1 working · 1 in review · 0/2 done"
      and who holds which task; a worker's cursor says "Task a · <title>"; two chats of Claude
      with `as="arch"` / `as="mep"` get two cursors.
- [ ] Write check: Codex without a task tries a write → `scope_denied`, nothing changes in the
      model; with its task it writes; outside the task's `element_ids` → `scope_denied`.
      With a master and `exclusive`, an app that never joined → `harness_managed`.
- [ ] Conductor: Claude and Codex call at the same time; no call runs after its app gave up; a
      call that waited 45 s returns `revit_busy`. A destructive call's "Yes" dialog still works.
- [ ] Assistant: `/harness` shows the board; `/harness accept <task>` reviews; `/harness next`
      says it is not in Lite; `/harness studio` opens Studio.
- [ ] Without any session: Claude Desktop alone works exactly as before (no new refusals, no
      `revit_busy`).
