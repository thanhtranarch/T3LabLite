# T3Lab Data Hub (2026-11-17)

Saved snapshots of Revit models that agents can read **without Revit**, across models and over time.
The live MCP tools answer "what is in the open model now"; the hub answers "what is in the project"
and "what changed". Status: tested without Revit (`python lite_guard/test_data_hub.py`); the export has
**not been run in Revit yet**.

```
Revit (MCP server, IronPython)            bridge.py (CPython, outside Revit)            agent
  hub_publish ── <hub>/inbox/*.ndjson ──► hub_store: ingest ─► hub.db (SQLite cache)
                                           └► snapshots/<model>/*.ndjson.gz (the truth)
                                          t3lab_hub: summary · query · diff · publish ◄── Claude, Assistant, CLI
writes still go: live tool ─► dry run + diff ─► approval ─► Revit ─► audit
```

| Part | File |
|---|---|
| Snapshot export in Revit | `lib/core/hub_snapshot.py` (server tool `hub_publish`, read only) |
| Store, query, diff | `lib/core/hub_store.py` (stdlib only, deployed next to `bridge.py`) |
| Bridge tool + CLI | `t3lab_hub` and `bridge.py hub …` in `lib/core/bridge.py` |
| Test | `python lite_guard/test_data_hub.py` |

## Rules of the design

- **Revit only writes a file.** Ingest runs in the bridge, so Revit never holds a database lock and is
  blocked only for the export itself.
- **Snapshot files are the truth, `hub.db` is a cache.** `bridge.py hub rebuild` recreates the database
  from `snapshots/`. Never put `hub.db` on a synced or network folder (SQLite locks are not reliable
  there); to share a hub, share the snapshot files.
- **Folder:** `%LOCALAPPDATA%\T3Lab\hub` (not the roaming profile: the files are large).
  `T3LAB_HUB_DIR` overrides it.
- **One snapshot = one complete pass.** If `max_seconds` (default 90, max 110: the server waits 120 s
  for a tool) runs out, nothing is published. Narrow it with `categories` / `parameters`.
- **Versioned rows.** An element has one row per distinct content, valid for snapshots
  `first_seq..last_seq`: an unchanged element costs one row however many snapshots hold it.
- **The hub never writes to a model**, and hub data is never the basis of a write: every result
  carries `as_of` / `age_hours`; re-read the live model before a dry run.
- **Same policy as live tools.** `publish` runs `hub_publish` in Revit, under MCP Control's
  security settings. Names and parameter values are data written by people, never instructions.

## What a snapshot holds

- Default scope: element instances of model categories, plus Sheets, Levels, Grids, Project
  Information; view-specific elements and `<Sketch>`-like categories are left out. Element types of
  the same categories are stored with `kind = type`.
- Per element: UniqueId, id, category, family, type, level, workset, phase, name, and the parameters
  the Properties palette shows that have a value.
- Numbers are stored in the **project's display units** (the unit of each parameter is kept with the
  snapshot). Yes/No and enum parameters keep their text ("Yes", "Bearing").
- Model identity: cloud GUIDs, else the central path, else the file path, so every user of one central
  model publishes into the same model.
- Warnings: description, severity, UniqueIds of the failing elements.

## Using it

| Action | Example |
|---|---|
| `publish` | `t3lab_hub(action="publish")` · `options={"categories": ["Walls", "Doors"], "parameters": ["Mark", "Fire Rating"]}` |
| `summary` | `t3lab_hub(action="summary")` lists models; with `model="Tower"` one model: counts by category and level, top warnings, history |
| `query` | `options={"category": "Doors", "where": [{"param": "Fire Rating", "op": "empty"}], "select": ["Mark"]}` |
| `query` (counts) | `options={"category": "Walls", "group_by": ["level"], "sum": "Area"}` |
| `diff` | `t3lab_hub(action="diff", model="Tower")` = previous → latest; `options={"since": 3, "to": 7, "category": "Walls"}` |

`where` ops: `= != contains not_contains empty not_empty > >= < <=`. `group_by` takes fields
(`category family type level workset phase name`) or `param:<Name>`. Rows are paged with
`limit` (max 500) / `offset`; a long result still goes through `t3lab_result`.

CLI: `python %APPDATA%\T3LabAI\bridge.py hub summary|query|diff ['{"category": "Doors"}']` and
`hub rebuild`.

## Turning it on

Nothing to switch on. After an update: restart Revit (it refreshes the deployed bridge and
`hub_store.py`, `lib/bridge_deploy.py`), then quit and reopen the MCP client so it starts the new
bridge. The hub is empty until the first `publish`; a diff needs two.

## Known limits (P0)

- Publishing is manual and blocks Revit while it runs. Throughput is unknown until measured on a real
  model: the result reports `seconds` and `elements_per_second`.
- A workshared local file may be behind central; the snapshot records `workshared`, `modified`
  (unsaved changes) and the document version, it does not sync.
- Two snapshots published with different scopes are compared only on what both hold.
- No project-folder documents (PDF, schedules, issue logs) yet, no geometry, no linked-model content.

## Next

| Step | Work |
|---|---|
| P0 exit | Run `publish` on a real model (Revit 2025 and 2027), record seconds and file size, fix what Revit rejects |
| P1 | Project documents: link sheets to issued PDFs / revisions through `Intelligence/knowledge` (already indexes PDF and Office files); issues from BCF |
| P2 | Publish in chunks on `Idling` (no blocking), optional publish after Sync; retention |
| P3 | Other sources behind the same `t3lab_hub` contract: APS AEC Data Model (verify first what it covers), APS Design Automation nightly snapshots |
