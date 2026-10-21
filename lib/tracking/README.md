# lib/tracking -- usage tracking (Lite only)

All the code that feeds the T3Lab Space usage dashboard
(`https://t3lab.space/api/revit/tracking`) lives here.

**t3lab_dev has no copy of this folder.** Do not delete it, and do not replace
it from dev. When you update `lib/` from dev, replace only the sub-folders you
need (`lib/GUI`, `lib/core`, ...), never `lib/` as a whole, or this folder
goes with it. `lite_guard/sync_from_dev.py` never copies into it, and
`lite_guard/guard.py` fails the commit if a file here goes missing.

| File | What it does |
|---|---|
| `hooks.py` | The three entry points below. They never raise and never block. |
| `service.py` | `TelemetryService`: builds the payload, anonymous user hash, sends it on a background thread. |
| `__init__.py` | Re-exports the entry points, so callers write `from tracking import track_...`. |

## Call sites (files shared with t3lab_dev)

Only these one-line calls live outside this folder. A copy from dev can still
drop them, so `lite_guard/manifest.json` checks each one.

| Entry point | Called from | Records |
|---|---|---|
| `track_session_start()` | `startup.py` | One row per Revit start / pyRevit reload |
| `track_ribbon_click()` | `lib/_cpython_bootstrap.py`, end of `init_cpython_paths()` | Every T3Lab ribbon button click |
| `track_mcp_call(...)` | `lib/core/server.py`, `_record_mcp_telemetry()` in `_handle_tool_call()` | Every MCP tool call, success and error |

If a dev copy overwrites one of those files, put the call back, wrapped in
`try: ... except Exception: pass` as it is now.

Tool folders (`T3Lab.tab/.../X.pushbutton`) hold no tracking code. A click is
tracked because the tool's `script.py` calls
`_cpython_bootstrap.init_cpython_paths()` near the top, which every tool does.
So a tool folder can be deleted and pasted from dev, as long as the dev
`script.py` keeps that call; `lite_guard` fails the commit if it does not.

## Settings

In `%APPDATA%\T3LabAI\mcp_paths.json`:

- `"tracking_enabled": false` turns tracking off.
- `"tracking_url": "..."` sends to another endpoint (e.g. a local test server).

`lib/Intelligence/telemetry.py` is not part of this. It is the T3Lab
Assistant's local timing log (JSONL on disk), it sends nothing, and it comes
from t3lab_dev.
