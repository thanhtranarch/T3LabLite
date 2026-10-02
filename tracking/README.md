# tracking -- usage tracking (Lite only)

All the code that feeds the T3Lab Space usage dashboard
(`https://t3lab.space/api/revit/tracking`) lives here.

**t3lab_dev has no copy of this folder.** Do not delete it, and do not replace
it from dev. It sits next to `lib/`, not inside it, so replacing `lib/` with
the dev copy does not touch it. `lite_guard/sync_from_dev.py` never copies into
it, and `lite_guard/guard.py` fails the commit if a file here goes missing.

| File | What it does |
|---|---|
| `tracking_hooks.py` | The three entry points below. They never raise and never block. |
| `tracking_service.py` | `TelemetryService`: builds the payload, anonymous user hash, sends it on a background thread. |

pyRevit only puts `lib/` on `sys.path`, so each call site adds this folder to
`sys.path` first. The modules carry a `tracking_` prefix so that adding the
folder cannot shadow any other module.

## Call sites (files shared with t3lab_dev)

Only these small blocks live outside this folder. A copy from dev can still
drop them, so `lite_guard/manifest.json` checks each one.

| Entry point | Called from | Records |
|---|---|---|
| `track_session_start()` | `startup.py` | One row per Revit start / pyRevit reload |
| `track_ribbon_click()` | `lib/_cpython_bootstrap.py`, end of `init_cpython_paths()` | Every T3Lab ribbon button click |
| `track_mcp_call(...)` | `lib/core/server.py`, `_record_mcp_telemetry()` in `_handle_tool_call()` | Every MCP tool call, success and error |

If a dev copy overwrites one of those files, put the block back as it is now:
add `<extension>/tracking` to `sys.path`, import from `tracking_hooks`, call,
all inside `try: ... except Exception: pass`.

## Settings

In `%APPDATA%\T3LabAI\mcp_paths.json`:

- `"tracking_enabled": false` turns tracking off.
- `"tracking_url": "..."` sends to another endpoint (e.g. a local test server).

`lib/Intelligence/telemetry.py` is not part of this. It is the T3Lab
Assistant's local timing log (JSONL on disk), it sends nothing, and it comes
from t3lab_dev.
