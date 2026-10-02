# lite_guard

T3LabLite is split from t3lab_dev, and code often comes across by copying
whole folders. A copy can overwrite a file that has Lite-only code, or a
"delete then paste" can remove a Lite-only tool. This happened in `5751bd2`:
the dev copies of `lib/_cpython_bootstrap.py` and `lib/core/server.py`
replaced the Lite ones, and the usage dashboard stopped getting ribbon and MCP
rows for a whole release.

`lite_guard` stops that in three places:

| Step | Tool | What it does |
|---|---|---|
| Copy | `sync_from_dev.py` | Copies from t3lab_dev with a dry run first. Never deletes. Skips `lite_only` paths. Blocks a file that would drop a tracking hook or a ribbon layout entry. |
| Commit | `githooks/pre-commit` | Runs `guard.py --staged` and stops the commit on an error. |
| Push / PR | `.github/workflows/lite-guard.yml` | Runs `guard.py` on GitHub, in case the hook is not installed. |

## Where usage tracking lives

| File | Hook | Records |
|---|---|---|
| `lib/Services/telemetry_service.py` | `TelemetryService` (Lite only) | Sends rows to `https://t3lab.space/api/revit/tracking` on a background thread |
| `startup.py` | `record_tool_usage(tool_name='session_start')` | One row per Revit start / pyRevit reload |
| `lib/_cpython_bootstrap.py` | `init_cpython_paths()` -> `_track_caller_script()` | Every ribbon button click (every `script.py` calls `init_cpython_paths()`) |
| `lib/core/server.py` | `_handle_tool_call()` -> `_record_mcp_telemetry()` | Every MCP tool call, success and error |

Opt-out: `"tracking_enabled": false` in `%APPDATA%\T3LabAI\mcp_paths.json`.

`lib/Intelligence/telemetry.py` is a different thing: a local JSONL log of
T3Lab Assistant turn timings. It sends nothing.

## Setup (once per clone)

```
python lite_guard/guard.py --install-hook
```

This sets `core.hooksPath` to `lite_guard/githooks` for this clone only.
GitHub Desktop and VS Code both run the hook.

## Bringing code over from t3lab_dev

```
python lite_guard/sync_from_dev.py D:/t3lab_dev lib/GUI "T3Lab.tab/Modeling & Datum.panel"
python lite_guard/sync_from_dev.py D:/t3lab_dev lib/GUI "T3Lab.tab/Modeling & Datum.panel" --apply
```

- `NEW` / `CHANGED`: copied with `--apply`.
- `BLOCKED`: not copied. Merge it by hand with the `git diff --no-index`
  command it prints, then run `guard.py`. `--force` copies it anyway.
- `KEEP`: only Lite has it. Left as is.
- `SKIP`: listed under `lite_only` in `manifest.json`.

If you still copy in Explorer: copy and overwrite, never delete the folder
first, then run `python lite_guard/guard.py` before you commit.

## Checks

```
python lite_guard/guard.py            # working tree
python lite_guard/guard.py --staged   # what the next commit records
```

- `markers`: the tracking hooks above, and `GITHUB_REPO` in
  `lib/core/updater.py` still points at T3LabLite.
- `tools`: every ribbon folder in `manifest.json` still exists, and every
  button has its `script.py` / `bundle.yaml`.
- `imports`: ribbon scripts and `lib/` modules only import `lib/` modules
  that exist.
- `xaml`: every `.xaml` file named in Python code exists.

Errors stop the commit. Warnings (for example a layout entry with no folder)
do not.

## Changing what is protected

- Added or removed a tool on purpose: `python lite_guard/guard.py --update`,
  then commit `manifest.json` with the change.
- New Lite-only code in a file shared with dev: add a `markers` entry.
- New Lite-only file or folder: add it to `lite_only`.
- Skip the hook once: `git commit --no-verify`.
