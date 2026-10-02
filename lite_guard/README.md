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

All tracking code is in `lib/tracking/` (Lite only, listed under `lite_only`).
See `lib/tracking/README.md`. Only three one-line calls sit in files shared
with t3lab_dev, and `markers` checks each one:

| File | Call | Records |
|---|---|---|
| `startup.py` | `track_session_start()` | One row per Revit start / pyRevit reload |
| `lib/_cpython_bootstrap.py` | `init_cpython_paths()` -> `track_ribbon_click()` | Every ribbon button click (every `script.py` calls `init_cpython_paths()`) |
| `lib/core/server.py` | `_handle_tool_call()` -> `_record_mcp_telemetry()` -> `track_mcp_call()` | Every MCP tool call, success and error |

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

If you copy in Explorer: replace only the sub-folders of `lib/` you are
updating (delete `lib/GUI`, paste the dev `lib/GUI`), never `lib/` as a whole
and never `lib/tracking/`. Replacing a whole tool folder
(`T3Lab.tab/.../X.pushbutton`) is fine as long as the dev `script.py` still
calls `_cpython_bootstrap.init_cpython_paths()`. Lite-only files inside a replaced sub-folder are
lost too, so run `python lite_guard/guard.py` before you commit.

## Checks

```
python lite_guard/guard.py            # working tree
python lite_guard/guard.py --staged   # what the next commit records
```

- `markers`: the files in `lib/tracking/` and the three calls above are
  still there, and `GITHUB_REPO` in `lib/core/updater.py` still points at
  T3LabLite.
- `tools`: every ribbon folder in `manifest.json` still exists, every
  button has its `script.py` / `bundle.yaml`, and every `script.py` calls
  `_cpython_bootstrap.init_cpython_paths()` (that call is what tracks the
  click; tool folders hold no tracking code of their own).
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
