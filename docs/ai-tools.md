# AI tools from ribbon tools (ported from t3lab_dev, 2026-11-23)

The ribbon tools' own logic, callable by an AI app through MCP and by the T3Lab Assistant without
opening the window. Ported from t3lab-revit-api PR #291 (ideas from HKUDS/CLI-Anything, Apache-2.0,
no code copied). **Not run in Revit yet** — see "Verify in Revit" below.

## How a tool is declared

One `ToolSpec` per tool in `lib/core/tools/<domain>.py`, collected by `lib/core/tool_specs.py`.
Its fields feed the MCP registry and dispatch (`lib/core/server.py`), the main-thread list
(`_WRITE_TOOLS`), the confirmation rule (`is_destructive`, which Lite's `core/mcp_policy.py` reads),
the read-only list, the Assistant subsets and prompt (`lib/Intelligence/`), the read-back
(`lib/core/write_verify.py`) and the Vietnamese titles. Never hand-add such a tool to those lists.

- `requires=(...)`: lib-relative files the tool needs. A spec whose files are missing is not
  registered (`tool_specs.SKIPPED`). Lite ships no View Template Transfer, so
  `transfer_view_templates` stays off.
- Every write tool has `dry_run` and a read-back mode; `ToolSpec.problems()` refuses one without.
- The MCP server runs IronPython 2.7: `python lite_guard/audit_ipy_compat.py --quiet` checks every
  module it can import (no f-strings, `casefold`, `open(encoding=)`, `json.dumps` without
  `ensure_ascii=False`, `except OSError` around `open`, utf-8-sig writes; use `Snippets._compat.fold`,
  `replace_file`, `io.open`, `core.jsonsafe.dumps`).

## Tools (23)

| Tool | From | Notes |
|---|---|---|
| `export_sheets` | BatchOut | PDF / DWG / DXF / DGN by Revit sheet set, numbers, prefix; BatchOut profiles and naming; every file checked |
| `rebar_wizard`, `inspect_rebar`, `rebar_tools`, `export_bvbs` | Rebar Wizard, Rebar Check, Cast Unit, Clone Drawing, BVBS Export | |
| `sync_views`, `generate_room_views` | Crop Sync, Datum Sync, SheetGen | |
| `auto_join`, `room_to_floor`, `rehost_to_level`, `cut_wall_openings`, `create_door_thresholds` | Auto Join, Room to Floor, Wall Adjust Base, Wall Cut Profile, Door Threshold | `room_to_floor` was hand-written in `server.py`; it is a spec now (same argument names, plus offset / structural / boundary) |
| `build_from_point_cloud`, `cad_to_elements`, `auto_dimension` | Point Cloud, CAD to Elements, AutoDimension | Lite's server has no `detect_point_cloud_elements`: call `build_from_point_cloud` without `detections` and it detects itself |
| `transfer_families`, `manage_groups`, `manage_worksets`, `smart_purge` | FamiTransfer, ManaGroup, ManaWorkset, Model Auditor | `transfer_families` runs on Lite's `family_transfer` (one target per call) |
| `schedule_roundtrip`, `link_models`, `cad_imports` | ManaSched, BatchLink, ManaDWG | |
| `open_tool_window` | — | MCP clients only: opens a T3Lab window for the person (`PostCommand`) |

Bridge (`lib/core/bridge.py`, Lite only, edited by hand): `t3lab_verify` checks exported files on
this computer; a call with a missing required argument, or a name one typo away from a real one, is
refused before it reaches Revit (`bad_arguments`, "did you mean"). `core/tool_cli.py` deploys next to
the bridge. Agents with a shell (Claude Code, Codex) list, read the help of, call and check these
tools from the command line, and Tool Hub recipes run over MCP (`t3lab_recipe`): `docs/cli.md`.

After a write, the server reads the model back and adds `verified` / `verification` to the result
(`core/write_verify.py`); `success` is never changed.

## What Lite's security settings do with them (MCP Control > Security)

- Read-only mode blocks every write tool above.
- A destructive call (delete worksets, purge, ungroup, overwrite families, delete CAD imports,
  unjoin, replace annotations) asks the person in Revit first, as for `delete_element`.
- Tested in `lite_guard/test_bridge_contract.py`.

## Not ported

Tiers / profiles / approval queue / turbo / gateway: they belong to t3lab_dev's bridge, which Lite
does not use. The bridge CLI and `t3lab_recipe` are ported on Lite's path (`docs/cli.md`). The BIM Harness
(`t3lab_harness`, `bridge.py harness …`) is ported on top of Lite's policy: `docs/harness.md`.

## Verify in Revit (restart Revit, do not Reload)

1. These windows had their logic moved to `lib/Snippets` / `lib/Services` and must work as before:
   BVBS Export, SheetGen, Room to Floor, Door Threshold, Wall Adjust Base, Wall Cut Profile,
   Point Cloud, CAD to Elements, AutoDimension, Batch Link, ManaDWG (also gets dev's fix for
   deleting pinned CAD imports), BatchOut ("Match Revit display" moved to `Services/sheet_export.py`).
2. One tool per row above from Claude Desktop (IronPython) and from the Assistant (CPython), with
   `dry_run` first.
3. `open_tool_window`: the command id of a pyRevit button, also inside a stack or pulldown.
