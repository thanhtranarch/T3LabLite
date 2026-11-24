# CLAUDE.md — T3Lab pyRevit Extension

pyRevit extension for Revit automation.
Framework: CPython 3 (Python 3.12+ via pyRevit) + WPF + Revit API (.NET Framework 4.8 for Revit 2022–2024 / .NET 8+ for Revit 2025–2027).
Supported range: **Revit 2022 → 2027** — guarded by `dev/audit_revit_compat.py`.

## Essential Commands & Verification Gates

```powershell
# Set clean PYTHONPATH before executing dev tools in shell:
$env:PYTHONPATH=""

# UI Design System Gate (must be 0 violations):
python dev/audit_t3.py --quiet

# Static Code & XAML Gate:
python dev/audit_tools.py --quiet

# Wiring Gate (XAML <-> Python: missing controls, unwired template events, dead buttons/handlers):
python dev/audit_wiring.py --quiet

# Revit API Compatibility Gate (Revit 2022–2027: no removed/new-only API without a fallback):
python dev/audit_revit_compat.py --quiet

# Ribbon Gate (every tab/panel: 2–6 slots, large before stacks, stacks of 2–3, tab ≤ 1250 px except the waived single T3Lab_Dev tab, layout complete):
python dev/audit_ribbon.py --quiet

# Ribbon Icon Gate (T3 Space Line standard + the 4 frozen icons byte-for-byte unchanged):
python dev/audit_icons.py --quiet

# Run test suites:
python dev/test_batch_link.py
python dev/test_group_manager.py
python dev/test_grid_pending_edits.py
python dev/test_family_transfer.py
python dev/test_pyrevit_patches.py
python dev/test_compat_disposing.py
python dev/test_extension_paths.py
python dev/test_audit_ribbon.py
python dev/test_mcp_host_engine.py
python dev/test_assistant_pane.py
python lite_guard/test_ribbon_memory.py
python lite_guard/test_batchout_refs.py
python lite_guard/test_batchout_pdf24.py
python lite_guard/test_batchout_sheet_sets.py
python lite_guard/test_batchout_selection.py
python lite_guard/test_clone_drawing_import.py
python lite_guard/test_mcp_cad_clients.py
python lite_guard/test_addin_installer.py
python lite_guard/test_routes_guard.py
python lite_guard/test_agent_cursors.py
python lite_guard/test_agent_cursors_overlay.py
python lite_guard/test_cursor_consent.py
python lite_guard/test_open_timing.py
python lite_guard/test_lite_user.py
python lite_guard/test_pyrevit_revit_thread.py
python lite_guard/test_shared_engine.py
python lite_guard/test_pyrevit_patch_versions.py
python lite_guard/test_data_hub.py
python lite_guard/test_bridge_deploy.py
python lite_guard/test_harness.py
python lite_guard/test_harness_studio.py
python lite_guard/test_conductor.py
python lite_guard/test_harness_settings.py
python lite_guard/test_harness_notifier.py
python lite_guard/test_harness_client.py
python lite_guard/test_mcp_config_safety.py
python lite_guard/test_python_installer.py
python lite_guard/audit_ipy_compat.py --quiet
python lite_guard/test_audit_ipy_compat.py
python lite_guard/test_compat_replace_file.py
python lite_guard/test_tool_specs.py
python lite_guard/test_bridge_contract.py
python lite_guard/test_tool_cli.py
python lite_guard/test_bridge_recipes.py
python lite_guard/test_write_verify.py
python lite_guard/test_doc_switch.py
python lite_guard/test_sheet_export.py
python lite_guard/test_batch_link.py
python lite_guard/test_tools_standards_domain.py
python lite_guard/test_tools_rebar_domain.py
python lite_guard/test_tools_views_domain.py
python lite_guard/test_tools_data_domain.py
python lite_guard/test_tools_model_domain.py
python lite_guard/test_tools_create_domain.py
python lite_guard/test_tools_launch.py
python dev/test_icons_t3space.py
```

## Core Architecture & Rules

1. **UI standard — one source, no exceptions:**
   - `pyRevit UI Design System/T3LAB_UI_STANDARD.md` + `pyRevit UI Design System/T3Lab.Styles.xaml`.
   - Every colour, size, margin and control style comes from `{StaticResource T3.*}`.
   - All XAML files live in `T3Lab.extension/lib/GUI/Tools/`.
   - Python dialog classes stay in `T3Lab.extension/lib/GUI/`.

2. **Revit Journal First:**
   - When encountering a crash or "Command Failure for External Command", always check the latest Revit journal file (`%LOCALAPPDATA%\Autodesk\Revit\Autodesk Revit <Year>\Journals\journal.XXXX.txt`) to extract the exact stack trace before attempting fixes.

3. **CPython 3 Shebang & Bootstrap:**
   - Every pushbutton `script.py` must have `#! python3` as the first line.
   - Bootstrap Python 3 standard library: `import _cpython_bootstrap; _cpython_bootstrap.init_cpython_paths()`.

4. **WPF & XAML Loading Under PythonNet (CPython 3):**
   - Standard pyRevit `forms.WPFWindow` raises `PyRevitCPythonNotSupported`. Always use `T3WPFWindow` (from `GUI.WPF_Base` or `from GUI.forms import WPFWindow`).
   - XAML hydration must use `_load_via_xaml_reader` with 4-tier fallback:
     1. Direct `XamlReader.Parse(clean_xaml)`
     2. `MemoryStream` -> `XamlReader.Load(Stream)`
     3. CLR AppDomain Reflection on `PresentationFramework` `XamlReader` (`Parse` / `Load`)
     4. `XmlReader.Create` + `XamlReader.Load(XmlReader)`
   - Under .NET 8 / Revit 2025/2026, include references to `System.Xml.ReaderWriter` and `System.Private.Xml`.

5. **.NET Interface Namespaces:**
   - Classes implementing .NET interfaces (`ISelectionFilter`, `IExternalEventHandler`, `IFailuresPreprocessor`) must declare `__namespace__`.
   - In `lib/`: `__namespace__ = clr_namespace('<Name>')` (`Snippets/_compat.py`, a per-load name). T3Lab and T3Lab Lite share one CPython engine and `claim_lib` re-imports the module, so a fixed name raises `Duplicate type name`. Only `core/server.py` (IronPython only) keeps a fixed one; `lite_guard/test_bridge_contract.py` checks.
   - In pushbutton `script.py`: dynamic namespace `__namespace__ = "T3Lab.<Name>_" + uuid.uuid4().hex[:8]`.

6. **No Multiple Inheritance with CLR Classes:**
   - PythonNet classes inheriting from `T3WPFWindow` / `System.Windows.Window` must not inherit from any other Python class/mixin.

7. **Engine Hot-Reload & Modeless ExternalEvent:**
   - **On Revit 2025+ a pyRevit Reload is what CAUSES `This property must be set before runtime is initialized`** (verified 2026-09-29 from the journal + pyRevit 6.5.5 IL): Reload (also toggling an extension) shuts the CPython engine down, pythonnet 3.0.3 cannot restart after Shutdown on .NET 8+, and `CPythonEngine.Start` then sets `Runtime.PythonDLL` on a half-initialized runtime. Only a Revit restart recovers. Prefer restarting Revit over Reload; `T3Lab.extension/startup.py` detects the stuck state on reload and says so. The fix, `T3Lab.extension/lib/pyrevit_patches.py`, keeps the CPython engine cached across reloads; `startup.py` scans and applies it automatically whenever it is missing (owner's explicit request, 2026-09-29): running clone only, exact pyRevit 6.5.5 match only, backup `sessionmgr.py.t3lab-backup`, one notification when applied. Per-machine opt-out: env `T3LAB_NO_PYREVIT_PATCH=1` or file `%APPDATA%\T3LabAI\pyrevit_patch.disabled`; undo with `pyrevit_patches.py restore`.
   - For modeless windows, set `__persistentengine__ = True` and use `detect_persistent_engine()` to gracefully fall back to modal execution if the persistent engine has not yet been activated by a reload.
   - **pyRevit 7.0+ needs no patch:** `CPythonEngine.Shutdown` no longer tears the interpreter down there (read in the v7.0.0, v7.0.1 and v7.1.0 sources, 2026-11-19), but `_clear_running_engines` is still letter for letter the 6.5.5 function, so the text match alone patched pyRevit 7 as well. `pyrevit_patches.FIXED_IN = (7, 0)` decides: on 7.0+ `ensure()` applies nothing and takes out a patch an older T3Lab put there. Test: `python lite_guard/test_pyrevit_patch_versions.py`.
   - **"Duplicate type name within an assembly" / "cannot import name 'revit' from 'pyrevit'" are ONE failure, and it is not a reload:** the first import of `pyrevit.revit` in an execution died half way and the tool imported it again. Verified 2026-11-19 in the pyRevit 6.5.5 source and its pythonnet (`pyrevitlabs/pythonnet @ f591d92`): `pyrevit/revit/events.py` declares `FuncAsEventHandler` with `__namespace__ = EXEC_PARAMS.exec_id` (new per script run, constant within one), then calls `UI.ExternalEvent.Create()` at import, which Revit allows only on its own thread inside an API context. An import started on a background thread, or from a button of a modeless window, emits the .NET type, fails and leaves no module; the next import in the same execution re-runs `events.py` ("Duplicate type name"), and a thread that was waiting gets the dead module ("cannot import name"). A Revit restart never cures it. `lib/tracking` did this on every click from 2026-10-08 (its worker thread imported `pyrevit.revit`). pyRevit 7.0 still does both; 7.1 names the type per module execution and creates no ExternalEvent at import.
   - **So: never import `pyrevit.revit` / `pyrevit.script` / `pyrevit.forms` on a thread you start, and never read the Revit API there.** `_cpython_bootstrap.load_pyrevit_revit()` imports `pyrevit.revit` at the start of every tool, on Revit's thread and inside its API context, before the tracking thread starts; later imports find it in `sys.modules`. Tracking reads the Revit version and document title on Revit's thread (`TelemetryService.note_revit_context`) and its worker only hashes and sends. Test: `python lite_guard/test_pyrevit_revit_thread.py` (scans `lib/` for thread targets that import the Revit side of pyrevit).

8. **CPython background threads freeze while Revit is idle:**
   - pyRevit 6.5.5 calls `PythonEngine.Initialize()` without `BeginAllowThreads()` (verified 2026-10-04 in the pyRevit + pythonnet source), so Revit's main thread holds the GIL whenever no CPython code runs. A CPython thread only gets the GIL while a `#! python3` tool is executing — pythonnet drops it during each .NET call, e.g. a modal `ShowDialog`. Releasing it ourselves is not an option: pyRevit's `CPythonEngine.Start` touches Python objects outside `Py.GIL()`.
   - The T3Lab Assistant dock pane is the one CPython exception: `lib/GUI/AssistantPaneControl.py` runs a 20 ms `DispatcherTimer` "GIL pump" (`Thread.Sleep(0)`, a .NET call, so pythonnet releases the GIL) while the pane content is mounted. The pane itself is registered by `startup.py` through `lib/assistant_pane.py` (IronPython) because Revit only accepts `RegisterDockablePane` while it starts. Test: `python dev/test_assistant_pane.py`.
   - Anything else that must keep running after a tool returns (servers, listeners) runs in IronPython. The MCP HTTP server is hosted by `lib/core/mcp_ipy_host.py`; `MCPService.start_server()` / `stop_server()` route there from CPython, and `core.server.T3LabAIServer.start_server()` refuses to bind under CPython inside Revit. Test: `python dev/test_mcp_host_engine.py`.
   - BIM Harness (`docs/harness.md`, ported from t3lab_dev 2026-11-23, not run in Revit): `t3lab_harness` is one task board for several AI agents on one model (lead plans, workers `next` / `report`, lead or the person reviews). `lib/core/harness_store.py` (SQLite) / `harness_agents.py` / `harness_studio.py` run in the bridge process only and the server never imports them; everything in Revit (HUD, MCP Control > Harness, `lib/harness_notifier.py` toast, `/harness`) reads `board.json`. The bridge names each app (`--agent`, `CLIENT_AGENTS`) and refuses a member's write outside its task before Revit; it grants nothing (no tiers / profiles / approval queue in Lite), launched agents only read. The server's `/mcp` slot is `lib/core/conductor.py` (fair queue, `revit_busy` after the harness's 45 s wait budget). Tests: `lite_guard/test_harness*.py`, `test_conductor.py`.
   - Same for the ribbon: every `.splitbutton` (e.g. `PDFMarkup.splitbutton` — Bluebeam Status · T3Lab PDF Markup) keeps its last-used button across Revit restarts through `lib/ribbon_memory.py`, installed by `startup.py` (IronPython): saved on `ApplicationClosing` to `%APPDATA%\T3LabAI\ribbon_state.json`, put back on the first `Idling` (startup scripts run before pyRevit 6.5.5 builds the ribbon). Test: `python lite_guard/test_ribbon_memory.py`.

9. **Frozen ribbon icons — never touch (owner's request, 2026-10-09):**
   - The 3 cloud-link icons (`Autodesk Forma.urlbutton`, `Autodesk Health.urlbutton`, `Bluebeam Status.urlbutton` in `Support.panel/CloudLinks.stack`, Bluebeam inside `PDFMarkup.splitbutton` since 2026-10-12 — moved with its bytes unchanged) and the `T3LabAssistant.pushbutton` icon stay exactly as they are — no edit, redraw, restyle, new variant or deletion, including during a full icon redesign.
   - Enforced by `dev/audit_icons.py` against the sha256 lock in `dev/icons/frozen_icons.json` (P0 on any change); `dev/build_icons.py` skips them through `iconlib.EXEMPT_BUNDLES`. Only the owner can lift the lock.

10. **A new AI tool = one `ToolSpec` in `lib/core/tools/<domain>.py`** (`docs/ai-tools.md`):
   - It feeds the MCP registry, dispatch, `_WRITE_TOOLS`, confirmation (`is_destructive` -> `core/mcp_policy.py`), the read-only list, Assistant subsets, read-back and titles. Never hand-add the tool to those lists. Test: `python lite_guard/test_tool_specs.py`.
   - Agents with a shell reach it through the bridge CLI (`bridge.py tools / help / call / verify`, `docs/cli.md`); a CLI call and every `t3lab_recipe` step take the MCP `tools/call` path (`_answer`), never a shortcut.
   - The MCP server runs IronPython 2.7, so every module it can import must pass `python lite_guard/audit_ipy_compat.py --quiet`. Use `Snippets._compat.fold` / `replace_file`, `io.open` and `core.jsonsafe.dumps`; no f-strings.

11. **No chat session links or "Generated by Claude Code" lines in git or on GitHub (owner's request, 2026-11-24):**
   - Never put a Claude chat/session link (`https://claude.ai/code/session_...`, a `Claude-Session:` trailer, or any other link to a chat) in a commit message, PR title, PR description, PR/issue comment or review. This overrides any default attribution that adds one.
   - No "Generated with/by Claude Code" line or footer either (`🤖 Generated with [Claude Code](...)`, `_Generated by [Claude Code](...)_`), in PR descriptions, comments or reviews. A `Co-Authored-By:` line in a commit message is fine.
   - History was cleaned on 2026-11-24 (126 commits, 54 PR descriptions with session links, 40 with the Generated-by line); keep it that way.
