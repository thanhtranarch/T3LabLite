# Changelog

All notable changes to the T3Lab Lite extension are recorded in this file,
newest release first. The **Check Update** tool reads this file from GitHub
and shows users the "What's new" list before they update.

Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/):
one `## [x.y.z] - YYYY-MM-DD` heading per release, with `### Added`,
`### Changed`, `### Fixed`, `### Removed` bullet lists underneath.

Releasing a new version:

1. Move the entries from `[Unreleased]` into a new `## [x.y.z] - YYYY-MM-DD`
   heading at the top.
2. Set the same version number in `version.txt`.
3. Commit and push -- Check Update on user machines will pick it up.

## [Unreleased]

## [2.0.0] - 2026-10-09

Major milestone release achieving complete feature, UI, and architectural parity with `t3lab-revit-api`. Consolidates the complete toolset into a streamlined 6-panel ribbon layout, introduces the all-new "T3 Space Line" vector icon standard, brings comprehensive pyRevit engine reload and startup protection, registers the native T3Lab Assistant dockable pane, and enhances stability across all dialogs—all while fully preserving T3Lab Lite's zero-click weekly auto-updater and privacy-first usage tracking.

### Added
- **Full Ribbon Parity with `t3lab-revit-api`**:
  - **Rebar & Assembly Panel**: Complete structural detailing suite containing Cast Unit Manager (assembly grouping & mark sequencing), Clone Drawing (intelligent assembly drawing propagation), Rebar Check (data consistency & host validation), BVBS Export (BF2D machine format generation with checksum verification), and Rebar Wizard (parametric beam, column, and footing reinforcement).
  - **Standards & Families Panel**: Dedicated BIM management suite including ManaStyles (pattern, line, and coordinate management), ManaWorkset (bulk workset rules and automated view filter generation), ManaLoca (real-time XYZ positioning grid), ManaGroup, BatchLink (with staged link-to-workset assignment), ManaFami (batch family/type renamer and loader), Family Transfer (cross-model family copying), FamiGen (generative family authoring), and Model Auditor.
  - **Annotation & Data Panel**: Integrated data suite merging Mana (AutoDimension, ManaDWG, ManaSelect), manaData (ManaContains, ManaPara, ManaSched with display unit sync), ManaAnno (text & dimension management), Make Pattern (vector hatch studio & `.pat` export), and IFC-SG Suite (CORENET X compliance & subtype mapping).
  - **Model & Datum Panel**: Full modeling suite with CAD to Elements, Point Cloud modeling, and Tools stack pulldowns for Finishes (Room To Floor, Door Threshold, Tile Layout), Reference (Image to Drafting, Property Line, Text to Element), and Modify (Auto Join, Split Elements, Wall Cut Profile, DatumSync, Auto Adjust Base Offset).
  - **Views & Sheets Panel**: Comprehensive documentation workflows with BatchOut (batch PDF/DWG/NWC exporter with live progress monitoring and export queue), ManaViews (bidirectional Excel schedule & view sync), ManaSheets, and ViewTools (SheetGen room-to-sheet generator, CropSync, PDF Import).
- **T3Lab Assistant Native Dockable Pane**:
  - Seamlessly registers as a native Revit `DockablePane` at startup using a lightweight IronPython host (`assistant_pane.py`), docking flush next to Revit's Properties pane.
  - Modeless GIL pumping ensures CPython worker threads and streaming LLM responses run smoothly without locking Revit.
- **Dynamic Assistant Greetings & Thread Input Lock**:
  - Non-repeating contextual greetings in both English and Vietnamese based on time of day (`lib/GUI/AssistantGreetings.py`).
  - `ThreadInputLock` utility (`lib/GUI/InputLock.py`) safely disables window interactions during long-running background operations.
- **MCP Server & Intelligence Modules**:
  - Self-healing MCP HTTP server (`mcp_ipy_host.py` / `Services/mcp_service.py`) hosted in IronPython to prevent GIL stalls, with automated bridge deployment to `%APPDATA%\T3LabAI\bridge.py`.
  - Zero-network file-based task watcher (`core/file_watcher.py`) watching `~/T3Lab_AI_Data/task.json`.
  - Advanced intelligence graph, planning, local assistant memory, and office text knowledge extraction.

### Changed
- **Redesigned Ribbon Icons (T3 Space Line Standard - Icon Standard 09)**:
  - All ribbon button icons across all 6 panels completely redesigned and audited against the T3 Space Line standard (52 vector icons, 0 errors, 0 warnings).
  - Features 2px stroke geometry, round caps and joins, integer-aligned paths, and adaptive light/dark palettes (`#000000`/`#F2F2F2` ink, `#666666`/`#A3A3A3` secondary, and `#EA680C`/`#FF8A3D` brand orange accent).
  - Cloud service links (Forma, Autodesk Health, Bluebeam) and T3Lab Assistant mascot icons preserved and locked.
- **Streamlined 6-Panel Ribbon Architecture**:
  - Reorganized panels into a logical workflow order: Support, Standards & Families, Model & Datum, Rebar & Assembly, Annotation & Data, Views & Sheets.
  - Added dedicated Settings pulldown under Support panel housing Feedback, MCP Control, LLMs Setting, ManaTabs, Ribbon Names, BG Theme, and Check Update.
- **Unified Modern UI Design System (`T3Standard`) Across 40+ Dialogs**:
  - Consistent window borders, modern headers, standardized footer bars (`T3.FooterBar`), status pills (`T3.StatusPill`), and callout banners (`T3.Callout`) implemented across all extension WPF windows.
- **Safe pyRevit & Revit Engine Startup**:
  - **Automatic Reload Patch** (`lib/pyrevit_patches.py`): Scans and automatically patches pyRevit `sessionmgr.py` on Revit 2025+ so pyRevit Reload operations no longer shut down the CPython engine ("This property must be set before runtime is initialized").
  - **PYREVIT_CPYVERSION Integer Sanitization** (`startup_fix_cpyversion.py`): Rewrites dotted version strings ("3.12.3" -> "3123") at startup to prevent `FormatException` Command Failure errors on pyRevit < 6.5.0.
  - **Startup Diagnostics & Mismatch Detection**: Validates Revit version compatibility (2022-2027), tests CPython engine DLL loads, and detects Python library vs runtime engine assembly mismatches with clear user instructions.
- **Weekly Auto-Update Preserved (Lite-Only)**:
  - Non-blocking weekly auto-update (`lib/core/updater.py`, `lib/core/update_schedule.py`) runs on first startup of the ISO week in a background thread; opt-out or daily frequency configurable via `%APPDATA%\T3LabAI\mcp_paths.json`.
  - On-demand updates available anytime via **Check Update** button.
  - Automatically retires legacy ribbon folders from copy-over updates to `_retired_ribbon` (`lib/core/ribbon_retire.py`).
- **Telemetry & Guard Rails Preserved (Lite-Only)**:
  - Automated session start, ribbon button click, and MCP tool call tracking (`lib/tracking/`) connecting to T3Lab Space.
  - Continuous regression prevention via `lite_guard` suite.
- **Decoupled User Configurations**:
  - Custom ribbon names, BG themes, and audit history moved out of the extension repository into `%APPDATA%\T3LabAI`, keeping the git working tree clean for seamless automated updates.

### Fixed
- **Check Update Execution**: Resolved issue where update prompts returned false under CPython; now safely requests Revit restart upon completion.
- **Units & Geometry Calculations**:
  - ManaSched: Fixed Excel re-import length conversions reading in project display units instead of raw feet.
  - AutoDimension: Fixed off-axis reference errors causing dimension strings to fail.
  - CAD to Elements: Fixed empty level and type dropdown lists.
  - Tile Layout: Fixed tile generation and CSV export errors.
  - Point Cloud: Fixed roof generation failures.
- **UI & State Bugs**:
  - Make Pattern: Resolved vector canvas hatch creation failures.
  - ManaSelect: Fixed selection category tiles snapping back to Explore and populated Quick Select lists.
  - IFC-SG Suite: Fixed subtype assignments corrupting upon column sort and resolved candidate overflow limits.
  - SheetGen: Fixed selection filtering including hidden search results during batch view/sheet creation.
  - ManaViews & ManaSheets: Fixed Excel I/O without requiring external openpyxl dependencies, and added amber highlight to modified cells.
  - ManaStyles & ManaFami: Fixed double-toggle behavior on window Maximize buttons.
  - BatchOut: Fixed empty sheets issue when opened from docked assistant context.
- **Engine & Import Stability**:
  - Fixed missing module imports causing `NameError` exceptions in file task watcher, MCP queries, View Manager, and Assistant task cards.
  - Resolved assembly duplicate type exceptions across repeated dialog launches in the same Revit session.

### Removed
- Removed unused 40 MB `mutool.exe` binary.
- Removed redundant Close/Cancel buttons in favor of standard title-bar controls.
- Purged obsolete local model audit history files from source control.

## [1.4.3] - 2026-10-02

### Fixed
- **Every CPython tool failing with "The input string '3.12.3' was not in a
  correct format"** (Command Failure for External Command). pyRevit builds
  before 6.5.0 seed `PYREVIT_CPYVERSION` as the dotted string `3.12.3`, then
  `int.Parse()` it on every `#! python3` click (pyRevit bug #3284).
  `startup.py` now rewrites it to the integer engine version (`3123`) first
  thing on every Revit start and pyRevit Reload, via the new
  `startup_fix_cpyversion.py` next to it. No-op on pyRevit 6.5.0+.
- **Anonymous usage statistics for ribbon tools and MCP calls.** 1.4.2
  (commit `5751bd2`) accidentally dropped these calls from
  `_cpython_bootstrap.init_cpython_paths()` and `core/server.py`, so only
  Revit start-ups were counted. Both are back; ribbon clicks are also found
  when pyRevit runs the script from a string (via the script's `__file__`),
  and IronPython falls back to .NET `WebClient` when `urllib2` HTTPS fails.
  Opt-out is unchanged: `"tracking_enabled": false` in settings.

## [1.4.2] - 2026-09-29

### Added
- **Family Transfer (`FamiTransfer`)**: New tool on the Modeling & Datum panel to copy Revit families and selected types between open projects and loaded link models. Features quick search, category filtering, source/target document selectors, conflict resolution (overwrite, rename, skip), and single-transaction undo.
- **Make Pattern**: Vector hatch studio (`patmaker`) for interactive authoring and export of Revit fill pattern (`.pat`) files with tiling controls and real-time preview.
- **Anonymous Telemetry**: Privacy-preserving, non-blocking asynchronous telemetry service (`telemetry_service.py`) for tracking session start, ribbon tool clicks, and MCP invocations with machine/user hash anonymization and opt-out support via settings.
- **Panel Theme Styling**: Added consistent background accent color (`#2CE07B00`) across all six T3Lab ribbon panel bundle configurations.

### Changed
- **Revit 2025+ Compatibility & Stability**: Added `pyrevit_patches.py` for safe pyRevit reload handling, .NET disposal wrappers, and enhanced API compatibility for Revit 2025/2026.
- **MCP Bridge Health**: Added health checks and auto-recovery for the MCP server bridge across Revit session restarts.
- **Dialog Error Handling**: Enhanced `ErrorGuard` protection and UI reliability across CropSync, DatumSync, and management dialogs.

## [1.4.1] - 2026-09-27

### Added
- **Check Update**: Restored the interactive Check Update tool on the Support panel and background daily update check in `startup.py`. Supports git fast-forward (`git pull --ff-only`) and direct zip fallback, changelog "What's new" preview, and pyRevit reload prompt.
- **CropSync**: New tool in Views & Sheets panel to synchronize crop regions, annotation crops, and crop view settings across selected views.
- **DatumSync**: New tool in Modeling & Datum panel to synchronize grid lines, levels, and datums across views.

### Changed
- **CPython 3 Migration**: Migrated scripts across all tools to CPython 3 runtime (`#! python3`) with robust engine path bootstrapping via `_cpython_bootstrap.py`.
- **Panel Renaming**: Renamed `Data & IFC-SG` panel to `Data`.
- **Ribbon Layout Refinement**: Streamlined button layouts across `Data.panel`, `Standards & Settings.panel`, `Modeling & Datum.panel`, `Annotation & Select.panel`, and `Support.panel`.

### Removed
- Removed deprecated/standalone tools from ribbon:
  - **IFC-SG Suite, BCF Reader, Foundation Volume** (from Data panel).
  - **ManaLoca, ManaStyles, ManaWorkset** (from Standards & Settings panel; functionality consolidated into Model Auditor).
  - **ManaFami, FamiGen** (from Modeling & Datum panel).
  - **ManaAnno** (from Annotation & Select panel).
  - Standalone **T3LabAssistant** ribbon pushbutton (now accessible natively via Dockable Pane and right-click context menu in Revit 2025+).

## [1.4.0] - 2026-09-15

### Added
- **ModelAuditor**: Smart Purge and Advanced Purge modules for comprehensive model health cleanup, scanning unreferenced views, unused families, dangerous operations, and worksets.
- **ManaGroup & BatchLink**: New tools for managing Revit/CAD link instances and grouping operations.
- **ManaWorkset, ManaLoca, ManaStyles**: Dedicated management tools for worksets, location coordinates, and object styles.
- **AI Teaching & Learning**: Trajectory capture and exemplar building for training custom assistant behaviors and workflows.
- **UI Themes**: Enhanced Revit Light/Dark mode compatibility and ribbon UI theme customization.
- **Tile Layout & AutoJoin**: Automated element joins and tile layout generation.
- **Sheet Manager & BatchOut**: Improved view placement, custom parameters service, and batch export workflow.

## [1.3.1] - 2026-08-05

### Fixed
- **Ribbon build error after updating to 1.3.0** (`UI build error for
  'T3LabLite': ... exists:Feedback`). 1.3.0 moved the **Feedback** and **MCP
  Control** buttons into the new *Assistant Tools* stack, but Revit keeps
  every ribbon item name for the life of the session -- so rebuilding the
  panel after the in-place update hit the names the old standalone buttons
  had already taken, and the whole T3Lab tab failed to load. The stacked
  buttons now use fresh bundle names (`SendFeedback`, `MCPPanel`); their
  ribbon labels are unchanged.
- **Check Update** now says that ribbon layout changes only take effect after
  a Revit restart, so a reload that leaves the tab looking incomplete is not
  mistaken for a broken update.

## [1.3.0] - 2026-08-05

### Added
- **T3Lab Assistant**: the AI assistant is back, rebuilt around real tool
  calling. It opens from the Support panel, docks as a native Revit pane next
  to Properties / Project Browser, and also sits in the right-click menu
  (Revit 2025+). Chat in Vietnamese or English to ask about the model or
  change it -- the assistant runs Revit tools and opens T3Lab tools for you,
  streams each step, and by default presents a plan and waits for your
  confirmation before it changes the model (deletes always ask; the wait can be
  turned off in LLMs Setting).
  - **Skills** -- reusable markdown instruction packs that activate on what you
    ask. 25 ship built-in (ISO 19650 naming, LOD, worksets, QA checklist, BEP,
    COBie handover, clash coordination, sheet/annotation standards and more).
    Type `/skill-name` to force one, or install extra skills straight from a
    GitHub repo -- Claude's `SKILL.md` format is read as-is.
  - **Knowledge (RAG)** -- index folders of PDF / TXT / MD and get answers with
    citations. Keyword search (BM25) always works offline; semantic search is
    optional and runs on a local Ollama embedding model.
  - **Projects** -- each project keeps its own chat history, custom
    instructions, default provider/model, knowledge folders, remembered facts
    ("remember ..." in chat) and daily scheduled prompts.
  - **Context and attachments** -- the assistant sees the active view and the
    current selection, and accepts attached files and images.
  - **PDF comment resolution** -- reads markup annotations from a PDF, traces
    them to the matching Revit sheet, and proposes a fix per comment that you
    run item by item.
  - **Spell check** -- proofreads every Text Note in the model or just the
    active view.
  - **Specialist routing** -- requests go to a focused agent (read-only data,
    model actions, modeling, QA, export, multi-document, knowledge, comments);
    multi-goal requests are planned as a graph, executed in parallel and
    re-routed when a step fails.
- **LLMs Setting** (Support ▸ Assistant Tools): one settings hub for every
  AI-powered T3Lab tool -- provider (Claude, OpenAI, DeepSeek, Ollama,
  LM Studio), model, API key or local server URL, live connection status per
  provider and a one-click test message; plus display name, "ask before model
  edits", deep-reasoning and maximum-quality toggles, chat detail, and the
  Projects / Knowledge / Skills tabs.
- **Assistant Tools** stack on the Support panel groups Feedback, MCP Control
  and LLMs Setting.
- **Pause / Stop** while a tool is running: Model Auditor, ManaSheets,
  ManaViews, SheetGen, FamiGen, IFC-SG, Tile Layout, Auto Join, Room To Floor,
  Image to Drafting, CAD to Elements (Wall / Floor / Beam) and Bulk Family
  Export now show a progress bar with Pause and Stop, from the shared
  `ProgressPauseMixin`.
- **Auto-update**: the first Revit start of each day checks GitHub for a newer
  release and downloads it in the background. The new version becomes active on
  the next Revit start (or pyRevit reload); a toast says so. Opt out with
  `"auto_update": false` in `%APPDATA%\T3LabAI\mcp_paths.json`.
- **MCP**: 12 new tools -- `manage_view`, `manage_sheet`,
  `manage_view_template`, `manage_document`, `manage_links`, `manage_material`,
  `manage_revision`, `edit_elements`, `create_detail_annotation`,
  `export_model`, `check_bad_geometry`, `collect_spellcheck_text`.

### Changed
- **Check Update**: version lookup, changelog reading and the git/zip update
  strategies moved into `lib/core/updater.py`, shared with the daily automatic
  check. The button behaves as before.
- **Assistant surface** follows Revit's Light/Dark UI theme (Revit 2024+)
  instead of a fixed light window.
- **MCP server / tool layer**: destructive calls are now declared as such so
  the Assistant always asks first; the active document is resolved even when
  Revit has no focused tab (start page, freshly opened session) instead of
  every tool failing; and navigating to an element activates its owner view,
  so view-specific elements (tags, dimensions, text notes) no longer trigger
  Revit's "No good view could be found" dialog.

## [1.2.0] - 2026-07-17

### Added
- `CHANGELOG.md` to track what each release changes.
- **Check Update**: shows a "What's new" summary from this changelog when a
  newer version is found.
- **BG Theme 2.0**: rebuilt as a full theme studio -- HSV colour picker
  (SV square + hue bar) with screen eyedropper, live-apply, named custom
  presets and recent colours for the model background; Sky/Horizon/Ground
  gradient backgrounds for 3D views (active view or all 3D views); Light/Dark
  Revit UI theme switching (Revit 2024+). SHIFT+Click still cycles
  Black → Gray → White.
- **PointCloud**: "Use View Crop" takes the extraction region straight from
  the active view's crop box (safest for large clouds), and custom regions
  are now defined by dragging a rectangle.
- **SheetGen**: title-block header strip setting (right / bottom / none,
  size in mm) so viewports avoid the title block frame.

### Changed
- **PointCloud**: detection engine reworked -- points are extracted in tiles
  with a density cap and progress reporting; walls are swept along project
  grid directions (not just 0°/90°) with checks that reject furniture,
  shelving and low MEP runs; floors and ceilings are detected from horizontal
  surfaces; doors and windows are hosted into their detected wall with the
  insertion point at the threshold / sill.
- **SheetGen**: interior elevations are created with the marker type that
  actually hosts the most views and named from their real view direction;
  the sheet layout preview now uses the true title block paper size and the
  same placement maths as the real viewports, so preview == result.

## [1.1.1] - 2026-07-15

### Added
- **CAD To Elements**: MEP routing support -- create Ducts, Pipes, Cable
  Trays and Conduits directly from CAD lines, with new MEP options in the
  dialog.

### Changed
- **Parameter Selector** dialog improvements.

## [1.0.1] - 2026-07-15

### Changed
- **Tile Layout**: engine refactored into a shared core library for more
  reliable layout generation.
- **BatchOut**: window is now modeless -- you can keep working in Revit
  while it stays open.

### Removed
- Legacy T3Lab Assistant window and unused test files.

## [1.0.0] - 2026-07-08

First tracked release.

### Added
- **Check Update** tool (Support panel): compares the installed version with
  the latest release on GitHub and updates via git or direct download.
- **MCP Control**: multi-instance bridge -- connect to more than one open
  Revit session.

### Removed
- T3Lab Assistant pane (replaced by the MCP-based workflow).
- Unused pushbutton icon assets.
