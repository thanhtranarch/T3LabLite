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

## [1.6.0] - 2026-10-09

Major UI modernization and ribbon regrouping matching t3lab-revit-api:
unified design system styles across all dialogs, six-panel ribbon layout,
new batch workset features, dynamic assistant greetings, and comprehensive stability fixes.

### Added
- **UI Design System Modernization**:
  - Full alignment with T3Lab modern UI standards across all WPF dialogs (`T3Standard`).
  - Standardized footer bar layout (`T3.FooterBar`): status text and copyright on the left, progress monitor and single clear primary action on the right.
  - Callout banners (`T3.Callout`) with context hints and shortcut tips across tools.
  - Status pills (`T3.StatusPill`) with string-bridged severity bindings for clean visual state transitions.
  - Segoe Fluent / MDL2 icons across navigation rails and buttons.
- **Dynamic Assistant Greetings** (`lib/GUI/AssistantGreetings.py`):
  - T3Lab Assistant welcome header now selects dynamic, non-repeating greetings based on time of day.
  - Dedicated English and Vietnamese greeting pools.
- **Batch Link Workset Management** (`lib/GUI/BatchLinkWorksets.py`):
  - Added Link Workset tab allowing per-link workset assignment in the grid.
  - Staged pending (amber), applied (green), and failed (red) status flow with bulk editor tools.
- **Thread Input Lock** (`lib/GUI/InputLock.py`):
  - Added `ThreadInputLock` utility to disable thread windows during modeless external event pumps, preventing errant clicks during long-running BatchOut operations.
- **Modular Point Cloud Analysis Service** (`lib/Services/point_cloud_analysis.py`):
  - Extracted point cloud geometry extraction and analysis pipeline to shared service for both UI and MCP tool workflows.

### Changed
- **Ribbon regrouped** -- restart Revit after updating (a pyRevit Reload cannot
  move buttons that already exist):
  - **Support**: Cloud Links, T3Lab Assistant, and a new **Settings** pulldown
    holding Feedback, MCP Control, LLMs Setting, ManaTabs, Ribbon Names,
    BG Theme and **Check Update**.
  - **Standards & Families** (was Standards & Settings): Standards stack
    (Model Auditor, ManaStyles, ManaWorkset), Managers stack (ManaLoca,
    ManaGroup, BatchLink), then ManaFami, Family Transfer and FamiGen.
  - **Model & Datum** (was Modeling & Datum): CAD to Elements and Point Cloud
    as large buttons, then a Tools stack of three pulldowns -- Finishes (Room
    To Floor, Door Threshold, Tile Layout), Reference (Image to Drafting,
    Property Line, Text to Element) and Modify (Auto Join, Split Elements,
    Wall Cut Profile, DatumSync, Auto Adjust Base Offset).
  - **Rebar & Assembly** moves next to Model & Datum.
  - **Annotation & Data** (Annotation & Select and Data merged): Mana stack,
    manaData stack, ManaAnno, Make Pattern, IFC-SG Suite.
  - **Views & Sheets**: View Tools stack (SheetGen, CropSync, PDF Import),
    ManaViews, ManaSheets, BatchOut.
- Ribbon icons follow the t3lab-revit-api icon set (28 buttons redrawn, a new
  Check Update icon).
- **Ribbon Names** keeps its maps in `%APPDATA%\T3LabAI\ribbon_names` instead of
  next to its script. Writing them inside the extension folder left a git
  install with local changes that blocked the next automatic update. Your
  existing names are carried over once.
- **Modernized Dialog Interfaces** -- Refactored layout, inputs, and controls across 40+ dialogs
  including AdvancedViewManager, AutoDimension, AutoJoin, BatchLink, BatchOut, CADToElements,
  DoorThreshold, FamiGen, ManaContains, ManaDWG, ManaFami, ManaGroup, ManaPara, ManaSched,
  ManaSelect, ManaSheets, ManaStyles, ManaViews, ModelAuditor, PDFImport, PointCloud, PropertyLine,
  RoomToFloor, SheetGen, SplitElements, T3LabAssistant, TextToElement, TileLayout, WallAdjustBase,
  and WallCutProfile.

### Fixed
- After a download-and-copy update (no git on the machine, or "Download latest
  version" in Check Update) the old panels no longer show up a second time:
  T3Lab moves the folders the update left behind to the `_retired_ribbon`
  folder of the extension on the next Revit start and says when a restart
  finishes the job (`lib/core/ribbon_retire.py`).
- **Check Update never updated**: its "Update now?" question always came back
  as "no" under CPython, so the click did nothing. After an update it now asks
  you to restart Revit instead of offering a pyRevit Reload -- on Revit 2025+ a
  Reload stops every T3Lab tool until Revit restarts.
- **Make Pattern**: Create Pattern failed every time.
- **ManaSelect**: the Quick Select, Select Similar, On Sheets and Warnings tiles
  snapped back to Explore; Quick Select now also fills its list when opened.
- **ManaSched**: importing Excel values back wrote lengths in feet (2500 mm
  became 2500 ft); values are read in the project's display units.
- **IFC-SG Suite**: after sorting a column, Apply to Selected wrote the subtype
  to other types.
- **Auto Dimension**: whole dimension strings were rejected with "Invalid
  number of references" when a grid or wall was slightly off axis.
- **Wall Cut Profile** and **Auto Adjust Base Offset** did not open at all;
  Pick then Apply now runs inside Revit's API context (the window closes while
  you pick and reopens with your inputs).
- **CAD to Elements**: the Level list and the wall / beam / MEP type lists were
  always empty, so every run stopped at "Select a Level."
- **Image to Drafting**: both tracing modes failed to load the tracer. Works on
  Revit 2022-2024; Revit 2025+ now says tracing is not available there yet.
- **Tile Layout**: Apply to Model created no tiles while reporting success, and
  Export CSV always failed.
- **Text to Element**: Pick items no longer picks while the dialog is still
  open (a known Revit crash pattern).
- **Point Cloud**: roofs were never created.
- **Room To Floor, Door Threshold, Point Cloud, Tile Layout, Wall Cut Profile,
  Auto Adjust Base Offset**: the second click in a session failed with
  "Duplicate type name within an assembly".
- **Split Elements** shows that splitting is not available yet instead of a
  file-not-found error, and Wall Cut Profile no longer offers "Edit Wall
  Profile", which did nothing.
- **ManaFami**: the Family Loader listed no families and Load loaded nothing;
  thumbnails now show, and Export List saves a `.csv`.
- **ManaStyles**: Duplicate did nothing for fill and line patterns; All / Clear
  / Custom now tick the boxes you see; Color Splasher says link sources need a
  View Filter instead of reporting 0 changes.
- **FamiGen**: Export & Place placed 0 instances.
- **Model Auditor**: Duplicate Elements detail rows and Select in Model work;
  run history is kept in `%APPDATA%\T3LabAI` instead of the extension folder
  (which also left git installs unable to update); status colours show.
- **BatchLink**: a link that fails to move to a workset no longer leaves its
  other instances half moved.
- **ManaGroup**: edited New Name cells turn amber.
- The Maximize button of ManaStyles and ManaFami toggled twice, so it did
  nothing.
- **ManaViews**: Excel export and Excel import failed every time.
- **PDF Import**: after unticking a view, All / None or switching mode, the
  PAGE column kept old numbers, so a view could get a different page than the
  one shown.
- **SheetGen**: Select All also ticked rooms hidden by the search, so Create
  made views and sheets for every room.
- **ManaViews / ManaSheets**: edited cells now turn amber before you apply.
- **ManaSheets**: Export always said it succeeded, even when it failed or fell
  back to CSV; Excel import works without openpyxl; import counted refused
  edits as "Updated".
- **BatchOut** opened from the docked T3Lab Assistant showed "Error loading
  sheets" and an empty window.
- **MCP Control**: the file watcher row showed an error and a disabled button
  while the watcher was running; the watcher started at Revit start is now
  reported as running.
- Missing imports that raised NameError: the file task watcher never started,
  MCP find_elements failed on name/level/type filters, View Manager Yes/No
  confirmations, IFC-SG subtype matching with more than 35 candidates, and the
  Assistant's task cards.
- The T3Lab Assistant's fallback for opening BatchOut looked for the script in
  the wrong folder. Image to Drafting (`potrace.exe`), the Assistant and the
  Assistant dock pane now find a button by its folder name, wherever it sits
  on the ribbon.

### Removed
- Redundant Cancel/Close buttons in dialog footer bars where the window title-bar close (X) is already standard.
- Obsolete local audit history JSON files from Model Auditor source tree.

## [1.5.0] - 2026-10-07

Every tool on the t3lab-revit-api ribbon is now on the T3Lab Lite ribbon too:
fourteen buttons that were missing, brought over with their current dialogs.

### Added
- **Rebar & Assembly** panel (new, the last panel on the tab), five tools for
  rebar and precast detailing:
  - **Cast Unit Manager** -- create assemblies with their rebar for many
    beams, columns or footings at once, sync loose rebar into its assembly,
    rename marks as a series and set rebar partitions by rule.
  - **Clone Drawing** -- copy the finished drawing of one assembly (views,
    sheet, annotations, tags, dimensions) to similar assemblies; anything that
    cannot be matched is listed with the reason.
  - **Rebar Check** -- data checks Revit does not run: rebar without a host,
    rebar missing from its assembly, assemblies without drawings, duplicate
    numbers, bars outside their host. Read-only until you press Fix.
  - **BVBS Export** -- write BVBS BF2D `.abs` files for bending machines;
    each file is read back and its checksums verified.
  - **Rebar Wizard** -- reinforce rectangular beams, columns and pad footings
    from a preset, added to the host's assembly.
- **Standards** stack on the Standards & Settings panel:
  - **ManaStyles** -- fill patterns, line styles, line patterns, Color
    Splasher and a coordinate editor.
  - **ManaWorkset** -- enable worksharing, create and delete worksets, assign
    elements to worksets by rule, generate workset view filters.
  - **ManaLoca** -- list elements of a view or level and edit their XYZ in a
    grid; stays open while you work.
- **IFC-SG Suite** on the Data panel: Subtype Assigner (Excel mapping to IFC
  Export Class and Predefined Type) and Compliance Checker (CORENET X rules).
- **ManaAnno** and **Make Pattern** on the Annotation & Select panel. ManaAnno
  finds, removes and renames Dimensions and Text Notes and edits dimension
  text; Make Pattern draws model and drafting hatch patterns on a vector canvas
  and creates them in Revit or exports `.pat`.
- **ManaFami** and **FamiGen** on the Modeling & Datum panel, next to Family
  Transfer. ManaFami batch-renames families and types and loads families;
  FamiGen creates families from CAD blocks, a JSON schema or built-in presets.
- **T3Lab Assistant** button on the Support panel. It runs the same assistant
  as the dock pane and the right-click menu.

### Changed
- **Restart Revit after updating.** The ribbon gained a panel and a stack, and
  a pyRevit Reload may not build them.
- **Automatic update now runs once a week instead of once a day, with no click
  needed.** On the first Revit start of each week (Monday to Sunday) T3Lab
  checks GitHub and installs the newest version in the background. A week only
  counts once GitHub answered, so an offline start is retried on the next day's
  first start. `"auto_update": false` still turns it off, and the new
  `"auto_update_interval": "daily"` keeps the old daily check. Check Update
  still updates on demand, and a manual check counts as that week's check.
  The schedule is in `lib/core/update_schedule.py`, tested by
  `lite_guard/test_update_schedule.py`.
- The restored tools use the latest t3lab-revit-api dialogs and windows
  (T3 design system), and the Revit 2022-2027 API helpers they were written
  against.
- **BG Theme** now keeps its colours in
  `%APPDATA%\T3LabAI\bg_theme\bg_theme_config.json` instead of a file inside
  the tool folder. An existing `dqt_bg_config.json` is copied over once and left
  where it is.
- Pause and Stop now also cover **FamiGen** and **IFC-SG Suite**; the Pause /
  Resume button shows an icon instead of an emoji.
- Shared code the new tools depend on was brought up to the t3lab-revit-api
  version, additions only: `GUI/WPF_Base.py` (rounded window and panel
  clipping, maximize), `GUI/ProgressPauseMixin.py`, `GUI/GridPendingEdits.py`,
  `Snippets/_compat.py` (Revit-version-safe rebar, assembly and family-parameter
  helpers), `Services/workset_service.py`, `core/paths.py` (`user_data_path`)
  and `core/extension_paths.py` (`find_bundle` / `bundle_path`: a button is
  found by its folder name, never by its panel path). `tab_path` was removed
  from `core/extension_paths.py`; the BG Theme service was its only caller.
- `lite_guard/manifest.json` records the new ribbon folders.

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
