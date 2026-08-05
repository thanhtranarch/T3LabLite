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

## [1.3.0] - 2026-08-05

> **Restart Revit to install this one — do not use pyRevit ▸ Reload.**
> Feedback and MCP Control move into the new Assistant Tools stack, and Revit
> will not let a ribbon button be moved once a session has created it. Reloading
> on top of a running 1.2.0 fails the whole T3Lab UI build with
> `...exists:Feedback` (`RibbonPanel.verifyNameExclusive`); restarting Revit
> clears it.

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
