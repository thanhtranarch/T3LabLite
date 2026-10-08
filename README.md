# T3Lab Lite

**T3Lab Lite** is a pyRevit extension running on **CPython 3** built for Revit users who want to work faster.
It covers batch export, sheet & view management, datum & crop synchronization, CAD-to-BIM conversion, model auditing, rebar and assembly detailing (cast units, drawing cloning, BVBS export), IFC-SG compliance, family generation and management, a built-in AI assistant that runs Revit tools from plain Vietnamese or English, and MCP integration that lets Claude AI work with Revit directly.

See [CHANGELOG.md](CHANGELOG.md) for what's new in each release.

---

#### Installation

T3Lab Lite is installed as a pyRevit extension.

    ▶ Install pyRevit

    ▶ Open Extensions Menu

    ▶ Add T3Lab Lite

    ▶ Click Install

#### Staying up to date

You never have to click anything. On the first Revit start of each week
(Monday to Sunday), T3Lab checks GitHub for a newer release and installs it in
the background. Revit keeps running the version it loaded, so the update takes
effect the next time you start Revit — a notification tells you when
that is worth doing. Restart Revit rather than using pyRevit ▸ Reload: on
Revit 2025+ a Reload stops every T3Lab tool until Revit restarts. If GitHub cannot be reached
(offline, blocked), the next day's first start tries again until the check
gets an answer, so a bad Monday does not cost you the week.

A clone with local commits or edits is only ever fast-forwarded, never
overwritten. Two settings in `%APPDATA%\T3LabAI\mcp_paths.json`:
`"auto_update": false` turns the automatic check off, and
`"auto_update_interval": "daily"` checks every day instead of every week.
**Check Update** (Support ▸ Settings) updates on demand either way, and
`%APPDATA%\T3LabAI\update.log` records what happened.

---

### Tools

The ribbon has the same six panels as t3lab-revit-api: **Support**, **Standards & Families**, **Model & Datum**, **Rebar & Assembly**, **Annotation & Data** and **Views & Sheets**. All tools run on the high-performance **CPython 3** engine.

Tools that process many elements — Model Auditor, ManaSheets, ManaViews, SheetGen, FamiGen, IFC-SG Suite, Tile Layout, Auto Join, Room To Floor, Image to Drafting, CAD to Elements and BatchOut — show a progress bar with **Pause** and **Stop** while they run, so a long batch can be paused or cancelled without killing Revit.

---

### Support

#### T3Lab Assistant
The AI assistant, in Revit. Ask about the model or tell it what to change, in Vietnamese or English — it calls real Revit tools, opens T3Lab tools for you, and by default presents a plan and waits for your confirmation before changing the model (deletes always ask). It opens from the **T3Lab Assistant** button on this panel as a window, docks as a native Revit pane beside Properties / Project Browser, and is also in the right-click menu (Revit 2025+).

- **Skills** — reusable instruction packs that activate on what you ask, 25 built in (ISO 19650 naming, LOD, worksets, QA checklist, BEP, COBie handover, clash coordination, sheet & annotation standards, …). Type `/skill-name` to force one, or install more from a GitHub repo — Claude's `SKILL.md` format is read as-is.
- **Knowledge (RAG)** — index folders of PDF / TXT / MD and get answers with citations. Keyword search works offline; semantic search is optional and runs on a local Ollama embedding model.
- **Projects** — per-project chat history, custom instructions, default model, knowledge folders, remembered facts, and daily scheduled prompts.
- **Context & attachments** — the assistant sees the active view and current selection, and accepts attached files and images.
- **PDF comments** — read markups from a PDF, trace them to the matching sheet, and resolve them item by item.
- **Spell check** — proofread every Text Note in the model or just the active view.

#### Settings (pulldown)
- **MCP Control** — start and stop the local MCP server that lets Claude AI (and other MCP clients) interact directly with Revit. Configure host, port, and authentication settings.
- **LLMs Setting** — the settings hub shared by every AI-powered T3Lab tool: provider (Claude, OpenAI, DeepSeek, Ollama, LM Studio), model, API key or local server URL, live connection status per provider and a test message; plus your display name, "ask before model edits", deep reasoning / maximum quality, chat detail, and the Projects, Knowledge and Skills tabs.
- **Feedback** — send feedback or suggestions to the T3Lab team directly from Revit.
- **ManaTabs** — hide or show Revit ribbon tabs to reduce clutter.
- **Ribbon Names** — shorten or restore ribbon tab names with inline editing and saved mappings.
- **BG Theme** — set the model-view background colour. Presets, RGB sliders, HEX input, and live preview. SHIFT+Click cycles Black → Gray → White.
- **Check Update** — compare the installed version (`version.txt`) with the latest release on GitHub and update via git or direct download — see [Staying up to date](#staying-up-to-date).
  - **Version Detection** — queries GitHub repository releases and mirrors (jsDelivr) to prevent rate limits.
  - **What's New Preview** — automatically extracts release notes from `CHANGELOG.md` between local and remote versions.
  - **Safe Updating** — pulls updates via `git pull --ff-only` when running from a git clone (protecting local edits) or downloads and extracts the release archive.
  - **Restart prompt** — tells you to restart Revit so the updated tools load.

#### Cloud Links
Quick links to Autodesk Forma, Autodesk Health dashboard, and Bluebeam Status.

---

### Standards & Families

#### Standards (Model Auditor · ManaStyles · ManaWorkset)
- **Model Auditor** — consolidated model health checks in one window:
  - **Model Check** — verify model standards and quality rules.
  - **Smart Purge** — safe, category-based model cleanup removing unreferenced views, unused families, and unplaced elements.
  - **Warnings** — review and address the Revit warning list.
  - **In-Place Models** — list and manage in-place family instances.
  - **Material List** — audit all materials used in the model.
- **ManaStyles** — manage fill patterns, line styles and line patterns, apply graphic override colours by category rule (Color Splasher), and view and adjust element XYZ coordinates in a grid.
- **ManaWorkset** — enable worksharing on a model, create and delete worksets, assign elements to worksets by rule (category, level, or type), and generate view filters from workset membership.

#### Managers (ManaLoca · ManaGroup · BatchLink)
- **ManaLoca** — list the elements of the active view or a level and edit their XYZ coordinates in a data grid, committed in one transaction. The window stays open while you work.
- **ManaGroup** — manage Revit Model Groups and Detail Groups: list group types and placed instances, count references, and audit unused group definitions.
- **BatchLink** — manage Revit and CAD link paths in bulk: verify link statuses, repath missing links, reload links across documents, and audit external dependencies.

#### ManaFami
Family manager. Batch Operations lists families, system types, model groups or assemblies; find and replace, add a prefix or suffix and change case, review every staged change in the grid, then apply it in one undoable step. Family Loader loads new families from a local folder (or a cloud catalogue you configure) into the project.

#### Family Transfer
Transfer families and selected types between open project documents or loaded links with category filtering, conflict resolution (overwrite, rename, skip), and single-transaction undo.

#### FamiGen
Create Revit families from external data: from CAD (scan imported DWG blocks and export each unique block as an `.rfa`), from a JSON schema (fully parametric families), or in batch from built-in presets. Can draft the schema with your configured AI provider.

---

### Model & Datum

#### CAD to Elements
Convert CAD linework into Walls, Floors, or Beams by layer and colour mapping.

#### Point Cloud to Model
Scan-to-BIM wizard that auto-detects Walls, Floors, Ceilings, Doors, Windows, Columns, Stairs, and Roof planes from a point cloud.

#### Tools (Finishes · Reference · Modify)
**Finishes**
- **Room To Floor** — create architectural or structural floors from selected room boundaries.
- **Door Threshold** — create threshold floors at the base of selected doors, sized to the opening and host wall.
- **Tile Layout** — 3-step wizard to extract floor boundaries, choose a tile pattern per floor, and place the generated tile layout on the active sheet.

**Reference**
- **Image to Drafting** — create a Drafting View and import an image from disk or clipboard.
- **Property Line** — create property lines from Lightbox parcel survey data. Supports metes-and-bounds descriptions and coordinate-based input.
- **Text to Element** — transfer Text Note content to element parameters via bounding-box overlap in the active view.

**Modify**
- **Auto Join** — automatically join intersecting elements by category rules (Shift+Click for quick join).
- **Split Elements** — split Walls, Columns, or Floors at selected levels, preserving parameters.
- **Wall Cut Profile** — cut wall profiles or create openings based on intersecting linked model elements.
- **DatumSync** — synchronize grid lines, levels, and reference planes across views. Align 2D/3D extents, datum bubbles, and visibility between a source view and target views to maintain clean documentation.
- **Auto Adjust Base Offset** — recalculate Base Offset when changing Base Constraint so elements keep their absolute elevation.

---

### Rebar & Assembly

#### Cast Unit Manager
Create assemblies with their rebar for many beams, columns or footings at once (from the selection or a filter), sync loose rebar into its assembly, rename marks as a series (prefix + start + step per assembly type), and set the rebar Partition by rule (assembly mark, level, host type or workset). Hosts in a group, from a link, or already in an assembly are skipped, with the reason shown before anything runs.

#### Clone Drawing
Copy the finished drawing of one assembly — views, sheet, annotations, tags, dimensions and spot elevations — to similar assemblies. Clone settings are chosen per object type and saved as presets; whatever cannot be matched on a target is listed with the reason. One undo step for the whole run; views, sheets and model elements are never deleted.

#### Rebar Check
Data checks Revit does not run: rebar without a valid host, rebar missing from its assembly, assemblies without drawings, duplicate numbers, bars outside their host, no partition, unknown shape. The scan is read-only; **Fix** syncs loose rebar into its assembly in one undo step.

#### BVBS Export
Write BVBS BF2D `.abs` files for bending machines from shape-driven rebar — one file per assembly or one for all. Preview every mark with the legs that will be written; free-form 3D bars and curved legs are listed as skipped, never exported. Each file is read back and its checksums verified.

#### Rebar Wizard
Reinforce rectangular beams, columns and pad footings from a preset: beam bottom and top bars with stirrups denser at both ends, column verticals with ties denser at top and bottom, and a two-layer pad-footing mesh. New bars join the host's assembly; presets are saved per user, and one Ctrl+Z undoes the whole run.

---

### Annotation & Data

#### Mana (ManaDWG · Auto Dimension · ManaSelect)
- **ManaDWG** — manage CAD imports and CAD links — list, rename, and delete DWG imports and links from a single interface.
- **Auto Dimension** — automatically create dimension chains for walls, columns, doors, lifts, and grids in the active or a chosen view.
- **ManaSelect** — smart selection manager with 4 modes: Quick Select (filter by parameter value or text), Select Similar (by type, family, or category), Select on Sheets (locate title blocks and CAD imports), and a Quick Filters sidebar.

#### manaData (ManaSched · ManaPara · ManaContains)
- **ManaSched** — export schedule data to Excel with formatting preserved, import updated values back into schedule rows, and duplicate schedules.
- **ManaPara** — Parameter Manager — transfer parameter values between elements by rule, assign Text Note content to element parameters via spatial overlap, and write schedule values into filled region parameters.
- **ManaContains** — find elements contained in Rooms, Areas, Spaces, Zones, Masses, or Scope Boxes. Assign parameter values to contained elements from their container, or aggregate element data back into the container.

#### ManaAnno
Annotation manager for Dimensions and Text Notes in one window: find dimensions or notes by type name or content and jump to their view, delete selected dimension instances, auto-rename Dimension and Text Note types from their properties, and set prefix, suffix, above, below or override text on dimensions.

#### Make Pattern
Vector hatch studio. Draw model and drafting fill patterns on a canvas with grid snapping and ortho lock, set the unit module and stagger shift (running bond), preview the pattern tiled over a large surface, then create it in Revit as a Fill Pattern / Filled Region or export an AutoCAD `.pat`. Linework can be imported from the selection in the active view.

#### IFC-SG Suite
Unified IFC-SG manager. **Subtype Assigner** loads mapping rules from Excel and assigns IFC Export Class and Predefined Type parameters; **Compliance Checker** verifies that required parameters exist and are filled, against CORENET X rules.

---

### Views & Sheets

#### View Tools (SheetGen · CropSync · PDF Import)
- **SheetGen** — create floor-plan views from a room list. Select rooms, choose a View Family Type and naming template, and generate all views in one transaction.
- **CropSync** — synchronize crop regions, annotation crops, and crop view settings across selected views to ensure consistent sheet alignment and view boundaries.
- **PDF Import** — import PDF pages into selected Revit views sequentially. Supports 150 / 300 / 600 DPI.

#### ManaViews
Browse and filter all views, rename views in bulk with naming rules, update view templates across multiple views, and remove unused views.

#### ManaSheets
Manage sheets in one unified interface — browse with live search, sync sheet data to/from Excel, place views on sheets, create sheet sets, and renumber sheets.

#### BatchOut
Export sheets to PDF, DWG, NWD, and IFC formats in batch.
Supports combined PDF, custom naming patterns, sheet ordering, and revision tracking.

---

### Network Traffic

Every connection is either **user-initiated** or the once-a-day update check,
which can be switched off (see [Staying up to date](#staying-up-to-date)).

| Component | Destination | When |
|---|---|---|
| Auto-update | `github.com` / `raw.githubusercontent.com` / `cdn.jsdelivr.net` | First Revit start of each week (retried the next day if GitHub was unreachable), unless `"auto_update": false` |
| T3Lab Assistant / AI tools | Only the provider you configure: `api.anthropic.com`, `api.openai.com`, `api.deepseek.com` — or nothing leaves the machine with Ollama (`localhost:11434`) / LM Studio (`localhost:1234`) | User sends a message or runs an AI-powered tool |
| Skills install | `api.github.com` — zipball of the repo you paste | User installs or updates skills |
| Knowledge (semantic search) | Ollama on `localhost`; the first enable downloads the embedding model (~270 MB) | User turns semantic search on |
| MCP Server | `localhost:8080` (host/port configurable) | Only while the MCP server is running |
| ManaFami (Cloud) | User-configured Vercel URL in `~/.t3lab/family_loader_config.json` | User opens the cloud family catalogue |
| Feedback | None — opens a `mailto:` link in the default email client | User sends feedback |
| Cloud Links | `acc.autodesk.com` / `health.autodesk.com` / `status.bluebeam.com` | Opens in the default browser on click |

API keys, chat history, projects, attachments and the knowledge index stay on
the machine under `%APPDATA%\T3LabAI`.
