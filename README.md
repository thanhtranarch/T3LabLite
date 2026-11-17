# T3Lab Lite

**T3Lab Lite** is a pyRevit extension running on **CPython 3** built for Revit users who want to work faster.
It covers batch export, sheet & view management, datum & crop synchronization, CAD-to-BIM conversion, model auditing, rebar and assembly detailing (cast units, drawing cloning, BVBS export), IFC-SG compliance, family generation and management, a built-in AI assistant that runs Revit tools from plain Vietnamese or English, MCP integration that lets Claude, ChatGPT (Codex) and other AI apps work with Revit — and with AutoCAD and other Autodesk apps through their T3Lab add-ins — and Agent Cursors, a live view of what each AI app is doing in the model.

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

The first time it opens it asks what to call you (once), or keeps the name you set in an earlier install; change it in **LLMs Setting → General**. The name is saved on this computer only, in `%APPDATA%\T3LabAI\t3lablite_user.json`, separate from T3Lab AI's files in the same folder.

- **Skills** — reusable instruction packs that activate on what you ask, 33 built in (ISO 19650 naming, LOD, worksets, QA checklist, BEP, COBie handover, clash coordination, sheet & annotation standards, construction wiki and drawing-completion rules, …). Type `/skill-name` to force one, or install more from a GitHub repo — Claude's `SKILL.md` format is read as-is.
- **Knowledge (RAG)** — index folders of PDF / TXT / MD and get answers with citations. Keyword search works offline; semantic search is optional and runs on a local Ollama embedding model.
- **Projects** — per-project chat history, custom instructions, default model, knowledge folders, remembered facts, and daily scheduled prompts.
- **Project memory shared with Claude** — facts the Assistant remembers ("remember that ...") and conventions saved by AI apps through MCP (`t3lab_memory_save`) are one memory per project. Claude Code's memory of the project's linked folders (CLAUDE.md, auto memory) is searched too, read only. A long memory reaches the prompt as an index and the Assistant searches the rest with `recall_memory`, so each message stays small.
- **Remote MCP servers (Claude)** — add MCP servers on the internet (https) in **LLMs Setting → Models**; Claude runs their tools through its MCP connector. Those calls do not go through "Ask before model edits".
- **Context & attachments** — the assistant sees the active view and current selection, and accepts attached files and images.
- **PDF comments** — read markups from a PDF, trace them to the matching sheet, and resolve them item by item.
- **Spell check** — proofread every Text Note in the model or just the active view.

#### Settings (pulldown)
- **MCP Control** — connect AI apps to Revit and other Autodesk apps. Two tabs:
  - **Connect** — pick the app at the top: **Revit**, or **AutoCAD**, **Civil 3D**, **Navisworks**, **3ds Max**. For Revit, step 1 starts and stops the local MCP server; for another app, it shows whether that app's T3Lab add-in (its MCP bridge) is installed. Step 2 connects Claude Desktop / ChatGPT (Codex) / Antigravity to the picked app in one click; **Manual config** at the bottom gives the JSON, TOML or `claude mcp add` entry for any other AI app. The last app you picked is remembered. See [MCP for other Autodesk apps](#mcp-for-other-autodesk-apps).
  - **Advanced** — auto-start, port (48884–48894), what AI apps may do (read-only mode, confirmations, code execution, toolset — see [MCP security](#mcp-security)), the file task watcher and teaching capture.
- **Agent Cursors** — see what AI apps do in Revit while they work (off by default). See [Agent Cursors](#agent-cursors).
- **LLMs Setting** — the settings hub shared by every AI-powered T3Lab tool: provider (Claude, OpenAI, Gemini, DeepSeek, any OpenAI-compatible server — OpenRouter, Azure OpenAI, vLLM/LiteLLM, Groq, Mistral, xAI —, Ollama, LM Studio), API key or server URL, live connection status per provider and a test message; plus your display name, "ask before model edits", auto model, deep reasoning / maximum quality, chat detail, and the Projects, Knowledge and Skills tabs. There is no model list to pick from: with **Auto model** (General) Claude picks Haiku, Sonnet or Opus for each request; other providers use their default model, and **Custom** uses the model name you type.
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
Quick links to Autodesk Forma and the Autodesk Health dashboard, plus **PDF Markup** — a split button with **Bluebeam Status** and **T3Lab PDF Markup** (opens [T3Lab PDF](https://pdf.t3lab.space/), offline PDF markup for drawings). The one you used last stays on top, also after Revit restarts.

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
- **BatchLink** — manage Revit links (status, reload, unload, pin, remove, Switch to Cloud), batch-link models from a folder or from Autodesk Docs as cloud links, set the workset each link sits on, and control per-view link display.

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
- **Wall Cut Profile** — create wall openings where ducts, pipes, cable trays, beams or columns cross walls, from the open model plus any linked models you tick.
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
- **ManaSched** — import an Excel or Google Sheets range as a schedule, drafting view or legend with its formatting and keep it linked (Update Tables); export schedule values to Excel, edit, review every change and apply in one undo step.
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

- **Export report** — after each export, `BatchOut_Report_<date_time>.xlsx` (or `.csv`) in the output folder lists every sheet / view and format with its status, the reason for a failure, the file path and size.
- **Profiles** — save the current settings as a profile (to a shared folder for your team) and import a profile from any folder.
- **Skip image reference files (DWG)** — Revit writes every image in a view or title block (logos) as a separate file next to the DWG. Turn this on to export the DWG without images and without those files. Off by default; saved in the profile.

---

### Agent Cursors

**Support ▸ Settings ▸ Agent Cursors** turns on a live overlay of what each AI
app does in Revit through MCP (off by default; the button lists the benefits
and drawbacks before you switch).

- **One cursor per AI app and action** — Read, Create, Modify, Delete, View,
  Export, Code — in the app's colour, with a plain caption of what it does
  ("Moving elements · 3 elements") and the MCP tool name under it.
- **Move, click, next** — the cursor visits each element the call works on;
  a box with the category and size ("Walls · 1200 × 300") appears as it
  arrives. New elements grow in, edits pulse, deletes show a red cross, moves
  leave a dashed ghost. A read or review scans the view instead.
- **Goal, tasks and changes** — an AI app shows its goal and its list of
  tasks over the model with the MCP tool `t3lab_narrate` ("Renumber the
  doors on Level 2": 1 Read the doors · 2 Number them · 3 Check). The bar
  ticks the tasks off, and while a task runs the cursor names it ("Step 2/3 ·
  Number them") instead of the tool. Code an AI runs is named by its
  description or by the code itself ("Editing stairs"), not "Running code".
  Changed values show under the box: `Comments: "" → Checked`.
- **Your OK on the bar** — when an AI app needs your OK (code, deletes,
  sync…), the question appears on the bar with **Allow** / **Deny** (and
  **Details** for the code) instead of a Revit dialog, so Revit stays free;
  it stays in sight even while the AI app is in front. With the cursors off,
  the Revit dialog is used as before. The first time an AI app works in
  Revit while the cursors are off, T3Lab offers to turn them on (once per
  session; "Don't ask again" stops it).
- **While it works** — long calls say "Working 8 s"; a call waiting for your
  OK pulses amber. Two AI apps editing the same element flash it red.
  Targets outside the view get an edge arrow with a count.
- **Follow** opens the views the AI works in, **Zoom** brings off-screen
  targets into view; **Select**, **Tile**, **Less motion**, **Speed** (Calm /
  Normal / Fast) and **Glow** (experimental: outlines the touched elements
  inside 3D views) switch on the bar at the bottom of the view.
- **When the AI goes idle** — a summary per AI app, **Select touched**
  (selects every element the AI changed; no model change) and **Replay** of
  the last 20 steps. **Keep bar** (on by default) keeps the bar and the
  dimmed cursors on screen between calls, so they do not blink on and off
  while you work with an AI app. The **X** on the bar turns the cursors off.

Each tool call does a little extra work (bounding boxes of up to 12
elements), Follow switches the active view, and the overlay shows up in
screenshots and screen recordings.

### MCP for other Autodesk apps

MCP Control → **Connect** registers more than Revit: pick **AutoCAD**,
**Civil 3D**, **Navisworks** or **3ds Max** and connect your AI app to it. Each app is reached through its own T3Lab add-in, which puts an
MCP bridge at `%APPDATA%\<folder>\bridge.py`; the AI app runs it with
`python -I` and the bridge finds the app by itself (no port).

| App | MCP server | Bridge |
|---|---|---|
| Revit | `t3lab-revit` | this extension (MCP Control starts the server) |
| AutoCAD | `t3lab-autocad` | `%APPDATA%\T3LabCAD\bridge.py` — T3LabCAD add-in, 2022-2027 |
| Civil 3D | `t3lab-autocad` | same as AutoCAD (T3LabCAD loads in Civil 3D too) |
| Navisworks | `t3lab-navisworks` | `%APPDATA%\T3LabNavisworks\bridge.py` — no add-in yet |
| 3ds Max | `t3lab-3dsmax` | `%APPDATA%\T3LabMax\bridge.py` — T3LabMax add-in, 2022-2027 |

For AutoCAD, Civil 3D and 3ds Max, **Connect** downloads the add-in from the
[t3lab_CADmcp](https://github.com/thanhtranarch/t3lab_CADmcp) release and
installs it before writing the AI-app entry; restart the app once after the
first install. Navisworks shows "No T3Lab add-in for Navisworks yet" and keeps
**Connect** off. A bridge somewhere else can be set with
`"<app>_bridge": "<path>"` (e.g. `"navisworks_bridge"`) in
`%APPDATA%\T3LabAI\mcp_paths.json`. MCP Control does not write to Claude
Desktop's config while Claude Desktop is running (it would drop the entry
when it quits) — quit it from the tray icon first.

### Data hub

AI apps can ask about a model **with Revit closed**, across models, and ask
what changed since last time. Revit saves a snapshot of the open model
(elements, types, parameters, warnings) and the MCP bridge keeps the
snapshots in a local database. Ask your AI app to "publish this model to the
hub", then for example "how many doors per level", "which doors have no fire
rating", "what changed since the last publish".

- One MCP tool, `t3lab_hub`: `publish` (needs Revit), `summary`, `query`,
  `diff`. Results are small (counts and filtered rows), so a question costs
  far fewer tokens than reading elements one by one.
- Stored in `%LOCALAPPDATA%\T3Lab\hub`: `snapshots\` holds the data,
  `hub.db` is a cache rebuilt from it with
  `python %APPDATA%\T3LabAI\bridge.py hub rebuild`. Do not move `hub.db`
  to OneDrive or a network folder.
- Nothing to switch on. After an update, restart Revit (it refreshes the
  bridge files), then restart the AI app. A diff needs two publishes.
- The hub never changes a model, and its data can be old: AI apps re-read the
  live model before any edit.

Details: [docs/data-hub.md](docs/data-hub.md).

### Slow tool opens

To see where the time goes when a tool opens, turn on the timing log
(PowerShell):

    New-Item "$env:APPDATA\T3LabAI\open_timing.on" -Force

then restart Revit and open the slow tools. Each window adds one line to
`%APPDATA%\T3LabAI\open_times.log` with the time of every step (bootstrap,
imports, XAML, wiring, the tool's own data load, layout, first draw); the
slowest step is marked `*`. Delete `open_timing.on` to stop. Send the log
with a bug report.

---

### Network Traffic

Every connection is **user-initiated**, the weekly update check, or
usage tracking. The update check can be switched off (see
[Staying up to date](#staying-up-to-date)); usage tracking is off until you
choose **Allow** in the question Revit shows once (see
[docs/PRIVACY.md](docs/PRIVACY.md)).

| Component | Destination | When |
|---|---|---|
| Auto-update | `github.com` / `raw.githubusercontent.com` / `cdn.jsdelivr.net` | First Revit start of each week (retried the next day if GitHub was unreachable), unless `"auto_update": false` |
| T3Lab Assistant remote MCP servers | Anthropic's servers connect to the https MCP servers you add (Claude only) | Every Assistant message while a server is on |
| T3Lab Assistant / AI tools | Only the provider you configure: `api.anthropic.com`, `api.openai.com`, `generativelanguage.googleapis.com`, `api.deepseek.com`, the Base URL you set for Custom — or nothing leaves the machine with Ollama (`localhost:11434`) / LM Studio (`localhost:1234`) | User sends a message or runs an AI-powered tool |
| Skills install | `api.github.com` — zipball of the repo you paste | User installs or updates skills |
| Knowledge (semantic search) | Ollama on `localhost`; the first enable downloads the embedding model (~270 MB) | User turns semantic search on |
| MCP Server | `127.0.0.1:48884`–`48894` only (never reachable from the network), bearer token from `%APPDATA%\T3LabAI\mcp_token.txt` | Only while the MCP server is running |
| Other Autodesk apps (MCP Control → Connect) | None — writes the AI app's local config file only; the add-in bridge talks to its app on this computer | User clicks Connect |
| Agent Cursors | None — drawn locally over Revit | Only while turned on |
| Usage tracking | `t3lab.space/api/revit/tracking` — a random install ID made on this computer, Revit version, a salted code of the model title (not the title), ribbon button or MCP tool name, result and redacted arguments; the server also sees the sender's IP address. This is still personal data (pseudonymous). Never sent: code, secrets, full paths, prompts and other free text, e-mails, phone numbers, user names. See [docs/PRIVACY.md](docs/PRIVACY.md) | Revit start, ribbon click, MCP tool call — **only after you choose Allow** (asked once on the first Revit start). Withdraw with `"tracking_enabled": false` in `%APPDATA%\T3LabAI\mcp_paths.json` |
| ManaFami (Cloud) | User-configured Vercel URL in `~/.t3lab/family_loader_config.json` | User opens the cloud family catalogue |
| Feedback | None — opens a `mailto:` link in the default email client | User sends feedback |
| Cloud Links | `acc.autodesk.com` / `health.autodesk.com` / `status.bluebeam.com` / `pdf.t3lab.space` | Opens in the default browser on click |

API keys, chat history, projects, attachments, the knowledge index, your
display name (`t3lablite_user.json`) and the timing log stay on the machine
under `%APPDATA%\T3LabAI`.

### MCP security

These rules apply to AI apps connected through MCP (Claude Desktop, Codex,
Cursor, ...). The T3Lab Assistant has its own "Ask before edits" setting.
Change them in **MCP Control → Advanced → Security**, or in
`%APPDATA%\T3LabAI\mcp_paths.json`. They apply from the next tool call.

| Setting | Default | What it does |
|---|---|---|
| `mcp_read_only` | `false` | `true`: AI apps only see and call tools that change neither the model nor any file |
| `mcp_confirm_destructive` | `true` | Revit asks Yes/No before delete, purge, ungroup, deleting a view template or link, save as, sync, close without saving, or a `run_batch` that contains any of these. No is the default answer |
| `mcp_code_execution` | `"ask"` | `send_code_to_revit`: `"off"` hides and refuses it, `"ask"` asks in Revit each time, `"allow"` runs it without asking |
| `mcp_toolset` | `"full"` | Which tools AI apps see: `"full"` (all), `"standard"` (no teaching / training tools), `"documentation"` (read tools plus views, sheets, annotation, schedules, revisions and export — nothing that models or deletes). Fewer tools means fewer tokens on every request. A hidden tool is refused, also inside `run_batch` |
| `auto_start_file_watcher` | `false` | The file task watcher (`~/T3Lab_AI_Data/task.json`) runs code with no token, so it only starts with Revit after you click **Start Watcher** once |

`t3lab_narrate` (the Agent Cursors goal banner) changes nothing, so it stays
available in read-only mode.

The Revit question says what the call does, e.g. *Delete 3 elements (id 101,
102, 103)?* For `send_code_to_revit` it shows the purpose the AI gives in the
`description` argument, and under it what T3Lab finds in the code itself:
deletes, parameter changes, created elements, categories, transaction names,
and file, network or program access. The AI cannot hide those by wording its
description differently. **Show details** shows the full code.

Tools are listed with MCP annotations (`readOnlyHint`, `destructiveHint`), so
clients that support them can ask before edits too. `run_batch` runs several
edits as one Revit undo step and rolls them all back if one fails.

### MCP construction wiki & drawing rules

AI apps get the same skill library as the T3Lab Assistant through the
read-only `t3lab_wiki` tool (`list`, `search`, `read`), and the MCP
`initialize` answer tells them to read the drawing rules before they create,
annotate, complete or issue drawings. Pages (in Vietnamese):

| Page | What it covers |
|---|---|
| `drawing-completion-rules` | When a sheet counts as done, the read → plan → fix → check workflow with T3Lab tools, and what the tools cannot do (tags are not de-duplicated, `create_dimension` takes grids only) |
| `architectural-drawing-rules` · `structural-drawing-rules` · `mep-drawing-rules` | What each drawing type must show: plans, ceilings, elevations, sections, details; foundation / framing plans, rebar notation, bar schedule; plumbing, HVAC, electrical, fire protection |
| `drawing-issue-checklist` | Read-only check before an issue: numbering, drawing index, title blocks, revisions, cross references |
| `drawing-conventions` · `drawing-set-structure` · `construction-glossary` | Paper sizes, scales, line weights, text heights, symbols, title block; design stages and sheet order; Vietnamese–English–Revit terms |

Your own pages in `%APPDATA%\T3LabAI\skills` and the active project's
`skills` folder are listed next to the built-in ones, and AI apps are told
that a project BEP wins over the generic examples in the built-in pages.
Pages switched off in **LLMs Setting → Skills** are hidden from AI apps as
well.

**Project memory.** `t3lab_memory` (read only: `list`, `search`, `read`)
returns what is remembered about the active T3Lab project: conventions saved
in the T3Lab Assistant or by an AI app with `t3lab_memory_save`, and Claude
Code's memory of the folders linked to the project (`CLAUDE.md`,
`.claude/CLAUDE.md`, `CLAUDE.local.md`, `~/.claude/projects/<folder>/memory`).
Also the resource `t3lab://memory/project`. Read-only mode hides
`t3lab_memory_save`; Claude Code's own files are never changed.

The same pages are also MCP **resources** (`t3lab://wiki/<page>`), so you can
attach one by hand in apps that show resources (Claude Desktop's **+** menu,
`@` in Claude Code). Three MCP **prompts** start the common workflows with the
right pages and tools: `complete-drawings`, `issue-drawings` and
`model-health-check` (read-only).

### MCP protocol versions

The server and the bridge speak MCP `2026-07-28` (no handshake: the version
travels with every request, `server/discover` describes the server) and the
earlier `2025-11-25`, `2025-06-18`, `2025-03-26` and `2024-11-05` (with the
`initialize` handshake), so old and new AI apps both connect. Both eras are
answered with Revit closed.
