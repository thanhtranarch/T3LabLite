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
3. Commit and push -- Check Update on user machines will pick it up, and the
   first Revit start after the update shows that section in a popup
   (`lib/core/whats_new.py`). Keep the first lines short: they fill the popup.

## [Unreleased]

### Added

- **FamiGen Studio (needs pyRevit 7.0+; the chat needs 7.1+):** the FamiGen
  button now opens a chat-and-3D window. Describe a family or drop a photo /
  PDF cut sheet; Claude builds it, you review it (detail levels, visual styles,
  types, sizes), then Create Family. Shift+Click opens the classic window.
- **FamiGen — families that resize and vary:** family types, reference planes,
  labelled dimensions and locked faces, repeated parts and detail levels;
  four new categories (Pipe Accessory, Duct Accessory, Sprinkler, Electrical
  Fixture) and a design spec for every category.
- **FamiGen — closer to your picture:** Claude first writes down what it sees
  (camera, scale, every part with size and colour), then models it; the
  check compares a render from the picture's own camera, up to twice. Pipes
  and taps get real elbows (`bend_radius`) instead of kinks. New **Ref** view.
- **FamiGen — every Revit form:** Swept Blend is new, Sweep profiles can be
  drawn as the cross-section seen along the path, voids cut the solids they
  overlap, and Claude is told which form fits which shape (not only
  Extrusion). Studio has a **Voids** switch.
- **FamiGen — models in one piece:** a new `Tube` form for curved rods, rails,
  handles and tube slides (smooth, no kinks), a check that names parts left
  floating (Claude fixes them once before you see the model), joint rules in
  every category spec, and smooth shading for round parts in Studio.
- **FamiGen — AI apps over MCP:** while Claude Desktop / Code or ChatGPT work
  with FamiGen, Studio switches to them by itself. New MCP tools
  `famigen_get_studio_state`, `famigen_validate_family`,
  `famigen_patch_family` and `famigen_list_templates`.

### Added

- **FamiGen — models that make sense in space:** every AI family now states its
  concept (what it is, how it is used, where it stands), each part's role, and
  the joints between parts. FamiGen checks the geometry against that and sends
  problems back to the AI: parts that do not really touch, joints that miss,
  stairs / slides / bridges that do not land at both ends, and anything
  blocking a route or the inside of a pipe.
- **FamiGen — hollow parts:** pipes, tube slides and tunnels are hollow
  (`"wall"` on Cylinder / Tube); basins, bowls and bins modelled as a solid
  block are reported. Each of the 17 category guides says what is hollow in
  that category and which joints to declare.

### Changed

- **FamiGen Studio — one title bar:** the window now shows only the T3 title
  bar (drag it to move, double-click to maximize). If pyRevit's own title bar
  cannot be hidden, the Studio hides its own instead.
- **FamiGen Studio follows the T3 design system:** T3 colours, type sizes,
  4 px spacing, 26 px rows and flat controls (no shadows or animation). The
  right panel now reads like Revit: a Properties panel with collapsible groups
  (Identity Data, Geometry, Dimensions) and a parts tree grouped by
  subcategory, with repeated parts folded into one row ("Post ×8"). View
  controls sit in a docked toolbar above the 3D view instead of floating on it.

### Fixed

- **FamiGen (classic window) — the Family JSON box was one line high:** it now
  fills its panel.

- **MCP Control — "module 'core.paths' has no attribute 'atomic_write'":**
  after an update installed while Revit was open, MCP Control kept an old copy
  of the settings code in memory and could not configure Claude Desktop,
  ChatGPT (Codex) or Antigravity. It now loads the new copy by itself.
- **Model Auditor, Mana Sheets — raw "cannot import name 'revit' from
  'pyrevit'" traceback:** these tools now show the T3Lab error dialog, which
  says to restart Revit, and log the error to `%APPDATA%\T3LabAI\errors.log`.

## [2.6.1] - 2026-11-18

BatchOut no longer fails with "Duplicate type name within an assembly", and the error dialog now tells you when Revit needs a restart.

### Fixed

- **BatchOut — "Duplicate type name within an assembly":** BatchOut's Revit
  event handler now gets a new .NET type name each time it loads, so BatchOut
  no longer fails when its code is loaded a second time in the same Revit
  session.
- When Python in a Revit session was reloaded (pyRevit > Reload, an extension
  switched off and on, or an update without a restart), the error dialog now
  says to restart Revit instead of "Try again". This covers "Duplicate type
  name within an assembly" and "cannot import name 'revit' from 'pyrevit'".

## [2.6.0] - 2026-11-18

Agent Cursors now show only the bar by default, stay on screen beside your AI app, and slow Revit down less.

### Changed

- Agent Cursors show only the bar by default: the agent, its goal and tasks,
  the current step and the last steps. No cursor, boxes or effects on the
  model, and no view switching, zooming, selecting, tiling or Glow. Turn on
  **On model** (options on the bar, or Settings > Agent Cursors) to get them
  back.

### Fixed

- Agent Cursors stay on screen while you work in the AI app next to Revit.
  They hide only when Revit is minimized or another window covers its views,
  and no longer flash up on every step.
- Boxes and the cursor land on the right element after Follow opens a view or
  Tile rearranges the views: Agent Cursors wait for Revit to finish the layout,
  and drop boxes drawn for the old layout.
- The cursor no longer jumps ahead mid-glide when the views change.
- Faster steps with Glow on: the view is redrawn only when the outlined
  elements change, not on every call.

## [2.5.1] - 2026-11-17

Agent Cursors no longer show an extra "Claude" cursor next to the AI app you use.

### Fixed
- **Agent Cursors — an extra "Claude" / "AI app" cursor next to your AI app:** when a step failed, or the AI app listed or switched documents, the T3Lab connection made a check of its own ("Listing open documents") and it was drawn as a second AI app. That check no longer draws a cursor, and every step of your AI app carries its name.

## [2.5.0] - 2026-11-17

Agent Cursors name the AI app you really use, and leave the model when it stops.

### Fixed
- **Agent Cursors — the wrong AI app name ("Claude" while using Antigravity):** a step that came without an app name was labelled "Claude". The name now comes from the app that started the connection (Antigravity, Claude Desktop, ChatGPT / Codex, Cursor), so it is right without connecting again; a step with no name at all reads "AI app". Restart Revit, then restart the AI app.
- **Agent Cursors — they stayed after the AI app stopped:** when the AI app goes quiet, its cursor, the boxes and the orange Glow outline now leave the model; **Keep bar** keeps only the bar (folded). They come back on its next step.

## [2.4.5] - 2026-11-17

Agent Cursors show the right AI app name: Antigravity is no longer shown as Claude.

### Fixed
- **Agent Cursors — the wrong AI app name:** Antigravity showed as "Claude". MCP Control's Connect now writes which app it connects (Claude Desktop, ChatGPT (Codex), Antigravity) into that app's settings, and the cursor and the bar use that name. Apps connected before this update show **Connect** again in MCP Control: click it once. An app that sends no name shows as "AI app" instead of a guess.

## [2.4.4] - 2026-11-17

Agent Cursors show the work on the model: one cursor per AI app, readable names, and the elements it builds highlighted as it goes.

### Changed
- **Agent Cursors — one cursor per AI app:** an app no longer leaves a second, faded cursor behind for each kind of action; its one cursor moves on and the action only tints the left edge of its tag.
- **Agent Cursors — readable names and tags:** Claude Desktop's agent mode shows as "Claude" instead of `local-agent-mode-t…`, other unknown apps get a clean name ("Acme Bot"). A tag flips to the left of or above its pointer near the edge of the view instead of being cut off, and two tags never cover each other.
- **Agent Cursors — the work shows on the model:** elements a call names are framed while it runs and the cursor goes to them. What the AI created or changed — also from code, which reports no ids — is highlighted for a few seconds with a chip ("Walls ×3 created") and a dotted line joining new elements in the order they were made (not with Less motion).
- **Agent Cursors — calmer reads:** a call with no element ("Listing levels") scans next to where the cursor last worked instead of jumping across the view; a new cursor starts in the middle of the view.

### Fixed
- **pyRevit "Routes servers failed activation" (WinError 10013) at Revit start:** pyRevit's Routes server used port 48884, the same port as the T3Lab MCP server. T3Lab now moves Routes to port 48870 when it is switched on, on every pyRevit version (6.x and 7.x); nothing changes when Routes is off.
- **Agent Cursors — boxes in plans and sections:** element boxes could be missing in plan and section views (they were only reliable in 3D).

## [2.4.3] - 2026-11-17

A new Agent Cursors bar with a minimal cursor and a Pause button, and Connect keeps the AutoCAD add-in up to date.

### Changed
- **Agent Cursors — new bar and a minimal cursor:** the cursor is a small pointer with one tag ("Claude · Moving 12 walls") instead of an arrow with two labels. The bar shows each AI app (its pointer is its avatar), a status (Working 8 s, Waiting for your OK, Paused, Idle), the goal with its progress, what runs now and the values it changed, and the last 3 steps; the 7 option chips moved into ⚙ Options. It follows Revit's light or dark theme, can be dragged anywhere (double-click puts it back) and folds to one line after 10 s without a call (the mouse opens it; Allow / Deny show on that line). Many targets get one frame round them all.
- **Agent Cursors — every option on by default:** Follow, Zoom, Select, Tile, Less motion, Glow and Keep bar start switched on. An option you already turned off stays off.
- **Agent Cursors — Pause:** ⏸ on the bar holds every model change until you click ▶; read-only tools keep running and the AI app is told it is paused.
- **Agent Cursors — Glow turns itself off** when Revit cannot draw it, with a note on the bar, instead of failing on every call.
- **Agent Cursors — options in the Settings dialog:** Settings › Agent Cursors now shows a checkbox for each option (Follow, Zoom, Select, Tile, Less motion, Glow, Keep bar) instead of a line of text. Tick what you want and click Turn On; when the cursors are on, the same dialog saves your changes or turns them off.
- **MCP Control — Connect keeps the AutoCAD add-in up to date:** with the add-in already installed, Connect also fetches the latest release (at most every 10 minutes; offline it just connects) and says when it updated it. The add-in's ribbon buttons have new icons that follow AutoCAD's light or dark theme.

## [2.4.2] - 2026-11-17

MCP Control installs the AutoCAD / Civil 3D add-in for you when you click Connect, and BatchOut exports with PDF24 again.

### Added
- **MCP Control — Connect installs the AutoCAD add-in:** pick AutoCAD or Civil 3D and click Connect: MCP Control downloads the T3Lab add-in (T3LabCAD) from GitHub and installs it for every AutoCAD and Civil 3D from 2022 to 2027, then connects your AI app. Restart AutoCAD once after the first install. Step 1 also lists the versions found on this PC. Civil 3D uses the AutoCAD add-in (`t3lab-autocad`).

### Removed
- **MCP Control — Navisworks, 3ds Max and Inventor** are no longer in the app list: only apps whose T3Lab add-in Connect installs in one click are shown (Revit, AutoCAD, Civil 3D).

### Fixed
- **BatchOut — PDF24 printer:** export failed with "The PrintToFile property cannot be set to false while the printer is Virtual". When Revit treats PDF24 as a virtual printer, BatchOut now prints to a file and turns it into a PDF with the Ghostscript that comes with PDF24; your printer's Print to file setting is put back afterwards.

## [2.4.1] - 2026-11-17

MCP fixes: AI-app tool calls no longer hang, Connect never wipes your AI-app config, Vietnamese names reach Revit intact, and read-only mode is enforced everywhere.

### Fixed
- **AI apps — every tool call hung for 2 minutes:** calls from Claude Desktop, ChatGPT (Codex) or Antigravity waited 120 s and then failed with "still running in Revit", and other calls queued behind them. The T3Lab Assistant was not affected.
- **MCP Control — Connect could wipe your AI-app config:** a Claude Desktop or Antigravity config saved with a BOM (Notepad, PowerShell), a comment or a trailing comma, or a Codex `config.toml` with a non-UTF-8 character, was replaced by the T3Lab entry alone, losing your other MCP servers and settings. Such a file is now read (BOM) or left untouched with a message saying what to fix.
- **MCP Control — "App switch failed: no attribute '_configure_all_btn'"** when picking another app in the Connect tab.
- **AI apps — Vietnamese text:** view, sheet and parameter names such as "Tầng 1" or "Đường ống" reached Revit garbled, and some characters closed the connection.
- **AI apps behind a proxy:** with a system or company proxy, Revit showed as "not reachable" and MCP Control showed the server as "Stopped" while it was running. Local calls to Revit no longer go through a proxy.
- **Read-only mode:** AI apps could still call the teaching and training tools (change teaching settings, start a fine-tune) while read-only mode was on. They are now blocked like every other change.
- **T3Lab Assistant — tools missing after turning on MCP read-only mode:** MCP Control's read-only mode and toolsets are for AI apps only; the Assistant now always keeps its full tool list (e.g. `place_wall`).
- **AI apps — placing elements by type name:** `place_wall` with a wall type, placing point-, line- and surface-based families, listing family types and finding title blocks failed with "Name" or "type not found".
- **MCP Control — Antigravity path:** opening MCP Control before Antigravity created its config pinned a path Antigravity never reads, so Connect showed "Connected" with no effect. The live path is now used.
- **MCP settings file:** `mcp_paths.json` (read-only mode, code execution, paths) is written safely, so two writers at once can no longer erase your security settings; a file that cannot be read is kept as `mcp_paths.json.unreadable`.
- **T3Lab and T3Lab Lite on one machine:** they no longer overwrite each other's MCP bridge at every Revit start.
- **AI apps — tool list after a reload:** when Revit answered with an error, AI apps saw only 3 tools until they restarted; they now get the last known list.
- **Agent Cursors — broken icon on the bar:** icons on the bar fall back from Segoe Fluent Icons to Segoe MDL2 Assets, so none shows as an empty box; the "next" mark of a task list uses an icon every Windows version has.

## [2.4.0] - 2026-11-17

Agent Cursors (see what AI apps do and answer them on the bar), a data hub AI apps can read with Revit closed, PDF24 for BatchOut, MCP Control in one Connect tab, and the Assistant asks your name.

### Added
- **Agent Cursors (Settings → Agent Cursors):** see what AI apps do in Revit. One coloured cursor per app and action (Read, Create, Modify, Delete, View, Export, Code) moves to the elements it works on, with an effect per action: new elements grow in, edits pulse, deletes show a red cross, moves leave a ghost. The cursor moves to each element and clicks it, a read or review scans the view, captions say what the AI does in plain English ("Moving elements · 3 elements"), boxes name the element ("Walls · 1200 × 300"), and the bar lists the last 5 steps. New: the AI can show its goal over the model, changed values show as "Mark: A-101 → A-102", long calls say "Working 8 s", a call waiting for your OK pulses amber, two AIs editing the same element flash red, off-screen targets get an edge arrow, and when the AI goes idle the bar shows a summary with **Select touched** and **Replay**. The bar also has **Less motion**, **Speed** (Calm / Normal / Fast) and an experimental **Glow** that outlines the touched elements inside 3D views (off by default). **Follow** opens the views the AI works in, **Zoom** brings off-screen elements into view, **Select** and **Tile** are optional. Off by default; the bar at the bottom of the view switches the options.
- **MCP Control — one Connect tab for Revit and other Autodesk apps:** pick the app (Revit, AutoCAD, Civil 3D, Navisworks, 3ds Max, Inventor), then connect Claude Desktop, ChatGPT (Codex) or Antigravity to it; the manual JSON / TOML / `claude mcp add` entry is folded at the bottom. AutoCAD connects through the T3LabCAD add-in (`t3lab-autocad`); the other apps connect once their T3Lab add-in is installed. The Setup, Manual Config and AutoCAD tabs are merged into this one.
- **BatchOut — Skip image reference files (DWG):** export DWG without the logo/image files Revit writes next to it. Off by default; saved in the profile.
- **Agent Cursors — tasks and your OK on the bar:** an AI app can list its tasks (`t3lab_narrate`); the bar ticks them off and the cursor names the task it is on ("Step 2/4 · Number the risers") instead of the tool name. Code is named by its description or by what the code does ("Editing stairs"), not "Running code in Revit". When an AI app needs your OK (code, deletes, sync…), the question appears on the bar with **Allow** / **Deny** (and **Details** for the code) instead of a Revit dialog, so Revit stays free. With the cursors off, the Revit dialog is used as before, and the first time an AI app works while they are off, T3Lab offers to turn them on (once per session).
- **Agent Cursors — Keep bar:** on by default; the bar and the dimmed cursors stay on screen between calls instead of blinking off and on.
- **Data hub (MCP tool `t3lab_hub`):** ask your AI app to "publish this model to the hub", then ask about it any time, even with Revit closed — counts by category or level, filters on parameters, and what changed since the last publish. Answers are small, so they cost few tokens. Stored on this computer in `%LOCALAPPDATA%\T3Lab\hub`; the hub never changes a model. See [docs/data-hub.md](https://github.com/thanhtranarch/T3LabLite/blob/main/docs/data-hub.md).
- **BatchOut — PDF Engine:** PDF options → Output Settings → PDF Engine: **Revit (built-in)** or **PDF24 printer**. With PDF24 Creator installed, each sheet is printed to PDF24 and saved straight into the output folder under the BatchOut file name; your PDF24 settings are put back afterwards. Combine joins the pages with pdf24-DocTool. Saved in the profile.

### Changed
- **T3Lab Assistant — your name:** the Assistant no longer greets everyone with a built-in name. The first time it opens it asks what to call you (once), or keeps the name you set in an earlier install. Change it in LLMs Setting → General. Saved in its own file, `%APPDATA%\T3LabAI\t3lablite_user.json`.
- **LLMs Setting → Models:** Model Selection is gone — Claude picks Haiku, Sonnet or Opus for each request (General → Auto model). Custom keeps its model name box.
- **MCP bridge kept up to date:** every Revit start refreshes the MCP bridge files AI apps use, so new MCP tools reach them after a restart of the AI app, without opening MCP Control. T3Lab and T3Lab Lite on one computer no longer overwrite each other's bridge.
- **Faster opens:** Visual & Styles opens without reading the whole model (the Color Splasher tab loads when you open it), and CAD to Elements shows its window first, then reads the DWG layers.

## [2.3.0] - 2026-11-04

BatchOut export report and shareable profiles, usage statistics that ask first, and an MCP option to run AI code without the Yes/No dialog.

### Added
- **BatchOut — export report:** after each export, BatchOut saves `BatchOut_Report_<date_time>.xlsx` (or `.csv`) in the output folder, one row per sheet/view and format: number, name, revision, status, why it failed, file path and size. Turn it off or pick CSV in the Queue step → Report.
- **BatchOut — profiles:** the Queue step has a Profile row. Pick a saved profile to apply it, **Save As** to keep the current settings (save to a shared folder for your team), **Import** to load a profile from any folder.
- **MCP — turn off the code dialog:** MCP Control → Advanced → Security has a new "Ask in Revit before each run" box under "Let AI apps run Python code". Untick it and code from AI apps runs with no Yes/No dialog in Revit. The code dialog itself also has a "Don't ask again for AI code" tick box; tick it and click Yes to stop the dialog. Tick "Ask" again to bring it back.

### Changed
- **Usage statistics now ask first:** nothing is sent until you choose **Allow** in a one-time question on the first Revit start (Vietnam's Personal Data Protection Law 91/2025/QH15). What is sent is smaller too: a random install ID instead of one built from your Windows user and computer name, a code instead of the model name, and no prompts, free text, e-mails, phone numbers or user folders. See [docs/PRIVACY.md](https://github.com/thanhtranarch/T3LabLite/blob/main/docs/PRIVACY.md).

## [2.2.0] - 2026-10-25

LLM fixes from `t3lab-revit-api` (no silent switch to another AI vendor, real timeouts and retries, API keys that work in both T3Lab builds), Gemini and Custom providers, project memory shared with Claude, and MCP protocol 2026-07-28.

### Added
- **Project memory shared with Claude:** the T3Lab Assistant and AI apps (Claude Desktop, Claude Code ...) now use one project memory. AI apps search it with the new MCP tool `t3lab_memory` (also the resource `t3lab://memory/project`) and save conventions with `t3lab_memory_save`; the Assistant sees them in its next chat. Claude Code's memory of the folders linked to the project (CLAUDE.md, auto memory) is searched too, read only — switch it off in LLMs Setting → General.
- **T3Lab Assistant — smaller prompts:** a long memory, or one with Claude Code notes, now reaches the prompt as an index (newest facts + counts) and the Assistant looks up the rest with `recall_memory` when the task needs it, instead of carrying every fact in every message.
- **MCP protocol 2026-07-28:** AI apps on the newest MCP version connect without a handshake (`server/discover`), next to apps on 2025-11-25 … 2024-11-05. Both work with Revit closed.
- **MCP resources and prompts:** every wiki page is an MCP resource (`t3lab://wiki/<page>`) you can attach in the AI app, and three prompts start the common jobs: `complete-drawings`, `issue-drawings`, `model-health-check`.
- **MCP toolsets (MCP Control → Advanced → Security):** Full, Standard (no teaching tools) or Documentation (read, views, sheets, annotation, export — nothing that models or deletes). Fewer tools = fewer tokens per request.
- **T3Lab Assistant — remote MCP servers (Claude):** LLMs Setting → Models → Remote MCP servers adds tools from MCP servers on the internet (https), run through Claude's MCP connector. Tokens are stored encrypted like API keys.
- **LLMs Setting — Gemini (Google):** models listed live from your key, vision and tools, Deep reasoning maps to Gemini's reasoning effort.
- **LLMs Setting — Custom (OpenAI-compatible):** OpenRouter, Azure OpenAI, Groq, Mistral, xAI or a company vLLM / LiteLLM server with a Base URL, optional key and model name (Azure `api-key` header detected automatically).
- **Batch Link — link from Autodesk Docs:** Link Models links cloud models straight from Autodesk Docs (path type Cloud), never their Desktop Connector copy — the default when the open model is a cloud model. A Desktop Connector folder narrows the list. **Manage Links → Switch to Cloud** re-points disk / Desktop Connector links to the same cloud model, keeping placement and worksets.
- **Batch Link — View Display → Copy to This View:** copies a view template's RVT link display, Model and Annotation category tabs included, onto the active view, and warns when the view's own template controls its link display.
- **MCP security (MCP Control → Advanced → Security):** read-only mode for AI apps; Revit asks Yes/No before an AI app deletes, purges, ungroups, syncs, saves as or closes without saving; `send_code_to_revit` shows the code and asks first (or can be turned off). The T3Lab Assistant is not affected.
- **MCP `run_batch` tool:** several edits run as one Revit undo step and are all rolled back if one fails.
- **MCP:** tools are listed with `readOnlyHint` / `destructiveHint` annotations; `ping` and protocol versions up to 2025-06-18.
- **Construction wiki & drawing rules for AI apps:** 8 new pages — when a sheet is finished and how to finish it with T3Lab tools, what architectural, structural (rebar notation, bar schedule) and MEP drawings must show, a check before issuing drawings, paper sizes / scales / line weights / symbols / title block, drawing set order and design stages, and a Vietnamese–English–Revit glossary. AI apps read them with the new read-only `t3lab_wiki` tool, and Revit's MCP server now tells them to read the drawing rules before drawing work. The T3Lab Assistant uses the same pages as skills.

### Changed
- **LLMs Setting — Save Key checks the key first:** a key the vendor rejects is never saved. The key box shows only the last 4 characters (`••••••••abcd`), a failure says what failed, where and what to do next, and you are warned if Windows could not encrypt the key. Picking a model saves it (no Save Model button).
- **LLMs Setting — keys from environment variables:** with no key saved, `ANTHROPIC_API_KEY`, `OPENAI_API_KEY` or `DEEPSEEK_API_KEY` is used (read only, never written to settings.json).
- **API keys:** plain keys left by an older build are encrypted once when Revit starts, not on every read.
- **Token usage and cost** are recorded for OpenAI streamed agent turns and for DeepSeek, Ollama and LM Studio too.
- **Tool Hub:** unchanged since 2.1.0 (the LLMs Setting tab and `/toolhub` keep working with the new key handling).
- **OpenAI:** agent replies stream live; GPT-5 / o-series reasoning models get the right token-limit parameter and Deep reasoning → high effort; image, audio and codex models are hidden from the list. A provider that rejects a request parameter is retried without it instead of failing the turn.
- **T3Lab Assistant dead code removed (~400 lines):** the switched-off reply action bar ("Try again", 👍/👎), the unwired Reset Chat handler, and the never-called `parse_command` / `parse_command_local` chain with the local model's old JSON prompt. No visible change.
- **T3Lab Assistant code split:** `T3LabAssistantDialog.py` went from 11,930 to about 6,740 lines; helpers and seven method groups (sessions, projects, commands, documents, plans, spellcheck, tool cards) moved to `lib/GUI/assistant/`. No behaviour change intended. `docs/assistant-request-paths.md` maps every path a chat message can take.
- **MCP:** the server reports the extension version (from `version.txt`) instead of 1.0.0.
- **MCP confirmations say what the AI is doing, not how much code it sends:** `send_code_to_revit` now asks in Revit with the purpose of the code, e.g. *Run AI code in Revit: Add prefix A- to sheet numbers?*, and lists what T3Lab finds in the code: deletes, parameter changes, created elements, categories, transaction names, and file or network access. It no longer shows "N line(s) of IronPython". Other confirmations name the change too (*Delete 3 elements (id …)?*, *Close "Tower A" without saving?*). AI apps now send a one-line `description` with their code. In the T3Lab Assistant, tool cards and approval cards show the same summary instead of raw JSON.
- **File task watcher** (`task.json` / `task.py`) no longer starts on its own: it runs code with no token. Click **Start Watcher** once to turn it on for later sessions. A task file left from before is never run.
- **Usage tracking** no longer sends code, secrets or full paths (paths go as the file name only). README now lists it under Network Traffic.
- **Wall Cut Profile:** new window in the T3 layout. Openings now come from elements of the open model; tick **Include linked models** to add Revit links, and tick as many links as you need (all are ticked at first, **Pick Links** picks them in the view). With it off, links are not used.

### Removed
- **MCP point cloud tools** (`list_point_clouds`, `analyze_point_cloud`, `detect_point_cloud_elements`): no AI app called them since usage tracking began. Scan-to-BIM stays in the **Point Cloud to Model** button.

### Fixed
- **API key shared with t3lab-revit-api no longer fails with 401:** both builds keep keys in `%APPDATA%\T3LabAI\settings.json`. Lite 2.0.0 sent the encrypted `dpapi:` text to the vendor as the key and saved plain keys over encrypted ones. Lite now reads and writes the same Windows-encrypted keys (DPAPI, current Windows user) as t3lab-revit-api, and a key saved on another Windows account says "Saved key can't be read on this Windows account — paste it again" instead of being sent.
- **No silent switch to another AI vendor:** when the chosen provider fails, the Assistant shows the error instead of sending your Revit context to a provider you did not pick. A local model (Ollama, LM Studio) never falls back to a cloud one.
- **Timeouts and retries:** every request has a real timeout; 408 / 429 / 5xx / "overloaded" replies are retried twice with backoff (`Retry-After` honoured). **Stop** in the Assistant aborts the request at once, and the agent's time budget ends a hung request as a timeout.
- **Dropped streams:** an answer cut off mid-stream continues from where it stopped instead of starting again; an agent turn the server ended early shows the "dropped mid-request" warning with Retry.
- **T3Lab Assistant:** when its language module cannot load (a broken install), a message the Assistant cannot answer gets the setup guidance instead of an error bubble (`keyword_parse` was undefined there).
- **T3Lab Assistant:** opening the Assistant no longer logs "Local engine offer failed" every time (it called a method that never existed).
- **MCP:** the port set in MCP Control is used again (48884–48894); the Manual Config snippet is valid JSON and matches the app setup; `tools/list` no longer waits behind a slow tool call; a call that timed out says whether it may still finish; failed tools are flagged `isError`; Claude no longer needs a restart when Revit created the token after it.
- **ManaSched:** text cells with Excel's Accounting format no longer show a stray `_(` … `_)` around the text in Revit. `[Red]`-style colour codes and `\` escapes in text formats are no longer printed either.
- **ManaSched:** Excel / Google Sheets cell borders now show in a schedule. Grid lines stay on (Revit hides every header border when they are off), borders use the line style's category id, and a line Excel saved on only one of two neighbouring cells (e.g. a header row's bottom border) reaches both.

## [2.1.0] - 2026-10-19

ManaSched v2 and the latest T3Lab Assistant, copied from `t3lab-revit-api`, plus the PDF Markup split button.

### Added
- **ManaSched v2 — Excel / Google Sheets tables in Revit** (Annotation & Data › manaData):
  - Import an Excel range as a schedule, drafting view or legend with its formatting (fonts, colours, borders, merged cells, column widths and row heights) and keep it linked to the file.
  - Pick the sheet and the cells in an Excel-like viewer.
  - Paste a Google Sheets link (normal, Publish-to-web or bare id) or browse to a `.gsheet` shortcut. No sign-in: the sheet must be shared "Anyone with the link" or published. Downloads are cached in `%APPDATA%\T3LabAI\gsheets`.
  - **Linked Tables**: see which tables are out of date ("Sheet changed", "Offline"); Update Tables, reload from another file, or unlink.
  - Schedule data round-trip: export schedule values to Excel, edit them, import and review every change, then apply in one undo step.
- **Tool Hub in the Assistant**: sign in with the e-mail link or the 6-digit code, synced tool catalog with offline search, `/toolhub` and `/save-tool` commands, tool recipes, and a `find_tools` search the model can call. LLMs Setting has a new **Tool Hub** tab (sign in, sync, my tools, search in one window).
- **Assistant — model per role**: the newest model of each provider by default; auto model switch (Haiku pre-reads, Sonnet/Opus take complex turns); cost in USD per turn and a daily budget.
- **Assistant — per-chat memory** (Claude Projects-style) and a context window sized to the selected model.
- **What's new popup**: the first Revit start after an update shows what changed in the new version, once per version (after skipped updates, every release since the last one shown). "Open the full changelog" opens this file on GitHub.
- **PDF Markup split button** (Support › Cloud Links): **Bluebeam Status** now shares a split button with the new **T3Lab PDF Markup**, which opens [T3Lab PDF](https://pdf.t3lab.space/) (offline PDF markup for drawings) in the default browser. The button used last stays on top, also after Revit restarts (saved when Revit closes, in `%APPDATA%\T3LabAI\ribbon_state.json`).

### Changed
- **Assistant**:
  - History budget = model window minus the measured fixed part (system prompt + tools), not a fixed share.
  - Large tool results are capped (32k cloud / 4k local) and old results inside a turn are folded; the full result is kept on disk.
  - Cloud answers up to 16k/32k tokens, 10-minute stream timeout, and the agent loop can pause.
  - Claude tool search + `defer_loading` keeps the cloud tool list short.
  - The model chip follows provider changes made outside the pane; no offline answers when Claude is configured.
  - Pure UI updates are posted with `BeginInvoke`, so the pane no longer blocks Revit while it streams.
- **API keys** in `%APPDATA%\T3LabAI\settings.json` (and the Tool Hub session) are now encrypted with Windows DPAPI for the current user. Existing plain keys keep working.
- `Bluebeam Status` moved into `CloudLinks.stack/PDFMarkup.splitbutton`. A copy-over update moves the old `CloudLinks.stack/Bluebeam Status.urlbutton` folder to `_retired_ribbon` on the next Revit start.

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
