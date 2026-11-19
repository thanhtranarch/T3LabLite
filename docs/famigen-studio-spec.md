# FamiGen Studio — specification

Status: implemented 2026-11 (Studio 1.0). Requires **pyRevit 7.0+**; two-way page
messages need **pyRevit 7.1+**. Owner: T3Lab. This page is the contract for the
code under `lib/FamilyGen/` and for the `famigen_*` MCP tools. The family JSON
contract itself is in [famigen-ai.md](famigen-ai.md) and
`lib/Intelligence/family_schema.py`.

## 1. Goal

A BIM modeller gets a correct, editable Revit family by chatting: describe the
object (or drop a photo, drawing or PDF cut sheet), review it in 3D with its
materials, types and sizes, then create the `.rfa`. No JSON on screen, no
add-in to compile.

| # | Requirement | Where |
|---|---|---|
| R1 | Chat with Claude directly (user's own API key from LLMs Setting) | `claude_designer.py`, `net_http.py` |
| R2 | Or work from an AI app over MCP (Claude Desktop / Code, ChatGPT Codex, …): Studio switches by itself | `studio_link.py`, server `famigen_*` tools |
| R3 | Auto-detect the family category, or let the user pin one | `family_schema.CATEGORY_GUIDE`, `get_category_spec` tool |
| R4 | Follow a reference image / PDF closely; compare a render with it once | `begin_check`, snapshot round |
| R5 | Complex families: repeated members, detail levels, voids | `repeat`, `detail_levels` |
| R6 | Families that flex: reference planes, labelled dimensions, face locks | `reference_planes`, `dimensions`, `locks` |
| R7 | Family types: size / finish variants | `types[]` |
| R11 | Models read as ONE built object: continuous curves, no floating parts | `Tube`, `spatial_check`, polish round |
| R12 | Models make sense in space: the AI states the concept, each part's role, what is hollow and how parts join; FamiGen checks it | `concept`, `role`, `wall`, `joints`, `spatial_check` (§5.3) |
| R8 | Revit-like review: Coarse / Medium / Fine, Wireframe / Hidden Line / Shaded / Consistent Colors / Realistic, 3D / Front / Right / Plan | `studio.js` |
| R9 | Same look as every T3 tool: ONE T3 title bar (pyRevit's own is hidden, or the page hides its own), footer bar, window border | `studio.css` (T3 tokens), `studio_host._go_frameless` |
| R10 | Never change the model before the user presses Create Family | controller + builder |

Non-goals (1.0): MEP connectors, nested families, formulas, arrays driven by a
count parameter, editing an existing `.rfa` (only "add to the open family").

## 2. User flows

1. **Describe → review → create.** Type in the chat (or pick a template), Send.
   Claude answers in one or two sentences plus its assumptions; the 3D view,
   Parts, Materials, Types, Sizes and Checks fill in. Create Family saves
   `<family_name>.rfa` to the chosen folder (optionally loads it).
2. **From a reference.** Attach, paste (Ctrl+V) or drop images, PDFs or text
   files. With "Compare with my reference" on, Studio renders the model (3D +
   front) and Claude compares it with the reference once, fixing what differs.
3. **Refine.** "Make the legs Ø50", "add a Walnut type", "make it parametric:
   Width, Depth, Height". Claude changes only what was asked.
4. **AI app over MCP.** While an AI app uses the `famigen_*` tools, Studio shows
   "MCP · <app>", messages typed in Studio go to the app's inbox, and every
   family the app proposes appears in Studio at once. The built-in Claude uses
   no tokens meanwhile. "Send to" (Auto / Built-in Claude / AI app) overrides.
5. **History.** My families lists the last 50 families from Claude, the T3Lab
   Assistant and MCP apps; one click reopens one.

## 3. Architecture

```
 Studio page (WebView2, three.js r158)          T3Lab MCP server (IronPython host)
   │  chrome.webview.postMessage                   │  famigen_get_studio_state / validate /
   ▼                                               │  propose / patch / create / list_templates
 studio_host.py ── output.on_message ──┐           │
   (pyRevit 7 output window)           ▼           ▼
                               studio_controller.py ◄── studio_link.py files ──► server
                                 │        │        (presence, activity, state, inbox)
                    claude_designer.py  preview_mesh.py        ProposalStore (proposals)
                         │                    │
                   net_http.py (.NET HttpClient, polled by DispatcherTimer)
                         │
                  api.anthropic.com/v1/messages
                                             builder.py (Revit API, ExternalEvent)
```

| Module | Engine | Revit API | Tests |
|---|---|---|---|
| `studio/studio.html, .css, .js`, `vendor/three.min.js` | WebView2 | no | `test_famigen_studio.PageTests` |
| `studio_page.py` (one-file page, version gate) | any | no | `PageTests` |
| `studio_controller.py` (protocol, preview payload) | CPython | no | `ControllerTests`, `PreviewTests`, `McpModeTests` |
| `claude_designer.py` (Claude session) | any | no | `test_famigen_claude.py` |
| `studio_link.py` (files shared with the server) | both | no | `McpModeTests` |
| `templates.py` (40 briefs, 9 groups) | any | no | `SchemaExtrasTests` |
| `net_http.py` (.NET HTTP) | CPython | no | needs Revit |
| `studio_host.py` (pyRevit glue, ExternalEvent) | CPython | yes | needs Revit |
| `builder.py` (+ types, repeat, constraints, detail levels) | any | yes | needs Revit |

### Threading (CLAUDE.md rule 8)

No CPython thread is started. HTTP runs in .NET; a `DispatcherTimer` (100 ms)
polls the Task on the UI thread, which already holds the GIL. The host's 1 s
watch timer calls `controller.tick()`. With a persistent engine the Studio is
modeless and Create Family goes through an `ExternalEvent`; without one, the
command stays alive in a nested `DispatcherFrame` (Revit's window disabled like
a modal dialog) and builds directly.

## 4. The Claude session (`claude_designer.py`)

* Request: `model` (from the live `/v1/models` list, quality role = Opus first,
  never hard-coded), `max_tokens` 32 000, `system` (contract + workflow +
  category spec or category guide), `tools` = `submit_family` (+
  `get_category_spec` in Auto-detect), `tool_choice: auto` (forced choice is a
  400 on Opus 5.5 / Sonnet 5.5 / Fable 5.1+), `output_config.effort`
  (Fast = low, Balanced = medium, Best = high, Deep = xhigh), server-side
  `fallbacks: "default"` (beta header), top-level `cache_control`.
* Optional features refused with a 400 are dropped for the session and the
  request is retried (older models / accounts keep working).
* Validate → repair: errors go back as an `is_error` tool_result naming every
  JSON path; at most 2 repair rounds, then the last draft is shown with its
  problems. `max_tokens` cut → one "be compact" round. `refusal` → plain message.
* The transcript is append-only (preserved thinking): assistant content is kept
  verbatim, every `tool_use` gets its `tool_result` first in the next user turn.
* Editor context: when the family on screen differs from Claude's last one
  (opened from history, MCP, a template), it is sent as "Current family JSON".
* Reference analysis: with a reference attached Claude must first call
  `record_analysis` (camera azimuth/elevation in FamiGen axes, view type,
  scale basis, overall size, every component with form, size, position,
  colour and count). Studio shows it as a "What Claude sees" card and adds a
  **Ref** view with that camera.
* Reference check: `begin_check(renders)` sends renders from the **reference
  camera** and the 3D view; Claude either submits a fix or answers
  "Matches …". Up to 2 rounds per user message.
* Attachments: PNG/JPG/GIF/WEBP ≤ 5 MB, PDF ≤ 32 MB, TXT/MD/CSV/JSON/XML ≤ 400 KB;
  at most 8 per message.

## 5. Schema v2 additions (all optional, v1/v2 files stay valid)

| Field | Shape | Builder |
|---|---|---|
| `types[]` | `{"name", "values": {param: value, material param: material}}` | first = current type renamed, others `NewType` |
| `SweptBlend` (form) | `{"path": [one Line/Arc3P], "profile", "top_profile"}` (local sections) | `NewSweptBlend` |
| `Sweep.profile_local` | `true` → profile drawn as the section seen along the path | profile placed at the path start |
| `Tube.bend_radius` | straight runs + true circular elbows of that centre-line radius (pipes, taps, handrails) | fillet arcs sampled every 12° |
| `Tube` (form) | `{"points": [[x,y,z], …], "radius", "smooth"?, "closed"?}` | written out as spline segments (Cylinders lengthened by r·tan(turn/2) so bends close) + a ball only at turns > 35° |
| `geometry[].repeat` | `{"count", "step":[dx,dy,dz]}` or `{"at":[[dx,dy,dz],…]}` (not Sweep, ≤ 200) | written out by `expand_schema` before building (preview does the same) |
| `geometry[].detail_levels` | `["coarse","medium","fine"]` subset | `FamilyElementVisibility` |
| `reference_planes[]` | `{"name", "axis": x/y/z, "offset" mm}` | `NewReferencePlane` (plan / front elevation) |
| `dimensions[]` | `{"parameter" (length), "from", "to"}` | `NewDimension` + `FamilyLabel` |
| `geometry[].locks` | `{x_min…z_max: plane}` | `NewAlignment` + `IsLocked`, only faces on the plane |
| categories | + Pipe Accessory, Duct Accessory, Sprinkler, Electrical Fixture (17) | missing template → Generic Model, re-categorised |

Template planes usable by name: `Center (Left/Right)`, `Center (Front/Back)`,
`Ref. Level`. Constraint failures are warnings, never a failed build. The
preview warns before Revit when a locked face does not lie on its plane.

### 5.2 The whole Revit form toolbox

| Revit tool | Schema | Typical use | Builder |
|---|---|---|---|
| Extrusion | `Extrusion` (+ `sketch_plane_x/y` for elevation outlines) | slabs, panels, gables, stringers | `NewExtrusion` |
| Blend | `Blend` (`profile` → `top_profile`) | tapers, hoppers, chamfered plinths, cones | `NewBlend` |
| Revolve | `Revolution` (axis, angles) | domes, bowls, knobs, balusters, rings | `NewRevolution` |
| Sweep | `Sweep` + `"profile_local": true` | handrails, mouldings, kerbs, slide troughs | `NewSweep` (profile placed at the path start) |
| Swept Blend | `SweptBlend` (ONE Line/Arc3P path, start + end local sections) | chutes, horns, transitions | `NewSweptBlend`; sections at the path ends, retried in the XY plane |
| Void Forms | `"is_solid": false` on any form | openings, cavities, hollow tubes, holes | void form + `SolidSolidCutUtils.AddCutBetweenSolids` with every solid it overlaps (best effort) |
| (FamiGen) | `Cylinder`, `Tube` | straight / curved rods | expanded to Cylinders (+ balls) |

Local cross-sections (`FamilyGen/form_frames.py`): at a path point with tangent
`t`, `u = Z × t` (sideways), `v = t × u` (up); `[u, v, 0]` in mm. Circles and
ellipses become Arc3P pieces so Revit gets plain world curves. The contract's
FORMS block tells Claude which tool fits which shape; Studio hides voids unless
**Voids** is ticked (the preview does not subtract them; Revit does).

### 5.1 Continuity (why models no longer look "in pieces")

| Cause of a disjointed model | Rule / mechanism |
|---|---|
| Curved members drawn as chained straight Cylinders | `Tube`: one centre line, spline-smoothed, joints closed |
| Parts placed with approximate coordinates, leaving gaps | Contract CONTINUITY rules: child starts at the parent's face coordinate; 1–5 mm overlap |
| Parts floating (no support) | `spatial_check`: parts whose SURFACES do not touch anything that reaches the lowest part → "Not connected: 'Lamp' …" warning (§5.3; bounding boxes are no longer enough) |
| Claude never told | `claude_designer`: a valid family with spatial findings gets up to `MAX_POLISH` (2) polish rounds (non-error tool_result naming them) before the user sees it |
| Faceted look of round parts | Studio shades Cylinder / Revolution / Sweep smoothly, prismatic parts keep crisp edges |
| Missing category know-how | Every category spec ends with "JOINTS & SMOOTH FORMS" (category-specific joint recipes) |

The check uses bounding boxes (5 mm contact tolerance): it catches floating
parts, not every hairline gap, and a diagonal member counts as touching what
its box overlaps. It runs in every preview, `famigen_validate_family` and
`famigen_propose_family` warnings.

### 5.3 Meaning: concept, roles, solid or hollow, joints, space

A model can be valid JSON, touch everywhere, and still be nonsense: a tube slide that
starts in mid-air beside its platform, a stair that ends under a floor, a basin modelled
as a block. The schema therefore carries what the family MEANS, and
`FamilyGen/spatial_check.py` checks the geometry against it.

| Field | What it says | Checked |
|---|---|---|
| `concept` `{what, use, placement}` | What the object is, who or what uses it and by which route, where it stands or hangs. Required from an AI (`meaning_errors`); optional in hand-written JSON. | Presence only. It is the first thing the model writes, so the geometry is derived from it. |
| `geometry[].role` | `support`, `surface`, `enclosure`, `container`, `conduit`, `access`, `guard`, `fixing`, `control`, `decor` | `container` / `conduit` must be hollow → **Hollow**. `access` needs a joint at each end → **Access**. |
| `wall` on `Cylinder` / `Tube` | Hollow pipe: inner radius = radius − wall. Preview shows the bore; the builder makes an extrusion with an inner loop (axis-aligned) or revolves the wall section (diagonal). Joint balls of a hollow Tube are shells. | `wall` < `radius`. |
| `joints[]` `{a, b, at}` | Part `a` meets part `b` (or `"ground"`) at point `at`. | Both parts within 25 mm of `at` (a point inside a part counts) → **Joint**, with the distance missed. |

Findings (`spatial_check.PREFIXES`), all added to the preview warnings, so the Issues tab,
the Claude polish round and `famigen_validate_family` show the same list:

| Finding | Rule |
|---|---|
| **Not connected** | Solids must reach the base through parts that really touch: vertex-to-surface distance ≤ max(10 mm, 0.2 % of the model), then parts set INTO another solid (glass in a wall), then faces that cross without a vertex near (a plank across joists). |
| **Joint** | A declared joint where one part is not at the point; or names a part that does not exist. |
| **Hollow** | A `container` / `conduit` that is a plain solid: no `wall`, no `inner_loops`, no void overlapping it, and (Revolution) a profile that stays on the axis. |
| **Access** | An `access` part named in fewer than two joints. |
| **Crossing** | Something inside the bore of a `conduit` or a conduit centre line inside a solid (pipe ends excepted); something in the 150–1400 mm above an `access` surface. Skipped for jointed pairs, for `fixing` / `control` / `guard` parts, and where a void opens the way. |

Prompt: `build_system_prompt` blocks MEANING, SOLID OR HOLLOW, JOINTS, SPACE; the session
workflow adds THINK IN THIS ORDER (concept → levels and footprints → roles → joints →
geometry); `record_analysis` gains `concept`, `levels`, `connections`, `hollow`; every
category spec ends with "MEANING, SOLID / HOLLOW, JOINTS" (what is hollow in THAT category,
which joints to declare).

Limits: mesh based (48-gon circles, sampled splines); inside tests exist for Extrusion,
Cylinder and full Revolution only; the Studio preview does not cut voids, the check does
account for them. It finds what is wrong in the model; it does not prove Revit can build it.
Cost: about 1.5 s for a 360-part / 90 000-triangle model, cached per geometry.

## 6. Page ↔ Python messages

Page → Python (`type`): `ready`, `send {text, json}`, `cancel`, `new_chat`,
`attach`, `attach_data {name, media_type, data}`, `remove_attachment {id}`,
`set_category`, `set_model`, `set_effort`, `set_check {on}`, `set_engine
{engine}`, `set_load`, `pick_folder`, `create {json, load}`, `json_changed`,
`history`, `load_history {id}`, `snapshot {token, views, images}`, `select
{label}`, `copy_prompt`, `window {action}`, `close`.

Python → page: `init`, `busy {on, text}`, `chat {role, text, assumptions?,
files?, thumbs?, meta?}`, `chat_reset`, `schema {json, source}`, `preview
{parts, materials, planes, dimensions, table, type_materials, issues, …}`,
`usage`, `attachments`, `status {text, tone}`, `folder`, `category`,
`detected`, `engine {pref, effective, client}`, `history`, `snapshot_request
{token, views}`, `created {lines, path}`, `error`.

`test_messages_have_handlers_on_both_sides` fails when one side gains a type
the other does not handle.

## 7. MCP: AI apps working with Studio

Files in `%APPDATA%\T3LabAI\famigen_studio` (`studio_link.py`):

| File | Writer → reader | Rule |
|---|---|---|
| `presence.json` | Studio → server | heartbeat every 2 s; open when < 10 s old |
| `mcp_activity.json` | server → Studio | every `famigen_*` call; app "connected" for 10 min |
| `state.json` | Studio → server | category, family on screen, proposal id, selected part |
| `inbox.json` + `inbox_files/` | Studio → server | messages typed in Studio in AI-app mode |

| Tool | Purpose | Token note |
|---|---|---|
| `famigen_get_schema {category?, include_guidance?, compact?}` | contract | `compact` drops the example |
| `famigen_get_studio_state {include_family?}` | what the user sees + unread messages | read first on "continue in FamiGen" |
| `famigen_validate_family {schema}` | errors / warnings / size, nothing stored | cheap loop before proposing |
| `famigen_propose_family {schema, note?}` | store + show (Studio if open, else the classic review window) | |
| `famigen_patch_family {proposal_id?, set?, parts?, note?}` | edit by part id, top-level fields | sends only the change |
| `famigen_list_templates {query?, limit?}` | template briefs | |
| `famigen_create_family {proposal_id \| schema, …}` | build + save `.rfa` | only after the user approves |

Recommended loop for an AI app: `get_studio_state` → (`list_templates`) →
`get_schema {compact}` → `validate_family` until clean → `propose_family` →
user reviews in Studio → `patch_family` for each change → user presses Create
Family (or the app calls `create_family` after an explicit yes).

## 8. T3Lab Assistant integration — plan

Today: the Assistant can call every `famigen_*` tool (family specialist), and
its proposals land in the shared store, so an open Studio shows them at once.

Next steps, in order:

1. **Handoff button.** In the Assistant, a "Open in FamiGen Studio" chip under
   a family answer: opens Studio on that proposal (ribbon command id +
   `load_history`).
2. **Assistant as an engine.** "Send to: T3Lab Assistant" in Studio: messages
   go to the Assistant's queue (same inbox file, a second reader), so users with
   a non-Claude provider (OpenAI, Gemini, local) can drive Studio too.
3. **Shared conversation.** Store Studio conversations next to Assistant
   projects so a project's memory ("our casework is 18 mm melamine") reaches
   the family prompt.
4. **One window.** Long term, host Studio inside the Assistant dock pane as a
   tab once pyRevit exposes a WebView2 dockable pane.

## 9. Category specs

`FamiGen.pushbutton/prompts/<category>.md` exists for all 17 categories:
anatomy, typical sizes (e.g. EN 1092-1 PN16 flange diameters for valves),
origin and axes, recommended forms, parameters and types, detail levels,
pitfalls. Fixed category → the spec goes into the system prompt; Auto-detect →
Claude fetches it with `get_category_spec` after choosing the category.

## 10. Verification

Automated (Linux, no Revit): `python lite_guard/test_famigen_claude.py`,
`python lite_guard/test_famigen_studio.py`, plus a Chromium run of the real
page against the real controller with a scripted Claude (light and dark).

Needs Revit (not verifiable here): WebView2 message round trip in pyRevit
7.1, `.NET HttpClient` through the user's proxy, modeless vs modal Studio,
ExternalEvent create, `NewType` / `RenameCurrentType`, reference planes,
labelled dimensions and face locks flexing per type, detail-level visibility,
the four new MCP categories' templates per Revit language, title-bar buttons
reaching the output window (`_runtime_output().window`).

## 11. Limits

* pyRevit < 7.0: Studio refuses to open and offers the classic window
  (Shift+Click opens it directly).
* Constraints lock planar faces only (not cylinder sides); repeated copies are
  not locked.
* Claude cost is the user's own API usage; the AI badge tooltip shows tokens
  and an estimate. AI-app mode uses the app's plan instead.
