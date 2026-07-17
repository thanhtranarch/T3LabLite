# T3Lab Lite

**T3Lab Lite** is a pyRevit extension built for Revit users who want to work faster.
It covers batch export, sheet & view management, family tools, IFC-SG compliance, CAD-to-BIM conversion, model auditing, and MCP integration that lets Claude AI work with Revit directly.

See [CHANGELOG.md](CHANGELOG.md) for what's new in each release.

---

#### Installation

T3Lab Lite is installed as a pyRevit extension.

    ▶ Install pyRevit

    ▶ Open Extensions Menu

    ▶ Add T3Lab Lite

    ▶ Click Install

---

### Tools

The ribbon is organised into six panels: **Support**, **Standards & Settings**, **Data & IFC-SG**, **Modeling & Datum**, **Annotation & Select**, and **Views & Sheets**.

---

### Support

#### MCP Control
Start and stop the local MCP server that lets Claude AI (and other MCP clients) interact directly with Revit.
Configure host, port, and authentication settings.

#### PDF Import
Import PDF pages into selected Revit views sequentially. Supports 150 / 300 / 600 DPI.

#### UI Theme & Tabs
- **ManaTabs** — hide or show Revit ribbon tabs to reduce clutter.
- **Ribbon Names** — shorten or restore ribbon tab names with inline editing and saved mappings.
- **BG Theme** — set the model-view background colour. Presets, RGB sliders, HEX input, and live preview. SHIFT+Click cycles Black → Gray → White.

#### Send Feedback
Send feedback or suggestions to the T3Lab team directly from Revit.

#### Cloud Links
Quick links to Autodesk Forma, Autodesk Health dashboard, and Bluebeam Status.

---

### Standards & Settings

#### Model Auditor
Consolidated model health checks in one window.
- **Model Check** — verify model standards and quality rules.
- **Warnings** — review and address the Revit warning list.
- **In-Place Models** — list and manage in-place family instances.
- **Material List** — audit all materials used in the model.

#### ManaLoca
List and adjust element locations in the current view or by level.
Edit XYZ coordinates in a data grid and commit changes back to the model. Modeless — stays open while you work.

#### ManaStyles
Unified visual style manager.
- **Fill Patterns / Line Styles / Line Patterns** — create, rename, and manage graphic styles.
- **Color Splasher** — apply graphic override colours to elements by category rule.
- **Coordinate Editor** — view and adjust element XYZ coordinates in a grid.

#### ManaWorkset
Enable worksharing, create or delete worksets, assign elements to worksets by rule (category, level, or type), and generate workset-based view filters.

---

### Data & IFC-SG

#### ManaSched
Export schedule data to Excel with formatting preserved, import updated values back into schedule rows, and duplicate schedules.

#### ManaPara
Parameter Manager — transfer parameter values between elements by rule, assign Text Note content to element parameters via spatial overlap, and write schedule values into filled region parameters.

#### ManaContains
Find elements contained in Rooms, Areas, Spaces, Zones, Masses, or Scope Boxes.
Assign parameter values to contained elements from their container, or aggregate element data back into the container.

#### IFC-SG Suite
Unified suite for IFC-SG submission in Singapore.
- **Subtype Assigner** — load mapping rules from Excel, assign IFC Export Class and Predefined Type parameters.
- **Compliance Checker** — verify required parameters exist and are filled based on CORENET X rules.

#### BCF Reader
Import BCF 2.1 files from IFC Delta Viewer and navigate issues in Revit.
Click an issue to jump the active view to the flagged element. Modeless — stays open while you work.

#### Foundation Volume
Write the Revit computed volume into a selected shared parameter on Structural Foundation elements, in one transaction.

---

### Modeling & Datum

#### Property Line
Create property lines from Lightbox parcel survey data. Supports metes-and-bounds descriptions and coordinate-based input.

#### Tile Layout
3-step wizard to extract floor boundaries, choose a tile pattern per floor, and place the generated tile layout on the active sheet.

#### Create Elements (CAD to BIM)
- **CAD to Elements** — convert CAD linework into Walls, Floors, or Beams by layer and colour mapping.
- **Room To Floor** — create architectural or structural floors from selected room boundaries.
- **Door Threshold** — create threshold floors at the base of selected doors, sized to the opening and host wall.
- **Point Cloud to Model** — Scan-to-BIM wizard that auto-detects Walls, Floors, Ceilings, Doors, Windows, Columns, Stairs, and Roof planes from a point cloud.
- **Image to Drafting** — create a Drafting View and import an image from disk or clipboard.
- **Text to Element** — transfer Text Note content to element parameters via bounding-box overlap in the active view.

#### ManaFami
Family Manager — browse loaded families by category, search and filter, load new families from disk or the cloud catalogue, and delete unused family types.

#### FamiGen
Create Revit families from external data.
- **From CAD** — scan imported DWG blocks and export each unique block as an `.rfa` family.
- **From JSON** — generate fully parametric families from a structured JSON schema.
- **Batch** — create standard families from built-in presets without any source file.

#### Element Adjust
- **Auto Join** — automatically join intersecting elements by category rules (Shift+Click for quick join).
- **Split Elements** — split Walls, Columns, or Floors at selected levels, preserving parameters.
- **Wall Cut Profile** — cut wall profiles or create openings based on intersecting linked model elements.
- **Auto Adjust Base Offset** — recalculate Base Offset when changing Base Constraint so elements keep their absolute elevation.

---

### Annotation & Select

#### ManaAnno
Unified tool to find, remove, and rename Dimensions and Text Notes.
Find dimensions or notes by keyword and jump to the view, delete instances or types, and auto-rename all types by their properties.

#### ManaDWG
Manage CAD imports and CAD links — list, rename, and delete DWG imports and links from a single interface.

#### Auto Dimension
Automatically create dimension chains for walls, columns, doors, lifts, and grids in the active or a chosen view.

#### ManaSelect
Smart selection manager with 4 modes: Quick Select (filter by parameter value or text), Select Similar (by type, family, or category), Select on Sheets (locate title blocks and CAD imports), and a Quick Filters sidebar.

---

### Views & Sheets

#### ManaViews
Browse and filter all views, rename views in bulk with naming rules, update view templates across multiple views, and remove unused views.

#### ManaSheets
Manage sheets in one unified interface — browse with live search, sync sheet data to/from Excel, place views on sheets, create sheet sets, and renumber sheets.

#### SheetGen
Create floor-plan views from a room list. Select rooms, choose a View Family Type and naming template, and generate all views in one transaction.

#### BatchOut
Export sheets to PDF, DWG, NWD, and IFC formats in batch.
Supports combined PDF, custom naming patterns, sheet ordering, and revision tracking.

---

### Network Traffic

All connections are **user-initiated**. Nothing runs on extension load.

| Component | Destination | When |
|---|---|---|
| MCP Server | `localhost:8080` (host/port configurable) | Only while the MCP server is running |
| ManaFami (Cloud) | User-configured Vercel URL in `~/.t3lab/family_loader_config.json` | User opens the cloud family catalogue |
| Feedback | None — opens a `mailto:` link in the default email client | User sends feedback |
| Cloud Links | `acc.autodesk.com` / `health.autodesk.com` / `status.bluebeam.com` | Opens in the default browser on click |
