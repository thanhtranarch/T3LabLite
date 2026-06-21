# T3Lab Lite

**T3Lab Lite** is a pyRevit extension built for Revit users who want to work faster.
It covers batch export, family tools, IFC-SG compliance, location management, and an AI assistant that understands Vietnamese.

---

#### Installation

T3Lab Lite is installed as a pyRevit extension.

    ▶ Install pyRevit

    ▶ Open Extensions Menu

    ▶ Add T3Lab Lite

    ▶ Click Install

---

### Tools

#### BatchOut
Export sheets to PDF, DWG, NWD, and IFC formats in batch.
Supports combined PDF, custom naming patterns, sheet ordering, and revision tracking.

#### BCF Reader
Import BCF files from IFC Delta Viewer and navigate issues in Revit.
Click an issue to jump the active view to the flagged element. Modeless — stays open while you work.

#### IFC-SG Suite
Unified suite for IFC-SG submission in Singapore.
- **Subtype Assigner** — load mapping rules from Excel, assign IFC Export Class and Predefined Type parameters.
- **Compliance Checker** — verify required parameters exist and are filled based on CORENET X rules.

#### Load Family
Browse and load Revit families from local folders with category filtering and search.

#### Bulk Family Export
Scan imported DWG/DXF files for block references and export each unique block as a separate `.rfa` family file.

#### JSON to Family
Build parametric Revit families from a JSON schema — parameters, reference planes, and geometry (Extrusion, Sweep, Revolve, Blend, Void). Must be run inside an open Family document.

#### Property Line
Create property lines from Lightbox parcel survey data.

#### Workset
- **Workset Management** — rule-based automatic workset assignment.
- **Workset Views** — create a view filtered to a single workset.
- **Click to Central** — create a new central file from a local model.

#### Mana Locations
List and adjust element locations in the current view or by level.
Edit XYZ coordinates in a data grid and commit changes back to the model. Modeless.

#### PDF Import
Import PDF pages into selected Revit views sequentially. Supports 150 / 300 / 600 DPI.

#### MCP Control
Start and stop the local MCP server that lets Claude AI interact directly with Revit.
Configure host, port, and authentication settings.

#### T3Lab Assistant
An AI assistant built into Revit. Type commands in **Vietnamese or English** to open and control T3Lab tools.
Works with the Claude API (your own Anthropic key) or local Ollama for fully offline use.

#### UI Utilities
- **BG Theme** — set the model-view background colour. Presets, RGB sliders, HEX input, and live preview.
- **ManaTabs** — hide or show Revit ribbon tabs to reduce clutter.
- **Ribbon Names** — shorten or restore ribbon tab names.

#### Send Feedback
Send feedback or suggestions to the T3Lab team directly from Revit.

#### Cloud Links
Quick links to Autodesk Forma, Autodesk Health dashboard, and Bluebeam Status.

---

### Network Traffic

All connections are **user-initiated**. Nothing runs on extension load.

| Component | Destination | When |
|---|---|---|
| T3Lab Assistant | `api.anthropic.com` | User sends a message with a Claude API key saved |
| Load Family (Cloud) | Author's Vercel deployment | User opens the cloud family catalogue |
| MCP Server | `localhost:8080` | Only while MCP server is running |
| Ollama | `localhost:11434` | Local inference only — never external |
