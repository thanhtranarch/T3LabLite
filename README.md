# T3Lab Lite

**T3Lab Lite** is a pyRevit extension built for Revit users who want to work faster.
It covers annotation management, batch export, family tools, and an AI assistant that understands Vietnamese.

---

#### Installation

T3Lab Lite is installed as a pyRevit extension.

    ▶ Install pyRevit

    ▶ Open Extensions Menu

    ▶ Add T3Lab Lite

    ▶ Click Install

---

### Tools

#### Annotation Manager
Find, delete, and rename Dimensions and Text Notes from a single window.

#### Dim Text / Upper Dim Text
Edit dimension text overrides and convert them to uppercase in one click.

#### Grids
Save and restore grid head/tail positions per view — useful before sharing or printing.

#### Reset Overrides
Remove all by-element graphic overrides and linework on selected elements in the active view.

#### BatchOut
Export sheets to PDF, DWG, DWF, DGN, IFC, NWD, and image formats in batch.
Supports combined PDF, custom naming patterns, sheet ordering, and revision tracking.

#### Load Family
Browse and load Revit families from local folders with category filtering and search.

#### JSON to Family
Build parametric Revit families from a JSON schema — parameters, reference planes, and geometry (Extrusion, Sweep, Revolve, Blend, Void). Must be run inside an open Family document.

#### Property Line
Create property lines from Lightbox parcel survey data.

#### Workset
Assign worksets to selected elements quickly.

#### T3Lab Assistant
An AI assistant built into Revit. Type commands in **Vietnamese or English** to open and control T3Lab tools.
Works with the Claude API (your own Anthropic key) or local Ollama for fully offline use.
Includes a local MCP server for AI-to-Revit communication.

---

### Network Traffic

All connections are **user-initiated**. Nothing runs on extension load.

| Component | Destination | When |
|---|---|---|
| T3Lab Assistant | `api.anthropic.com` | User sends a message with a Claude API key saved |
| Load Family (Cloud) | Author's Vercel deployment | User opens the cloud family catalogue |
| API Learner | `revitapidocs.com` | Refreshes BatchOut API data (30-day cache, Fridays only) |
| Ollama | `localhost:11434` | Local inference only — never external |
| MCP Server | `localhost:8080` | Only while MCP is running |
