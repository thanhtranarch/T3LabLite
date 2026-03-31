# T3Lab Lite

Extension for [pyRevit](https://github.com/eirannejad/pyRevit) — tools for annotation, batch export, project utilities, and AI-assisted automation inside Autodesk Revit.

## Installation

* Install [pyRevit](https://github.com/eirannejad/pyRevit) 4.8+
* Clone this repo **as** `T3LabLite.extension` into your pyRevit extensions folder:
  ```
  git clone https://github.com/thanhtranarch/T3LabLite "%APPDATA%\pyRevit\Extensions\T3LabLite.extension"
  ```
  Or download the ZIP, extract, rename the folder to `T3LabLite.extension`, and place it in `%APPDATA%\pyRevit\Extensions\`
* Reload pyRevit

> **Note:** The repository root **is** the extension folder — there is no wrapper subdirectory.

## Tools

#### Annotation Manager
* Unified window to find, delete, and rename Dimensions and Text Notes
* Dimension tab: find by type name, delete instances, auto-rename dimension types
* Text Note tab: find by content, delete types, auto-rename text note types

#### Dim Text / Upper Dim Text
* Edit and uppercase dimension text overrides on selected dimensions

#### Grids
* Save and restore grid head/tail positions per view (Save, Restore, Restore All)

#### Reset Overrides
* Resets all by-element graphic overrides and linework on selected elements in the active view

#### BatchOut
* Batch export sheets to PDF, DWG, DWF, DGN, IFC, NWD, and image formats
* Supports combined PDF, sheet ordering, custom naming patterns, and revision tracking

#### Load Family
* Browse and load Revit families from local folders with category filtering and search
* Optional cloud catalogue (connects to author's Vercel deployment — see Network Traffic)

#### JSON to Family
* Generate parametric Revit families from a JSON schema
* Creates parameters, reference planes, and geometry (Extrusion, Sweep, Revolve, Blend, Void)
* Must be run inside an open Family document

#### Property Line
* Create property lines from Lightbox parcel survey data

#### Workset
* Assign worksets to selected elements

#### T3Lab Assistant
* Natural-language AI assistant — type commands in Vietnamese or English to open and control T3Lab tools
* Supports Claude API (your own Anthropic key) or local [Ollama](https://ollama.com/) for offline inference
* Includes a local MCP server (Start / Stop) for AI-to-Revit communication

## Network Traffic

All connections are **user-initiated**. Nothing runs on extension load.

* `lib/t3lab_assistant.py` → `https://api.anthropic.com/v1/messages` — only when the user sends a message and has saved a Claude API key in Settings
* `lib/GUI/FamilyLoaderCloudDialog.py` → author's Vercel deployment — only when the user opens Load Family (Cloud) to fetch a family catalogue or download a `.rfa` file. A Vercel bypass token is hardcoded in the source at line 90.
* `lib/api_learner.py` + `lib/api_updater.py` → `https://www.revitapidocs.com` — reads public API docs to keep BatchOut's export calls compatible across Revit versions (30-day local cache, refreshes on Fridays)
* `lib/local_llm.py` → `http://localhost:11434` — Ollama only, local machine, never external
* `lib/core/server.py` → listens on `localhost:8080` — only while MCP is running; note `Access-Control-Allow-Origin: *` is set

## Credits

* [Ehsan Iran-Nejad](https://github.com/eirannejad) for pyRevit
* Reset Overrides script based on original by Daria Ivanciucova
