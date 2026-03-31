# T3Lab Lite

A pyRevit extension for Autodesk Revit that adds tools for annotation management,
batch export, project utilities, and AI-assisted automation.

**Author:** Tran Tien Thanh
**Contact:** trantienthanh909@gmail.com
**LinkedIn:** [linkedin.com/in/sunarch7899](https://linkedin.com/in/sunarch7899/)

---

## Requirements

- Autodesk Revit 2020 or later
- [pyRevit](https://github.com/eirannejad/pyRevit) 4.8+

---

## Installation

The repository root **is** the extension folder. pyRevit requires the folder to be
named `T3LabLite.extension`.

**Option A — git clone (recommended)**

```
git clone https://github.com/thanhtranarch/T3LabLite "%APPDATA%\pyRevit\Extensions\T3LabLite.extension"
```

**Option B — ZIP download**

1. Download and extract the ZIP
2. Rename the extracted folder to `T3LabLite.extension`
3. Move it to `%APPDATA%\pyRevit\Extensions\`

Then reload pyRevit: **pyRevit tab → Reload**.

---

## Tools

All tools appear in the **T3LabLite** ribbon tab.

### Annotation Panel

| Tool | Description |
|------|-------------|
| **Annotation Manager** | Unified window to find, delete, and rename Dimensions and Text Notes. Dimension tab: find by type keyword, delete instances, auto-rename types. Text Note tab: find by content, delete types, auto-rename types. |
| **Dim Text** | Edit dimension text overrides on selected dimensions |
| **Upper Dim Text** | Convert dimension text overrides to uppercase |
| **Save Grids** | Save current grid head/tail positions for later restoration |
| **Restore Grids** | Restore selected grids to their saved positions |
| **Restore All Grids** | Restore all grids to their saved positions |
| **Reset Overrides** | Reset all by-element graphic overrides and linework on selected elements in the active view |

---

### Export Panel

| Tool | Description |
|------|-------------|
| **BatchOut** | Batch export sheets to PDF, DWG, DWF, DGN, IFC, NWD, and image formats. Supports combined PDF, sheet ordering, custom naming patterns, and revision tracking. |

---

### Project Panel

| Tool | Description |
|------|-------------|
| **Load Family** | Browse and load Revit families from local folders with category filtering and search. Optional cloud catalogue (see Network Traffic). |
| **JSON to Family** | Generate parametric Revit families from a JSON schema — creates parameters, reference planes, and geometry (Extrusion, Sweep, Revolve, Blend, Void). Must be run inside an open Family document. |
| **Property Line** | Create property lines from Lightbox parcel data |
| **Workset** | Assign worksets to selected elements |

---

### AI Connection Panel

| Tool | Description |
|------|-------------|
| **T3Lab Assistant** | Natural-language assistant — type commands in Vietnamese or English to open and control T3Lab tools. Includes MCP server controls (Start / Stop) and a Settings button to configure the AI backend. |

Two AI backends are supported (configured via Settings inside the assistant):

- **Claude API** — sends messages to `api.anthropic.com` using your own Anthropic API key
- **Local LLM** — uses [Ollama](https://ollama.com/) running on `localhost:11434`, fully offline

---

## Project Structure

The repository root is the extension folder itself:

```
T3LabLite.extension/          ← this repository's root
├── extension.json            # pyRevit extension manifest
├── T3LabLite.tab/            # Ribbon tab definition
│   ├── Annotation.panel/
│   ├── Export.panel/
│   ├── Project.panel/
│   └── AI Connection.panel/
├── checks/                   # Model quality check scripts
├── commands/                 # Standalone command scripts
└── lib/                      # Shared Python libraries
    ├── GUI/                  # WPF dialogs (.xaml + .py)
    ├── Renaming/             # Find & replace base classes
    ├── Selection/            # Element selection utilities
    ├── Snippets/             # Reusable Revit API snippets
    ├── config/               # Settings and learned patterns
    ├── core/                 # Local MCP server
    └── ui/                   # Button state and settings dialog
```

---

## Network Traffic

All outbound connections are **user-initiated or opt-in**. Nothing runs automatically
on extension load.

| File | Destination | Trigger |
|------|-------------|---------|
| `lib/t3lab_assistant.py:362` | `https://api.anthropic.com/v1/messages` | User sends a message in T3Lab Assistant **and** has saved a Claude API key in Settings. The user's own key is used; no data is stored by T3Lab. |
| `lib/GUI/FamilyLoaderCloudDialog.py:84` | `https://t3stu-...vercel.app/api/families` | User opens the **Load Family (Cloud)** dialog. Fetches a family catalogue JSON from the author's Vercel deployment. A Vercel deployment-protection bypass token is hardcoded in the source at line 90. |
| `lib/GUI/FamilyLoaderCloudDialog.py:275` | Same Vercel deployment | User selects and loads a cloud family. Downloads the `.rfa` file to a local temp folder. |
| `lib/api_learner.py:64` | `https://www.revitapidocs.com/{version}/` | When BatchOut's SmartAPIAdapter refreshes its API compatibility cache (30-day TTL). Reads public documentation pages; no user data is sent. |
| `lib/api_updater.py:63` | `https://www.revitapidocs.com` | Same cache refresh mechanism, also scheduled on Fridays. |
| `lib/local_llm.py:66` | `http://localhost:11434` | Ollama backend only — **local machine only**, never leaves the host. |
| `lib/core/server.py:452` | Listens on `localhost:{port}` (default 8080) | User clicks **Start MCP** inside T3Lab Assistant. Binds to localhost only. Note: `Access-Control-Allow-Origin: *` is set, so any local browser tab can connect to it while the server is running. |

No telemetry, analytics, or automatic update checks run on startup.

---

## License

MIT License — see [LICENSE](LICENSE) for details.

For bugs or feature requests, please [open an issue](https://github.com/thanhtranarch/T3LabLite/issues).
