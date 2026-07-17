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

### Removed
- **MCP**: `show_assistant_pane` tool — the T3Lab Assistant pane no longer
  appears in the MCP client's tool list (the pane itself was already retired
  in favour of the MCP-based workflow).

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
