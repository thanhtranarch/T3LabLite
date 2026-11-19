# FamiGen: the right parameters for each category

Code: `lib/FamilyGen/param_profiles.py` · preview flex: `lib/FamilyGen/flex.py` ·
tests: `lite_guard/test_famigen_studio.py` (`ProfileTests`, `FlexTests`).

## Why

A family is only useful in a project when the sizes people change are real
parameters that flex the geometry. Claude often returns a correct-looking model
with fixed sizes (no reference planes, no labelled dimensions), so changing
"Width" in Revit does nothing. The Studio now grades every family against what
its category needs and can ask Claude for exactly the missing sizes.

## How a family is graded

1. **Profile per category** – `PROFILES[category]` lists the *required* sizes
   (what people schedule, tag or flex) and *optional* ones, each with the axis it
   measures (X = left/right, Y = front/back, Z = up; `None` = diameter or count).
   Examples: Door → Width (X), Height (Z), Thickness (Y); Window → Width, Height,
   Sill Height; Pipe Accessory → Nominal Diameter, Length; Lighting Fixture →
   Diameter, Height. Every supported category has a profile (tested).
2. **Match** – a parameter matches by name or a known alias (`ALIASES`: "W",
   "Overall Width", "DN", "Face to face"…).
3. **State per size**
   - `driven`: the parameter exists and a labelled dimension on the right axis
     uses it → it resizes the family, in Revit and in the Studio preview;
   - `present`: the parameter exists but drives nothing;
   - `missing`: not in the family.
4. **Score** = (2 × driven + present) / (2 × required) → 0–100 %. Shown in the
   Sizes tab as "Parameter fit · <Category> · N of M sizes resize" (optional
   sizes a family does not use read "not used"), every row with its
   state in words.
5. **Fix** – "Make it parametric" sends `param_profiles.brief(schema)` to Claude:
   only the sizes that do not resize yet, with their axes, asking for reference
   planes, face locks and labelled dimensions while keeping the current look.
   Every template brief already ends with its category's sizes
   (`template_note`).

## Office profiles

An office can replace any category's profile with
`%APPDATA%\T3LabAI\famigen_profiles.json` (read on every check; a bad file is
ignored):

```json
{"Door": {"required": [["Width", "x", "rough width"], ["Height", "z", "leaf height"]],
          "optional": [["Frame Width", "x", "jamb"]]}}
```

## Trying sizes before Revit

In the Types tab, length / number values are editable (Enter, Tab, arrows ±10 mm,
Shift+arrows ±100 mm). The Studio stores the value in the type (or the default)
and re-draws the preview with `flex.py`: a labelled dimension moves its
non-template plane, planes dimensioned from it follow, and faces locked to a
moving plane stretch the part. Parts without locks stay put, as in Revit, and the
status line says when a parameter drives nothing.

## Type or instance

Every profile size has a recommended scope: **type** by default (every placed
copy shares it), **instance** where each placement differs - Window *Sill
Height*, *Mounting Height* (plumbing, electrical, specialty equipment),
*Suspension Length* (lighting). Parameter fit shows the scope of each size and
"→ instance" / "→ type" when the family has it the other way. An office profile
row may carry the scope as a 4th item: `["Sill Height", "z", "sill", "instance"]`.

## Shared parameters

Schedules, tags and IFC need the office's shared parameters, not same-named
family parameters.

- The Studio reads the shared parameter file Revit uses (`Manage > Shared
  Parameters`; the path is read when the Studio opens) as text -
  `FamilyGen/shared_params.py`, UTF-16 or UTF-8. Parameter fit shows the file
  and marks sizes that are *shared* or *shareable* (same name, compatible type).
- **Apply standards** (Sizes tab, no AI): sets each profile size to its
  recommended scope and marks every parameter found in the file
  `"shared": true` with its `"guid"`.
- `parameters[].shared` / `.guid` in the family JSON: the builder adds the
  `ExternalDefinition` (by GUID, else by name) with
  `FamilyManager.AddParameter(ExternalDefinition, GroupTypeId, isInstance)`
  (`Snippets/_compat.add_shared_family_parameter`, Revit 2022-2027). Not in
  the file -> a family parameter and a warning in the build report.

## Next steps (not done yet)

- IFC: map profile sizes to IFC property sets (e.g. `Pset_DoorCommon`).
- Check formulas and min/max ranges (e.g. Door Width 600–1200).
- NEEDS VERIFICATION in Revit: that families built after "Make it parametric"
  flex the same way the preview does.
