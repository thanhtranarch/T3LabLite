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
   Sizes tab as "Parameter fit · <Category> · <score>%", every row with its
   state in words.
5. **Fix** – "Make it parametric" sends `param_profiles.brief(schema)` to Claude:
   only the sizes that do not resize yet, with their axes, asking for reference
   planes, face locks and labelled dimensions while keeping the current look.
   Every template brief already ends with its category's sizes
   (`template_note`).

## Trying sizes before Revit

In the Types tab, length / number values are editable (Enter, Tab, arrows ±10 mm,
Shift+arrows ±100 mm). The Studio stores the value in the type (or the default)
and re-draws the preview with `flex.py`: a labelled dimension moves its
non-template plane, planes dimensioned from it follow, and faces locked to a
moving plane stretch the part. Parts without locks stay put, as in Revit, and the
status line says when a parameter drives nothing.

## Next steps (not done yet)

- Instance vs type: mark which profile sizes should be instance parameters
  (e.g. Sill Height, Mounting Height) and grade that too.
- Shared parameters / IFC: map profile names to the office's shared parameter
  file so schedules and IFC export pick them up.
- Per-office profiles: read an optional JSON override in
  `%APPDATA%\T3LabAI\famigen_profiles.json`.
- Check formulas and min/max ranges (e.g. Door Width 600–1200).
- NEEDS VERIFICATION in Revit: that families built after "Make it parametric"
  flex the same way the preview does.
