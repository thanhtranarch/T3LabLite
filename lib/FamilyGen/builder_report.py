# -*- coding: utf-8 -*-
"""English summary of a FamilyGen.builder report (pure Python, no Revit).

Shared by the classic FamiGen dialog, FamiGen Studio and the MCP tool
``famigen_create_family``. ``FamilyGen.builder.report_lines`` re-exports it.
"""


def report_lines(report, limit=12):
    """English summary lines for a build report (what / where / next)."""
    lines = ["Built {} of {} part(s).".format(report.get('built', 0), report.get('total', 0))]
    if report.get('saved_path'):
        lines.append("Saved to: {}".format(report['saved_path']))
    if report.get('loaded'):
        lines.append("Loaded into the active project.")
    mats = report.get('materials_created', []) + report.get('materials_reused', [])
    if mats:
        lines.append("Materials: {} ({} new); {} material parameter(s)."
                     .format(len(mats), len(report.get('materials_created', [])),
                             len(report.get('material_parameters', []))))
    if report.get('parameters_created') or report.get('parameters_set'):
        lines.append("Parameters: {} created, {} set.".format(
            len(report.get('parameters_created', [])), len(report.get('parameters_set', []))))
    if report.get('subcategories'):
        lines.append("Subcategories: {}.".format(', '.join(report['subcategories'])))
    if report.get('reference_planes') or report.get('dimensions') or report.get('locks'):
        lines.append("Parametric: {} reference plane(s), {} labelled dimension(s), {} locked face(s)."
                     .format(len(report.get('reference_planes') or []),
                             len(report.get('dimensions') or []), report.get('locks') or 0))
    if report.get('types'):
        lines.append("Types: {}.".format(', '.join(report['types'])))
    for title, key in (("Skipped", 'skipped'), ("Warnings", 'warnings')):
        items = report.get(key) or []
        if items:
            lines.append("")
            lines.append("{} ({}):".format(title, len(items)))
            for item in items[:limit]:
                lines.append("  - {}".format(item))
            if len(items) > limit:
                lines.append("  ... and {} more".format(len(items) - limit))
    return lines
