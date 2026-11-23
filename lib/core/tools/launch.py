# -*- coding: utf-8 -*-
"""Open a T3Lab ribbon window for the person at the machine (MCP clients; the
Assistant has its own launcher, open_t3lab_tool).

The button is found on the live ribbon (UIApplication.GetRibbonPanels; pyRevit
names an item after its bundle folder) and posted with PostCommand, so the
window opens after this call returns instead of blocking Revit's queue.
Doc: docs/agents/tool-specs.md (P2). Test: dev/test_tools_launch.py.
"""
from __future__ import unicode_literals

import difflib
import re

from core.tool_specs import ToolSpec, by_ribbon

_SUFFIXES = ('.pushbutton', '.smartbutton')


def _key(text):
    return re.sub(r'[^a-z0-9]', '', '{}'.format(text or '').lower())


def match_button(buttons, wanted):
    """(button dict, candidates): exact on folder / title / alias, else the
    single button containing the words; candidates when ambiguous or unknown."""
    want = _key(re.sub(r'\.(push|smart|url)button$', '', '{}'.format(wanted or '').strip()))
    if not want:
        return None, []

    def keys(b):
        stem = b['button'].rsplit('.', 1)[0]
        return [_key(stem), _key(b.get('title'))] + [_key(a) for a in b.get('aliases') or []]

    exact = [b for b in buttons if want in keys(b)]
    if len(exact) == 1:
        return exact[0], []
    partial = [b for b in buttons if any(want in k for k in keys(b) if k)]
    if len(partial) == 1:
        return partial[0], []
    pool = partial or buttons
    titles = sorted(set(b.get('title') or b['button'] for b in pool))
    near = difflib.get_close_matches('{}'.format(wanted), titles, n=8, cutoff=0.3)
    return None, near or titles[:8]


def command_ids(tab, path):
    """Command id strings an API ribbon button may have: the full path through
    pulldowns / splits, then tab + panel + button only."""
    out = ['CustomCtrl_%CustomCtrl_%' + '%'.join([tab] + path)]
    if len(path) > 2:
        out.append('CustomCtrl_%CustomCtrl_%' + '%'.join([tab, path[0], path[-1]]))
    return out


def _find_on_ribbon(uiapp, stem, tabs):
    """(tab, [panel name, ..., item name]) of the live item named `stem`."""
    def walk(items, trail):
        for item in items or []:
            try:
                name = item.Name
            except Exception:
                continue
            if name == stem:
                return trail + [name]
            try:
                children = item.GetItems()
            except Exception:
                children = None
            if children:
                found = walk(children, trail + [name])
                if found:
                    return found
        return None

    for tab in tabs:
        try:
            panels = uiapp.GetRibbonPanels(tab)
        except Exception:
            continue
        for panel in panels or []:
            try:
                found = walk(panel.GetItems(), [panel.Name])
            except Exception:
                found = None
            if found:
                return tab, found
    return None, None


def _post(uiapp, stem):
    """(True, how) once Revit queued the button's command, else (False, why)."""
    import ribbon_memory
    from Autodesk.Revit.UI import RevitCommandId
    tab, path = _find_on_ribbon(uiapp, stem, ribbon_memory.tab_names())
    if path is None:
        return False, 'the button is not on the ribbon (tab hidden or not loaded)'
    for text in command_ids(tab, path):
        try:
            cid = RevitCommandId.LookupCommandId(text)
            if cid is not None and uiapp.CanPostCommand(cid):
                uiapp.PostCommand(cid)
                return True, 'PostCommand'
        except Exception:
            continue
    return False, 'Revit did not accept a command for the button'


def open_tool_window(doc, args):
    from Services import tool_discovery
    buttons = [b for b in tool_discovery.scan_all_buttons()
               if b.get('kind') != 'url' and b['button'].endswith(_SUFFIXES)]
    button, candidates = match_button(buttons, args.get('tool'))
    if button is None:
        return {'error': 'No T3Lab window matches "{}".'.format(args.get('tool') or ''),
                'candidates': candidates}
    title = button.get('title') or button['button']
    panel = (button.get('panel') or '').rsplit('.', 1)[0]
    out = {'tool': title, 'where': 'T3Lab tab > {} > {}'.format(panel, title)}
    headless = by_ribbon().get(button['button'])
    if headless:
        out['headless_tools'] = headless
        out['hint'] = ('{} does the same without the window; prefer it unless the '
                       'person wants to see the window.').format(', '.join(headless))
    if args.get('dry_run'):
        out.update(dry_run=True, opened=False)
        return out
    from Autodesk.Revit.UI import UIApplication
    ok, how = _post(UIApplication(doc.Application), button['button'].rsplit('.', 1)[0])
    if not ok:
        out.update(opened=False, error='Could not open {} ({}). Ask the person to open it '
                                       'from {}.'.format(title, how, out['where']))
        return out
    out.update(opened=True, note='The window opens in Revit after this call; it does not '
                                 'wait for the person to finish.')
    return out


SPECS = [
    ToolSpec(
        'open_tool_window',
        ('Open a T3Lab ribbon tool window in Revit for the person at the machine (BatchOut, '
         'Rebar Wizard, ManaSheets...). Use it when the person wants the window or a step '
         'needs their eyes; the result names any tool that does the job headless.'),
        open_tool_window,
        properties={'tool': {'type': 'string', 'description': 'Button title or folder name.'},
                    'dry_run': {'type': 'boolean', 'description': 'Find it, do not open.'}},
        required=['tool'], effect='read', main_thread=True, placement='external',
        title_vi='Mở cửa sổ tool T3Lab', keywords=['mở tool', 'mở cửa sổ', 'open window'],
        needs_doc=True),
]
