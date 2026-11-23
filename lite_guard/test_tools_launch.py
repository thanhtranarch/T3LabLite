# -*- coding: utf-8 -*-
"""core/tools/launch.py: open_tool_window finds the ribbon button and posts it.

Run:  python lite_guard/test_tools_launch.py      (fake ribbon + Revit UI API)
"""
import os
import sys
import tempfile
import types

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_launch_test_')

from core import tool_specs as TS       # noqa: E402
from core.tools import launch as L      # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


BUTTONS = [
    {'button': 'RebarWizard.pushbutton', 'title': 'Rebar Wizard', 'aliases': [],
     'panel': 'Rebar & Assembly.panel', 'kind': 'script'},
    {'button': 'RebarCheck.pushbutton', 'title': 'Rebar Check', 'aliases': [],
     'panel': 'Rebar & Assembly.panel', 'kind': 'script'},
    {'button': 'BatchOut.pushbutton', 'title': 'Batch Out', 'aliases': ['BatchOut'],
     'panel': 'Views & Sheets.panel', 'kind': 'script'},
    {'button': 'Autodesk Forma.urlbutton', 'title': 'Forma', 'aliases': [],
     'panel': 'Support.panel', 'kind': 'url'},
]

print('[match]')
b, _c = L.match_button(BUTTONS, 'rebar wizard')
check('title, spaces and case ignored', b['button'] == 'RebarWizard.pushbutton')
b, _c = L.match_button(BUTTONS, 'BatchOut.pushbutton')
check('folder name', b['button'] == 'BatchOut.pushbutton')
b, cands = L.match_button(BUTTONS, 'rebar')
check('ambiguous -> candidates', b is None and set(cands) == {'Rebar Wizard', 'Rebar Check'},
      cands)
b, cands = L.match_button(BUTTONS, 'rebr wizrd')
check('typo -> close candidates', b is None and 'Rebar Wizard' in cands, cands)
check('command ids: full path, then tab/panel/button',
      L.command_ids('T3Lab', ['Views', 'Tools', 'Crop']) ==
      ['CustomCtrl_%CustomCtrl_%T3Lab%Views%Tools%Crop', 'CustomCtrl_%CustomCtrl_%T3Lab%Views%Crop'])

print('[ribbon]')


class Item(object):
    def __init__(self, name, children=None):
        self.Name, self.children = name, children

    def GetItems(self):
        if self.children is None:
            raise AttributeError('no items')
        return self.children


class Panel(Item):
    pass


posted = []


class UIApp(object):
    def __init__(self, app=None):
        pass

    def GetRibbonPanels(self, tab):
        if tab != 'T3Lab':
            raise Exception('no tab')
        return [Panel('Views & Sheets', [Item('BatchOut')]),
                Panel('Rebar & Assembly', [Item('Tools', [Item('RebarWizard')])])]

    def CanPostCommand(self, cid):
        return True

    def PostCommand(self, cid):
        posted.append(cid)


KNOWN = {'CustomCtrl_%CustomCtrl_%T3Lab%Rebar & Assembly%RebarWizard'}
ui = types.ModuleType('Autodesk.Revit.UI')
ui.UIApplication = UIApp
ui.RevitCommandId = types.SimpleNamespace(LookupCommandId=lambda s: s if s in KNOWN else None)
for name in ('Autodesk', 'Autodesk.Revit'):
    sys.modules.setdefault(name, types.ModuleType(name))
sys.modules['Autodesk.Revit.UI'] = ui
import ribbon_memory     # noqa: E402
ribbon_memory.tab_names = lambda extension_dir=None: ['Missing', 'T3Lab']
from Services import tool_discovery   # noqa: E402
tool_discovery.scan_all_buttons = lambda: BUTTONS

tab, path = L._find_on_ribbon(UIApp(), 'RebarWizard', ['Missing', 'T3Lab'])
check('found through a stack / pulldown', (tab, path) == ('T3Lab', ['Rebar & Assembly', 'Tools',
                                                                    'RebarWizard']), path)
doc = types.SimpleNamespace(Application=object())
r = TS.run('open_tool_window', doc, {'tool': 'Rebar Wizard'})
check('opened via the tab/panel/button id', r['opened'] and posted == list(KNOWN), (r, posted))
check('names the headless tool', r['headless_tools'] == ['rebar_wizard'] and 'prefer' in r['hint'])
r = TS.run('open_tool_window', doc, {'tool': 'BatchOut', 'dry_run': True})
check('dry run finds, does not post', r['dry_run'] and not r['opened'] and len(posted) == 1
      and r['headless_tools'] == ['export_sheets'], r)
r = TS.run('open_tool_window', doc, {'tool': 'BatchOut'})
check('no command id -> error says where to click', not r['opened']
      and 'Views & Sheets > Batch Out' in r['error'], r)
r = TS.run('open_tool_window', doc, {'tool': 'Forma'})
check('url buttons are not windows', 'No T3Lab window' in r['error'], r)
spec = TS.get('open_tool_window')
check('MCP only, read tier, main thread', spec.placement == ('external',)
      and spec.effect == 'read' and spec.main_thread)

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
