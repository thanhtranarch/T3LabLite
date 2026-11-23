# -*- coding: utf-8 -*-
"""Snippets/doc_switch.py: match an open document and activate it without a click,
local files and Autodesk Docs / BIM 360 models alike, and its use by the server's
switch_active_document / list_open_documents. Linux, no Revit (fakes).

Module part ported from t3lab_dev dev/test_doc_switch.py.
Run:  python3 lite_guard/test_doc_switch.py      Exit 0 = all pass.
"""
import os
import shutil
import sys
import tempfile
import types

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
_TMP = tempfile.mkdtemp(prefix='t3lab_docswitch_')
# Settings, token and data folder stay out of the real profile.
os.environ['APPDATA'] = os.environ['HOME'] = os.environ['USERPROFILE'] = _TMP

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


class Options(object):
    pass


from core import paths                  # noqa: E402
paths.save_settings({'tracking_enabled': False})
from core import server as srv_mod      # noqa: E402  (before the fakes: real import)

_db = types.ModuleType('Autodesk.Revit.DB')
for _name in ('FilteredElementCollector', 'ViewSheet', 'BuiltInCategory', 'Level', 'ElementId',
              'Transaction', 'ElementLevelFilter'):
    setattr(_db, _name, None)
_db.OpenOptions = Options
_autodesk = sys.modules.setdefault('Autodesk', types.ModuleType('Autodesk'))
_revit_ns = sys.modules.setdefault('Autodesk.Revit', types.ModuleType('Autodesk.Revit'))
_autodesk.Revit, _revit_ns.DB = _revit_ns, _db
sys.modules['Autodesk.Revit.DB'] = _db

from Snippets import doc_switch as DS   # noqa: E402


class Doc(object):
    IsLinked = False

    def __init__(self, title, path='', cloud=False):
        self.Title, self.PathName, self.IsModelInCloud = title, path, cloud

    def GetCloudModelPath(self):
        return ('cloud', self.Title)


class UIApp(object):
    def __init__(self, docs=(), active=None):
        self.calls = []
        self.Application = types.SimpleNamespace(Documents=list(docs))
        self.ActiveUIDocument = types.SimpleNamespace(Document=active) if active else None

    def OpenAndActivateDocument(self, *args):
        self.calls.append(args)
        return types.SimpleNamespace(Document=Doc('Opened', args[0]))


A = Doc('Tower_A', r'C:\Jobs\Tower_A.rvt')
B = Doc('Tower_B', 'Autodesk Docs://Jobs/Tower_B.rvt', cloud=True)
C = Doc('Project1')                                    # never saved
D = Doc('Tower_A - Podium', r'C:\Jobs\Podium.rvt')
docs = [A, B, C, D]

print('[match]')
check('exact title wins over a substring', DS.match(docs, 'tower_a')[0] is A)
check('file path', DS.match(docs, r'c:\jobs\podium.rvt')[0] is D)
check('file name without extension', DS.match(docs, 'Podium')[0] is D)
check('cloud model by file name', DS.match(docs, 'Tower_B.rvt')[0] is B)
doc, err = DS.match(docs, 'Tower')
check('ambiguous substring lists candidates', doc is None
      and set(err['candidates']) == {'Tower_A', 'Tower_B', 'Tower_A - Podium'}, err)
check('nothing matches -> (None, None)', DS.match(docs, 'Hospital') == (None, None))
check('empty query -> error', 'required' in DS.match(docs, '  ')[1]['error'])

print('[activate]')
ui = UIApp()
check('local file: activated by its path', DS.activate(ui, A) == (True, 'file')
      and ui.calls[-1] == (A.PathName,))
check('cloud model: activated by its cloud ModelPath, no detach',
      DS.activate(ui, B) == (True, 'cloud') and ui.calls[-1][0] == ('cloud', 'Tower_B')
      and isinstance(ui.calls[-1][1], Options) and ui.calls[-1][2] is False)
n = len(ui.calls)
ok, why = DS.activate(ui, C)
check('unsaved: not activated, says why, no API call', not ok and 'never been saved' in why
      and len(ui.calls) == n, why)
check('activation info for list_open_documents',
      [DS.activation(d)['kind'] for d in (A, B, C)] == ['file', 'cloud', 'unsaved']
      and DS.activation(C)['can_activate'] is False)

print('[server: switch_active_document]')
srv = srv_mod.T3LabAIServer()
_pyrevit = types.ModuleType('pyrevit')
_pyrevit.revit = types.SimpleNamespace(doc=None, uidoc=None)
_pyrevit.HOST_APP = types.SimpleNamespace(uiapp=None)
_saved_pyrevit = sys.modules.get('pyrevit')
sys.modules['pyrevit'] = _pyrevit


def switch(query, uiapp):
    return srv._execute_tool_in_context('switch_active_document',
                                        {'path_or_title': query}, uiapp=uiapp)


try:
    ui = UIApp(docs, active=A)
    res = switch('Tower_B', ui)
    check('cloud model activated without a click', res.get('success') and res.get('kind') == 'cloud'
          and ui.calls and ui.calls[-1][0] == ('cloud', 'Tower_B'), res)
    ui = UIApp(docs, active=A)
    res = switch('podium', ui)
    check('local file activated by its path', res.get('success') and res.get('kind') == 'file'
          and ui.calls == [(D.PathName,)], res)
    ui = UIApp(docs, active=A)
    res = switch('Tower_A', ui)
    check('already active: no-op, no API call', res.get('success') and ui.calls == []
          and 'already the active' in res.get('note', ''), res)
    ui = UIApp(docs, active=A)
    res = switch('Project1', ui)
    check('never saved: says why, no API call', 'never been saved' in res.get('error', '')
          and ui.calls == [] and 'Project1' in res.get('open_documents', []), res)
    ui = UIApp(docs, active=A)
    res = switch('Tower', ui)
    check('ambiguous: candidates, no API call', res.get('candidates') and ui.calls == [], res)
    disk = os.path.join(_TMP, 'Hospital.rvt')
    with open(disk, 'w') as f:
        f.write('x')
    ui = UIApp(docs, active=A)
    res = switch(disk, ui)
    check('full path of a file not open: opened from disk', res.get('opened_from_disk')
          and ui.calls == [(disk,)], res)
    ui = UIApp(docs, active=A)
    res = switch('Hospital', ui)
    check('not open, not a file: error lists open documents', 'No open document' in
          res.get('error', '') and ui.calls == [], res)

    print('[server: list_open_documents]')
    ui = UIApp(docs + [type('Linked', (Doc,), {'IsLinked': True})('Link')], active=A)
    _pyrevit.HOST_APP.uiapp = ui
    _pyrevit.revit.doc = A
    res = srv._execute_tool_in_context('list_open_documents', {}, uiapp=ui)
    rows = dict((d['title'], d) for d in res.get('documents', []))
    check('linked documents left out', sorted(rows) == sorted(d.Title for d in docs), rows)
    check('kind + can_activate per document',
          [(rows[t]['kind'], rows[t]['can_activate']) for t in ('Tower_A', 'Tower_B', 'Project1')]
          == [('file', True), ('cloud', True), ('unsaved', False)], rows)
    check('active one marked, unsaved path kept', rows['Tower_A']['is_active']
          and rows['Project1']['path'] == '(unsaved)', rows)
finally:
    if _saved_pyrevit is None:
        sys.modules.pop('pyrevit', None)
    else:
        sys.modules['pyrevit'] = _saved_pyrevit
    shutil.rmtree(_TMP, ignore_errors=True)

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
