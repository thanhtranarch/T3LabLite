# -*- coding: utf-8 -*-
"""core/write_verify.py + T3LabAIServer._verify_write: the read-back after a write.

Run:  python lite_guard/test_write_verify.py      (no Revit: the model is a fake)
"""

import ast
import os
import sys
import tempfile
import time

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(REPO, 'lib')
sys.path.insert(0, LIB)
os.environ['APPDATA'] = tempfile.mkdtemp(prefix='t3lab_wv_test_')

from core import write_verify as W   # noqa: E402
import core.server as S               # noqa: E402

FAILURES = []


def check(name, cond, detail=''):
    if cond:
        print('  ok    {}'.format(name))
    else:
        FAILURES.append(name)
        print('  FAIL  {}  {}'.format(name, detail))


MODEL = {101: 'Level 1', 202: 'A-101 Plan', 303: 'Wall'}
exists = MODEL.__contains__
name_of = MODEL.get
NOW = time.time()


def run(tool, args, result, changes=(1, 0, 0), started=NOW):
    return W.verify(tool, args, result, started, exists, name_of, changes)


# ── ids ─────────────────────────────────────────────────────────────────────
print('[ids]')
check('reported_ids: *_id, *_ids, rows, no workset/bools', sorted(W.reported_ids({
    'level_id': 101, 'element_ids': [303, '202'], 'created': [{'id': 5}],
    'workset_id': 9, 'success': True, 'count': 3})) == [5, 101, 202, 303])
r = run('create_level', {'name': 'Level 1', 'elevation': 3}, {'success': True, 'level_id': 101})
check('create ok: id exists, change recorded, name matches', r == {'verified': True, 'checks': 3}, r)
r = run('create_level', {'elevation': 3}, {'success': True, 'level_id': 999})
check('reported id missing -> not verified', not r['verified']
      and 'not in the model: [999]' in r['failed'][0], r)
r = run('create_level', {'elevation': 3}, {'success': True, 'level_id': 101}, changes=(0, 0, 0))
check('success with no recorded change -> not verified',
      r['failed'] == ['Revit recorded no model change'], r)
r = run('create_level', {'elevation': 3}, {'success': True, 'level_id': 101}, changes=None)
check('no DocumentChanged hook -> change check skipped', r == {'verified': True, 'checks': 1}, r)
r = run('create_level', {'name': 'L1', 'elevation': 3}, {'success': True, 'level_id': 101})
check('Revit renamed it -> warning, still verified', r['verified']
      and 'asked "L1"' in r['warnings'][0], r)
r = run('delete_element', {'element_ids': [303, 7]}, {'deleted': [{'id': 7}, {'id': 303}],
                                                      'deleted_count': 2}, changes=(0, 0, 2))
check('delete: survivor caught', not r['verified'] and '[303]' in r['failed'][0], r)
r = run('delete_element', {'element_ids': [7]}, {'deleted': [{'id': 7}], 'deleted_count': 1})
check('delete ok', r['verified'], r)

# ── skipped ─────────────────────────────────────────────────────────────────
print('[skipped]')
check('error result', run('create_level', {}, {'error': 'x'}) is None)
check('success False (turbo rollback)', run('create_level', {}, {'success': False}) is None)
check('dry run', run('set_parameter', {'dry_run': True}, {'dry_run': True, 'diff': []}) is None)
check('tool without checks', run('select_elements', {}, {'success': True}) is None)
check('switch_active_document untouched', run('switch_active_document', {}, {'ok': 1}) is None)

# ── values ──────────────────────────────────────────────────────────────────
print('[values]')
check('same_value text', W.same_value('Concrete', ' concrete ') == 'ok')
check('same_value number + unit', W.same_value('3000', '3000 mm') == 'ok')
check('same_value comma decimal', W.same_value('2,5', '2.5 m') == 'ok')
check('same_value numbers differ -> units', W.same_value('3000', '3 m') == 'units')
check('same_value text differs', W.same_value('Fire 60', 'Fire 90 min') == 'differs'
      and W.same_value('A', 'B') == 'differs')


def sp(asked, after):
    return run('set_parameter', {'element_id': 303, 'parameter_name': 'Mark', 'value': asked},
               {'success': True, 'element_id': 303, 'applied': True, 'dry_run': False,
                'diff': [{'element_id': 303, 'field': 'Mark', 'before': 'X', 'after': after}]})


check('set_parameter landed', sp('W-01', 'W-01')['verified'])
r = sp('W-01', 'W-02')
check('set_parameter stored something else', not r['verified']
      and 'reads "W-02"' in r['failed'][0], r)
r = sp('3000', '3 m')
check('set_parameter numbers differ -> warning only', r['verified'] and 'units' in r['warnings'][0])
r = run('rename_element', {'element_id': 202, 'new_name': 'A-102 Plan'},
        {'success': True, 'element_id': 202})
check('rename not applied -> warning names both', 'A-101 Plan' in r['warnings'][0], r)

# ── exports ─────────────────────────────────────────────────────────────────
print('[exports]')
out = tempfile.mkdtemp(prefix='t3lab_wv_out_')


def w(name, data, age=0):
    p = os.path.join(out, name)
    with open(p, 'wb') as f:
        f.write(data)
    if age:
        os.utime(p, (NOW - age, NOW - age))
    return p


w('old.pdf', b'%PDF-1.4', age=3600)
w('A-101.pdf', b'%PDF-1.7')
w('A-102.pdf', b'%PDF-1.7')
started = NOW - 1
r = W.verify('export_sheets_pdf', {}, {'success': True, 'sheet_count': 2, 'output_folder': out,
                                       'combined': False}, started, exists)
check('pdf: 2 new files for 2 sheets, old file ignored', r == {'verified': True, 'checks': 2}, r)
r = W.verify('export_sheets_pdf', {}, {'success': True, 'sheet_count': 3, 'output_folder': out,
                                       'combined': False}, started, exists)
check('pdf: fewer files than sheets', not r['verified'] and 'expected 3' in r['failed'][0], r)
w('A-103.pdf', b'<html>')
r = W.verify('export_sheets_pdf', {}, {'success': True, 'sheet_count': 3, 'output_folder': out,
                                       'combined': False}, started, exists)
check('pdf: bad header caught', not r['verified'] and 'A-103.pdf' in r['failed'][0], r)
w('S1.dwg', b'AC1032....')
w('S1-View.dwg', b'AC1032....')
r = W.verify('export_dwg', {}, {'success': True, 'count': 1, 'output_folder': out},
             started, exists)
check('dwg: xref files allowed (>= count)', r['verified'], r)
r = W.verify('export_image', {}, {'success': True, 'files': [], 'output_folder': out},
             started, exists)
check('image: no file reported', not r['verified'], r)
r = W.verify('export_dwg', {}, {'success': True, 'count': 1,
                                'output_folder': os.path.join(out, 'nope')}, started, exists)
check('unreadable folder', not r['verified'] and 'not readable' in r['failed'][0], r)

# ── annotate ────────────────────────────────────────────────────────────────
print('[annotate]')
res = W.annotate({'success': True}, {'verified': False, 'checks': 1, 'failed': ['x']})
check('failure: flag + hint, success untouched', res['verified'] is False
      and res['success'] is True and 'Do not report this as done' in res['verification_hint']
      and res['verification'] == {'checks': 1, 'failed': ['x']}, res)
res = W.annotate({'success': True}, {'verified': True, 'checks': 2})
check('ok: compact', res == {'success': True, 'verified': True, 'verification': {'checks': 2}})
check('no report: result unchanged', W.annotate({'a': 1}, None) == {'a': 1})

# ── server hook ─────────────────────────────────────────────────────────────
print('[server]')


class _Elem(object):
    def __init__(self, name):
        self.Name = name


class _Doc(object):
    Title = 'Tower.rvt'

    def GetElement(self, eid):
        return _Elem(MODEL[eid]) if eid in MODEL else None


class _UIDoc(object):
    Document = _Doc()


class _App(object):
    ActiveUIDocument = _UIDoc()


class _Call(object):
    added, modified, deleted = 1, 0, 0


class _Task(object):
    tool_name = 'create_level'
    arguments = {'name': 'Level 1', 'elevation': 3}
    audit = _Call()
    result = {'success': True, 'level_id': 101}


S.make_eid = lambda i: i
S.elem_name = lambda e: e.Name
srv = object.__new__(S.T3LabAIServer)
srv._t3lab_audit_doc_changed = object()
res = srv._verify_write(_App(), _Task(), NOW)
check('server: verified + 3 checks', res.get('verified') is True
      and res['verification'] == {'checks': 3}, res)
t = _Task()
t.result = {'success': True, 'level_id': 999}
res = srv._verify_write(_App(), t, NOW)
check('server: missing id flagged', res['verified'] is False and 'verification_hint' in res, res)
srv2 = object.__new__(S.T3LabAIServer)       # DocumentChanged never hooked
t = _Task()
t.audit = None
t.result = {'success': True, 'level_id': 101}
check('server: no hook, no audit -> no change check',
      srv2._verify_write(_App(), t, NOW)['verification'] == {'checks': 2})


class _Boom(object):
    @property
    def ActiveUIDocument(self):
        raise RuntimeError('API context')


t = _Task()
t.result = {'success': True, 'level_id': 101}
check('server: read failure never fails the call',
      srv._verify_write(_Boom(), t, NOW) == {'success': True, 'level_id': 101})
src = open(os.path.join(LIB, 'core', 'server.py'), encoding='utf-8').read()
check('handler calls _verify_write for writes only, tolerates an old server object',
      "verify = getattr(self.server, '_verify_write', None) if is_write else None" in src
      and 'task.result = verify(app, task, started)' in src)

# ── IronPython 2.7: no py3-only syntax in the module ────────────────────────
print('[ironpython]')
wv_src = open(os.path.join(LIB, 'core', 'write_verify.py'), encoding='utf-8').read()
tree = ast.parse(wv_src)
bad = [type(n).__name__ for n in ast.walk(tree)
       if isinstance(n, (ast.JoinedStr, ast.Nonlocal, ast.AsyncFunctionDef, ast.NamedExpr))
       or (isinstance(n, ast.arguments) and (n.kwonlyargs or n.posonlyargs))]
check('write_verify.py parses as py2-compatible syntax', not bad
      and 'print(' not in wv_src.replace("'print('", ''), bad)

print('')
if FAILURES:
    print('{} FAILED: {}'.format(len(FAILURES), ', '.join(FAILURES)))
    sys.exit(1)
print('all passed')
