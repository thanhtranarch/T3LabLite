# -*- coding: utf-8 -*-
"""Read-back after a model write: does the model now say what the tool said?

"Never trust exit 0" (HKUDS/CLI-Anything) applied to Revit edits. Runs on the
Revit main thread right after the write (server handler), for MCP clients and
the Assistant alike; reads only, never fails a call. The Revit side is passed
in as callables, so this module has no Revit import (IronPython 2.7 + CPython 3).
Doc: docs/agents/write-verify.md. Test: dev/test_write_verify.py.
"""
from __future__ import unicode_literals

import os
import re

from core import tool_specs as _specs     # tools declared in core/tools/<domain>.py

MAX_IDS = 200            # existence checks per call
CLOCK_SLACK = 2.0        # s: file mtime vs the call's start

# Successful calls of these must leave a change in DocumentChanged.
MUST_CHANGE = frozenset([
    'create_level', 'create_grid', 'create_sheet', 'create_room', 'place_wall',
    'create_point_based_element', 'create_line_based_element',
    'create_surface_based_element', 'create_text_note', 'create_view',
    'duplicate_view', 'copy_elements', 'create_dimension', 'create_schedule',
    'create_detail_annotation', 'create_structural_framing_system',
]) | _specs.names(verify='must_change')
# Every element id these report must exist afterwards (none of them deletes).
ID_TOOLS = MUST_CHANGE | frozenset([
    'move_elements', 'rotate_element', 'rename_element', 'set_parameter',
    'add_view_to_sheet', 'place_views_on_sheets', 'apply_view_template',
    'set_element_workset', 'create_view_filter', 'tag_elements', 'tag_all_walls',
    'tag_all_rooms', 'color_elements', 'revit_override_color', 'join_geometry',
]) | _specs.names(verify='ids')
# Tools that read their own result back and list what is missing in
# `not_in_model_after` (e.g. core/tools/views.py).
SELF_CHECKED = _specs.names(verify='self')
# Tools whose result lists the files they wrote in `files` (+ `exported`).
FILE_LIST_TOOLS = _specs.names(verify='files')
# Not element ids.
_NOT_ELEMENT = frozenset(['workset_id', 'call_id', 'result_id', 'approval_id', 'job_id',
                          'session_id', 'task_id'])
_ID_KEY = re.compile(r'^(?:[a-z0-9]+_)*ids?$')
_NUM = re.compile(r'^\s*([-+]?\d+(?:[.,]\d+)?)')


def _ints(value):
    out = []
    for v in (value if isinstance(value, (list, tuple)) else [value]):
        if isinstance(v, bool):
            continue
        if isinstance(v, dict):
            v = v.get('id', v.get('element_id'))
        try:
            out.append(int(v))
        except (TypeError, ValueError):
            continue
    return out


def reported_ids(result):
    """Element ids a result names: *_id / *_ids keys at the top level and the
    id / element_id of rows one level down."""
    ids = []
    for key, value in result.items():
        if key in _NOT_ELEMENT or not _ID_KEY.match('{}'.format(key)):
            if isinstance(value, list) and value and isinstance(value[0], dict):
                for row in value:
                    ids += _ints(row.get('id', row.get('element_id')))
            continue
        ids += _ints(value)
    seen, out = set(), []
    for i in ids:
        if i > 0 and i not in seen:
            seen.add(i)
            out.append(i)
    return out


def _number(text):
    m = _NUM.match('{}'.format(text or ''))
    return float(m.group(1).replace(',', '.')) if m else None


def same_value(asked, got):
    """'ok' | 'units' (numbers differ: maybe a unit conversion) | 'differs'."""
    a = '{}'.format(asked if asked is not None else '').strip()
    g = '{}'.format(got if got is not None else '').strip()
    if a.lower() == g.lower():
        return 'ok'
    na, ng = _number(a), _number(g)
    if na is not None and ng is not None:
        if abs(na - ng) <= 1e-6 * max(1.0, abs(na)):
            return 'ok'
        return 'units'
    return 'differs'


def _files_since(folder, ext, since):
    try:
        names = os.listdir(folder)
    except (OSError, TypeError):
        return None
    out = []
    for n in names:
        p = os.path.join(folder, n)
        if n.lower().endswith('.' + ext) and os.path.isfile(p):
            try:
                if os.path.getmtime(p) >= since - CLOCK_SLACK:
                    out.append(p)
            except OSError:
                continue
    return sorted(out)


def _file_problem(path):
    try:
        from core import tool_cli
        r = tool_cli.verify_file(path)
        return None if r.get('ok') else '{}: {}'.format(os.path.basename(path), r.get('problem'))
    except ImportError:
        if not os.path.isfile(path) or os.path.getsize(path) == 0:
            return '{}: missing or empty'.format(os.path.basename(path))
        return None


def _check_files(tool, args, result, started, fail, count):
    folder = result.get('output_folder')
    if tool == 'export_image':
        files = [f for f in (result.get('files') or []) if isinstance(f, (str, type('')))]
        if not files:
            fail('export_image reported no file')
    else:
        ext = 'pdf' if tool == 'export_sheets_pdf' else 'dwg'
        files = _files_since(folder, ext, started)
        if files is None:
            fail('output folder not readable: {}'.format(folder))
            return
        if tool == 'export_sheets_pdf':
            expected = 1 if result.get('combined') else int(result.get('sheet_count') or 0)
            if len(files) != expected:
                fail('{} new PDF file(s) in the folder, expected {}'.format(len(files), expected))
        else:
            expected = int(result.get('count') or 0)
            if len(files) < expected:      # unmerged views add xref DWGs, never fewer
                fail('{} new DWG file(s) in the folder, expected at least {}'.format(
                    len(files), expected))
    for f in files:
        count()
        problem = _file_problem(f)
        if problem:
            fail(problem)


def verify(tool, args, result, started, exists, name_of=None, changes=None):
    """{'verified': bool, 'checks': n, 'failed'?: [...], 'warnings'?: [...]}, or
    None when nothing applies (error, dry run, unknown tool).

    exists(id) -> bool and name_of(id) -> str read the model; changes =
    (added, modified, deleted) from DocumentChanged during the call, or None.
    """
    if not isinstance(result, dict) or result.get('error') or result.get('success') is False:
        return None
    args = args if isinstance(args, dict) else {}
    if args.get('dry_run') or result.get('dry_run') or result.get('applied') is False:
        return None                 # previews and read operations of write tools
    failed, warnings, n = [], [], [0]

    def fail(text):
        failed.append(text)

    def count():
        n[0] += 1

    if tool in ID_TOOLS:
        ids = reported_ids(result)
        missing = []
        for i in ids[:MAX_IDS]:
            count()
            if not exists(i):
                missing.append(i)
        if missing:
            fail('reported element(s) not in the model: {}'.format(missing[:20]))
        if len(ids) > MAX_IDS:
            warnings.append('checked {} of {} ids'.format(MAX_IDS, len(ids)))

    if tool == 'delete_element':
        gone = _ints(result.get('deleted') or [])
        still = []
        for i in gone[:MAX_IDS]:
            count()
            if exists(i):
                still.append(i)
        if still:
            fail('reported deleted but still in the model: {}'.format(still[:20]))

    if tool in MUST_CHANGE and changes is not None:
        count()
        if not any(changes):
            fail('Revit recorded no model change')

    if name_of is not None:
        expected = {'create_level': ('level_id', args.get('name')),
                    'create_sheet': ('sheet_id', args.get('sheet_name')),
                    'duplicate_view': ('new_view_id', args.get('name')),
                    'rename_element': ('element_id', args.get('new_name'))}.get(tool)
        if expected and expected[1] and _ints(result.get(expected[0]) or args.get(expected[0])):
            eid = _ints(result.get(expected[0]) or args.get(expected[0]))[0]
            count()
            got = name_of(eid)
            if got != expected[1]:
                # Revit makes names unique ("Level 1 1"); the tool may say so.
                warnings.append('name in the model is "{}", asked "{}"'.format(got, expected[1]))

    if tool == 'set_parameter':
        rows = [r for r in (result.get('diff') or []) if isinstance(r, dict)]
        if rows:
            count()
            verdict = same_value(args.get('value'), rows[-1].get('after'))
            if verdict == 'differs':
                fail('parameter now reads "{}", asked "{}"'.format(rows[-1].get('after'),
                                                                   args.get('value')))
            elif verdict == 'units':
                warnings.append('parameter reads "{}" for "{}": check units'.format(
                    rows[-1].get('after'), args.get('value')))

    if tool in SELF_CHECKED or result.get('not_in_model_after'):
        count()
        lost = result.get('not_in_model_after') or []
        if lost:
            fail('reported done but not in the model: {}'.format(lost[:20]))

    if tool in FILE_LIST_TOOLS and result.get('files') is not None:
        files = result.get('files') or []
        if len(files) != result.get('exported', len(files)):
            warnings.append('only the first {} files were checked'.format(len(files)))
        for f in files:
            count()
            problem = _file_problem(f)
            if problem:
                fail(problem)

    if tool in ('export_sheets_pdf', 'export_dwg', 'export_image'):
        _check_files(tool, args, result, started, fail, count)

    if not n[0] and not failed:
        return None
    out = {'verified': not failed, 'checks': max(n[0], 1)}
    if failed:
        out['failed'] = failed
    if warnings:
        out['warnings'] = warnings
    return out


HINT = ('The tool reported success, but reading the model back disagrees (see '
        'verification.failed). Do not report this as done. Check the model before '
        'retrying: a retry may apply the change twice.')


def annotate(result, report):
    """Fold a verify() report into the tool result (in place)."""
    if not report or not isinstance(result, dict):
        return result
    result['verified'] = report['verified']
    result['verification'] = dict((k, v) for k, v in report.items() if k != 'verified')
    if not report['verified']:
        result['verification_hint'] = HINT
    return result
