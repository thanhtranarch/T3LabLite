# -*- coding: utf-8 -*-
"""
send_code_to_revit errors -> short, fixable summary (not a full traceback).

summarize(): exception + the user's failing line + a hint for known mistakes.
classify(): same rule table as a category key; reused by
dev/analyze_mcp_audit.py so offline reports match what the model is told.
Why: send_code = 74 % of MCP calls, 15 % failed (tracking 2026-10-03..17).
IronPython 2.7 + CPython 3, pure Python.
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__ = "Code Errors"

import re
import traceback

# exec() filename of the user's code: tells its frames from the server's.
CODE_FILENAME = '<send_code>'
MESSAGE_CHARS = 600
SOURCE_CHARS = 200
MAX_FRAMES = 3

# Names send_code_to_revit puts in scope next to doc / uidoc / app.
SCOPE_HINT = ('In scope: doc, uidoc, app, result, output, eid_value, make_eid, '
              'elem_name, revit_version. Import Revit types yourself: '
              'from Autodesk.Revit.DB import ...')

_F_STRING = re.compile(r'''(?<![\w'"])(?:[rR]?[fF]|[fF][rR])['"]''')

# (category, type-name regex, message regex, hint); first match wins.
# Type = class name, so .NET exceptions from IronPython match too.
_RULES = [
    ('integer_value', None, r'IntegerValue',
     'ElementId.IntegerValue no longer exists on Revit 2026+. Use '
     'eid_value(element_id) (in scope), which works on every version.'),
    ('elementid_value', None, r"ElementId.*attribute.*'Value'|'Value'.*ElementId",
     'ElementId.Value does not exist before Revit 2024. Use '
     'eid_value(element_id) (in scope), which works on every version.'),
    ('removed_api', None, r'BuiltInParameterGroup|ParameterGroup',
     'BuiltInParameterGroup was removed in Revit 2025. Use GroupTypeId '
     '(Definition.GetGroupTypeId()).'),
    ('removed_api', None, r'DisplayUnitType|\bUnitType\b|ParameterType',
     'This unit/parameter type API was removed. Use UnitTypeId / SpecTypeId '
     'and Definition.GetDataType().'),
    ('no_transaction', None,
     r'no open transaction|Modification of the document is forbidden|'
     r'Modifying is forbidden',
     "Model edits need a transaction: t = Transaction(doc, 'T3Lab'); "
     't.Start(); ...; t.Commit() (from Autodesk.Revit.DB import Transaction).'),
    ('transaction_open', None,
     r'Starting a new transaction is not permitted|already (?:started|open)',
     'A transaction is already open. Commit it before starting another, or '
     'use SubTransaction inside it.'),
    ('wrong_id_type', None, r'expected ElementId|ElementId.*got int',
     'This API wants an ElementId, not an int: wrap it with make_eid(n).'),
    ('none_result', None, r"'NoneType' object (?:has no attribute|is not)",
     'Something returned None: doc.GetElement() on a missing id, or '
     'LookupParameter()/get_Parameter() for a parameter this element does not '
     'have. Check for None before using it.'),
    ('no_overload', None,
     r'No constructor matches|No method matches|expected \w+, got|'
     r'Multiple targets could match',
     'No Revit API overload matched the arguments. Check types: lengths are '
     'float feet, ids are ElementId, collections need a .NET List.'),
    ('unicode', r'^Unicode(?:En|De)codeError$', None,
     'IronPython 2.7: Revit names are already unicode. Do not .encode()/'
     ".decode() them; use u'...' literals and '{}'.format()."),
    ('name_error', r'^(?:NameError|UnboundLocalError)$', None, SCOPE_HINT),
]

_PY3_SYNTAX = re.compile(r':=|\bnonlocal\b|^\s*async\s+def\b|\)\s*->\s*\w',
                         re.MULTILINE)


def _type_name(exc):
    try:
        return type(exc).__name__
    except Exception:
        return 'Exception'


def _message(exc):
    try:
        text = u'{}'.format(exc)
    except Exception:
        try:
            text = repr(exc)
        except Exception:
            text = ''
    text = text.strip()
    if len(text) > MESSAGE_CHARS:
        text = text[:MESSAGE_CHARS] + '...'
    return text


def classify(type_name, message, code=''):
    """(category, hint) for one error, or (None, None) when no rule fits."""
    type_name = type_name or ''
    message = message or ''
    code = code or ''
    if type_name in ('SyntaxError', 'IndentationError', 'TabError'):
        if _F_STRING.search(code):
            return ('f_string',
                    "This runs on IronPython 2.7: no f-strings. Use "
                    "'{}'.format(x) instead.")
        if _PY3_SYNTAX.search(code):
            return ('py3_syntax',
                    'This runs on IronPython 2.7 (Python 2 syntax): no type '
                    'hints, walrus, nonlocal or async.')
        return ('syntax', 'This runs on IronPython 2.7 (Python 2 syntax).')
    for category, type_rx, msg_rx, hint in _RULES:
        if type_rx and not re.search(type_rx, type_name):
            continue
        if msg_rx and not re.search(msg_rx, message):
            continue
        return (category, hint)
    return (None, None)


def _code_line(code, lineno):
    try:
        lines = (code or '').splitlines()
        if lineno and 0 < int(lineno) <= len(lines):
            return lines[int(lineno) - 1].strip()[:SOURCE_CHARS]
    except Exception:
        pass
    return None


def summarize(exc, code, tb=None):
    """Short error dict for a failed send_code_to_revit run.

    {'error': '<Type>: <message>', 'line': n, 'source': '<that line>',
     'frames': ['line 3 in helper', ...], 'category': ..., 'hint': ...}
    Keys without a value are left out. Never raises.
    """
    try:
        type_name = _type_name(exc)
        message = _message(exc)
        out = {'error': u'{}: {}'.format(type_name, message) if message
               else type_name}
        lineno = None
        frames = []
        if type_name in ('SyntaxError', 'IndentationError', 'TabError'):
            lineno = getattr(exc, 'lineno', None)
        elif tb is not None:
            try:
                entries = traceback.extract_tb(tb)
            except Exception:
                entries = []
            mine = [e for e in entries if e[0] == CODE_FILENAME]
            if mine:
                lineno = mine[-1][1]
                for entry in mine[-MAX_FRAMES:]:
                    frames.append(u'line {} in {}'.format(entry[1], entry[2]))
        if lineno:
            out['line'] = lineno
            source = _code_line(code, lineno)
            if source:
                out['source'] = source
        if len(frames) > 1:
            out['frames'] = frames
        category, hint = classify(type_name, message, code)
        if category:
            out['category'] = category
            out['hint'] = hint
        return out
    except Exception:
        return {'error': u'{}'.format(exc)}
