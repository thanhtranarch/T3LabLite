# -*- coding: utf-8 -*-
"""
What one MCP tool call DOES, in words a person can check before saying Yes.

Used by the Revit confirmation for AI apps (core.mcp_policy.confirm_text) and
by the T3Lab Assistant's tool and approval cards, instead of a tool name with
raw JSON or "N line(s) of code".

send_code_to_revit has two sources, kept apart on purpose:
  - `description`, the AI's own one-line purpose for the code, and
  - code_findings(): what T3Lab reads in the code itself (transactions,
    deletes, parameter writes, categories, file and network access). The AI
    cannot word these away: a description that leaves out a delete still
    shows "Deletes elements" under it.

Plain Python, no Revit import: runs under IronPython 2.7 (MCP host) and
CPython 3 (Assistant, lite_guard/test_mcp_server.py).
"""

from __future__ import unicode_literals

import re

CODE_TOOL = 'send_code_to_revit'

# Where the headline of a code call came from (describe_code()['source']).
SOURCE_AI = 'ai'                    # the AI's `description` argument
SOURCE_TRANSACTION = 'transaction'  # Transaction(doc, "name") in the code
SOURCE_CODE = 'code'                # T3Lab's own reading of the code

_HEADLINE_LIMIT = 160
_VALUE_LIMIT = 40
_NAMES_LIMIT = 6

# Transaction names that say nothing about the change.
_GENERIC_TX = frozenset([
    '', 't', 'tx', 'tr', 'trans', 'transaction', 'ai', 'mcp', 'code', 'script',
    'python', 'run', 'run code', 'ai code', 't3lab', 't3lab ai', 't3lab mcp',
    'mcp code', 'change', 'changes', 'edit', 'update', 'modify',
])

# Model changes, most serious first. Each pattern stands for a Revit API
# call; the phrase is what the confirmation lists. _CREATES has no pattern:
# code_findings() sets it from _CREATE_RE, which also names the classes.
_CREATES = 'Creates elements'
_ACTIONS = (
    (r'(?<!File)(?<!Directory)\.Delete\s*\(', 'Deletes elements'),
    (r'\bSynchronizeWithCentral\s*\(',
     'Syncs with central (publishes to the whole team)'),
    (r'\b(?P<recv>\w+)\.SaveAs\s*\(', 'Saves the model under a new name'),
    (r'\b(?P<recv>\w+)\.Save\s*\(', 'Saves the model'),
    (r'\b(?P<recv>\w+)\.Close\s*\(', 'Closes a document'),
    (r'\b(?:OpenDocumentFile|OpenAndActivateDocument)\s*\(',
     'Opens a document'),
    (None, _CREATES),
    (r'\.(?:Set|SetValueString)\s*\(', 'Changes parameter values'),
    (r'\.(?:Name|SheetNumber|Number)\s*=(?!=)', 'Renames or renumbers elements'),
    (r'\.ChangeTypeId\s*\(', 'Changes element types'),
    (r'\bElementTransformUtils\.(?:Move|Rotate)Elements?\s*\('
     r'|\.Location\.(?:Move|Rotate)\s*\('
     r'|\.Location\.(?:Point|Curve)\s*=(?!=)', 'Moves or rotates elements'),
    (r'\bElementTransformUtils\.MirrorElements?\s*\(', 'Mirrors elements'),
    (r'\bElementTransformUtils\.CopyElements?\s*\(|\.Duplicate\s*\(',
     'Copies or duplicates elements'),
    (r'\.Pinned\s*=(?!=)', 'Pins or unpins elements'),
    (r'\b(?:Hide|Unhide)Elements\s*\(|\bSetElementOverrides\s*\('
     r'|\bIsolateElementsTemporary\s*\(|\bSetCategoryHidden\s*\(',
     'Changes what views show (hide / graphic overrides)'),
    (r'\bLoadFamily(?:Symbol)?\s*\(', 'Loads families'),
    (r'\b(?P<recv>\w+)\.Export\s*\(', 'Exports files (PDF / DWG / IFC / image)'),
)

# These do not change the model, so they leave the code "read only".
_UI_ACTIONS = (
    (r'\bSetElementIds\s*\(', 'Changes the selection'),
    (r'\.ActiveView\s*=(?!=)|\bRequestViewChange\s*\(',
     'Switches the active view'),
)

# Reaching outside Revit: always listed, whatever else the code does.
_OUTSIDE = (
    (r'\bos\.(?:remove|unlink|rmdir|removedirs)\s*\(|\bshutil\.rmtree\s*\('
     r'|\b(?:File|Directory)\.Delete\s*\(', 'deletes files or folders on disk'),
    (r'\bopen\s*\([^)]*[\'"][wax]b?\+?[\'"]'
     r'|\bFile\.(?:Write|Append|Copy|Move)\w*\s*\(|\bStreamWriter\s*\('
     r'|\bshutil\.(?:copy\w*|move)\s*\(|\bos\.(?:rename|makedirs|mkdir)\s*\(',
     'writes files on disk'),
    (r'\b(?:urllib2?|httplib|http\.client|requests|socket|ftplib|smtplib)\b'
     r'|\b(?:WebClient|HttpClient|WebRequest|HttpWebRequest|TcpClient)\b',
     'uses the network'),
    (r'\b(?:subprocess|os\.system|os\.popen|os\.startfile)\b'
     r'|\bProcess\.Start\s*\(', 'starts other programs'),
)

_ACTIONS_RE = [(re.compile(p) if p else None, text) for p, text in _ACTIONS]
_UI_ACTIONS_RE = [(re.compile(p), text) for p, text in _UI_ACTIONS]
_OUTSIDE_RE = [(re.compile(p), text) for p, text in _OUTSIDE]

# `recv` in a pattern above is the object the method is called on.
# writer.Close() or wb.Save() is a file or Excel, not the Revit model.
_NOT_DOC_RECEIVER = re.compile(
    r'(?i)writer|reader|stream|file|workbook|excel|zip|client|conn|socket'
    r'|response|^(?:f|fp|fh|fs|sw|sr|wb|ws|book|sheet|xl|r|req)$')

_COMMENT_RE = re.compile(r'^\s*#.*$', re.M)
_HAS_TX_RE = re.compile(r'\b(?:Transaction|TransactionGroup)\s*\(')
_TX_NAME_RE = re.compile(
    r'\b(?:Transaction|TransactionGroup)\s*\(\s*[\w.]+\s*,\s*[uUrR]?([\'"])(.+?)\1')
_TX_START_RE = re.compile(r'\.Start\s*\(\s*[uUrR]?([\'"])(.+?)\1\s*\)')
_CATEGORY_RE = re.compile(r'\bBuiltInCategory\.OST_(\w+)')
_CLASS_RE = re.compile(
    r'\bOfClass\s*\(\s*(?:clr\.GetClrType\s*\(\s*|typeof\s*\(\s*)?(?:\w+\.)*(\w+)\s*\)')
_PARAM_RE = re.compile(
    r'\b(?:LookupParameter|GetParameters)\s*\(\s*[uUrR]?([\'"])(.+?)\1')
_BIP_RE = re.compile(r'\bBuiltInParameter\.(\w+)')
_CREATE_RE = re.compile(r'\b([A-Z]\w*)\.Create\s*\(|\.New([A-Z]\w*)\s*\(')
_CAMEL_RE = re.compile(r'(?<=[a-z0-9])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])')

# Geometry and helper objects: `Line.Create...` is not a model element.
_NOT_ELEMENTS = frozenset([
    'Arc', 'Line', 'LineBound', 'LineUnbound', 'Ellipse', 'NurbSpline',
    'HermiteSpline', 'CylindricalHelix', 'CurveLoop', 'CurveArray',
    'CurveArrArray', 'Plane', 'Transform', 'XYZ', 'UV', 'Frame',
    'BoundingBoxXYZ', 'Outline', 'Options', 'ElementId', 'List', 'Solid',
    'Color', 'FilteredElementCollector', 'Transaction', 'TransactionGroup',
])


def _text(value):
    try:
        return u'{}'.format(value)
    except Exception:
        return u''


def _clip(text, limit):
    """One line, at most `limit` characters."""
    text = u' '.join(_text(text).split())
    if len(text) <= limit:
        return text
    return text[:limit - 1].rstrip() + u'…'


def _plural(n, word, plural=None):
    return u'{} {}'.format(n, word if n == 1 else (plural or word + u's'))


def _names(items, limit=_NAMES_LIMIT):
    """'a, b, c' -- or 'a, b, c +4 more'."""
    items = [_text(i) for i in items]
    shown = u', '.join(items[:limit])
    if len(items) > limit:
        shown += u' +{} more'.format(len(items) - limit)
    return shown


def _unique(items):
    seen, out = set(), []
    for item in items:
        key = item.lower()
        if item and key not in seen:
            seen.add(key)
            out.append(item)
    return out


def _matches(rx, code):
    for m in rx.finditer(code):
        recv = m.groupdict().get('recv')
        if recv is None or not _NOT_DOC_RECEIVER.search(recv):
            return True
    return False


def _pretty_category(ost_name):
    """'StructuralColumns' -> 'Structural Columns', 'MEPSpaces' -> 'MEP Spaces'."""
    return _CAMEL_RE.sub(u' ', ost_name).strip()


def _as_list(value):
    return list(value) if isinstance(value, (list, tuple)) else []


def _str_arg(args, key):
    return _clip(args.get(key) or u'', _VALUE_LIMIT)


# ── send_code_to_revit ──────────────────────────────────────────────────────

def code_findings(code):
    """What the code itself does, read with patterns (no exec, no parse: it is
    IronPython 2.7 and may not even compile under CPython 3).

    Returns a dict:
      transactions  named transactions, in order
      has_transaction
      actions       model changes, e.g. 'Deletes elements'
      ui_actions    selection / active view changes
      outside       'writes files on disk', 'uses the network', ...
      creates       element classes it creates ('Wall', 'ViewSheet')
      categories    'Walls', 'Structural Columns'
      classes       OfClass(...) targets
      parameters    LookupParameter names and BuiltInParameter members
      read_only     no transaction and nothing above that changes anything
    """
    code = _COMMENT_RE.sub(u'', _text(code))
    tx_names = [m.group(2).strip() for m in _TX_NAME_RE.finditer(code)]
    tx_names += [m.group(2).strip() for m in _TX_START_RE.finditer(code)]
    creates = _unique([m.group(1) or m.group(2)
                       for m in _CREATE_RE.finditer(code)
                       if (m.group(1) or m.group(2)) not in _NOT_ELEMENTS])
    actions = [text for rx, text in _ACTIONS_RE
               if (bool(creates) if rx is None else _matches(rx, code))]
    outside = [text for rx, text in _OUTSIDE_RE if rx.search(code)]
    has_tx = bool(_HAS_TX_RE.search(code))
    return {
        'transactions': _unique(tx_names),
        'has_transaction': has_tx,
        'actions': actions,
        'ui_actions': [text for rx, text in _UI_ACTIONS_RE if rx.search(code)],
        'outside': outside,
        'creates': creates,
        'categories': _unique([_pretty_category(m.group(1))
                               for m in _CATEGORY_RE.finditer(code)]),
        'classes': _unique([m.group(1) for m in _CLASS_RE.finditer(code)]),
        'parameters': _unique([m.group(2).strip() for m in _PARAM_RE.finditer(code)]
                              + [m.group(1) for m in _BIP_RE.finditer(code)]),
        'read_only': not (has_tx or actions or outside),
    }


def _works_on(found):
    return _unique(found['categories'] + found['classes'])


def _finding_lines(found):
    """Bullet lines (no bullet sign) for the confirmation, most serious first."""
    lines = []
    for action in found['actions']:
        if action == _CREATES and found['creates']:
            lines.append(u'Creates: {}'.format(_names(found['creates'])))
        elif action == u'Changes parameter values' and found['parameters']:
            lines.append(u'Changes parameter values ({})'.format(
                _names(found['parameters'])))
        else:
            lines.append(action)
    lines.extend(found['ui_actions'])
    works_on = _works_on(found)
    if works_on:
        lines.append(u'Works on: {}'.format(_names(works_on)))
    if (found['parameters']
            and u'Changes parameter values' not in found['actions']):
        lines.append(u'Reads parameters: {}'.format(_names(found['parameters'])))
    named = [t for t in found['transactions'] if t]
    if named:
        lines.append(u'Transaction: {}'.format(
            u', '.join(u'"{}"'.format(_clip(t, 60)) for t in named[:3])))
    elif found['has_transaction']:
        lines.append(u'Opens a transaction (changes the model)')
    for item in found['outside']:
        lines.append(u'Outside Revit: {}'.format(item))
    if found['read_only']:
        lines.append(u'No transaction found: it should only read the model')
    return lines


def _code_headline(found):
    """Headline from the code alone, for when the AI gave no description."""
    useful = [t for t in found['transactions']
              if t.strip().lower() not in _GENERIC_TX]
    if useful:
        return _clip(u', '.join(useful[:3]), _HEADLINE_LIMIT), SOURCE_TRANSACTION
    works_on = _names(_works_on(found), 4)
    if found['actions']:
        parts = [found['actions'][0]] + [a.lower() for a in found['actions'][1:2]]
        head = u', '.join(parts) + (u' — ' + works_on if works_on else u'')
    else:
        parts = [u'Reads ' + works_on if works_on else u'Reads the model']
        parts += [a.lower() for a in found['ui_actions'][:1]]
        parts += found['outside'][:1]
        head = u', '.join(parts)
    return _clip(head, _HEADLINE_LIMIT), SOURCE_CODE


def describe_code(arguments):
    """{'headline', 'source', 'findings'} for one send_code_to_revit call.

    headline: the AI's `description` when it gave one, else the transaction
    name, else T3Lab's reading of the code. findings always come from the
    code itself.
    """
    arguments = arguments if isinstance(arguments, dict) else {}
    found = code_findings(arguments.get('code') or u'')
    purpose = _clip(arguments.get('description') or u'', _HEADLINE_LIMIT)
    if purpose:
        headline, source = purpose, SOURCE_AI
    else:
        headline, source = _code_headline(found)
    return {'headline': headline.rstrip(u'.'), 'source': source,
            'findings': _finding_lines(found)}


# ── Every other tool ────────────────────────────────────────────────────────

def _target(args, noun=u'element'):
    """Which elements an element_ids / category / selection tool acts on."""
    ids = _as_list(args.get('element_ids'))
    if ids:
        return u'{} (id {})'.format(_plural(len(ids), noun), _names(ids, 5))
    category = _str_arg(args, 'category')
    if category:
        return u'all {} in the active view'.format(category)
    return u'the selected elements'


def _generic(tool_name, args):
    """'Create level — elevation: 9, name: Level 3' for tools with no own text."""
    title = _text(tool_name).replace(u'_', u' ').strip()
    title = title[:1].upper() + title[1:]
    op = _str_arg(args, 'operation')
    if op:
        title += u': ' + op.replace(u'_', u' ')
    parts = []
    for key in sorted(args):
        value = args[key]
        if key == 'operation' or value is None or value == u'' or value == []:
            continue
        if isinstance(value, bool):
            shown = u'yes' if value else u'no'
        elif isinstance(value, (list, tuple)):
            shown = (_names(value, 3) if len(value) <= 3
                     else _plural(len(value), u'item'))
        elif isinstance(value, dict):
            shown = u'{…}'
        else:
            shown = _clip(value, _VALUE_LIMIT)
        parts.append(u'{}: {}'.format(key.replace(u'_', u' '), shown))
    if len(parts) > 4:
        parts = parts[:4] + [u'…']
    return title + (u' — ' + u', '.join(parts) if parts else u'')


def _delete_element(a):
    ids = _as_list(a.get('element_ids'))
    what = (u'{} (id {})'.format(_plural(len(ids), u'element'), _names(ids, 5))
            if ids else u'elements (no id given)')
    if a.get('dry_run'):
        return (u'List what deleting {} would remove (dry run, nothing is '
                u'deleted)'.format(what), [])
    return (u'Delete {}'.format(what),
            [u'Elements that depend on them (tags, dimensions, hosted '
             u'elements) are deleted too.'])


def _edit_elements(a):
    op = _str_arg(a, 'operation').lower()
    target = _target(a)
    if op == 'ungroup':
        return u'Ungroup {}'.format(target), []
    if op == 'group':
        return u'Group {}'.format(target), []
    if op == 'mirror':
        return (u'{} {} across the {} axis'.format(
            u'Mirror a copy of' if a.get('copy') else u'Mirror', target,
            _str_arg(a, 'axis') or u'x'), [])
    if op == 'change_type':
        new_type = u': '.join(v for v in (_str_arg(a, 'family_name'),
                                          _str_arg(a, 'type_name')) if v)
        return u'Change {} to type {}'.format(target, new_type or u'?'), []
    return _generic('edit_elements', a), []


def _manage_view_template(a):
    op = _str_arg(a, 'operation').lower()
    name = _str_arg(a, 'template_name') or u'id {}'.format(a.get('template_id'))
    if op == 'delete':
        if a.get('force'):
            return (u'Delete view template "{}", even if views use it'.format(name),
                    [u'Every view that uses it loses the template.'])
        return u'Delete view template "{}"'.format(name), []
    if op == 'rename':
        return u'Rename view template "{}" to "{}"'.format(
            name, _str_arg(a, 'new_name')), []
    if op == 'duplicate':
        return u'Duplicate view template "{}" as "{}"'.format(
            name, _str_arg(a, 'new_name')), []
    return _generic('manage_view_template', a), []


def _manage_links(a):
    op = _str_arg(a, 'operation').lower()
    links = _as_list(a.get('link_names')) or [
        u'id {}'.format(i) for i in _as_list(a.get('link_ids'))]
    names = _names(links, 4) or u'?'
    if op == 'delete':
        return (u'Delete link: {}'.format(names),
                [u'Tags and dimensions that depend on the link are deleted too.'])
    if op in ('reload', 'unload', 'pin', 'unpin'):
        return u'{} link: {}'.format(op.capitalize(), names), []
    return _generic('manage_links', a), []


def _manage_document(a):
    op = _str_arg(a, 'operation').lower()
    if op == 'save':
        return u'Save the model', []
    if op == 'save_as':
        head = u'Save the model as {}'.format(
            _clip(a.get('path') or u'?', 120))
        if a.get('overwrite'):
            head += u', replacing the existing file'
        return head, []
    if op == 'sync_with_central':
        notes = [u'Publishes your work to the whole team; can run for minutes.']
        if a.get('relinquish', True):
            notes.append(u'Gives back the worksets you own afterwards.')
        return (u'Sync with central, comment "{}"'.format(
            _clip(a.get('comment') or u'', 80)), notes)
    return _generic('manage_document', a), []


def _purge_unused(a):
    if a.get('dry_run', True):
        return (u'List unused family types and view templates (dry run, '
                u'nothing is deleted)', [])
    return (u'Delete every unused family type and view template', [])


def _close_document(a):
    doc = _clip(a.get('path_or_title') or u'?', 80)
    if a.get('save'):
        return u'Save and close "{}"'.format(doc), []
    return (u'Close "{}" without saving'.format(doc),
            [u'Changes made since the last save are lost.'])


def _check_bad_geometry(a):
    if a.get('deep_probe'):
        return (u'Deep-scan geometry for export crashes',
                [u'Calls Face.Triangulate(), the same call that crashes '
                 u'exports: it can close Revit without saving.'])
    return _generic('check_bad_geometry', a), []


def _run_batch(a):
    steps = [s for s in _as_list(a.get('steps')) if isinstance(s, dict)]
    lines = []
    for i, step in enumerate(steps):
        head, _details = summarize(step.get('name', u'?'),
                                   step.get('arguments') or {})
        lines.append(u'{}. {}'.format(i + 1, head))
    return (u'Run {} as one undo step'.format(_plural(len(steps), u'step')),
            lines)


def _set_parameter(a):
    return (u'Set {} = "{}" on element {}'.format(
        _str_arg(a, 'parameter_name'), _str_arg(a, 'value'),
        a.get('element_id')), [])


def _bulk_set_parameter(a):
    ids = _as_list(a.get('element_ids'))
    if ids:
        where = _plural(len(ids), u'element')
    else:
        where = u'all {}'.format(_str_arg(a, 'category') or u'elements')
        if a.get('filter_parameter'):
            where += u' where {} = "{}"'.format(_str_arg(a, 'filter_parameter'),
                                                _str_arg(a, 'filter_value'))
    return u'Set {} = "{}" on {}'.format(
        _str_arg(a, 'parameter_name'), _str_arg(a, 'value'), where), []


def _rename_element(a):
    return u'Rename element {} to "{}"'.format(
        a.get('element_id'), _str_arg(a, 'new_name')), []


def _offset(a):
    return u', '.join(u'{}={}'.format(k, a.get(k)) for k in ('dx', 'dy', 'dz')
                      if a.get(k) not in (None, u''))


def _move_elements(a):
    return u'Move {} by {}'.format(_target(a), _offset(a) or u'?'), []


def _copy_elements(a):
    return u'Copy {}, offset {}'.format(_target(a), _offset(a) or u'?'), []


def _rotate_element(a):
    return u'Rotate {} by {}°'.format(_target(a), a.get('angle_degrees')), []


def _send_code(a):
    info = describe_code(a)
    return info['headline'], info['findings']


_DESCRIBERS = {
    CODE_TOOL: _send_code,
    'delete_element': _delete_element,
    'edit_elements': _edit_elements,
    'manage_view_template': _manage_view_template,
    'manage_links': _manage_links,
    'manage_document': _manage_document,
    'purge_unused': _purge_unused,
    'close_document': _close_document,
    'check_bad_geometry': _check_bad_geometry,
    'run_batch': _run_batch,
    'set_parameter': _set_parameter,
    'bulk_set_parameter': _bulk_set_parameter,
    'rename_element': _rename_element,
    'move_elements': _move_elements,
    'copy_elements': _copy_elements,
    'rotate_element': _rotate_element,
}


def summarize(tool_name, arguments):
    """(headline, details) for one tool call.

    headline: one line saying what the call does ("Delete 3 elements (id
    101, 102, 103)"). details: short lines that back it up -- for code, what
    T3Lab found in it; for run_batch, one line per step. Never raises.
    """
    args = arguments if isinstance(arguments, dict) else {}
    try:
        describer = _DESCRIBERS.get(tool_name)
        if describer is not None:
            headline, details = describer(args)
        else:
            headline, details = _generic(tool_name, args), []
        return _clip(headline, 2 * _HEADLINE_LIMIT), list(details)
    except Exception:
        return _text(tool_name), []
