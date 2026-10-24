# -*- coding: utf-8 -*-
"""
MCP prompts for EXTERNAL clients (Claude Desktop, Claude Code ... via
bridge.py): the T3Lab workflows a user picks by name -- "/complete-drawings"
in Claude Code, the prompt menu in Claude Desktop -- instead of typing the
whole procedure.

Each prompt is a user message that points the model at the right wiki pages
(core.mcp_wiki) and T3Lab tools, in the order the drawing rules ask for:
read the rules, list what exists, show the plan, act, check. Nothing here
touches Revit; the server answers prompts/list and prompts/get on the HTTP
thread.

Plain Python, IronPython 2.7 + CPython 3. lite_guard/test_mcp_server.py checks
that every tool and wiki page a prompt names is real.

Author: Tran Tien Thanh
"""

from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__ = "MCP Prompts"

from core import mcp_wiki

_REPORT = ('Report what was done and what was not, with element ids, in the '
           'user\'s language.')

PROMPTS = [
    {
        'name': 'complete-drawings',
        'title': 'Complete drawings',
        'description': ('Finish sheets/views so they meet the T3Lab drawing '
                        'rules: dimensions, tags, notes, titles, scales.'),
        'arguments': [
            {'name': 'scope', 'required': False,
             'description': ('What to finish, e.g. "sheet A-101", "all level '
                             'floor plans". Default: the active view or sheet.')},
            {'name': 'discipline', 'required': False,
             'description': 'architectural, structural or mep (default: from the sheets).'},
        ],
        'tools': ['t3lab_wiki', 'revit_get_active_view', 'revit_list_sheets',
                  'revit_list_views', 'get_current_view_elements', 'run_batch',
                  'create_dimension', 'tag_elements', 'create_text_note'],
        'pages': [mcp_wiki.START_PAGE, 'architectural-drawing-rules',
                  'structural-drawing-rules', 'mep-drawing-rules'],
        'template': (
            'Complete these drawings in the open Revit model: {scope}.\n'
            'Discipline: {discipline}.\n\n'
            '1. Read the rules first with t3lab_wiki: page "drawing-completion-'
            'rules", then the page for the discipline (architectural-drawing-'
            'rules, structural-drawing-rules or mep-drawing-rules).\n'
            '2. List what exists (revit_get_active_view, revit_list_sheets, '
            'revit_list_views, get_current_view_elements) and compare it with '
            'the rules: missing dimensions, tags, notes, view titles, scales.\n'
            '3. Show me the plan with targets and counts before you change '
            'anything.\n'
            '4. Make the changes with the T3Lab tools (create_dimension, '
            'tag_elements, create_text_note ...); group many edits in run_batch '
            'so they are one undo step.\n'
            '5. Check the result with a read tool. ' + _REPORT),
    },
    {
        'name': 'issue-drawings',
        'title': 'Issue drawings',
        'description': ('Check sheets against the issue checklist, add the '
                        'revision, then export PDF and/or DWG.'),
        'arguments': [
            {'name': 'sheets', 'required': True,
             'description': 'Sheets to issue, e.g. "A-101 to A-120" or "all structural sheets".'},
            {'name': 'revision', 'required': False,
             'description': 'Revision to add, e.g. "P02 - For coordination". Default: none.'},
            {'name': 'format', 'required': False,
             'description': 'pdf, dwg or both (default: pdf).'},
        ],
        'tools': ['t3lab_wiki', 'revit_list_sheets', 'manage_revision',
                  'export_sheets_pdf', 'export_dwg', 'get_model_warnings'],
        'pages': ['drawing-issue-checklist'],
        'template': (
            'Issue these sheets: {sheets}.\n'
            'Revision: {revision}. Format: {format}.\n\n'
            '1. Read t3lab_wiki page "drawing-issue-checklist".\n'
            '2. Check every sheet against it (revit_list_sheets for numbers, '
            'names and title block data; get_model_warnings for open '
            'warnings). Never invent sheet numbers or revision codes: follow '
            'what the model already shows.\n'
            '3. List the problems you found and the sheets that pass, and ask '
            'me before you fix or skip anything.\n'
            '4. Add the revision with manage_revision when one is given.\n'
            '5. Export with export_sheets_pdf and/or export_dwg, then give me '
            'the output folder and the file count. ' + _REPORT),
    },
    {
        'name': 'model-health-check',
        'title': 'Model health check',
        'description': ('Read-only audit: warnings, bad geometry, model size, '
                        'with the fixes that matter most first.'),
        'arguments': [
            {'name': 'focus', 'required': False,
             'description': 'Optional focus, e.g. "warnings", "file size", "levels and grids".'},
        ],
        'tools': ['get_model_health', 'get_model_warnings', 'audit_model',
                  'check_bad_geometry', 'analyze_model_statistics'],
        'pages': [],
        'template': (
            'Check the health of the open Revit model. Focus: {focus}.\n\n'
            'Use read tools only -- change nothing: get_model_health, '
            'get_model_warnings, audit_model, check_bad_geometry, '
            'analyze_model_statistics.\n'
            'Then give me a short report: the biggest problems first, each with '
            'its count, why it matters and the fix (with the T3Lab tool that '
            'would do it). ' + _REPORT),
    },
]

_DEFAULTS = {
    'scope': 'the active view or sheet',
    'discipline': 'from the sheets themselves',
    'revision': 'none',
    'format': 'pdf',
    'focus': 'everything',
}


class PromptError(ValueError):
    """Bad prompts/get request -> JSON-RPC -32602 (Invalid params)."""


def _find(name):
    for prompt in PROMPTS:
        if prompt['name'] == name:
            return prompt
    return None


def prompt_list():
    """prompts/list result."""
    return {'prompts': [dict((k, prompt[k]) for k in
                             ('name', 'title', 'description', 'arguments'))
                        for prompt in PROMPTS]}


def get_prompt(name, arguments=None):
    """prompts/get result; raises PromptError on a bad name or a missing
    required argument."""
    prompt = _find(u'{}'.format(name or u''))
    if prompt is None:
        raise PromptError('Unknown prompt "{}". Known: {}.'.format(
            name, ', '.join(p['name'] for p in PROMPTS)))
    arguments = arguments if isinstance(arguments, dict) else {}
    values = {}
    for arg in prompt['arguments']:
        value = u'{}'.format(arguments.get(arg['name']) or u'').strip()
        if not value:
            if arg.get('required'):
                raise PromptError('Prompt "{}" needs "{}".'.format(
                    prompt['name'], arg['name']))
            value = _DEFAULTS.get(arg['name'], u'not given')
        values[arg['name']] = value
    return {
        'description': prompt['description'],
        'messages': [{'role': 'user',
                      'content': {'type': 'text',
                                  'text': prompt['template'].format(**values)}}],
    }
