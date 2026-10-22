# -*- coding: utf-8 -*-
"""
MCP policy for EXTERNAL clients (Claude Desktop, Codex, Cursor ... via
bridge.py): what they may see, what they may call, and what needs a "Yes" in
Revit first.

The in-app T3Lab Assistant does not come through here. It calls
T3LabAIServer._execute_tool directly and has its own gate (ToolGate in
Intelligence/agent_loop.py), so nothing below changes the Assistant.

Plain Python, no Revit import: lite_guard/test_mcp_server.py runs it on CI.

Settings (in %APPDATA%\\T3LabAI\\mcp_paths.json, also set from MCP Control >
Advanced > Security):

  mcp_read_only            false   true = external clients see and call read
                                   tools only; nothing can change the model.
  mcp_confirm_destructive  true    Ask in Revit before delete / purge /
                                   ungroup / save as / sync / close without
                                   saving, and before a batch that does any
                                   of those.
  mcp_code_execution       "ask"   send_code_to_revit: "off" (hidden and
                                   refused), "ask" (Revit shows the code and
                                   asks each time), "allow" (runs at once).
"""

from __future__ import unicode_literals

import json

from core import call_summary

# Newest first. initialize answers with the client's version when it is one
# of these, otherwise with the newest one -- the spec's negotiation rule.
SUPPORTED_PROTOCOLS = ('2025-06-18', '2025-03-26', '2024-11-05')
LATEST_PROTOCOL = SUPPORTED_PROTOCOLS[0]

READ_ONLY_KEY = 'mcp_read_only'
CONFIRM_KEY = 'mcp_confirm_destructive'
CODE_KEY = 'mcp_code_execution'

CODE_OFF = 'off'
CODE_ASK = 'ask'
CODE_ALLOW = 'allow'
CODE_MODES = (CODE_OFF, CODE_ASK, CODE_ALLOW)

DEFAULT_POLICY = {
    'read_only': False,
    'confirm_destructive': True,
    'code_execution': CODE_ASK,
}

# Tools that run caller-supplied code.
CODE_TOOLS = frozenset(['send_code_to_revit'])

# Tools that may delete or overwrite something that cannot simply be redone:
# destructiveHint=true in tools/list. Every other modifying tool only adds or
# edits (destructiveHint=false). Which CALLS need a "Yes" is decided per call
# by T3LabAIServer.is_destructive(), which also looks at the arguments.
DESTRUCTIVE_HINT_TOOLS = frozenset([
    'delete_element', 'edit_elements', 'manage_view_template', 'manage_links',
    'manage_document', 'purge_unused', 'close_document', 'send_code_to_revit',
    'run_batch',
])

# Answered by the server without the Revit request lock: they never touch
# the Revit API, so they must not queue behind a slow tools/call (a
# tools/list stuck there ran past the bridge's 15 s list timeout and the
# client got the stale cache).
LOCK_FREE_METHODS = frozenset(['initialize', 'ping', 'tools/list'])

# How much of the code / arguments the confirmation dialog shows.
_CONFIRM_TEXT_LIMIT = 1800

# `instructions` of the initialize result (MCP 2025-03-26+; older clients
# ignore the field). Clients that support it (Claude Desktop, Claude Code) put
# it in the model's system prompt, so this is how an AI app learns that the
# drawing rules exist before its first tool call. bridge.py answers initialize
# itself and keeps a copy (it cannot import lib/); lite_guard/test_mcp_server.py
# checks the two stay identical.
SERVER_INSTRUCTIONS = (
    'T3Lab drives the Revit model open on this machine.\n'
    'Drawings (sheets, views, dimensions, tags, notes, revisions, PDF/DWG '
    'issue):\n'
    '1. Before you create, annotate, complete or issue drawings, read the rules '
    'with t3lab_wiki: page "drawing-completion-rules" first, then the page for '
    'the drawing type (architectural-drawing-rules, structural-drawing-rules, '
    'mep-drawing-rules), and "drawing-issue-checklist" before an issue. They '
    'say what a finished sheet must show and which tools to use.\n'
    '2. t3lab_wiki also holds construction knowledge (scales, line weights, '
    'symbols, title block, rebar notation, design stages, Vietnamese-English '
    'terms): search it instead of guessing.\n'
    '3. The project\'s own BEP / standard wins over the wiki. Never invent '
    'project codes, sheet numbers or revision codes: follow what the model '
    'already shows and say you inferred it.\n'
    '4. Read before you write: list what exists, show the plan with targets and '
    'counts before a batch change, then check the result with a read tool.\n'
    '5. Report what was done and what was not, with element ids, in the '
    'user\'s language.')


def negotiate_protocol(requested):
    """The protocol version to answer `initialize` with."""
    return requested if requested in SUPPORTED_PROTOCOLS else LATEST_PROTOCOL


def _as_bool(value, default):
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    text = u'{}'.format(value).strip().lower()
    if text in ('1', 'true', 'yes', 'on'):
        return True
    if text in ('0', 'false', 'no', 'off'):
        return False
    return default


def policy_from_settings(settings):
    """Policy dict from a loaded mcp_paths.json dict. Bad values -> defaults."""
    settings = settings if isinstance(settings, dict) else {}
    code = u'{}'.format(settings.get(CODE_KEY) or CODE_ASK).strip().lower()
    if code not in CODE_MODES:
        code = CODE_ASK
    return {
        'read_only': _as_bool(settings.get(READ_ONLY_KEY),
                              DEFAULT_POLICY['read_only']),
        'confirm_destructive': _as_bool(settings.get(CONFIRM_KEY),
                                        DEFAULT_POLICY['confirm_destructive']),
        'code_execution': code,
    }


def load_policy():
    """Current policy, read fresh from mcp_paths.json on every call so a
    change in MCP Control applies to the next call without a restart."""
    try:
        from core import paths
        return policy_from_settings(paths.load_settings())
    except Exception:
        return dict(DEFAULT_POLICY)


def save_policy(read_only=None, confirm_destructive=None, code_execution=None):
    """Persist the given fields; the others keep their stored values."""
    from core import paths
    data = paths.load_settings() or {}
    if read_only is not None:
        data[READ_ONLY_KEY] = bool(read_only)
    if confirm_destructive is not None:
        data[CONFIRM_KEY] = bool(confirm_destructive)
    if code_execution is not None:
        mode = u'{}'.format(code_execution).strip().lower()
        if mode not in CODE_MODES:
            raise ValueError('code_execution must be one of {}'.format(
                ', '.join(CODE_MODES)))
        data[CODE_KEY] = mode
    paths.save_settings(data)
    return policy_from_settings(data)


def tool_visible(tool_name, policy, read_only_names):
    """Whether tools/list shows `tool_name` to external clients."""
    if tool_name in CODE_TOOLS and policy.get('code_execution') == CODE_OFF:
        return False
    if policy.get('read_only') and tool_name not in read_only_names:
        return False
    return True


def check_call(tool_name, policy, read_only_names, destructive):
    """Decide one external tools/call before it runs.

    Returns (decision, message): decision is 'run', 'confirm' or 'block';
    message explains a block to the caller.
    """
    if tool_name in CODE_TOOLS:
        mode = policy.get('code_execution')
        if mode == CODE_OFF:
            return 'block', (
                '{} is turned off on this machine (MCP Control > Advanced > '
                'Security). Use the regular T3Lab tools instead, or ask the '
                'user to allow code execution.'.format(tool_name))
        if policy.get('read_only'):
            return 'block', _read_only_message(tool_name)
        if mode == CODE_ASK:
            return 'confirm', None
        return 'run', None
    if policy.get('read_only') and tool_name not in read_only_names:
        return 'block', _read_only_message(tool_name)
    if destructive and policy.get('confirm_destructive', True):
        return 'confirm', None
    return 'run', None


def _read_only_message(tool_name):
    return ('Read-only mode is on (MCP Control > Advanced > Security), so {} '
            'was not run and the model was not changed. Tell the user what '
            'you wanted to do; they can turn read-only mode off.'.format(tool_name))


def tool_annotations(tool_name, read_only_names):
    """MCP tool annotations (spec 2025-03-26+). Hints only -- clients use them
    to decide when to ask the user; enforcement stays in check_call."""
    read_only = tool_name in read_only_names
    ann = {
        'readOnlyHint': read_only,
        # Everything here acts on the local Revit session, nothing on the web.
        'openWorldHint': False,
    }
    if not read_only:
        ann['destructiveHint'] = tool_name in DESTRUCTIVE_HINT_TOOLS
    return ann


def list_entry(tool, read_only_names):
    """A copy of one registry entry with annotations added. The registry dict
    itself is left alone: the Assistant converts it for Anthropic/OpenAI/
    Ollama, which do not take an `annotations` key."""
    entry = dict(tool)
    ann = dict(entry.get('annotations') or {})
    for key, value in tool_annotations(entry.get('name', ''), read_only_names).items():
        ann.setdefault(key, value)
    entry['annotations'] = ann
    return entry


def is_error_result(result):
    """True when a tool's own result dict reports failure, so tools/call can
    set isError and the client sees a failed call, not a successful one."""
    if not isinstance(result, dict):
        return False
    if result.get('success') is False:
        return True
    err = result.get('error')
    return bool(err) and result.get('success') is not True


def _clip(text, limit=_CONFIRM_TEXT_LIMIT):
    text = u'{}'.format(text or u'')
    if len(text) <= limit:
        return text
    return text[:limit] + u'\n... ({} more characters not shown)'.format(len(text) - limit)


def _args_text(arguments):
    try:
        return json.dumps(arguments or {}, indent=2, ensure_ascii=False, sort_keys=True)
    except Exception:
        return u'{}'.format(arguments)


_CODE_SOURCE_NOTE = {
    call_summary.SOURCE_AI: u'(Purpose as stated by the AI.)',
    call_summary.SOURCE_TRANSACTION: u'(Purpose taken from the transaction '
                                     u'name in the code.)',
    call_summary.SOURCE_CODE: u'(The AI gave no purpose; this is T3Lab\'s '
                              u'reading of the code.)',
}


def _bullets(lines, limit=12):
    shown = [u'• {}'.format(line) for line in lines[:limit]]
    if len(lines) > limit:
        shown.append(u'• ... {} more'.format(len(lines) - limit))
    return u'\n'.join(shown)


def confirm_text(tool_name, arguments, client=None):
    """(main_instruction, content, expanded) for the Revit confirmation.

    Says WHAT the call does (core.call_summary): for code, its purpose and
    what T3Lab found in it -- not how many lines it has.
    """
    who = client or u'An AI app connected through MCP'
    arguments = arguments if isinstance(arguments, dict) else {}
    if tool_name in CODE_TOOLS:
        info = call_summary.describe_code(arguments)
        content = [u'{} wants to run Python code with full access to the '
                   u'open model.'.format(who),
                   _CODE_SOURCE_NOTE.get(info['source'], u'')]
        if info['findings']:
            content += [u'', u'Found in the code by T3Lab:',
                        _bullets(info['findings'])]
        content += [u'', u'Click "Show details" to read the code. Run it only '
                         u'if you trust what it does.']
        return (u'Run AI code in Revit: {}?'.format(info['headline']),
                u'\n'.join(content),
                _clip(arguments.get('code') or u''))
    headline, details = call_summary.summarize(tool_name, arguments)
    expanded = _clip(u'Tool: {}\n{}'.format(tool_name, _args_text(arguments)))
    if tool_name == 'run_batch':
        return (u'{}, including a delete or other destructive change?'.format(
                    headline),
                u'{} wants to run these steps as one undo step:\n{}'.format(
                    who, _clip(u'\n'.join(details), 600)),
                expanded)
    content = u'{} wants to do this on the open model.'.format(who)
    if details:
        content += u'\n\n' + _bullets(details)
    content += (u'\n\nThis can delete or overwrite work. Click "Show details" '
                u'to see the exact request.')
    return u'{}?'.format(headline), content, expanded


def origin_allowed(origin):
    """DNS-rebinding guard: a request carrying an Origin header must come from
    a page on this machine. Native clients (the bridge, curl) send none."""
    if not origin:
        return True
    origin = u'{}'.format(origin).strip().lower()
    for scheme in (u'http://', u'https://'):
        if origin.startswith(scheme):
            host = origin[len(scheme):].split(u'/', 1)[0]
            if host.startswith(u'['):           # [::1]:port
                host = host.split(u']', 1)[0] + u']'
            else:
                host = host.split(u':', 1)[0]
            return host in (u'localhost', u'127.0.0.1', u'[::1]')
    return False


def token_matches(header_value, token):
    """Constant-time check of an `Authorization: Bearer <token>` header."""
    if not token:
        return False
    expected = u'Bearer ' + token
    given = u'{}'.format(header_value or u'')
    try:
        import hmac
        compare = getattr(hmac, 'compare_digest', None)
        if compare is not None:
            return compare(given.encode('utf-8'), expected.encode('utf-8'))
    except Exception:
        pass
    if len(given) != len(expected):
        return False
    diff = 0
    for a, b in zip(given, expected):
        diff |= ord(a) ^ ord(b)
    return diff == 0
