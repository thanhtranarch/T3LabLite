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
  mcp_toolset              "full"  Which tools external clients see:
                                   "full" (all), "standard" (no teaching /
                                   training tools), "documentation" (read
                                   tools + views, sheets, annotation, export).
                                   Fewer tools = fewer tokens per request for
                                   clients without tool search.
"""

from __future__ import unicode_literals

import json

from core import call_summary

# Newest first. The server is "dual-era" (MCP 2026-07-28, Versioning):
#  - MODERN versions carry the version, client info and capabilities in every
#    request's params._meta and have no handshake; `server/discover` tells a
#    client what this server speaks, and a version not listed here gets
#    UnsupportedProtocolVersionError.
#  - LEGACY versions open with `initialize`, which answers with the client's
#    version when it is one of these, otherwise with the newest legacy one --
#    never with a modern one, which has no `initialize`.
MODERN_PROTOCOLS = ('2026-07-28',)
LEGACY_PROTOCOLS = ('2025-11-25', '2025-06-18', '2025-03-26', '2024-11-05')
SUPPORTED_PROTOCOLS = MODERN_PROTOCOLS + LEGACY_PROTOCOLS
LATEST_PROTOCOL = SUPPORTED_PROTOCOLS[0]
LATEST_LEGACY_PROTOCOL = LEGACY_PROTOCOLS[0]

# params._meta keys of a modern request / result (MCP 2026-07-28).
META_PROTOCOL = 'io.modelcontextprotocol/protocolVersion'
META_CLIENT_INFO = 'io.modelcontextprotocol/clientInfo'
META_CLIENT_CAPS = 'io.modelcontextprotocol/clientCapabilities'
META_SERVER_INFO = 'io.modelcontextprotocol/serverInfo'

# JSON-RPC error codes. -32020..-32099 belong to the MCP spec.
ERR_HEADER_MISMATCH = -32020
ERR_UNSUPPORTED_VERSION = -32022
ERR_INVALID_PARAMS = -32602
# Legacy (2025-11-25 and earlier) "resource not found"; modern uses -32602.
ERR_LEGACY_RESOURCE_NOT_FOUND = -32002

SERVER_NAME = 'T3Lab Revit MCP Server'

# How long a client may reuse a list result before asking again (modern
# CacheableResult). Short: the read-only switch and the toolset in MCP Control
# change tools/list, and a user's new wiki page changes resources/list.
LIST_TTL_MS = 60000
# Results describe this machine's Revit and policy: never for shared caches.
CACHE_SCOPE = 'private'
CACHEABLE_METHODS = frozenset(['tools/list', 'resources/list', 'resources/read',
                               'resources/templates/list', 'prompts/list',
                               'server/discover'])

# What the server offers. listChanged stays False: the server never pushes
# (no subscriptions/listen); clients re-list after LIST_TTL_MS instead.
SERVER_CAPABILITIES = {
    'tools': {'listChanged': False},
    'resources': {'listChanged': False, 'subscribe': False},
    'prompts': {'listChanged': False},
}

READ_ONLY_KEY = 'mcp_read_only'
CONFIRM_KEY = 'mcp_confirm_destructive'
CODE_KEY = 'mcp_code_execution'
TOOLSET_KEY = 'mcp_toolset'

CODE_OFF = 'off'
CODE_ASK = 'ask'
CODE_ALLOW = 'allow'
CODE_MODES = (CODE_OFF, CODE_ASK, CODE_ALLOW)

TOOLSET_FULL = 'full'
TOOLSET_STANDARD = 'standard'
TOOLSET_DOCUMENTATION = 'documentation'
TOOLSETS = (TOOLSET_FULL, TOOLSET_STANDARD, TOOLSET_DOCUMENTATION)

DEFAULT_POLICY = {
    'read_only': False,
    'confirm_destructive': True,
    'code_execution': CODE_ASK,
    'toolset': TOOLSET_FULL,
}

# Teaching / fine-tuning controls and the connection smoke test: only the
# owner's "teach then train" sessions need them. Hidden by every toolset but
# "full".
TRAINING_TOOLS = frozenset([
    't3lab_begin_teaching', 't3lab_end_teaching', 't3lab_set_teaching_mode',
    't3lab_mark_sandbox', 't3lab_build_exemplars', 't3lab_train_model',
    't3lab_training_status', 'say_hello',
])

# The "documentation" toolset: every read tool plus what finishing and issuing
# drawings needs (views, sheets, annotation, schedules, revisions, export).
# lite_guard/test_mcp_server.py checks every name is a real tool.
DOCUMENTATION_TOOLS = frozenset([
    'add_view_to_sheet', 'apply_view_template', 'bulk_set_parameter',
    'color_elements', 'create_detail_annotation', 'create_dimension',
    'create_schedule', 'create_sheet', 'create_text_note', 'create_view',
    'create_view_filter', 'duplicate_view', 'export_dwg', 'export_image',
    'export_sheets_pdf', 'manage_revision', 'manage_sheet', 'manage_view',
    'manage_view_template', 'place_views_on_sheets', 'rename_element',
    'revit_override_color', 'run_batch', 'set_parameter', 'tag_all_rooms',
    'tag_all_walls', 'tag_elements',
])

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
LOCK_FREE_METHODS = frozenset(['initialize', 'ping', 'tools/list',
                               'server/discover', 'resources/list',
                               'resources/templates/list', 'resources/read',
                               'prompts/list', 'prompts/get'])

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
    """The protocol version to answer `initialize` with (legacy handshake)."""
    return requested if requested in LEGACY_PROTOCOLS else LATEST_LEGACY_PROTOCOL


def server_version():
    """The extension version (version.txt) -- serverInfo.version."""
    import io
    import os
    path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__)))), 'version.txt')
    try:
        with io.open(path, 'r', encoding='utf-8') as f:
            return f.read().strip() or u'0.0.0'
    except Exception:
        return u'0.0.0'


def server_info():
    return {'name': SERVER_NAME, 'version': server_version()}


def request_meta(request):
    """params._meta of a JSON-RPC request, or {}."""
    params = request.get('params') if isinstance(request, dict) else None
    meta = params.get('_meta') if isinstance(params, dict) else None
    return meta if isinstance(meta, dict) else {}


def modern_version(request):
    """The protocol version a MODERN request declares in params._meta, or
    None for a legacy request (one that relies on the initialize handshake).
    A declared but unsupported version is returned as is -- check it with
    version_error()."""
    meta = request_meta(request)
    if META_PROTOCOL not in meta:
        return None
    return u'{}'.format(meta.get(META_PROTOCOL) or u'')


def version_error(request_id, requested):
    """UnsupportedProtocolVersionError response for a modern request whose
    version this server does not speak; None when it does."""
    if requested in MODERN_PROTOCOLS:
        return None
    return {'jsonrpc': '2.0', 'id': request_id,
            'error': {'code': ERR_UNSUPPORTED_VERSION,
                      'message': 'Unsupported protocol version',
                      'data': {'supported': list(SUPPORTED_PROTOCOLS),
                               'requested': requested}}}


def header_mismatch(request, header_version):
    """HeaderMismatchError response when the MCP-Protocol-Version HTTP header
    and the request's _meta disagree; None otherwise (or when either is
    missing)."""
    declared = modern_version(request)
    header_version = u'{}'.format(header_version or u'').strip()
    if not declared or not header_version or declared == header_version:
        return None
    return {'jsonrpc': '2.0', 'id': request.get('id'),
            'error': {'code': ERR_HEADER_MISMATCH,
                      'message': ('MCP-Protocol-Version header ({}) does not '
                                  'match the request _meta ({})').format(
                                      header_version, declared)}}


def discover_result(instructions=None, info=None):
    """server/discover result (before finish_result adds resultType/_meta)."""
    result = {'supportedVersions': list(SUPPORTED_PROTOCOLS),
              'capabilities': json.loads(json.dumps(SERVER_CAPABILITIES))}
    if instructions:
        result['instructions'] = instructions
    if info:
        result['_meta'] = {META_SERVER_INFO: info}
    return result


def initialize_result(requested, instructions=None, info=None):
    """initialize result for a legacy client."""
    result = {'protocolVersion': negotiate_protocol(requested),
              'capabilities': json.loads(json.dumps(SERVER_CAPABILITIES)),
              'serverInfo': info or server_info()}
    if instructions:
        result['instructions'] = instructions
    return result


def finish_result(response, method, modern, info=None):
    """Add the fields a MODERN client expects on every result: resultType,
    serverInfo in _meta, and ttlMs/cacheScope on cacheable ones. Legacy
    responses and errors pass through untouched. Returns the response."""
    if not modern or not isinstance(response, dict):
        return response
    result = response.get('result')
    if not isinstance(result, dict):
        return response
    result.setdefault('resultType', 'complete')
    meta = result.get('_meta')
    if not isinstance(meta, dict):
        meta = {}
        result['_meta'] = meta
    meta.setdefault(META_SERVER_INFO, info or server_info())
    if method in CACHEABLE_METHODS:
        result.setdefault('ttlMs', LIST_TTL_MS)
        result.setdefault('cacheScope', CACHE_SCOPE)
    return response


def resource_not_found(request_id, uri, modern):
    """Error for resources/read of an unknown URI: -32602 in the modern spec,
    -32002 before it."""
    return {'jsonrpc': '2.0', 'id': request_id,
            'error': {'code': (ERR_INVALID_PARAMS if modern
                               else ERR_LEGACY_RESOURCE_NOT_FOUND),
                      'message': 'Resource not found',
                      'data': {'uri': uri}}}


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
        'toolset': _toolset(settings.get(TOOLSET_KEY)),
    }


def _toolset(value):
    name = u'{}'.format(value or TOOLSET_FULL).strip().lower()
    return name if name in TOOLSETS else TOOLSET_FULL


def toolset_allows(tool_name, toolset, read_only_names):
    """Whether the active toolset includes `tool_name`."""
    toolset = _toolset(toolset)
    if toolset == TOOLSET_FULL:
        return True
    if tool_name in TRAINING_TOOLS:
        return False
    if toolset == TOOLSET_DOCUMENTATION:
        return tool_name in read_only_names or tool_name in DOCUMENTATION_TOOLS
    return True


def batch_outside_toolset(arguments, toolset, read_only_names):
    """Sorted step tool names of a run_batch call that the toolset hides --
    run_batch must not be a way around it."""
    steps = (arguments or {}).get('steps') if isinstance(arguments, dict) else None
    names = set()
    for step in steps if isinstance(steps, list) else []:
        name = step.get('name') if isinstance(step, dict) else None
        if name and not toolset_allows(name, toolset, read_only_names):
            names.add(name)
    return sorted(names)


def load_policy():
    """Current policy, read fresh from mcp_paths.json on every call so a
    change in MCP Control applies to the next call without a restart."""
    try:
        from core import paths
        return policy_from_settings(paths.load_settings())
    except Exception:
        return dict(DEFAULT_POLICY)


def save_policy(read_only=None, confirm_destructive=None, code_execution=None,
                toolset=None):
    """Persist the given fields; the others keep their stored values."""
    from core import paths
    data = paths.load_settings() or {}
    if toolset is not None:
        name = u'{}'.format(toolset).strip().lower()
        if name not in TOOLSETS:
            raise ValueError('toolset must be one of {}'.format(', '.join(TOOLSETS)))
        data[TOOLSET_KEY] = name
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
    return toolset_allows(tool_name, policy.get('toolset'), read_only_names)


def check_call(tool_name, policy, read_only_names, destructive):
    """Decide one external tools/call before it runs.

    Returns (decision, message): decision is 'run', 'confirm' or 'block';
    message explains a block to the caller.
    """
    if not toolset_allows(tool_name, policy.get('toolset'), read_only_names):
        return 'block', toolset_message(tool_name, policy.get('toolset'))
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


def toolset_message(tool_name, toolset):
    return ('{} is not in the "{}" toolset chosen on this machine (MCP Control '
            '> Advanced > Security). Use the listed tools, or ask the user to '
            'pick the "full" toolset.'.format(tool_name, _toolset(toolset)))


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
