"""T3Lab MCP stdio bridge.

Forwards JSON-RPC requests from an MCP client (Claude Desktop / Claude Code)
to the T3Lab HTTP server running inside Revit.

Deployment: this file is copied by the extension to a machine-stable
location — %APPDATA%/T3LabAI/bridge.py — and the Claude config points at
that copy (see Services/mcp_service.py, MCPService.deploy_bridge). Users can
download, move or update the extension freely; the MCP config never breaks
because it never references a path inside the extension folder.

Attach-proof protocol handling: the MCP handshake no longer requires a live
Revit instance —

  - initialize / ping      -> answered locally by the bridge, so the client
                              always attaches even when Revit is closed.
  - tools/list             -> served live from Revit when reachable (and the
                              manifest is cached to
                              %APPDATA%/T3LabAI/tools_cache.json); when Revit
                              is down the cached manifest is served instead.
  - tools/call, Revit down -> structured error payload telling the model to
                              start Revit / the T3Lab server, instead of a
                              broken session.

Multi-window (multi-instance) support: each running Revit instance hosts its
own server on the first free port in 48884-48894. The bridge tracks a
"current" port and re-routes automatically:

  - list_open_documents   -> fans out to every alive instance and merges the
                             results, so documents in other Revit windows are
                             visible to the AI client.
  - switch_active_document-> if the requested document is open in another
                             Revit window, the bridge switches its current
                             port to that instance; every subsequent tool
                             call then targets that window.
  - any other call        -> forwarded to the current port; if that instance
                             is gone (connection refused, or it stopped
                             answering /health — hung zombie socket, foreign
                             server such as pyRevit Routes on the port), the
                             bridge fails over to another alive instance.

Result-size guard: the MCP client rejects any tool result over 1 MB outright
("Tool result is too large. Maximum size is 1MB.") — it does not truncate it,
so the model receives nothing and the whole call is wasted. Every response is
therefore capped here before it goes out; see MAX_RESULT_BYTES below.

Bridge-level tools (served entirely by the bridge, injected into tools/list
on top of the in-Revit server's manifest):

  - list_revit_instances  -> probe the port range and report every running
                             Revit instance, its open documents, and which
                             instance the bridge currently targets.
  - switch_revit_instance -> explicitly retarget ALL subsequent tool calls at
                             another instance, by port or by a document open
                             in it (the by-document form delegates to the
                             cross-instance switch_active_document handling,
                             so the target window is activated too).
"""

import sys
import os
import json
import urllib.request
import urllib.error
from concurrent.futures import ThreadPoolExecutor

BRIDGE_VERSION = '2.7.0'

PORT_MIN, PORT_MAX = 48884, 48894
PROBE_TIMEOUT = 1.0    # /health probe per port (run in parallel)
TARGET_PROBE_TIMEOUT = 4.0  # explicit single-port checks get a fairer wait —
                            # a busy instance (single-threaded server mid-tool)
                            # can miss the fast range scan yet still be alive
CALL_TIMEOUT = 130     # matches the server's 120s ExternalEvent wait + margin
LIST_TIMEOUT = 15      # tools/list must answer fast or fall back to cache

# ─── Result-size guard ──────────────────────────────────────────────────────
# The client's 1 MB tool-result ceiling is a hard reject, and no client-side
# setting raises it (MAX_MCP_OUTPUT_TOKENS governs the separate 25k-token
# limit; the per-tool annotation below tops out at 500k chars). So an
# oversized payload has to be shrunk on this side: the longest lists are
# halved until the response fits, and the payload says so, giving the model
# real partial data plus a paging hint instead of an error toast.
MAX_RESULT_BYTES = 800000   # margin under the client's hard 1 MB ceiling
MIN_KEPT_ITEMS = 5          # never shrink a list below this many rows
SHRINK_PASSES = 60          # repeated halving bottoms out long before this

# Tools whose results are legitimately large (full parameter dumps, schedule
# rows, model audits). Without an annotation the client persists anything over
# its default threshold to disk and hands the model a file reference instead
# of the data; these read tools are more useful inline. Deliberately well
# under both the client's 500k-char annotation ceiling and MAX_RESULT_BYTES —
# the point is to skip the disk round-trip on a mid-size read, not to dump a
# whole model into the conversation.
LARGE_RESULT_CHARS = 200000
LARGE_RESULT_TOOLS = (
    'analyze_model_statistics',
    'audit_model',
    'get_all_parameters',
    'get_current_view_elements',
    'get_material_quantities',
    'get_model_warnings',
    'get_schedule_data',
    'query_stored_data',
    'revit_list_sheets',
    'revit_list_views',
)

# Protocol versions, newest first -- copies of core.mcp_policy (this file is
# deployed on its own and cannot import lib/; lite_guard/test_mcp_server.py
# keeps them identical). Dual-era (MCP 2026-07-28): a MODERN request carries
# its version in params._meta and needs no handshake (server/discover is
# answered here); a LEGACY client opens with initialize, answered here too.
MODERN_PROTOCOLS = ('2026-07-28',)
LEGACY_PROTOCOLS = ('2025-11-25', '2025-06-18', '2025-03-26', '2024-11-05')
SUPPORTED_PROTOCOLS = MODERN_PROTOCOLS + LEGACY_PROTOCOLS
DEFAULT_PROTOCOL = LEGACY_PROTOCOLS[0]   # initialize with an unknown version

META_PROTOCOL = 'io.modelcontextprotocol/protocolVersion'
META_SERVER_INFO = 'io.modelcontextprotocol/serverInfo'
SERVER_NAME = 'T3Lab Revit MCP Server'
LIST_TTL_MS = 60000
CACHE_SCOPE = 'private'
CACHEABLE_METHODS = frozenset(['tools/list', 'resources/list', 'resources/read',
                               'resources/templates/list', 'prompts/list',
                               'server/discover'])
SERVER_CAPABILITIES = {
    'tools': {'listChanged': False},
    'resources': {'listChanged': False, 'subscribe': False},
    'prompts': {'listChanged': False},
}

# `instructions` of the initialize result: points the AI app at the drawing
# rules (t3lab_wiki) before its first tool call. A copy of
# core.mcp_policy.SERVER_INSTRUCTIONS -- this file is deployed on its own and
# cannot import lib/. lite_guard/test_mcp_server.py keeps the two identical.
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

# Tools implemented by the bridge itself (multi-instance control lives here —
# a server inside one Revit process cannot see the other processes). Injected
# into every tools/list response on top of the server manifest; never cached.
BRIDGE_TOOLS = [
    {
        'name': 'list_revit_instances',
        'description': ('List running Revit instances. Each open Revit window is a separate '
                        'process hosting its own T3Lab server on a port in {}-{}; this reports '
                        'every alive instance, the documents open in it, and which instance '
                        'tool calls currently target. Use switch_revit_instance (by port) or '
                        'switch_active_document (by document) to retarget.'
                        ).format(PORT_MIN, PORT_MAX),
        'inputSchema': {'type': 'object', 'properties': {}, 'required': []},
    },
    {
        'name': 'switch_revit_instance',
        'description': ('Point ALL subsequent tool calls at another running Revit instance '
                        '(separate Revit window/process). Pass port from list_revit_instances, '
                        'or document (a title/file name open in the target instance) — the '
                        'by-document form also activates that document\'s window, like '
                        'switch_active_document.'),
        'inputSchema': {
            'type': 'object',
            'properties': {
                'port': {
                    'type': 'integer',
                    'description': 'Instance port from list_revit_instances ({}-{})'.format(
                        PORT_MIN, PORT_MAX),
                },
                'document': {
                    'type': 'string',
                    'description': ('Document title, file name or .rvt path open in the '
                                    'target instance (alternative to port)'),
                },
            },
            'required': [],
        },
    },
]


def _data_dir():
    """%APPDATA%/T3LabAI — same folder as mcp_token.txt / mcp_paths.json."""
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI')


def _read_token():
    """Read the shared-secret token T3LabAIServer writes on first run to
    %APPDATA%\\T3LabAI\\mcp_token.txt. All Revit instances on the machine
    share this file, so one token authenticates against every instance."""
    try:
        with open(os.path.join(_data_dir(), 'mcp_token.txt'), 'r') as f:
            return f.read().strip()
    except Exception:
        return ''


# ─── Tools-manifest cache (lets tools/list answer with Revit closed) ─────────

def _cache_file():
    return os.path.join(_data_dir(), 'tools_cache.json')


def _load_cache():
    try:
        with open(_cache_file(), 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _load_cached_list(key):
    items = _load_cache().get(key)
    return items if isinstance(items, list) else []


def _load_cached_tools():
    return _load_cached_list('tools')


def _save_cached_list(key, items):
    """Keep the last live list of `key` (tools, resources, prompts ...) so
    the matching list request still answers while Revit is closed."""
    try:
        os.makedirs(_data_dir(), exist_ok=True)
        data = _load_cache()
        data['bridge_version'] = BRIDGE_VERSION
        data[key] = items
        tmp = _cache_file() + '.tmp'
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump(data, f)
        os.replace(tmp, _cache_file())
    except Exception:
        pass


def _save_cached_tools(tools):
    _save_cached_list('tools', tools)


# Token read again from disk after a 401. The bridge starts with the client
# (Claude Desktop ...), often before Revit ever ran on this machine, so the
# token file may not exist yet at launch -- without this every call failed
# with 401 until the client was restarted.
_LATEST_TOKEN = ['']


def _post_once(port, request, token, timeout):
    # 127.0.0.1, NOT localhost: on Windows "localhost" resolves to ::1 first
    # and the IPv6 connect attempt takes ~2s to fail before falling back to
    # IPv4 — adding a flat 2s to EVERY tool call (the server binds IPv4 only).
    data = json.dumps(request).encode('utf-8')
    headers = {'Content-Type': 'application/json'}
    if token:
        headers['Authorization'] = 'Bearer ' + token
    version = _declared_version(request)
    if version:
        headers['MCP-Protocol-Version'] = version
    req = urllib.request.Request(
        f"http://127.0.0.1:{port}/mcp", data=data, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as response:
            return json.loads(response.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        # A modern-protocol error (unsupported version, header mismatch)
        # comes back as 400 with a JSON-RPC body: pass it on as is.
        if e.code == 400:
            try:
                body = json.loads(e.read().decode('utf-8'))
            except Exception:
                body = None
            if isinstance(body, dict) and body.get('jsonrpc') == '2.0':
                return body
        raise


def _post_mcp(port, request, token, timeout=CALL_TIMEOUT):
    """Forward one JSON-RPC request dict to an instance; returns parsed reply."""
    token = _LATEST_TOKEN[0] or token
    try:
        return _post_once(port, request, token, timeout)
    except urllib.error.HTTPError as e:
        if e.code != 401:
            raise
        fresh = _read_token()
        if not fresh or fresh == token:
            raise
        _LATEST_TOKEN[0] = fresh
        return _post_once(port, request, fresh, timeout)


def _is_conn_refused(exc):
    if isinstance(exc, urllib.error.URLError):
        return isinstance(exc.reason, ConnectionRefusedError)
    return isinstance(exc, ConnectionRefusedError)


def _port_alive(port, timeout=PROBE_TIMEOUT):
    """True if a T3Lab server answers /health on this port."""
    try:
        req = urllib.request.Request(f"http://127.0.0.1:{port}/health")
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.status == 200
    except urllib.error.HTTPError as e:
        # Transitional: servers built between the dead-code sweep af2a9ab
        # (2026-09-26, deleted do_GET) and its restore answer GET /health
        # with 501 "Unsupported method" while serving POST /mcp normally, and
        # a Revit window keeps that build until it restarts. Count such a
        # listener as alive — but only a Python BaseHTTP one: a zombie socket
        # still times out, and a foreign server (pyRevit Routes runs on
        # HttpListener) carries no BaseHTTP banner.
        server = (e.headers.get('Server') or '') if e.headers else ''
        return e.code == 501 and server.startswith('BaseHTTP')
    except Exception:
        return False


def _alive_ports():
    """Probe the whole port range in parallel; ~PROBE_TIMEOUT total."""
    ports = range(PORT_MIN, PORT_MAX + 1)
    with ThreadPoolExecutor(max_workers=len(ports)) as pool:
        return [p for p, ok in zip(ports, pool.map(_port_alive, ports)) if ok]


def _tool_payload(response):
    """Parse the JSON text payload of a tools/call response, or None."""
    try:
        return json.loads(response['result']['content'][0]['text'])
    except Exception:
        return None


def _tool_response(request_id, payload):
    """Build a tools/call JSON-RPC response around a payload dict."""
    return {
        'jsonrpc': '2.0',
        'id': request_id,
        'result': {'content': [{'type': 'text',
                                'text': json.dumps(payload, indent=2)}]},
    }


def _rpc_error(request_id, message, code=-32603):
    return {'jsonrpc': '2.0', 'id': request_id,
            'error': {'code': code, 'message': message}}


# ─── Keeping results under the client's 1 MB ceiling ────────────────────────

def _response_bytes(response):
    """Wire size of a response — what the client actually measures."""
    try:
        return len(json.dumps(response).encode('utf-8'))
    except Exception:
        return 0


def _longest_list(payload):
    """(owner_dict, key, list) for the longest list anywhere in the payload.
    Only lists held under a dict key are candidates — those are the ones that
    can be replaced with a shorter copy, and every server payload is a dict
    of lists. Iterative: a deeply nested payload must not blow the stack."""
    best = (None, None, None)
    stack = [payload]
    while stack:
        node = stack.pop()
        if isinstance(node, dict):
            for key, value in node.items():
                if isinstance(value, list):
                    if best[2] is None or len(value) > len(best[2]):
                        best = (node, key, value)
                    stack.append(value)
                elif isinstance(value, dict):
                    stack.append(value)
        elif isinstance(node, list):
            stack.extend(v for v in node if isinstance(v, (dict, list)))
    return best


def _shrink_once(payload, shrunk):
    """Halve the longest list in the payload, recording what was dropped.
    False when nothing is left worth shrinking."""
    owner, key, biggest = _longest_list(payload)
    if owner is None or len(biggest) <= MIN_KEPT_ITEMS:
        return False
    keep = max(MIN_KEPT_ITEMS, len(biggest) // 2)
    # First sighting of a key carries the real total — later passes see the
    # already-shortened list, so the original count must not be overwritten.
    entry = shrunk.setdefault(key, {'total': len(biggest)})
    entry['returned'] = keep
    owner[key] = biggest[:keep]
    return True


def _hard_cut(response):
    """Last resort for a payload that cannot be shrunk structurally (plain
    text, or one enormous value): cut the raw text until it fits. A cut
    result the model can read beats a rejected one it never sees."""
    try:
        item = response['result']['content'][0]
        text = item.get('text') or ''
    except Exception:
        return response
    marker = ('\n\n... [T3Lab bridge: result cut here — the MCP client rejects '
              'tool results over 1 MB. The remainder was NOT returned; re-run '
              'with narrower filters or paging (offset/count) to get it.]')
    for _ in range(SHRINK_PASSES):
        if _response_bytes(response) <= MAX_RESULT_BYTES:
            break
        text = text[:max(1000, len(text) // 2)]
        item['text'] = text + marker
    return response


def _cap_result(response):
    """Keep any outgoing response under the client's hard 1 MB result cap.
    Applied to everything the bridge writes, so a tool that forgot its own
    paging cannot cost the model the entire call."""
    if _response_bytes(response) <= MAX_RESULT_BYTES:
        return response
    try:
        item = response['result']['content'][0]
        payload = json.loads(item['text']) if item.get('type') == 'text' else None
    except Exception:
        payload = None
    if not isinstance(payload, dict):
        return _hard_cut(response)

    shrunk = {}
    while len(shrunk) < SHRINK_PASSES:
        if not _shrink_once(payload, shrunk):
            break
        item['text'] = json.dumps(payload, indent=2)
        if _response_bytes(response) <= MAX_RESULT_BYTES:
            break

    if shrunk:
        payload['truncated'] = True
        payload['truncated_lists'] = shrunk
        payload['truncation_note'] = (
            'Lists were shortened by the T3Lab bridge because the full result '
            'exceeded the MCP client 1 MB limit. Do NOT count or conclude '
            'from these lists — use the total in truncated_lists. Re-run with '
            'a narrower filter, or with offset/count paging where the tool '
            'supports it, to read the rest.')
        item['text'] = json.dumps(payload, indent=2)

    if _response_bytes(response) > MAX_RESULT_BYTES:
        return _hard_cut(response)
    return response


def _documents_on(port, token):
    """list_open_documents on one instance; [] on any failure (including
    instances still running an older build without the tool)."""
    try:
        resp = _post_mcp(port, {
            'jsonrpc': '2.0', 'id': 'bridge-list',
            'method': 'tools/call',
            'params': {'name': 'list_open_documents', 'arguments': {}},
        }, token, timeout=10)
        return (_tool_payload(resp) or {}).get('documents', []) or []
    except Exception:
        return []


def _documents_map(ports, token):
    """{port: [docs]} across instances, probed in parallel so one busy
    Revit (single-threaded server mid-tool) can't stall the whole fan-out."""
    with ThreadPoolExecutor(max_workers=max(len(ports), 1)) as pool:
        return dict(zip(ports, pool.map(
            lambda p: _documents_on(p, token), ports)))


def _match_tier(doc, query):
    """Mirror the server's match precedence. Lower tier = stronger match;
    None = no match. Tiers: 0 exact title, 1 exact path, 2 file name
    (with/without extension), 3 title substring."""
    q = (query or '').strip().lower()
    if not q:
        return None
    title = (doc.get('title') or '').lower()
    path = doc.get('path') or ''
    if path == '(unsaved)':
        path = ''
    if title == q:
        return 0
    try:
        if path and os.path.normcase(os.path.normpath(path)) == \
                os.path.normcase(os.path.normpath(query.strip())):
            return 1
    except Exception:
        pass
    base = os.path.basename(path).lower() if path else ''
    if base and (base == q or os.path.splitext(base)[0] == q):
        return 2
    if q in title:
        return 3
    return None


def _ensure_alive_target(current):
    """One-time (per bridge process) validation of the startup target port.

    Port 48884 (PORT_MIN) is not guaranteed to host a T3Lab server: a killed
    Revit can leave a hung zombie socket on it, and the pyRevit Routes server
    binds 0.0.0.0:48884 when enabled. Both accept connections without ever
    answering MCP — a request then dies with a timeout or an HTTP error, NOT
    connection-refused, so the conn-refused failover never fired and
    tools/list silently fell back to the (possibly empty) cache. Probe
    /health once up front and retarget to the first alive instance."""
    if current.get('validated'):
        return
    current['validated'] = True
    if _port_alive(current['port'], timeout=PROBE_TIMEOUT):
        return
    ports = _alive_ports()
    if ports:
        current['port'] = ports[0]


def _forward(request, current, token, timeout=CALL_TIMEOUT):
    """Forward to the current instance; fail over to another alive instance
    when the current one is gone — connection refused (that Revit was
    closed) or no longer answering /health (hung socket, foreign server on
    the port). Errors from a still-healthy instance (HTTP 500, a long tool
    call timing out) are NOT retried elsewhere — re-running the call on a
    different instance would target the wrong document."""
    _ensure_alive_target(current)
    try:
        return _post_mcp(current['port'], request, token, timeout)
    except Exception as e:
        if not _is_conn_refused(e) and _port_alive(current['port']):
            raise
        for port in _alive_ports():
            if port != current['port']:
                resp = _post_mcp(port, request, token, timeout)
                current['port'] = port
                return resp
        raise


def _handle_list(request, current, token):
    """Fan list_open_documents out to every alive instance. With a single
    instance the plain server response passes through untouched."""
    ports = _alive_ports()
    if len(ports) <= 1:
        return _forward(request, current, token)

    docs_map = _documents_map(ports, token)
    instances, total = [], 0
    for port in ports:
        docs = docs_map.get(port, [])
        for d in docs:
            d['instance_port'] = port
        total += len(docs)
        instances.append({
            'instance_port': port,
            'is_current_connection': port == current['port'],
            'documents': docs,
        })
    if total == 0:
        # Every instance reported nothing — most likely they run an older
        # build without list_open_documents. Pass the plain server response
        # through instead of a misleading empty merge.
        return _forward(request, current, token)
    return _tool_response(request.get('id'), {
        'instances': instances,
        'document_count': total,
        'note': ('Several Revit windows (instances) are open. '
                 'switch_active_document works on documents in ANY of them — '
                 'the bridge re-routes tool calls to the right window '
                 'automatically.'),
    })


def _handle_switch(request, current, token):
    """Route switch_active_document to whichever instance has the document.

    Order matters: other instances are checked BEFORE letting the current
    instance open the path from disk, otherwise a full path to a file that
    is already open in another Revit window would be re-opened (read-only)
    in the current one instead of switching windows.
    """
    rid = request.get('id')
    args = (request.get('params') or {}).get('arguments') or {}
    query = args.get('path_or_title') or ''

    ports = _alive_ports()
    multi = len(ports) > 1

    # Single window — plain forward (server opens from disk if needed).
    if not multi:
        return _forward(request, current, token)

    docs_map = _documents_map(ports, token)

    # Current instance wins when it has any match (the server applies the
    # exact precedence and ambiguity rules itself).
    if any(_match_tier(d, query) is not None
           for d in docs_map.get(current['port'], [])):
        return _forward(request, current, token)

    # Look for the document in the other Revit windows, best tier first.
    matches = []
    for port in ports:
        if port == current['port']:
            continue
        for d in docs_map.get(port, []):
            tier = _match_tier(d, query)
            if tier is not None:
                matches.append((tier, port, d))
    if matches:
        best = min(m[0] for m in matches)
        matches = [m for m in matches if m[0] == best]

    if len(matches) == 1:
        _, port, _doc = matches[0]
        resp = _post_mcp(port, request, token)
        payload = _tool_payload(resp)
        if payload and payload.get('success'):
            current['port'] = port
            payload['instance_port'] = port
            payload['switched_window'] = True
            payload['note'] = ('Document is open in another Revit window — the '
                               'bridge now routes ALL tool calls to that window '
                               '(port {}).').format(port)
            return _tool_response(rid, payload)
        return resp

    if len(matches) > 1:
        return _tool_response(rid, {
            'error': ('Ambiguous: "{}" matches documents in several Revit '
                      'windows.').format(query),
            'candidates': [{'instance_port': p,
                            'title': d.get('title'),
                            'path': d.get('path')} for _, p, d in matches],
        })

    # No match in any other window — let the current instance handle it
    # (open from disk, or report a proper error). Enrich that error with
    # what the other windows DO have open, so the AI can self-correct.
    resp = _forward(request, current, token)
    payload = _tool_payload(resp)
    if payload and payload.get('error'):
        other = {str(port): [d.get('title') for d in docs_map.get(port, [])]
                 for port in ports if port != current['port']}
        if any(other.values()):
            payload['documents_in_other_windows'] = other
            return _tool_response(rid, payload)
    return resp


def _server_info():
    return {'name': SERVER_NAME, 'version': BRIDGE_VERSION}


def _capabilities():
    return json.loads(json.dumps(SERVER_CAPABILITIES))


def _handle_initialize_local(request):
    """Answer initialize without touching Revit — attach always succeeds.
    Legacy handshake: never answers with a modern version."""
    params = request.get('params') or {}
    protocol = params.get('protocolVersion')
    if protocol not in LEGACY_PROTOCOLS:
        protocol = DEFAULT_PROTOCOL
    return {
        'jsonrpc': '2.0',
        'id': request.get('id'),
        'result': {
            'protocolVersion': protocol,
            'capabilities': _capabilities(),
            'serverInfo': _server_info(),
            'instructions': SERVER_INSTRUCTIONS,
        },
    }


def _handle_discover_local(request):
    """server/discover (MCP 2026-07-28) without touching Revit, so a modern
    client attaches with Revit closed exactly like a legacy one."""
    return {
        'jsonrpc': '2.0',
        'id': request.get('id'),
        'result': {
            'supportedVersions': list(SUPPORTED_PROTOCOLS),
            'capabilities': _capabilities(),
            'instructions': SERVER_INSTRUCTIONS,
        },
    }


def _declared_version(request):
    """Version a modern request declares in params._meta, else None."""
    params = request.get('params') if isinstance(request, dict) else None
    meta = params.get('_meta') if isinstance(params, dict) else None
    if not isinstance(meta, dict) or META_PROTOCOL not in meta:
        return None
    return str(meta.get(META_PROTOCOL) or '')


def _version_error(request_id, requested):
    return {'jsonrpc': '2.0', 'id': request_id,
            'error': {'code': -32022,
                      'message': 'Unsupported protocol version',
                      'data': {'supported': list(SUPPORTED_PROTOCOLS),
                               'requested': requested}}}


def _finish_modern(response, method):
    """resultType, serverInfo and cache hints on a modern result (copy of
    core.mcp_policy.finish_result)."""
    result = response.get('result') if isinstance(response, dict) else None
    if not isinstance(result, dict):
        return response
    result.setdefault('resultType', 'complete')
    meta = result.get('_meta')
    if not isinstance(meta, dict):
        meta = {}
        result['_meta'] = meta
    meta.setdefault(META_SERVER_INFO, _server_info())
    if method in CACHEABLE_METHODS:
        result.setdefault('ttlMs', LIST_TTL_MS)
        result.setdefault('cacheScope', CACHE_SCOPE)
    return response


def _handle_tools_list(request, current, token):
    """Live manifest from Revit when reachable (refreshing the cache);
    cached manifest when not — so a client that attached while Revit was
    closed still gets the full tool list. The bridge's own multi-instance
    tools are appended on top either way (only the server manifest is
    cached, so they never duplicate)."""
    try:
        resp = _forward(request, current, token, timeout=LIST_TIMEOUT)
        tools = (resp.get('result') or {}).get('tools')
        if isinstance(tools, list) and tools:
            _save_cached_tools(tools)
        else:
            tools = []
    except Exception:
        tools = _load_cached_tools()
    return {'jsonrpc': '2.0', 'id': request.get('id'),
            'result': {'tools': _annotate_large_results(tools + BRIDGE_TOOLS)}}


def _annotate_large_results(tools):
    """Raise the client's per-tool inline-result threshold for the read tools
    that legitimately return a lot of data, so their output reaches the model
    directly instead of being persisted to disk as a file reference. Applied
    after _save_cached_tools so the cache stays a verbatim server manifest."""
    for tool in tools:
        if tool.get('name') not in LARGE_RESULT_TOOLS:
            continue
        meta = tool.get('_meta')
        if not isinstance(meta, dict):
            meta = {}
            tool['_meta'] = meta
        meta.setdefault('anthropic/maxResultSizeChars', LARGE_RESULT_CHARS)
    return tools


def _handle_instances(request, current, token):
    """list_revit_instances — probe the port range and report every running
    Revit instance with its open documents. Served entirely by the bridge."""
    rid = request.get('id')
    ports = _alive_ports()
    if not ports:
        return _revit_down_response(rid)

    docs_map = _documents_map(ports, token)
    instances = []
    for port in ports:
        docs = docs_map.get(port, [])
        instances.append({
            'port': port,
            'is_current': port == current['port'],
            'active_document': next(
                (d.get('title') for d in docs if d.get('is_active')), None),
            'documents': [{'title': d.get('title'), 'path': d.get('path')}
                          for d in docs],
        })

    payload = {
        'instances': instances,
        'count': len(instances),
        'current_port': current['port'],
        'note': ('Tool calls go to the instance marked is_current. Retarget with '
                 'switch_revit_instance (by port) or switch_active_document (by '
                 'document — re-routes across instances automatically).'),
    }
    if current['port'] not in ports:
        payload['warning'] = ('The instance the bridge last targeted (port {}) is gone — '
                              'the next tool call fails over to another alive instance '
                              'automatically.').format(current['port'])
    return _tool_response(rid, payload)


def _handle_switch_instance(request, current, token):
    """switch_revit_instance — explicit retarget by port, or by document
    (delegated to the cross-instance switch_active_document handling so the
    target document's window is activated as well)."""
    rid = request.get('id')
    args = (request.get('params') or {}).get('arguments') or {}
    port = args.get('port')
    document = (args.get('document') or '').strip()

    if port is None and not document:
        return _tool_response(rid, {
            'error': 'Pass port (from list_revit_instances) or document.',
        })

    if port is not None:
        try:
            port = int(port)
        except (TypeError, ValueError):
            return _tool_response(rid, {'error': 'port must be an integer.'})
        # Probe the requested port directly with a generous timeout — a busy
        # instance can miss the fast parallel range scan yet still be alive,
        # and refusing a switch to a live instance is worse than waiting.
        if not _port_alive(port, timeout=TARGET_PROBE_TIMEOUT):
            ports = _alive_ports()
            if not ports:
                return _revit_down_response(rid)
            return _tool_response(rid, {
                'error': 'No running Revit instance on port {}.'.format(port),
                'alive_ports': ports,
                'hint': 'Call list_revit_instances to see what is running.',
            })
        already = port == current['port']
        current['port'] = port
        docs = _documents_on(port, token)
        return _tool_response(rid, {
            'success': True,
            'port': port,
            'already_current': already,
            'active_document': next(
                (d.get('title') for d in docs if d.get('is_active')), None),
            'documents': [d.get('title') for d in docs],
            'note': ('All tool calls now go to the Revit instance on port {}. '
                     'They target its ACTIVE document — use switch_active_document '
                     'to activate a different one.').format(port),
        })

    # By document — same semantics as switch_active_document, which already
    # finds the document across instances, switches the bridge's port and
    # activates the window.
    synthetic = {
        'jsonrpc': '2.0',
        'id': rid,
        'method': 'tools/call',
        'params': {'name': 'switch_active_document',
                   'arguments': {'path_or_title': document}},
    }
    return _handle_switch(synthetic, current, token)


def _revit_down_response(request_id):
    return _tool_response(request_id, {
        'error': 'Revit is not reachable — no T3Lab MCP server found on '
                 'ports {}-{}.'.format(PORT_MIN, PORT_MAX),
        'hint': ('Start Revit and make sure the T3Lab server is running '
                 '(it auto-starts with the T3Lab extension; otherwise use '
                 'the MCP Control button on the T3Lab ribbon). Then retry '
                 'this tool — no reconnect/restart of the client is needed.'),
    })


# List methods served live from Revit (the wiki pages as resources, the
# workflow prompts) and cached like tools/list, so they still answer with
# Revit closed: with the last live list, else an empty one. An empty list --
# never an error -- also answers an older Revit build without them.
LIST_METHODS = {
    'resources/list':           'resources',
    'resources/templates/list': 'resourceTemplates',
    'prompts/list':             'prompts',
}
# Single-item reads that need the live server.
READ_METHODS = frozenset(['resources/read', 'prompts/get'])


def _handle_cached_list(request, current, token):
    method = request.get('method')
    key = LIST_METHODS[method]
    items = None
    try:
        resp = _forward(request, current, token, timeout=LIST_TIMEOUT)
        result = resp.get('result') if isinstance(resp, dict) else None
        if isinstance(result, dict) and isinstance(result.get(key), list):
            items = result[key]
            _save_cached_list(key, items)
    except Exception:
        pass
    if items is None:
        items = _load_cached_list(key)
    return {'jsonrpc': '2.0', 'id': request.get('id'), 'result': {key: items}}


def _handle_read(request, current, token):
    try:
        return _forward(request, current, token, timeout=LIST_TIMEOUT)
    except Exception as e:
        if not _is_conn_refused(e) and _alive_ports():
            raise
        return _rpc_error(request.get('id'),
                          'Revit is not reachable, so {} cannot be answered. '
                          'Start Revit with the T3Lab extension and retry.'.format(
                              request.get('method')))


def _handle(request, current, token):
    """One request -> response. Modern requests (version in params._meta)
    are checked here and their results get the modern fields."""
    declared = _declared_version(request)
    if declared is not None and declared not in MODERN_PROTOCOLS:
        return _version_error(request.get('id'), declared)
    response = _dispatch(request, current, token)
    if declared is not None:
        response = _finish_modern(response, request.get('method'))
    return response


def _dispatch(request, current, token):
    method = request.get('method')
    if method == 'initialize':
        return _handle_initialize_local(request)
    if method == 'server/discover':
        return _handle_discover_local(request)
    if method == 'ping':
        return {'jsonrpc': '2.0', 'id': request.get('id'), 'result': {}}
    if method in LIST_METHODS:
        return _handle_cached_list(request, current, token)
    if method in READ_METHODS:
        return _handle_read(request, current, token)
    if method == 'tools/list':
        return _handle_tools_list(request, current, token)
    if method == 'tools/call':
        name = (request.get('params') or {}).get('name')
        try:
            if name == 'list_revit_instances':
                return _handle_instances(request, current, token)
            if name == 'switch_revit_instance':
                return _handle_switch_instance(request, current, token)
            if name == 'list_open_documents':
                return _handle_list(request, current, token)
            if name == 'switch_active_document':
                return _handle_switch(request, current, token)
            return _forward(request, current, token)
        except Exception as e:
            if _is_conn_refused(e):
                return _revit_down_response(request.get('id'))
            raise

    # Anything else (subscriptions/listen, logging ...): the in-Revit server
    # does not implement it, so forwarding just trades a clean protocol answer
    # for a connection error whenever Revit happens to be closed. -32601 is what
    # the spec asks for and every client already knows how to ignore it.
    return _rpc_error(request.get('id'),
                      'Method not supported by the T3Lab Revit bridge: {}'.format(method),
                      code=-32601)


def main():
    port = PORT_MIN
    # Read port from arguments if provided
    for arg in sys.argv[1:]:
        try:
            port = int(arg)
            break
        except ValueError:
            pass

    current = {'port': port}
    token = _read_token()

    while True:
        line = sys.stdin.readline()
        if not line:
            break
        try:
            request = json.loads(line)
        except Exception:
            continue

        # Notifications (no 'id') are handled locally — the in-Revit server
        # keeps no MCP session state, so forwarding them only adds a failed
        # connect per notification whenever Revit is closed.
        if 'id' not in request:
            continue

        try:
            response = _handle(request, current, token)
        except Exception as e:
            response = _rpc_error(
                request.get('id'),
                f"T3Lab Revit server connection failed: {str(e)}")

        try:
            response = _cap_result(response)
        except Exception:
            pass          # a guard failure must never eat the response itself

        sys.stdout.write(json.dumps(response) + '\n')
        sys.stdout.flush()


if __name__ == '__main__':
    main()
