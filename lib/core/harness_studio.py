# -*- coding: utf-8 -*-
"""Harness Studio: the task board as a node canvas in the browser
(docs/agents/harness-studio.md).

Runs in the bridge process (`bridge.py harness studio`), never inside Revit,
and listens on 127.0.0.1 only. Reads come from the store; every change goes
through the gateway as the person at the machine (owner=true, transport cli),
so roles, shared memory and the approval re-check stay in one place.
Board text is written by agents: the page shows it as text, never as HTML.
Test: dev/test_harness_studio.py.
"""
import collections
import hmac
import json
import os
import re
import threading
import time
import uuid

DEFAULT_PORT = 48881            # 48880 = MCP router, 48884-48892 = Revit servers
PORT_TRIES = 10
PAGE_FILE = 'harness_studio.html'
RUNNING_FILE = 'studio.json'    # {port, token} of the Studio that is up, so a second start reuses it
TOKEN_HEADER = 'X-T3Lab-Studio'
LOCAL_HOSTS = ('127.0.0.1', 'localhost', '[::1]', '::1')
MAX_BODY = 512 * 1024
OWNER_ACTIONS = ('open', 'plan', 'remove', 'review', 'retry', 'note', 'close')
AGENT_ACTIONS = ('agent_add', 'agent_remove')
APPROVAL_ACTIONS = ('approve', 'reject')

FLOW_FORMAT = 'harness-flow/1'
FLOWS_DIR = 'flows'
TASK_FIELDS = ('key', 'title', 'detail', 'depends_on', 'assignee', 'done_when', 'scope')
_FLOW_RE = re.compile(r'^[a-z0-9][a-z0-9._-]{0,39}$')

RUN_IDLE_WAIT_S = 5
RUN_MAX_IDLE = 360              # 30 min with nothing moving (review / approval pending)
QUEUE_TTL_S = 10


class StudioError(Exception):
    def __init__(self, message, code='bad_arguments', **extra):
        Exception.__init__(self, message)
        self.payload = dict({'error': message, 'error_code': code}, **extra)


# ─── Flows: a plan saved for reuse ───────────────────────────────────────────

def _flow_path(root, name):
    name = u'{}'.format(name or '').strip().lower()
    if not _FLOW_RE.match(name):
        raise StudioError('A flow name is 1-40 characters: a-z, 0-9, ".", "_" or "-".')
    return name, os.path.join(root, FLOWS_DIR, name + '.json')


def flow_tasks(tasks):
    """Plan fields only: a flow never carries status, owner or results."""
    out = []
    for task in tasks or []:
        if isinstance(task, dict):
            item = dict((k, task[k]) for k in TASK_FIELDS if task.get(k) not in (None, '', []))
            item['key'] = item.get('key') or task.get('task')
            out.append(item)
    return out


def list_flows(root):
    out = []
    try:
        names = sorted(os.listdir(os.path.join(root, FLOWS_DIR)))
    except OSError:
        return out
    for name in names:
        if name.endswith('.json'):
            try:
                flow = load_flow(root, name[:-5])
            except StudioError:
                continue
            out.append({'name': flow['name'], 'goal': flow.get('goal') or '',
                        'tasks': len(flow.get('tasks') or []), 'saved': flow.get('saved')})
    return out


def load_flow(root, name):
    name, path = _flow_path(root, name)
    try:
        with open(path, encoding='utf-8') as f:
            flow = json.load(f)
    except (OSError, ValueError):
        raise StudioError('No flow "{}".'.format(name), 'flow_not_found')
    if not isinstance(flow, dict) or flow.get('format') != FLOW_FORMAT:
        raise StudioError('"{}" is not a harness flow.'.format(name), 'flow_not_found')
    return dict(flow, name=name)


def save_flow(root, name, goal, tasks, now=time.time):
    name, path = _flow_path(root, name)
    tasks = flow_tasks(tasks)
    if not tasks:
        raise StudioError('There is no task to save.')
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = '{}.{}.tmp'.format(path, uuid.uuid4().hex[:8])
    with open(tmp, 'w', encoding='utf-8') as f:
        json.dump({'format': FLOW_FORMAT, 'name': name, 'goal': u'{}'.format(goal or ''),
                   'tasks': tasks, 'saved': now()}, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)
    return {'saved': name, 'tasks': len(tasks)}


def delete_flow(root, name):
    name, path = _flow_path(root, name)
    try:
        os.remove(path)
    except OSError:
        raise StudioError('No flow "{}".'.format(name), 'flow_not_found')
    return {'deleted': name}


# ─── Runner: the launched agents, round after round ──────────────────────────

class Runs(object):
    """`harness run --loop` on a thread, so the page can start and stop it."""

    def __init__(self, call, wait_s=RUN_IDLE_WAIT_S, max_idle=RUN_MAX_IDLE):
        self._call = call           # call(action, args) -> payload
        self._wait_s, self._max_idle = wait_s, max_idle
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread = None
        self._state = {'running': False, 'session': None, 'rounds': 0, 'ended': ''}
        self._log = collections.deque(maxlen=40)

    def view(self):
        with self._lock:
            return dict(self._state, stopping=self._state['running'] and self._stop.is_set(),
                        log=list(self._log))

    def start(self, session, parallel=None):
        with self._lock:
            if self._state['running']:
                raise StudioError('The agents are already running.', 'already_running')
            self._stop.clear()
            self._log.clear()
            self._state = {'running': True, 'session': session, 'rounds': 0, 'ended': ''}
            self._thread = threading.Thread(target=self._loop, args=(session, parallel),
                                            name='harness-studio-run', daemon=True)
            self._thread.start()
        return self.view()

    def stop(self):
        self._stop.set()            # the round in flight finishes: agents always report
        return self.view()

    def join(self, timeout=None):
        if self._thread is not None:
            self._thread.join(timeout)

    def _end(self, why):
        with self._lock:
            self._state.update(running=False, ended=why)

    def _loop(self, session, parallel):
        idle = 0
        while not self._stop.is_set():
            args = {'session': session} if session else {}
            if parallel:
                args['parallel'] = parallel
            try:
                payload = self._call('run', args)
            except Exception as ex:
                return self._end('{}: {}'.format(type(ex).__name__, ex))
            ran = payload.get('ran') if isinstance(payload, dict) else None
            if not isinstance(ran, list):
                return self._end((payload or {}).get('error') if isinstance(payload, dict)
                                 else 'The runner did not answer.')
            if not ran:
                return self._end('No launched agent is registered (Agents > Add agent).')
            with self._lock:
                self._state['rounds'] += 1
                for item in ran:
                    if item.get('task'):
                        self._log.append(dict(item, at=time.time()))
            if all(r.get('all_done') for r in ran):
                return self._end('Every task is done.')
            idle = 0 if any(r.get('task') for r in ran) else idle + 1
            if idle > self._max_idle:
                return self._end('Nothing moved for {} minutes.'.format(
                    int(idle * self._wait_s / 60)))
            if idle and self._stop.wait(self._wait_s):
                break
        self._end('Stopped.')


# ─── Studio: state for the page, actions from the page ───────────────────────

class Studio(object):
    instances_tool = 'list_revit_instances'          # read-only bridge tool, no Revit call

    def __init__(self, store, agents, gateway, harness_tool, approval_tool,
                 approvals=None, root=None, token=None):
        self.store, self.agents, self.gateway = store, agents, gateway
        self.harness_tool, self.approval_tool = harness_tool, approval_tool
        self._approvals = approvals or (lambda: [])
        self.root = root or store.harness_dir()
        self.token = token or uuid.uuid4().hex + uuid.uuid4().hex
        self.runs = Runs(lambda action, args: self.call(action, args, owner=False))
        self.last_request = time.time()
        self._queue = {'at': 0.0, 'value': None, 'busy': False}
        self._revit = {'at': 0.0, 'value': None, 'busy': False}

    def call(self, action, args=None, owner=True):
        args = dict(args or {}, action=action)
        if owner:
            args['owner'] = True
        return self.gateway(self.harness_tool, args)

    def _revit_queue(self, session):
        """Who holds Revit (gateway status). Refreshed off the request thread:
        a closed Revit makes the health probe slow on Windows."""
        queue = self._queue
        if not queue['busy'] and time.time() - queue['at'] > QUEUE_TTL_S:
            queue['busy'] = True

            def refresh():
                try:
                    out = self.call('status', {'session': session}, owner=False)
                    queue['value'] = out.get('revit_queue') if isinstance(out, dict) else None
                except Exception:
                    queue['value'] = None
                queue.update(at=time.time(), busy=False)
            threading.Thread(target=refresh, name='harness-studio-queue', daemon=True).start()
        return queue['value']

    def _revit_probe(self):
        """{ok, documents, active} for the Start screen and the ready check, from
        list_revit_instances; None until the first answer. Off the request thread."""
        probe = self._revit
        if not probe['busy'] and time.time() - probe['at'] > QUEUE_TTL_S:
            probe['busy'] = True

            def refresh():
                try:
                    out = self.gateway(self.instances_tool, {})
                    found = out.get('instances') if isinstance(out, dict) else None
                    if isinstance(found, list) and found:
                        docs, active = [], None
                        for inst in found:
                            for doc in inst.get('documents') or []:
                                if doc.get('title') and doc['title'] not in docs:
                                    docs.append(doc['title'])
                            if inst.get('is_current') or active is None:
                                active = inst.get('active_document') or active
                        probe['value'] = {'ok': True, 'documents': docs, 'active': active}
                    else:
                        probe['value'] = {'ok': False, 'documents': [], 'active': None}
                except Exception:
                    probe['value'] = {'ok': False, 'documents': [], 'active': None}
                probe.update(at=time.time(), busy=False)
            threading.Thread(target=refresh, name='harness-studio-revit', daemon=True).start()
        return probe['value']

    def state(self, session=None):
        board = self.store.Harness(self.root)
        sessions = board.sessions()
        ids = [s['session'] for s in sessions]
        opened = [s['session'] for s in sessions if s['status'] == 'open']
        sid = session if session in ids else (opened[0] if opened else None)
        out = {'setup': self.store.load_config(self.root), 'sessions': sessions, 'session': sid,
               'board': None, 'stats': None, 'revit_queue': None, 'revit': self._revit_probe(),
               'agents': self.agents.load(self.root) if self.agents else {},
               'platforms': list(getattr(self.agents, 'PLATFORMS', ())),
               'approvals': self._approvals(), 'run': self.runs.view(),
               'flows': list_flows(self.root), 'now': time.time()}
        if sid:
            out['board'] = board.status(sid, full=True)
            out['stats'] = board.stats(sid)
            if out['board']['status'] == 'open':
                out['revit_queue'] = self._revit_queue(sid)
        return out

    def act(self, action, args=None):
        args = args if isinstance(args, dict) else {}
        if action in OWNER_ACTIONS:
            return self.call(action, args)
        if action in AGENT_ACTIONS:
            return self.call(action, args, owner=False)
        if action in APPROVAL_ACTIONS:
            return self.gateway(self.approval_tool,
                                {'action': action, 'approval_id': args.get('approval_id')})
        raise StudioError('Unknown action "{}".'.format(action))

    def run(self, op, args=None):
        args = args if isinstance(args, dict) else {}
        if op == 'start':
            return self.runs.start(args.get('session'), args.get('parallel'))
        if op == 'stop':
            return self.runs.stop()
        raise StudioError('run: start or stop.')

    def flow(self, op, args=None):
        args = args if isinstance(args, dict) else {}
        if op == 'save':
            board = self.store.Harness(self.root).status(args.get('session'), full=True)
            return save_flow(self.root, args.get('name'), args.get('goal') or board.get('goal'),
                             board.get('tasks'))
        if op == 'delete':
            return delete_flow(self.root, args.get('name'))
        if op == 'get':
            return load_flow(self.root, args.get('name'))
        if op == 'apply':
            flow = load_flow(self.root, args.get('name'))
            session = args.get('session')
            if not session:
                opened = self.call('open', {
                    'goal': args.get('goal') or flow.get('goal') or flow['name'],
                    'revit_document': args.get('revit_document') or '',
                    'review': args.get('review', True),
                    'exclusive': args.get('exclusive'), 'write_mode': args.get('write_mode')})
                if not isinstance(opened, dict) or opened.get('error'):
                    return opened
                session = opened['session']
            out = self.call('plan', {'session': session, 'tasks': flow['tasks']})
            return dict(out, session=session) if isinstance(out, dict) else out
        raise StudioError('flow: save, apply, get or delete.')


# ─── HTTP ────────────────────────────────────────────────────────────────────

def _page():
    with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), PAGE_FILE), 'rb') as f:
        return f.read()


def make_handler(studio):
    from http.server import BaseHTTPRequestHandler
    from urllib.parse import parse_qs, urlparse

    class StudioHandler(BaseHTTPRequestHandler):
        protocol_version = 'HTTP/1.1'
        server_version = 'T3LabHarnessStudio'

        def log_message(self, fmt, *args):
            pass

        def _send(self, code, body, kind='application/json'):
            if not isinstance(body, bytes):
                body = json.dumps(body, ensure_ascii=False).encode('utf-8')
            self.send_response(code)
            self.send_header('Content-Type', kind + '; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.send_header('Content-Security-Policy',
                             "default-src 'self'; style-src 'unsafe-inline'; "
                             "script-src 'unsafe-inline'; frame-ancestors 'none'")
            self.end_headers()
            self.wfile.write(body)

        def _refusal(self):
            """Why this request is not from the owner's own browser, or ''."""
            host = (self.headers.get('Host') or '').rsplit(':', 1)[0].lower()
            if host not in LOCAL_HOSTS:         # DNS rebinding
                return 'Host not allowed.'
            origin = self.headers.get('Origin')
            if origin and (urlparse(origin).hostname or '') not in ('127.0.0.1', 'localhost',
                                                                   '::1'):
                return 'Origin not allowed.'
            return ''

        def _authorized(self):
            given = self.headers.get(TOKEN_HEADER) or ''
            return hmac.compare_digest(given.encode('utf-8'), studio.token.encode('utf-8'))

        def do_GET(self):
            url = urlparse(self.path)
            studio.last_request = time.time()
            refusal = self._refusal()
            if refusal:
                return self._send(403, {'error': refusal})
            if url.path == '/':
                try:
                    return self._send(200, _page(), 'text/html')
                except OSError:
                    return self._send(500, {'error': PAGE_FILE + ' is not next to the bridge.'})
            if url.path != '/api/state':
                return self._send(404, {'error': 'Not found'})
            if not self._authorized():
                return self._send(401, {'error': 'Open the link printed by '
                                                 '`bridge.py harness studio`.'})
            try:
                session = (parse_qs(url.query).get('session') or [''])[0]
                self._send(200, studio.state(session or None))
            except Exception as ex:
                self._send(200, getattr(ex, 'payload', None)
                           or {'error': '{}: {}'.format(type(ex).__name__, ex)})

        def do_POST(self):
            path = urlparse(self.path).path
            studio.last_request = time.time()
            refusal = self._refusal()
            if refusal:
                return self._send(403, {'error': refusal})
            if not self._authorized():
                return self._send(401, {'error': 'Missing or wrong Studio token.'})
            try:
                length = int(self.headers.get('Content-Length') or 0)
                if length > MAX_BODY:
                    return self._send(413, {'error': 'Request too large.'})
                body = json.loads(self.rfile.read(length).decode('utf-8') or '{}')
                if not isinstance(body, dict):
                    raise ValueError('not an object')
            except (ValueError, UnicodeDecodeError):
                return self._send(400, {'error': 'The request body is not a JSON object.'})
            op, args = body.get('action'), body.get('args')
            try:
                if path == '/api/act':
                    out = studio.act(op, args)
                elif path == '/api/run':
                    out = studio.run(op, args)
                elif path == '/api/flow':
                    out = studio.flow(op, args)
                else:
                    return self._send(404, {'error': 'Not found'})
            except Exception as ex:             # store and agent errors carry a payload
                out = getattr(ex, 'payload', None) or {
                    'error': '{}: {}'.format(type(ex).__name__, ex)}
            self._send(200, out if out is not None else {})

    return StudioHandler


def make_server(studio, port=None, host='127.0.0.1'):
    """The server on `port`, or on the first free of DEFAULT_PORT.. when None."""
    from http.server import ThreadingHTTPServer
    ports = [port] if port is not None else range(DEFAULT_PORT, DEFAULT_PORT + PORT_TRIES)
    error = None
    for candidate in ports:
        try:
            httpd = ThreadingHTTPServer((host, candidate), make_handler(studio))
        except OSError as ex:
            error = ex
            continue
        httpd.daemon_threads = True
        return httpd
    raise error


def _running(root, urlopen=None):
    """URL of a Studio already serving this board, or ''."""
    try:
        with open(os.path.join(root, RUNNING_FILE), encoding='utf-8') as f:
            info = json.load(f)
        port, token = int(info['port']), u'{}'.format(info['token'])
        from urllib.request import ProxyHandler, Request, build_opener
        urlopen = urlopen or build_opener(ProxyHandler({})).open     # never through a proxy
        request = Request('http://127.0.0.1:{}/api/state'.format(port),
                          headers={TOKEN_HEADER: token})
        with urlopen(request, timeout=2) as response:
            if response.status == 200:
                return 'http://127.0.0.1:{}/?k={}'.format(port, token)
    except Exception:
        pass
    return ''


def _idle_watch(studio, httpd, minutes, sleep=time.sleep):
    """Stop the server once no page has polled for `minutes` and no agent runs
    (a Studio started from Revit has no terminal to Ctrl+C)."""
    while True:
        sleep(15)
        if studio.runs.view()['running']:
            studio.last_request = time.time()
        elif time.time() - studio.last_request > minutes * 60:
            httpd.shutdown()
            return


def serve(store, agents, gateway, harness_tool, approval_tool, port=None, open_browser=True,
          approvals=None, say=None, idle_exit_min=None):
    """Blocks until Ctrl+C (or idle_exit_min without a page). Exit code for the CLI."""
    say = say or (lambda text: None)

    def browse(url):
        if open_browser:
            try:
                import webbrowser
                webbrowser.open(url)
            except Exception:
                pass

    root = store.harness_dir()
    url = _running(root)
    if url:
        say('Harness Studio is already running: {}'.format(url))
        browse(url)
        return 0
    studio = Studio(store, agents, gateway, harness_tool, approval_tool, approvals)
    try:
        httpd = make_server(studio, port)
    except OSError as ex:
        say('Harness Studio could not listen: {}. Pass --studio-port=<free port>.'.format(ex))
        return 1
    url = 'http://127.0.0.1:{}/?k={}'.format(httpd.server_address[1], studio.token)
    running = os.path.join(root, RUNNING_FILE)
    try:
        os.makedirs(root, exist_ok=True)
        with open(running, 'w', encoding='utf-8') as f:
            json.dump({'port': httpd.server_address[1], 'token': studio.token,
                       'pid': os.getpid()}, f)
    except OSError:
        pass                    # a second start then opens its own Studio
    say('Harness Studio: {}'.format(url))
    say('Only this machine can open it. Ctrl+C stops it.')
    if idle_exit_min:
        threading.Thread(target=_idle_watch, args=(studio, httpd, idle_exit_min),
                         name='harness-studio-idle', daemon=True).start()
    browse(url)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        studio.runs.stop()
        httpd.server_close()
        try:
            os.remove(running)
        except OSError:
            pass
    return 0
