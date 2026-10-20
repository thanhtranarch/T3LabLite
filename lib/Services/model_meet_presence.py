# -*- coding: utf-8 -*-
"""
Model Meet — presence: who is walking the same model right now, and chat.

Everyone who opens the same model joins the same ROOM. The room code is a hash
of the model's identity (cloud model GUID, else central path, else project
number + title): the identity itself never leaves the PC, and two people can
always meet by typing the same code by hand.

What is shared with the room — nothing else:
    character name + look, Revit user name (to match the Activity page),
    position (model feet) + view direction, level name, walking/standing,
    and the chat lines people type.

Two transports, same interface:

  FolderTransport    a folder everyone can reach: the central model's server
                     folder (office network), or a OneDrive / Dropbox / ACC
                     Desktop Connector folder (worldwide, as fast as that
                     service syncs). Files under <folder>/T3LabMeet/<room>/.
  SupabaseTransport  the Internet, through a free Supabase project (two tables,
                     docs/model-meet.md has the SQL). Real time: ~1 s.

PresenceClient runs the network on a background thread. Model Meet is a modal
window, so CPython threads get the GIL while it is open (CLAUDE.md rule 8); the
UI thread only swaps snapshots under a lock and never waits on the network.

Pure Python (+ System.Net when inside Revit). Tested by dev/test_model_meet.py.

Author: T3Lab
"""
from __future__ import division

import base64
import hashlib
import io
import json
import os
import random
import re
import threading
import time
import uuid
import zlib
from datetime import datetime

from Snippets import model_meet_walk as mw

PROTOCOL = 1
TTL_SECONDS = 20.0          # no update for this long = left the room
INTERVAL = 1.0              # seconds between two network rounds
HTTP_TIMEOUT = 6.0
NAME_MAX = 60
TEXT_MAX = 400
HISTORY = 50                # chat lines a newcomer receives
FOLDER_CHAT_KEEP = 200      # chat files read from a shared folder
STALE_FILE_SECONDS = 24 * 3600
FIRST_SIGHT_MAX = 120.0     # s; how old a presence file may look the first time it is seen
REPLACE_TRIES = 3           # a shared file another PC is reading cannot be replaced for a moment

FOLDER = "folder"
SUPABASE = "supabase"

PRESENCE_TABLE = "t3_meet_presence"
MESSAGE_TABLE = "t3_meet_messages"


class TransportError(Exception):
    """A failure the user can act on; the message is plain English."""


# ── identity ─────────────────────────────────────────────────────────────────
def room_seed(cloud_guid=u"", central=u"", number=u"", title=u""):
    """What makes two copies of a model 'the same project'."""
    if cloud_guid:
        return u"cloud:" + cloud_guid.strip().lower()
    if central:
        return u"central:" + central.strip().replace(u"\\", u"/").lower()
    title = re.sub(r"\.rvt$", u"", (title or u"").strip(), flags=re.IGNORECASE)   # Explorer may hide it
    return u"project:{}|{}".format((number or u"").strip().lower(), title.lower())


def room_code(seed):
    """'K7QM-2XDA' — 8 base32 characters of sha1(seed)."""
    digest = hashlib.sha1((seed or u"").encode("utf-8")).digest()
    code = base64.b32encode(digest).decode("ascii")[:8]
    return code[:4] + u"-" + code[4:]


def normalize_room(text):
    """User-typed room code -> canonical form, or '' when unusable.

    Letters and digits only (4–32), upper case; 8 characters get the dash back.
    """
    raw = re.sub(r"[^A-Za-z0-9]", u"", text or u"").upper()
    if not 4 <= len(raw) <= 32:
        return u""
    return raw[:4] + u"-" + raw[4:] if len(raw) == 8 else raw


def user_key(name):
    """File- and URL-safe key of a user name, stable on every PC."""
    clean = (name or u"").strip()
    slug = re.sub(r"[^a-z0-9]+", u"-", clean.lower()).strip(u"-")[:24] or u"user"
    return u"{}-{:08x}".format(slug, zlib.crc32(clean.lower().encode("utf-8")) & 0xffffffff)


def _text(value, limit):
    if value is None:
        return u""
    if not isinstance(value, str):
        value = u"{}".format(value)
    value = u" ".join(value.split())
    return value[:limit]


def _number(value, default=0.0, limit=1e7):
    try:
        v = float(value)
    except (TypeError, ValueError):
        return default
    if v != v or abs(v) > limit:        # NaN or absurd
        return default
    return v


# ── wire format ──────────────────────────────────────────────────────────────
class Peer(object):
    """Someone in the room, as last reported."""

    __slots__ = ("key", "name", "user", "look", "x", "y", "z", "yaw", "pitch", "level", "status", "ts")

    def __init__(self, key, name, x, y, z, yaw, pitch, level, status, ts, user=u"", look=None):
        self.key = key
        self.name = name
        self.user = user or name
        self.look = look or mw.clean_look({}, name)
        self.x, self.y, self.z = x, y, z
        self.yaw, self.pitch = yaw, pitch
        self.level = level
        self.status = status
        self.ts = ts

    @property
    def position(self):
        return (self.x, self.y, self.z)


class Message(object):
    __slots__ = ("id", "key", "name", "text", "ts")

    def __init__(self, id_, key, name, text, ts):
        self.id = id_
        self.key = key
        self.name = name
        self.text = text
        self.ts = ts

    @property
    def time_label(self):
        try:
            return datetime.fromtimestamp(self.ts).strftime("%H:%M")
        except Exception:
            return u""


def clean_name(name):
    """The name as the room sees it — and the one user_key must be made from."""
    return _text(name, NAME_MAX) or u"Someone"


def make_state(name, position, yaw, pitch, level=u"", status=u"walking", now=None,
               user=u"", look=None):
    """Presence payload. `position` in MODEL feet (internal coordinates), the eye."""
    name = clean_name(name)
    return {
        "v": PROTOCOL,
        "key": user_key(name),
        "name": name,
        "user": _text(user, NAME_MAX) or name,
        "look": mw.clean_look(look, name),
        "x": round(float(position[0]), 3),
        "y": round(float(position[1]), 3),
        "z": round(float(position[2]), 3),
        "yaw": round(float(yaw), 1),
        "pitch": round(float(pitch), 1),
        "level": _text(level, NAME_MAX),
        "status": _text(status, 20),
        "ts": float(now if now is not None else time.time()),
    }


def read_state(data):
    """Peer from a payload someone else wrote, or None. Never trusts the input."""
    if not isinstance(data, dict):
        return None
    name = _text(data.get("name"), NAME_MAX)
    if not name:
        return None
    key = user_key(name)
    return Peer(key, name,
                _number(data.get("x")), _number(data.get("y")), _number(data.get("z")),
                _number(data.get("yaw"), limit=1e4), _number(data.get("pitch"), limit=90.0),
                _text(data.get("level"), NAME_MAX), _text(data.get("status"), 20),
                _number(data.get("ts"), limit=1e11),
                _text(data.get("user"), NAME_MAX), mw.clean_look(data.get("look"), name))


def make_message(name, text, now=None):
    name = _text(name, NAME_MAX) or u"Someone"
    return {"v": PROTOCOL, "id": uuid.uuid4().hex, "key": user_key(name), "name": name,
            "text": _text(text, TEXT_MAX),
            "ts": float(now if now is not None else time.time())}


def read_message(data):
    if not isinstance(data, dict):
        return None
    name = _text(data.get("name"), NAME_MAX)
    text = _text(data.get("text"), TEXT_MAX)
    if not name or not text:
        return None
    mid = re.sub(r"[^0-9a-f]", u"", _text(data.get("id"), 40).lower()) or uuid.uuid4().hex
    return Message(mid, user_key(name), name, text, _number(data.get("ts"), limit=1e11))


def parse_iso(stamp):
    """Epoch seconds of an ISO-8601 timestamp ('2026-10-10T08:15:30.12+00:00'), or None."""
    if not stamp:
        return None
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})(\.\d+)?"
                 r"(Z|[+-]\d{2}:?\d{2})?$", stamp.strip())
    if not m:
        return None
    y, mo, d, h, mi, s = (int(m.group(i)) for i in range(1, 7))
    frac = float(m.group(7)) if m.group(7) else 0.0
    import calendar
    epoch = calendar.timegm((y, mo, d, h, mi, s, 0, 0, 0)) + frac
    tz = m.group(8)
    if tz and tz != "Z":
        sign = 1 if tz[0] == "+" else -1
        digits = tz[1:].replace(":", "")
        epoch -= sign * (int(digits[:2]) * 3600 + int(digits[2:4]) * 60)
    return epoch


# ── transports ───────────────────────────────────────────────────────────────
def _write_json_atomic(path, data, tries=REPLACE_TRIES):
    """Write through a temp file + rename. On a Windows share the rename fails for
    a moment while another PC (or the OneDrive client) has the file open: retry,
    and never leave the temp file behind."""
    folder = os.path.dirname(path)
    if not os.path.isdir(folder):
        os.makedirs(folder)
    tmp = u"{}.{}.tmp".format(path, uuid.uuid4().hex[:6])
    try:
        with io.open(tmp, "w", encoding="utf-8") as fh:
            fh.write(json.dumps(data, ensure_ascii=True))
        for attempt in range(tries):
            try:
                os.replace(tmp, path)
                return
            except PermissionError:
                if attempt == tries - 1:
                    raise
                time.sleep(0.05 * (attempt + 1))
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:
                pass


def _read_json(path):
    try:
        with io.open(path, "r", encoding="utf-8") as fh:
            return json.load(fh)
    except Exception:
        return None


class FolderTransport(object):
    """Presence and chat as small JSON files in a shared folder."""

    kind = FOLDER

    def __init__(self, root, room, clock=time.time):
        if not root:
            raise TransportError(u"Choose a shared folder first.")
        self.root = root
        self.room = room
        self.clock = clock
        self.room_dir = os.path.join(root, u"T3LabMeet", room)
        self.presence_dir = os.path.join(self.room_dir, u"presence")
        self.chat_dir = os.path.join(self.room_dir, u"chat")
        self._checked = False
        self._observed = {}         # peer key -> (ts, local time that ts was first seen)

    def describe(self):
        return u"shared folder {}".format(self.root)

    def check(self):
        """The chosen folder must exist — a typo must not silently create one."""
        if self._checked:
            return
        if not os.path.isdir(self.root):
            raise TransportError(u"The folder '{}' cannot be reached. Check the path or the "
                                 u"network connection.".format(self.root))
        try:
            for folder in (self.presence_dir, self.chat_dir):
                if not os.path.isdir(folder):
                    os.makedirs(folder)
        except OSError as ex:
            raise TransportError(u"Cannot write to '{}': {}".format(self.root, ex))
        self._checked = True

    def publish(self, state):
        self.check()
        try:
            _write_json_atomic(os.path.join(self.presence_dir, state["key"] + u".json"), state)
        except PermissionError:
            pass                    # someone is reading it right now; the next round writes

    def leave(self, key):
        try:
            os.remove(os.path.join(self.presence_dir, key + u".json"))
        except OSError:
            pass

    def peers(self):
        """People whose file keeps changing. Judged on THIS PC's clock: a peer is
        here while its timestamp changed within TTL_SECONDS of local time, so a
        wrong clock on another PC, or a slow cloud sync, does not drop it. The
        first time a file is seen its own timestamp says how old it is (capped)."""
        now = self.clock()
        out = []
        try:
            names = os.listdir(self.presence_dir)
        except OSError:
            return out
        for name in names:
            if not name.endswith(u".json"):
                continue
            path = os.path.join(self.presence_dir, name)
            peer = read_state(_read_json(path))
            if peer is None:
                continue
            seen = self._observed.get(peer.key)
            if seen is None or seen[0] != peer.ts:
                if seen is None:
                    since = now - min(max(now - peer.ts, 0.0), FIRST_SIGHT_MAX)
                else:
                    since = now
                self._observed[peer.key] = seen = (peer.ts, since)
            if now - seen[1] <= TTL_SECONDS:
                out.append(peer)
            elif now - peer.ts > STALE_FILE_SECONDS:
                try:
                    os.remove(path)         # left by a Revit that crashed long ago
                except OSError:
                    pass
        return out

    def send(self, message):
        self.check()
        name = u"{:013d}_{}_{:04x}.json".format(int(message["ts"] * 1000), message["key"],
                                               random.randint(0, 0xffff))
        _write_json_atomic(os.path.join(self.chat_dir, name), message)

    def messages(self, cursor=None):
        """(new messages oldest first, cursor).

        The cursor is the set of files already read among the newest
        FOLDER_CHAT_KEEP — not "the newest name": file names start with the
        writer's clock, and a message synced late (OneDrive) or written by a PC
        whose clock is behind sorts BEFORE what was already read. A file that
        cannot be read yet (half synced) stays unread and is tried again."""
        try:
            names = sorted(n for n in os.listdir(self.chat_dir) if n.endswith(u".json"))
        except OSError:
            return [], cursor
        recent = names[-FOLDER_CHAT_KEEP:]
        if cursor is None:
            pick = recent[-HISTORY:]
            done = set(recent) - set(pick)
        else:
            done = set(n for n in recent if n in cursor)
            pick = [n for n in recent if n not in cursor]
        out = []
        for name in pick:
            msg = read_message(_read_json(os.path.join(self.chat_dir, name)))
            if msg is not None:
                out.append(msg)
                done.add(name)
        out.sort(key=lambda m: m.ts)
        return out, frozenset(done)


def _http(method, url, body=None, headers=None, timeout=HTTP_TIMEOUT):
    """(status, text). System.Net inside Revit, urllib elsewhere. Non-2xx returned."""
    headers = dict(headers or {})
    try:
        import System
        import System.Net as Net
        import System.IO as IO
        import System.Text as Text
    except Exception:
        Net = None
    if Net is not None:
        req = Net.WebRequest.Create(url)
        req.Method = method
        req.Timeout = int(timeout * 1000)
        content_type = headers.pop("Content-Type", None)
        accept = headers.pop("Accept", None)
        if accept:
            req.Accept = accept
        for k, v in headers.items():
            req.Headers.Add(k, v)
        if body is not None:
            payload = Text.Encoding.UTF8.GetBytes(body)
            req.ContentType = content_type or "application/json"
            req.ContentLength = payload.Length
            stream = req.GetRequestStream()
            try:
                stream.Write(payload, 0, payload.Length)
            finally:
                stream.Close()
        try:
            resp = req.GetResponse()
        except Net.WebException as wex:
            resp = wex.Response
            if resp is None:
                raise TransportError(u"No answer from the relay: {}".format(wex.Message))
        try:
            status = int(System.Convert.ToInt32(resp.StatusCode))
            reader = IO.StreamReader(resp.GetResponseStream(), Text.Encoding.UTF8)
            try:
                return status, reader.ReadToEnd()
            finally:
                reader.Close()
        finally:
            resp.Close()

    import urllib.request
    import urllib.error
    data = body.encode("utf-8") if body is not None else None
    req = urllib.request.Request(url, data=data, method=method)
    for k, v in headers.items():
        req.add_header(k, v)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return resp.status, resp.read().decode("utf-8", "replace")
    except urllib.error.HTTPError as err:
        return err.code, err.read().decode("utf-8", "replace")
    except Exception as ex:
        raise TransportError(u"No answer from the relay: {}".format(ex))


class SupabaseTransport(object):
    """Presence and chat in two Supabase tables, over its REST API."""

    kind = SUPABASE

    def __init__(self, url, key, room, http=None):
        url = (url or u"").strip().rstrip(u"/")
        if not re.match(r"^https://[\w.-]+(:\d+)?$", url):
            raise TransportError(u"Enter the Supabase project URL, like https://abcd1234.supabase.co")
        if not (key or u"").strip():
            raise TransportError(u"Enter the Supabase anon (public) key.")
        self.base = url + u"/rest/v1/"
        self.key = key.strip()
        self.room = room
        self.http = http or _http

    def describe(self):
        return u"Internet relay {}".format(self.base.split(u"/rest/")[0])

    def _headers(self, prefer=None):
        h = {"apikey": self.key, "Authorization": "Bearer " + self.key,
             "Content-Type": "application/json", "Accept": "application/json"}
        if prefer:
            h["Prefer"] = prefer
        return h

    def _call(self, method, path, body=None, prefer=None):
        status, text = self.http(method, self.base + path,
                                 json.dumps(body) if body is not None else None,
                                 self._headers(prefer))
        if status in (401, 403):
            raise TransportError(u"The relay refused the key ({}). Check the anon key and "
                                 u"the table policies in docs/model-meet.md.".format(status))
        if status == 404:
            raise TransportError(u"The relay has no Model Meet tables yet. Run the SQL in "
                                 u"docs/model-meet.md in the Supabase SQL editor.")
        if not 200 <= status < 300:
            raise TransportError(u"The relay answered {}: {}".format(status, (text or u"")[:160]))
        if not text:
            return None
        try:
            return json.loads(text)
        except ValueError:
            return None

    def publish(self, state):
        self._call("POST", u"{}?on_conflict=room,user_key".format(PRESENCE_TABLE),
                   [{"room": self.room, "user_key": state["key"], "payload": state}],
                   prefer="resolution=merge-duplicates,return=minimal")

    def leave(self, key):
        self._call("DELETE", u"{}?room=eq.{}&user_key=eq.{}".format(PRESENCE_TABLE, self.room, key),
                   prefer="return=minimal")

    def peers(self):
        rows = self._call("GET", u"{}?select=user_key,payload,updated_at&room=eq.{}".format(
            PRESENCE_TABLE, self.room)) or []
        stamped = [(parse_iso(r.get("updated_at")), r) for r in rows if isinstance(r, dict)]
        times = [t for t, _r in stamped if t is not None]
        if not times:
            return []
        server_now = max(times)          # our own row was just written: ~server time
        out = []
        for t, row in stamped:
            if t is None or server_now - t > TTL_SECONDS:
                continue
            peer = read_state(row.get("payload"))
            if peer is not None:
                peer.ts = t
                out.append(peer)
        return out

    def send(self, message):
        self._call("POST", MESSAGE_TABLE,
                   [{"room": self.room, "user_key": message["key"], "payload": message}],
                   prefer="return=minimal")

    def messages(self, cursor=None):
        if cursor is None:
            rows = self._call("GET", u"{}?select=id,payload&room=eq.{}&order=id.desc&limit={}".format(
                MESSAGE_TABLE, self.room, HISTORY)) or []
            rows = list(reversed(rows))
        else:
            rows = self._call("GET", u"{}?select=id,payload&room=eq.{}&id=gt.{}&order=id.asc&limit=100".format(
                MESSAGE_TABLE, self.room, int(cursor))) or []
        out = []
        newest = cursor
        for row in rows:
            if not isinstance(row, dict):
                continue
            msg = read_message(row.get("payload"))
            if msg is not None:
                out.append(msg)
            try:
                newest = max(int(row.get("id")), int(newest or 0))
            except (TypeError, ValueError):
                pass
        return out, (newest if newest is not None else 0)


def make_transport(settings, room):
    """Transport for the saved settings. Raises TransportError with a fix-it message."""
    if settings.get("transport") == SUPABASE:
        return SupabaseTransport(settings.get("supabase_url"), settings.get("supabase_key"), room)
    return FolderTransport(settings.get("folder"), room)


# ── the background client ────────────────────────────────────────────────────
class PresenceClient(object):
    """Publishes my state and collects the room on a background thread.

    UI thread: set_state() / say() / snapshot() — never blocks on the network.
    """

    def __init__(self, transport, me_key, interval=INTERVAL):
        self.transport = transport
        self.me_key = me_key
        self.interval = interval
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._thread = None
        self._state = None
        self._outbox = []
        self._peers = []
        self._inbox = []
        self._error = u""
        self._cursor = None
        self._rounds = 0
        self.last_ok = 0.0

    # UI side
    def set_state(self, state):
        with self._lock:
            self._state = state

    def say(self, message):
        with self._lock:
            self._outbox.append(message)

    def snapshot(self):
        """(peers, new messages since the last call, error text)."""
        with self._lock:
            inbox, self._inbox = self._inbox, []
            return list(self._peers), inbox, self._error

    @property
    def running(self):
        return self._thread is not None and self._thread.is_alive()

    def start(self):
        if self.running:
            return
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, name="T3Lab Model Meet")
        self._thread.daemon = True
        self._thread.start()

    def stop(self, timeout=2.0):
        """Leave the room. Returns once the thread said goodbye or `timeout`
        passed — a thread stuck in a slow call leaves when that call returns,
        and the room drops it after TTL_SECONDS anyway."""
        self._stop.set()
        thread = self._thread
        if thread is not None:
            thread.join(timeout)

    # network side
    def _leave(self):
        try:
            self.transport.leave(self.me_key)
        except Exception:
            pass

    def round(self):
        """One network round. Raises on failure (the loop reports it)."""
        with self._lock:
            state = self._state
            outbox, self._outbox = self._outbox, []
        sent = 0
        try:
            if state is not None:
                self.transport.publish(state)
            for message in outbox:
                self.transport.send(message)
                sent += 1
        except Exception:
            with self._lock:
                self._outbox[0:0] = outbox[sent:]       # nothing unsent is ever lost
            raise
        peers = [p for p in self.transport.peers() if p.key != self.me_key]
        messages = []
        if self._rounds % 2 == 0 or outbox:
            messages, self._cursor = self.transport.messages(self._cursor)
        self._rounds += 1
        with self._lock:
            self._peers = peers
            self._inbox.extend(messages)
            self._error = u""
        self.last_ok = time.time()

    def _loop(self):
        delay = self.interval
        while not self._stop.is_set():
            try:
                self.round()
                delay = self.interval
            except Exception as ex:
                text = u"{}".format(ex) if isinstance(ex, TransportError) else \
                    u"Connection problem: {}".format(ex)
                with self._lock:
                    self._error = text
                delay = min(15.0, max(2.0, delay * 2))
            self._stop.wait(delay)
        self._leave()


# ── settings ─────────────────────────────────────────────────────────────────
DEFAULTS = {
    "display_name": u"",
    "transport": FOLDER,
    "folder": u"",
    "supabase_url": u"",
    "supabase_key": u"",
    "auto_connect": False,
    "rooms": {},            # room seed hash -> room code typed by the user
    "folders": {},          # room seed hash -> shared folder used for that model
    "look": {},             # the character: colour, build, hat, height
    "models": [],           # models walked on this PC, newest first (remember_model)
    "world": {},            # View panel: quality, colour mode, collision, frame rate shown
    "bookmarks": {},        # room seed hash -> [bookmark] (Snippets/model_meet_places.py)
}
MAX_MODELS = 30


def settings_path():
    from core.paths import user_data_path
    return user_data_path("model_meet", "settings.json")


def load_settings(path=None):
    data = dict(DEFAULTS)
    stored = _read_json(path or settings_path())
    if isinstance(stored, dict):
        for k in DEFAULTS:
            if k in stored and isinstance(stored[k], type(DEFAULTS[k])):
                data[k] = stored[k]
    for key in ("rooms", "folders", "look", "world", "bookmarks"):
        data[key] = dict(data[key])         # never hand out the DEFAULTS dicts themselves
    data["models"] = list(data["models"])
    return data


def save_settings(data, path=None):
    clean = dict((k, data.get(k, DEFAULTS[k])) for k in DEFAULTS)
    _write_json_atomic(path or settings_path(), clean)


def remember_model(settings, key, title, created, edited, borrowed=0, when=None):
    """Put this model at the top of the 'Your models' list (one entry per model)."""
    entry = {"key": key, "title": _text(title, 120) or u"Untitled",
             "last": when or datetime.now().strftime("%Y-%m-%d %H:%M"),
             "created": int(created), "edited": int(edited), "borrowed": int(borrowed)}
    rest = [m for m in settings.get("models") or [] if isinstance(m, dict) and m.get("key") != key]
    settings["models"] = [entry] + rest[:MAX_MODELS - 1]
    return settings["models"]


def seed_hash(seed):
    """Key under which a typed room code is remembered for one model."""
    return hashlib.sha1((seed or u"").encode("utf-8")).hexdigest()[:16]
