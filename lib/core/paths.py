# -*- coding: utf-8 -*-
"""
Shared, user-editable path settings for the MCP layer.

Every machine-specific path used by the MCP server / Control dialog — the
task-watcher data directory, the Python interpreter used to launch
bridge.py, the Claude Desktop config file — is resolved through a single
JSON file:

    %APPDATA%\\T3LabAI\\mcp_paths.json

That file lives outside the repo (same folder as mcp_token.txt), so it is
never committed and never carries any developer's local paths. The first
time a path is needed, a sensible per-OS default is computed and written
back to the file; from then on the file is the source of truth and can be
hand-edited to relocate any of these paths without touching code.
"""

from __future__ import unicode_literals

import io
import json
import os
import threading
import time

# One lock for every read-modify-write of mcp_paths.json in this process: the
# HTTP threads (teaching settings), the UI (MCP Control) and the overlay all
# write it, and an unlocked writer can read {} mid-write and save it back.
_LOCK = threading.RLock()


def settings_dir():
    """The %APPDATA%\\T3LabAI directory, created on first use."""
    app_data = os.environ.get('APPDATA', os.path.expanduser('~'))
    d = os.path.join(app_data, 'T3LabAI')
    try:
        if not os.path.isdir(d):
            os.makedirs(d)
    except OSError:
        pass
    return d


def user_data_path(*parts, **kwargs):
    """Path of a per-user T3Lab data file under %APPDATA%\\T3LabAI.

    Tools keep their state HERE, never inside the extension folder: the clone
    is updated by git and shared, and state written next to the scripts was
    committed to the public repo (client names, machine paths — 2026-10-02).

    legacy=[old paths]: when the new file does not exist yet, the first old
    file that does is copied over once, so moving a writer never loses the
    user's settings. The old file is left untouched.
    """
    legacy = kwargs.get('legacy') or ()
    path = os.path.join(settings_dir(), *parts)
    folder = os.path.dirname(path)
    try:
        if not os.path.isdir(folder):
            os.makedirs(folder)
    except OSError:
        pass
    if not os.path.exists(path):
        for old in legacy:
            if old and os.path.isfile(old):
                try:
                    import shutil
                    shutil.copyfile(old, path)
                    break
                except Exception:
                    pass
    return path


def settings_file():
    """Path to the shared path-settings JSON file."""
    return os.path.join(settings_dir(), 'mcp_paths.json')


def replace_file(tmp, dst):
    """Move tmp over dst so dst is never missing or half-written.
    CPython: os.replace. IronPython 2.7 has none: File.Copy(overwrite)."""
    replace = getattr(os, 'replace', None)
    if replace is not None:
        replace(tmp, dst)
        return
    try:
        from System.IO import File
    except ImportError:
        File = None
    if File is not None and os.path.exists(dst):
        File.Copy(tmp, dst, True)
        os.remove(tmp)
        return
    if os.path.exists(dst):
        os.remove(dst)
    os.rename(tmp, dst)


def atomic_write(path, data):
    """Write bytes to `path` through a temp file next to it (one retry when
    antivirus / an indexer briefly holds the file)."""
    tmp = path + '.t3lab-tmp'
    with io.open(tmp, 'wb') as f:
        f.write(data)
    try:
        replace_file(tmp, path)
    except (IOError, OSError):
        time.sleep(0.25)
        replace_file(tmp, path)


def _read_settings():
    """(dict, ok). ok is False when the file exists but cannot be used as a
    settings object (bad JSON, not an object, unreadable)."""
    path = settings_file()
    if not os.path.isfile(path):
        return {}, True
    try:
        with io.open(path, 'rb') as f:
            raw = f.read()
        # utf-8-sig: Notepad / PowerShell 5.1 save UTF-8 with a BOM.
        text = raw.decode('utf-8-sig').strip()
        if not text:
            return {}, True
        data = json.loads(text)
    except Exception:
        return {}, False
    if not isinstance(data, dict):
        return {}, False
    return data, True


def load_settings():
    """Return the settings dict, or {} if the file is missing/unreadable."""
    with _LOCK:
        return _read_settings()[0]


def save_settings(data):
    """Persist the settings dict, best-effort.

    Written through a temp file, so a crash or a concurrent reader never sees
    a truncated file. A file that existed but could not be parsed is kept as
    mcp_paths.json.unreadable first, so hand-edited values are not lost."""
    with _LOCK:
        try:
            path = settings_file()
            _old, ok = _read_settings()
            if not ok:
                try:
                    with io.open(path, 'rb') as f:
                        bad = f.read()
                    with io.open(path + '.unreadable', 'wb') as f:
                        f.write(bad)
                except Exception:
                    pass
            text = json.dumps(data, indent=2)
            if not isinstance(text, bytes):
                text = text.encode('utf-8')
            atomic_write(path, text)
        except Exception:
            pass


def get_setting(key, default_fn):
    """
    Read `key` from mcp_paths.json.

    If absent (first run, or the user cleared it), compute a default via
    default_fn() and persist it — so the JSON file becomes the editable
    source of truth for every call after this one.
    """
    with _LOCK:
        data = load_settings()
        value = data.get(key)
        if value:
            return value
        value = default_fn()
        data[key] = value
        save_settings(data)
        return value


def set_setting(key, value):
    """Explicitly set/override `key` in mcp_paths.json."""
    with _LOCK:
        data = load_settings()
        data[key] = value
        save_settings(data)
