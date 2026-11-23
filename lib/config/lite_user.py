# -*- coding: utf-8 -*-
"""T3Lab Lite: the name the Assistant calls its user. Lite only.

Kept in its own file, %APPDATA%\\T3LabAI\\t3lablite_user.json, next to (not
inside) the user_profile.json / settings.json that T3Lab AI (t3lab_dev) also
writes on the same machine, so the two extensions never overwrite each
other's name (owner, 2026-11-17).

- Never a built-in name: "Thạnh" was the shipped default, and LLM Settings
  pre-filled it, so Save stored it as every user's name.
- First call: a name the user set in an earlier install (user_profile.json,
  then settings.json "username") is copied over once, unless it is that
  shipped default.
- No name anywhere: needs_name() is True until the user answers the one-time
  question (show_assistant_dialog) or cancels it.

Engine-agnostic (IronPython 2.7 + CPython 3). Test: lite_guard/test_lite_user.py.
"""
from __future__ import unicode_literals

import datetime
import io
import json
from core import jsonsafe
import os

FILE_NAME = 't3lablite_user.json'
SHIPPED_DEFAULTS = (u'Thạnh', u'Thanh')     # never a real choice
MAX_LEN = 40


def _dir():
    base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI')


def path():
    return os.path.join(_dir(), FILE_NAME)


def clean(name):
    """Trimmed name, or '' when empty or a shipped default."""
    text = u'{}'.format(name or u'').strip()[:MAX_LEN]
    return u'' if text in SHIPPED_DEFAULTS else text


def _read():
    try:
        with io.open(path(), 'r', encoding='utf-8') as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _write(data):
    try:
        if not os.path.isdir(_dir()):
            os.makedirs(_dir())
        data['updated_at'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        text = jsonsafe.dumps(data, indent=2)
        with io.open(path(), 'w', encoding='utf-8') as f:
            f.write(text)
        return True
    except Exception:
        return False


def _earlier_install_name():
    """A name the user typed in an earlier install, or ''."""
    try:
        with io.open(os.path.join(_dir(), 'user_profile.json'), 'r', encoding='utf-8') as f:
            name = clean(((json.load(f) or {}).get('user') or {}).get('name'))
        if name:
            return name
    except Exception:
        pass
    try:
        with io.open(os.path.join(_dir(), 'settings.json'), 'r', encoding='utf-8') as f:
            return clean((json.load(f) or {}).get('username'))
    except Exception:
        return u''


def get_name():
    """The user's name, or '' (callers greet without a name)."""
    data = _read()
    name = clean(data.get('name'))
    if name or data.get('migrated'):
        return name
    name = _earlier_install_name()
    data.update({'name': name, 'migrated': True,
                 'source': 'earlier install' if name else ''})
    _write(data)
    return name


def set_name(name):
    """Save the name the user typed; '' clears it. Returns the stored name."""
    data = _read()
    name = clean(name)
    data.update({'name': name, 'migrated': True, 'asked': True, 'source': 'typed'})
    _write(data)
    return name


def needs_name():
    """True while no name is known and the user was never asked."""
    return not get_name() and not _read().get('asked')


def mark_asked():
    """The user cancelled the question: do not ask again (Settings can set it)."""
    data = _read()
    data['asked'] = True
    data.setdefault('migrated', True)
    return _write(data)
