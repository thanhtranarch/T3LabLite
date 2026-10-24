# -*- coding: utf-8 -*-
"""
Assistant text helpers — never-raising conversions used on every chat turn:
reasoning-block hiding, exception text, JSON text and tool-call summaries.

Pure Python (no .NET, no Revit).

Split out of GUI/T3LabAssistantDialog.py, which re-imports every name
here: `from GUI.T3LabAssistantDialog import <name>` keeps working.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Text"

import json
import re


_THINK_BLOCK_RE = re.compile(
    r'<(think|thinking|reasoning|reflection)>[\s\S]*?</\1\s*>', re.IGNORECASE)
_THINK_OPEN_RE = re.compile(
    r'<(think|thinking|reasoning|reflection)>[\s\S]*\Z', re.IGNORECASE)


def _hide_reasoning(text):
    """Hide model chain-of-thought from anything user-visible.

    An unterminated opening tag (mid-stream) hides everything after it.
    Guards the LIVE stream; providers strip the final text themselves.
    """
    if not text:
        return text
    out = _THINK_BLOCK_RE.sub('', text)
    out = _THINK_OPEN_RE.sub('', out)
    return out.lstrip()


def _exc_text(exc):
    """Exception message as unicode. Never raises.

    `str(exc)` under IronPython 2.7 throws UnicodeEncodeError on non-ASCII
    messages, from inside the error handler, killing the chat turn.
    """
    try:
        return u"{}".format(exc)
    except Exception:
        pass
    try:
        r = repr(exc)
        return r.decode('utf-8', 'replace') if isinstance(r, bytes) else u"{}".format(r)
    except Exception:
        return u"<unprintable error>"


def _json_safe(obj):
    """Recursively coerce byte strings inside `obj` to unicode. Never raises.

    Legacy code-page bytes from .NET/OS (non-UTF-8) are the one input json's
    encoder cannot survive; see _json_text.
    """
    if isinstance(obj, bytes):
        try:
            return obj.decode('utf-8')
        except Exception:
            return obj.decode('latin-1', 'replace')
    if isinstance(obj, dict):
        out = {}
        for k, v in obj.items():
            out[_json_safe(k)] = _json_safe(v)
        return out
    if isinstance(obj, (list, tuple, set)):
        return [_json_safe(v) for v in obj]
    return obj


def _json_text(obj, sort_keys=False):
    """`json.dumps` as unicode that can never kill a chat turn.

    IronPython 2.7's encoder decodes byte strings as UTF-8; code-page bytes
    raised past every handler up to _route_input ("'unknown' codec can't
    decode byte 0xe0"). Sanitising first removes that input.
    """
    try:
        out = json.dumps(_json_safe(obj), ensure_ascii=False,
                         sort_keys=sort_keys)
        return out.decode('utf-8', 'replace') if isinstance(out, bytes) \
            else u"{}".format(out)
    except Exception:
        pass
    try:
        # ASCII-escaped via jsonsafe: plain `ensure_ascii=True` would raise on
        # the same non-ASCII input that failed rung one.
        from core import jsonsafe
        return jsonsafe.dumps(_json_safe(obj), sort_keys=sort_keys,
                              default=_exc_text)
    except Exception:
        return u"<unserializable {}>".format(type(obj).__name__)


def _call_summary(name, args):
    """(headline, details): what a tool call does, in words — never raises.

    The tool cards show this instead of raw JSON (and send_code_to_revit's
    code); see core.call_summary.
    """
    try:
        from core.call_summary import summarize
        return summarize(name, args)
    except Exception:
        return (u"{}".format(name), [])
