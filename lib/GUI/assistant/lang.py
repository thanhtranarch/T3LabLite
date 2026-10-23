# -*- coding: utf-8 -*-
"""
Assistant language — which language (VI / EN) the assistant answers in, the
time-of-day greeting and relative timestamps.

`_LAST_USER_VIET` is process state written by `_is_viet_text` and
`_note_user_language`; read it through this module (`lang._LAST_USER_VIET`),
never through a `from ... import` copy. GUI.T3LabAssistantDialog serves the
live value for old callers.

Split out of GUI/T3LabAssistantDialog.py, which re-imports every name
here: `from GUI.T3LabAssistantDialog import <name>` keeps working.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Language"

# Pure Python: a failure here is a broken install, so let it surface.
from GUI import AssistantGreetings as _greetings
from GUI.assistant.deps import HAS_LANGUAGE, _lang_detect


_VIET_CHARS = (u"àáâãèéêìíòóôõùúýăđơưạảấầẩẫậắằẳẵặẹẻẽếềểễệỉịọỏốồổỗộớờởỡợ"
               u"ụủứừửữựỳỵỷỹ")


def _reply_language():
    """User's language preference: 'auto' | 'vi' | 'en'. Never raises."""
    try:
        from config.settings import get_settings
        return get_settings().get_reply_language()
    except Exception:
        return 'auto'


def _is_viet_text(text):
    """True when the assistant should answer this turn in Vietnamese.

    The setting decides; 'auto' detects from the text (2026-07-28: hard-wired
    False mixed languages, unbranched VI strings still rendered).
    """
    lang = _reply_language()
    if lang == 'vi':
        return True
    if lang == 'en':
        return False
    if not text:
        return _ui_viet()
    viet = _detect_viet(text, default=_ui_viet())
    if viet:
        # Turn-less messages (startup notices, card labels) follow the user.
        global _LAST_USER_VIET
        _LAST_USER_VIET = True
    return viet


def _detect_viet(text, default=True):
    """Language verdict for one message, with a fallback for the undecidable.

    `Intelligence.language` weighs evidence (diacritic-less typing,
    code-switching) and returns `default` when there is none, so a bare tool
    name ("BatchOut") changes nothing.
    """
    if HAS_LANGUAGE:
        try:
            return _lang_detect.is_vietnamese(text, default=default)
        except Exception:
            pass
    return any(c in _VIET_CHARS for c in (text or u'').lower())


# Language of the latest user message, for turn-less strings; None until typed.
_LAST_USER_VIET = None


def _note_user_language(text):
    """Record the language of a user message. Call once per submitted turn.

    An undecidable message ("ok", a tool name) leaves the session language alone.
    """
    global _LAST_USER_VIET
    if text:
        _LAST_USER_VIET = _detect_viet(
            text, default=(True if _LAST_USER_VIET is None
                           else _LAST_USER_VIET))


def _relative_time(ts_text, viet=True):
    """'hôm nay 16:09' / '3 ngày trước' from a stored 'Y-m-d H:M:S' stamp.

    Returns u'' for anything unparseable (old/hand-edited history). Never raises.
    """
    if not ts_text:
        return u""
    try:
        import datetime as _dt
        then = _dt.datetime.strptime(u"{}".format(ts_text).strip(),
                                     "%Y-%m-%d %H:%M:%S")
    except Exception:
        return u""
    try:
        now = _dt.datetime.now()
        days = (now.date() - then.date()).days
        clock = then.strftime("%H:%M")
        if days <= 0:
            return (u"hôm nay lúc {}" if viet else u"today at {}").format(clock)
        if days == 1:
            return (u"hôm qua lúc {}" if viet else u"yesterday at {}").format(clock)
        if days < 7:
            return (u"{} ngày trước" if viet else u"{} days ago").format(days)
        return then.strftime("%d/%m/%Y" if viet else "%d %b %Y")
    except Exception:
        return u""


def _time_greeting(hour=None, viet=None):
    """Time-of-day greeting phrase (no name). Never raises.

    Buckets/phrases: GUI/AssistantGreetings.py. Used by the onboarding card;
    the welcome headline uses the richer pool (_render_greeting).
    """
    try:
        if viet is None:
            viet = _ui_viet()
        return _greetings.base_greeting(hour, bool(viet))
    except Exception:
        return u"Xin chào" if viet else u"Hello"


def _ui_viet():
    """Language for UI strings with no user text of their own."""
    lang = _reply_language()
    if lang == 'vi':
        return True
    if lang == 'en':
        return False
    return bool(_LAST_USER_VIET)
