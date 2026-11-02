# -*- coding: utf-8 -*-
"""
Consent for usage tracking (Lite only).

Nothing is sent until the user says yes. Vietnam's Personal Data Protection
Law (Luat 91/2025/QH15, Art. 9) requires consent that is voluntary, clear and
verifiable, and silence or no answer is not consent. GDPR says the same for
EU users. So:

  - No answer stored  -> not sent. Revit asks once, on the first Idling after
                         start (like the What's new popup).
  - "Allow"           -> sent. The answer, its time and CONSENT_VERSION are
                         stored, so it can be shown later (verifiable).
  - "Don't allow"     -> never sent, never asked again for this version.
  - Closing the dialog -> nothing stored, asked again on the next Revit start.

Bump CONSENT_VERSION whenever the data sent changes in a way users must
know about: everyone is asked again and old "Allow" answers stop counting.

Withdraw at any time: set "tracking_consent" -> "value": false in
%APPDATA%\\T3LabAI\\mcp_paths.json, or the old switch
"tracking_enabled": false, which always wins.

Runs under IronPython 2.7 (startup.py) and CPython 3. Never raises.
Test: python3 lite_guard/test_tracking_privacy.py
"""
from __future__ import unicode_literals

from datetime import datetime

CONSENT_KEY = 'tracking_consent'
LEGACY_SWITCH = 'tracking_enabled'
CONSENT_VERSION = 1
PRIVACY_URL = 'https://github.com/thanhtranarch/T3LabLite/blob/main/docs/PRIVACY.md'

# ask_when_idle() states
PENDING = 'pending'
SKIPPED = 'skipped'
FAILED = 'failed'

# show_dialog() answers
ALLOW = 'allow'
DENY = 'deny'
NO_ANSWER = None


def _paths():
    from core import paths
    return paths


def _settings():
    try:
        return _paths().load_settings() or {}
    except Exception:
        return {}


def current_answer(settings=None):
    """True (allowed), False (refused) or None (never answered for this version)."""
    settings = _settings() if settings is None else settings
    if settings.get(LEGACY_SWITCH) is False:
        return False
    record = settings.get(CONSENT_KEY)
    if not isinstance(record, dict):
        return None
    if record.get('version') != CONSENT_VERSION:
        # A refusal stays a refusal; an old "Allow" must be asked again.
        return False if record.get('value') is False else None
    value = record.get('value')
    return value if isinstance(value, bool) else None


def has_consent(settings=None):
    """Send only on an explicit, current "Allow"."""
    return current_answer(settings) is True


def record_answer(allowed, via='dialog'):
    """Store the answer with its time and version (verifiable consent)."""
    try:
        _paths().set_setting(CONSENT_KEY, {
            'value': bool(allowed),
            'version': CONSENT_VERSION,
            'at': datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ'),
            'via': via,
        })
        return True
    except Exception:
        return False


DIALOG_TITLE = 'T3Lab Lite'
DIALOG_INSTRUCTION = 'Help improve T3Lab Lite with usage statistics?'
DIALOG_BODY = (
    'If you allow it, T3Lab Lite sends t3lab.space a small record each time '
    'Revit starts, you click a T3Lab button or an AI app calls a T3Lab MCP '
    'tool:\n'
    '  - the tool name and whether it worked\n'
    '  - your Revit version\n'
    '  - a random ID made on this computer (not your name or computer name)\n'
    '  - a code made from the model name (not the name itself)\n\n'
    'Never sent: your model, code, file paths, prompts, text you type, '
    'API keys. The server also sees your IP address.\n\n'
    'Used only to see which tools people use and which fail. '
    'Nothing is sent unless you choose Allow. You can change your mind at '
    'any time.'
)
ALLOW_TEXT = 'Allow usage statistics'
DENY_TEXT = "Don't allow"
FOOTER = 'Privacy notice: ' + PRIVACY_URL


def show_dialog():
    """Revit TaskDialog. Returns ALLOW, DENY or NO_ANSWER (closed)."""
    from Autodesk.Revit.UI import (TaskDialog, TaskDialogCommandLinkId,
                                   TaskDialogCommonButtons, TaskDialogIcon,
                                   TaskDialogResult)
    dialog = TaskDialog(DIALOG_TITLE)
    dialog.MainInstruction = DIALOG_INSTRUCTION
    dialog.MainIcon = TaskDialogIcon.TaskDialogIconInformation
    dialog.MainContent = DIALOG_BODY
    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, ALLOW_TEXT)
    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, DENY_TEXT)
    dialog.CommonButtons = TaskDialogCommonButtons.Close
    # Enter must never mean yes.
    dialog.DefaultButton = TaskDialogResult.CommandLink2
    dialog.FooterText = FOOTER
    result = dialog.Show()
    if result == TaskDialogResult.CommandLink1:
        return ALLOW
    if result == TaskDialogResult.CommandLink2:
        return DENY
    return NO_ANSWER


def _idling_app(revit_handle):
    candidates = [revit_handle]
    try:
        from pyrevit import HOST_APP
        candidates.append(getattr(HOST_APP, 'uiapp', None))
    except Exception:
        pass
    for app in candidates:
        if app is not None and hasattr(app, 'Idling'):
            return app
    return None


def ask_when_idle(revit_handle=None, show=None, on_allow=None):
    """Ask on the first Idling when this version was never answered.

    `on_allow` runs after an "Allow" (startup sends the session row then).
    Returns (state, detail).
    """
    try:
        if current_answer() is not None:
            return SKIPPED, 'already answered'
        app = _idling_app(revit_handle)
        if app is None:
            return FAILED, 'no Idling event'
        show = show or show_dialog

        def on_idling(sender, args):
            try:
                app.Idling -= on_idling
            except Exception:
                pass
            # A pyRevit Reload before this Idling hooks a second handler.
            if current_answer() is not None:
                return
            try:
                answer = show()
            except Exception:
                return
            if answer == ALLOW:
                record_answer(True)
                if on_allow is not None:
                    try:
                        on_allow()
                    except Exception:
                        pass
            elif answer == DENY:
                record_answer(False)

        app.Idling += on_idling
        return PENDING, 'asks on first Idling'
    except Exception as exc:
        return FAILED, str(exc)
