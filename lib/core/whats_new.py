# -*- coding: utf-8 -*-
""""What's new" popup once per version, after an update (Lite only).

startup.py calls install() on every Revit start. When version.txt is newer
than the version last shown on this machine (%APPDATA%\\T3LabAI\\mcp_paths.json
-> "whats_new_seen"), the first Idling shows a TaskDialog with that release's
CHANGELOG.md section (every release since the last one shown, after skipped
updates). A machine that never saw a popup gets the current release only.
The version is marked as seen before the dialog opens, so a Revit crash or a
pyRevit Reload never shows it twice; a Revit closed before its first Idling
shows it on the next start.

extract_whats_new() is also what Check Update lists before updating
(core/updater.py re-exports it).

Runs under IronPython 2.7 (startup.py) and CPython 3 (tests):
lite_guard/test_whats_new.py. Nothing here raises into startup.py.
"""
from __future__ import unicode_literals

import io
import os
import re

EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
VERSION_FILE = 'version.txt'
CHANGELOG_FILE = 'CHANGELOG.md'
CHANGELOG_URL = 'https://github.com/thanhtranarch/T3LabLite/blob/main/CHANGELOG.md'

SEEN_KEY = 'whats_new_seen'
MAIN_LINES = 14         # in the dialog body; the rest goes under "Show details"
MAX_LINES = 60

# install() states
PENDING = 'pending'
SKIPPED = 'skipped'
FAILED = 'failed'

_LINK_RE = re.compile(r'\[([^\]]+)\]\(([^)\s]+)\)')


def parse_version(text):
    """'1.2.3' -> (1, 2, 3); tolerates stray characters."""
    parts = []
    for token in (text or '').strip().split('.'):
        digits = ''.join(ch for ch in token if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts) if parts else (0,)


def _clean(text):
    """Markdown -> plain text for a TaskDialog / forms.alert."""
    text = _LINK_RE.sub(lambda m: '{} ({})'.format(m.group(1), m.group(2)), text)
    return text.replace('**', '').replace('`', '')


def extract_whats_new(changelog_text, local_version, max_lines=20, upto=None,
                      exact=None):
    """Display lines for every release newer than local_version (and not newer
    than upto), or for the `exact` release only, from Keep-a-Changelog
    '## [x.y.z] - date' headings. Nested bullets keep one indent level.
    Returns [] when nothing matches."""
    local = parse_version(local_version)
    top = parse_version(upto) if upto else None
    only = parse_version(exact) if exact else None
    include = False
    out = []
    for raw in (changelog_text or '').replace('\r\n', '\n').split('\n'):
        line = raw.strip()
        if line.startswith('## '):
            header = line[3:].strip()
            ver_text = header.lstrip('[').split(']')[0].strip()
            if not any(ch.isdigit() for ch in ver_text):     # [Unreleased]
                include = False
                continue
            ver = parse_version(ver_text)
            if only is not None:
                include = ver == only
            else:
                include = ver > local and (top is None or ver <= top)
            if include:
                if out:
                    out.append('')
                out.append('Version ' + header.replace('[', '').replace(']', ''))
        elif include:
            if line.startswith('### '):
                out.append(line[4:].strip() + ':')
            elif line.startswith('- ') or line.startswith('* '):
                nested = len(raw) - len(raw.lstrip()) >= 2
                out.append(('      - ' if nested else '  • ') + line[2:].strip())
            elif line and out and out[-1].startswith(('  • ', '      - ')):
                out[-1] += ' ' + line                          # wrapped bullet
    if len(out) > max_lines:
        out = out[:max_lines] + ['  ...']
    return [_clean(item) for item in out]


def _read_text(path):
    try:
        with io.open(path, 'r', encoding='utf-8-sig') as fh:
            return fh.read()
    except Exception:
        return ''


def local_version(ext_dir=EXTENSION_DIR):
    return _read_text(os.path.join(ext_dir, VERSION_FILE)).strip() or '0.0.0'


def is_pending(current, seen):
    """Show when this machine has not seen `current` yet; a downgrade never shows."""
    return not seen or parse_version(current) > parse_version(seen)


def notes_for(changelog_text, current, seen, max_lines=MAX_LINES):
    """Lines for the popup: releases after `seen` up to `current`, or the
    `current` release alone when nothing was ever shown."""
    if seen:
        return extract_whats_new(changelog_text, seen, max_lines=max_lines, upto=current)
    return extract_whats_new(changelog_text, '0', max_lines=max_lines, exact=current)


def _settings():
    from core import paths
    return paths


def mark_seen(version):
    try:
        _settings().set_setting(SEEN_KEY, version)
        return True
    except Exception:
        return False


def _open_url(url):
    from System.Diagnostics import Process, ProcessStartInfo
    psi = ProcessStartInfo(url)
    psi.UseShellExecute = True        # .NET 8 (Revit 2025+) defaults to False
    Process.Start(psi)


def show_dialog(version, lines):
    """Revit TaskDialog; "Show details" holds what does not fit the body."""
    from Autodesk.Revit.UI import (TaskDialog, TaskDialogCommandLinkId,
                                   TaskDialogCommonButtons, TaskDialogIcon,
                                   TaskDialogResult)
    dialog = TaskDialog('T3Lab Lite')
    dialog.MainInstruction = "What's new in T3Lab Lite {}".format(version)
    dialog.MainIcon = TaskDialogIcon.TaskDialogIconInformation
    dialog.MainContent = '\n'.join(lines[:MAIN_LINES])
    if len(lines) > MAIN_LINES:
        dialog.ExpandedContent = '\n'.join(lines[MAIN_LINES:])
    dialog.AddCommandLink(TaskDialogCommandLinkId.CommandLink1,
                          'Open the full changelog')
    dialog.CommonButtons = TaskDialogCommonButtons.Close
    dialog.DefaultButton = TaskDialogResult.Close
    dialog.FooterText = 'Shown once after each T3Lab Lite update.'
    if dialog.Show() == TaskDialogResult.CommandLink1:
        try:
            _open_url(CHANGELOG_URL)
        except Exception:
            pass


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


def show_when_idle(app, version, lines, show=None):
    """Show on the first Idling (Revit's window is up by then), once."""
    show = show or show_dialog

    def on_idling(sender, args):
        try:
            app.Idling -= on_idling
        except Exception:
            pass
        # A pyRevit Reload before this Idling hooks a second handler.
        if not is_pending(version, _settings().load_settings().get(SEEN_KEY)):
            return
        mark_seen(version)
        try:
            show(version, lines)
        except Exception:
            pass

    app.Idling += on_idling


def install(revit_handle, ext_dir=EXTENSION_DIR):
    """Arm the popup when this version is new here. Returns (state, detail)."""
    try:
        current = local_version(ext_dir)
        seen = _settings().load_settings().get(SEEN_KEY)
        if not is_pending(current, seen):
            return SKIPPED, 'already shown for ' + current
        lines = notes_for(_read_text(os.path.join(ext_dir, CHANGELOG_FILE)), current, seen)
        if not lines:
            mark_seen(current)
            return SKIPPED, 'no CHANGELOG entry for ' + current
        app = _idling_app(revit_handle)
        if app is None:
            return FAILED, 'no Idling event at startup'
        show_when_idle(app, current, lines)
        return PENDING, current
    except Exception as exc:
        return FAILED, str(exc)
