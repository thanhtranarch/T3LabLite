# -*- coding: utf-8 -*-
"""Move ribbon folders an update left behind out of the extension (Lite only).

After 1.5.0 the T3Lab ribbon was regrouped the way t3lab-revit-api has it:
six panels (Support, Standards & Families, Model & Datum, Rebar & Assembly,
Annotation & Data, Views & Sheets), and the buttons moved into new panels,
stacks and pulldowns.

A `git pull` removes the old folders, but the download-and-copy update (no
git.exe on the machine, or "Download latest version" in Check Update) only
ADDS files, so every old panel folder stays on disk. pyRevit 6 builds every
folder of a tab or panel, listed under `layout:` or not, so those leftovers
showed up as a second copy of the old panels, with buttons that point at the
old scripts.

startup.py calls retire_old_ribbon_folders() on every Revit start. Once the
new layout is on disk it moves each leftover in RETIRED to
<extension>\\_retired_ribbon\\<timestamp>\\ -- nothing is deleted. pyRevit only
reads `*.tab` folders at the extension root, so the backup never reaches the
ribbon; a plain rename on the same drive cannot leave a half-copied folder
behind (a folder Windows keeps locked is simply retried on the next start).
Per-user settings that older versions kept inside those folders are carried
over to %APPDATA%\\T3LabAI first. pyRevit has already read the leftovers for
the session that is starting, so they disappear from the next Revit start on.

Runs under IronPython 2.7 (startup.py) and CPython 3 (tests):
lite_guard/test_ribbon_retire.py.
"""

import datetime
import io
import os
import shutil

EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# Folders the regroup moved away, relative to the extension folder.
RETIRED = (
    'T3Lab.tab/Standards & Settings.panel',
    'T3Lab.tab/Data.panel',
    'T3Lab.tab/Modeling & Datum.panel',
    'T3Lab.tab/Annotation & Select.panel',
    'T3Lab.tab/Support.panel/UI.stack',
    'T3Lab.tab/Support.panel/Assistant Tools.stack',
    'T3Lab.tab/Support.panel/PDF import.pushbutton',
    'T3Lab.tab/Support.panel/CheckUpdate.pushbutton',
    'T3Lab.tab/Views & Sheets.panel/SheetGen.pushbutton',
    'T3Lab.tab/Views & Sheets.panel/CropSync.pushbutton',
)

# All present = the new layout is on disk. A half-applied update leaves the
# old ribbon alone rather than taking buttons away.
NEW_LAYOUT = (
    'T3Lab.tab/Support.panel/Settings.pulldown/bundle.yaml',
    'T3Lab.tab/Standards & Families.panel/bundle.yaml',
    'T3Lab.tab/Model & Datum.panel/bundle.yaml',
    'T3Lab.tab/Annotation & Data.panel/bundle.yaml',
    'T3Lab.tab/Views & Sheets.panel/ViewTools.stack/bundle.yaml',
)

# Where the leftovers go, inside the extension folder (never a `*.tab`).
BACKUP_DIR = '_retired_ribbon'

# Per-user files older versions wrote inside a retired folder:
# (file under the retired folder, user_data_path() parts the tool reads now).
USER_FILES = (
    ('T3Lab.tab/Support.panel/UI.stack/Ribbon Names.pushbutton/dqt_ribbon_map.json',
     ('ribbon_names', 'ribbon_map.json')),
    ('T3Lab.tab/Support.panel/UI.stack/Ribbon Names.pushbutton/dqt_ribbon_originals.json',
     ('ribbon_names', 'ribbon_originals.json')),
    ('T3Lab.tab/Support.panel/UI.stack/Ribbon Names.pushbutton/dqt_ribbon_state.json',
     ('ribbon_names', 'ribbon_state.json')),
    ('T3Lab.tab/Support.panel/UI.stack/BG Theme.pushbutton/dqt_bg_config.json',
     ('bg_theme', 'bg_theme_config.json')),
    ('T3Lab.tab/Standards & Settings.panel/Standards.stack/ManaLoca.pushbutton/session.json',
     ('manaloca', 'session.json')),
)


def _abs(root, rel):
    return os.path.join(root, *rel.split('/'))


def _keep_user_file(old_file, parts, data_dir):
    """Copy a per-user file out of a retired folder, unless the tool already
    has its own copy under %APPDATA%\\T3LabAI."""
    if not os.path.isfile(old_file):
        return False
    if data_dir is None:
        from core.paths import user_data_path
        user_data_path(*parts, legacy=[old_file])
        return True
    target = os.path.join(data_dir, *parts)
    if os.path.exists(target):
        return False
    folder = os.path.dirname(target)
    if not os.path.isdir(folder):
        os.makedirs(folder)
    shutil.copyfile(old_file, target)
    return True


def _log(backup_root, lines):
    try:
        if not os.path.isdir(backup_root):
            os.makedirs(backup_root)
        with io.open(os.path.join(backup_root, 'retired.log'), 'a', encoding='utf-8') as fh:
            for line in lines:
                fh.write(u'{}\n'.format(line))
    except Exception:
        pass


def retire_old_ribbon_folders(extension_dir=None, backup_root=None, data_dir=None):
    """Move the leftover folders in RETIRED out of the extension.

    Returns the list of moved folders (relative paths); empty when there is
    nothing to do. Never raises -- it runs during Revit startup.

    data_dir is for tests: where per-user files go instead of
    %APPDATA%\\T3LabAI (core.paths.user_data_path).
    """
    root = extension_dir or EXTENSION_DIR
    moved = []
    try:
        if not all(os.path.isfile(_abs(root, rel)) for rel in NEW_LAYOUT):
            return moved
        present = [rel for rel in RETIRED if os.path.isdir(_abs(root, rel))]
        if not present:
            return moved

        backup_root = backup_root or os.path.join(root, BACKUP_DIR)
        stamp = datetime.datetime.now().strftime('%Y%m%d-%H%M%S')
        log = [u'[{}] retiring old ribbon folders from {}'.format(stamp, root)]

        for old_rel, parts in USER_FILES:
            try:
                if _keep_user_file(_abs(root, old_rel), parts, data_dir):
                    log.append(u'  kept settings: {}'.format(old_rel))
            except Exception as ex:
                log.append(u'  could not keep {}: {}'.format(old_rel, ex))

        for rel in present:
            dest = os.path.join(backup_root, stamp, *rel.split('/'))
            try:
                parent = os.path.dirname(dest)
                if not os.path.isdir(parent):
                    os.makedirs(parent)
                os.rename(_abs(root, rel), dest)
                moved.append(rel)
                log.append(u'  moved: {}'.format(rel))
            except Exception as ex:
                log.append(u'  could not move {}: {}'.format(rel, ex))

        _log(backup_root, log)
    except Exception:
        pass
    return moved


def notify_retired(moved):
    """One toast after a cleanup: the leftovers are still on this session's
    ribbon, so tell the user a restart finishes the update."""
    if not moved:
        return
    try:
        from pyrevit import forms
        forms.toast(
            "T3Lab removed {} ribbon folder(s) left over from the previous "
            "version (backup: the _retired_ribbon folder of the extension). "
            "Restart Revit once so the ribbon shows only the new layout.".format(len(moved)),
            title="T3Lab ribbon updated",
            appid="T3Lab Lite")
    except Exception:
        pass
