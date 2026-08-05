# -*- coding: utf-8 -*-
"""
T3Lab Update Core
=================
Everything needed to find out whether a newer version of the extension is
published on GitHub, and to bring the local copy up to date. No UI lives
here: `Support.panel/CheckUpdate.pushbutton` drives it interactively, and
`startup.py` drives it silently once a day.

Update strategies, in order of preference:
    1. `git pull --ff-only` -- when the extension is a git clone and git.exe
       is available. Fast-forward only, so local commits or edits are never
       overwritten; the pull just fails and is reported.
    2. The repository zip, copied over the extension folder. Used when the
       extension is *not* a git clone (a plain download install).

Daily auto-update
-----------------
`startup.py` calls `start_daily_update()` on every Revit start. The first
start of a calendar day checks GitHub on a background thread and updates
silently; every later start that day is a no-op. The downloaded code becomes
active on the next Revit start -- the running session keeps the code it already
loaded, so the user is told with a toast.

A pyRevit reload picks up new *script* code, but not a new *ribbon layout*:
Revit ribbon items cannot be moved, renamed or removed once a session has
created them, so a release that reorganises panels (1.3.0 moved Feedback and
MCPControl into the Assistant Tools stack) fails its UI build on reload with
"...exists:<button>" from RibbonPanel.verifyNameExclusive. Restarting Revit is
the only fix, which is why every update message asks for a restart first.

Settings live in %APPDATA%\\T3LabAI\\mcp_paths.json:
    "auto_update"       -- false turns the daily check off (default true)
    "last_update_check" -- 'YYYY-MM-DD' stamp of the last check

Progress is logged to %APPDATA%\\T3LabAI\\update.log.
"""

from __future__ import unicode_literals

import os
import shutil
import tempfile
import threading
import time
from datetime import datetime

import clr

clr.AddReference('System')

from System.Net import (WebClient, ServicePointManager,
                        SecurityProtocolType, CredentialCache)
from System.Text import Encoding
from System.Diagnostics import Process, ProcessStartInfo


# ── Locations ───────────────────────────────────────────────────────────────────
# lib/core/updater.py -> lib/core -> lib -> T3Lab.extension/
EXTENSION_DIR = os.path.dirname(os.path.dirname(os.path.dirname(
    os.path.abspath(__file__))))

GITHUB_REPO    = "thanhtranarch/T3LabLite"
GITHUB_BRANCH  = "main"
VERSION_FILE   = "version.txt"
CHANGELOG_FILE = "CHANGELOG.md"

# Several sources for the same file -- raw.githubusercontent.com rate-limits
# per IP (HTTP 429), which shared office networks hit easily. jsDelivr is a
# CDN mirror of the repository and is effectively rate-limit free.
REMOTE_VERSION_URLS = [
    "https://raw.githubusercontent.com/{repo}/{branch}/{vfile}".format(
        repo=GITHUB_REPO, branch=GITHUB_BRANCH, vfile=VERSION_FILE),
    "https://cdn.jsdelivr.net/gh/{repo}@{branch}/{vfile}".format(
        repo=GITHUB_REPO, branch=GITHUB_BRANCH, vfile=VERSION_FILE),
]
REMOTE_CHANGELOG_URLS = [
    "https://raw.githubusercontent.com/{repo}/{branch}/{cfile}".format(
        repo=GITHUB_REPO, branch=GITHUB_BRANCH, cfile=CHANGELOG_FILE),
    "https://cdn.jsdelivr.net/gh/{repo}@{branch}/{cfile}".format(
        repo=GITHUB_REPO, branch=GITHUB_BRANCH, cfile=CHANGELOG_FILE),
]
REMOTE_ZIP_URL = "https://github.com/{repo}/archive/refs/heads/{branch}.zip".format(
    repo=GITHUB_REPO, branch=GITHUB_BRANCH)

# Settings keys in mcp_paths.json
ENABLED_KEY = 'auto_update'
STAMP_KEY   = 'last_update_check'

# Seconds to wait after Revit startup before touching the network, so the
# ribbon is fully built and the user is never waiting on us.
STARTUP_DELAY = 20.0

_LOG_MAX_BYTES = 256 * 1024


# ============================================================
# LOGGING
# ============================================================
def _log_file():
    try:
        from core import paths as _paths
        return os.path.join(_paths.settings_dir(), 'update.log')
    except Exception:
        return os.path.join(os.path.expanduser('~'), 'T3Lab_update.log')


def log(message):
    """Append a timestamped line to update.log. Never raises."""
    try:
        path = _log_file()
        if os.path.isfile(path) and os.path.getsize(path) > _LOG_MAX_BYTES:
            os.remove(path)
        with open(path, 'a') as handle:
            handle.write("[{}] {}\n".format(
                datetime.now().strftime('%Y-%m-%d %H:%M:%S'), message))
    except Exception:
        pass


# ============================================================
# VERSION HELPERS
# ============================================================
def enable_tls12():
    """Make sure HTTPS calls work on older .NET defaults (IronPython)."""
    try:
        ServicePointManager.SecurityProtocol = (
            ServicePointManager.SecurityProtocol | SecurityProtocolType.Tls12)
    except Exception:
        pass


def read_local_version():
    version_path = os.path.join(EXTENSION_DIR, VERSION_FILE)
    try:
        with open(version_path, 'r') as vfile:
            text = vfile.read().strip()
            return text if text else "0.0.0"
    except Exception:
        # No version file yet -- treat as an old install so update is offered
        return "0.0.0"


def clean_version_text(text):
    return (text or "").strip().lstrip('﻿').strip()


def parse_version(text):
    """'1.2.3' -> (1, 2, 3); tolerates stray characters."""
    parts = []
    for token in (text or "").strip().split('.'):
        digits = ''.join(ch for ch in token if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts) if parts else (0,)


def _new_web_client():
    """WebClient with headers and proxy credentials for corporate networks."""
    client = WebClient()
    client.Encoding = Encoding.UTF8
    try:
        client.Headers.Add("User-Agent", "T3Lab-CheckUpdate/1.0")
        client.Headers.Add("Cache-Control", "no-cache")
        client.UseDefaultCredentials = True
        if client.Proxy is not None:
            client.Proxy.Credentials = CredentialCache.DefaultCredentials
    except Exception:
        pass
    return client


def _fetch_remote_version_git():
    """Read the remote version through git -- immune to web rate limits."""
    code, _, _ = run_command(
        "git", "fetch origin {}".format(GITHUB_BRANCH), cwd=EXTENSION_DIR)
    if code != 0:
        return None
    code, stdout, _ = run_command(
        "git", "show origin/{}:{}".format(GITHUB_BRANCH, VERSION_FILE),
        cwd=EXTENSION_DIR)
    if code != 0:
        return None
    return clean_version_text(stdout) or None


def fetch_remote_version():
    """Latest published version string. Raises when no source is reachable."""
    # Preferred: ask git directly when the extension is a clone
    if git_usable():
        text = _fetch_remote_version_git()
        if text:
            return text
        log("git version check failed, falling back to HTTP")

    last_error = None
    for attempt in range(2):
        if attempt:
            time.sleep(3)  # brief pause before the retry round
        for url in REMOTE_VERSION_URLS:
            client = _new_web_client()
            try:
                text = clean_version_text(client.DownloadString(url))
                if text:
                    return text
            except Exception as ex:
                last_error = ex
                log("version check failed for {}: {}".format(url, ex))
            finally:
                client.Dispose()
    raise last_error or Exception("No version source was reachable.")


# ============================================================
# CHANGELOG ("WHAT'S NEW") HELPERS
# ============================================================
def fetch_remote_changelog():
    """Return the CHANGELOG.md text from the repository, or None."""
    if git_usable():
        # origin was already fetched by the version check
        code, stdout, _ = run_command(
            "git", "show origin/{}:{}".format(GITHUB_BRANCH, CHANGELOG_FILE),
            cwd=EXTENSION_DIR)
        if code == 0 and stdout.strip():
            return stdout

    for url in REMOTE_CHANGELOG_URLS:
        client = _new_web_client()
        try:
            text = client.DownloadString(url)
            if text and text.strip():
                return text
        except Exception as ex:
            log("changelog fetch failed for {}: {}".format(url, ex))
        finally:
            client.Dispose()
    return None


def extract_whats_new(changelog_text, local_version, max_lines=20):
    """Collect changelog lines for every release newer than local_version.

    Expects Keep-a-Changelog style headings: '## [x.y.z] - date'.
    Returns a list of display lines (may be empty).
    """
    local = parse_version(local_version)
    include = False
    out = []
    for raw in (changelog_text or "").replace('\r\n', '\n').split('\n'):
        line = raw.strip()
        if line.startswith('## '):
            header = line[3:].strip()
            ver_text = header.lstrip('[').split(']')[0].strip()
            # skip non-release headings such as [Unreleased]
            if not any(ch.isdigit() for ch in ver_text):
                include = False
                continue
            include = parse_version(ver_text) > local
            if include:
                if out:
                    out.append("")
                out.append("Version " + header.replace('[', '').replace(']', ''))
        elif include:
            if line.startswith('### '):
                out.append(line[4:].strip() + ":")
            elif line.startswith('- ') or line.startswith('* '):
                out.append("  • " + line[2:].strip())
            elif line and out and out[-1].startswith("  • "):
                # continuation of a wrapped bullet line
                out[-1] += " " + line
    if len(out) > max_lines:
        out = out[:max_lines] + ["  ..."]
    return [item.replace("**", "") for item in out]


def get_whats_new_text(local_version):
    """'What's new' block for the update prompt; empty string on any problem."""
    try:
        changelog = fetch_remote_changelog()
        if not changelog:
            return ""
        notes = extract_whats_new(changelog, local_version)
        if not notes:
            return ""
        return "\n\nWhat's new:\n" + "\n".join(notes)
    except Exception as ex:
        log("could not build what's-new text: {}".format(ex))
        return ""


# ============================================================
# UPDATE HELPERS
# ============================================================
def run_command(exe, args, cwd=None):
    psi = ProcessStartInfo()
    psi.FileName = exe
    psi.Arguments = args
    psi.CreateNoWindow = True
    psi.UseShellExecute = False
    psi.RedirectStandardOutput = True
    psi.RedirectStandardError = True
    if cwd:
        psi.WorkingDirectory = cwd
    try:
        psi.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0"
    except Exception:
        pass
    proc = Process.Start(psi)
    stdout = proc.StandardOutput.ReadToEnd()
    stderr = proc.StandardError.ReadToEnd()
    proc.WaitForExit()
    return proc.ExitCode, stdout, stderr


def is_git_clone():
    """True when the extension folder is a git working copy."""
    return os.path.isdir(os.path.join(EXTENSION_DIR, '.git'))


def git_usable():
    """True when the extension is a git clone and git.exe is available."""
    if not is_git_clone():
        return False
    try:
        code, _, _ = run_command("git", "--version")
        return code == 0
    except Exception:
        return False


def _restore_generated_bytecode():
    """Undo local changes to tracked .pyc files so a pull is not blocked.

    The repository tracks __pycache__ bytecode, and pyRevit's CPython engine
    rewrites those files as it imports lib/ -- so a user who never edited
    anything can still end up with a dirty working tree that `git pull` refuses
    to fast-forward over.

    Only bytecode is restored, and only when nothing else is modified: real
    edits must block the pull, never be thrown away.
    """
    code, stdout, _ = run_command("git", "diff --name-only", cwd=EXTENSION_DIR)
    if code != 0:
        return
    changed = [line.strip() for line in stdout.splitlines() if line.strip()]
    if not changed:
        return

    suffixes = ('.pyc', '.pyo')
    if not all(path.endswith(suffixes) for path in changed):
        log("working tree has local edits -- leaving them untouched")
        return

    specs = ' '.join('"*{}"'.format(sfx) for sfx in suffixes
                     if any(path.endswith(sfx) for path in changed))
    code, _, stderr = run_command(
        "git", "checkout -- {}".format(specs), cwd=EXTENSION_DIR)
    log("restored {} regenerated bytecode file(s), exit {} {}".format(
        len(changed), code, stderr.strip()))


def update_with_git():
    """Fast-forward the local clone to the latest remote commit."""
    _restore_generated_bytecode()
    code, stdout, stderr = run_command("git", "pull --ff-only", cwd=EXTENSION_DIR)
    output = (stdout + "\n" + stderr).strip()
    log("git pull exit {}:\n{}".format(code, output))
    return code == 0, output


def update_with_zip():
    """Download the repository zip and copy it over the extension folder.

    Note: files removed upstream are not deleted locally -- this is a
    copy-over update, good enough for script bundles.
    """
    work_dir = os.path.join(tempfile.gettempdir(), "t3lab_update")
    shutil.rmtree(work_dir, ignore_errors=True)
    os.makedirs(work_dir)

    zip_path = os.path.join(work_dir, "t3lab_latest.zip")
    client = _new_web_client()
    try:
        client.DownloadFile(REMOTE_ZIP_URL, zip_path)
    finally:
        client.Dispose()

    clr.AddReference('System.IO.Compression.FileSystem')
    from System.IO.Compression import ZipFile
    extract_dir = os.path.join(work_dir, "extracted")
    ZipFile.ExtractToDirectory(zip_path, extract_dir)

    # The zip contains a single top folder, e.g. 'T3LabLite-main'
    top_dirs = [d for d in os.listdir(extract_dir)
                if os.path.isdir(os.path.join(extract_dir, d))]
    if not top_dirs:
        raise Exception("Downloaded archive has an unexpected layout.")
    src_root = os.path.join(extract_dir, top_dirs[0])

    failed = []
    for folder, dirs, files in os.walk(src_root):
        dirs[:] = [d for d in dirs if d not in ('.git', '__pycache__')]
        rel = os.path.relpath(folder, src_root)
        dest_folder = EXTENSION_DIR if rel == '.' else os.path.join(EXTENSION_DIR, rel)
        if not os.path.isdir(dest_folder):
            os.makedirs(dest_folder)
        for name in files:
            try:
                shutil.copy2(os.path.join(folder, name),
                             os.path.join(dest_folder, name))
            except Exception as ex:
                failed.append("{}\\{} ({})".format(rel, name, ex))

    shutil.rmtree(work_dir, ignore_errors=True)
    return failed


# ============================================================
# DAILY AUTO-UPDATE
# ============================================================
def _today():
    return datetime.now().strftime('%Y-%m-%d')


def is_auto_update_enabled():
    """Read "auto_update" from mcp_paths.json; default on.

    Read through load_settings rather than get_setting -- get_setting treats a
    stored false as "absent" and would overwrite the user's opt-out.
    """
    try:
        from core import paths as _paths
        value = _paths.load_settings().get(ENABLED_KEY)
        if value is None:
            _paths.set_setting(ENABLED_KEY, True)  # make it discoverable/editable
            return True
        return bool(value)
    except Exception:
        return True


def checked_today():
    """True when the daily check already ran today."""
    try:
        from core import paths as _paths
        return _paths.load_settings().get(STAMP_KEY) == _today()
    except Exception:
        return False


def stamp_today():
    """Record that today's check has been attempted.

    Stamped before the network work, not after: a machine that is offline all
    day should try once, not on every single Revit start.
    """
    try:
        from core import paths as _paths
        _paths.set_setting(STAMP_KEY, _today())
    except Exception:
        pass


def auto_apply_update():
    """Apply the update without asking anything. Returns (ok, message).

    A git clone is only ever fast-forwarded. The zip copy-over is deliberately
    *not* used as a fallback there -- it would overwrite local commits or
    uncommitted edits, which is exactly what an unattended update must not do.
    """
    if is_git_clone():
        if not git_usable():
            return False, ("extension is a git clone but git.exe was not found "
                           "-- use Check Update to update manually")
        ok, output = update_with_git()
        if ok:
            return True, "updated via git pull"
        return False, "git pull --ff-only failed: {}".format(output[:400])

    failed = update_with_zip()
    if failed:
        return True, "updated via zip, {} file(s) could not be replaced: {}".format(
            len(failed), "; ".join(failed[:5]))
    return True, "updated via zip"


def run_daily_update(force=False):
    """One full check-and-update pass. Returns (status, detail).

    status is one of: 'disabled', 'skipped', 'current', 'updated', 'failed'.
    Never raises -- this runs unattended during Revit startup.
    """
    try:
        if not force:
            if not is_auto_update_enabled():
                return 'disabled', 'auto_update is off in mcp_paths.json'
            if checked_today():
                return 'skipped', 'already checked today'
            stamp_today()

        enable_tls12()
        local_text = read_local_version()

        try:
            remote_text = fetch_remote_version()
        except Exception as ex:
            log("version check failed: {}".format(ex))
            return 'failed', "could not reach GitHub: {}".format(ex)

        if parse_version(remote_text) <= parse_version(local_text):
            log("up to date (installed {}, latest {})".format(local_text, remote_text))
            return 'current', remote_text

        log("new version found: {} -> {}".format(local_text, remote_text))
        ok, detail = auto_apply_update()
        if not ok:
            log("auto-update skipped: {}".format(detail))
            return 'failed', detail

        log("auto-update finished: {} (now {})".format(detail, remote_text))
        return 'updated', remote_text

    except Exception as ex:
        log("auto-update crashed: {}".format(ex))
        return 'failed', str(ex)


def _notify_updated(new_version):
    """Tell the user the files changed and the new code needs a reload."""
    try:
        from pyrevit import forms
        forms.toast(
            "T3Lab Lite {} was downloaded. "
            "Restart Revit to use it -- a pyRevit reload cannot apply ribbon "
            "changes.".format(new_version),
            title="T3Lab update ready",
            appid="T3Lab Lite")
    except Exception as ex:
        log("could not show update toast: {}".format(ex))


def _daily_worker(delay):
    if delay:
        time.sleep(delay)
    status, detail = run_daily_update()
    if status == 'updated':
        _notify_updated(detail)


def start_daily_update(delay=STARTUP_DELAY):
    """Kick off the once-a-day update check on a background thread.

    Called from startup.py on every Revit start. Returns True when a check was
    actually started (first start of the day, auto-update on).
    """
    if not is_auto_update_enabled():
        return False
    if checked_today():
        return False

    worker = threading.Thread(target=_daily_worker, args=(delay,))
    worker.daemon = True
    worker.start()
    return True
