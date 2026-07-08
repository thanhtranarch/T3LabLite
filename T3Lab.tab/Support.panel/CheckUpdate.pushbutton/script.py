# -*- coding: utf-8 -*-
"""
Check Update

Small support tool that checks whether a newer version of the T3Lab
extension is available on GitHub and updates the local copy to the
latest release.

How it works:
    1. Reads the installed version from <extension>/version.txt
    2. Downloads version.txt from the GitHub repository (main branch)
    3. If a newer version exists, updates the extension:
         - 'git pull' when the extension is a git clone and git is available
         - otherwise downloads the repository zip and copies it over
    4. Offers to reload pyRevit so the new version is active immediately

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__title__   = "Check\nUpdate"
__author__  = "Tran Tien Thanh"
__version__ = "1.0.0"

# IMPORT LIBRARIES
# ==============================================================================
import os
import time
import shutil
import tempfile
import clr

clr.AddReference('System')

from System.Net import (WebClient, ServicePointManager,
                        SecurityProtocolType, CredentialCache)
from System.Text import Encoding
from System.Diagnostics import Process, ProcessStartInfo

from pyrevit import forms, script

extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))

# DEFINE VARIABLES
# ==============================================================================
GITHUB_REPO   = "thanhtranarch/T3LabLite"
GITHUB_BRANCH = "main"
VERSION_FILE  = "version.txt"

# Several sources for the same file -- raw.githubusercontent.com rate-limits
# per IP (HTTP 429), which shared office networks hit easily. jsDelivr is a
# CDN mirror of the repository and is effectively rate-limit free.
REMOTE_VERSION_URLS = [
    "https://raw.githubusercontent.com/{repo}/{branch}/{vfile}".format(
        repo=GITHUB_REPO, branch=GITHUB_BRANCH, vfile=VERSION_FILE),
    "https://cdn.jsdelivr.net/gh/{repo}@{branch}/{vfile}".format(
        repo=GITHUB_REPO, branch=GITHUB_BRANCH, vfile=VERSION_FILE),
]
REMOTE_ZIP_URL = "https://github.com/{repo}/archive/refs/heads/{branch}.zip".format(
    repo=GITHUB_REPO, branch=GITHUB_BRANCH)

logger = script.get_logger()


# CLASS/FUNCTIONS
# ==============================================================================

# ============================================================
# VERSION HELPERS
# ============================================================
def _enable_tls12():
    """Make sure HTTPS calls work on older .NET defaults (IronPython)."""
    try:
        ServicePointManager.SecurityProtocol = (
            ServicePointManager.SecurityProtocol | SecurityProtocolType.Tls12)
    except Exception:
        pass


def _read_local_version():
    version_path = os.path.join(extension_dir, VERSION_FILE)
    try:
        with open(version_path, 'r') as vfile:
            text = vfile.read().strip()
            return text if text else "0.0.0"
    except Exception:
        # No version file yet -- treat as an old install so update is offered
        return "0.0.0"


def _clean_version_text(text):
    return (text or "").strip().lstrip(u'\ufeff').strip()


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
    code, _, _ = _run_command(
        "git", "fetch origin {}".format(GITHUB_BRANCH), cwd=extension_dir)
    if code != 0:
        return None
    code, stdout, _ = _run_command(
        "git", "show origin/{}:{}".format(GITHUB_BRANCH, VERSION_FILE),
        cwd=extension_dir)
    if code != 0:
        return None
    return _clean_version_text(stdout) or None


def _fetch_remote_version():
    # Preferred: ask git directly when the extension is a clone
    if _git_usable():
        text = _fetch_remote_version_git()
        if text:
            return text
        logger.debug("git version check failed, falling back to HTTP")

    last_error = None
    for attempt in range(2):
        if attempt:
            time.sleep(3)  # brief pause before the retry round
        for url in REMOTE_VERSION_URLS:
            client = _new_web_client()
            try:
                text = _clean_version_text(client.DownloadString(url))
                if text:
                    return text
            except Exception as ex:
                last_error = ex
                logger.debug("version check failed for %s: %s", url, ex)
            finally:
                client.Dispose()
    raise last_error or Exception("No version source was reachable.")


def _parse_version(text):
    """'1.2.3' -> (1, 2, 3); tolerates stray characters."""
    parts = []
    for token in (text or "").strip().split('.'):
        digits = ''.join(ch for ch in token if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    return tuple(parts) if parts else (0,)


# ============================================================
# UPDATE HELPERS
# ============================================================
def _run_command(exe, args, cwd=None):
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


def _git_usable():
    """True when the extension is a git clone and git.exe is available."""
    if not os.path.isdir(os.path.join(extension_dir, '.git')):
        return False
    try:
        code, _, _ = _run_command("git", "--version")
        return code == 0
    except Exception:
        return False


def _update_with_git():
    """Fast-forward the local clone to the latest remote commit."""
    code, stdout, stderr = _run_command("git", "pull --ff-only", cwd=extension_dir)
    log = (stdout + "\n" + stderr).strip()
    logger.debug("git pull output:\n%s", log)
    return code == 0, log


def _update_with_zip():
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
        dest_folder = extension_dir if rel == '.' else os.path.join(extension_dir, rel)
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


def _offer_reload(new_version):
    res = forms.alert(
        "T3Lab has been updated to version {}.\n\n"
        "Reload pyRevit now to start using the new version?".format(new_version),
        title="Update complete",
        options=["Reload pyRevit now", "Later"])
    if res == "Reload pyRevit now":
        try:
            from pyrevit.loader.sessionmgr import reload_pyrevit
            reload_pyrevit()
        except Exception as ex:
            logger.error("Automatic reload failed: %s", ex)
            forms.alert(
                "Could not reload automatically.\n\n"
                "Please click pyRevit > Reload to finish the update.",
                title="Reload required")


# ============================================================
# MAIN FLOW
# ============================================================
def _run_update(remote_text):
    if _git_usable():
        ok, log = _update_with_git()
        if ok:
            _offer_reload(remote_text)
            return
        res = forms.alert(
            "Update via git failed:\n\n{}\n\n"
            "Try downloading the latest version directly instead? "
            "This overwrites the extension files.".format(log[:800]),
            title="Update failed",
            options=["Download latest version", "Cancel"])
        if res != "Download latest version":
            return

    try:
        failed = _update_with_zip()
    except Exception as ex:
        logger.error("Zip update failed: %s", ex)
        forms.alert(
            "Could not download or apply the update.\n\n{}".format(ex),
            title="Update failed")
        return

    if failed:
        forms.alert(
            "Updated with warnings -- {} file(s) could not be replaced:\n\n{}".format(
                len(failed), "\n".join(failed[:15])),
            title="Update finished with warnings")
    _offer_reload(remote_text)


def main():
    _enable_tls12()

    local_text = _read_local_version()
    try:
        remote_text = _fetch_remote_version()
    except Exception as ex:
        logger.error("Version check failed: %s", ex)
        forms.alert(
            "Could not check the latest version online.\n"
            "The server may be busy or rate-limited -- please try again "
            "in a few minutes.\n\n{}".format(ex),
            title="Check Update",
            exitscript=True)
        return

    if _parse_version(remote_text) <= _parse_version(local_text):
        forms.alert(
            "T3Lab is up to date.\n\n"
            "Installed version:  {}\n"
            "Latest version:      {}".format(local_text, remote_text),
            title="Check Update")
        return

    res = forms.alert(
        "A new version of T3Lab is available!\n\n"
        "Installed version:  {}\n"
        "Latest version:      {}\n\n"
        "Update now?".format(local_text, remote_text),
        title="Check Update",
        options=["Update now", "Not now"])
    if res == "Update now":
        _run_update(remote_text)


# MAIN SCRIPT
# ==============================================================================
if __name__ == '__main__':
    main()
