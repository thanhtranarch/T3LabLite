# -*- coding: utf-8 -*-
"""
Check Update

Small support tool that checks whether a newer version of the T3Lab
extension is available on GitHub and updates the local copy to the
latest release.

How it works:
    1. Reads the installed version from <extension>/version.txt
    2. Downloads version.txt from the GitHub repository (main branch)
    3. If a newer version exists, shows "What's new" from CHANGELOG.md
       and updates the extension:
         - 'git pull' when the extension is a git clone and git is available
         - otherwise downloads the repository zip and copies it over
    4. Offers to reload pyRevit so the new version is active immediately

The check/download logic lives in lib/core/updater.py, shared with the
once-a-day automatic update that startup.py runs in the background.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
Linkedin: linkedin.com/in/sunarch7899/
"""

__title__   = "Check\nUpdate"
__author__  = "Tran Tien Thanh"
__version__ = "1.1.0"

# IMPORT LIBRARIES
# ==============================================================================
import os
import sys

# Path setup -- script.py lives 3 levels below T3Lab.extension/
SCRIPT_DIR = os.path.dirname(__file__)
EXT_DIR = os.path.dirname(os.path.dirname(os.path.dirname(SCRIPT_DIR)))
LIB_DIR = os.path.join(EXT_DIR, 'lib')
if LIB_DIR not in sys.path:
    sys.path.insert(0, LIB_DIR)

from pyrevit import forms, script

from core import updater

logger = script.get_logger()


# CLASS/FUNCTIONS
# ==============================================================================
def _offer_reload(new_version):
    # Revit never removes a ribbon item once created, so a reload can only
    # update buttons that kept their bundle name. A release that moves or
    # renames buttons is fully applied on the next Revit start.
    res = forms.alert(
        "T3Lab has been updated to version {}.\n\n"
        "Reload pyRevit now to start using the new version?\n"
        "If the T3Lab ribbon looks incomplete afterwards, restart Revit "
        "once -- ribbon layout changes only apply on a fresh start.".format(
            new_version),
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


def _run_update(remote_text):
    if updater.git_usable():
        ok, log = updater.update_with_git()
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
        failed = updater.update_with_zip()
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


# MAIN FLOW
# ==============================================================================
def main():
    updater.enable_tls12()

    local_text = updater.read_local_version()
    try:
        remote_text = updater.fetch_remote_version()
    except Exception as ex:
        logger.error("Version check failed: %s", ex)
        forms.alert(
            "Could not check the latest version online.\n"
            "The server may be busy or rate-limited -- please try again "
            "in a few minutes.\n\n{}".format(ex),
            title="Check Update",
            exitscript=True)
        return

    # A manual check counts as today's check, so the automatic one on the
    # next Revit start does not repeat the same work.
    updater.stamp_today()

    if updater.parse_version(remote_text) <= updater.parse_version(local_text):
        forms.alert(
            "T3Lab is up to date.\n\n"
            "Installed version:  {}\n"
            "Latest version:      {}".format(local_text, remote_text),
            title="Check Update")
        return

    res = forms.alert(
        u"A new version of T3Lab is available!\n\n"
        u"Installed version:  {}\n"
        u"Latest version:      {}{}\n\n"
        u"Update now?".format(local_text, remote_text,
                              updater.get_whats_new_text(local_text)),
        title="Check Update",
        options=["Update now", "Not now"])
    if res == "Update now":
        _run_update(remote_text)


# MAIN SCRIPT
# ==============================================================================
if __name__ == '__main__':
    main()
