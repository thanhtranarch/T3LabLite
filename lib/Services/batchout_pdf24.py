# -*- coding: utf-8 -*-
"""BatchOut PDF engine "PDF24 printer": find the PDF24 Creator printer, point
its auto-save at the export folder while BatchOut prints, put the user's own
PDF24 settings back afterwards.

PDF24 reads HKLM\\SOFTWARE\\PDF24\\Services\\PDF (the default "PDF24" printer),
then every value under the same HKCU path replaces it, unless HKLM has
NoUserOverride=1 (PDF24 Creator manual 11, §4.17 + §6.5). We only write HKCU.
The values we replace are saved to %APPDATA%\\T3LabAI\\pdf24_restore.json first,
so a Revit crash mid-export is undone on the next run.

Pure Python, engine-agnostic (IronPython 2.7 + CPython 3); the registry and
clock are injectable. Test: lite_guard/test_batchout_pdf24.py. NOT verified with
PDF24 + Revit yet.
"""
from __future__ import unicode_literals

import io
import json
import os
import time

SERVICE_KEY = 'SOFTWARE\\PDF24\\Services\\PDF'
DEVICES_KEY = 'Software\\Microsoft\\Windows NT\\CurrentVersion\\Devices'
HKCU, HKLM = 'HKCU', 'HKLM'
REG_SZ, REG_DWORD = 1, 4
RESTORE_FILE = 'pdf24_restore.json'
INSTALL_URL = 'https://tools.pdf24.org/en/creator'

# Auto-save for one BatchOut run: no dialog, no Explorer window, no command,
# overwrite (BatchOut already reserved a unique name), never the open Creator.
AUTO_SAVE_VALUES = (
    ('Handler', 'autoSave', REG_SZ),
    ('AutoSaveOverwriteFile', 1, REG_DWORD),
    ('AutoSaveUseFileChooser', 0, REG_DWORD),
    ('AutoSaveOpenDir', 0, REG_DWORD),
    ('AutoSaveShowProgress', 0, REG_DWORD),
    ('AutoSaveUseFileCmd', 0, REG_DWORD),
    ('LoadInCreatorIfOpen', 0, REG_DWORD),
)
TOUCHED = tuple(v[0] for v in AUTO_SAVE_VALUES) + ('AutoSaveDir', 'AutoSaveFilename')

DOCTOOL = 'pdf24-DocTool.exe'
GS_EXE = 'gswin64c.exe'
PRN_SUFFIX = '.t3lab.prn'       # Revit's print-to-file output for a "virtual" PDF24


# ─── Pure helpers ──────────────────────────────────────────────────────────────

def pick_printer(names):
    """The PDF24 PDF printer among installed printer names, or None.
    "PDF24" itself first; PDF24 Fax / other services are not PDF printers."""
    names = [n for n in (names or []) if n]
    for n in names:
        if n.strip().lower() == 'pdf24':
            return n
    for n in names:
        low = n.lower()
        if 'pdf24' in low and 'fax' not in low:
            return n
    return None


def auto_save_filename(name):
    """PDF24 runs AutoSaveFilename through strftime: '%' must be doubled.
    No extension: PDF24 adds '.pdf'."""
    name = (name or '').strip()
    if name.lower().endswith('.pdf'):
        name = name[:-4]
    return name.replace('%', '%%')


def _norm_paper(name):
    s = (name or '').upper().replace('_', ' ').replace('-', ' ').strip()
    for prefix in ('ISO ', 'ISO'):
        if s.startswith(prefix) and len(s) > len(prefix) and s[len(prefix)] in 'ABC':
            s = s[len(prefix):]
    return ' '.join(s.split())


def match_paper(wanted, available):
    """The printer paper name for BatchOut's size ("A1", "ISO_A1", "ANSI_D"),
    or None. Exact first, then a printer name that starts with it ("A1 (594 x
    841 mm)"); never "A1" -> "A10"."""
    want = _norm_paper(wanted)
    if not want or want == 'USE SHEET SIZE':
        return None
    normed = [(a, _norm_paper(a)) for a in (available or []) if a]
    for a, n in normed:
        if n == want:
            return a
    for a, n in normed:
        if n.startswith(want) and not n[len(want):len(want) + 1].isalnum():
            return a
    return None


def doctool_paths(env=None):
    env = os.environ if env is None else env
    out = []
    for var in ('ProgramFiles', 'ProgramW6432', 'ProgramFiles(x86)'):
        base = env.get(var)
        if base:
            path = os.path.join(base, 'PDF24', DOCTOOL)
            if path not in out:
                out.append(path)
    return out


def find_doctool(env=None, exists=os.path.isfile):
    for path in doctool_paths(env):
        if exists(path):
            return path
    return None


def join_command(doctool, files, out_dir, out_name):
    """argv for pdf24-DocTool: join `files` in order, no profile (keeps the
    printed quality), into out_dir/out_name.pdf (manual §11.1.5)."""
    name = out_name if out_name.lower().endswith('.pdf') else out_name + '.pdf'
    return ([doctool, '-join', '-profile', 'special/justAssemble',
             '-outputDir', out_dir, '-outputFile', name] + list(files))


def sniff(path):
    """'pdf' | 'ps' | 'other' for a printed file, None when unreadable. A PJL
    header (UEL + @PJL lines) may come before %!PS."""
    try:
        with io.open(path, 'rb') as f:
            head = f.read(4096)
    except (IOError, OSError):
        return None
    if b'%PDF-' in head[:1024]:
        return 'pdf'
    if b'%!PS' in head:
        return 'ps'
    return 'other'


def ghostscript_paths(env=None):
    """Ghostscript PDF24 ships (PDF24\\gs\\bin), for PostScript -> PDF."""
    env = os.environ if env is None else env
    out = []
    for var in ('ProgramFiles', 'ProgramW6432', 'ProgramFiles(x86)'):
        base = env.get(var)
        if base:
            path = os.path.join(base, 'PDF24', 'gs', 'bin', GS_EXE)
            if path not in out:
                out.append(path)
    return out


def find_ghostscript(env=None, exists=os.path.isfile):
    for path in ghostscript_paths(env):
        if exists(path):
            return path
    return None


def ps_to_pdf_command(gs, src, dst):
    """argv: PostScript (Revit printed to file) -> PDF, print quality."""
    return [gs, '-dBATCH', '-dNOPAUSE', '-dSAFER', '-q', '-sDEVICE=pdfwrite',
            '-dPDFSETTINGS=/prepress', '-sOutputFile=' + dst, src]


def prn_path(folder, name):
    return os.path.join(folder, name + PRN_SUFFIX)


def wait_for_pdf(expected, prev_mtime, started_at, timeout, taken=(),
                 clock=None, sleep=None, poll=0.25, extra=()):
    """Path of the PDF PDF24 wrote for one job, or None after `timeout` s.

    PDF24 saves after SubmitPrint returns, so poll. Ready = nonempty, newer
    than prev_mtime, same size on two polls, openable for writing (PDF24 no
    longer holds it). A different new .pdf in the folder (PDF24 kept its own
    filename pattern) counts too, unless already in `taken`; so do the
    `extra` paths (Revit's print-to-file output). Check the result with sniff().
    """
    clock = clock or getattr(time, 'monotonic', time.time)
    sleep = sleep or time.sleep
    folder = os.path.dirname(expected)
    taken = set(p.lower() for p in taken)
    deadline = clock() + max(0.0, timeout)
    last = {}
    while True:
        found = _candidates(expected, folder, started_at, prev_mtime, taken)
        found += [p for p in extra if p not in found and _fresh(p, started_at - 1.0)]
        for path in found:
            size = _size(path)
            if size and last.get(path) == size and _unlocked(path):
                return path
            last[path] = size
        if clock() >= deadline:
            return None
        sleep(poll)


def _candidates(expected, folder, started_at, prev_mtime, taken):
    out = []
    if _fresh(expected, prev_mtime):
        out.append(expected)
    try:
        names = os.listdir(folder)
    except OSError:
        names = []
    for name in sorted(names):
        path = os.path.join(folder, name)
        if (path.lower() == expected.lower() or not name.lower().endswith('.pdf')
                or path.lower() in taken):
            continue
        if _fresh(path, started_at - 1.0):
            out.append(path)
    return out


def _fresh(path, after):
    try:
        return os.path.isfile(path) and (after is None or os.path.getmtime(path) > after)
    except OSError:
        return False


def _size(path):
    try:
        return os.path.getsize(path)
    except OSError:
        return 0


def _unlocked(path):
    try:
        with io.open(path, 'r+b'):
            return True
    except (IOError, OSError):
        return False


# ─── Registry ─────────────────────────────────────────────────────────────────

class WinRegistry(object):
    """HKCU/HKLM through winreg (CPython) or _winreg (IronPython)."""

    def __init__(self):
        try:
            import winreg as reg
        except ImportError:
            import _winreg as reg
        self._reg = reg
        self._hives = {HKCU: reg.HKEY_CURRENT_USER, HKLM: reg.HKEY_LOCAL_MACHINE}

    def read(self, hive, key, name):
        """(value, kind) or None."""
        try:
            with self._reg.OpenKey(self._hives[hive], key) as k:
                value, kind = self._reg.QueryValueEx(k, name)
                return value, kind
        except OSError:
            return None

    def names(self, hive, key):
        out = []
        try:
            with self._reg.OpenKey(self._hives[hive], key) as k:
                i = 0
                while True:
                    try:
                        out.append(self._reg.EnumValue(k, i)[0])
                    except OSError:
                        break
                    i += 1
        except OSError:
            pass
        return out

    def write(self, hive, key, name, value, kind):
        with self._reg.CreateKey(self._hives[hive], key) as k:
            self._reg.SetValueEx(k, name, 0, kind, value)

    def delete(self, hive, key, name):
        try:
            with self._reg.OpenKey(self._hives[hive], key, 0, self._reg.KEY_SET_VALUE) as k:
                self._reg.DeleteValue(k, name)
        except OSError:
            pass


def default_registry():
    try:
        return WinRegistry()
    except Exception:
        return None


def installed_printers(registry=None):
    """Printer names for the current user (HKCU ...\\Devices), then .NET."""
    registry = registry or default_registry()
    names = registry.names(HKCU, DEVICES_KEY) if registry is not None else []
    if names:
        return names
    try:
        import clr
        clr.AddReference('System.Drawing')
        from System.Drawing.Printing import PrinterSettings
        return [str(n) for n in PrinterSettings.InstalledPrinters]
    except Exception:
        return []


def find_printer(registry=None):
    return pick_printer(installed_printers(registry))


# ─── Auto-save session ─────────────────────────────────────────────────────────

def _state_dir():
    base = os.environ.get('APPDATA', '') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI')


class AutoSave(object):
    """Point the PDF24 printer's auto-save at a folder for one run.

        session = AutoSave(); session.start(folder)
        session.set_filename(name)     # before each SubmitPrint
        session.restore()              # always, in finally
    """

    def __init__(self, registry=None, state_dir=None):
        self.registry = registry if registry is not None else default_registry()
        self.state_path = os.path.join(state_dir or _state_dir(), RESTORE_FILE)
        self.active = False

    def blocked_reason(self):
        """Why auto-save cannot be set for this user, or ''."""
        if self.registry is None:
            return 'the Windows registry is not reachable from this Python engine'
        flag = self.registry.read(HKLM, SERVICE_KEY, 'NoUserOverride')
        if flag is not None and str(flag[0]).strip() in ('1', 'True'):
            return ('PDF24 is locked by an administrator (NoUserOverride=1), so '
                    'BatchOut cannot point its auto-save at the export folder')
        return ''

    def recover(self):
        """Undo a run that never reached restore() (Revit crashed). True if it did."""
        saved = self._load()
        if saved is None:
            return False
        self._put_back(saved)
        self._forget()
        return True

    def start(self, folder):
        self.recover()
        saved = {}
        for name in TOUCHED:
            found = self.registry.read(HKCU, SERVICE_KEY, name)
            saved[name] = list(found) if found is not None else None
        self._save(saved)
        self.active = True
        for name, value, kind in AUTO_SAVE_VALUES:
            self.registry.write(HKCU, SERVICE_KEY, name, value, kind)
        self.registry.write(HKCU, SERVICE_KEY, 'AutoSaveDir', folder, REG_SZ)

    def set_filename(self, name):
        self.registry.write(HKCU, SERVICE_KEY, 'AutoSaveFilename',
                            auto_save_filename(name), REG_SZ)

    def restore(self):
        """The user's own PDF24 values back. Never raises."""
        if not self.active:
            return
        self.active = False
        try:
            saved = self._load()
            if saved is not None:
                self._put_back(saved)
            self._forget()
        except Exception:
            pass

    def _put_back(self, saved):
        for name in TOUCHED:
            if name not in saved:
                continue
            prev = saved[name]
            if prev is None:
                self.registry.delete(HKCU, SERVICE_KEY, name)
            else:
                self.registry.write(HKCU, SERVICE_KEY, name, prev[0], int(prev[1]))

    def _save(self, saved):
        folder = os.path.dirname(self.state_path)
        if not os.path.isdir(folder):
            os.makedirs(folder)
        text = json.dumps(saved, indent=2, ensure_ascii=True)
        if isinstance(text, bytes):
            text = text.decode('ascii')
        with io.open(self.state_path, 'w', encoding='utf-8') as f:
            f.write(text)

    def _load(self):
        try:
            with io.open(self.state_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return data if isinstance(data, dict) else None
        except Exception:
            return None

    def _forget(self):
        try:
            os.remove(self.state_path)
        except OSError:
            pass
