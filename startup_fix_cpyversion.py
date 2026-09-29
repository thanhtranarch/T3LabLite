# -*- coding: utf-8 -*-
"""
startup_fix_cpyversion.py  --  KHÔNG XÓA FILE NÀY / DO NOT DELETE
=================================================================
Called from the very top of startup.py (same folder, on purpose: easy to see).

Fixes, for every '#! python3' tool on the ribbon:

    Command Failure for External Command
    Revit encountered a The input string '3.12.3' was not in a correct format.

Vì sao (pyRevit bug #3284, upstream đã sửa từ pyRevit 6.5.0):

  1. Bộ nạp session C# của pyRevit (EnvDictionarySeeder) ghi cứng
     PYREVIT_CPYVERSION = "3.12.3" vào env dictionary của AppDomain.
  2. Mỗi lần bấm một tool CPython, CPythonEngine.GetPythonDll() chạy
     int.Parse(PYREVIT_CPYVERSION) -> "3.12.3" không phải số nguyên ->
     FormatException -> Revit hiện hộp thoại "Command Failure".
  3. Giá trị đúng là dạng số nguyên của engine: "3123" (thư mục CPY3123).

Cách sửa: bộ nạp seed env dictionary TRƯỚC khi chạy startup.py của các
extension, và runtime đọc lại dictionary đó ở MỖI lần bấm tool. Nên startup.py
(chạy bằng IronPython, không bị lỗi này) chỉ cần đổi "3.12.3" -> "3123" là mọi
tool CPython chạy lại bình thường. Chạy ở mỗi lần mở Revit và mỗi lần pyRevit
Reload (bộ nạp seed lại mỗi lần). Giá trị đã là số nguyên (pyRevit 6.5.0+) thì
không làm gì -> an toàn để giữ mãi.

Kết quả ghi vào ~/T3Lab_AI_Data/bootstrap_status.log khi có sửa hoặc lỗi.

Runs under IronPython 2.7 / 3.4 (startup.py) and CPython 3 (tests):
no f-strings, no Python-3-only syntax.
"""
import os

# Must match pyRevit: DomainStorageKeys.EnvVarsDictKey / EnvDictionaryKeys.CPYVersion
# (dev/pyRevitLabs.PyRevit.Runtime/EnvVariables.cs).
ENV_DICT_KEY = "PYREVITEnvVarsDict"
CPYVERSION_KEY = "PYREVIT_CPYVERSION"

LAST_RESULT = "not run"


def _digits(text):
    return "".join(c for c in (text or "") if c.isdigit())


def _pyrevit_home():
    """Root of the running pyRevit clone, or None."""
    try:
        from pyrevit import HOME_DIR
        if HOME_DIR:
            return HOME_DIR
    except Exception:
        pass
    try:
        import pyrevit
        # <clone>/pyrevitlib/pyrevit/__init__.py -> <clone>
        return os.path.dirname(os.path.dirname(os.path.dirname(
            os.path.abspath(pyrevit.__file__))))
    except Exception:
        return None


def installed_engine_versions(home=None):
    """Integer versions of the CPython engines in the running pyRevit clone.

    ``<clone>/bin/cengines/CPY3123`` -> 3123. These are the numbers
    ``clone.GetCPythonEngine()`` accepts, so the fixed value must be one of them.
    """
    home = home or _pyrevit_home()
    found = []
    if not home:
        return found
    cengines = os.path.join(home, "bin", "cengines")
    try:
        names = os.listdir(cengines)
    except Exception:
        return found
    for name in names:
        number = name[3:]
        if (name.upper().startswith("CPY") and number.isdigit()
                and os.path.isdir(os.path.join(cengines, name))):
            found.append(int(number))
    return sorted(found)


def resolve(current, available):
    """Integer version string the pyRevit runtime can int.Parse(), or None.

    "3.12.3" -> "3123" when that engine is installed (or nothing could be
    scanned); otherwise the newest installed engine.
    """
    digits = _digits(current)
    if digits and int(digits) > 0 and (not available or int(digits) in available):
        return str(int(digits))
    if available:
        return str(max(available))
    return None


def _get(data, key):
    try:
        return data.get(key)        # IronPython: PythonDictionary is a dict
    except Exception:
        pass
    try:
        return data[key]            # CPython (pythonnet): .NET indexer
    except Exception:
        return None


def _log(message):
    try:
        import datetime
        path = os.path.join(os.path.expanduser("~"), "T3Lab_AI_Data",
                            "bootstrap_status.log")
        if not os.path.isdir(os.path.dirname(path)):
            os.makedirs(os.path.dirname(path))
        with open(path, "a") as fh:
            fh.write("[%s] cpyversion fix: %s\n" % (
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), message))
    except Exception:
        pass


def apply(data=None, available=None):
    """Make PYREVIT_CPYVERSION an integer string. Returns a status; never raises.

    `data` / `available` are for tests; startup.py calls apply() with no args.
    """
    global LAST_RESULT
    try:
        if data is None:
            from System import AppDomain
            data = AppDomain.CurrentDomain.GetData(ENV_DICT_KEY)
        if data is None:
            LAST_RESULT = "skipped: pyRevit env dictionary not found"
            return LAST_RESULT

        current = _get(data, CPYVERSION_KEY)
        if current is None:
            LAST_RESULT = "skipped: %s not set" % CPYVERSION_KEY
            return LAST_RESULT
        current = str(current)
        if current.isdigit():
            LAST_RESULT = "ok: %s" % current     # pyRevit 6.5.0+: nothing to do
            return LAST_RESULT

        if available is None:
            available = installed_engine_versions()
        fixed = resolve(current, available)
        if not fixed:
            LAST_RESULT = "failed: no CPython engine matches '%s'" % current
            _log(LAST_RESULT)
            return LAST_RESULT

        data[CPYVERSION_KEY] = fixed
        LAST_RESULT = "fixed: '%s' -> '%s'" % (current, fixed)
        _log(LAST_RESULT)
    except Exception as exc:
        LAST_RESULT = "failed: %s: %s" % (type(exc).__name__, exc)
        _log(LAST_RESULT)
    return LAST_RESULT
