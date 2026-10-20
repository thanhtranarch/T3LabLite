# -*- coding: utf-8 -*-
"""
T3Lab Assistant — Dockable Pane, phần CPython.

startup.py (IronPython) đăng ký pane lúc Revit khởi động và cất một `Border`
rỗng (host) vào AppDomain — xem lib/assistant_pane.py. Module này là nửa còn
lại, chạy từ nút T3Lab Assistant (`#! python3`):

  * show_docked(): dựng T3LabAssistantWindow MỘT lần, tách Content của nó gắn
    vào host, rồi hiện / ẩn pane. Bấm nút lần sau chỉ hiện / ẩn — hội thoại
    vẫn còn nguyên.
  * Bơm GIL: pyRevit 6.5.5 không nhả GIL sau PythonEngine.Initialize(), nên
    luồng chính của Revit giữ GIL suốt lúc rảnh và mọi thread nền CPython đứng
    im (CLAUDE.md luật 8). Cửa sổ modal không bị vì ShowDialog là một lời gọi
    .NET (pythonnet nhả GIL trong đó); pane thì modeless. Một DispatcherTimer
    20ms gọi Thread.Sleep(0) — lời gọi .NET nên pythonnet nhả GIL, thread nền
    đang chờ (stream LLM, chờ ExternalEvent) chạy tiếp. Không thread nào chờ
    thì mỗi tick chỉ tốn vài micro giây.
  * Engine CPython tắt (pyRevit Reload chưa vá, hoặc đóng Revit): tháo nội dung
    ra, trả dòng hướng dẫn về, dừng mọi timer — để Revit không gọi vào một
    interpreter đã chết.

Removed 2026-07-27: AssistantPaneController (~350 lines) and its
Tools/AssistantPane.xaml (~1370 lines) — dead code, recoverable from git.
Removed 2026-10-07: the CPython AssistantPaneProvider. Registration moved to
startup.py, the only code that runs while Revit still accepts it.
"""

from __future__ import unicode_literals

import os
import sys

import clr
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')
clr.AddReference('RevitAPIUI')

from System import Guid
from System.Threading import Thread as _NetThread

# ─── Path bootstrap ────────────────────────────────────────────────────────────
_GUI_DIR  = os.path.dirname(__file__)                         # lib/GUI
_LIB_DIR  = os.path.dirname(_GUI_DIR)                        # lib
_EXT_DIR  = os.path.dirname(_LIB_DIR)                        # T3Lab.extension
for _p in (_LIB_DIR, _EXT_DIR):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import assistant_pane as _pane

# ─── Shared pane GUID (registered by startup.py via lib/assistant_pane.py) ─────
ASSISTANT_PANE_GUID = Guid(_pane.PANE_GUID)

# ─── Narrowest the pane content will lay itself out at (DIP) ───────────────────
# Below this Revit simply clips the right edge. It used to be 380, which is
# wider than a typical dock: a pane 400px wide at 125% display scaling is only
# 320 DIP, so the greeting, the composer hint, the project/mode row and the
# copyright were all cut off on the right. The layout now adapts down to this
# floor (T3LabAssistantWindow._apply_narrow_layout / _apply_compact_layout).
PANE_MIN_WIDTH = 240

# ─── GIL pump cadence ──────────────────────────────────────────────────────────
# 20ms = at most 20ms extra latency per streamed chunk; the tick costs a few
# microseconds when no background thread is waiting for the GIL.
GIL_PUMP_MS = 20

# The mounted window, the pump timer and what the host showed before it live in
# AppDomain data, not module globals: a pyRevit reload can re-import lib/ while
# the engine (and the mounted content) stays alive, and fresh globals would
# then mount a second assistant and start a second pump. unmount() clears them
# when the engine shuts down, so a key that is set always belongs to the
# running interpreter.
_KEY_WINDOW      = 'T3Lab.AssistantPane.Window'
_KEY_PUMP        = 'T3Lab.AssistantPane.Pump'
_KEY_PLACEHOLDER = 'T3Lab.AssistantPane.Placeholder'
_HOOKED = {'done': False}


def _get(key):
    try:
        from System import AppDomain
        return AppDomain.CurrentDomain.GetData(key)
    except Exception:
        return None


def _set(key, value):
    try:
        from System import AppDomain
        AppDomain.CurrentDomain.SetData(key, value)
    except Exception:
        pass


def _same(a, b):
    """Reference identity for .NET objects. pythonnet returns a fresh wrapper
    for a plain WPF element on every property read, so `is` is always False."""
    if a is None or b is None:
        return False
    try:
        from System import Object
        return bool(Object.ReferenceEquals(a, b))
    except Exception:
        return a is b


# ─── Debug log ─────────────────────────────────────────────────────────────────
# OFF by default. This used to append to
# ~/T3Lab_AI_Data/dockable_pane_startup.log on every Revit start, forever, and
# nothing ever read it; failures already reach the pyRevit logger below. Set
# T3LAB_PANE_DEBUG=1 only while diagnosing the pane.
_LOG_ENABLED = bool(os.environ.get("T3LAB_PANE_DEBUG"))
_LOG_MAX_BYTES = 256 * 1024
_LOG_PATH = os.path.join(os.path.expanduser("~"), "T3Lab_AI_Data",
                         "dockable_pane_startup.log")


def _log_pane(msg):
    """Append a timestamped line to the debug log. Never raises. No-op unless
    debugging is explicitly enabled."""
    if not _LOG_ENABLED:
        return
    try:
        import datetime
        import io
        _d = os.path.dirname(_LOG_PATH)
        if not os.path.isdir(_d):
            os.makedirs(_d)
        # Truncate rather than grow without bound.
        try:
            if os.path.getsize(_LOG_PATH) > _LOG_MAX_BYTES:
                os.remove(_LOG_PATH)
        except Exception:
            pass
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with io.open(_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(u"[{}] [Pane] {}\n".format(stamp, msg))
    except Exception:
        pass


# ─── GIL pump ──────────────────────────────────────────────────────────────────

def _pump_tick(sender, e):
    # Any .NET call makes pythonnet release the GIL for its duration; a thread
    # that has been waiting for it is handed the GIL before this returns.
    _NetThread.Sleep(0)


def _start_gil_pump():
    if _get(_KEY_PUMP) is not None:
        return
    try:
        from System import TimeSpan
        from System.Windows.Threading import DispatcherTimer
        timer = DispatcherTimer()
        timer.Interval = TimeSpan.FromMilliseconds(GIL_PUMP_MS)
        timer.Tick += _pump_tick
        timer.Start()
        _set(_KEY_PUMP, timer)
    except Exception as ex:
        _log_pane(u"GIL pump unavailable: {}".format(ex))


def _stop_gil_pump():
    timer = _get(_KEY_PUMP)
    _set(_KEY_PUMP, None)
    if timer is not None:
        try:
            timer.Stop()
        except Exception:
            pass


# ─── Mount / unmount ───────────────────────────────────────────────────────────

def _mounted():
    """(window, content) currently in the pane, or (None, None)."""
    win = _get(_KEY_WINDOW)
    if win is None:
        return None, None
    return win, getattr(win, '_pane_content', None)


def _stop_window(win):
    if win is None:
        return
    try:
        win._on_closing(None, None)     # stops the window's own timers
    except Exception:
        pass


def _mount(host, win):
    """Move the window's content into the pane host."""
    content = win.Content
    win.Content = None
    # Floor for the hosted pane content. Kept low on purpose — see
    # PANE_MIN_WIDTH; the layout itself adapts above it.
    try:
        content.MinWidth = PANE_MIN_WIDTH
    except Exception:
        pass
    if _get(_KEY_PLACEHOLDER) is None:
        _set(_KEY_PLACEHOLDER, host.Child)
    host.Child = content
    win._pane_content = content
    _set(_KEY_WINDOW, win)              # keeps the window (and its handlers) alive
    _start_gil_pump()
    _hook_engine_shutdown()


def unmount():
    """Detach the assistant and stop everything that calls into Python.

    Runs when the CPython engine shuts down. Never raises.
    """
    win = _get(_KEY_WINDOW)
    _set(_KEY_WINDOW, None)
    _stop_gil_pump()
    _stop_window(win)
    host = _pane.get_host()
    placeholder = _get(_KEY_PLACEHOLDER)
    if host is not None and placeholder is not None:
        try:
            host.Child = placeholder
        except Exception:
            pass


def _hook_engine_shutdown():
    """Unmount before the interpreter goes away. pythonnet runs its shutdown
    handlers while Python is still alive; atexit covers a Py_Finalize."""
    if _HOOKED['done']:
        return
    _HOOKED['done'] = True
    try:
        import atexit
        atexit.register(unmount)
    except Exception:
        pass
    try:
        from Python.Runtime import PythonEngine, ShutdownHandler
        PythonEngine.AddShutdownHandler(ShutdownHandler(unmount))
    except Exception as ex:
        _log_pane(u"shutdown handler not added: {}".format(ex))


# ─── Entry point ───────────────────────────────────────────────────────────────

def show_docked(make_window, uiapp):
    """Show the assistant in Revit's dockable pane.

    make_window: zero-argument callable returning a T3LabAssistantWindow built
    with is_docked=True. Called only when nothing is mounted yet.

    Returns (True, "shown" | "hidden" | "loaded") or (False, reason) when the
    pane cannot be used — the caller then opens the window instead.
    """
    if uiapp is None:
        return False, "no UIApplication"
    if not _pane.pane_exists():
        return False, ("the dock pane is not registered in this Revit session "
                       "(see %APPDATA%\\T3LabAI\\engine_check.log)")
    host = _pane.get_host()
    if host is None:
        return False, "the dock pane host is missing"
    try:
        from Autodesk.Revit.UI import DockablePaneId
        pane = uiapp.GetDockablePane(DockablePaneId(ASSISTANT_PANE_GUID))
    except Exception as ex:
        return False, u"GetDockablePane failed: {}".format(ex)
    if pane is None:
        return False, "GetDockablePane returned nothing"

    win, content = _mounted()
    if content is None or not _same(host.Child, content):
        _stop_window(win)               # an orphan from a replaced host
        _mount(host, make_window())
        pane.Show()
        return True, "loaded"
    _hook_engine_shutdown()             # this module may be a fresh import

    # IsVisible, not pane.IsShown(): a pane tabbed behind another one is
    # "shown" but not on screen, and the button should bring it forward.
    if content.IsVisible:
        pane.Hide()
        return True, "hidden"
    pane.Show()
    return True, "shown"
