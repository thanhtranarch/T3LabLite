# -*- coding: utf-8 -*-
"""T3Lab Assistant dockable pane — phần IronPython (đăng ký lúc Revit khởi động).

Revit chỉ nhận RegisterDockablePane trong lúc khởi động. startup.py là file duy
nhất của T3Lab chạy lúc đó, và nó chạy IronPython (có chủ ý — xem docstring của
startup.py). Còn Assistant là CPython (`#! python3`), nên pane được dựng thành
hai nửa nối với nhau bằng một phần tử WPF dùng chung:

  1. startup.py -> register(): tạo một `Border` rỗng (host) với dòng hướng dẫn,
     đăng ký pane trỏ vào host, rồi cất host vào AppDomain dưới HOST_KEY.
  2. Nút T3Lab Assistant (CPython) -> GUI/AssistantPaneControl.show_docked():
     lấy host từ AppDomain, gắn nội dung cửa sổ Assistant vào, rồi Show().

AppDomain là chỗ duy nhất hai engine Python cùng nhìn thấy: host là một object
.NET thường nên engine nào cũng dùng được.

Module này phải chạy được trên IronPython 2.7 / 3.4 (startup.py) và CPython
(test, hằng số cho AssistantPaneControl): không f-string, không cú pháp chỉ có
ở Python 3. Class IDockablePaneProvider chỉ được dựng bên trong register(), nên
import module từ CPython không đăng ký kiểu .NET nào.

Mọi hàm ở đây không bao giờ ném lỗi ra ngoài: startup.py không được làm Revit
khởi động thất bại.
"""
from __future__ import unicode_literals

import os
import sys

PANE_GUID = "7F3A9B2E-C4D1-4E8F-A6B5-1234567890AB"   # = core/server.py show_assistant_pane
PANE_TITLE = "T3Lab Assistant"
HOST_KEY = "T3Lab.AssistantPaneHost"
PLACEHOLDER_TEXT = ("T3Lab Assistant is not loaded yet. Click T3Lab Assistant "
                    "on the T3Lab tab to load it here.")

# register() states
REGISTERED = "registered"
EXISTS = "exists"
FAILED = "failed"


def _err(exc):
    try:
        return "%s" % (getattr(exc, "Message", None) or exc)
    except Exception:
        return "<unprintable error>"


def pane_id():
    from System import Guid
    from Autodesk.Revit.UI import DockablePaneId
    return DockablePaneId(Guid(PANE_GUID))


def pane_exists():
    """True when the pane is registered in this Revit session. Never raises."""
    try:
        from Autodesk.Revit.UI import DockablePane
        return bool(DockablePane.PaneExists(pane_id()))
    except Exception:
        return False


def get_host():
    """The host Border registered at startup, or None. Never raises."""
    try:
        from System import AppDomain
        return AppDomain.CurrentDomain.GetData(HOST_KEY)
    except Exception:
        return None


def _lib_dir():
    return os.path.dirname(os.path.abspath(__file__))


def _theme_host(host, text):
    """Paint host + placeholder with the Assistant's own Revit palette.

    GUI/RevitTheme.py writes T3Theme* brushes for the current Revit theme
    (light / dark). The assistant content binds the same keys, so the pane
    keeps one look before and after it loads. Without it the placeholder
    falls back to the system window colours.
    """
    try:
        if _lib_dir() not in sys.path:
            sys.path.insert(0, _lib_dir())
        from GUI import RevitTheme
        RevitTheme.apply(host)
        from System.Windows.Controls import Border, TextBlock
        host.SetResourceReference(Border.BackgroundProperty, "T3ThemeChatBg")
        text.SetResourceReference(TextBlock.ForegroundProperty, "T3ThemeMuted")
        return True
    except Exception:
        pass
    try:
        from System.Windows import SystemColors
        host.Background = SystemColors.WindowBrush
        text.Foreground = SystemColors.GrayTextBrush
    except Exception:
        pass
    return False


def build_host():
    """The pane's root element: a Border showing the 'not loaded' hint.

    AssistantPaneControl swaps Child for the assistant content and puts this
    hint back when the CPython engine shuts down.
    """
    from System.Windows import HorizontalAlignment, VerticalAlignment, Thickness, TextWrapping, TextAlignment
    from System.Windows.Controls import Border, TextBlock
    text = TextBlock()
    text.Text = PLACEHOLDER_TEXT
    text.TextWrapping = TextWrapping.Wrap
    text.TextAlignment = TextAlignment.Center
    text.HorizontalAlignment = HorizontalAlignment.Center
    text.VerticalAlignment = VerticalAlignment.Center
    text.Margin = Thickness(16)
    host = Border()
    host.Child = text
    _theme_host(host, text)
    return host


def apply_initial_state(data):
    """Dock on the right as its own panel, the way Properties sits on the left.

    Revit honours InitialState only the first time the pane appears on a
    machine; afterwards it remembers wherever the user docked it.
    """
    try:
        from Autodesk.Revit.UI import DockablePaneState, DockPosition
        state = DockablePaneState()
        state.DockPosition = DockPosition.Right
        data.InitialState = state
        return True
    except Exception:
        return False


def _make_provider(host):
    """IDockablePaneProvider handing Revit the shared host. IronPython only."""
    import uuid
    from Autodesk.Revit.UI import (IDockablePaneProvider, EditorInteraction,
                                   EditorInteractionType)

    class _AssistantPaneProvider(IDockablePaneProvider):
        # Ignored by IronPython; keeps the type unique if this ever runs on
        # pythonnet (rule S15).
        __namespace__ = "T3Lab.AssistantPaneProvider_" + uuid.uuid4().hex[:8]

        def SetupDockablePane(self, data):
            data.FrameworkElement = host
            # A fresh install shows nothing until the button is clicked; from
            # then on Revit remembers whether the user left the pane open.
            try:
                data.VisibleByDefault = False
            except Exception:
                pass
            apply_initial_state(data)
            # Clicking into the pane must not cancel a command in progress.
            try:
                data.EditorInteraction = EditorInteraction(EditorInteractionType.KeepAlive)
            except Exception:
                pass

    return _AssistantPaneProvider()


def _app_candidates(revit_handle):
    """Objects that may own RegisterDockablePane, most likely first."""
    found = []
    if revit_handle is not None:
        found.append(("__revit__", revit_handle))
    try:
        from pyrevit import HOST_APP
        uiapp = getattr(HOST_APP, "uiapp", None)
        if uiapp is not None:
            found.append(("HOST_APP.uiapp", uiapp))
    except Exception:
        pass
    # On some pyRevit builds __revit__ is the DB.Application; UIApplication
    # has a public constructor taking it.
    if revit_handle is not None and not hasattr(revit_handle, "RegisterDockablePane"):
        try:
            from Autodesk.Revit.UI import UIApplication
            found.append(("UIApplication(__revit__)", UIApplication(revit_handle)))
        except Exception:
            pass
    loader_app = _loader_uicontrolled_app()
    if loader_app is not None:
        found.append(("PyRevitLoaderApplication", loader_app))
    return [(name, app) for name, app in found if hasattr(app, "RegisterDockablePane")]


def _loader_uicontrolled_app():
    """The UIControlledApplication pyRevit's loader kept from OnStartup, or None."""
    try:
        from System import AppDomain
        from System.Reflection import BindingFlags
        flags = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static
        for asm in AppDomain.CurrentDomain.GetAssemblies():
            try:
                if not asm.GetName().Name.startswith("pyRevitLoader"):
                    continue
                kind = asm.GetType("PyRevitLoader.PyRevitLoaderApplication")
                field = kind.GetField("_uiControlledApplication", flags) if kind else None
                app = field.GetValue(None) if field else None
                if app is not None:
                    return app
            except Exception:
                continue
    except Exception:
        pass
    return None


def register(revit_handle):
    """Register the pane. Returns (state, detail); never raises.

    Must run while Revit starts (startup.py). On a pyRevit reload the pane
    already exists and this returns EXISTS without touching it.
    """
    try:
        import clr
        clr.AddReference("RevitAPIUI")
        clr.AddReference("PresentationFramework")
        clr.AddReference("PresentationCore")
        clr.AddReference("WindowsBase")
    except Exception as exc:
        return FAILED, "Revit UI API unavailable: " + _err(exc)
    if pane_exists():
        return EXISTS, ""
    try:
        host = build_host()
        provider = _make_provider(host)
    except Exception as exc:
        return FAILED, "could not build the pane host: " + _err(exc)
    candidates = _app_candidates(revit_handle)
    if not candidates:
        return FAILED, "no UIApplication / UIControlledApplication available at startup"
    errors = []
    for name, app in candidates:
        try:
            app.RegisterDockablePane(pane_id(), PANE_TITLE, provider)
        except Exception as exc:
            errors.append("%s: %s" % (name, _err(exc)))
            continue
        try:
            from System import AppDomain
            AppDomain.CurrentDomain.SetData(HOST_KEY, host)
        except Exception as exc:
            return FAILED, "registered, but the host could not be shared: " + _err(exc)
        return REGISTERED, name
    return FAILED, "; ".join(errors)
