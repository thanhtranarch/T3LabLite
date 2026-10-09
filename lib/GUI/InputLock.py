# -*- coding: utf-8 -*-
"""Shut off user input to every window of the current thread for a while.

This is what WPF's ShowDialog does before it runs its modal loop: every
visible, enabled top-level window of the UI thread (Revit's main window, the
floating palettes, the calling tool window) is disabled, and exactly those are
enabled again afterwards. Windows keep painting while disabled; they only stop
receiving mouse and keyboard input.

BatchOut needs it for its modeless export. The export runs inside
ExternalEvent.Execute on Revit's UI thread, so the queue can only repaint if
that thread pumps messages between two native export calls. A pump with input
live would also deliver every click the user queued meanwhile, to Revit and to
BatchOut, in the middle of the API call. With this lock held, the pump is the
same as the one the modal path has always run inside ShowDialog.

Win32 through ctypes, imported lazily: when ctypes or user32 is unavailable
acquire() returns False and the caller must not pump.
"""


class ThreadInputLock(object):
    """Disable this thread's visible top-level windows; release() restores them."""

    def __init__(self, api=None):
        # api = (user32, kernel32, enum_proc_type); tests inject doubles.
        self._api = api
        self._disabled = []

    @property
    def active(self):
        return bool(self._disabled)

    def acquire(self):
        """Disable the windows. True when at least one window is now locked."""
        if self._disabled:
            return True
        try:
            user32, kernel32, enum_proc = self._win32()
            found = []

            def visit(hwnd, _lparam):
                try:
                    if hwnd and user32.IsWindowVisible(hwnd) and user32.IsWindowEnabled(hwnd):
                        found.append(hwnd)
                except Exception:
                    pass
                return True

            # Keep the callback object alive for the whole enumeration.
            callback = enum_proc(visit)
            user32.EnumThreadWindows(kernel32.GetCurrentThreadId(), callback, 0)
            for hwnd in found:
                user32.EnableWindow(hwnd, False)
                self._disabled.append(hwnd)
        except Exception:
            # Never leave half the windows disabled.
            self.release()
            return False
        return self.active

    def release(self):
        """Enable again exactly the windows acquire() disabled. Safe to repeat."""
        disabled, self._disabled = self._disabled, []
        if not disabled:
            return
        try:
            user32 = self._win32()[0]
        except Exception:
            return
        for hwnd in reversed(disabled):
            try:
                if user32.IsWindow(hwnd):
                    user32.EnableWindow(hwnd, True)
            except Exception:
                pass

    def _win32(self):
        if self._api is None:
            import ctypes
            from ctypes import wintypes
            # Own WinDLL instances: argtypes set here must not change the shared
            # ctypes.windll.user32 another module may configure differently.
            user32 = ctypes.WinDLL('user32')
            kernel32 = ctypes.WinDLL('kernel32')
            enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
            user32.EnumThreadWindows.argtypes = [wintypes.DWORD, enum_proc, wintypes.LPARAM]
            user32.EnumThreadWindows.restype = wintypes.BOOL
            for name in ('IsWindow', 'IsWindowVisible', 'IsWindowEnabled'):
                function = getattr(user32, name)
                function.argtypes = [wintypes.HWND]
                function.restype = wintypes.BOOL
            user32.EnableWindow.argtypes = [wintypes.HWND, wintypes.BOOL]
            user32.EnableWindow.restype = wintypes.BOOL
            kernel32.GetCurrentThreadId.argtypes = []
            kernel32.GetCurrentThreadId.restype = wintypes.DWORD
            self._api = (user32, kernel32, enum_proc)
        return self._api
