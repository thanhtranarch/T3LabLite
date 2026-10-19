# -*- coding: utf-8 -*-
"""Windows DPAPI wrapper for secrets kept in %APPDATA%\\T3LabAI (API keys,
the Tool Hub session).

A protected value is stored as ``"dpapi:" + base64(ProtectedData.Protect(...))``
with scope CurrentUser: only the same Windows account on the same machine can
read it back, so a copied settings.json leaks nothing.

Without .NET (tests, Linux dev boxes) ``protect`` is the identity function and
``unprotect`` returns plaintext unchanged, so the same file keeps working.
Runs under both CPython 3 (pythonnet) and IronPython 2.7.
"""
from __future__ import unicode_literals

import base64

PREFIX = 'dpapi:'

_protector = None      # (protect(bytes)->bytes, unprotect(bytes)->bytes) or False


def _dotnet_protector():
    try:
        import clr  # noqa: F401
        # .NET Framework 4.8 (Revit 2022-2024): System.Security.dll.
        # .NET 8 (Revit 2025+): the System.Security.Cryptography.ProtectedData
        # assembly Revit ships. Neither -> plaintext, as before DPAPI.
        for asm in ('System.Security', 'System.Security.Cryptography.ProtectedData'):
            try:
                clr.AddReference(asm)
            except Exception:
                pass
        from System.Security.Cryptography import ProtectedData, DataProtectionScope
        from System import Array, Byte
    except Exception:
        return False

    def _to_net(data):
        return Array[Byte](bytearray(data))

    def _from_net(arr):
        return bytes(bytearray(arr))

    def protect(data):
        return _from_net(ProtectedData.Protect(
            _to_net(data), None, DataProtectionScope.CurrentUser))

    def unprotect(data):
        return _from_net(ProtectedData.Unprotect(
            _to_net(data), None, DataProtectionScope.CurrentUser))

    return protect, unprotect


def set_protector(pair):
    """Tests inject a fake ``(protect, unprotect)``; ``None`` resets detection."""
    global _protector
    _protector = pair


def _get():
    global _protector
    if _protector is None:
        _protector = _dotnet_protector()
    return _protector


def available():
    return bool(_get())


def is_protected(value):
    try:
        return value.startswith(PREFIX)
    except AttributeError:
        return False


def protect(value):
    """Plaintext -> stored form. Empty values and already-protected ones pass."""
    if not value or is_protected(value):
        return value
    pair = _get()
    if not pair:
        return value
    try:
        blob = pair[0](value.encode('utf-8'))
        return PREFIX + base64.b64encode(blob).decode('ascii')
    except Exception:
        return value


def unprotect(value):
    """Stored form -> plaintext. Returns None when the blob cannot be opened
    (another Windows account, another machine): the caller asks again."""
    if not value or not is_protected(value):
        return value
    pair = _get()
    if not pair:
        return None
    try:
        blob = base64.b64decode(value[len(PREFIX):].encode('ascii'))
        return pair[1](blob).decode('utf-8')
    except Exception:
        return None
