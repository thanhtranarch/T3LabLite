# -*- coding: utf-8 -*-
"""
T3Lab Telemetry Service
=======================
Records worldwide usage metrics for T3Lab Lite:
  1. Ribbon pushbutton clicks in Revit
  2. MCP AI tool invocations with sub-tool name, parameters, and purpose

Data is sent asynchronously in a background thread to the T3Lab Space backend
(default: https://t3lab.space/api/revit/tracking).

Privacy (see docs/PRIVACY.md and docs/compliance/):
  - Nothing is sent without the user's explicit "Allow" (tracking/consent.py).
    Silence is not consent (Luat 91/2025/QH15, Art. 9).
  - The rows are still personal data (pseudonymous ID + the IP the server
    sees), so keep the minimum:
      * user_hash is a random ID made on this machine, not derived from the
        Windows user or computer name;
      * doc_name is a salted code of the model title, never the title (a
        local copy's title ends with the Windows user name, and project names
        are client information);
      * MCP arguments go through redact_arguments: no code, secrets, free
        text (prompts, descriptions, messages), personal fields, full paths;
        e-mails, phone numbers and user folders are scrubbed from the rest
        and from error messages.
  - Zero-latency: never blocks Revit UI or MCP execution.
  - Fail-safe: all network exceptions are caught and silenced.

Lite only. The rest of the extension calls it through tracking/hooks.py.
"""

from __future__ import unicode_literals

import os
import sys
import json
import time
import re
import uuid
import hashlib
import threading
from datetime import datetime

# Safe standard library imports across Python 2/3 and IronPython
try:
    from urllib.request import Request, urlopen
    from urllib.error import URLError
except ImportError:
    # Python 2 / IronPython fallback
    try:
        from urllib2 import Request, urlopen, URLError
    except ImportError:
        Request = None
        urlopen = None

DEFAULT_TRACKING_URL = "https://t3lab.space/api/revit/tracking"

# Argument keys whose values never leave the machine.
_CODE_KEYS = ("code", "script")
_SECRET_HINTS = ("password", "passwd", "token", "secret", "api_key", "apikey")
# Keys holding a path: only the last part (file or folder name) is sent, so
# Windows user names and client folder structures stay local.
_PATH_HINTS = ("path", "folder", "directory", "dir", "file")
# Free text written by the user or the AI (prompts, goals, notes): it can hold
# anything, names and client details included, so only its length is sent.
_FREE_TEXT_KEYS = ("goal", "instruction", "prompt", "purpose", "description",
                   "query", "reason", "message", "text", "comment", "note",
                   "question", "content", "body", "summary")
# Keys that name or reach a person.
_PERSONAL_HINTS = ("email", "e_mail", "phone", "mobile", "user", "author",
                   "owner", "person", "address", "contact")
_MAX_TEXT = 300

# Scrubbed from every string that is still sent, error messages included.
_SCRUBBERS = (
    (re.compile(r"[\w.+-]+@[\w-]+(\.[\w-]+)+"), "(email)"),
    (re.compile(r"(?<![\w.])(\+?84|0)[\s.-]?\d{2,3}([\s.-]?\d{3}){2}(?![\w.])"), "(phone)"),
    (re.compile(r"([A-Za-z]:[\\/]+(Users|Documents and Settings)[\\/]+)[^\\/\s\"']+", re.I), r"\1(user)"),
    (re.compile(r"(/(home|Users)/)[^/\s\"']+"), r"\1(user)"),
)


def scrub_text(text, limit=_MAX_TEXT):
    """Remove e-mails, phone numbers and user folder names; cap the length."""
    if not isinstance(text, str):
        return text
    for pattern, repl in _SCRUBBERS:
        text = pattern.sub(repl, text)
    if limit and len(text) > limit:
        text = text[:limit - 3] + "..."
    return text


def _scrub_any(value):
    if isinstance(value, str):
        return scrub_text(value)
    if isinstance(value, (list, tuple)):
        return [_scrub_any(v) for v in value]
    if isinstance(value, dict):
        return redact_arguments(value)      # nested keys get the same rules
    return value


def _last_path_part(text):
    parts = [p for p in text.replace("\\", "/").split("/") if p]
    return parts[-1] if parts else ""


def _redact_value(key, value):
    k = (key or "").lower()
    if k in _CODE_KEYS:
        text = value if isinstance(value, str) else str(value or "")
        return "(code not sent, {} lines)".format(len(text.splitlines()) or 1)
    if any(h in k for h in _SECRET_HINTS):
        return "(not sent)"
    if k in _FREE_TEXT_KEYS:
        text = value if isinstance(value, str) else str(value or "")
        return "(text not sent, {} chars)".format(len(text))
    if any(h in k for h in _PERSONAL_HINTS):
        return "(not sent)"
    if any(h in k for h in _PATH_HINTS):
        if isinstance(value, str):
            return scrub_text(_last_path_part(value))
        if isinstance(value, (list, tuple)):
            return [scrub_text(_last_path_part(v)) if isinstance(v, str) else v
                    for v in value]
    return _scrub_any(value)


def redact_arguments(arguments):
    """Copy of an MCP arguments dict that is safe to send (see above).
    Nested steps (run_batch) are redacted the same way."""
    if not isinstance(arguments, dict):
        return {}
    out = {}
    for k, v in arguments.items():
        if k == "steps" and isinstance(v, list):
            out[k] = [dict(s, arguments=redact_arguments(s.get("arguments")))
                      if isinstance(s, dict) else s for s in v]
        else:
            out[k] = _redact_value(k, v)
    return out


INSTALL_ID_KEY = "tracking_install_id"


class TelemetryService(object):
    _cached_user_hash = None
    _cached_revit_version = None

    @classmethod
    def is_enabled(cls):
        """True only after the user chose "Allow" (tracking/consent.py)."""
        try:
            from tracking.consent import has_consent
            return has_consent()
        except Exception:
            return False

    @classmethod
    def get_tracking_url(cls):
        """Get the destination URL for telemetry."""
        try:
            from core import paths as _paths
            settings = _paths.load_settings()
            return settings.get("tracking_url") or DEFAULT_TRACKING_URL
        except Exception:
            return DEFAULT_TRACKING_URL

    @classmethod
    def get_user_hash(cls):
        """Random ID made once on this machine and kept in mcp_paths.json.

        Not derived from the Windows user or computer name, so it cannot be
        turned back into either. Delete "tracking_install_id" to get a new one.
        """
        if cls._cached_user_hash:
            return cls._cached_user_hash
        try:
            from core import paths as _paths
            value = _paths.load_settings().get(INSTALL_ID_KEY)
            if not (isinstance(value, str) and re.match(r"^[0-9a-f]{16}$", value)):
                value = uuid.uuid4().hex[:16]
                _paths.set_setting(INSTALL_ID_KEY, value)
            cls._cached_user_hash = value
            return value
        except Exception:
            return "anon_user"

    @classmethod
    def doc_code(cls, title):
        """Salted code of a model title: the same model gives the same code on
        this machine, but the title itself never leaves it."""
        if not title:
            return ""
        raw = "{}|{}".format(cls.get_user_hash(), title)
        return "doc_" + hashlib.sha256(raw.encode("utf-8")).hexdigest()[:12]

    @classmethod
    def get_revit_context(cls):
        """Resolve current Revit version and open document title safely."""
        revit_ver = cls._cached_revit_version
        doc_title = None

        try:
            from pyrevit import revit, HOST_APP
            if not revit_ver:
                revit_ver = getattr(HOST_APP, "version", None)
                if not revit_ver and getattr(revit, "doc", None):
                    app = revit.doc.Application
                    revit_ver = str(app.VersionNumber)
                cls._cached_revit_version = str(revit_ver or "unknown")

            if getattr(revit, "doc", None) and revit.doc:
                doc_title = revit.doc.Title
        except Exception:
            pass

        return revit_ver or "unknown", doc_title or ""

    @classmethod
    def extract_mcp_purpose(cls, tool_name, arguments):
        """
        Extract or construct a human-readable 'purpose / intent' for an MCP tool call.
        Fulfills: 'đối với mcp sẽ liệt kê ra tool nhỏ nào trong mcp được sử dụng và sử dụng mục đích gì'.
        """
        if not arguments or not isinstance(arguments, dict):
            arguments = {}

        # 1. Free text (goal, prompt, description...) is never sent
        #    (redact_arguments), so the purpose comes from the tool alone.

        # 2. Code execution tool
        if tool_name == "send_code_to_revit":
            # The code itself is never sent (redact_arguments).
            return "Thực thi mã tùy biến qua Revit API"

        # 3. Model construction tools
        if tool_name == "place_wall":
            wall_type = arguments.get("wall_type") or arguments.get("type_name") or "Standard"
            level = arguments.get("level") or arguments.get("level_name") or "Level 1"
            return "Tạo tường '{wall_type}' tại tầng '{level}'".format(wall_type=wall_type, level=level)

        if tool_name == "create_point_based_element":
            fam = arguments.get("family_name") or arguments.get("family") or "Element"
            typ = arguments.get("type_name") or arguments.get("type") or ""
            return "Đặt cấu kiện {fam} ({typ})".format(fam=fam, typ=typ).strip()

        if tool_name == "create_room":
            name = arguments.get("name") or arguments.get("room_name") or "Room"
            return "Tạo phòng '{name}'".format(name=name)

        if tool_name == "tag_all_walls":
            view = arguments.get("view_name") or "Active View"
            return "Đánh nhãn (Tag) toàn bộ tường trên view '{view}'".format(view=view)

        if tool_name == "tag_all_rooms":
            view = arguments.get("view_name") or "Active View"
            return "Đánh nhãn (Tag) toàn bộ phòng trên view '{view}'".format(view=view)

        if tool_name == "create_grid":
            name = arguments.get("name") or arguments.get("grid_name") or ""
            return "Tạo trục lưới công trình (Grid {name})".format(name=name).strip()

        if tool_name == "create_level":
            name = arguments.get("name") or arguments.get("level_name") or ""
            elev = arguments.get("elevation") or ""
            return "Tạo tầng (Level {name} cao độ {elev})".format(name=name, elev=elev).strip()

        if tool_name == "export_sheets_pdf":
            return "Xuất bản vẽ Revit sang định dạng PDF"

        if tool_name == "audit_model" or tool_name == "get_model_health":
            return "Kiểm tra và audit chất lượng mô hình BIM"

        if tool_name == "bulk_set_parameter":
            param = arguments.get("parameter_name") or "Parameter"
            return "Gán hàng loạt giá trị cho tham số '{param}'".format(param=param)

        if tool_name == "manage_view_template" or tool_name == "apply_view_template":
            tmpl = arguments.get("template_name") or "ViewTemplate"
            return "Áp dụng View Template '{tmpl}'".format(tmpl=tmpl)

        # 4. Fallback summary of top key-value arguments
        summaries = []
        for k, v in arguments.items():
            if k in ("self", "doc", "uiapp"):
                continue
            v_str = str(v)
            if len(v_str) > 40:
                v_str = v_str[:37] + "..."
            summaries.append("{}: {}".format(k, v_str))
            if len(summaries) >= 3:
                break

        if summaries:
            return "Thực thi {} ({})".format(tool_name, ", ".join(summaries))
        return "Gọi công cụ {}".format(tool_name)

    @classmethod
    def record_mcp_call(cls, tool_name, arguments=None, result=None, duration_ms=None):
        """
        Record an MCP tool execution.
        Fulfills: 'đối với mcp sẽ liệt kê ra tool nhỏ nào trong mcp được sử dụng và sử dụng mục đích gì'.
        """
        if not cls.is_enabled():
            return

        status = "success"
        err_msg = None
        if isinstance(result, dict):
            if result.get("isError") or result.get("error"):
                status = "error"
                err_msg = scrub_text(str(result.get("error") or "Execution failed"))

        arguments = redact_arguments(arguments)
        purpose = cls.extract_mcp_purpose(tool_name, arguments)

        # Clean sanitized details payload
        details = {}
        if isinstance(arguments, dict):
            for k, v in arguments.items():
                try:
                    # Filter out huge payloads
                    val_str = json.dumps(v)
                    if len(val_str) < 1000:
                        details[k] = v
                    else:
                        details[k] = "(large payload truncated)"
                except Exception:
                    details[k] = str(v)[:200]

        cls._dispatch_async({
            "tool_name": tool_name,
            "tool_type": "mcp",
            "panel": "MCP / AI",
            "purpose": purpose,
            "details": details,
            "status": status,
            "error_message": err_msg,
            "execution_time_ms": duration_ms,
        })

    @classmethod
    def record_tool_usage(cls, tool_name, tool_type="ribbon", panel=None, purpose=None,
                          details=None, status="success", error_message=None, duration_ms=None):
        """Record a general tool / ribbon click."""
        if not cls.is_enabled():
            return

        cls._dispatch_async({
            "tool_name": tool_name,
            "tool_type": tool_type,
            "panel": panel,
            "purpose": purpose or "Người dùng kích hoạt nút trên Revit Ribbon",
            "details": details or {},
            "status": status,
            "error_message": scrub_text(error_message) if error_message else error_message,
            "execution_time_ms": duration_ms,
        })

    @classmethod
    def _dispatch_async(cls, payload):
        """Send tracking payload in a background daemon thread."""
        def _worker():
            try:
                revit_ver, doc_title = cls.get_revit_context()
                payload["revit_version"] = revit_ver
                payload["doc_name"] = cls.doc_code(doc_title)
                payload["user_hash"] = cls.get_user_hash()
                payload["timestamp"] = datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%S.%fZ")

                url = cls.get_tracking_url()
                json_data = json.dumps(payload).encode("utf-8")

                sent = False
                if urlopen is not None and Request is not None:
                    try:
                        req = Request(url, data=json_data, headers={
                            "Content-Type": "application/json",
                            "User-Agent": "T3Lab-Lite-Telemetry/1.0",
                        })
                        urlopen(req, timeout=3.5)
                        sent = True
                    except Exception:
                        # IronPython's urllib2 HTTPS can fail where .NET works
                        if sys.platform != "cli":
                            raise
                if not sent:
                    # .NET WebClient fallback for IronPython
                    import clr
                    clr.AddReference("System")
                    from System.Net import WebClient
                    client = WebClient()
                    client.Headers.Add("Content-Type", "application/json")
                    client.Headers.Add("User-Agent", "T3Lab-Lite-Telemetry/1.0")
                    client.UploadString(url, "POST", json.dumps(payload))
            except Exception:
                # Silently catch all network issues -- telemetry must never break user workflows
                pass

        try:
            t = threading.Thread(target=_worker)
            t.daemon = True
            t.start()
        except Exception:
            pass
