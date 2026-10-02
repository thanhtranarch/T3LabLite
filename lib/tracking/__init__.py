# -*- coding: utf-8 -*-
"""
T3Lab Lite usage tracking -- everything that feeds the T3Lab Space dashboard.

  hooks.py    the three entry points the rest of the extension calls
  service.py  TelemetryService: payload, anonymous user hash, async send

Lite only: t3lab_dev has no copy of this folder, so never delete it or
replace it from dev. See README.md here and lite_guard/.
"""
from tracking.hooks import (  # noqa: F401
    track_mcp_call,
    track_ribbon_click,
    track_session_start,
)
