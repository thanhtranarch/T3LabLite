# -*- coding: utf-8 -*-
"""
Assistant theme — chat-surface colours that follow Revit's own UI theme.

`_tb` / `_trgb` / `_theme_color` read a token; `_bind_*` make an element
FOLLOW a token (DynamicResource from code). Falls back to fixed light colours
when GUI.RevitTheme cannot be imported.

Split out of GUI/T3LabAssistantDialog.py, which re-imports every name
here: `from GUI.T3LabAssistantDialog import <name>` keeps working.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Theme"

from GUI.assistant._log import logger

# ─── Theme: follow Revit's own UI theme ───────────────────────────────────────
# All chat-surface colours come from here (no hex at call sites) so the window
# re-skins on Revit Light/Dark (2024+). See GUI/RevitTheme.py.
try:
    from GUI import RevitTheme as _theme

    def _tb(token):
        """SolidColorBrush for a theme token."""
        return _theme.brush(token)

    def _trgb(token):
        """(r, g, b) for a theme token."""
        return _theme.rgb(token)

    def _theme_color(token):
        """System.Windows.Media.Color for a theme token."""
        return _theme.color(token)

    def _bind_theme(element, prop_name, token):
        """Make `element.<prop_name>` FOLLOW a theme token instead of copying it.

        _tb() returns a frozen brush, so painted elements kept their colours
        after a Light/Dark switch. SetResourceReference is code-behind
        DynamicResource: repaints on key change, no re-render (mid-stream
        messages untouched). Falls back to plain assignment when prop_name is
        not a DependencyProperty.
        """
        try:
            from System.Windows import DependencyProperty
            dp = DependencyProperty.FromName(prop_name, element.GetType())
            if dp is not None:
                element.SetResourceReference(dp, _theme.RESOURCE_PREFIX + token)
                return True
        except Exception:
            pass
        try:
            setattr(element, prop_name, _tb(token))
        except Exception:
            pass
        return False

    def _bind_fg(element, token):
        """Foreground follows `token`."""
        return _bind_theme(element, 'Foreground', token)

    def _bind_bg(element, token):
        """Background follows `token`."""
        return _bind_theme(element, 'Background', token)

    def _bind_border(element, token):
        """BorderBrush follows `token`."""
        return _bind_theme(element, 'BorderBrush', token)

    def _bind_stroke(element, token):
        """Stroke follows `token` (Path / Ellipse / Line icons)."""
        return _bind_theme(element, 'Stroke', token)
except Exception as _theme_ex:                       # pragma: no cover
    logger.warning("RevitTheme unavailable, using light defaults: {}".format(
        _theme_ex))
    _theme = None
    # Keep in step with RevitTheme._LIGHT (paints the window when it can't import).
    _LIGHT_FALLBACK = {
        'AppBg': (245, 245, 245), 'ChatBg': (255, 255, 255),
        'ComposerBg': (255, 255, 255), 'CardBg': (255, 255, 255),
        'CardBorder': (222, 222, 222), 'Divider': (236, 236, 236),
        'PaneEdge': (222, 222, 222), 'UserBubbleBg': (219, 232, 246),
        'UserBubbleText': (0, 0, 0), 'BotText': (0, 0, 0),
        'Ink': (60, 60, 60), 'Muted': (90, 90, 90), 'Faint': (171, 171, 171),
        'Accent': (6, 150, 215), 'AccentSoft': (214, 233, 247),
        'Blue': (0, 110, 175), 'Success': (11, 122, 74), 'Danger': (196, 43, 28),
        'Amber': (245, 158, 11), 'IconFg': (145, 145, 145),
        'IconFgHover': (60, 60, 60), 'IconHoverBg': (234, 234, 234),
        'InputText': (60, 60, 60), 'InputCaret': (60, 60, 60),
        'CodeBg': (245, 245, 245), 'CodeFg': (28, 28, 28),
        'ScrollThumb': (193, 193, 193), 'SelectedBg': (227, 227, 227),
    }

    def _trgb(token):
        return _LIGHT_FALLBACK.get(token, _LIGHT_FALLBACK['Ink'])

    def _tb(token):
        from System.Windows.Media import SolidColorBrush, Color
        r, g, b = _trgb(token)
        return SolidColorBrush(Color.FromRgb(r, g, b))

    def _theme_color(token):
        from System.Windows.Media import Color
        r, g, b = _trgb(token)
        return Color.FromRgb(r, g, b)

    # Same surface as the themed branch; nothing to follow, so assign once.
    def _bind_theme(element, prop_name, token):
        try:
            setattr(element, prop_name, _tb(token))
        except Exception:
            pass
        return False

    def _bind_fg(element, token):
        return _bind_theme(element, 'Foreground', token)

    def _bind_bg(element, token):
        return _bind_theme(element, 'Background', token)

    def _bind_border(element, token):
        return _bind_theme(element, 'BorderBrush', token)

    def _bind_stroke(element, token):
        return _bind_theme(element, 'Stroke', token)
