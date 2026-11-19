# -*- coding: utf-8 -*-
"""Agent cursors overlay: the WPF windows that draw what core/agent_cursors.py
reports. Two owned, non-activating windows: a click-through canvas over the
visible views and a HUD bar. Engine-agnostic (IronPython 2.7 + CPython 3).
"""

from __future__ import unicode_literals

import os
import sys
import time

from core.agent_cursors import (
    ACTIONS, BAND_STRIPE, BAR_W, BOX_HOLD_S, COUNT_BAR_W, DWELL_MS, FEED_SHOWN, FEED_SIZE,
    FRESH_HOLD_S, IDLE_HIDE_S, MAX_STOPS, MOVE_MS, OPTION_GROUPS, OPTION_LABELS, REPLAY_SIZE,
    STATES, SUMMARY_S, SWEEP_W, TAG_MAX, action_style, agent_style, ago_text,
    band_motion, band_rows, band_style, band_wash, bar_tone, classify_action,
    cluster_stops, compact_line, count_bar, count_items, describe_call, fill_width,
    footer_text, front_hides, hop_ms, hud_offset, hud_palette, hud_place, hud_state,
    motion_off, offscreen_markers, on_model, options, place_tags, plan_progress,
    scan_points, set_enabled, set_hud_offset, set_option, short_text, should_collapse,
    speed_factor, status_text, tour_schedule, union_box)


class OverlayHost(object):
    """Owns the overlay. Under IronPython it lives on its own STA thread with
    its own Dispatcher, so cursors keep moving ("Working 8 s") while Revit's
    main thread runs a long tool call. Under CPython a thread of ours would
    freeze whenever Revit is idle (CLAUDE.md rule 8), so it stays on the
    calling thread. call() is safe from any thread. Env
    T3LAB_CURSORS_MAIN_THREAD=1 forces the calling thread."""

    def __init__(self, owner_hwnd, callbacks=None):
        self.callbacks = callbacks or {}
        self.revit_hwnd = owner_hwnd    # Revit's main window, also on our own thread
        self.threaded = False
        self._ov = None
        self._dispatcher = None
        if sys.platform == 'cli' and os.environ.get('T3LAB_CURSORS_MAIN_THREAD') != '1':
            try:
                self._start_thread()
            except Exception:
                self._ov = None
        if self._ov is None:
            self._ov = _Overlay(owner_hwnd, self)
            self._dispatcher = self._ov.win.Dispatcher

    @property
    def alive(self):
        return self._ov is not None and self._ov.alive

    def _start_thread(self):
        import threading
        from System.Threading import Thread, ThreadStart, ApartmentState
        ready = threading.Event()
        box = {}

        def run():
            try:
                from System.Windows.Threading import Dispatcher
                box['ov'] = _Overlay(None, self)
                box['disp'] = Dispatcher.CurrentDispatcher
            except Exception as ex:
                box['error'] = ex
            finally:
                ready.set()
            if 'ov' in box:
                Dispatcher.Run()

        thread = Thread(ThreadStart(run))
        thread.SetApartmentState(ApartmentState.STA)
        thread.IsBackground = True
        thread.Name = 'T3Lab Agent Cursors'
        thread.Start()
        ready.wait(5.0)
        if 'ov' not in box:
            raise RuntimeError(box.get('error') or 'overlay thread did not start')
        self._ov, self._dispatcher, self.threaded = box['ov'], box['disp'], True

    def call(self, name, *args, **kwargs):
        """Run overlay.<name>(...) on the overlay's thread. Never raises."""
        ov = self._ov

        def run():
            try:
                getattr(ov, name)(*args, **kwargs)
            except Exception:
                pass
        try:
            if self._dispatcher.CheckAccess():
                run()
            else:
                from System import Action
                self._dispatcher.BeginInvoke(Action(run))
        except Exception:
            pass


def _NO_STYLE(enum):
    # `WindowStyle.None` is a syntax error in Python.
    return getattr(enum, 'None')


_GWL_EXSTYLE = -20
_WS_EX_TRANSPARENT = 0x20
_WS_EX_TOOLWINDOW = 0x80
_WS_EX_NOACTIVATE = 0x08000000
_MAX_TRANSIENT = 400        # animation handlers kept alive (see _hook)
BUSY_GAP_S = 4              # a gap this short between calls still reads "Working"
# Glyph font with fallback: Windows 11 ships Segoe Fluent Icons (a superset);
# a codepoint missing in the first font is taken from the next, never tofu
# (owner, 2026-11-17: a broken icon on the bar).
ICON_FONTS = 'Segoe Fluent Icons, Segoe MDL2 Assets'
# The agent's pointer: a bare tip, no tail (owner, 2026-11-17: minimal cursor,
# the caption stays). The bar draws it small as the agent's avatar.
CURSOR_PATH = 'M1,1 L1,14 L5,10.2 L12,9.8 Z'


def _set_exstyle(window, bits):
    """Add extended window styles (click-through / no focus); best effort."""
    try:
        import ctypes
        from System.Windows.Interop import WindowInteropHelper
        hwnd = WindowInteropHelper(window).Handle.ToInt64()
        user32 = ctypes.windll.user32
        style = user32.GetWindowLongW(ctypes.c_void_p(hwnd), _GWL_EXSTYLE)
        user32.SetWindowLongW(ctypes.c_void_p(hwnd), _GWL_EXSTYLE, style | bits)
    except Exception:
        pass


def _opacity_property():
    from System.Windows import UIElement
    return UIElement.OpacityProperty


class _Overlay(object):
    """Two owned, non-activating windows: a click-through canvas over the
    visible views for cursors, boxes and effects, and a small HUD with the
    agents, their action counts, the views they opened and live toggles."""

    def __init__(self, owner_hwnd, host=None):
        import clr
        for ref in ('PresentationCore', 'PresentationFramework', 'WindowsBase'):
            try:
                clr.AddReference(ref)
            except Exception:
                pass
        from System.Windows import Window, WindowStyle, ResizeMode, SizeToContent
        from System.Windows.Controls import Canvas
        from System.Windows.Media import Brushes

        self.alive = True
        self.host = host
        self._owner = owner_hwnd
        self._goals = {}            # agent -> goal banner text
        self._plans = {}            # agent -> {'tasks': [...], 'step': n} (t3lab_narrate)
        self._asks = []             # consent questions on the bar: dicts, oldest first
        self._working = {}          # agent -> (started, caption)
        self._waiting = {}          # agent -> caption
        self._acting = {}           # agent -> action of its running / waiting call
        self._replay_step = None    # the step a replay shows now (band_rows `last`)
        self._loops = []            # (target, property) of the bar's endless animations
        self._drawn = {}            # what the bar last showed: only a change moves (_motion)
        self._leaving = None        # token while the cursors fade out (_enter_summary)
        self._focus = {}            # agent -> frames round what its running call names
        self._history = []          # replayable show() records, oldest first
        self._replay = None         # [records, index] while replaying
        self._summary = None        # time the idle summary card appeared
        self._touched = 0
        self._marks = []            # off-screen arrows of the last call
        self._fx = {}               # Effects switches of the call being drawn (_effects)
        self._cursors = {}          # agent -> dict(group, tx, ...): one cursor per agent
        self._counts = {}           # agent -> {action: n}
        self._views = []            # [(view id, name)] the agents worked in
        self._feed = []             # last steps, newest first (HUD)
        self._actions = 0
        self._last = 0.0
        self._scale = None
        self._origin = (0.0, 0.0)   # screen px of the canvas top-left
        self._handlers = []         # strong refs: delegates must outlive the GC
        self._transient = []

        self.win = Window()
        self.win.WindowStyle = _NO_STYLE(WindowStyle)
        self.win.AllowsTransparency = True
        self.win.Background = Brushes.Transparent
        self.win.ResizeMode = ResizeMode.NoResize
        self.win.ShowInTaskbar = False
        self.win.ShowActivated = False
        self.win.Focusable = False
        self.win.IsHitTestVisible = False
        self.win.Title = 'T3Lab Agent Cursors'
        self.canvas = Canvas()
        self.canvas.IsHitTestVisible = False
        self.win.Content = self.canvas
        self._own(self.win)
        self._hook(self.win, 'SourceInitialized',
                   lambda s, e: _set_exstyle(self.win, _WS_EX_TRANSPARENT |
                                             _WS_EX_TOOLWINDOW | _WS_EX_NOACTIVATE))
        self._hook(self.win, 'Closed', lambda s, e: setattr(self, 'alive', False))

        self._theme = 'dark'                # set_theme() follows Revit
        self._pal = hud_palette('dark')
        self._paused = False        # Pause on the bar (agent_cursors.set_paused)
        self._folded = False        # the user folded the bar to one line
        self._hover = False
        self._drawer = False        # options drawer (⚙) open
        self._note = ''             # replaces the footer text (replay progress)
        self._failed = 0
        self._last_changes = []     # 'Mark: A-101 → A-102' of the last step
        self._view_dip = None       # active view (DIP) the bar sits in
        self._status_label = None

        self.hud = Window()
        self.hud.WindowStyle = _NO_STYLE(WindowStyle)
        self.hud.AllowsTransparency = True
        self.hud.Background = Brushes.Transparent
        self.hud.ResizeMode = ResizeMode.NoResize
        self.hud.ShowInTaskbar = False
        self.hud.ShowActivated = False
        self.hud.SizeToContent = SizeToContent.WidthAndHeight
        self.hud.Title = 'T3Lab Agent Cursors'
        self._own(self.hud)
        self._hook(self.hud, 'SourceInitialized',
                   lambda s, e: _set_exstyle(self.hud, _WS_EX_TOOLWINDOW | _WS_EX_NOACTIVATE))
        self._build_hud()

        from System import TimeSpan
        from System.Windows.Threading import DispatcherTimer
        self._timer = DispatcherTimer()
        self._timer.Interval = TimeSpan.FromSeconds(1)
        self._hook(self._timer, 'Tick', self._on_tick)
        self._replay_timer = DispatcherTimer()
        self._hook(self._replay_timer, 'Tick', self._on_replay_tick)
        if not owner_hwnd:
            # Own thread: no Win32 owner (a cross-thread owner ties input
            # queues to Revit's busy thread), so stay on top while Revit is
            # in front and hide while another app is.
            self.win.Topmost = self.hud.Topmost = True
            self._front_timer = DispatcherTimer()
            self._front_timer.Interval = TimeSpan.FromMilliseconds(400)
            self._hook(self._front_timer, 'Tick', self._on_front_tick)
            self._front_timer.Start()
        self._wanted = False        # overlay should be visible when Revit is in front
        self._front_ok = True       # nothing covers Revit's views (_on_front_tick)
        self._canvas_px = None      # screen px rect the canvas covers (_place)
        self._resting = False       # idle with Keep bar on: summary shown, nothing hidden

    # ── plumbing ──────────────────────────────────────────────────────────────

    def _hook(self, source, event, fn, transient=False):
        keep = self._transient if transient else self._handlers
        keep.append(fn)
        if transient and len(keep) > _MAX_TRANSIENT:
            del keep[:len(keep) - _MAX_TRANSIENT]
        handler = getattr(source, event)
        handler += fn
        setattr(source, event, handler)

    def _own(self, window):
        if not self._owner:
            return
        try:
            from System.Windows.Interop import WindowInteropHelper
            WindowInteropHelper(window).Owner = self._owner
        except Exception:
            pass

    @staticmethod
    def _brush(hex_color, alpha=255):
        from System.Windows.Media import SolidColorBrush, Color
        h = hex_color.lstrip('#')
        return SolidColorBrush(Color.FromArgb(alpha, int(h[0:2], 16),
                                              int(h[2:4], 16), int(h[4:6], 16)))

    def _dpi(self):
        if self._scale:
            return self._scale
        try:
            from System.Windows import PresentationSource
            from System.Windows.Interop import WindowInteropHelper
            # The window may not be shown yet (Revit covered): create its
            # handle so the scale is real, not the 1.0 fallback.
            WindowInteropHelper(self.win).EnsureHandle()
            src = PresentationSource.FromVisual(self.win)
            self._scale = float(src.CompositionTarget.TransformToDevice.M11) or 1.0
        except Exception:
            return 1.0
        return self._scale

    def _local(self, box):
        """Screen px box -> canvas DIP (x0, y0, x1, y1)."""
        s = self._dpi()
        ox, oy = self._origin
        return ((box[0] - ox) / s, (box[1] - oy) / s, (box[2] - ox) / s, (box[3] - oy) / s)

    @staticmethod
    def _text(text, size=11, bold=True, color='#FFFFFF', glyph=False):
        from System.Windows import FontWeights
        from System.Windows.Controls import TextBlock
        from System.Windows.Media import FontFamily
        tb = TextBlock()
        tb.Text = text
        tb.FontFamily = FontFamily(ICON_FONTS if glyph else 'Segoe UI')
        tb.FontSize = size
        tb.FontWeight = FontWeights.SemiBold if bold and not glyph else FontWeights.Normal
        tb.Foreground = _Overlay._brush(color)
        return tb

    @staticmethod
    def _pill(child, color, radius=4, alpha=255):
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border
        b = Border()
        b.Background = _Overlay._brush(color, alpha)
        b.CornerRadius = CornerRadius(radius)
        b.Padding = Thickness(6, 2, 6, 2)
        b.Child = child
        return b

    def _row(self, *children):
        from System.Windows import VerticalAlignment
        from System.Windows.Controls import StackPanel, Orientation
        row = StackPanel()
        row.Orientation = Orientation.Horizontal
        for child in children:
            child.VerticalAlignment = VerticalAlignment.Center
            row.Children.Add(child)
        return row

    def _glyph_label(self, glyph, label, color, size=10):
        from System.Windows import Thickness
        icon = self._text(glyph, size, color=color, glyph=True)
        icon.Margin = Thickness(0, 0, 4, 0)
        return self._row(icon, self._text(label, size, color=color))

    def _anim(self, frames, done=None, ease=False):
        """DoubleAnimationUsingKeyFrames from [(value, ms)]; ease = slow down
        into each frame instead of a straight line."""
        from System import TimeSpan
        from System.Windows.Media.Animation import (
            DoubleAnimationUsingKeyFrames, LinearDoubleKeyFrame, KeyTime)
        curve = None
        if ease:
            from System.Windows.Media.Animation import CubicEase, EasingMode
            curve = CubicEase()
            curve.EasingMode = EasingMode.EaseOut
        anim = DoubleAnimationUsingKeyFrames()
        for k, (value, ms) in enumerate(frames):
            at = KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(ms))
            if curve is not None and k:
                from System.Windows.Media.Animation import EasingDoubleKeyFrame
                anim.KeyFrames.Add(EasingDoubleKeyFrame(value, at, curve))
            else:
                anim.KeyFrames.Add(LinearDoubleKeyFrame(value, at))
        if done is not None:
            self._hook(anim, 'Completed', done, transient=True)
        return anim

    # Motion on the bar. It is decoration: Animation Off skips it and a failure
    # never costs the bar (_motion). The bar is rebuilt on every refresh, so
    # only what changed since the last draw moves (self._drawn, band_motion).

    def _motion(self, fn, *args):
        if self._reduce():
            return
        try:
            fn(*args)
        except Exception:
            pass

    def _loop(self, target, prop, frames):
        """An endless animation, remembered so a rebuild can stop it."""
        from System.Windows.Media.Animation import RepeatBehavior
        anim = self._anim(frames)
        anim.RepeatBehavior = RepeatBehavior.Forever
        target.BeginAnimation(prop, anim)
        self._loops.append((target, prop))

    def _stop_loops(self):
        for target, prop in self._loops:
            try:
                target.BeginAnimation(prop, None)
            except Exception:
                pass
        self._loops = []

    def _fade_in(self, element, ms=180):
        element.BeginAnimation(_opacity_property(), self._anim(((0.0, 0), (1.0, ms))))

    def _slide_in(self, element, dy=6.0, ms=240):
        """Fade in while moving `dy` px into place."""
        from System.Windows.Media import TranslateTransform
        tt = TranslateTransform()
        element.RenderTransform = tt
        tt.BeginAnimation(_translate_y(), self._anim(((dy, 0), (0.0, ms)), ease=True))
        self._fade_in(element, ms)

    def _pop(self, element, cx, cy, scale=1.45, ms=260):
        """Swell and settle: something on it just changed."""
        self._scale_anim(element, cx, cy, ((scale, 0), (1.0, ms)))

    def _shake(self, element):
        from System.Windows.Media import TranslateTransform
        tt = TranslateTransform()
        element.RenderTransform = tt
        tt.BeginAnimation(_translate_x(), self._anim(
            ((0.0, 0), (-6.0, 60), (6.0, 120), (-4.0, 180), (0.0, 240))))

    def _grow(self, element, start, end, ms=320):
        """Width from `start` to `end` (the element already has `end`)."""
        from System.Windows import FrameworkElement
        if abs(end - start) >= 1.0:
            element.BeginAnimation(FrameworkElement.WidthProperty,
                                   self._anim(((start, 0), (end, ms)), ease=True))

    def _breathe(self, element, low=0.45, ms=1400):
        self._loop(element, _opacity_property(), ((1.0, 0), (low, ms // 2), (1.0, ms)))

    def _pulse(self, element, delay_ms, hold_ms, done=None):
        """Fade in at delay, hold, fade out — one animation: a second
        BeginAnimation on Opacity would replace the first."""
        element.BeginAnimation(_opacity_property(), self._anim(
            ((0.0, delay_ms), (1.0, delay_ms + 180), (1.0, delay_ms + 180 + hold_ms),
             (0.0, delay_ms + 780 + hold_ms)), done))

    def _remover(self, element):
        def remove(s, e):
            try:
                self.canvas.Children.Remove(element)
            except Exception:
                pass
        return remove

    def _scale_anim(self, element, cx, cy, frames):
        """Scale `element` around its own (cx, cy)."""
        from System.Windows.Media import ScaleTransform
        st = ScaleTransform()
        st.CenterX, st.CenterY = cx, cy
        element.RenderTransform = st
        for prop in (ScaleTransform.ScaleXProperty, ScaleTransform.ScaleYProperty):
            st.BeginAnimation(prop, self._anim(frames))

    # ── HUD ───────────────────────────────────────────────────────────────────
    # One bar at the bottom of the view (owner, 2026-11-17 redesign):
    #   header   cursor-mark avatars · title · status pill · ⏸ ⚙ — ✕ (drag to move)
    #   drawer   options (⚙), grouped switches + speed
    #   goal     goal + "Step k/n" + progress bar per agent with a plan
    #   bands    what runs now, one colour block per agent (band_rows), else the
    #            last step + before→after chips; a replay fills its block to k / n
    #   feed     last FEED_SHOWN steps: action stripe, ✓ / ✕, how long ago
    #   agents   one row per agent with its counts as a stacked bar (2+ agents, or idle)
    #   footer   "12 actions · 1 failed · 2 views" · Select touched · Replay
    # Folded (— or idle COLLAPSE_S): one line, mouse over opens it again.

    def _c(self, key):
        return self._pal[key]

    def _build_hud(self):
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border, StackPanel
        self._bar = Border()
        self._bar.CornerRadius = CornerRadius(8)
        self._bar.Padding = Thickness(12, 10, 12, 10)
        col = StackPanel()
        self._bar.Child = col
        self._ask_panel = StackPanel()          # Allow / Deny cards (ask())
        col.Children.Add(self._ask_panel)
        self._full = StackPanel()
        col.Children.Add(self._full)
        self._compact = StackPanel()
        col.Children.Add(self._compact)
        self.hud.Content = self._bar
        self._hook(self.hud, 'MouseEnter', lambda s, e: self._set_hover(True))
        self._hook(self.hud, 'MouseLeave', lambda s, e: self._set_hover(False))
        self._apply_theme()

    def _apply_theme(self):
        self._bar.Background = self._brush(self._c('bg'), 246)
        self._bar.BorderBrush = self._brush(self._c('line'))
        self._bar.BorderThickness = _thickness(1)

    def set_theme(self, theme):
        """Revit switched light / dark (agent_cursors._sync_theme)."""
        self._theme = 'light' if theme == 'light' else 'dark'
        self._pal = hud_palette(theme)
        self._apply_theme()
        if self.hud.IsVisible:
            self._refresh_hud()

    def set_paused(self, on):
        self._paused = bool(on)
        if self.hud.IsVisible:
            self._refresh_hud()

    def _set_hover(self, on):
        self._hover = on
        if on:
            self._last_hover = time.time()
        if self.hud.IsVisible:
            self._refresh_hud()

    # small parts

    def _mark(self, color, size=16):
        """The agent's cursor, small: its avatar on the bar."""
        from System.Windows.Controls import Canvas
        from System.Windows.Media import Geometry, ScaleTransform
        from System.Windows.Shapes import Path
        holder = Canvas()
        holder.Width = size * 0.8
        holder.Height = size
        arrow = Path()
        arrow.Data = Geometry.Parse(CURSOR_PATH)
        arrow.Fill = self._brush(color)
        arrow.RenderTransform = ScaleTransform(size / 15.0, size / 15.0)
        holder.Children.Add(arrow)
        return holder

    def _icon_button(self, glyph, tip, action, active=False):
        from System.Windows import Thickness, CornerRadius, HorizontalAlignment, VerticalAlignment
        from System.Windows.Controls import Border
        from System.Windows.Input import Cursors
        icon = self._text(glyph, 11, color=self._c('text2'), glyph=True)
        icon.HorizontalAlignment = HorizontalAlignment.Center
        icon.VerticalAlignment = VerticalAlignment.Center
        btn = Border()
        btn.Width = btn.Height = 24
        btn.CornerRadius = CornerRadius(4)
        btn.Background = self._brush(self._c('chip') if active else self._c('bg'), 255 if active else 1)
        btn.Child = icon
        btn.Margin = Thickness(2, 0, 0, 0)
        btn.Cursor = Cursors.Hand
        btn.ToolTip = tip
        self._hook(btn, 'MouseEnter', lambda s, e: setattr(btn, 'Background', self._brush(self._c('chip'))),
                   transient=True)
        self._hook(btn, 'MouseLeave', lambda s, e: setattr(btn, 'Background', self._brush(
            self._c('chip') if active else self._c('bg'), 255 if active else 1)), transient=True)

        def click(s, e):
            try:
                e.Handled = True
                action()
                self._refresh_hud()
            except Exception:
                pass
        self._hook(btn, 'MouseLeftButtonDown', _swallow, transient=True)
        self._hook(btn, 'MouseLeftButtonUp', click, transient=True)
        return btn

    def _button(self, glyph, label, tip, action, primary=False):
        """A text button on the bar; refreshes the bar after `action`."""
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border
        from System.Windows.Input import Cursors
        fg = self._c('on_accent') if primary else self._c('text')
        btn = Border()
        btn.CornerRadius = CornerRadius(4)
        btn.Padding = Thickness(8, 3, 10, 3)
        btn.Background = self._brush(self._c('accent') if primary else self._c('chip'))
        btn.Child = self._glyph_label(glyph, label, fg, 11.5)
        btn.Margin = Thickness(6, 0, 0, 0)
        btn.Cursor = Cursors.Hand
        btn.ToolTip = tip

        def click(s, e):
            try:
                e.Handled = True
                action()
                self._refresh_hud()
            except Exception:
                pass
        self._hook(btn, 'MouseLeftButtonDown', _swallow, transient=True)
        self._hook(btn, 'MouseLeftButtonUp', click, transient=True)
        return btn

    def _status_pill(self, state, seconds=None):
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border
        from System.Windows.Shapes import Ellipse
        color = self._c(STATES[state][1])
        dot = Ellipse()
        dot.Width = dot.Height = 7
        dot.Fill = self._brush(color)
        dot.Margin = Thickness(0, 0, 6, 0)
        pill = Border()
        pill.CornerRadius = CornerRadius(10)
        pill.Padding = Thickness(7, 1, 9, 2)
        pill.Background = self._brush(color, 40)
        label = self._text(status_text(state, seconds), 11.5, bold=False, color=color)
        self._status_label = label if state == 'working' else None
        if state in ('working', 'waiting'):
            self._motion(self._breathe, dot, 0.3, 1100)
        pill.Child = self._row(dot, label)
        pill.Margin = Thickness(8, 0, 0, 0)
        return pill

    def _switch(self, on, flip):
        from System.Windows import Thickness, CornerRadius, HorizontalAlignment
        from System.Windows.Controls import Border
        from System.Windows.Input import Cursors
        from System.Windows.Shapes import Ellipse
        knob = Ellipse()
        knob.Width = knob.Height = 12
        knob.Fill = self._brush('#FFFFFF')
        knob.HorizontalAlignment = HorizontalAlignment.Right if on else HorizontalAlignment.Left
        track = Border()
        track.Width, track.Height = 28, 16
        track.CornerRadius = CornerRadius(8)
        track.Padding = Thickness(2)
        track.Background = self._brush(self._c('accent') if on else self._c('track'))
        track.Child = knob
        track.Cursor = Cursors.Hand
        self._hook(track, 'MouseLeftButtonUp', flip, transient=True)
        return track

    def _line_row(self, left, right=None, margin_top=0):
        """Grid: `left` stretches, `right` hugs the right edge."""
        from System.Windows import Thickness, GridLength, GridUnitType, VerticalAlignment
        from System.Windows.Controls import Grid, ColumnDefinition
        grid = Grid()
        for width in (GridLength(1, GridUnitType.Star), GridLength.Auto):
            cd = ColumnDefinition()
            cd.Width = width
            grid.ColumnDefinitions.Add(cd)
        left.VerticalAlignment = VerticalAlignment.Center
        grid.Children.Add(left)
        if right is not None:
            right.VerticalAlignment = VerticalAlignment.Center
            Grid.SetColumn(right, 1)
            grid.Children.Add(right)
        grid.Margin = Thickness(0, margin_top, 0, 0)
        return grid

    def _tone(self, action):
        """The action's colour as this bar draws it (agent_cursors.bar_tone)."""
        return bar_tone(action_style(action)[2], self._theme)

    def _counts_row(self, agent, counts):
        """An agent's counts: one stacked bar in the actions' colours (each
        part grows from the width it had), then glyph + number per action."""
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border
        parts = []
        was, now = self._drawn.get('counts') or {}, self._drawn.setdefault('counts_now', {})
        for k, (action, width) in enumerate(count_bar(counts, COUNT_BAR_W, 1.0)):
            seg = Border()
            seg.Width, seg.Height = width, 6
            seg.CornerRadius = CornerRadius(1)
            seg.Background = self._brush(self._tone(action))
            seg.Margin = Thickness(1 if k else 0, 0, 0, 0)
            seg.ToolTip = u'{} {}'.format(action_style(action)[0], counts.get(action, 0))
            self._motion(self._grow, seg, was.get((agent, action), 0.0), width)
            now[(agent, action)] = width
            parts.append(seg)
        for action, n in count_items(counts):
            label, glyph, _color = action_style(action)
            icon = self._text(glyph, 11, color=self._tone(action), glyph=True)
            icon.Margin = Thickness(6, 0, 3, 0)
            icon.ToolTip = u'{} {}'.format(label, n)
            parts.append(icon)
            parts.append(self._text('{}'.format(n), 11, color=self._c('muted')))
        return self._row(*parts)

    # the bar

    def _state(self):
        # A run of calls stays "Working" in the gaps between them.
        recent = [st['t'] for st in self._feed[:1] if not st.get('kind')]
        busy = self._working or bool(recent and time.time() - recent[0] < BUSY_GAP_S)
        return hud_state(busy, self._waiting, self._asks, self._paused)

    def _refresh_hud(self):
        from System.Windows import Visibility
        state = self._state()
        idle_s = time.time() - self._last if state == 'idle' else 0
        folded = should_collapse(self._folded, idle_s, self._hover, self._asks)
        self._full.Visibility = Visibility.Collapsed if folded else Visibility.Visible
        self._compact.Visibility = Visibility.Visible if folded else Visibility.Collapsed
        self._ask_panel.Visibility = (Visibility.Collapsed if folded else Visibility.Visible)
        self._stop_loops()
        if folded:
            self._fill_compact(state)
        else:
            self._fill_full(state)
        if self._drawn.get('folded') != folded:     # opened, folded or first shown
            self._drawn['folded'] = folded
            self._motion(self._fade_in, self._compact if folded else self._full)
        self._place_hud()

    def _seconds(self):
        if not self._working:
            return None
        return time.time() - min(started for started, _c in self._working.values())

    def _agents(self):
        return sorted(set(self._counts) | set(self._goals) | set(self._plans) |
                      set(self._working))

    def _fill_compact(self, state):
        from System.Windows import Thickness
        from System.Windows.Input import Cursors
        from System.Windows.Shapes import Ellipse
        self._compact.Children.Clear()
        agents = self._agents()
        parts = []
        for agent in agents[:3]:
            mark = self._mark(agent_style(agent)[1], 14)
            mark.Margin = Thickness(0, 0, 4, 0)
            parts.append(mark)
        dot = Ellipse()
        dot.Width = dot.Height = 7
        dot.Fill = self._brush(self._c(STATES[state][1]))
        dot.Margin = Thickness(4, 0, 8, 0)
        if state in ('working', 'waiting'):
            self._motion(self._breathe, dot, 0.3, 1100)
        parts.append(dot)
        caption, progress = self._now_caption()
        if self._asks:
            item = self._asks[0]
            caption = u'{} waits for your OK'.format(agent_style(item['agent'])[0])
        text = self._text(compact_line(caption or status_text(state), progress), 12.5, bold=False,
                          color=self._c('text'))
        parts.append(text)
        if self._asks:
            item = self._asks[0]
            parts.append(self._button(u'', 'Allow', 'Run it',
                                      lambda rid=item['rid']: self._answer(rid, True), primary=True))
            parts.append(self._button(u'', 'Deny', 'Do not run it',
                                      lambda rid=item['rid']: self._answer(rid, False)))
        expand = self._icon_button(u'', 'Open the bar', self._unfold)
        expand.Margin = Thickness(8, 0, 0, 0)
        parts.append(expand)
        row = self._row(*parts)
        row.Cursor = Cursors.SizeAll
        self._hook(row, 'MouseLeftButtonDown', self._drag, transient=True)
        self._compact.Children.Add(row)

    def _unfold(self):
        self._folded = False
        self._last = time.time()

    def _toggle_fold(self):
        self._folded = not self._folded
        if not self._folded:
            self._last = time.time()

    def _now_caption(self):
        """(caption, plan progress) for the one-line bar: the newest running
        call, else the step a replay shows, else the last step."""
        progress = None
        for agent in self._agents():
            progress = plan_progress(self._plans.get(agent))
            if progress:
                break
        if self._working:
            key = max(self._working, key=lambda k: self._working[k][0])
            return self._working[key][1], progress
        rows = self._bands()[0]
        if not rows:
            return '', progress
        if rows[0]['state'] == 'replay':
            return u'{} \u00b7 {}'.format(rows[0]['note'], rows[0]['caption']), progress
        return rows[0]['caption'], progress

    def _last_step(self):
        for step in self._feed:
            if not step.get('kind'):
                return step
        return None

    def _bands(self):
        """(rows, agents left out, the feed step a band already shows) for
        band_rows. An agent with an Allow / Deny card gets no band: the card
        says it."""
        asked = set(item['agent'] for item in self._asks)
        working = [(agent, self._acting.get(agent), work[1])
                   for agent, work in self._working.items() if agent not in self._waiting]
        waiting = [(agent, self._acting.get(agent), caption)
                   for agent, caption in self._waiting.items() if agent not in asked]
        last, replay = self._last_step(), None
        if self._replay is not None and self._replay_step is not None:
            last, replay = self._replay_step, (self._replay[1], len(self._replay[0]))
        rows, more = band_rows(working, waiting, last, replay)
        on_band = last if rows and rows[0]['state'] == 'done' else None
        return rows, more, on_band

    def _fill_full(self, state):
        from System.Windows import Thickness
        from System.Windows.Controls import StackPanel
        full = self._full
        full.Children.Clear()
        full.Width = BAR_W
        agents = self._agents()
        drawn = self._drawn
        drawn['counts_now'], drawn['plan_now'] = {}, {}

        # header
        left = []
        for agent in agents[:4]:
            mark = self._mark(agent_style(agent)[1], 16)
            mark.Margin = Thickness(0, 0, 4, 0)
            left.append(mark)
        title = (agent_style(agents[0])[0] if len(agents) == 1 else
                 '{} agents'.format(len(agents)) if agents else 'Agent Cursors')
        name = self._text(title, 13, color=self._c('text'))
        name.Margin = Thickness(4, 0, 0, 0)
        left.append(name)
        left.append(self._status_pill(state, self._seconds()))
        buttons = self._row(
            self._icon_button(u'' if self._paused else u'',
                              'Resume: let the agents change the model again' if self._paused else
                              'Pause: hold model changes (read-only tools still run)',
                              self._toggle_pause, active=self._paused),
            self._icon_button(u'', 'Options', self._toggle_drawer, active=self._drawer),
            self._icon_button(u'', 'Fold to one line (mouse over opens it)',
                              self._toggle_fold),
            self._icon_button(u'', 'Turn agent cursors off (Settings > Agent Cursors '
                              'turns them back on)', lambda: set_enabled(False)))
        head = self._line_row(self._row(*left), buttons)
        head.Background = self._brush(self._c('bg'), 1)     # hit-testable: drag to move
        head.ToolTip = 'Drag to move the bar · double-click to put it back'
        self._hook(head, 'MouseLeftButtonDown', self._drag, transient=True)
        full.Children.Add(head)

        if self._drawer:
            full.Children.Add(self._options_drawer())

        # goal + progress, per agent with a goal or a plan
        for agent in agents:
            goal = self._goals.get(agent, '')
            progress = plan_progress(self._plans.get(agent))
            if not goal and not progress:
                continue
            label = goal or 'Plan'
            if len(agents) > 1:
                label = u'{} · {}'.format(agent_style(agent)[0], label)
            right = None
            if progress:
                right = self._text(u'Step {}/{}'.format(min(progress[0] + 1, progress[1]),
                                                        progress[1]),
                                   11.5, bold=False, color=self._c('text2'))
            line = self._text(u'Goal · ' + short_text(label, 52), 11.5, bold=False,
                              color=self._c('text2'))
            full.Children.Add(self._line_row(line, right, 10))
            if progress:
                full.Children.Add(self._progress_bar(progress, agent))

        # bands: what runs now, one colour block per agent, else the last step
        rows, more, on_band = self._bands()
        motions, drawn['bands'] = band_motion(rows, drawn.get('bands') or {})
        fill_from = drawn.get('fill', 0.0)
        drawn['fill'] = 0.0
        for k, row in enumerate(rows):
            try:
                band = self._action_band(row, len(agents) > 1, 4 if k else 10, motions[k],
                                         fill_from)
            except Exception:       # one block must never cost the feed and the footer
                band = self._text(row['caption'], 12.5, color=self._c('text'))
                band.Margin = Thickness(0, 10, 0, 0)
            full.Children.Add(band)
        if more:
            extra = self._text(u'+{} more'.format(more), 11, bold=False,
                               color=self._c('muted'))
            extra.Margin = Thickness(BAND_STRIPE + 9, 3, 0, 0)
            full.Children.Add(extra)

        # feed (without the step a band already shows)
        shown = [s for s in self._feed
                 if s.get('kind') != 'goal' and s is not on_band][:FEED_SHOWN]
        top = (shown[0].get('t'), shown[0]['caption']) if shown else None
        if shown:
            feed = StackPanel()
            feed.Margin = Thickness(0, 6, 0, 0)
            for k, step in enumerate(shown):
                row = self._feed_row(step, len(agents) > 1)
                if k == 0 and top != drawn.get('feed'):     # it drops in from the band
                    self._motion(self._slide_in, row, -6.0)
                feed.Children.Add(row)
            full.Children.Add(feed)
        drawn['feed'] = top

        # agents with their counts (several agents, or a summary when idle)
        if len(agents) > 1 or (state == 'idle' and self._counts):
            for agent in agents:
                tag, color = agent_style(agent)
                mark = self._mark(color, 14)
                mark.Margin = Thickness(0, 0, 8, 0)
                row = self._line_row(self._row(mark, self._text(tag, 12.5, bold=False,
                                                                color=self._c('text'))),
                                     self._counts_row(agent, self._counts.get(agent, {})), 6)
                full.Children.Add(row)

        drawn['counts'], drawn['plan'] = drawn.pop('counts_now'), drawn.pop('plan_now')

        # footer
        from System.Windows.Controls import Border
        rule = Border()
        rule.Height = 1
        rule.Background = self._brush(self._c('line'))
        rule.Margin = Thickness(0, 10, 0, 8)
        full.Children.Add(rule)
        note = self._note if self._note else footer_text(self._actions, self._failed,
                                                         len(self._views))
        if self._views and not self._note:
            names = [v[1] for v in self._views[-3:]]
            note_tip = u'Views: ' + u' · '.join(names)
        else:
            note_tip = None
        foot_text = self._text(note, 11.5, bold=False, color=self._c('muted'))
        if note_tip:
            foot_text.ToolTip = note_tip
        actions = self._row(
            self._button(u'', 'Select touched',
                         'Select every element the agents changed (no model change)',
                         self._select_touched),
            self._button(u'', 'Replay', 'Replay the last {} steps'.format(REPLAY_SIZE),
                         self.replay))
        full.Children.Add(self._line_row(foot_text, actions))

    def _progress_bar(self, progress, agent=''):
        from System.Windows import Thickness, CornerRadius, HorizontalAlignment
        from System.Windows.Controls import Border
        done, total = progress
        track = Border()
        track.Height = 4
        track.CornerRadius = CornerRadius(2)
        track.Background = self._brush(self._c('track'))
        track.Margin = Thickness(0, 4, 0, 0)
        fill = Border()
        fill.Height = 4
        fill.CornerRadius = CornerRadius(2)
        fill.Background = self._brush(self._c('accent'))
        fill.HorizontalAlignment = HorizontalAlignment.Left
        fill.Width = fill_width(done, total, BAR_W, 4.0)
        self._motion(self._grow, fill, (self._drawn.get('plan') or {}).get(agent, 0.0), fill.Width)
        self._drawn.setdefault('plan_now', {})[agent] = fill.Width
        track.Child = fill
        return track

    def _action_band(self, row, with_agent, margin_top=10, motion=None, fill_from=0.0):
        """One band_rows row as a block of its action's colour (band_style):
        solid stripe, a wash of the same colour, glyph + caption in the bar's
        own text colour, the action's name on the right. Running: a soft
        block crosses it and the stripe breathes; replay: it fills to k / n
        from `fill_from`. motion (band_motion): the band slides in, its glyph
        pops, or it shakes."""
        from System.Windows import (Thickness, CornerRadius, TextWrapping, GridLength,
                                    GridUnitType, HorizontalAlignment, VerticalAlignment)
        from System.Windows.Controls import Border, Grid, ColumnDefinition, StackPanel
        label, glyph, color = band_style(row['action'], self._theme, row['state'], row['failed'])
        wash = band_wash(self._theme)
        band = Border()
        band.CornerRadius = CornerRadius(4)
        band.Background = self._brush(color, wash)
        band.ClipToBounds = True
        band.Margin = Thickness(0, margin_top, 0, 0)
        layers = Grid()
        band.Child = layers
        if row['progress']:
            done, total = row['progress']
            fill = Border()
            fill.HorizontalAlignment = HorizontalAlignment.Left
            fill.Width = fill_width(done, total, BAR_W)
            fill.CornerRadius = CornerRadius(4) if done >= total else CornerRadius(4, 0, 0, 4)
            fill.Background = self._brush(color, wash)
            self._motion(self._grow, fill, fill_from, fill.Width)
            self._drawn['fill'] = fill.Width
            layers.Children.Add(fill)
        elif row['state'] == 'running':
            self._motion(self._sweep, layers, color, wash)
        stripe = Border()
        stripe.Width = BAND_STRIPE
        stripe.HorizontalAlignment = HorizontalAlignment.Left
        stripe.CornerRadius = CornerRadius(4, 0, 0, 4)
        stripe.Background = self._brush(color)
        if row['state'] in ('running', 'waiting'):
            self._motion(self._breathe, stripe)
        layers.Children.Add(stripe)

        body = StackPanel()
        body.Margin = Thickness(BAND_STRIPE + 9, 6, 10, 7)
        layers.Children.Add(body)
        head = Grid()
        for width in (GridLength.Auto, GridLength(1, GridUnitType.Star), GridLength.Auto):
            cd = ColumnDefinition()
            cd.Width = width
            head.ColumnDefinitions.Add(cd)
        lead = []
        if with_agent and row['agent']:
            mark = self._mark(agent_style(row['agent'])[1], 13)
            mark.Margin = Thickness(0, 0, 6, 0)
            mark.ToolTip = agent_style(row['agent'])[0]
            lead.append(mark)
        icon = self._text(glyph, 12, color=color, glyph=True)
        icon.Margin = Thickness(0, 0, 8, 0)
        lead.append(icon)
        left = self._row(*lead)
        left.VerticalAlignment = VerticalAlignment.Top
        left.Margin = Thickness(0, 2, 0, 0)
        head.Children.Add(left)
        text = self._text(row['caption'], 12.5, color=self._c('text'))
        text.TextWrapping = TextWrapping.Wrap
        Grid.SetColumn(text, 1)
        head.Children.Add(text)
        note = self._text(row['note'], 11, bold=False, color=self._c('text2'))
        note.Margin = Thickness(10, 1, 0, 0)
        note.VerticalAlignment = VerticalAlignment.Top
        Grid.SetColumn(note, 2)
        head.Children.Add(note)
        body.Children.Add(head)
        if row['state'] in ('done', 'replay') and self._last_changes:
            from System.Windows.Controls import WrapPanel
            from System.Windows.Media import FontFamily
            chips = WrapPanel()
            chips.Margin = Thickness(0, 4, 0, 0)
            for change in self._last_changes[:3]:
                chip = self._pill(self._text(short_text(change, 40), 11, bold=False,
                                             color=self._c('text')), self._c('chip'), radius=2)
                chip.Child.FontFamily = FontFamily('Consolas')
                chip.Margin = Thickness(0, 0, 6, 2)
                chips.Children.Add(chip)
            body.Children.Add(chips)
        if motion == 'enter':
            self._motion(self._slide_in, band)
        elif motion == 'shake':
            self._motion(self._shake, band)
        elif motion == 'pop':
            self._motion(self._pop, icon, 6.0, 6.0)
        return band

    def _sweep(self, layers, color, wash):
        """A soft block of the band's colour that crosses it while the call
        runs; one pass takes 1.4 s at Normal."""
        from System.Windows import HorizontalAlignment, Point
        from System.Windows.Controls import Border
        from System.Windows.Media import (TranslateTransform, LinearGradientBrush, GradientStop,
                                          Color)
        block = Border()
        block.Width = SWEEP_W
        block.HorizontalAlignment = HorizontalAlignment.Left
        block.IsHitTestVisible = False
        try:
            h = color.lstrip('#')
            r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
            soft = LinearGradientBrush()
            soft.StartPoint, soft.EndPoint = Point(0.0, 0.5), Point(1.0, 0.5)
            for alpha, offset in ((0, 0.0), (min(255, wash * 2), 0.5), (0, 1.0)):
                soft.GradientStops.Add(GradientStop(Color.FromArgb(alpha, r, g, b), offset))
            block.Background = soft
        except Exception:
            block.Background = self._brush(color, wash)     # a plain block still reads
        tt = TranslateTransform()
        tt.X = -SWEEP_W
        block.RenderTransform = tt
        layers.Children.Add(block)
        self._loop(tt, _translate_x(), ((-SWEEP_W, 0), (BAR_W, int(1400 * speed_factor()))))

    def _feed_row(self, step, with_agent):
        """Stripe in the action's colour · ✓ / ✕ / ⚠ · [agent] caption · how
        long ago. The stripes line up under the band's: one colour column."""
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border
        if step.get('kind') in ('warn', 'note'):
            glyph, tint = u'', self._c('warn')
        elif step['failed']:
            glyph, tint = u'', self._c('bad')
        else:
            glyph, tint = u'', self._c('ok')
        stripe = Border()
        stripe.Width, stripe.Height = 3, 12
        stripe.CornerRadius = CornerRadius(1.5)
        stripe.Background = self._brush(tint if step.get('kind') else
                                        self._tone(step.get('action')))
        stripe.Margin = Thickness(1, 0, BAND_STRIPE + 5, 0)
        if not step.get('kind'):
            stripe.ToolTip = action_style(step.get('action'))[0]
        icon = self._text(glyph, 10, color=tint, glyph=True)
        icon.Margin = Thickness(0, 0, 8, 0)
        parts = [stripe, icon]
        if with_agent:
            mark = self._mark(agent_style(step['agent'])[1], 12)
            mark.Margin = Thickness(0, 0, 6, 0)
            parts.append(mark)
        parts.append(self._text(short_text(step['caption'], 50), 12, bold=False,
                                color=self._c('text2')))
        ago = self._text(ago_text(time.time() - step.get('t', time.time())), 11, bold=False,
                         color=self._c('faint'))
        row = self._line_row(self._row(*parts), ago, 3)
        row.ToolTip = step.get('tool') or ''
        return row

    def _options_drawer(self):
        """⚙: the options as switches in groups and what each icon means."""
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border, StackPanel
        box = Border()
        box.Background = self._brush(self._c('card'))
        box.CornerRadius = CornerRadius(4)
        box.Padding = Thickness(10, 4, 10, 8)
        box.Margin = Thickness(0, 10, 0, 0)
        col = StackPanel()
        box.Child = col
        opts = options()
        for group, names in OPTION_GROUPS:
            head = self._text(group.upper(), 10.5, bold=False, color=self._c('faint'))
            head.Margin = Thickness(0, 6, 0, 2)
            col.Children.Add(head)
            for name in names:
                label, tip = OPTION_LABELS[name][:2]

                def flip(s, e, name=name):
                    try:
                        set_option(name, not options()[name])
                        self._refresh_hud()
                    except Exception:
                        pass
                text = self._text(label, 12.5, bold=False, color=self._c('text'))
                text.ToolTip = tip
                col.Children.Add(self._line_row(text, self._switch(opts[name], flip), 3))
        head = self._text('ACTIONS', 10.5, bold=False, color=self._c('faint'))
        head.Margin = Thickness(0, 6, 0, 2)
        col.Children.Add(head)
        col.Children.Add(self._legend())
        return box

    def _legend(self):
        """What each action's icon and colour mean: icon + name, no swatch
        (owner, 2026-11-19)."""
        from System.Windows import Thickness
        from System.Windows.Controls import WrapPanel
        wrap = WrapPanel()
        for key, label, glyph, _color in ACTIONS:
            icon = self._text(glyph, 11.5, color=self._tone(key), glyph=True)
            icon.Margin = Thickness(0, 0, 4, 0)
            item = self._row(icon, self._text(label, 11.5, bold=False, color=self._c('text2')))
            item.Margin = Thickness(0, 2, 12, 0)
            wrap.Children.Add(item)
        return wrap

    def _toggle_drawer(self):
        self._drawer = not self._drawer

    def _toggle_pause(self):
        fn = (self.host.callbacks if self.host is not None else {}).get('pause')
        if fn is not None:
            fn(not self._paused)        # agent_cursors.set_paused -> set_paused() here
        else:
            self._paused = not self._paused

    # position: bottom centre of the active view, moved by the user's drag

    def _place_hud(self):
        if self._view_dip is None or not self.hud.IsVisible:
            return
        self.hud.UpdateLayout()
        size = (self.hud.ActualWidth or 420, self.hud.ActualHeight or 120)
        try:
            offset = hud_offset()
        except Exception:
            offset = (0.0, 0.0)
        self.hud.Left, self.hud.Top = hud_place(self._view_dip, size, offset)

    def _drag(self, sender, e):
        try:
            if e.ClickCount == 2:
                set_hud_offset(0, 0)
                self._place_hud()
                return
            self.hud.DragMove()
            if self._view_dip is None:
                return
            size = (self.hud.ActualWidth or 420, self.hud.ActualHeight or 120)
            home = hud_place(self._view_dip, size, (0.0, 0.0))
            set_hud_offset(self.hud.Left - home[0], self.hud.Top - home[1])
        except Exception:
            pass

    def held(self, agent, caption):
        """A change refused while paused (agent_cursors.paused_refusal)."""
        self._feed.insert(0, {'agent': agent, 'action': 'modify', 'tool': '', 'failed': True,
                              'kind': 'warn', 't': time.time(),
                              'caption': u'Held while paused: {}'.format(caption)})
        del self._feed[FEED_SIZE:]
        self._last = time.time()
        if self.hud.IsVisible:
            self._refresh_hud()

    def notice(self, text):
        self._feed.insert(0, {'agent': '', 'action': 'read', 'tool': '', 'failed': False,
                              'kind': 'note', 't': time.time(), 'caption': text})
        del self._feed[FEED_SIZE:]
        if self.hud.IsVisible:
            self._refresh_hud()

    # ── cursors ──────────────────────────────────────────────────────────────

    def _cursor(self, agent, action=None):
        """Minimal cursor (owner, 2026-11-17): a bare pointer tip in the
        agent's colour and one tag, "Claude · Moving 12 walls". One per
        agent: the action only changes the icon at the head of the tag (two
        cursors of one agent sat on top of each other, 2026-11-17)."""
        cur = self._cursors.get(agent)
        if cur is not None:
            if action:
                self._accent(cur, action)
            return cur
        from System.Windows import Thickness, CornerRadius, Visibility
        from System.Windows.Controls import Canvas, Border
        from System.Windows.Media import TranslateTransform, ScaleTransform, Geometry
        from System.Windows.Media.Effects import DropShadowEffect
        from System.Windows.Shapes import Path, Ellipse
        tag, color = agent_style(agent)

        group = Canvas()
        arrow = Path()
        arrow.Data = Geometry.Parse(CURSOR_PATH)
        arrow.Fill = self._brush(color)            # no outline (owner, 2026-11-17)
        shadow = DropShadowEffect()
        shadow.BlurRadius, shadow.ShadowDepth, shadow.Opacity = 2, 0.5, 0.3
        arrow.Effect = shadow
        press = ScaleTransform()            # a "click" squeezes the arrow to its tip
        arrow.RenderTransform = press
        group.Children.Add(arrow)

        spin = Ellipse()                    # pulses in the tag while the call runs
        spin.Width = spin.Height = 6
        spin.Fill = self._brush('#FFFFFF')
        spin.Margin = Thickness(0, 0, 5, 0)
        spin.Visibility = Visibility.Collapsed
        glyph = self._text('', 10, color='#FFFFFF', glyph=True)     # the action (_accent)
        glyph.Margin = Thickness(0, 0, 5, 0)
        glyph.Visibility = Visibility.Collapsed
        name = self._text(tag, 11, color='#FFFFFF')
        state = self._text('', 11, bold=False, color='#FFFFFF')     # " · what it does"
        pill = Border()
        pill.CornerRadius = CornerRadius(6)
        pill.Padding = Thickness(7, 1, 7, 2)
        pill.Background = self._brush(color)
        pill.Child = self._row(spin, glyph, name, state)
        Canvas.SetLeft(pill, 12)
        Canvas.SetTop(pill, 12)
        group.Children.Add(pill)

        tx = TranslateTransform()
        group.RenderTransform = tx
        self.canvas.Children.Add(group)
        # A new cursor starts mid-view (a little apart per agent), not in a corner.
        slot = len(self._cursors)
        x, y = 40.0 + 36 * slot, 40.0 + 24 * slot
        bounds = self._bounds()
        if bounds is not None:
            x = (bounds[0] + bounds[2]) / 2.0 + 36 * slot
            y = (bounds[1] + bounds[3]) / 2.0 + 24 * slot
        cur = {'group': group, 'tx': tx, 'press': press, 'spin': spin, 'pill': pill,
               'glyph': glyph, 'name': name, 'state': state, 'color': color,
               'caption': '', 'x': x, 'y': y, 'action': None, 'new': True}
        tx.X, tx.Y = cur['x'], cur['y']
        self._cursors[agent] = cur
        if action:
            self._accent(cur, action)
        return cur

    def _wake(self, cur):
        """Full opacity; a cursor shown for the first time fades in and its
        tag grows out of the tip."""
        cur['group'].BeginAnimation(_opacity_property(), None)
        cur['group'].Opacity = 1.0
        if cur.pop('new', False):
            self._motion(self._fade_in, cur['group'], 220)
            self._motion(self._pop, cur['pill'], 0.0, 0.0, 0.6, 240)

    def _accent(self, cur, action):
        """The action shows as its icon at the head of the agent's tag, in
        the tag's own ink: no block, no colour (owner, 2026-11-19). The old
        coloured edge did not read (amber on Claude's terracotta). The icon
        pops when the action changes."""
        from System.Windows import Visibility
        if cur.get('action') == action:
            return
        changed = cur.get('action') is not None
        cur['action'] = action
        cur['glyph'].Text = action_style(action)[1]
        cur['glyph'].Visibility = Visibility.Visible
        if changed:
            self._motion(self._pop, cur['glyph'], 5.0, 6.0, 1.6, 240)

    def _bounds(self):
        """The canvas in its own DIP (l, t, r, b), or None before it is placed."""
        try:
            w, h = float(self.win.Width), float(self.win.Height)
        except Exception:
            return None
        if not (w > 80 and h > 80):         # also False for NaN
            return None
        return (0.0, 0.0, w, h)

    def _layout_tags(self, lead=None):
        """Put every tag where it is readable: inside the view (flipped left
        of / above its pointer near an edge) and never over another tag. The
        `lead` agent (the one that just moved) picks first."""
        from System.Windows import Size
        from System.Windows.Controls import Canvas
        keys = sorted(self._cursors, key=lambda k: (k != lead, k))
        inf = float('inf')
        items = []
        for key in keys:
            cur = self._cursors[key]
            try:
                cur['pill'].Measure(Size(inf, inf))
                size = cur['pill'].DesiredSize
                w, h = float(size.Width), float(size.Height)
            except Exception:
                w, h = 160.0, 20.0
            items.append((cur['x'], cur['y'], w, h))
        for key, (dx, dy) in zip(keys, place_tags(items, self._bounds())):
            Canvas.SetLeft(self._cursors[key]['pill'], dx)
            Canvas.SetTop(self._cursors[key]['pill'], dy)

    def _cursor_state(self, cur, state, caption=None):
        """idle | working (spinner) | waiting (amber, pulsing) | failed (red);
        `caption` (what it does) follows the name, None keeps the current one."""
        from System.Windows import Visibility
        pill, spin = cur['pill'], cur['spin']
        pill.BeginAnimation(_opacity_property(), None)
        pill.Opacity = 1.0
        spin.BeginAnimation(_opacity_property(), None)
        spin.Visibility = Visibility.Collapsed
        pill.Background = self._brush({'waiting': '#F59E0B', 'failed': '#DC2626'}.get(
            state, cur['color']))
        ink = '#1F2937' if state == 'waiting' else '#FFFFFF'
        for part in ('glyph', 'name', 'state'):
            cur[part].Foreground = self._brush(ink)
        if state == 'waiting':
            caption = u'Waiting for your OK'
        if caption is not None:
            cur['caption'] = caption
        text = cur.get('caption') or ''
        cur['state'].Text = u' \u00b7 ' + short_text(text, TAG_MAX) if text else ''
        if state == 'working':
            spin.Visibility = Visibility.Visible
            if not self._reduce():
                from System.Windows.Media.Animation import RepeatBehavior
                anim = self._anim(((1.0, 0), (0.25, 450), (1.0, 900)))
                anim.RepeatBehavior = RepeatBehavior.Forever
                spin.BeginAnimation(_opacity_property(), anim)
        elif state == 'waiting' and not self._reduce():
            from System.Windows.Media.Animation import RepeatBehavior
            anim = self._anim(((1.0, 0), (0.5, 500), (1.0, 1000)))
            anim.RepeatBehavior = RepeatBehavior.Forever
            pill.BeginAnimation(_opacity_property(), anim)

    def _tour(self, cur, sched, shake=False):
        """Glide through sched [(x, y, arrive_ms, leave_ms)] on one keyframe
        path per axis (a second BeginAnimation would cancel the first); a
        failed call ends with a shake."""
        from System import TimeSpan
        from System.Windows.Media import TranslateTransform
        from System.Windows.Media.Animation import (
            DoubleAnimationUsingKeyFrames, EasingDoubleKeyFrame, LinearDoubleKeyFrame,
            KeyTime, CubicEase, EasingMode)
        ease = CubicEase()
        ease.EasingMode = EasingMode.EaseInOut

        def at(ms):
            return KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(ms))
        try:
            start = (float(cur['tx'].X), float(cur['tx'].Y))   # where it is mid-tour
        except Exception:
            start = (cur['x'], cur['y'])
        for axis, prop in ((0, TranslateTransform.XProperty), (1, TranslateTransform.YProperty)):
            anim = DoubleAnimationUsingKeyFrames()
            anim.KeyFrames.Add(LinearDoubleKeyFrame(start[axis], at(0)))
            for stop in sched:
                anim.KeyFrames.Add(EasingDoubleKeyFrame(stop[axis], at(stop[2]), ease))
                anim.KeyFrames.Add(LinearDoubleKeyFrame(stop[axis], at(stop[3])))
            if shake and axis == 0:
                end, t = sched[-1][0], sched[-1][3]
                for dx, dt in ((-6, 60), (6, 120), (-4, 180), (0, 240)):
                    anim.KeyFrames.Add(LinearDoubleKeyFrame(end + dx, at(t + dt)))
            cur['tx'].BeginAnimation(prop, anim)
        cur['x'], cur['y'] = sched[-1][0], sched[-1][1]

    def _press(self, cur, times):
        """Squeeze the arrow at each arrival: the agent clicks."""
        from System.Windows.Media import ScaleTransform
        frames = [(1.0, 0)]
        for t in times:
            frames += [(1.0, max(t - 1, frames[-1][1] + 1)), (0.78, t + 70), (1.0, t + 170)]
        for prop in (ScaleTransform.ScaleXProperty, ScaleTransform.ScaleYProperty):
            cur['press'].BeginAnimation(prop, self._anim(frames))

    # ── effects ──────────────────────────────────────────────────────────────

    def _trail(self, x0, y0, x1, y1, color):
        """Dashed path from where the cursor was to where it goes."""
        from System.Windows.Media import DoubleCollection
        from System.Windows.Shapes import Line
        if abs(x1 - x0) + abs(y1 - y0) < 24:
            return
        line = Line()
        line.X1, line.Y1, line.X2, line.Y2 = x0, y0, x1, y1
        line.Stroke = self._brush(color, 170)
        line.StrokeThickness = 1.5
        dash = DoubleCollection()
        dash.Add(3)
        dash.Add(3)
        line.StrokeDashArray = dash
        line.Opacity = 0
        self.canvas.Children.Insert(0, line)
        line.BeginAnimation(_opacity_property(), self._anim(
            ((0.0, 0), (0.9, 120), (0.9, MOVE_MS), (0.0, MOVE_MS + 700)),
            self._remover(line)))

    def _ripple(self, x, y, color, delay_ms, size=36):
        """Click ring where the cursor lands."""
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Ellipse
        ring = Ellipse()
        ring.Width = ring.Height = size
        ring.Stroke = self._brush(color)
        ring.StrokeThickness = 2
        ring.Fill = self._brush(color, 40)
        ring.Opacity = 0
        Canvas.SetLeft(ring, x - size / 2.0)
        Canvas.SetTop(ring, y - size / 2.0)
        self.canvas.Children.Add(ring)
        self._scale_anim(ring, size / 2.0, size / 2.0,
                         ((0.1, delay_ms), (0.1, delay_ms), (1.0, delay_ms + 520)))
        ring.BeginAnimation(_opacity_property(), self._anim(
            ((0.0, delay_ms), (0.9, delay_ms + 40), (0.0, delay_ms + 560)),
            self._remover(ring)))

    def _frame(self, x0, y0, x1, y1, color, dashed=False, thickness=1.0, fill_alpha=18):
        from System.Windows.Controls import Canvas
        from System.Windows.Media import DoubleCollection
        from System.Windows.Shapes import Rectangle
        frame = Rectangle()
        frame.Width = max(x1 - x0, 2)
        frame.Height = max(y1 - y0, 2)
        frame.Stroke = self._brush(color)
        frame.StrokeThickness = thickness
        frame.Fill = self._brush(color, fill_alpha)
        if dashed:
            dash = DoubleCollection()
            dash.Add(4)
            dash.Add(3)
            frame.StrokeDashArray = dash
        Canvas.SetLeft(frame, x0)
        Canvas.SetTop(frame, y0)
        return frame

    def _draw_focus(self, agent, b, color):
        """What a running call names: a soft frame that breathes until the
        call ends (_clear_focus), so the work shows on the model meanwhile."""
        frame = self._frame(b[0] - 2, b[1] - 2, b[2] + 2, b[3] + 2, color,
                            thickness=1.5, fill_alpha=28)
        self.canvas.Children.Insert(0, frame)
        self._focus.setdefault(agent, []).append(frame)
        if not self._reduce():
            from System.Windows.Media.Animation import RepeatBehavior
            anim = self._anim(((1.0, 0), (0.45, 700), (1.0, 1400)))
            anim.RepeatBehavior = RepeatBehavior.Forever
            frame.BeginAnimation(_opacity_property(), anim)

    def _clear_focus(self, agent):
        for frame in self._focus.pop(agent, ()):
            try:
                frame.BeginAnimation(_opacity_property(), None)
                self.canvas.Children.Remove(frame)
            except Exception:
                pass

    def _draw_label(self, b, text, action, delay_ms, hold_ms):
        """'Walls ×3 created' above what the call made or changed."""
        from System.Windows import Thickness
        from System.Windows.Controls import Canvas
        _label, glyph, color = action_style(action)
        chip = self._pill(self._glyph_label(glyph, text, '#FFFFFF', 11), color)
        chip.Padding = Thickness(8, 3, 8, 3)
        chip.Opacity = 0
        Canvas.SetLeft(chip, max(b[0], 4))
        Canvas.SetTop(chip, b[1] - 26 if b[1] >= 30 else b[3] + 26)
        self.canvas.Children.Add(chip)
        self._pulse(chip, delay_ms, hold_ms, self._remover(chip))

    def _build_trail(self, sched, color, hold_ms):
        """Several new elements in a row: a dotted line joins them in the
        order they were made, one segment per arrival."""
        from System.Windows.Media import DoubleCollection
        from System.Windows.Shapes import Line
        end = sched[-1][2]
        for k in range(1, len(sched)):
            (x0, y0), (x1, y1), at = sched[k - 1][:2], sched[k][:2], sched[k][2]
            if abs(x1 - x0) + abs(y1 - y0) < 12:
                continue
            line = Line()
            line.X1, line.Y1, line.X2, line.Y2 = x0, y0, x1, y1
            line.Stroke = self._brush(color, 200)
            line.StrokeThickness = 2
            dash = DoubleCollection()
            dash.Add(1)
            dash.Add(3)
            line.StrokeDashArray = dash
            line.Opacity = 0
            self.canvas.Children.Insert(0, line)
            line.BeginAnimation(_opacity_property(), self._anim(
                ((0.0, at - 120), (1.0, at), (1.0, end + hold_ms), (0.0, end + hold_ms + 600)),
                self._remover(line)))

    def _draw_box(self, x0, y0, x1, y1, size, action, delay_ms, fresh=False):
        """One target box with the action's effect; `size` is the badge text
        (category · size). fresh = just created / changed: a soft highlight
        that stays FRESH_HOLD_S."""
        from System.Windows import Thickness
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Rectangle, Shape
        from System.Windows.Media import TranslateTransform
        label, glyph, color = action_style(action)
        group = Canvas()
        group.Opacity = 0
        frame = self._frame(x0, y0, x1, y1, color, dashed=(action == 'code' and not fresh),
                            thickness=1.5 if fresh else 1.0, fill_alpha=52 if fresh else 18)
        group.Children.Add(frame)
        for hx, hy in ((x0, y0), (x1, y0), (x0, y1), (x1, y1)):
            handle = Rectangle()
            handle.Width = handle.Height = 4
            handle.Fill = self._brush(color)       # solid corner, no white (owner)
            Canvas.SetLeft(handle, hx - 2)
            Canvas.SetTop(handle, hy - 2)
            group.Children.Add(handle)
        if self._fx.get('labels', True):
            badge = self._pill(self._glyph_label(glyph, size, '#FFFFFF'), color, radius=2)
            badge.Padding = Thickness(4, 1, 4, 1)
            Canvas.SetLeft(badge, max(x0, (x0 + x1) / 2.0 - 60))
            Canvas.SetTop(badge, y1 + 4)
            group.Children.Add(badge)
        self.canvas.Children.Insert(0, group)    # under the cursors
        hold = int((FRESH_HOLD_S if fresh else BOX_HOLD_S) * 1000 * speed_factor())
        fx = None if self._reduce() else ('create' if fresh and action == 'code' else action)
        start = delay_ms + 180

        if fx == 'create':
            # Grows out of its centre.
            self._scale_anim(group, (x0 + x1) / 2.0, (y0 + y1) / 2.0,
                             ((0.4, delay_ms), (0.4, delay_ms),
                              (1.08, start + 160), (1.0, start + 260)))
        elif fx == 'modify':
            frame.BeginAnimation(Shape.StrokeThicknessProperty, self._anim(
                ((1.0, start), (3.0, start + 200), (1.0, start + 400),
                 (3.0, start + 600), (1.0, start + 800))))
        elif fx == 'read':
            line = Rectangle()
            line.Width = max(x1 - x0, 2)
            line.Height = 2
            line.Fill = self._brush(color, 200)
            Canvas.SetLeft(line, x0)
            Canvas.SetTop(line, y0)
            tt = TranslateTransform()
            line.RenderTransform = tt
            group.Children.Add(line)
            h = max(y1 - y0 - 2, 0)
            tt.BeginAnimation(TranslateTransform.YProperty, self._anim(
                ((0.0, start), (h, start + 700), (0.0, start + 1400))))
        elif fx == 'code':
            frame.BeginAnimation(Shape.StrokeDashOffsetProperty, self._anim(
                ((0.0, start), (-42.0, start + hold))))
        elif fx == 'export':
            up = self._pill(self._text('', 11, glyph=True), color, radius=9)
            Canvas.SetLeft(up, x1 - 10)
            Canvas.SetTop(up, y0 - 10)
            tt = TranslateTransform()
            up.RenderTransform = tt
            group.Children.Add(up)
            tt.BeginAnimation(TranslateTransform.YProperty, self._anim(
                ((0.0, start), (-28.0, start + 900))))
        self._pulse(group, delay_ms, hold, self._remover(group))

    def _draw_gone(self, x0, y0, x1, y1, delay_ms):
        """A deleted element: red frame with a cross that shrinks away."""
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Line
        _l, _g, color = action_style('delete')
        group = Canvas()
        group.Opacity = 0
        group.Children.Add(self._frame(x0, y0, x1, y1, color, dashed=True, fill_alpha=40))
        for ax, ay, bx, by in ((x0, y0, x1, y1), (x0, y1, x1, y0)):
            line = Line()
            line.X1, line.Y1, line.X2, line.Y2 = ax, ay, bx, by
            line.Stroke = self._brush(color, 200)
            line.StrokeThickness = 1.5
            group.Children.Add(line)
        self.canvas.Children.Insert(0, group)
        hold = 900
        self._scale_anim(group, (x0 + x1) / 2.0, (y0 + y1) / 2.0,
                         ((1.0, delay_ms + 180 + hold), (0.15, delay_ms + 780 + hold)))
        self._pulse(group, delay_ms, hold, self._remover(group))

    def _draw_ghost(self, old, new, color, delay_ms):
        """Where a moved element was, with an arrow to where it is now."""
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Line
        group = Canvas()
        group.Opacity = 0
        group.Children.Add(self._frame(old[0], old[1], old[2], old[3], color,
                                       dashed=True, fill_alpha=0))
        line = Line()
        line.X1, line.Y1 = (old[0] + old[2]) / 2.0, (old[1] + old[3]) / 2.0
        line.X2, line.Y2 = (new[0] + new[2]) / 2.0, (new[1] + new[3]) / 2.0
        line.Stroke = self._brush(color, 190)
        line.StrokeThickness = 1.5
        group.Children.Add(line)
        head = self._ripple_dot(line.X2, line.Y2, color)
        group.Children.Add(head)
        self.canvas.Children.Insert(0, group)
        self._pulse(group, delay_ms, int(BOX_HOLD_S * 1000), self._remover(group))

    def _ripple_dot(self, x, y, color):
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Ellipse
        dot = Ellipse()
        dot.Width = dot.Height = 7
        dot.Fill = self._brush(color)
        Canvas.SetLeft(dot, x - 3.5)
        Canvas.SetTop(dot, y - 3.5)
        return dot

    def _flash_view(self, geo, text, color):
        """Frame the whole view and name it: the agent opened / switched to it."""
        from System.Windows import Thickness
        from System.Windows.Controls import Canvas
        x0, y0, x1, y1 = self._local(geo['rect'])
        group = Canvas()
        group.Opacity = 0
        group.Children.Add(self._frame(x0 + 2, y0 + 2, x1 - 2, y1 - 2, color,
                                       thickness=3, fill_alpha=0))
        chip = self._pill(self._glyph_label('', text, '#FFFFFF', 11), color)
        chip.Padding = Thickness(8, 4, 8, 4)
        Canvas.SetLeft(chip, x0 + 12)
        Canvas.SetTop(chip, y0 + 12)
        group.Children.Add(chip)
        self.canvas.Children.Insert(0, group)
        self._pulse(group, 0, 1600, self._remover(group))

    # ── public ────────────────────────────────────────────────────────────────

    # ── states other than a finished call ─────────────────────────────────────

    def _place(self, rect):
        """Cover `rect` (screen px) with the canvas; shift cursors to match."""
        left, top, right, bottom = rect
        self._canvas_px = (left, top, right, bottom)
        if not self.win.IsVisible:
            self._wanted = True
            # Shown only when nothing covers Revit: showing it anyway let it
            # flash up on every call until the next front tick hid it again.
            if self._front_ok:
                self.win.Width = self.win.Height = 1
                self.win.Show()
        scale = self._dpi()
        if self._origin != (left, top):
            dx = (self._origin[0] - left) / scale
            dy = (self._origin[1] - top) / scale
            self._clear_effects()
            for cur in self._cursors.values():
                # From where it is now (mid-glide too), not from where its
                # last tour ends: that made the cursor jump ahead.
                try:
                    x, y = float(cur['tx'].X), float(cur['tx'].Y)
                except Exception:
                    x, y = cur['x'], cur['y']
                cur['x'], cur['y'] = x + dx, y + dy
                cur['tx'].BeginAnimation(_translate_x(), None)
                cur['tx'].BeginAnimation(_translate_y(), None)
                cur['tx'].X, cur['tx'].Y = cur['x'], cur['y']
        self._origin = (left, top)
        self.win.Left, self.win.Top = left / scale, top / scale
        self.win.Width = (right - left) / scale
        self.win.Height = (bottom - top) / scale

    def _clear_effects(self):
        """The views moved (tile, follow, resize): boxes, frames, trails and
        arrows were drawn for the old layout and would sit beside their
        elements. Drop them; the cursors stay (_place shifts them)."""
        try:
            from System import Object
            groups = [cur['group'] for cur in self._cursors.values()]
            for child in list(self.canvas.Children):
                if not any(Object.ReferenceEquals(child, g) for g in groups):
                    self.canvas.Children.Remove(child)
        except Exception:
            pass
        self._focus = {}
        self._marks = []

    def _show_hud(self, rect):
        """rect = the active view (screen px); None = a call with no view to
        draw in: the bar stays where it last was, or stays away if it never
        had a place."""
        if rect is not None:
            scale = self._dpi()
            self._view_dip = (rect[0] / scale, rect[1] / scale, rect[2] / scale, rect[3] / scale)
            if self._canvas_px is None or not self.win.IsVisible:
                self._canvas_px = tuple(rect)   # bar only: the view the bar sits in
        if self._view_dip is None:
            return
        if not self.hud.IsVisible and (self._front_ok or self._asks):
            self.hud.Show()
        self._refresh_hud()             # places it too (_place_hud)
        self._timer.Start()

    def working(self, agent, action, caption, tool_name, rect, task_text='',
                targets=(), canvas=None):
        """A call has started: spin, say "Working", count the seconds. With a
        plan, the cursor names the current task and the call goes underneath.
        targets = screen boxes of the elements the call names: framed until
        it ends, the cursor goes to the first. None: the cursor stays where
        it last worked. rect None = no view to draw in: the bar only."""
        self._end_summary()
        self._stop_replay()             # a live call wins over a replay
        if task_text:
            caption, tool_name = task_text, u'{} · {}'.format(caption, tool_name)
        self._working[agent] = (time.time(), caption)
        self._acting[agent] = action
        if not self._model():
            self._bar_only()            # the default: the bar, nothing on the model
        elif rect is not None:
            try:
                self._draw_working(agent, action, caption, rect, targets, canvas)
            except Exception:
                pass            # the drawing is decoration: the bar still says Working
        self._last = time.time()
        self._show_hud(rect)

    def _draw_working(self, agent, action, caption, rect, targets, canvas):
        """The running call on the model: its cursor, the frames round what it
        names, the path to the first."""
        if canvas or not self.win.IsVisible or self._origin == (0.0, 0.0):
            self._place(canvas or rect)
        cur = self._cursor(agent, action)
        self._cursor_state(cur, 'working', caption)
        self._wake(cur)
        self._clear_focus(agent)
        if targets:
            acolor = action_style(action)[2]
            boxes = [self._local(b) for b in targets]
            for b in boxes:
                self._draw_focus(agent, b, acolor)
            point = ((boxes[0][0] + boxes[0][2]) / 2.0, (boxes[0][1] + boxes[0][3]) / 2.0)
            reduce = self._reduce()
            if not reduce and self._effects()['trail']:
                self._trail(cur['x'], cur['y'], point[0], point[1], acolor)
            self._tour(cur, tour_schedule([point], 1 if reduce else int(hop_ms(1) * speed_factor()), 0))
        self._layout_tags(agent)

    def waiting(self, agent, tool_name, caption, on):
        """The call waits for the user's OK: an amber, pulsing pill."""
        cur = self._cursors.get(agent)
        if on:
            self._waiting[agent] = caption
            self._acting.setdefault(agent, classify_action(tool_name))
            if cur is None and self.win.IsVisible and self._model():
                cur = self._cursor(agent, classify_action(tool_name))
                self._wake(cur)
            if cur is not None:
                self._cursor_state(cur, 'waiting')
        else:
            self._waiting.pop(agent, None)
            if agent not in self._working:
                self._acting.pop(agent, None)
            if cur is not None:
                work = self._working.get(agent)
                self._cursor_state(cur, 'working' if work else 'idle', work[1] if work else None)
        if cur is not None:
            self._layout_tags(agent)
        self._last = time.time()
        if self.hud.IsVisible:
            self._refresh_hud()

    def narrate(self, agent, goal, plan=None):
        """t3lab_narrate: the agent's goal and task list on the HUD + a feed line
        when the goal changes."""
        changed = goal != self._goals.get(agent)
        self._goals[agent] = goal
        if plan and plan.get('tasks'):
            self._plans[agent] = plan
        if goal and changed:
            self._feed.insert(0, {'agent': agent, 'action': 'read', 'caption': u'Goal: ' + goal,
                                  'tool': 't3lab_narrate', 'failed': False, 'kind': 'goal',
                                  't': time.time()})
            del self._feed[FEED_SIZE:]
        self._last = time.time()
        if self.hud.IsVisible:
            self._refresh_hud()

    # ── consent on the bar (agent_cursors.ask) ───────────────────────────────

    def ask(self, rid, agent, tool_name, caption, main, content, expanded, always_text, rect):
        """A call waits for the user's OK: a card with Allow / Deny on the bar
        and the cursor pulsing, instead of a modal Revit dialog."""
        self._end_summary()
        if self._model():
            self._place(rect)
        else:
            self._bar_only()
        self._asks.append({'rid': rid, 'agent': agent, 'tool': tool_name, 'caption': caption,
                           'main': main, 'content': content, 'expanded': expanded,
                           'always_text': always_text})
        self.waiting(agent, tool_name, caption, True)
        self._render_asks()
        self._show_hud(rect)
        try:
            self.hud.Topmost = True         # seen even while the AI app is in front
        except Exception:
            pass

    def _answer(self, rid, approved, always=False):
        item = next((a for a in self._asks if a['rid'] == rid), None)
        if item is None:
            return
        self._asks.remove(item)
        self.waiting(item['agent'], item['tool'], item['caption'], False)
        self._render_asks()
        fn = (self.host.callbacks if self.host is not None else {}).get('answer')
        if fn is not None:
            fn(rid, approved, always)
        self._last = time.time()

    def cancel_ask(self, rid):
        item = next((a for a in self._asks if a['rid'] == rid), None)
        if item is not None:
            self._asks.remove(item)
            self.waiting(item['agent'], item['tool'], item['caption'], False)
            self._render_asks()

    def _render_asks(self):
        from System.Windows import (Thickness, CornerRadius, TextWrapping, Visibility,
                                    HorizontalAlignment)
        from System.Windows.Controls import Border, StackPanel, CheckBox, ScrollViewer
        from System.Windows.Input import Cursors
        from System.Windows.Media import FontFamily
        self._ask_panel.Children.Clear()
        for item in self._asks:
            tag, color = agent_style(item['agent'])
            card = Border()
            card.Background = self._brush(self._c('card'))
            card.BorderBrush = self._brush(self._c('warn'))
            card.BorderThickness = Thickness(1)
            card.CornerRadius = CornerRadius(6)
            card.Padding = Thickness(10, 8, 10, 8)
            card.Margin = Thickness(0, 0, 0, 8)
            card.Width = BAR_W
            body = StackPanel()
            card.Child = body
            head = self._glyph_label('\uE7BA', u'{} asks for your OK'.format(tag), self._c('warn'), 11)
            body.Children.Add(head)
            title = self._text(item['main'] or item['caption'], 12, color=self._c('text'))
            title.TextWrapping = TextWrapping.Wrap
            title.Margin = Thickness(0, 4, 0, 0)
            body.Children.Add(title)
            if item['content']:
                text = self._text(item['content'][:900], 11, bold=False, color=self._c('text2'))
                text.TextWrapping = TextWrapping.Wrap
                text.Margin = Thickness(0, 4, 0, 0)
                body.Children.Add(text)
            if item['expanded']:
                code = self._text(item['expanded'][:6000], 10.5, bold=False, color=self._c('text2'))
                code.FontFamily = FontFamily('Consolas')
                scroll = ScrollViewer()
                scroll.MaxHeight = 180
                scroll.Content = code
                scroll.Visibility = Visibility.Collapsed
                scroll.Margin = Thickness(0, 4, 0, 0)
                toggle = self._pill(self._glyph_label('\uE70D', 'Details', self._c('text2')), self._c('chip'))
                toggle.Cursor = Cursors.Hand
                toggle.ToolTip = 'Show what will run'

                def flip(s, e, scroll=scroll):
                    scroll.Visibility = (Visibility.Collapsed if scroll.Visibility ==
                                         Visibility.Visible else Visibility.Visible)
                self._hook(toggle, 'MouseLeftButtonUp', flip, transient=True)
                toggle.Margin = Thickness(0, 6, 0, 0)
                toggle.HorizontalAlignment = HorizontalAlignment.Left
                body.Children.Add(toggle)
                body.Children.Add(scroll)
            always = None
            if item['always_text']:
                always = CheckBox()
                always.Content = self._text(item['always_text'], 11, bold=False, color=self._c('text2'))
                always.Margin = Thickness(0, 6, 0, 0)
                body.Children.Add(always)
            allow = self._pill(self._glyph_label('\uE73E', 'Allow', self._c('on_accent'), 11), self._c('accent'))
            deny = self._pill(self._glyph_label('\uE711', 'Deny', self._c('text'), 11), self._c('chip'))
            for chip, ok in ((allow, True), (deny, False)):
                chip.Cursor = Cursors.Hand
                chip.Padding = Thickness(10, 3, 10, 3)
                chip.Margin = Thickness(0, 0, 8, 0)

                def click(s, e, rid=item['rid'], ok=ok, always=always):
                    self._answer(rid, ok, bool(always is not None and always.IsChecked))
                self._hook(chip, 'MouseLeftButtonUp', click, transient=True)
            buttons = self._row(allow, deny)
            buttons.Margin = Thickness(0, 8, 0, 0)
            body.Children.Add(buttons)
            self._ask_panel.Children.Add(card)
        if not self._asks:
            try:
                self.hud.Topmost = False
            except Exception:
                pass

    def _effects(self):
        """The Effects switches for the call being drawn, read once per call:
        {'trail', 'ripple', 'labels'} (agent_cursors.OPTIONS; on when unreadable).
        Glow is the main thread's (agent_cursors_glow)."""
        try:
            opts = options()
        except Exception:
            opts = {}
        self._fx = dict((name, bool(opts.get(name, True))) for name in ('trail', 'ripple', 'labels'))
        return self._fx

    def _reduce(self):
        """Animation is Off (agent_cursors.motion_off)."""
        try:
            return motion_off()
        except Exception:
            return False

    # ── a finished call ──────────────────────────────────────────────────────

    def show(self, agent, action, tool_name, failed, snap, followed=(),
             args=None, n_targets=0, changes=(), clashes=(), touched=0, goal='',
             record=True, task_text='', created='', fresh=()):
        """created = 'Walls ×3 created' (agent_cursors.work_label), fresh =
        the ids it is about: those boxes stay highlighted longer. record =
        False is a replayed step: it is drawn, not counted or listed again."""
        self._end_summary()
        if record:
            self._stop_replay()         # a live call wins over a replay
            self._history.append((agent, action, tool_name, failed, _light(snap), list(followed),
                                  args, n_targets, list(changes), list(clashes), touched, goal,
                                  task_text, created, list(fresh)))
            del self._history[:-REPLAY_SIZE]
        self._touched = touched or self._touched
        if goal and agent not in self._goals:
            self._goals[agent] = goal
        label, _glyph, acolor = action_style(action)
        caption, detail = self._step_text(tool_name, args, n_targets, failed, task_text)
        self._working.pop(agent, None)
        self._waiting.pop(agent, None)
        self._acting.pop(agent, None)
        for vid, name, is_new in followed:
            if vid not in [v[0] for v in self._views]:
                self._views.append((vid, name))
        if self._model():
            try:
                self._draw_step(agent, action, failed, snap, followed, n_targets, changes,
                                clashes, created, fresh, caption, label, acolor)
            except Exception:
                self._clear_focus(agent)    # decoration: the bar still gets the step
        else:
            self._bar_only()            # the default: the bar, nothing on the model
        self._log_step(agent, action, tool_name, failed, caption, detail, changes, clashes,
                       record)
        self._show_hud(snap['active']['rect'])

    @staticmethod
    def _step_text(tool_name, args, n_targets, failed, task_text='', document=''):
        """(caption, detail) of a finished call: with a plan the task leads
        and what the call did goes underneath."""
        caption = describe_call(tool_name, args, n_targets, failed)
        detail = tool_name
        if task_text:
            detail = u'{} · {}'.format(caption, tool_name)
            caption = (u'Failed: ' if failed else u'') + task_text
        if document:            # a file that is open but not on screen
            caption = u'{} · in {}'.format(caption, document)
        return caption, detail

    def done(self, agent, action, tool_name, failed, args=None, n_targets=0, goal='',
             task_text=''):
        """A call ended with nothing to draw (no document or view left, or
        notify() failed half way): the bar gets the step and stops saying
        Working. Before 2026-11-19 such a call never ended on the bar."""
        self._end_summary()
        self._stop_replay()
        if goal and agent not in self._goals:
            self._goals[agent] = goal
        caption, detail = self._step_text(tool_name, args, n_targets, failed, task_text)
        self._working.pop(agent, None)
        self._waiting.pop(agent, None)
        self._acting.pop(agent, None)
        self._clear_focus(agent)
        cur = self._cursors.get(agent)
        if cur is not None:
            self._cursor_state(cur, 'failed' if failed else 'idle', caption)
        self._log_step(agent, action, tool_name, failed, caption, detail, (), (), True)
        self._show_hud(None)

    def _bar_only(self):
        """Bar only (On model off): the bar may show, the canvas never does;
        a cursor left from before the switch goes."""
        self._wanted = True
        if self._cursors or self.canvas.Children.Count:
            for cur in self._cursors.values():
                cur['group'].BeginAnimation(_opacity_property(), None)
            self.canvas.Children.Clear()
            self._cursors = {}
            self._focus = {}
        if self.win.IsVisible:
            self.win.Hide()

    def _model(self):
        try:
            return on_model()
        except Exception:
            return False

    def _draw_step(self, agent, action, failed, snap, followed, n_targets, changes,
                   clashes, created, fresh, caption, label, acolor):
        """The step on the model (On model): cursor tour, boxes, effects."""
        self._place(snap['rect'])
        f = speed_factor()
        reduce = self._reduce()
        fx = self._effects()
        cur = self._cursor(agent, action)
        self._clear_focus(agent)
        self._cursor_state(cur, 'failed' if failed else 'idle', caption)
        if followed:
            first = followed[0]
            self._flash_view(snap['active'], '{} · {}'.format(
                'Opened' if first[2] else 'Switched to', first[1]), action_style('view')[2])

        # Stops: every box / deleted box, visited in order (move, click, next).
        stops, ghosts = [], []
        for geo in snap['views']:
            for item in geo['boxes']:
                b = self._local(item['box'])     # the tip lands on the element
                stops.append((((b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0), 'box', b, item))
            for item in geo['gone']:
                b = self._local(item['box'])
                stops.append((((b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0), 'gone', b, item))
            ghosts.extend(geo['ghosts'])
        boxes = [st for st in stops if st[1] == 'box']
        fresh_ids = set(fresh)
        fresh_boxes = [st[2] for st in boxes if st[3].get('id') in fresh_ids]
        if cluster_stops(len(boxes)):
            # Many targets: one frame round them all and one stop, not a box each.
            u = union_box([st[2] for st in boxes])
            text = u'\u00d7{} · {}'.format(max(n_targets, len(boxes)), label)
            stops = [(((u[0] + u[2]) / 2.0, (u[1] + u[3]) / 2.0), 'box', u,
                      {'label': text, 'size': text, 'fresh': bool(fresh_boxes)})] + [
                st for st in stops if st[1] != 'box']
        active_local = self._local(snap['active']['rect'])
        if stops:
            points = [st[0] for st in stops[:MAX_STOPS]]
        else:
            # Nothing to point at (a read, a list, a review): a small scan
            # next to where this cursor last worked, not across the view.
            slot = sorted(self._cursors).index(agent)
            n = 1 if failed or reduce else (3 if action in ('read', 'view', 'export', 'code') else 2)
            points = scan_points((cur['x'], cur['y']), active_local,
                                 self._actions * 31 + slot * 7 + 1, n)
        hop = 1 if reduce else int(hop_ms(len(points)) * f)
        sched = tour_schedule(points, hop, 0 if reduce else int(DWELL_MS * f))
        change_by_id = dict(changes)
        clash_ids = dict(clashes)
        for k, (_pt, kind, b, item) in enumerate(stops):
            arrive = sched[min(k, len(sched) - 1)][2]
            if kind == 'box':
                self._draw_box(b[0], b[1], b[2], b[3], item.get('label') or item['size'],
                               action, arrive,
                               fresh=item.get('fresh') or item.get('id') in fresh_ids)
                if fx['labels'] and item.get('id') in change_by_id:
                    self._draw_chip(b[0], b[3] + 24, change_by_id[item['id']], acolor, arrive)
                if item.get('id') in clash_ids:
                    self._flash_clash(b, arrive)
            else:
                self._draw_gone(b[0], b[1], b[2], b[3], arrive)
        for old, new in ghosts:
            self._draw_ghost(self._local(old['box']), self._local(new['box']),
                             acolor, sched[0][2])
        self._draw_offscreen(active_local, snap.get('offscreen') or (), acolor)
        if created and fresh_boxes and not failed:
            # What the call made or changed, named once above it; it stays.
            hold = int(FRESH_HOLD_S * 1000 * f)
            if fx['labels']:
                self._draw_label(union_box(fresh_boxes), created, action, sched[-1][2], hold)
            if not reduce and fx['trail'] and len(fresh_boxes) > 1 and len(stops) == len(boxes) > 1:
                self._build_trail(sched, acolor, hold)
        if not reduce and fx['trail']:
            self._trail(cur['x'], cur['y'], points[0][0], points[0][1], acolor)
        self._tour(cur, sched, shake=failed and not reduce)
        self._layout_tags(agent)
        if not reduce and fx['ripple']:
            ring = '#DC2626' if failed else acolor
            for x, y, arrive, _leave in sched:
                self._ripple(x, y, ring, arrive)
            self._press(cur, [st[2] for st in sched])
        self._wake(cur)

    def _log_step(self, agent, action, tool_name, failed, caption, detail, changes, clashes,
                  record):
        """The step on the bar: feed row, counts, changed values."""
        step = {'agent': agent, 'action': action, 'caption': caption, 'tool': detail,
                'failed': failed, 't': time.time()}
        self._last_changes = [text for _i, text in changes]
        if record:                  # a replay neither counts nor lists a step twice
            self._feed.insert(0, step)
            for elem_id, others in clashes:
                self._feed.insert(0, {
                    'agent': agent, 'action': action, 'tool': tool_name, 'failed': True,
                    'kind': 'warn', 't': time.time(),
                    'caption': u'{} and {} both edited element {}'.format(
                        agent_style(agent)[0], ', '.join(agent_style(o)[0] for o in others),
                        elem_id)})
            del self._feed[FEED_SIZE:]
            counts = self._counts.setdefault(agent, {})
            counts[action] = counts.get(action, 0) + 1
            self._actions += 1
            self._failed += 1 if failed else 0
            self._note = ''
        else:
            self._replay_step = step    # the band fills to k / n with this step on it
        self._last = time.time()

    # ── chips, conflicts, off-screen arrows ──────────────────────────────────

    def _draw_chip(self, x, y, text, color, delay_ms):
        """'Comments: "" → Checked' under a box."""
        from System.Windows import Thickness
        from System.Windows.Controls import Canvas
        chip = self._pill(self._glyph_label('\uE70F', text, '#111827', 10), '#FEF3C7', radius=2)
        chip.BorderBrush = self._brush(color)
        chip.BorderThickness = Thickness(1)
        chip.Opacity = 0
        Canvas.SetLeft(chip, x)
        Canvas.SetTop(chip, y)
        self.canvas.Children.Add(chip)
        self._pulse(chip, delay_ms, int(BOX_HOLD_S * 1000 * speed_factor()) + 1500,
                    self._remover(chip))

    def _flash_clash(self, b, delay_ms):
        """Two agents edited this element: a thick red frame flashes."""
        from System.Windows.Controls import Canvas
        group = Canvas()
        group.Opacity = 0
        group.Children.Add(self._frame(b[0] - 4, b[1] - 4, b[2] + 4, b[3] + 4, '#DC2626',
                                       thickness=3, fill_alpha=30))
        self.canvas.Children.Insert(0, group)
        frames = [(0.0, delay_ms)]
        for k in range(3):
            frames += [(1.0, delay_ms + 150 + 400 * k), (0.2, delay_ms + 350 + 400 * k)]
        frames += [(1.0, delay_ms + 1400), (1.0, delay_ms + 3000), (0.0, delay_ms + 3600)]
        group.BeginAnimation(_opacity_property(), self._anim(frames, self._remover(group)))

    def _draw_offscreen(self, rect, points, color):
        """Edge arrows (with a count) toward targets outside the view."""
        from System.Windows import Thickness
        from System.Windows.Controls import Canvas
        from System.Windows.Media import Geometry, RotateTransform
        from System.Windows.Shapes import Path
        local = [((px - self._origin[0]) / self._dpi(), (py - self._origin[1]) / self._dpi())
                 for px, py in points]
        for x, y, angle, count in offscreen_markers(rect, local):
            group = Canvas()
            group.Opacity = 0
            arrow = Path()
            arrow.Data = Geometry.Parse('M0,-9 L16,0 L0,9 L4,0 Z')
            arrow.Fill = self._brush(color)
            arrow.RenderTransform = RotateTransform(angle)
            Canvas.SetLeft(arrow, x)
            Canvas.SetTop(arrow, y)
            group.Children.Add(arrow)
            badge = self._pill(self._text('{}'.format(count), 10), color, radius=8)
            badge.Padding = Thickness(5, 0, 5, 0)
            Canvas.SetLeft(badge, x - 22)
            Canvas.SetTop(badge, y - 22)
            group.Children.Add(badge)
            group.ToolTip = '{} target{} off screen'.format(count, '' if count == 1 else 's')
            self.canvas.Children.Add(group)
            self._pulse(group, 200, int(BOX_HOLD_S * 1000), self._remover(group))

    # ── replay and the idle summary ─────────────────────────────────────────

    def replay(self):
        """Play the last steps again (screen positions as they were then).
        Not while a call runs: it would clear that call's cursor."""
        if not self._history or self._replay is not None:
            return
        if self._working:
            self._note = u'Replay waits until the running call ends'
            return
        from System import TimeSpan
        self._end_summary()
        self._note = ''
        for cur in self._cursors.values():
            cur['group'].BeginAnimation(_opacity_property(), None)
        self.canvas.Children.Clear()
        self._cursors = {}
        self._focus = {}
        self._replay = [list(self._history), 0]
        self._replay_timer.Interval = TimeSpan.FromMilliseconds(int(1500 * speed_factor()))
        self._replay_timer.Start()
        self._on_replay_tick(None, None)

    def _on_replay_tick(self, sender, e):
        try:
            if self._replay is None:
                self._replay_timer.Stop()
                return
            records, k = self._replay
            if k >= len(records):
                self._stop_replay(u'Replay done · {} steps'.format(len(records)))
                self._refresh_hud()
                return
            rec = records[k]
            self._replay[1] = k + 1
            # show() redraws the bar: its band carries the step and fills to k / n.
            self.show(rec[0], rec[1], rec[2], rec[3], rec[4], rec[5], args=rec[6],
                      n_targets=rec[7], changes=rec[8], clashes=rec[9], touched=rec[10],
                      goal=rec[11], record=False, task_text=rec[12], created=rec[13],
                      fresh=rec[14])
        except Exception:
            self._stop_replay()

    def _stop_replay(self, note=''):
        """End a replay (done, failed, or a live call arrived); `note` goes
        to the footer."""
        if self._replay is None:
            return
        self._replay = None
        self._replay_step = None
        self._replay_timer.Stop()
        self._note = note

    def _keep(self):
        try:
            return bool(options().get('keep'))
        except Exception:
            return False

    def _enter_summary(self, keep=False):
        """Agents idle: cursors, boxes and Glow leave (owner, 2026-11-17: they
        stayed after the agent stopped); the bar lists each agent's counts
        and Keep bar keeps only the bar (folded). They fade out; a call that
        arrives meanwhile keeps them (_end_summary)."""
        self._summary = time.time()
        self._resting = keep
        fn = (self.host.callbacks if self.host is not None else {}).get('idle')
        if fn is not None:
            fn()                        # Glow off the model (main thread)
        if self._reduce() or not self._cursors:
            self._leave()
            return
        token = self._leaving = object()

        def gone(s, e):
            if self._leaving is token:      # nothing came back meanwhile
                self._leave()
        try:
            self.canvas.BeginAnimation(_opacity_property(),
                                       self._anim(((1.0, 0), (0.0, 420)), gone))
        except Exception:
            self._leave()

    def _leave(self):
        """Take cursors and boxes off the views and hide the canvas window."""
        self._leaving = None
        for cur in self._cursors.values():
            cur['group'].BeginAnimation(_opacity_property(), None)
        self.canvas.Children.Clear()
        self._cursors = {}
        self._focus = {}
        self.win.Hide()
        self.canvas.BeginAnimation(_opacity_property(), None)
        self.canvas.Opacity = 1.0

    def _end_summary(self):
        if self._leaving is not None:       # back before the cursors had gone
            self._leaving = None
            self.canvas.BeginAnimation(_opacity_property(), None)
            self.canvas.Opacity = 1.0
        if self._summary is not None:
            self._summary = None
            self._resting = False

    def _select_touched(self):
        fn = (self.host.callbacks if self.host is not None else {}).get('select_touched')
        if fn is not None:
            fn()

    def hide(self):
        for item in list(self._asks):          # closing the bar answers No
            self._answer(item['rid'], False)
        self._timer.Stop()
        for cur in self._cursors.values():
            cur['group'].BeginAnimation(_opacity_property(), None)
        self.canvas.Children.Clear()
        self._cursors = {}
        self._focus = {}
        self._counts = {}
        self._views = []
        self._feed = []
        self._actions = 0
        self._transient = []
        self._working = {}
        self._waiting = {}
        self._acting = {}
        self._goals = {}
        self._plans = {}
        self._failed = 0
        self._last_changes = []
        self._note = ''
        self._drawer = False
        self._folded = False
        self._replay = None
        self._replay_step = None
        self._replay_timer.Stop()
        self._stop_loops()
        self._drawn = {}
        self._end_summary()
        self._wanted = False
        self.win.Hide()
        self.hud.Hide()

    def _on_tick(self, sender, e):
        try:
            now = time.time()
            if self._working and self._status_label is not None:
                self._status_label.Text = status_text('working', self._seconds())
            if self._working or self._waiting or self._asks or self._replay is not None:
                self._last = now
                return
            if self.hud.IsVisible and self._fold_due():
                self._refresh_hud()             # idle long enough: fold to one line
            if self.hud.IsMouseOver:
                self._last = now            # reading the HUD / toggling
                if self._summary is not None:
                    self._summary = now
            elif self._summary is not None:
                if now - self._summary >= SUMMARY_S and not self._keep():
                    self.hide()
            elif now - self._last >= IDLE_HIDE_S:
                keep = self._keep()
                if self._actions or keep:
                    self._enter_summary(keep)
                    self._refresh_hud()
                else:
                    self.hide()
        except Exception:
            pass

    def _fold_due(self):
        from System.Windows import Visibility
        folded = self._compact.Visibility == Visibility.Visible
        want = should_collapse(self._folded, time.time() - self._last, self._hover, self._asks)
        return want != folded

    def _on_front_tick(self, sender, e):
        """Own thread only (the windows are Topmost): hide while Revit is
        minimized or another app's window covers its views; stay visible
        while the AI app sits beside Revit or on another screen."""
        try:
            front, minimized, rect = _front_info(self.host.revit_hwnd if self.host else None)
            self._front_ok = ok = not front_hides(front, minimized, rect, self._canvas_px)
            model = self._model()
            for window in (self.win, self.hud):
                if window is self.win and not model:
                    continue                # bar only: the canvas stays hidden
                if window is self.hud and self._asks:
                    if not window.IsVisible:
                        window.Visibility = _visibility('Visible')
                    continue                # a question waits: stay in sight
                if not ok and window.IsVisible:
                    window.Visibility = _visibility('Hidden')
                elif ok and self._wanted and not window.IsVisible and (
                        window is self.hud or self._summary is None):
                    window.Visibility = _visibility('Visible')
        except Exception:
            pass

def _swallow(sender, e):
    e.Handled = True


def _translate_x():
    from System.Windows.Media import TranslateTransform
    return TranslateTransform.XProperty


def _translate_y():
    from System.Windows.Media import TranslateTransform
    return TranslateTransform.YProperty


def _thickness(n):
    from System.Windows import Thickness
    return Thickness(n)


def _visibility(name):
    from System.Windows import Visibility
    return getattr(Visibility, name)


_PID = [None]


def _hwnd_int(hwnd):
    try:
        return int(hwnd.ToInt64())      # System.IntPtr
    except AttributeError:
        return int(hwnd or 0)


def _front_info(revit_hwnd):
    """(foreground window is Revit, Revit minimized, foreground window rect in
    screen px or None). Unknown -> (True, False, None): keep showing."""
    try:
        import ctypes

        class _RECT(ctypes.Structure):
            _fields_ = [('left', ctypes.c_long), ('top', ctypes.c_long),
                        ('right', ctypes.c_long), ('bottom', ctypes.c_long)]

        user32 = ctypes.windll.user32
        if _PID[0] is None:
            from System.Diagnostics import Process
            _PID[0] = Process.GetCurrentProcess().Id
        minimized = False
        main = _hwnd_int(revit_hwnd)
        if main:
            minimized = bool(user32.IsIconic(ctypes.c_void_p(main)))
        hwnd = user32.GetForegroundWindow()
        pid = ctypes.c_ulong(0)
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value == _PID[0]:
            return True, minimized, None
        r = _RECT()
        if not hwnd or not user32.GetWindowRect(ctypes.c_void_p(hwnd), ctypes.byref(r)):
            return False, minimized, None
        return False, minimized, (r.left, r.top, r.right, r.bottom)
    except Exception:
        return True, False, None


def _light(snap):
    """A snapshot without Revit objects, kept for Replay."""
    views = []
    for geo in snap.get('views') or ():
        views.append(dict((k, v) for k, v in geo.items() if k != 'view'))
    active = views[0] if views else {}
    out = dict((k, v) for k, v in snap.items() if k not in ('views', 'active'))
    out['views'], out['active'] = views, active
    return out
