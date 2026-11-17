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
    ACTIONS, BOX_HOLD_S, DWELL_MS, FEED_SIZE, MAX_STOPS, MOVE_MS, IDLE_HIDE_S,
    OPTIONS, OPTION_LABELS, REPLAY_SIZE, SUMMARY_S, action_style,
    agent_style, classify_action, counts_text, describe_call, hop_ms, next_speed, plan_rows,
    offscreen_markers, options, review_points, set_enabled, set_option, set_speed,
    speed, speed_factor, summary_lines, tour_schedule)


class OverlayHost(object):
    """Owns the overlay. Under IronPython it lives on its own STA thread with
    its own Dispatcher, so cursors keep moving ("Working 8 s") while Revit's
    main thread runs a long tool call. Under CPython a thread of ours would
    freeze whenever Revit is idle (CLAUDE.md rule 8), so it stays on the
    calling thread. call() is safe from any thread. Env
    T3LAB_CURSORS_MAIN_THREAD=1 forces the calling thread."""

    def __init__(self, owner_hwnd, callbacks=None):
        self.callbacks = callbacks or {}
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
# Glyph font with fallback: Windows 11 ships Segoe Fluent Icons (a superset);
# a codepoint missing in the first font is taken from the next, never tofu
# (owner, 2026-11-17: a broken icon on the bar).
ICON_FONTS = 'Segoe Fluent Icons, Segoe MDL2 Assets'


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
        self._working = {}          # 'agent|action' -> (started, caption)
        self._waiting = {}          # 'agent|action' -> caption
        self._history = []          # replayable show() records, oldest first
        self._replay = None         # [records, index] while replaying
        self._summary = None        # time the idle summary card appeared
        self._touched = 0
        self._marks = []            # off-screen arrows of the last call
        self._cursors = {}          # 'agent|action' -> dict(group, tx, ...)
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

    def _anim(self, frames, done=None):
        """DoubleAnimationUsingKeyFrames from [(value, ms)]."""
        from System import TimeSpan
        from System.Windows.Media.Animation import (
            DoubleAnimationUsingKeyFrames, LinearDoubleKeyFrame, KeyTime)
        anim = DoubleAnimationUsingKeyFrames()
        for value, ms in frames:
            anim.KeyFrames.Add(LinearDoubleKeyFrame(
                value, KeyTime.FromTimeSpan(TimeSpan.FromMilliseconds(ms))))
        if done is not None:
            self._hook(anim, 'Completed', done, transient=True)
        return anim

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

    def _build_hud(self):
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border, StackPanel, Orientation
        from System.Windows.Input import Cursors

        bar = Border()
        bar.Background = self._brush('#1F2937', 235)
        bar.CornerRadius = CornerRadius(8)
        bar.Padding = Thickness(12, 8, 8, 8)
        col = StackPanel()
        bar.Child = col

        self._ask_panel = StackPanel()          # Allow / Deny cards (ask())
        col.Children.Add(self._ask_panel)

        top = StackPanel()
        top.Orientation = Orientation.Horizontal
        col.Children.Add(top)
        self._agents_panel = StackPanel()
        top.Children.Add(self._agents_panel)

        self._goal_panel = StackPanel()
        self._goal_panel.Margin = Thickness(16, 0, 0, 0)
        top.Children.Add(self._goal_panel)

        close = self._text('', 11, color='#CBD5E1', glyph=True)
        close.Margin = Thickness(8, 2, 4, 0)
        close.Cursor = Cursors.Hand
        close.ToolTip = ('Turn agent cursors off. Settings > Agent Cursors '
                         'turns them back on.')
        self._hook(close, 'MouseLeftButtonUp', lambda s, e: set_enabled(False))
        top.Children.Add(close)

        from System.Windows.Controls import WrapPanel
        self._toggles = WrapPanel()
        self._toggles.MaxWidth = 560
        self._toggles.Margin = Thickness(0, 6, 0, 0)
        col.Children.Add(self._toggles)

        self._views_text = self._text('', 11, bold=False, color='#CBD5E1')
        self._views_text.Margin = Thickness(0, 4, 0, 0)
        col.Children.Add(self._views_text)
        self._summary_panel = StackPanel()
        self._summary_panel.Margin = Thickness(0, 6, 0, 0)
        col.Children.Add(self._summary_panel)
        self._feed_panel = StackPanel()
        self._feed_panel.Margin = Thickness(0, 6, 0, 0)
        col.Children.Add(self._feed_panel)
        self.hud.Content = bar

    def _toggle(self, name, on):
        from System.Windows import Thickness
        from System.Windows.Input import Cursors
        label, tip = OPTION_LABELS[name]
        glyph = '' if on else ''        # CheckMark / Checkbox
        chip = self._pill(self._glyph_label(glyph, label, '#FFFFFF' if on else '#94A3B8'),
                          '#374151' if on else '#1F2937')
        chip.Margin = Thickness(4, 0, 0, 0)
        chip.Cursor = Cursors.Hand
        chip.ToolTip = tip

        def flip(s, e):
            try:
                set_option(name, not options()[name])
                self._refresh_hud()
            except Exception:
                pass
        self._hook(chip, 'MouseLeftButtonUp', flip, transient=True)
        return chip

    def _refresh_hud(self):
        from System.Windows import Thickness, CornerRadius, HorizontalAlignment
        from System.Windows.Controls import Border
        self._agents_panel.Children.Clear()
        for agent in sorted(self._counts):
            tag, color = agent_style(agent)
            dot = Border()
            dot.Width = dot.Height = 18
            dot.CornerRadius = CornerRadius(9)
            dot.Background = self._brush(color)
            dot.Margin = Thickness(0, 0, 8, 0)
            letter = self._text(tag[:1].upper(), 9)
            letter.HorizontalAlignment = HorizontalAlignment.Center
            dot.Child = letter
            parts = [dot, self._text(tag, 11.5)]
            counts = self._counts[agent]
            for key, label, glyph, acolor in ACTIONS:
                if not counts.get(key):
                    continue
                chip = self._pill(self._glyph_label(glyph, '{}'.format(counts[key]), '#FFFFFF'),
                                  acolor, radius=2)
                chip.Margin = Thickness(6, 0, 0, 0)
                chip.ToolTip = '{} · {} {}'.format(tag, label, counts[key])
                parts.append(chip)
            row = self._row(*parts)
            row.Margin = Thickness(0, 1, 0, 1)
            row.ToolTip = '{}: {}'.format(tag, counts_text(counts))
            self._agents_panel.Children.Add(row)
        self._goal_panel.Children.Clear()
        for agent in sorted(set(self._goals) | set(self._plans)):
            tag, color = agent_style(agent)
            if self._goals.get(agent):
                line = self._glyph_label('\uE7C1', u'{} · {}'.format(tag, self._goals[agent]),
                                         '#FDE68A', 11)
                self._goal_panel.Children.Add(line)
            for state, text in plan_rows(self._plans.get(agent) or {}):
                glyph, tint = {'done': ('\uE73E', '#4ADE80'), 'current': ('\uE768', '#FDE68A'),
                               'todo': ('\uE739', '#94A3B8')}[state]
                row = self._glyph_label(glyph, text, tint if state != 'done' else '#CBD5E1', 11)
                row.Margin = Thickness(16, 1, 0, 0)
                row.ToolTip = u'{} · {}'.format(tag, {'done': 'done', 'current': 'now',
                                                      'todo': 'next'}[state])
                self._goal_panel.Children.Add(row)
        self._toggles.Children.Clear()
        opts = options()
        for name, _k, _d in OPTIONS:
            self._toggles.Children.Add(self._toggle(name, opts[name]))
        self._toggles.Children.Add(self._button(
            '\uE916', 'Speed: {}'.format(speed().capitalize()),
            'Calm, Normal or Fast motion', lambda: set_speed(next_speed(speed()))))
        self._toggles.Children.Add(self._button(
            '\uE768', 'Replay', 'Replay the last {} steps'.format(REPLAY_SIZE), self.replay))
        self._feed_panel.Children.Clear()
        for step in self._feed:
            self._feed_panel.Children.Add(self._feed_row(step))
        n = self._actions
        names = [v[1] for v in self._views[-3:]]
        more = len(self._views) - len(names)
        views = ' · '.join(names) + (' · +{}'.format(more) if more > 0 else '')
        self._views_text.Text = '{} action{}{}'.format(
            n, '' if n == 1 else 's', ('  ·  Views: ' + views) if views else '')

    def _button(self, glyph, label, tip, action, color='#E5E7EB'):
        """A clickable chip on the HUD; refreshes the HUD after `action`."""
        from System.Windows import Thickness
        from System.Windows.Input import Cursors
        chip = self._pill(self._glyph_label(glyph, label, color), '#374151')
        chip.Margin = Thickness(4, 0, 0, 0)
        chip.Cursor = Cursors.Hand
        chip.ToolTip = tip

        def click(s, e):
            try:
                action()
                self._refresh_hud()
            except Exception:
                pass
        self._hook(chip, 'MouseLeftButtonUp', click, transient=True)
        return chip

    def _feed_row(self, step):
        """[agent dot] [action glyph] caption · agent   ok / failed."""
        from System.Windows import Thickness, CornerRadius
        from System.Windows.Controls import Border
        tag, color = agent_style(step['agent'])
        label, glyph, acolor = action_style(step['action'])
        if step.get('kind') == 'warn':
            glyph, acolor = '\uE7BA', '#F87171'
        elif step.get('kind') == 'goal':
            glyph, acolor = '\uE7C1', '#FDE68A'
        dot = Border()
        dot.Width = dot.Height = 8
        dot.CornerRadius = CornerRadius(4)
        dot.Background = self._brush(color)
        dot.Margin = Thickness(0, 0, 6, 0)
        icon = self._text(glyph, 10, color=acolor, glyph=True)
        icon.Margin = Thickness(0, 0, 6, 0)
        text = self._text(step['caption'], 11, bold=False, color='#F8FAFC')
        who = self._text(u'  · ' + tag, 10.5, bold=False, color='#94A3B8')
        state = self._text('' if step.get('kind') else ('\uE711' if step['failed'] else '\uE73E'), 10,
                           color='#F87171' if step['failed'] else '#4ADE80', glyph=True)
        state.Margin = Thickness(8, 0, 0, 0)
        row = self._row(dot, icon, text, who, state)
        row.Margin = Thickness(0, 1, 0, 1)
        row.ToolTip = step['tool']
        return row

    # ── cursors ──────────────────────────────────────────────────────────────

    def _cursor(self, agent, action):
        key = '{}|{}'.format(agent, action)
        cur = self._cursors.get(key)
        if cur is not None:
            return cur
        from System.Windows import Thickness
        from System.Windows.Controls import Canvas, StackPanel
        from System.Windows.Media import TranslateTransform, ScaleTransform, Geometry
        from System.Windows.Shapes import Path
        tag, color = agent_style(agent)
        label, glyph, acolor = action_style(action)

        group = Canvas()
        arrow = Path()
        arrow.Data = Geometry.Parse('M0,0 L0,17 L4.6,12.8 L7.6,19.6 L10.4,18.4 L7.4,11.6 L13.4,11.6 Z')
        arrow.Fill = self._brush(color)
        arrow.Stroke = self._brush('#FFFFFF')
        arrow.StrokeThickness = 1.2
        press = ScaleTransform()            # a "click" squeezes the arrow to its tip
        arrow.RenderTransform = press
        group.Children.Add(arrow)
        spin = self._spinner(color)
        group.Children.Add(spin)

        tags = StackPanel()
        Canvas.SetLeft(tags, 12)
        Canvas.SetTop(tags, 18)
        name = self._pill(self._row(self._text(tag), self._action_chip(glyph, label, acolor)),
                          color)
        tags.Children.Add(name)
        tool = self._text('', 11, color='#111827')          # what it does
        tool_name = self._text('', 9.5, bold=False, color='#6B7280')   # MCP tool
        lines = StackPanel()
        lines.Children.Add(tool)
        lines.Children.Add(tool_name)
        tool_pill = self._pill(lines, '#FFFFFF', alpha=238)
        tool_pill.Margin = Thickness(0, 2, 0, 0)
        tool_pill.BorderBrush = self._brush(acolor)
        tool_pill.BorderThickness = Thickness(1)
        tags.Children.Add(tool_pill)
        group.Children.Add(tags)

        tx = TranslateTransform()
        group.RenderTransform = tx
        self.canvas.Children.Add(group)
        # Cursors start at different spots so a swarm never stacks up.
        slot = len(self._cursors)
        cur = {'group': group, 'tx': tx, 'tool': tool, 'tool_name': tool_name,
               'tool_pill': tool_pill, 'press': press, 'color': acolor, 'spin': spin, 'x': 40.0 + 36 * slot, 'y': 40.0 + 24 * slot}
        tx.X, tx.Y = cur['x'], cur['y']
        self._cursors[key] = cur
        return cur

    def _action_chip(self, glyph, label, color):
        from System.Windows import Thickness
        chip = self._pill(self._glyph_label(glyph, label, '#FFFFFF', 9.5), color, radius=2)
        chip.Padding = Thickness(4, 0, 4, 0)
        chip.Margin = Thickness(6, 0, 0, 0)
        return chip

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

    def _draw_box(self, x0, y0, x1, y1, size, action, delay_ms):
        """One target box; `size` is the badge text (category · size)."""
        """One target box with the action's effect."""
        from System.Windows import Thickness
        from System.Windows.Controls import Canvas
        from System.Windows.Shapes import Rectangle, Shape
        from System.Windows.Media import TranslateTransform
        label, glyph, color = action_style(action)
        group = Canvas()
        group.Opacity = 0
        frame = self._frame(x0, y0, x1, y1, color, dashed=(action == 'code'))
        group.Children.Add(frame)
        for hx, hy in ((x0, y0), (x1, y0), (x0, y1), (x1, y1)):
            handle = Rectangle()
            handle.Width = handle.Height = 5
            handle.Fill = self._brush('#FFFFFF')
            handle.Stroke = self._brush(color)
            handle.StrokeThickness = 1
            Canvas.SetLeft(handle, hx - 2.5)
            Canvas.SetTop(handle, hy - 2.5)
            group.Children.Add(handle)
        badge = self._pill(self._glyph_label(glyph, size, '#FFFFFF'), color, radius=2)
        badge.Padding = Thickness(4, 1, 4, 1)
        Canvas.SetLeft(badge, max(x0, (x0 + x1) / 2.0 - 60))
        Canvas.SetTop(badge, y1 + 4)
        group.Children.Add(badge)
        self.canvas.Children.Insert(0, group)    # under the cursors
        hold = int(BOX_HOLD_S * 1000 * speed_factor())
        fx = None if self._reduce() else action
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

    def _spinner(self, color):
        """Ring that turns around the arrow while the call runs (hidden)."""
        from System.Windows import Visibility
        from System.Windows.Media import RotateTransform, DoubleCollection
        from System.Windows.Shapes import Ellipse
        from System.Windows.Controls import Canvas
        ring = Ellipse()
        ring.Width = ring.Height = 22
        ring.Stroke = self._brush(color)
        ring.StrokeThickness = 2
        dash = DoubleCollection()
        dash.Add(6)
        dash.Add(4)
        ring.StrokeDashArray = dash
        Canvas.SetLeft(ring, -11)
        Canvas.SetTop(ring, -11)
        rot = RotateTransform()
        rot.CenterX = rot.CenterY = 11
        ring.RenderTransform = rot
        ring.Visibility = Visibility.Collapsed
        return ring

    def _spin(self, cur, on):
        from System.Windows import Visibility
        from System.Windows.Media import RotateTransform
        from System.Windows.Media.Animation import DoubleAnimation, RepeatBehavior
        from System import TimeSpan
        from System.Windows import Duration
        ring = cur['spin']
        if on and not self._reduce():
            anim = DoubleAnimation()
            anim.From, anim.To = 0.0, 360.0
            anim.Duration = Duration(TimeSpan.FromMilliseconds(900))
            anim.RepeatBehavior = RepeatBehavior.Forever
            ring.Visibility = Visibility.Visible
            ring.RenderTransform.BeginAnimation(RotateTransform.AngleProperty, anim)
        else:
            ring.RenderTransform.BeginAnimation(RotateTransform.AngleProperty, None)
            ring.Visibility = Visibility.Collapsed

    def _place(self, rect):
        """Cover `rect` (screen px) with the canvas; shift cursors to match."""
        left, top, right, bottom = rect
        if not self.win.IsVisible:
            self.win.Width = self.win.Height = 1
            self._wanted = True
            self.win.Show()
        scale = self._dpi()
        if self._origin != (left, top):
            dx = (self._origin[0] - left) / scale
            dy = (self._origin[1] - top) / scale
            for cur in self._cursors.values():
                cur['x'] += dx
                cur['y'] += dy
                cur['tx'].BeginAnimation(_translate_x(), None)
                cur['tx'].BeginAnimation(_translate_y(), None)
                cur['tx'].X, cur['tx'].Y = cur['x'], cur['y']
        self._origin = (left, top)
        self.win.Left, self.win.Top = left / scale, top / scale
        self.win.Width = (right - left) / scale
        self.win.Height = (bottom - top) / scale

    def _show_hud(self, rect):
        scale = self._dpi()
        self._refresh_hud()
        if not self.hud.IsVisible:
            self.hud.Show()
        self.hud.UpdateLayout()
        aleft, _atop, aright, abottom = rect
        hud_w = self.hud.ActualWidth or 320
        self.hud.Left = aleft / scale + ((aright - aleft) / scale - hud_w) / 2.0
        self.hud.Top = abottom / scale - (self.hud.ActualHeight or 48) - 16
        self._timer.Start()

    def working(self, agent, action, caption, tool_name, rect, task_text=''):
        """A call has started: spin, say "Working", count the seconds. With a
        plan, the cursor names the current task and the call goes underneath."""
        self._end_summary()
        if not self.win.IsVisible or self._origin == (0.0, 0.0):
            self._place(rect)
        cur = self._cursor(agent, action)
        key = '{}|{}'.format(agent, action)
        if task_text:
            caption, tool_name = task_text, u'{} · {}'.format(caption, tool_name)
        self._working[key] = (time.time(), caption)
        cur['tool'].Text = u'Working · {}'.format(caption)
        cur['tool_name'].Text = tool_name
        self._spin(cur, True)
        cur['group'].BeginAnimation(_opacity_property(), None)
        cur['group'].Opacity = 1.0
        self._last = time.time()
        self._show_hud(rect)

    def waiting(self, agent, tool_name, caption, on):
        """The call waits for the user's OK: an amber, pulsing pill."""
        action = classify_action(tool_name)
        key = '{}|{}'.format(agent, action)
        cur = self._cursors.get(key)
        if on:
            self._waiting[key] = caption
            if cur is None:
                if not self.win.IsVisible:
                    return          # nothing on screen yet: no view to draw in
                cur = self._cursor(agent, action)
                cur['group'].Opacity = 1.0
            cur['tool'].Text = u'Waiting for your OK · {}'.format(caption)
            cur['tool_pill'].BorderBrush = self._brush('#F59E0B')
            cur['tool_pill'].BorderThickness = _thickness(2)
            if not self._reduce():
                from System.Windows.Media.Animation import RepeatBehavior
                anim = self._anim(((1.0, 0), (0.45, 500), (1.0, 1000)))
                anim.RepeatBehavior = RepeatBehavior.Forever
                cur['tool_pill'].BeginAnimation(_opacity_property(), anim)
        else:
            self._waiting.pop(key, None)
            if cur is not None:
                cur['tool_pill'].BeginAnimation(_opacity_property(), None)
                cur['tool_pill'].Opacity = 1.0
                cur['tool_pill'].BorderThickness = _thickness(1)
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
                                  'tool': 't3lab_narrate', 'failed': False, 'kind': 'goal'})
            del self._feed[FEED_SIZE:]
        self._last = time.time()
        if self.hud.IsVisible:
            self._refresh_hud()

    # ── consent on the bar (agent_cursors.ask) ───────────────────────────────

    def ask(self, rid, agent, tool_name, caption, main, content, expanded, always_text, rect):
        """A call waits for the user's OK: a card with Allow / Deny on the bar
        and the cursor pulsing, instead of a modal Revit dialog."""
        self._end_summary()
        self._place(rect)
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
            card.Background = self._brush('#111827')
            card.BorderBrush = self._brush('#F59E0B')
            card.BorderThickness = Thickness(1)
            card.CornerRadius = CornerRadius(6)
            card.Padding = Thickness(10, 8, 10, 8)
            card.Margin = Thickness(0, 0, 0, 8)
            card.MaxWidth = 560
            body = StackPanel()
            card.Child = body
            head = self._glyph_label('\uE7BA', u'{} asks for your OK'.format(tag), '#FDE68A', 11)
            body.Children.Add(head)
            title = self._text(item['main'] or item['caption'], 12, color='#FFFFFF')
            title.TextWrapping = TextWrapping.Wrap
            title.Margin = Thickness(0, 4, 0, 0)
            body.Children.Add(title)
            if item['content']:
                text = self._text(item['content'][:900], 11, bold=False, color='#E5E7EB')
                text.TextWrapping = TextWrapping.Wrap
                text.Margin = Thickness(0, 4, 0, 0)
                body.Children.Add(text)
            if item['expanded']:
                code = self._text(item['expanded'][:6000], 10.5, bold=False, color='#D1D5DB')
                code.FontFamily = FontFamily('Consolas')
                scroll = ScrollViewer()
                scroll.MaxHeight = 180
                scroll.Content = code
                scroll.Visibility = Visibility.Collapsed
                scroll.Margin = Thickness(0, 4, 0, 0)
                toggle = self._pill(self._glyph_label('\uE70D', 'Details', '#E5E7EB'), '#374151')
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
                always.Content = self._text(item['always_text'], 11, bold=False, color='#E5E7EB')
                always.Margin = Thickness(0, 6, 0, 0)
                body.Children.Add(always)
            allow = self._pill(self._glyph_label('\uE73E', 'Allow', '#FFFFFF', 11), '#16A34A')
            deny = self._pill(self._glyph_label('\uE711', 'Deny', '#FFFFFF', 11), '#4B5563')
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

    def _reduce(self):
        try:
            return bool(options().get('reduce'))
        except Exception:
            return False

    # ── a finished call ──────────────────────────────────────────────────────

    def show(self, agent, action, tool_name, failed, snap, followed=(),
             args=None, n_targets=0, changes=(), clashes=(), touched=0, goal='',
             record=True, task_text=''):
        self._end_summary()
        if record:
            self._history.append((agent, action, tool_name, failed, _light(snap), list(followed),
                                  args, n_targets, list(changes), list(clashes), touched, goal,
                                  task_text))
            del self._history[:-REPLAY_SIZE]
        self._touched = touched or self._touched
        if goal and agent not in self._goals:
            self._goals[agent] = goal
        self._place(snap['rect'])
        f = speed_factor()
        reduce = self._reduce()

        label, _glyph, acolor = action_style(action)
        caption = describe_call(tool_name, args, n_targets, failed)
        detail = tool_name
        if task_text:           # the plan's task leads; what the call did goes underneath
            detail = u'{} · {}'.format(caption, tool_name)
            caption = (u'Failed: ' if failed else u'') + task_text
        key = '{}|{}'.format(agent, action)
        cur = self._cursor(agent, action)
        self._working.pop(key, None)
        self._waiting.pop(key, None)
        self._spin(cur, False)
        cur['tool_pill'].BeginAnimation(_opacity_property(), None)
        cur['tool_pill'].Opacity = 1.0
        cur['tool_pill'].BorderThickness = _thickness(1)
        cur['tool'].Text = caption
        cur['tool_name'].Text = detail
        cur['tool_pill'].BorderBrush = self._brush('#DC2626' if failed else acolor)

        for vid, name, is_new in followed:
            if vid not in [v[0] for v in self._views]:
                self._views.append((vid, name))
        if followed:
            first = followed[0]
            self._flash_view(snap['active'], '{} · {}'.format(
                'Opened' if first[2] else 'Switched to', first[1]), action_style('view')[2])

        # Stops: every box / deleted box, visited in order (move, click, next).
        stops, ghosts = [], []
        for geo in snap['views']:
            for item in geo['boxes']:
                b = self._local(item['box'])
                stops.append(((b[2] - 4, b[1] + 4), 'box', b, item))
            for item in geo['gone']:
                b = self._local(item['box'])
                stops.append((((b[0] + b[2]) / 2.0, (b[1] + b[3]) / 2.0), 'gone', b, item))
            ghosts.extend(geo['ghosts'])
        active_local = self._local(snap['active']['rect'])
        if stops:
            points = [st[0] for st in stops[:MAX_STOPS]]
        else:
            # Nothing to point at (a read, a list, a review): scan the active
            # view, a few clicks along a path of its own per agent and action.
            slot = sorted(self._cursors).index(key)
            n = 1 if failed or reduce else (3 if action in ('read', 'view', 'export', 'code') else 2)
            points = review_points(active_local, self._actions * 31 + slot * 7 + 1, n)
        hop = 1 if reduce else int(hop_ms(len(points)) * f)
        sched = tour_schedule(points, hop, 0 if reduce else int(DWELL_MS * f))
        change_by_id = dict(changes)
        clash_ids = dict(clashes)
        for k, (_pt, kind, b, item) in enumerate(stops):
            arrive = sched[min(k, len(sched) - 1)][2]
            if kind == 'box':
                self._draw_box(b[0], b[1], b[2], b[3], item.get('label') or item['size'],
                               action, arrive)
                if item.get('id') in change_by_id:
                    self._draw_chip(b[0], b[3] + 24, change_by_id[item['id']], acolor, arrive)
                if item.get('id') in clash_ids:
                    self._flash_clash(b, arrive)
            else:
                self._draw_gone(b[0], b[1], b[2], b[3], arrive)
        for old, new in ghosts:
            self._draw_ghost(self._local(old['box']), self._local(new['box']),
                             acolor, sched[0][2])
        self._draw_offscreen(active_local, snap.get('offscreen') or (), acolor)
        if not reduce:
            self._trail(cur['x'], cur['y'], points[0][0], points[0][1], acolor)
        self._tour(cur, sched, shake=failed and not reduce)
        if not reduce:
            ring = '#DC2626' if failed else acolor
            for x, y, arrive, _leave in sched:
                self._ripple(x, y, ring, arrive)
            self._press(cur, [st[2] for st in sched])
        cur['group'].BeginAnimation(_opacity_property(), None)
        cur['group'].Opacity = 1.0

        self._feed.insert(0, {'agent': agent, 'action': action, 'caption': caption,
                              'tool': detail, 'failed': failed})
        for elem_id, others in clashes:
            self._feed.insert(0, {
                'agent': agent, 'action': action, 'tool': tool_name, 'failed': True,
                'kind': 'warn', 'caption': u'{} and {} both edited element {}'.format(
                    agent_style(agent)[0], ', '.join(agent_style(o)[0] for o in others), elem_id)})
        del self._feed[FEED_SIZE:]
        if record:                  # a replay does not count twice
            counts = self._counts.setdefault(agent, {})
            counts[action] = counts.get(action, 0) + 1
            self._actions += 1
        self._last = time.time()
        self._show_hud(snap['active']['rect'])

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
            arrow.Stroke = self._brush('#FFFFFF')
            arrow.StrokeThickness = 1
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
        """Play the last steps again (screen positions as they were then)."""
        if not self._history or self._replay is not None:
            return
        from System import TimeSpan
        self._end_summary()
        for cur in self._cursors.values():
            cur['group'].BeginAnimation(_opacity_property(), None)
        self.canvas.Children.Clear()
        self._cursors = {}
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
                self._replay = None
                self._replay_timer.Stop()
                self._views_text.Text = u'Replay done · {} steps'.format(len(records))
                return
            rec = records[k]
            self._replay[1] = k + 1
            self.show(rec[0], rec[1], rec[2], rec[3], rec[4], rec[5], args=rec[6],
                      n_targets=rec[7], changes=rec[8], clashes=rec[9], touched=rec[10],
                      goal=rec[11], record=False, task_text=rec[12])
            self._views_text.Text = u'Replay {} / {}'.format(k + 1, len(records))
        except Exception:
            self._replay = None
            self._replay_timer.Stop()

    def _keep(self):
        try:
            return bool(options().get('keep'))
        except Exception:
            return False

    def _enter_summary(self, keep=False):
        """Agents idle: the HUD adds a summary card. Keep bar on: cursors only
        dim and both windows stay, so the next call does not make them blink;
        off: the canvas is cleared and hidden."""
        from System.Windows import Thickness
        self._summary = time.time()
        self._resting = keep
        for cur in self._cursors.values():
            cur['group'].BeginAnimation(_opacity_property(), None)
            cur['group'].Opacity = 0.45
        if not keep:
            self.canvas.Children.Clear()
            self._cursors = {}
            self.win.Hide()
        self._summary_panel.Children.Clear()
        title = self._text(u'Agents are idle', 12)
        self._summary_panel.Children.Add(title)
        for line in summary_lines(self._counts, self._touched):
            row = self._text(line, 11, bold=False, color='#E5E7EB')
            row.Margin = Thickness(0, 2, 0, 0)
            self._summary_panel.Children.Add(row)
        buttons = self._row(
            self._button('\uE8B3', 'Select touched',
                         'Select every element the agents changed (no model change)',
                         self._select_touched),
            self._button('\uE768', 'Replay', 'Replay the last steps', self.replay))
        buttons.Margin = Thickness(-4, 6, 0, 0)
        self._summary_panel.Children.Add(buttons)

    def _end_summary(self):
        if self._summary is not None:
            self._summary = None
            self._resting = False
            self._summary_panel.Children.Clear()

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
        self._counts = {}
        self._views = []
        self._feed = []
        self._actions = 0
        self._transient = []
        self._working = {}
        self._waiting = {}
        self._goals = {}
        self._plans = {}
        self._replay = None
        self._replay_timer.Stop()
        self._end_summary()
        self._wanted = False
        self.win.Hide()
        self.hud.Hide()

    def _on_tick(self, sender, e):
        try:
            now = time.time()
            for key, (started, caption) in list(self._working.items()):
                cur = self._cursors.get(key)
                if cur is not None and key not in self._waiting:
                    cur['tool'].Text = u'Working {} s · {}'.format(int(now - started), caption)
            if self._working or self._waiting or self._asks or self._replay is not None:
                self._last = now
                return
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

    def _on_front_tick(self, sender, e):
        """Own thread only: show while Revit (this process) is in front."""
        try:
            front = _process_in_front()
            for window in (self.win, self.hud):
                if window is self.hud and self._asks:
                    if not window.IsVisible:
                        window.Visibility = _visibility('Visible')
                    continue                # a question waits: stay in sight
                if not front and window.IsVisible:
                    window.Visibility = _visibility('Hidden')
                elif front and self._wanted and not window.IsVisible and (
                        window is self.hud or self._summary is None or self._resting):
                    window.Visibility = _visibility('Visible')
        except Exception:
            pass

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


def _process_in_front():
    """True when the foreground window belongs to this process (Revit)."""
    try:
        import ctypes
        from System.Diagnostics import Process
        user32 = ctypes.windll.user32
        hwnd = user32.GetForegroundWindow()
        pid = ctypes.c_ulong(0)
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        return pid.value == Process.GetCurrentProcess().Id
    except Exception:
        return True


def _light(snap):
    """A snapshot without Revit objects, kept for Replay."""
    views = []
    for geo in snap.get('views') or ():
        views.append(dict((k, v) for k, v in geo.items() if k != 'view'))
    active = views[0] if views else {}
    out = dict((k, v) for k, v in snap.items() if k not in ('views', 'active'))
    out['views'], out['active'] = views, active
    return out
