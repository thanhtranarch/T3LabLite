# -*- coding: utf-8 -*-
"""Settings > Agent Cursors window: the on/off switch, every option, and what
the feature gives and costs. Laid out like BatchOut's Format page (owner,
2026-11-19): cards with a sunken header over a two-column body.

No Revit API and no settings I/O here: the button script passes the current
values in and saves what ask() returns. Checkboxes are found by name,
`opt_<option>` in Tools/AgentCursorsSettings.xaml; an option without one is
kept as it was. Test: dev/test_agent_cursors_overlay.py (on a fake WPF).
"""

import os

try:
    from GUI.WPF_Base import T3WPFWindow as _Window
except Exception:           # no WPF (tests, a broken engine): ask() falls back
    _Window = object

_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'AgentCursorsSettings.xaml')
MASTER = 'model'            # T3LabLite: the option every drawing option hangs under
FREE = ('model', 'keep')    # options that act without MASTER (agent_cursors.effective)
HEIGHT, HEIGHT_OPEN = 480, 620      # as tall as the cards; taller while the details card is open
MODEL_ROW = 64                      # the "Show" group: without it the second card is cut off
_UP, _DOWN = u"", u""


class AgentCursorsSettingsWindow(_Window):
    """result: None on Cancel / X, else {'enabled': bool, option: bool, ...}."""

    def __init__(self, enabled, options, sections=()):
        _Window.__init__(self, _XAML)
        self.result = None
        self._setup(enabled, options, sections)

    def _setup(self, enabled, options, sections):
        """enabled = the feature is on now (the button then says Save);
        options = {name: bool}; sections = [(heading, [line], glyph)]."""
        self._options = dict(options)
        self._boxes = {}
        for name, value in self._options.items():
            box = getattr(self, 'opt_' + name, None)
            if box is not None:
                box.IsChecked = bool(value)
                self._boxes[name] = box
        self._height = HEIGHT
        if MASTER in self._boxes and getattr(self, 'row_model', None) is not None:
            self.row_model.Visibility = self._visibility(True)
            self._height = HEIGHT + MODEL_ROW
        self.Height = self._height
        self.chk_enabled.IsChecked = True       # opened to turn it on, or it is on
        self.btn_ok.Content = 'Save' if enabled else 'Turn On'
        self._fill(sections)
        self._sync()

    @staticmethod
    def _visibility(shown):
        from System.Windows import Visibility
        return Visibility.Visible if shown else Visibility.Collapsed

    def _style(self, key):
        try:
            return self.FindResource(key)
        except Exception:
            return None

    def _fill(self, sections):
        """Benefits left, drawbacks right: one bullet row per line."""
        from System.Windows import Thickness, TextWrapping, GridLength, GridUnitType, VerticalAlignment
        from System.Windows.Controls import Grid, ColumnDefinition, TextBlock
        panels = (getattr(self, 'pnl_benefits', None), getattr(self, 'pnl_drawbacks', None))
        for panel, section in zip(panels, sections):
            if panel is None:
                continue
            lines, glyph = section[1], section[2] if len(section) > 2 else u""
            for line in lines:
                row = Grid()
                for width in (GridLength.Auto, GridLength(1, GridUnitType.Star)):
                    column = ColumnDefinition()
                    column.Width = width
                    row.ColumnDefinitions.Add(column)
                icon = TextBlock()
                icon.Style = self._style('T3.Icon.Muted')
                icon.Text = glyph
                icon.Margin = Thickness(0, 4, 8, 0)
                icon.VerticalAlignment = VerticalAlignment.Top
                row.Children.Add(icon)
                text = TextBlock()
                text.Style = self._style('T3.Caption')
                text.Text = line
                text.TextWrapping = TextWrapping.Wrap
                Grid.SetColumn(text, 1)
                row.Children.Add(text)
                row.Margin = Thickness(0, 0, 0, 8)
                panel.Children.Add(row)

    def _sync(self):
        """Grey out what cannot act: everything while the feature is unticked,
        the drawing options while MASTER is. A greyed box keeps its tick."""
        on = bool(self.chk_enabled.IsChecked)
        master = self._boxes.get(MASTER)
        draws = on and (master is None or bool(master.IsChecked))
        for name, box in self._boxes.items():
            box.IsEnabled = on if name in FREE else draws
        try:                                    # BatchOut: the card of a ticked format is inked
            self.card_options.BorderBrush = self.FindResource('T3.Ink' if on else 'T3.Border')
        except Exception:
            pass

    def values(self):
        out = dict(self._options)
        out.update((name, bool(box.IsChecked)) for name, box in self._boxes.items())
        out['enabled'] = bool(self.chk_enabled.IsChecked)
        return out

    # ── handlers (AgentCursorsSettings.xaml) ────────────────────────────────

    def enabled_clicked(self, sender, e):
        self._sync()

    def model_clicked(self, sender, e):
        self._sync()

    def details_header_clicked(self, sender, e):
        """Fold / unfold the benefits card; the window grows with it (not
        while it is maximized: Height is then the size it restores to)."""
        from System.Windows import Visibility, WindowState
        opening = self.details_body.Visibility != Visibility.Visible
        self.details_body.Visibility = self._visibility(opening)
        self.details_arrow.Text = _UP if opening else _DOWN
        if self.WindowState == WindowState.Normal:
            self.Height = max(HEIGHT_OPEN, self._height) if opening else self._height

    def ok_clicked(self, sender, e):
        self.result = self.values()
        self.Close()

    def cancel_clicked(self, sender, e):
        self.result = None
        self.Close()


def ask(enabled, options, sections=()):
    """Show the window. {'enabled': bool, option: bool, ...} on Turn On /
    Save, None on Cancel. Raises when the window cannot be built: the caller
    falls back to the plain checkbox dialog."""
    window = AgentCursorsSettingsWindow(enabled, options, sections)
    window.ShowDialog()
    return window.result
