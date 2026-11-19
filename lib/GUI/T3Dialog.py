# -*- coding: utf-8 -*-
"""
T3Dialog — Modern replacement for Revit TaskDialog and forms.alert.
Provides Info, Warning, Error (Danger), and Confirm modal dialogs conforming to T3 UI standard.
"""

import os

# Python.NET 3 refuses to convert a plain int into a .NET enum, so every
# Visibility assignment must use the real System.Windows.Visibility member --
# `x.Visibility = 2` raises "int can not be converted to Enum implicitly".
from System.Windows import Visibility as _Visibility

try:
    from GUI.WPF_Base import T3WPFWindow as _WPFWindow
except Exception:
    try:
        from WPF_Base import T3WPFWindow as _WPFWindow
    except Exception:
        # pyrevit.forms raises PyRevitCPythonNotSupported on *any* attribute
        # access under CPython, so this last resort must stay guarded.
        try:
            from pyrevit import forms as _forms
            _WPFWindow = _forms.WPFWindow
        except Exception:
            _WPFWindow = object


_XAML = os.path.join(os.path.dirname(__file__), 'Tools', 'T3Dialog.xaml')


class T3Dialog(_WPFWindow):
    """Universal T3 Modal Dialog for alerts, warnings, errors and confirmations."""

    MODE_INFO = "info"
    MODE_WARNING = "warning"
    MODE_ERROR = "error"
    MODE_CONFIRM = "confirm"
    MODE_PROMPT = "prompt"      # confirm + one-line text answer (ask_for_string)

    def __init__(self, message, title="Notification", details=None,
                 mode="info", ok_text="OK", cancel_text="Cancel", danger=False, owner=None,
                 default_text=u"", options=None, subtitle=None, sections=None,
                 sections_title="Details"):
        _WPFWindow.__init__(self, _XAML)
        self.result = False
        self.value = None
        self.option_values = {}
        self._option_boxes = []
        self._option_sync = None
        self._mode = mode
        self._danger = danger

        if owner:
            try:
                self.Owner = owner
            except Exception:
                pass


        # Wire Title and Texts
        if hasattr(self, 'dlg_title') and self.dlg_title:
            self.dlg_title.Text = title
        if subtitle and getattr(self, 'dlg_subtitle', None) is not None:
            self.dlg_subtitle.Text = subtitle
        if hasattr(self, 'txt_message') and self.txt_message:
            self.txt_message.Text = message
        if hasattr(self, 'txt_details') and self.txt_details:
            if details:
                self.txt_details.Text = str(details)
                self.txt_details.Visibility = _Visibility.Visible
            else:
                self.txt_details.Text = ""
                self.txt_details.Visibility = _Visibility.Collapsed

        # Configure Style & Icon based on mode
        self._configure_appearance(mode, ok_text, cancel_text, danger)
        if mode == self.MODE_PROMPT and getattr(self, 'txt_input', None) is not None:
            self.txt_input.Text = default_text or u""
            self.txt_input.Visibility = _Visibility.Visible
            try:
                self.txt_input.Focus()
                self.txt_input.SelectAll()
            except Exception:
                pass

        if options:
            self._add_options(options)
            if sections and self._add_sections(sections_title, sections):
                self.txt_details.Visibility = _Visibility.Collapsed
            self._fit_height()
            try:
                self.Loaded += self._on_layout_changed      # measured again once it is real
            except Exception:
                pass

        # Wire Events
        if hasattr(self, 'btn_ok') and self.btn_ok:
            self.btn_ok.Click += self._on_ok
        if hasattr(self, 'btn_cancel') and self.btn_cancel:
            self.btn_cancel.Click += self._on_cancel
        if hasattr(self, 'btn_close_chrome') and self.btn_close_chrome:
            self.btn_close_chrome.Click += self._on_cancel

    def _configure_appearance(self, mode, ok_text, cancel_text, danger):
        # Configure buttons
        if hasattr(self, 'btn_ok') and self.btn_ok:
            self.btn_ok.Content = ok_text
            if danger:
                danger_style = self.TryFindResource("T3.Button.Danger")
                if danger_style:
                    self.btn_ok.Style = danger_style

        if hasattr(self, 'btn_cancel') and self.btn_cancel:
            self.btn_cancel.Content = cancel_text
            if mode in (self.MODE_CONFIRM, self.MODE_PROMPT):
                self.btn_cancel.Visibility = _Visibility.Visible
            else:
                self.btn_cancel.Visibility = _Visibility.Collapsed

        # Configure Icon Glyph and Color
        # Segoe MDL2 glyphs from the T3 table (T3LAB_UI_STANDARD.md, section Icon):
        # E946 Info · E7BA Warning · E783 Error
        glyph = u"\uE946"
        fill_res = "T3.Success.Fill"
        fg_res = "T3.Success.Text"

        if mode == self.MODE_WARNING:
            glyph = u"\uE7BA"
            fill_res = "T3.Warning.Fill"
            fg_res = "T3.Warning.Text"
        elif mode == self.MODE_ERROR or danger:
            glyph = u"\uE783"
            fill_res = "T3.Danger.Fill"
            fg_res = "T3.Danger.Text"
        elif mode == self.MODE_CONFIRM:
            glyph = u"\uE7BA" if danger else u"\uE946"
            fill_res = "T3.Warning.Fill" if danger else "T3.SurfaceSunken"
            fg_res = "T3.Warning.Text" if danger else "T3.Text"

        if hasattr(self, 'icon_glyph') and self.icon_glyph:
            self.icon_glyph.Text = glyph
            fg_brush = self.TryFindResource(fg_res)
            if fg_brush:
                self.icon_glyph.Foreground = fg_brush

        if hasattr(self, 'icon_bg') and self.icon_bg:
            bg_brush = self.TryFindResource(fill_res)
            if bg_brush:
                self.icon_bg.Background = bg_brush

    def _style(self, key):
        try:
            return self.FindResource(key)
        except Exception:
            return None

    def _text(self, text, style_key, wrap=True):
        from System.Windows import TextWrapping
        from System.Windows.Controls import TextBlock
        block = TextBlock()
        style = self._style(style_key)
        if style is not None:
            block.Style = style
        block.Text = text or u""
        if wrap:
            block.TextWrapping = TextWrapping.Wrap
        return block

    def _add_options(self, options):
        """One row per (key, label, description, checked[, requires]) in
        pnl_options: a T3.CheckBox and, beside it, what the option does (it
        used to hide in a tooltip). `requires` = key of the option that must
        be ticked for this one to act (chains are followed): the row greys
        out with that box and keeps its own tick."""
        panel = getattr(self, 'pnl_options', None)
        if panel is None:
            return
        from System.Windows import Thickness, GridLength, VerticalAlignment
        from System.Windows.Controls import CheckBox, Grid, RowDefinition
        style = self._style('T3.CheckBox')
        rows = {}
        for index, option in enumerate(options):
            key, label, note, checked = option[:4]
            row = RowDefinition()
            row.Height = GridLength.Auto
            panel.RowDefinitions.Add(row)
            box = CheckBox()
            if style is not None:
                box.Style = style
            box.Content = label
            box.IsChecked = bool(checked)
            box.Margin = Thickness(0, 0, 16, 8)
            box.VerticalAlignment = VerticalAlignment.Center
            Grid.SetRow(box, index)
            panel.Children.Add(box)
            text = self._text(note, 'T3.Caption')
            text.Margin = Thickness(0, 0, 0, 8)
            Grid.SetRow(text, index)
            Grid.SetColumn(text, 1)
            panel.Children.Add(text)
            self._option_boxes.append((key, box))
            rows[key] = (box, text, option[4] if len(option) > 4 else None)
        panel.Visibility = _Visibility.Visible
        self.Width = max(self.Width, 560)

        def live(key, depth=0):
            # Ticked all the way up: an option under "On model" is off when
            # the on/off switch above that is unticked.
            requires = rows[key][2]
            if requires not in rows or depth > len(rows):
                return True
            return bool(rows[requires][0].IsChecked) and live(requires, depth + 1)

        def sync(sender=None, e=None):
            for key, (box, text, requires) in rows.items():
                if requires in rows:
                    on = live(key)
                    box.IsEnabled = on
                    text.Opacity = 1.0 if on else 0.5
        masters = set(requires for _b, _t, requires in rows.values() if requires in rows)
        if masters:
            self._option_sync = sync        # pythonnet: the delegate must outlive this call
            for key in masters:
                rows[key][0].Click += sync
            sync()

    def _add_sections(self, title, sections):
        """Long reading under the options, folded: [(heading, [line][, glyph])].
        Each line is a bullet with a hanging indent (glyph = a T3 icon,
        default a check). False when the XAML has no place for it."""
        expander = getattr(self, 'exp_details', None)
        panel = getattr(self, 'pnl_sections', None)
        if expander is None or panel is None:
            return False
        from System.Windows import Thickness, GridLength, GridUnitType, VerticalAlignment
        from System.Windows.Controls import Grid, ColumnDefinition
        for section in sections:
            heading, lines = section[0], section[1]
            glyph = section[2] if len(section) > 2 else u"\uE73E"
            head = self._text(heading.upper(), 'T3.Label', wrap=False)
            head.Margin = Thickness(0, 8, 0, 4)
            panel.Children.Add(head)
            for line in lines:
                row = Grid()
                for width in (GridLength.Auto, GridLength(1, GridUnitType.Star)):
                    column = ColumnDefinition()
                    column.Width = width
                    row.ColumnDefinitions.Add(column)
                icon = self._text(glyph, 'T3.Icon.Muted', wrap=False)
                icon.Margin = Thickness(0, 4, 8, 0)
                icon.VerticalAlignment = VerticalAlignment.Top
                row.Children.Add(icon)
                text = self._text(line, 'T3.Body.Secondary')
                Grid.SetColumn(text, 1)
                row.Children.Add(text)
                row.Margin = Thickness(0, 0, 0, 4)
                panel.Children.Add(row)
        expander.Header = title
        expander.Visibility = _Visibility.Visible
        try:
            expander.Expanded += self._on_layout_changed
            expander.Collapsed += self._on_layout_changed
        except Exception:
            pass
        return True

    def _on_layout_changed(self, sender, e):
        self._fit_height()

    def _fit_height(self):
        """As tall as the content needs, title bar and footer (48 + 48)
        included. The options dialog had a fixed 440: a scroll box when the
        text was long, a gap when it was short."""
        body = getattr(self, 'pnl_body', None)
        try:
            from System.Windows import Size
            body.Measure(Size(max(float(self.Width) - 74.0, 100.0), float('inf')))
            need = float(body.DesiredSize.Height) + 98.0
            self.Height = min(max(need, float(self.MinHeight)), 680.0)
        except Exception:
            self.Height = max(self.Height, 440)

    def _on_ok(self, sender, e):
        self.result = True
        if getattr(self, 'txt_input', None) is not None:
            self.value = self.txt_input.Text
        self.option_values = dict((k, bool(b.IsChecked)) for k, b in self._option_boxes)
        self.Close()

    def _on_cancel(self, sender, e):
        self.result = False
        self.Close()


# ── LAST-RESORT FALLBACK ─────────────────────────────────────────────────────
# A dialog whose whole job is to carry a message must never be the thing that
# swallows it. If the WPF window cannot be built for any reason -- XamlReader
# unavailable in this engine, a stale module in the persistent CPython engine,
# a broken resource -- fall back to Revit's own TaskDialog so the user still
# reads the message instead of a traceback.

def _task_dialog(message, title, details, buttons=1):
    """Revit's native TaskDialog. Returns True on OK/Yes, False otherwise."""
    from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult
    dialog = TaskDialog(title or "T3Lab")
    dialog.MainInstruction = message or ""
    if details:
        dialog.MainContent = details
    if buttons == 2:
        dialog.CommonButtons = (TaskDialogCommonButtons.Yes |
                                TaskDialogCommonButtons.No)
        dialog.DefaultButton = TaskDialogResult.No
        return dialog.Show() == TaskDialogResult.Yes
    dialog.CommonButtons = TaskDialogCommonButtons.Close
    dialog.Show()
    return True


def _run(build, message, title, details, buttons=1):
    """Show the T3 dialog, or degrade to TaskDialog rather than raise."""
    try:
        dlg = build()
        dlg.ShowDialog()
        return dlg.result if buttons == 2 else True
    except Exception as exc:
        try:
            note = details or ""
            # Say why the styled dialog is missing: silently swapping chrome
            # looks like a different bug next time somebody reports it.
            if note:
                note += "\n\n"
            note += "(T3 dialog unavailable: %s)" % exc
            return _task_dialog(message, title, note, buttons)
        except Exception:
            return False if buttons == 2 else True


def show_info(message, title="Information", details=None, owner=None):
    """Show an info modal dialog."""
    return _run(lambda: T3Dialog(message, title=title, details=details,
                                 mode=T3Dialog.MODE_INFO, owner=owner),
                message, title, details)


def show_warning(message, title="Warning", details=None, owner=None):
    """Show a warning modal dialog."""
    return _run(lambda: T3Dialog(message, title=title, details=details,
                                 mode=T3Dialog.MODE_WARNING, owner=owner),
                message, title, details)


def show_error(message, title="Error", details=None, owner=None):
    """Show an error modal dialog."""
    return _run(lambda: T3Dialog(message, title=title, details=details,
                                 mode=T3Dialog.MODE_ERROR, ok_text="Close",
                                 owner=owner),
                message, title, details)


def confirm(message, title="Confirm Action", ok_text="Proceed", cancel_text="Cancel",
            danger=False, details=None, owner=None):
    """Show a confirmation dialog. Returns True if confirmed, False otherwise."""
    return _run(lambda: T3Dialog(message, title=title, details=details,
                                 mode=T3Dialog.MODE_CONFIRM, ok_text=ok_text,
                                 cancel_text=cancel_text, danger=danger,
                                 owner=owner),
                message, title, details, buttons=2)


def confirm_options(message, options, title="Confirm Action", ok_text="OK",
                    cancel_text="Cancel", details=None, owner=None, subtitle=None,
                    sections=None, sections_title="Details"):
    """Confirm with checkboxes, one row each with what it does.
    `options` = [(key, label, description, checked[, requires])], requires =
    key of the option this one depends on. `sections` = [(heading, [line]
    [, glyph])]: long reading, folded under `sections_title`; it replaces
    `details`, which stays the plain text of the TaskDialog fallback.
    Returns {key: bool} on OK, None on Cancel. Without the WPF dialog it
    falls back to a Yes/No TaskDialog and returns the given defaults."""
    try:
        dlg = T3Dialog(message, title=title, details=details, mode=T3Dialog.MODE_CONFIRM,
                       ok_text=ok_text, cancel_text=cancel_text, owner=owner, options=options,
                       subtitle=subtitle, sections=sections, sections_title=sections_title)
        dlg.ShowDialog()
        return dlg.option_values if dlg.result else None
    except Exception as exc:
        try:
            note = (details + "\n\n" if details else "") + "(T3 dialog unavailable: %s)" % exc
            if _task_dialog(message, title, note, buttons=2):
                return dict((o[0], bool(o[3])) for o in options)
        except Exception:
            pass
        return None


def ask_for_string(message, default=u"", title="Input", owner=None):
    """One line of text from the user, or None on Cancel / no dialog."""
    try:
        dlg = T3Dialog(message, title=title, mode=T3Dialog.MODE_PROMPT,
                       ok_text="OK", owner=owner, default_text=default)
        dlg.ShowDialog()
        return dlg.value if dlg.result else None
    except Exception:
        return None
