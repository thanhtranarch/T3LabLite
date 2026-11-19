# -*- coding: utf-8 -*-
"""FamiGen Studio host: the web UI in pyRevit's WebView2 output window.

Requires pyRevit 7.0+ (the output window renders with WebView2/Chromium) and,
for the page <-> Python messages, 7.1+ (``output.on_message`` /
``output.post_message``). ``requirement()`` and ``open_studio()`` check both
and say exactly what to update.

Pieces (docs/famigen-studio-spec.md section 4):

* ``build_page()``  - one self-contained HTML file (CSS, three.js r158 and the
  Studio script inlined) in ``%APPDATA%\\T3LabAI\\famigen_studio``, so nothing
  depends on how WebView2 resolves relative file URLs.
* ``StudioController`` (pure) - the message protocol, Claude session, preview.
* ``DotNetPoster`` - Claude requests in .NET, polled on the UI thread.
* ``_StudioAction`` - ExternalEvent that builds the family when the window is
  modeless (persistent engine). Without a persistent engine the command keeps
  running in a nested dispatcher frame until the window closes, with Revit's
  main window disabled like a modal dialog, so Create Family runs directly.
"""
import io
import json
import os
import traceback
import uuid

import clr

clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System import TimeSpan                                         # noqa: E402
from System.Windows.Threading import (Dispatcher, DispatcherFrame,  # noqa: E402
                                      DispatcherTimer)
from Autodesk.Revit.UI import ExternalEvent, IExternalEventHandler  # noqa: E402

from FamilyGen import studio_controller as sc                       # noqa: E402
from FamilyGen.studio_page import (MESSAGING_MESSAGE, build_page,   # noqa: E402
                                   page_folder, pyrevit_version, requirement,
                                   version_text)

WINDOW_SIZE = (1440, 900)

_ACTIVE = {'studio': None}       # keeps the modeless Studio alive (GC) and single


# ── adapters ─────────────────────────────────────────────────────────────────

class _Prefs(object):
    """Small JSON settings file: folder, load, category, model, effort."""

    def __init__(self, path=None):
        self.path = path or os.path.join(page_folder(), 'famigen_studio.json')
        self._data = {}
        try:
            with io.open(self.path, 'r', encoding='utf-8') as handle:
                data = json.loads(handle.read())
            if isinstance(data, dict):
                self._data = data
        except Exception:
            self._data = {}

    def get(self, key, default=None):
        return self._data.get(key, default)

    def set(self, key, value):
        self._data[key] = value

    def save(self):
        folder = os.path.dirname(self.path)
        if not os.path.isdir(folder):
            os.makedirs(folder)
        with io.open(self.path, 'w', encoding='utf-8') as handle:
            handle.write(json.dumps(self._data, ensure_ascii=False, indent=2))


class _ClaudeAdapter(object):
    """Key and live model list from LLMs Setting's Claude provider."""

    def __init__(self):
        self._provider = None
        self._models = None

    def _p(self):
        if self._provider is None:
            from Intelligence.claude_provider import ClaudeProvider
            self._provider = ClaudeProvider()
        return self._provider

    def api_key(self):
        try:
            from config.settings import T3LabAISettings
            return T3LabAISettings().get_api_key('Claude')
        except Exception:
            return None

    def models(self):
        if self._models is None:
            self._models = list(self._p().get_models() or [])
        return self._models

    def default_model(self):
        # Family design is geometry reasoning: the quality role (Opus first,
        # or the model pinned for that role in LLMs Setting).
        try:
            return self._p().pick_quality_model()
        except Exception:
            return None


def _make_action_handler():
    # A new .NET type name per load: this module can be imported again in the
    # same engine (pyRevit reload, a dev reload), and PythonNet refuses a second
    # type with the same name ("Duplicate type name within an assembly").
    class _StudioAction(IExternalEventHandler):
        __namespace__ = 'T3Lab.FamiGenStudio_' + uuid.uuid4().hex[:8]

        def __init__(self):
            self.action = None

        def Execute(self, uiapp):
            action, self.action = self.action, None
            if action is not None:
                action(uiapp)

        def GetName(self):
            return 'T3Lab FamiGen Studio'

    return _StudioAction()


class _RevitAdapter(object):
    def __init__(self, doc, app, output, modal):
        self.doc = doc
        self.app = app
        self.output = output
        self.modal = modal
        self.into_family = bool(getattr(doc, 'IsFamilyDocument', False)) and modal
        self.frameless = False      # set by open_studio once pyRevit's own title bar is hidden
        self._handler = None
        self._event = None
        if not modal:
            self._handler = _make_action_handler()
            self._event = ExternalEvent.Create(self._handler)   # needs API context: here it has

    # dialogs (UI thread; no API context needed)
    def pick_files(self):
        from Microsoft.Win32 import OpenFileDialog
        dlg = OpenFileDialog()
        dlg.Title = 'Reference images or PDF spec sheets'
        dlg.Filter = ('Images and PDF|*.png;*.jpg;*.jpeg;*.gif;*.webp;*.pdf|'
                      'Images|*.png;*.jpg;*.jpeg;*.gif;*.webp|PDF|*.pdf')
        dlg.Multiselect = True
        return list(dlg.FileNames) if dlg.ShowDialog() else []

    def pick_json_file(self):
        from Microsoft.Win32 import OpenFileDialog
        dlg = OpenFileDialog()
        dlg.Title = 'Open a family JSON'
        dlg.Filter = 'Family JSON|*.json|All files|*.*'
        return dlg.FileName if dlg.ShowDialog() else None

    def save_json_file(self, name, text):
        from Microsoft.Win32 import SaveFileDialog
        dlg = SaveFileDialog()
        dlg.Title = 'Save the family JSON'
        dlg.Filter = 'Family JSON|*.json'
        dlg.FileName = u'{}.json'.format(name)
        if not dlg.ShowDialog():
            return None
        with io.open(dlg.FileName, 'w', encoding='utf-8') as handle:
            handle.write(text)
        return dlg.FileName

    def pick_folder(self, initial):
        try:
            from pyrevit import forms
            return forms.pick_folder(title='Choose the folder to save families in')
        except Exception:
            return None

    def copy_text(self, text):
        from System.Windows import Clipboard
        Clipboard.SetText(text)

    def close(self):
        try:
            self.output.close()
        except Exception:
            pass

    def _window(self):
        # pyRevit 7's output window (a WPF Window) behind the public wrapper;
        # the title-bar buttons of the page do nothing if it cannot be reached.
        try:
            return self.output._runtime_output().window
        except Exception:
            return getattr(self.output, 'window', None)

    def window_state(self, action):
        from System.Windows import WindowState
        win = self._window()
        if win is None:
            return
        if action == 'drag':
            _drag_window(win)
        elif action == 'minimize':
            win.WindowState = WindowState.Minimized
        elif win.WindowState == WindowState.Maximized:
            win.WindowState = WindowState.Normal
        else:
            win.WindowState = WindowState.Maximized

    # Revit API
    def create_family(self, schema, folder, load, callback):
        def _work(uiapp):
            from FamilyGen import builder
            doc, app = self.doc, self.app
            if uiapp is not None:
                try:
                    if uiapp.ActiveUIDocument is not None:
                        doc = uiapp.ActiveUIDocument.Document
                    app = uiapp.Application
                except Exception:
                    pass
            try:
                if self.into_family:
                    report = builder.build_into_document(doc, schema)
                else:
                    report = builder.create_family(
                        app, schema, folder, project_doc=doc,
                        load_into_project=load and not doc.IsFamilyDocument)
            except builder.FamilyBuildError as ex:
                callback(None, u'{}'.format(ex))
                return
            except Exception as ex:
                callback(None, u'Revit could not create "{}": {}. Simplify the part named in '
                               u'the warnings and try again.'.format(schema.get('family_name'), ex))
                return
            callback(report, None)

        if self.modal or self._event is None:
            _work(None)
        else:
            self._handler.action = _work
            self._event.Raise()


# ── window chrome ────────────────────────────────────────────────────────────
# pyRevit's output window is a MahApps MetroWindow with its own dark title bar
# (verified in pyRevit's ScriptConsole.cs). The Studio page draws the T3 title
# bar, so the host hides pyRevit's; when that is not possible the page hides
# its own instead (init.frameless = False) - never two title bars.

def _go_frameless(win):
    """Hide the output window's own title bar and buttons. True when it worked."""
    if win is None:
        return False
    try:
        win.UseNoneWindowStyle = True        # MetroWindow: no title bar, no window buttons
        win.ShowTitleBar = False
    except Exception:
        return False
    for name in ('ShowCloseButton', 'ShowMinButton', 'ShowMaxRestoreButton'):
        try:
            setattr(win, name, False)
        except Exception:
            pass
    try:
        from System.Windows import Visibility
        for name in ('LeftWindowCommands', 'RightWindowCommands'):
            commands = getattr(win, name, None)
            if commands is not None:
                commands.Visibility = Visibility.Collapsed
    except Exception:
        pass
    return True


def _drag_window(win):
    """Start a native window move (the page's title bar was pressed).

    The press happened inside WebView2, so WPF's DragMove refuses it; handing
    the window a caption click is the usual way to move a frameless WebView2
    window."""
    try:
        import ctypes
        from System.Windows.Interop import WindowInteropHelper
        hwnd = WindowInteropHelper(win).Handle.ToInt64()
        user32 = ctypes.windll.user32
        user32.SendMessageW.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t]
        user32.ReleaseCapture()
        user32.SendMessageW(ctypes.c_void_p(hwnd), 0x00A1, 2, 0)     # WM_NCLBUTTONDOWN, HTCAPTION
    except Exception:
        pass


# ── open ─────────────────────────────────────────────────────────────────────

def _window_closed(output):
    try:
        return not output.supports_messages
    except Exception:
        return True


def _set_revit_enabled(uiapp, enabled):
    try:
        import ctypes
        hwnd = uiapp.MainWindowHandle.ToInt64()
        ctypes.windll.user32.EnableWindow(ctypes.c_void_p(hwnd), bool(enabled))
    except Exception:
        pass


def detect_persistent_engine():
    """True when this command runs on a persistent engine; a modeless Studio
    needs one, otherwise pyRevit drops the interpreter when the script ends.

    Its own copy on purpose. The button's script used to import the one in
    GUI.ManaSelectDialog, and Services/tool_discovery reads every
    `from GUI.<X>Dialog import` in a script as "this button opens X": FamiGen
    picked up Mana Select's keywords and the Assistant could no longer tell
    "Mana Select" from FamiGen (2026-11-19).
    """
    try:
        from pyrevit import EXEC_PARAMS
        try:
            return bool(EXEC_PARAMS.needs_persistent_engine)
        except Exception:
            pass
        cfg_json = EXEC_PARAMS.script_runtime.ScriptRuntimeConfigs.EngineConfigs
        if cfg_json:
            return bool(json.loads(cfg_json).get('persistent', False))
    except Exception:
        pass
    return False


def _pyrevit_output():
    """pyRevit's real output window - never `script.get_output()` here.

    Under CPython the bootstrap swaps `get_output` in pyrevit.script and
    pyrevit.output for a silent stand-in, because no other T3Lab tool uses the
    output window (_cpython_bootstrap.install_script_shim). That stand-in
    answers every attribute, `on_message` and `open_page` included, so the
    Studio passed its own checks and then opened nothing: on pyRevit 7 it
    could only ever work on the first CPython click of a Revit session, before
    the swap. `PyRevitOutputWindow()` is what the unswapped `get_output()`
    returns (pyRevit 6.5.5 and 7.1.0 sources).
    """
    from pyrevit import output as pyrevit_output
    return pyrevit_output.PyRevitOutputWindow()


def open_studio(doc, uiapp, persistent):
    """Open FamiGen Studio. Returns (state, message).

    state: 'opened' (modeless), 'closed' (modal run finished), or 'unsupported'
    with the message to show. ``persistent`` = the command runs on a
    persistent engine (``detect_persistent_engine``).
    """
    ok, message = requirement()
    if not ok:
        return 'unsupported', message
    output = _pyrevit_output()
    if not hasattr(output, 'on_message') or not hasattr(output, 'post_message'):
        return 'unsupported', MESSAGING_MESSAGE.format(version_text(pyrevit_version()))

    current = _ACTIVE['studio']
    if current is not None and not _window_closed(current['output']):
        return 'unsupported', (u'FamiGen Studio is already open. Switch to its window '
                               u'(it may be behind Revit).')

    try:
        output.set_title('FamiGen Studio')
        output.set_width(WINDOW_SIZE[0])
        output.set_height(WINDOW_SIZE[1])
    except Exception:
        pass
    output.open_page(build_page())

    modal = not persistent
    revit = _RevitAdapter(doc, doc.Application, output, modal)
    prefs = _Prefs()
    # "frameless": false in famigen_studio.json keeps pyRevit's own title bar.
    if prefs.get('frameless', True) is not False:
        revit.frameless = _go_frameless(revit._window())
    from FamilyGen.net_http import DotNetPoster
    from FamilyGen.proposals import ProposalStore
    poster = DotNetPoster()
    controller = sc.StudioController(
        post=lambda msg: output.post_message(msg), http=poster, revit=revit,
        claude=_ClaudeAdapter(), prefs=prefs, store=ProposalStore(),
        pyrevit_version=version_text(pyrevit_version()))

    def _on_message(message):
        try:
            controller.handle(message)
        except Exception:
            traceback.print_exc()

    if not output.on_message(_on_message):
        poster.dispose()
        return 'unsupported', MESSAGING_MESSAGE.format(version_text(pyrevit_version()))

    state = {'output': output, 'controller': controller, 'poster': poster, 'revit': revit}
    watch = DispatcherTimer()
    watch.Interval = TimeSpan.FromMilliseconds(1000)
    frame = DispatcherFrame() if modal else None

    def _watch(sender, args):
        if not _window_closed(output):
            try:
                controller.tick()       # heartbeat, AI app over MCP, new proposals
            except Exception:
                traceback.print_exc()
            return
        watch.Stop()
        poster.dispose()
        controller.close_link()
        if _ACTIVE['studio'] is state:
            _ACTIVE['studio'] = None
        if frame is not None:
            frame.Continue = False

    watch.Tick += _watch
    state['watch'] = watch
    _ACTIVE['studio'] = state
    watch.Start()
    if not modal:
        return 'opened', u''

    # No persistent engine: keep this command (and its API context) alive
    # until the window closes, and block Revit's own window like a modal
    # dialog so no other command starts meanwhile.
    _set_revit_enabled(uiapp, False)
    try:
        Dispatcher.PushFrame(frame)
    finally:
        _set_revit_enabled(uiapp, True)
        watch.Stop()
        poster.dispose()
        controller.close_link()
        _ACTIVE['studio'] = None
    return 'closed', u''
