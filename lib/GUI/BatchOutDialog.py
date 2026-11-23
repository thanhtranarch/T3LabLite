# -*- coding: utf-8 -*-
"""BatchOut Dialog — Advanced batch exporter for sheets and views to PDF, DWG, DWF, NWD, IFC.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__  = "Tran Tien Thanh"
__title__   = "BatchOut Dialog"

# IMPORT LIBRARIES
import os
import sys

# CPython engine paths come from lib/_cpython_bootstrap.py below.


_script_dir = os.path.dirname(__file__)
_ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(_script_dir)))
_lib_dir = os.path.join(_ext_dir, 'lib')
if _lib_dir not in sys.path:
    sys.path.insert(0, _lib_dir)

try:
    import _cpython_bootstrap
    _cpython_bootstrap.init_cpython_paths()   # also hot-reloads an edited WPF_Base
except Exception:
    pass


import clr
import json
import re
import time
from datetime import datetime
from collections import defaultdict

clr.AddReference('System.Windows.Forms')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')
clr.AddReference('System')
from System.Windows.Forms import (FolderBrowserDialog, DialogResult,
                                  OpenFileDialog, SaveFileDialog)
from System.Windows import Visibility, WindowState
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, UriKind, Action, GC
from System.ComponentModel import INotifyPropertyChanged, PropertyChangedEventArgs
from System.Threading import Thread, ThreadStart
from System.Windows.Threading import DispatcherPriority

from pyrevit import revit, DB, UI, forms, script, EXEC_PARAMS
from GUI.WPF_Base import T3WPFWindow, to_items_source
from GUI.InputLock import ThreadInputLock
from Autodesk.Revit.DB import (
    Transaction, FilteredElementCollector, BuiltInCategory,
    ViewSheet, ViewSheetSet, DWGExportOptions, DWFExportOptions,
    DGNExportOptions, ExportDWGSettings, ACADVersion, PDFExportOptions,
    ImageExportOptions, ImageFileType, ImageResolution,
    PropOverrideMode, View, ViewPlan, ViewSection, View3D,
    ViewSchedule, ViewDrafting, ViewType,
)
from Autodesk.Revit.UI import IExternalEventHandler, ExternalEvent

from System.Collections.Generic import List


extension_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
lib_dir = os.path.join(extension_dir, 'lib')
if lib_dir not in sys.path:
    sys.path.append(lib_dir)

from Snippets._compat import disposing, eid_value, make_eid
from core.paths import user_data_path
from Services import batchout_report, batchout_profiles, batchout_refs, batchout_pdf24
from Services.SheetManager.viewsheet_sets_service import ViewSheetSetsService
from Services.batchout_report import DEFAULT_COLUMNS as REPORT_DEFAULT_COLUMNS

try:
    from Intelligence.api_learner import SmartAPIAdapter, RevitAPILearner
    HAS_API_LEARNER = True
    _api_learner_err = None
except Exception as _api_learner_ex:
    HAS_API_LEARNER = False
    _api_learner_err = _api_learner_ex


try:
    from Autodesk.Revit.DB import IFCExportOptions, IFCVersion
    HAS_IFC = True
except:
    HAS_IFC = False

try:
    from Autodesk.Revit.DB import NavisworksExportOptions
    HAS_NAVISWORKS = True
except:
    HAS_NAVISWORKS = False

# DEFINE VARIABLES
logger = script.get_logger()
# CPython has no ScriptOutput.GetDefault; safe_output() returns a no-op window.
try:
    from _cpython_bootstrap import safe_output
    output = safe_output()
except Exception:
    output = script.get_output()


# ── Machine state, not user files ──
# Profiles stay in Documents\T3Lab_BatchOut_Profiles (user's shareable files).
# Crash breadcrumbs are per-machine state under %APPDATA%\T3LabAI\batchout: the
# marker is fsync'd around EVERY native export call = one OneDrive upload per
# sheet in Documents. Bad Geometry findings (checks/badgeometry_check.py) live in
# %APPDATA%\T3LabAI\diagnostics for the same reason; dev/test_user_data_moves.py
# keeps both sides on the same path.
CRASH_MARKER_NAME = '_export_crash_marker.json'
CRASH_HISTORY_NAME = '_export_crash_history.json'
LEGACY_DIAG_FOLDER = os.path.join(os.path.expanduser('~'), 'Documents', 'T3Lab_Diagnostics')
BAD_GEOMETRY_NAME = '_bad_geometry.json'


def _adopt_legacy_file(old, new):
    """Move a TRANSIENT state file an older build left behind to its new home.

    Moved, not copied: a copy would come back from the old folder after every
    delete and report the same old crash each time. Never raises.
    """
    try:
        if old and os.path.isfile(old) and not os.path.exists(new):
            import shutil
            shutil.move(old, new)
    except Exception:
        pass


def crash_state_files(profiles_folder):
    """(marker, history) paths of the native-crash breadcrumbs.

    History is durable, so an older one is copied once; the transient marker
    is moved.
    """
    history = user_data_path(
        'batchout', CRASH_HISTORY_NAME,
        legacy=[os.path.join(profiles_folder, CRASH_HISTORY_NAME)])
    marker = user_data_path('batchout', CRASH_MARKER_NAME)
    _adopt_legacy_file(os.path.join(profiles_folder, CRASH_MARKER_NAME), marker)
    return marker, history


def bad_geometry_file():
    """Findings of the Bad Geometry check, read by safe mode.

    Same path as checks/badgeometry_check.py FINDINGS_FILE; an older file is
    copied once.
    """
    return user_data_path(
        'diagnostics', BAD_GEOMETRY_NAME,
        legacy=[os.path.join(LEGACY_DIAG_FOLDER, BAD_GEOMETRY_NAME)])


def _build_view_type_labels():
    """Display label per ViewType (getattr: a member missing in some Revit
    build never breaks import). Also the `view_type_filter` entries, built by
    _fill_view_type_filter() from the loaded views.
    """
    labels = {}
    for name, label in (
            ("FloorPlan", "Floor Plan"),
            ("CeilingPlan", "Ceiling Plan"),
            ("Elevation", "Elevation"),
            ("Section", "Section"),
            ("ThreeD", "3D View"),
            ("Schedule", "Schedule"),
            ("DraftingView", "Drafting"),
            ("Legend", "Legend"),
            ("EngineeringPlan", "Engineering"),
            ("AreaPlan", "Area Plan"),
            ("Detail", "Detail"),
            ("Rendering", "Rendering"),
            ("Walkthrough", "Walkthrough"),
            ("Report", "Report"),
            ("ColumnSchedule", "Column Schedule"),
            ("PanelSchedule", "Panel Schedule"),
            ("CostReport", "Cost Report"),
            ("LoadsReport", "Loads Report"),
            ("PresureLossReport", "Pressure Loss Report"),
    ):
        member = getattr(ViewType, name, None)
        if member is not None:
            labels[member] = label
    return labels


def _build_non_exportable_view_types():
    """ViewTypes kept out of the Views list: internal/browser pseudo-views and
    sheets (own mode). Everything else (legends, schedules, renderings) is listed."""
    types = []
    for name in ("Internal", "ProjectBrowser", "SystemBrowser",
                 "Undefined", "DrawingSheet"):
        member = getattr(ViewType, name, None)
        if member is not None:
            types.append(member)
    return set(types)


VIEW_TYPE_LABELS = _build_view_type_labels()
NON_EXPORTABLE_VIEW_TYPES = _build_non_exportable_view_types()


# DWG "Match Revit display": shared with the MCP / Assistant export (Services/sheet_export.py).
from Services.sheet_export import _enum_member, apply_revit_display_fidelity  # noqa: E402
from Services import sheet_export as _sheet_export                           # noqa: E402
_sheet_export.logger = logger     # DWG option warnings to the pyRevit output, as before


def detect_persistent_engine():
    """True only when __persistentengine__ is baked into the loaded command
    metadata (pyRevit bakes it at ribbon build — needs a Reload).

    Accessor differs per pyRevit runtime build (2026-07-15: needs_persistent_engine
    raised AttributeError on a 2023 runtime):
    - newer: EXEC_PARAMS.needs_persistent_engine (EngineConfigs.PersistentEngine)
    - older: script_runtime.ScriptRuntimeConfigs.EngineConfigs, JSON with "persistent"
    Defaults to False (modal fallback) when undetectable.
    """
    try:
        return bool(EXEC_PARAMS.needs_persistent_engine)
    except Exception:
        pass
    try:
        cfg_json = EXEC_PARAMS.script_runtime.ScriptRuntimeConfigs.EngineConfigs
        if cfg_json:
            return bool(json.loads(cfg_json).get('persistent', False))
    except Exception as ex:
        logger.debug("Persistent engine detection failed: {}".format(ex))
    return False

# F6: the api_learner import ran before `logger` existed; surface its failure once.
if not HAS_API_LEARNER:
    logger.warning(
        "BatchOut Intelligence helper unavailable (api_learner); "
        "smart API adaptation disabled: {}".format(_api_learner_err))

# From the Application, not the document: `revit.doc` is None from the Assistant
# pane / with no project open, and failed at import time.
from Snippets._host import get_revit_version
REVIT_VERSION = get_revit_version()

# CLASS/FUNCTIONS
try:
    _Reactive = getattr(forms, 'Reactive', object)
except Exception:
    _Reactive = object


class SheetItem(_Reactive):
    """Sheet row in the list."""
    def __init__(self, sheet, is_selected=False, lazy=False):
        self._property_changed_handlers = []
        self.Sheet = sheet
        self._is_selected = bool(is_selected)
        self.SheetNumber = sheet.SheetNumber
        self.SheetName = sheet.Name
        self.Status = "Ready"
        self.Progress = 0
        self.Size = "-"
        self.Orientation = "-"
        self.Revision = ""
        self.RevisionDate = ""
        self.RevisionDescription = ""
        self.DrawnBy = ""
        self.CheckedBy = ""
        self.CustomFilename = ""

        if not lazy:
            self._load_revision_params()

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        val = bool(value)
        if self._is_selected != val:
            self._is_selected = val
            self.OnPropertyChanged("IsSelected")

    def add_PropertyChanged(self, handler):
        if handler not in self._property_changed_handlers:
            self._property_changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)

    def OnPropertyChanged(self, property_name):
        try:
            if PropertyChangedEventArgs is not None:
                args = PropertyChangedEventArgs(property_name)
                for handler in list(self._property_changed_handlers):
                    try:
                        handler(self, args)
                    except Exception:
                        pass
        except Exception:
            pass

    def _load_revision_params(self):
        """Load revision and metadata parameters. Called deferred for fast startup."""
        try:
            rev_param = self.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION)
            self.Revision = rev_param.AsString() if rev_param else ""
        except:
            self.Revision = ""

        try:
            rev_date_param = self.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION_DATE)
            self.RevisionDate = rev_date_param.AsString() if rev_date_param else ""
        except:
            self.RevisionDate = ""

        try:
            rev_desc_param = self.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION_DESCRIPTION)
            self.RevisionDescription = rev_desc_param.AsString() if rev_desc_param else ""
        except:
            self.RevisionDescription = ""

        try:
            drawn_param = self.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_DRAWN_BY)
            self.DrawnBy = drawn_param.AsString() if drawn_param else ""
        except:
            self.DrawnBy = ""

        try:
            checked_param = self.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_CHECKED_BY)
            self.CheckedBy = checked_param.AsString() if checked_param else ""
        except:
            self.CheckedBy = ""

    def __repr__(self):
        return "{} - {}".format(self.SheetNumber, self.SheetName)


class ViewItem(_Reactive):
    """View row in the list."""

    _VIEW_TYPE_MAP = VIEW_TYPE_LABELS

    def __init__(self, view, is_selected=False, lazy=False):
        self._property_changed_handlers = []
        self.View = view
        self._is_selected = bool(is_selected)
        self.ViewName = view.Name
        self.SheetNumber = view.Name  # alias for column compatibility
        self.SheetName = ""

        # ViewType / Scale are direct properties — cheap, always loaded
        try:
            self.ViewType = self._VIEW_TYPE_MAP.get(view.ViewType, str(view.ViewType))
        except:
            self.ViewType = "Unknown"
        self.SheetName = self.ViewType

        try:
            self.Scale = "1:{}".format(view.Scale) if hasattr(view, 'Scale') and view.Scale else "-"
        except:
            self.Scale = "-"

        self.Phase = "-"
        self.ViewTemplate = "-"
        if not lazy:
            self._load_extra_data()

        self.Status = "Ready"
        self.Progress = 0
        self.CustomFilename = ""
        self.Size = self.Scale
        self.Revision = self.Phase
        self.Orientation = "-"

    @property
    def IsSelected(self):
        return self._is_selected

    @IsSelected.setter
    def IsSelected(self, value):
        val = bool(value)
        if self._is_selected != val:
            self._is_selected = val
            self.OnPropertyChanged("IsSelected")

    def add_PropertyChanged(self, handler):
        if handler not in self._property_changed_handlers:
            self._property_changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)

    def OnPropertyChanged(self, property_name):
        try:
            if PropertyChangedEventArgs is not None:
                args = PropertyChangedEventArgs(property_name)
                for handler in list(self._property_changed_handlers):
                    try:
                        handler(self, args)
                    except Exception:
                        pass
        except Exception:
            pass

    def _load_extra_data(self):
        """Load Phase and ViewTemplate — deferred for startup performance."""
        try:
            phase_param = self.View.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
            if phase_param:
                phase_id = phase_param.AsElementId()
                if phase_id and phase_id != DB.ElementId.InvalidElementId:
                    phase_elem = self.View.Document.GetElement(phase_id)
                    self.Phase = phase_elem.Name if phase_elem else "-"
        except:
            pass

        try:
            template_id = self.View.ViewTemplateId
            if template_id and template_id != DB.ElementId.InvalidElementId:
                template = self.View.Document.GetElement(template_id)
                self.ViewTemplate = template.Name if template else "-"
        except:
            pass

        self.Revision = self.Phase  # keep alias in sync

    def __repr__(self):
        return "{} ({})".format(self.ViewName, self.ViewType)


# Queue row states: ExportManager.xaml colours the PROGRESS cell by this exact
# text (dev/test_batchout_live_progress.py).
QUEUE_QUEUED = "Queued"
QUEUE_EXPORTING = "Exporting..."
QUEUE_EXPORTED = "Exported"
QUEUE_FAILED = "Failed"
QUEUE_SKIPPED = "Skipped"
QUEUE_NOT_EXPORTED = "Not exported"


def _clock(seconds):
    """mm:ss, or h:mm:ss from one hour on."""
    seconds = int(max(0, seconds))
    hours, rest = divmod(seconds, 3600)
    minutes, secs = divmod(rest, 60)
    if hours:
        return "{}:{:02d}:{:02d}".format(hours, minutes, secs)
    return "{:02d}:{:02d}".format(minutes, secs)


def format_progress_summary(done, total, percent, elapsed=None, running=False):
    """Right-hand label under the queue: '3 / 12 · 25% · 00:41 · ~02:03 left'.

    The ETA uses the average of finished items; shown only once one is done and
    while running.
    """
    parts = []
    if total:
        parts.append("{} / {}".format(done, total))
    if total or percent > 0:
        parts.append("{}%".format(percent))
    if elapsed is not None:
        parts.append(_clock(elapsed))
        if running and total and 0 < done < total:
            parts.append("~{} left".format(_clock(elapsed * (total - done) / float(done))))
    return u" \u00b7 ".join(parts)


class ExportPreviewItem(object):
    """Export preview row for a SheetItem or ViewItem."""
    def __init__(self, item, format_name, size, orientation):
        self._property_changed_handlers = []
        if hasattr(item, 'SheetNumber'):
            self.SheetNumber = item.SheetNumber
            self.SheetName = item.SheetName
            self.ItemName = "{} - {}".format(item.SheetNumber, item.SheetName)
        elif hasattr(item, 'ViewName'):
            self.SheetNumber = item.ViewName
            self.SheetName = item.ViewType
            self.ItemName = "{} ({})".format(item.ViewName, item.ViewType)
        else:
            self.SheetNumber = "Unknown"
            self.SheetName = "Unknown"
            self.ItemName = "Unknown"

        # Views alias Revision to Phase, so only a sheet's revision is reported.
        self.Revision = getattr(item, 'Revision', '') if hasattr(item, 'Sheet') else ''
        self.OutputFile = ''  # set on a confirmed export; read by the report
        self.Format = format_name
        self.Size = size
        self.Orientation = orientation
        self._progress = 0
        self._status = ""
        self.ProgressText = ""

    @property
    def Progress(self):
        return self._progress

    @Progress.setter
    def Progress(self, value):
        if self._progress != value:
            self._progress = value
            self.OnPropertyChanged("Progress")

    @property
    def Status(self):
        return self._status

    @Status.setter
    def Status(self, value):
        if self._status != value:
            self._status = value
            self.OnPropertyChanged("Status")

    def add_PropertyChanged(self, handler):
        if handler not in self._property_changed_handlers:
            self._property_changed_handlers.append(handler)

    def remove_PropertyChanged(self, handler):
        if handler in self._property_changed_handlers:
            self._property_changed_handlers.remove(handler)

    def OnPropertyChanged(self, property_name):
        try:
            if PropertyChangedEventArgs is not None:
                args = PropertyChangedEventArgs(property_name)
                for handler in list(self._property_changed_handlers):
                    try:
                        handler(self, args)
                    except Exception:
                        pass
        except Exception:
            pass


class ExportProfile(object):
    """Export profile with all settings."""
    def __init__(self, name="", description=""):
        self.Name = name
        self.Description = description
        self.CreatedDate = datetime.now().strftime("%Y-%m-%d %H:%M")

        self.ExportPDF = True
        self.ExportDWG = False
        self.ExportDGN = False
        self.ExportDWF = False
        self.ExportNWD = False
        self.ExportIFC = False
        self.ExportIMG = False

        self.PDFPaperSize = "Use Sheet Size"
        self.PDFOrientation = "Landscape"
        self.PDFZoom = "Fit to Page"
        self.PDFHideRefPlanes = False
        self.PDFHideScopeBoxes = False
        self.PDFHideCropBoundaries = False
        self.PDFHideUnreferencedTags = False
        self.PDFViewLinksBlue = False
        self.CombinePDF = False
        # "Revit" (doc.Export) or "PDF24" (PDF24 Creator printer, Services/batchout_pdf24.py)
        self.PDFEngine = "Revit"

        self.DWGVersion = "AutoCAD 2013"
        self.CADExportSetup = "Use setup from file"
        self.CADExportViewsOnSheets = False
        self.CADExportLinksAsExternal = False
        # Colors / line weights / linetype scale as the Revit view shows them
        self.CADMatchRevitDisplay = True
        # No image files next to the DWG (logos are dropped from the DWG)
        self.CADSkipImageReferences = False

        self.OutputFolder = os.path.join(os.path.expanduser('~'), 'Documents', 'Revit Exports')
        self.SplitByFormat = False
        self.ReverseOrder = False

        self.NamingPattern = "{SheetNumber}"

        # Export report (Services/batchout_report.py); ReportColumns has no UI —
        # edit it in the profile JSON to pick/reorder columns.
        self.WriteReport = True
        self.ReportFormat = "XLSX"
        self.ReportColumns = list(REPORT_DEFAULT_COLUMNS)

    def to_dict(self):
        """Dict for JSON."""
        return {
            'Name': self.Name,
            'Description': self.Description,
            'CreatedDate': self.CreatedDate,
            'ExportPDF': self.ExportPDF,
            'ExportDWG': self.ExportDWG,
            'ExportDGN': self.ExportDGN,
            'ExportDWF': self.ExportDWF,
            'ExportNWD': self.ExportNWD,
            'ExportIFC': self.ExportIFC,
            'ExportIMG': self.ExportIMG,
            'PDFPaperSize': self.PDFPaperSize,
            'PDFOrientation': self.PDFOrientation,
            'PDFZoom': self.PDFZoom,
            'PDFHideRefPlanes': self.PDFHideRefPlanes,
            'PDFHideScopeBoxes': self.PDFHideScopeBoxes,
            'PDFHideCropBoundaries': self.PDFHideCropBoundaries,
            'PDFHideUnreferencedTags': self.PDFHideUnreferencedTags,
            'PDFViewLinksBlue': self.PDFViewLinksBlue,
            'CombinePDF': self.CombinePDF,
            'PDFEngine': self.PDFEngine,
            'DWGVersion': self.DWGVersion,
            'CADExportSetup': self.CADExportSetup,
            'CADExportViewsOnSheets': self.CADExportViewsOnSheets,
            'CADExportLinksAsExternal': self.CADExportLinksAsExternal,
            'CADMatchRevitDisplay': self.CADMatchRevitDisplay,
            'CADSkipImageReferences': self.CADSkipImageReferences,
            'OutputFolder': self.OutputFolder,
            'SplitByFormat': self.SplitByFormat,
            'ReverseOrder': self.ReverseOrder,
            'NamingPattern': self.NamingPattern,
            'WriteReport': self.WriteReport,
            'ReportFormat': self.ReportFormat,
            'ReportColumns': list(self.ReportColumns),
        }

    @staticmethod
    def from_dict(data):
        """Create profile from a dict; rejects non-dicts (the profiles folder
        also holds bookkeeping JSON, e.g. the crash history list)."""
        if not isinstance(data, dict):
            raise ValueError(
                "not a profile object (got {})".format(type(data).__name__))
        profile = ExportProfile()
        for key, value in data.items():
            if hasattr(profile, key):
                setattr(profile, key, value)
        return profile


from Snippets._compat import clr_namespace   # per-load .NET type name (CLAUDE.md rule 5)
try:
    from Services.batchout_handler import BatchOutEventHandler
except Exception:
    import uuid
    class BatchOutEventHandler(IExternalEventHandler):
        __namespace__ = clr_namespace('BatchOut')
        def __init__(self, logger):
            self._logger = logger
            self._queue = []

        def add(self, action):
            self._queue.append(action)

        def Execute(self, uiapp):
            try:
                actions = self._queue
                self._queue = []
                for action in actions:
                    try:
                        action()
                    except Exception as ex:
                        try:
                            self._logger.error(
                                "BatchOut external event failed: {}".format(ex))
                        except Exception:
                            pass
            except Exception:
                pass

        def GetName(self):
            return "T3Lab BatchOut Handler"


class _PendingExportError(RuntimeError):
    """Revit still owns a pending transaction; later exports must not run."""


class ExportManagerWindow(T3WPFWindow):
    """Export Manager Window."""
    # Click= trong DataTemplate nằm ngoài namescope của window: thiếu cờ này
    # handler của từng dòng không bao giờ chạy.
    WIRE_TEMPLATED_CLICKS = True

    # Auto-saved "latest setup" — lives beside profiles but is not listed as one
    LATEST_SETUP_FILENAME = '_latest_setup.json'

    def __init__(self):
        try:
            xaml_file_path = os.path.join(os.path.dirname(__file__), 'Tools', 'ExportManager.xaml')
            if not os.path.exists(xaml_file_path):
                ext_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__))))
                xaml_file_path = os.path.join(ext_dir, 'lib', 'GUI', 'Tools', 'ExportManager.xaml')
            T3WPFWindow.__init__(self, xaml_file_path)

            self.doc = revit.doc

            # Modeless only when pyRevit keeps this engine alive; otherwise
            # callbacks hit torn-down globals → fatal crash (2026-07-14,
            # journal 1006). See detect_persistent_engine().
            self.modeless = detect_persistent_engine()

            # ExternalEvent used ONLY on the modeless path (_run_in_api_context);
            # modal marshals through the Dispatcher.
            self._api_handler = BatchOutEventHandler(logger)
            self._api_event = None
            if self.modeless:
                # Create() throws outside API context (Assistant pane handler or
                # its direct-export thread) — degrade to modal instead of dying.
                # Mirrors server.ensure_external_event().
                try:
                    self._api_event = ExternalEvent.Create(self._api_handler)
                except BaseException as ee_ex:
                    logger.debug(
                        "ExternalEvent.Create unavailable (window constructed "
                        "outside API context) — falling back to modal: "
                        "{}".format(ee_ex))
                    self.modeless = False

            self.all_sheets = []
            self.filtered_sheets = []
            self.all_views = []
            self.filtered_views = []
            self.export_items = []
            self._export_running = False
            self.selection_mode = "sheets"  # "sheets" or "views"
            self.profiles = []  # List of ExportProfile objects
            self._filling_profiles = False
            self._report_columns = list(REPORT_DEFAULT_COLUMNS)
            self._failure_reasons = {}
            self._last_output_file = ''
            self.profiles_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'T3Lab_BatchOut_Profiles')
            self.latest_setup_file = os.path.join(self.profiles_folder, self.LATEST_SETUP_FILENAME)

            # Native-crash breadcrumb: a marker surviving a native export call ==
            # that sheet+format killed Revit; confirmed kills go into a durable
            # history so they can be skipped. Paths: crash_state_files().
            self._crash_marker_file, self._crash_history_file = \
                crash_state_files(self.profiles_folder)
            self._crash_history = self._read_crash_history()
            self._pending_crash = self._promote_crash_marker()
            self._skipped_fatal = []
            # Safe-mode exports, reported at the end (not a normal export's file).
            self._safe_applied = []
            self._safe_mode = False
            # Known-fatal items export anyway by default (a silently missing file
            # is worse than a warned crash); only the pre-export prompt sets this.
            self._skip_fatal = False
            # Every item+format that produced NO file, with reason — shown in the
            # completion dialog.
            self._failed_items = []
            # Filenames already written during THIS run (see _reserve_filename).
            self._used_filenames = set()
            self._verify_misses = 0

            self._titleblock_size_cache = {}  # Sheet ID -> (width_mm, height_mm)
            self._paper_size_cache = {}  # Sheet ID -> (size_name, orientation)
            self._titleblock_cache_loaded = False
            self._batch_updating = False

            self.naming_pattern = type('_NP', (), {'Text': '{SheetNumber}'})()


            if HAS_API_LEARNER:
                try:
                    # No web check here: the old weekly revitapidocs.com probe was
                    # a sync download on the UI thread (100 s timeout).
                    self.api_adapter = SmartAPIAdapter(self.doc, REVIT_VERSION)
                except Exception as adapter_ex:
                    logger.warning("Could not initialize Smart API Adapter: {}".format(adapter_ex))
                    self.api_adapter = None
            else:
                self.api_adapter = None

            default_folder = os.path.join(os.path.expanduser('~'), 'Documents', 'Revit Exports')
            self.output_folder.Text = default_folder

            self.load_cad_export_setups()
            self.load_sheet_sets_for_filter()
            self.load_sheets()
            self.load_profiles()

            if not HAS_NAVISWORKS:
                self.export_nwd.IsEnabled = False
                self.export_nwd.ToolTip = "Navisworks export not available — NavisworksExportOptions missing in this Revit version"

            if not HAS_IFC:
                self.export_ifc.IsEnabled = False
                self.export_ifc.ToolTip = "IFC export not available — IFCExportOptions missing in this Revit version"

            self._restore_latest_setup()

            # After _restore_latest_setup so this status wins the display
            self._warn_if_previous_crash()

            # In code: EventSetters in Styles don't work with pyRevit
            self.sheets_listview.PreviewMouseLeftButtonDown += self.listview_clicked

            self.sheets_listview.SelectionChanged += self.sheets_selection_changed

            self.main_tabs.SelectionChanged += self.tab_changed

            # Auto-save latest setup on any close path
            self.Closing += self._window_closing_save_setup
            self.Closed += self._window_closed_dispose

            self.update_navigation_buttons()

        except Exception as ex:
            logger.error("Error initializing BatchOut window: {}".format(ex))
            raise


    def load_profiles(self, select_name=None):
        """Fill the profile picker from the profiles folder; stray JSON is skipped."""
        self.profiles = []
        try:
            if not os.path.exists(self.profiles_folder):
                os.makedirs(self.profiles_folder)
            for name, _path, data in batchout_profiles.list_profiles(self.profiles_folder):
                profile = ExportProfile.from_dict(data)
                profile.Name = name
                self.profiles.append(profile)
            logger.info("Loaded {} profiles".format(len(self.profiles)))
        except Exception as ex:
            logger.error("Error loading profiles: {}".format(ex))

        # Filling the combo raises SelectionChanged; it must not apply a profile.
        self._filling_profiles = True
        try:
            self.profile_combo.Items.Clear()
            selected = -1
            for index, profile in enumerate(self.profiles):
                self.profile_combo.Items.Add(profile.Name)
                if select_name is not None and profile.Name == select_name:
                    selected = index
            self.profile_combo.SelectedIndex = selected
        except Exception as ex:
            logger.debug("Profile picker not filled: {}".format(ex))
        finally:
            self._filling_profiles = False

    def profile_selected(self, sender, e):
        if self._filling_profiles:
            return
        index = self.profile_combo.SelectedIndex
        if 0 <= index < len(self.profiles):
            profile = self.profiles[index]
            self.apply_profile_to_ui(profile)
            self.status_text.Text = "Profile '{}' applied".format(profile.Name)

    def import_profile_clicked(self, sender, e):
        """Copy a profile JSON from any folder (shared drive, email) into the picker."""
        dialog = OpenFileDialog()
        dialog.Title = "Import BatchOut Profile"
        dialog.Filter = "BatchOut profile (*.json)|*.json"
        if dialog.ShowDialog() != DialogResult.OK:
            return
        source = dialog.FileName
        try:
            data = batchout_profiles.read_profile_file(source)
        except (OSError, ValueError) as ex:
            forms.alert("'{}' was not imported: {}.\n\nChoose a .json file saved with "
                        "BatchOut's Save As button.".format(os.path.basename(source), ex),
                        title="Profile Not Imported")
            return
        try:
            same_folder = os.path.normcase(os.path.dirname(os.path.abspath(source))) == \
                os.path.normcase(os.path.abspath(self.profiles_folder))
            if same_folder:
                name = str(data['Name']).strip()
            else:
                name = batchout_profiles.unique_profile_name(
                    [p.Name for p in self.profiles], data['Name'])
                data['Name'] = name
                batchout_profiles.write_profile_file(
                    batchout_profiles.unique_profile_path(self.profiles_folder, name), data)
        except (OSError, ValueError) as ex:
            forms.alert("The profile was read but could not be saved to\n{}\n\n{}\n\n"
                        "Check that the folder is writable, then import again.".format(
                            self.profiles_folder, ex), title="Profile Not Imported")
            return
        self.load_profiles(select_name=name)
        self.apply_profile_to_ui(ExportProfile.from_dict(data))
        self.status_text.Text = "Profile '{}' imported and applied".format(name)

    def save_profile_clicked(self, sender, e):
        """Save the current settings as a profile; any folder works (share it with the team)."""
        dialog = SaveFileDialog()
        dialog.Title = "Save BatchOut Profile"
        dialog.Filter = "BatchOut profile (*.json)|*.json"
        dialog.InitialDirectory = self.profiles_folder
        index = self.profile_combo.SelectedIndex
        dialog.FileName = batchout_profiles.safe_filename(
            self.profiles[index].Name if 0 <= index < len(self.profiles) else "BatchOut Profile")
        if dialog.ShowDialog() != DialogResult.OK:
            return
        path = dialog.FileName
        folder = os.path.dirname(path)
        # A leading '_' would hide the file from the picker (bookkeeping prefix).
        stem = batchout_profiles.safe_filename(os.path.splitext(os.path.basename(path))[0])
        path = os.path.join(folder, stem + batchout_profiles.PROFILE_EXT)
        profile = self.get_current_settings_as_profile()
        profile.Name = stem
        profile.Description = "Saved on {}".format(datetime.now().strftime("%Y-%m-%d %H:%M"))
        try:
            batchout_profiles.write_profile_file(path, profile.to_dict())
        except OSError as ex:
            forms.alert("The profile was not saved to\n{}\n\n{}\n\nChoose a folder you can "
                        "write to and try again.".format(path, ex), title="Profile Not Saved")
            return
        if os.path.normcase(os.path.abspath(folder)) == os.path.normcase(
                os.path.abspath(self.profiles_folder)):
            self.load_profiles(select_name=stem)
            self.status_text.Text = "Profile '{}' saved".format(stem)
        else:
            self.status_text.Text = "Profile '{}' saved to {} — teammates can import it".format(
                stem, folder)


    def save_latest_setup(self):
        """Auto-save current settings as the latest setup (restored on next open)."""
        try:
            if not os.path.exists(self.profiles_folder):
                os.makedirs(self.profiles_folder)
            profile = self.get_current_settings_as_profile()
            profile.Name = "Latest Setup"
            profile.Description = "Auto-saved on {}".format(datetime.now().strftime("%Y-%m-%d %H:%M"))
            with open(self.latest_setup_file, 'w') as f:
                json.dump(profile.to_dict(), f, indent=2)
            logger.debug("Latest setup auto-saved to {}".format(self.latest_setup_file))
        except Exception as ex:
            logger.warning("Could not auto-save latest setup: {}".format(ex))

    def _restore_latest_setup(self):
        """Restore the auto-saved latest setup from the previous session, if any."""
        try:
            if not os.path.exists(self.latest_setup_file):
                return
            with open(self.latest_setup_file, 'r') as f:
                data = json.load(f)
            self.apply_profile_to_ui(ExportProfile.from_dict(data))
            self.status_text.Text = "Latest setup restored"
            logger.info("Latest setup restored")
        except Exception as ex:
            logger.warning("Could not restore latest setup: {}".format(ex))

    def _window_closing_save_setup(self, sender, e):
        """Closing hook — persist settings so the next session starts where the user left off."""
        self.save_latest_setup()

    def _window_closed_dispose(self, sender, e):
        """Dispose the ExternalEvent once the window is gone.

        With __persistentengine__ the window outlives its close; an undisposed
        event crashes Revit at shutdown (0xc0000005 'Invalid WPFWndProxy',
        journal.0090, 2026-07-15).
        """
        # Safety net: Revit must never stay locked behind a closed window.
        lock = getattr(self, '_input_lock', None)
        if lock is not None:
            lock.release()
        if getattr(self, '_api_handler', None) is not None:
            self._api_handler.clear()
        try:
            if self._api_event is not None:
                self._api_event.Dispose()
        except Exception as ex:
            logger.debug("ExternalEvent dispose failed: {}".format(ex))
        self._api_event = None

    def _run_in_api_context(self, action):
        """Run a callable inside valid Revit API context.

        Modeless: queue through the ExternalEvent (fires when Revit idles).
        Modal: Dispatcher callbacks pump INSIDE the command's API context
        (blocked in ShowDialog). The ExternalEvent would only fire after the
        modal window closes — Revit never idles meanwhile.
        """
        if self.modeless:
            if self._api_event is None:
                self.status_text.Text = "Revit connection is closed. Reopen BatchOut."
                return False
            self._api_handler.add(action)
            try:
                outcome = str(self._api_event.Raise())
                if outcome not in ('Accepted', 'Pending'):
                    raise RuntimeError("Revit rejected the request: " + outcome)
            except Exception as ex:
                self._api_handler.remove(action)
                self.status_text.Text = "Could not queue the action. Try again when Revit is ready."
                logger.error("BatchOut could not queue action: {}".format(ex))
                return False
        else:
            try:
                self.Dispatcher.BeginInvoke(DispatcherPriority.Background,
                                            Action(action))
            except Exception as ex:
                self.status_text.Text = "Could not queue the action. Reopen BatchOut and retry."
                logger.error("BatchOut dispatcher failed: {}".format(ex))
                return False
        return True

    def get_current_settings_as_profile(self):
        """Capture current UI settings as a profile."""
        profile = ExportProfile()

        profile.ExportPDF = self.export_pdf.IsChecked if self.export_pdf.IsChecked is not None else False
        profile.ExportDWG = self.export_dwg.IsChecked if self.export_dwg.IsChecked is not None else False
        profile.ExportDGN = self.export_dgn.IsChecked if self.export_dgn.IsChecked is not None else False
        profile.ExportDWF = self.export_dwf.IsChecked if self.export_dwf.IsChecked is not None else False
        profile.ExportNWD = self.export_nwd.IsChecked if self.export_nwd.IsChecked is not None else False
        profile.ExportIFC = self.export_ifc.IsChecked if self.export_ifc.IsChecked is not None else False
        profile.ExportIMG = self.export_img.IsChecked if self.export_img.IsChecked is not None else False

        if self.pdf_paper_size.SelectedItem:
            profile.PDFPaperSize = self.pdf_paper_size.SelectedItem.Content
        profile.PDFOrientation = "Landscape" if self.pdf_landscape.IsChecked else "Portrait"
        profile.PDFZoom = "Fit to Page" if self.pdf_fit_to_page.IsChecked else "Custom Zoom"
        profile.PDFHideRefPlanes = self.pdf_hide_ref_planes.IsChecked if self.pdf_hide_ref_planes.IsChecked is not None else False
        profile.PDFHideScopeBoxes = self.pdf_hide_scope_boxes.IsChecked if self.pdf_hide_scope_boxes.IsChecked is not None else False
        profile.PDFHideCropBoundaries = self.pdf_hide_crop_boundaries.IsChecked if self.pdf_hide_crop_boundaries.IsChecked is not None else False
        profile.PDFHideUnreferencedTags = self.pdf_hide_unreferenced_tags.IsChecked if self.pdf_hide_unreferenced_tags.IsChecked is not None else False
        profile.PDFViewLinksBlue = self.pdf_view_links_blue.IsChecked if self.pdf_view_links_blue.IsChecked is not None else False
        profile.CombinePDF = self.combine_pdf.IsChecked if self.combine_pdf.IsChecked is not None else False
        profile.PDFEngine = self._pdf_engine()

        if self.dwg_version.SelectedItem:
            profile.DWGVersion = self.dwg_version.SelectedItem.Content
        if self.cad_export_setup.SelectedItem:
            profile.CADExportSetup = self.cad_export_setup.SelectedItem.Content
        profile.CADExportViewsOnSheets = self.cad_export_views_on_sheets.IsChecked if self.cad_export_views_on_sheets.IsChecked is not None else False
        profile.CADExportLinksAsExternal = self.cad_export_links_as_external.IsChecked if self.cad_export_links_as_external.IsChecked is not None else False
        profile.CADMatchRevitDisplay = bool(self.cad_match_revit_display.IsChecked)
        profile.CADSkipImageReferences = bool(self.cad_skip_image_references.IsChecked)

        profile.OutputFolder = self.output_folder.Text if self.output_folder.Text else ""
        profile.SplitByFormat = self.save_split_by_format.IsChecked if self.save_split_by_format.IsChecked is not None else False
        profile.ReverseOrder = self.reverse_order.IsChecked if self.reverse_order.IsChecked is not None else False

        profile.NamingPattern = self.naming_pattern.Text if self.naming_pattern.Text else "{SheetNumber}-{SheetName}"

        profile.WriteReport = bool(self.write_report.IsChecked)
        profile.ReportFormat = self._selected_report_format()
        profile.ReportColumns = list(self._report_columns)

        return profile

    def _selected_report_format(self):
        item = self.report_format.SelectedItem
        return batchout_report.normalize_format(item.Content if item is not None else None)

    def apply_profile_to_ui(self, profile):
        """Apply profile settings to UI controls."""
        try:
            self.export_pdf.IsChecked = profile.ExportPDF
            self.export_dwg.IsChecked = profile.ExportDWG
            self.export_dgn.IsChecked = profile.ExportDGN
            self.export_dwf.IsChecked = profile.ExportDWF
            self.export_nwd.IsChecked = profile.ExportNWD
            self.export_ifc.IsChecked = profile.ExportIFC
            self.export_img.IsChecked = profile.ExportIMG

            # Formats unavailable in this Revit version must never come back checked
            if not HAS_NAVISWORKS:
                self.export_nwd.IsChecked = False
            if not HAS_IFC:
                self.export_ifc.IsChecked = False

            for i in range(self.pdf_paper_size.Items.Count):
                if self.pdf_paper_size.Items[i].Content == profile.PDFPaperSize:
                    self.pdf_paper_size.SelectedIndex = i
                    break

            if profile.PDFOrientation == "Landscape":
                self.pdf_landscape.IsChecked = True
            else:
                self.pdf_portrait.IsChecked = True

            if profile.PDFZoom == "Fit to Page":
                self.pdf_fit_to_page.IsChecked = True
            else:
                self.pdf_zoom_custom.IsChecked = True

            self.pdf_hide_ref_planes.IsChecked = profile.PDFHideRefPlanes
            self.pdf_hide_scope_boxes.IsChecked = profile.PDFHideScopeBoxes
            self.pdf_hide_crop_boundaries.IsChecked = profile.PDFHideCropBoundaries
            self.pdf_hide_unreferenced_tags.IsChecked = profile.PDFHideUnreferencedTags
            self.pdf_view_links_blue.IsChecked = profile.PDFViewLinksBlue
            self.combine_pdf.IsChecked = profile.CombinePDF
            self.pdf_engine.SelectedIndex = 1 if profile.PDFEngine == "PDF24" else 0

            for i in range(self.dwg_version.Items.Count):
                if self.dwg_version.Items[i].Content == profile.DWGVersion:
                    self.dwg_version.SelectedIndex = i
                    break

            for i in range(self.cad_export_setup.Items.Count):
                if self.cad_export_setup.Items[i].Content == profile.CADExportSetup:
                    self.cad_export_setup.SelectedIndex = i
                    break

            self.cad_export_views_on_sheets.IsChecked = profile.CADExportViewsOnSheets
            self.cad_export_links_as_external.IsChecked = profile.CADExportLinksAsExternal
            self.cad_match_revit_display.IsChecked = bool(profile.CADMatchRevitDisplay)
            self.cad_skip_image_references.IsChecked = bool(profile.CADSkipImageReferences)

            self.output_folder.Text = profile.OutputFolder
            if profile.SplitByFormat:
                self.save_split_by_format.IsChecked = True
            else:
                self.save_same_folder.IsChecked = True
            self.reverse_order.IsChecked = profile.ReverseOrder

            self.naming_pattern.Text = profile.NamingPattern

            self.write_report.IsChecked = bool(profile.WriteReport)
            report_format = batchout_report.normalize_format(profile.ReportFormat)
            for i in range(self.report_format.Items.Count):
                if str(self.report_format.Items[i].Content) == report_format:
                    self.report_format.SelectedIndex = i
                    break
            self._report_columns = batchout_report.normalize_columns(profile.ReportColumns)

            self.status_text.Text = "Profile '{}' loaded successfully".format(profile.Name)

        except Exception as ex:
            logger.error("Error applying profile to UI: {}".format(ex))
            forms.alert("Error applying profile:\n{}".format(str(ex)))


    # size_name -> (width_mm, height_mm), landscape
    PAPER_SIZES_MM = {
        "A0": (1189, 841),
        "A1": (841, 594),
        "A2": (594, 420),
        "A3": (420, 297),
        "A4": (297, 210),
        "A5": (210, 148),
        "B0": (1414, 1000),
        "B1": (1000, 707),
        "B2": (707, 500),
        "B3": (500, 353),
        "B4": (353, 250),
        "B5": (250, 176),
        "ANSI A": (279, 216),
        "ANSI B": (432, 279),
        "ANSI C": (559, 432),
        "ANSI D": (864, 559),
        "ANSI E": (1118, 864),
        "ARCH A": (305, 229),
        "ARCH B": (457, 305),
        "ARCH C": (610, 457),
        "ARCH D": (914, 610),
        "ARCH E": (1219, 914),
        "ARCH E1": (1067, 762),
        "Letter": (279, 216),
        "Legal": (356, 216),
        "Tabloid": (432, 279),
        "Ledger": (432, 279),
    }

    def _batch_load_titleblock_sizes(self):
        """Load all titleblock dimensions in ONE query (not one collector per sheet)."""
        try:
            self._titleblock_size_cache = {}

            all_titleblocks = FilteredElementCollector(self.doc)\
                .OfCategory(BuiltInCategory.OST_TitleBlocks)\
                .WhereElementIsNotElementType()\
                .ToElements()

            # OwnerViewId (sheet id) -> (width_mm, height_mm)
            for tb in all_titleblocks:
                try:
                    owner_view_id = tb.OwnerViewId
                    if owner_view_id and owner_view_id != DB.ElementId.InvalidElementId:
                        width_param = tb.get_Parameter(DB.BuiltInParameter.SHEET_WIDTH)
                        height_param = tb.get_Parameter(DB.BuiltInParameter.SHEET_HEIGHT)

                        if width_param and height_param:
                            width_mm = width_param.AsDouble() * 304.8
                            height_mm = height_param.AsDouble() * 304.8
                            self._titleblock_size_cache[eid_value(owner_view_id)] = (width_mm, height_mm)
                except:
                    continue

            self._titleblock_cache_loaded = True
            logger.debug("Batch loaded {} titleblock sizes".format(len(self._titleblock_size_cache)))

        except Exception as ex:
            logger.debug("Error batch loading titleblocks: {}".format(ex))
            self._titleblock_size_cache = {}
            self._titleblock_cache_loaded = True

    def _detect_paper_size_from_dimensions(self, width_mm, height_mm):
        """(size_name, orientation) from dimensions, ±10 mm; common sizes first."""
        # Normalize to landscape
        if width_mm > height_mm:
            orientation = "Landscape"
        else:
            orientation = "Portrait"
            width_mm, height_mm = height_mm, width_mm

        tolerance = 10  # mm

        priority_sizes = ["A3", "A1", "A2", "A0", "A4", "ARCH D", "ARCH E", "ANSI D", "ANSI E"]

        for size_name in priority_sizes:
            if size_name in self.PAPER_SIZES_MM:
                std_width, std_height = self.PAPER_SIZES_MM[size_name]
                if (abs(width_mm - std_width) < tolerance and
                    abs(height_mm - std_height) < tolerance):
                    return (size_name, orientation)

        for size_name, (std_width, std_height) in self.PAPER_SIZES_MM.items():
            if size_name in priority_sizes:
                continue
            if (abs(width_mm - std_width) < tolerance and
                abs(height_mm - std_height) < tolerance):
                return (size_name, orientation)

        return ("Use Sheet Size", orientation)

    def get_paper_format_enum(self, size_str):
        """Map paper size string (e.g. A0, A1, A2, A3, A4, ISO_A3) to Autodesk.Revit.DB.ExportPaperFormat enum."""
        if not size_str:
            return DB.ExportPaperFormat.Default
            
        s = size_str.upper().strip()
        if s == "A0" or "ISO_A0" in s:
            return DB.ExportPaperFormat.ISO_A0
        elif s == "A1" or "ISO_A1" in s:
            return DB.ExportPaperFormat.ISO_A1
        elif s == "A2" or "ISO_A2" in s:
            return DB.ExportPaperFormat.ISO_A2
        elif s == "A3" or "ISO_A3" in s:
            return DB.ExportPaperFormat.ISO_A3
        elif s == "A4" or "ISO_A4" in s:
            return DB.ExportPaperFormat.ISO_A4
        elif "ANSI_A" in s or "ANSI A" in s:
            return DB.ExportPaperFormat.ANSI_A
        elif "ANSI_B" in s or "ANSI B" in s:
            return DB.ExportPaperFormat.ANSI_B
        elif "ANSI_C" in s or "ANSI C" in s:
            return DB.ExportPaperFormat.ANSI_C
        elif "ANSI_D" in s or "ANSI D" in s:
            return DB.ExportPaperFormat.ANSI_D
        elif "ANSI_E" in s or "ANSI E" in s:
            return DB.ExportPaperFormat.ANSI_E
        elif "ARCH_A" in s or "ARCH A" in s:
            return DB.ExportPaperFormat.ARCH_A
        elif "ARCH_B" in s or "ARCH B" in s:
            return DB.ExportPaperFormat.ARCH_B
        elif "ARCH_C" in s or "ARCH C" in s:
            return DB.ExportPaperFormat.ARCH_C
        elif "ARCH_D" in s or "ARCH D" in s:
            return DB.ExportPaperFormat.ARCH_D
        elif "ARCH_E" in s or "ARCH E" in s:
            return DB.ExportPaperFormat.ARCH_E
        elif "ARCH_E1" in s or "ARCH E1" in s:
            return DB.ExportPaperFormat.ARCH_E1
        
        return DB.ExportPaperFormat.Default

    def get_sheet_paper_size_and_orientation(self, sheet):
        """(paper_size, orientation) from the cached Title Block dimensions."""
        try:
            self._ensure_titleblock_cache()
            sheet_id = eid_value(sheet.Id)

            if sheet_id in self._paper_size_cache:
                return self._paper_size_cache[sheet_id]

            if sheet_id in self._titleblock_size_cache:
                width_mm, height_mm = self._titleblock_size_cache[sheet_id]
                result = self._detect_paper_size_from_dimensions(width_mm, height_mm)
                self._paper_size_cache[sheet_id] = result
                return result

            # Not cached = no titleblock (e.g. START VIEW). NEVER use a view-scoped
            # FilteredElementCollector(doc, sheet.Id): it initialises view graphics
            # via RevitWorker.exe and deadlocks/crashes on complex or cloud start views.
            result = ("Use Sheet Size", "Landscape")
            self._paper_size_cache[sheet_id] = result
            return result

        except Exception as ex:
            logger.debug("Could not auto-detect paper size for sheet: {}".format(ex))
            return ("Use Sheet Size", "Landscape")

    def load_cad_export_setups(self):
        """Load available DWG export setups from the document."""
        try:
            self.cad_export_setup.Items.Clear()

            from System.Windows.Controls import ComboBoxItem
            default_item = ComboBoxItem()
            default_item.Content = "Use setup from file (Default)"
            self.cad_export_setup.Items.Add(default_item)

            collector = FilteredElementCollector(self.doc)\
                .OfClass(ExportDWGSettings)

            for setup in collector:
                try:
                    setup_name = setup.Name if hasattr(setup, 'Name') else "Setup {}".format(eid_value(setup.Id))
                    item = ComboBoxItem()
                    item.Content = setup_name
                    item.Tag = setup
                    self.cad_export_setup.Items.Add(item)
                except:
                    pass

            self.cad_export_setup.SelectedIndex = 0

        except Exception as ex:
            logger.warning("Could not load CAD export setups: {}".format(ex))
            from System.Windows.Controls import ComboBoxItem
            default_item = ComboBoxItem()
            default_item.Content = "Use setup from file (Default)"
            self.cad_export_setup.Items.Add(default_item)
            self.cad_export_setup.SelectedIndex = 0

    # Raster quality, highest -> lowest. Presentation is the default; we step DOWN
    # only for a sheet proven to crash the PDF rasteriser higher up (fatal
    # "Invalid array ... Arr.cpp" on link-heavy views). 300-DPI "High" is print-grade.
    RASTER_LADDER = ("Presentation", "High", "Medium", "Low")

    def _apply_pdf_quality(self, pdf_options, level=None):
        """Force high-fidelity color + raster quality on PDF export.

        Revit rasterizes patches under transparent / overlapping filled regions;
        at low raster quality their fine linework is dropped ('missing lines').
        Guarded: enum / property names vary between Revit versions.

        `level` = starting rung of RASTER_LADDER (None = Presentation); falls
        further down if the enum value is missing in this build.
        """
        # Raster quality: lever for dropped lines AND the Arr.cpp memory spike
        try:
            rq_enum = getattr(DB, 'RasterQualityType', None)
            if rq_enum is not None and hasattr(pdf_options, 'RasterQuality'):
                start = self.RASTER_LADDER.index(level) if level in self.RASTER_LADDER else 0
                quality = None
                for name in self.RASTER_LADDER[start:]:
                    quality = getattr(rq_enum, name, None)
                    if quality is not None:
                        break
                if quality is not None:
                    pdf_options.RasterQuality = quality
        except Exception as ex:
            logger.debug("Could not set PDF RasterQuality in Revit {}: {}".format(REVIT_VERSION, ex))

        # Full color so nothing collapses to line/grey
        try:
            color_enum = getattr(DB, 'ColorDepthType', None)
            if color_enum is not None and hasattr(pdf_options, 'ColorDepth'):
                color_val = getattr(color_enum, 'Color', None)
                if color_val is not None:
                    pdf_options.ColorDepth = color_val
        except Exception as ex:
            logger.debug("Could not set PDF ColorDepth in Revit {}: {}".format(REVIT_VERSION, ex))

        return pdf_options

    # ── Native-crash breadcrumb ──
    # Revit's renderer can hard-crash while regenerating a sheet ("Invalid array
    # ... Arr.cpp" or 0xc0000005 on a render thread, after "Problems during
    # faceting"), in PDF and DWG (VHLR) alike — uncatchable from Python. A durable
    # marker names the sheet+format handed to Revit; if it survives, the next
    # session promotes it to the crash history, which queues that combination
    # LAST and offers safe mode. Skipping is only the user's explicit choice.

    def _read_crash_marker(self):
        """Return the in-flight crash-marker dict from a prior aborted export."""
        try:
            if os.path.exists(self._crash_marker_file):
                with open(self._crash_marker_file, 'r') as f:
                    return json.load(f)
        except Exception as ex:
            logger.debug("Could not read crash marker: {}".format(ex))
        return None

    def _read_crash_history(self):
        """Sheet+format combinations that hard-crashed Revit before. Never raises."""
        try:
            if os.path.exists(self._crash_history_file):
                with open(self._crash_history_file, 'r') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        return data
        except Exception as ex:
            logger.debug("Could not read crash history: {}".format(ex))
        return []

    def _promote_crash_marker(self):
        """A leftover marker = that sheet+format killed Revit. Record it in the
        history, clear the marker, and return it for the UI."""
        pc = self._read_crash_marker()
        if not pc:
            return None
        try:
            # safe_level is in the key: a crash at a STRONGER rung is new info;
            # dropping it would retry a proven-failed rung forever.
            key = (pc.get('doc'), pc.get('sheet'), pc.get('format'), pc.get('safe_level'))
            known = set((e.get('doc'), e.get('sheet'), e.get('format'), e.get('safe_level'))
                        for e in self._crash_history)
            if key not in known:
                self._crash_history.append(pc)
                self._write_crash_history()
        except Exception as ex:
            logger.debug("Could not promote crash marker: {}".format(ex))
        self._clear_crash_marker()
        return pc

    def _write_crash_history(self):
        """Persist the confirmed-fatal list (last 50 entries)."""
        try:
            folder = os.path.dirname(self._crash_history_file)
            if not os.path.exists(folder):
                os.makedirs(folder)
            with open(self._crash_history_file, 'w') as f:
                json.dump(self._crash_history[-50:], f, indent=2)
        except Exception as ex:
            logger.debug("Could not write crash history: {}".format(ex))

    def _write_crash_marker(self, sheet_number, fmt, level=None, safe_level=None):
        """Persist the sheet+format handed to the native exporter RIGHT NOW and
        its safe-mode rung (lets the next session escalate). fsync'd so it
        survives an immediate kill."""
        try:
            folder = os.path.dirname(self._crash_marker_file)
            if not os.path.exists(folder):
                os.makedirs(folder)
            data = {
                'sheet': sheet_number,
                'format': fmt,
                'raster': level,
                'safe_level': safe_level,
                # Rung NAME too: the index only means something against the
                # writing build's SAFE_LADDER.
                'safe_rung': (self.SAFE_LADDER[safe_level]
                              if safe_level is not None
                              and 0 <= safe_level < len(self.SAFE_LADDER) else None),
                'doc': self.doc.Title,
                'time': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            with open(self._crash_marker_file, 'w') as f:
                json.dump(data, f)
                f.flush()
                try:
                    os.fsync(f.fileno())
                except Exception:
                    pass
        except Exception as ex:
            logger.debug("Could not write crash marker: {}".format(ex))

    def _clear_crash_marker(self):
        """Remove the marker once the native export call returned."""
        try:
            if os.path.exists(self._crash_marker_file):
                os.remove(self._crash_marker_file)
        except Exception as ex:
            logger.debug("Could not clear crash marker: {}".format(ex))

    def _is_known_fatal(self, sheet_number, fmt):
        """True if this document+sheet+format has already hard-crashed Revit.

        Still exported (queued last, safe mode offered, user warned); skipped only
        on the user's explicit choice.
        """
        try:
            title = self.doc.Title
            for entry in self._crash_history:
                if (entry.get('sheet') == sheet_number
                        and entry.get('format') == fmt
                        and (not entry.get('doc') or entry.get('doc') == title)):
                    return True
        except Exception:
            pass
        return False

    # ── Safe-mode export ──
    # The fatal geometry is usually one or two elements: neutralise them, export,
    # roll back (model unchanged). Still produces the file, unlike skipping.

    def _bad_geometry_ids(self):
        """ElementIds the Bad Geometry check flagged for THIS model ([] if not run)."""
        try:
            findings = bad_geometry_file()
            if not os.path.exists(findings):
                return []
            with open(findings, 'r') as f:
                data = json.load(f)
            if data.get('doc') and data.get('doc') != self.doc.Title:
                return []
            return [make_eid(e['id']) for e in data.get('elements', []) if 'id' in e]
        except Exception as ex:
            logger.debug("Could not read bad-geometry list: {}".format(ex))
            return []

    # Categories carrying the spline solids behind "SilSolver::eval ... Hermite
    # surface" crashes (SPH bath sheets). Fallback for the 'categories' rung
    # when the Bad Geometry check named nothing.
    SUSPECT_CATEGORIES = (
        DB.BuiltInCategory.OST_PlumbingFixtures,
        DB.BuiltInCategory.OST_SpecialityEquipment,
        DB.BuiltInCategory.OST_GenericModel,
        DB.BuiltInCategory.OST_Casework,
        DB.BuiltInCategory.OST_Furniture,
    )

    # Cumulative mitigations, ordered by how much of the drawing they destroy —
    # least first. We climb only past a rung PROVEN to crash (each failure costs
    # a Revit session + cloud reopen), so few attempts matter.
    # 'wireframe' is second: it hides NOTHING and skips hidden-line removal /
    # face tessellation, where every fatal journal signature comes from
    # (facetFaceWithSpecifiedAlgorithm, silhouette edges, SilSolver::eval).
    # Hides first would have emptied interior elevations (30-Jul crash, SPH
    # ELEVATIONS 3, 4).
    # NOTE: history stores an INDEX (`safe_level`) — reordering invalidates old
    # levels; markers also record the NAME (`safe_rung`).
    SAFE_LADDER = ('elements', 'wireframe', 'coarse', 'categories', 'links')

    def _next_safe_level(self, sheet_number, fmt):
        """Rung to start on: one above the highest rung that already crashed (0 if none)."""
        highest = -1
        try:
            title = self.doc.Title
            for entry in self._crash_history:
                if (entry.get('sheet') == sheet_number
                        and entry.get('format') == fmt
                        and (not entry.get('doc') or entry.get('doc') == title)):
                    recorded = entry.get('safe_level')
                    if recorded is not None and recorded > highest:
                        highest = recorded
        except Exception:
            pass
        return highest + 1

    def _safe_ladder_exhausted(self, sheet_number, fmt):
        """True once every rung has crashed — nothing left to try automatically."""
        return self._next_safe_level(sheet_number, fmt) >= len(self.SAFE_LADDER)

    def _begin_safe_mode(self, sheet, level, fmt=None):
        """Open a TransactionGroup, apply every SAFE_LADDER rung up to `level`,
        return (group, level_used) for the caller to roll back; (None, level)
        if nothing changed. (24-Jul journal.1084: Coarse alone still crashed
        sheet 1402.)"""
        views = []
        try:
            for vid in sheet.GetAllPlacedViews():
                view = self.doc.GetElement(vid)
                if view is not None and not view.IsTemplate:
                    views.append(view)
        except Exception as ex:
            logger.error("Could not read placed views: {}".format(ex))
        if not views:
            return None, level

        group = None
        try:
            group = DB.TransactionGroup(self.doc, "BatchOut safe-mode export")
            group.Start()
            trans = Transaction(self.doc, "Simplify views for safe export")
            trans.Start()

            applied = []
            # A rung that changes nothing (e.g. 'elements' with no report) must
            # not waste a session: keep climbing until something applies.
            top = min(level, len(self.SAFE_LADDER) - 1)
            rung = 0
            while rung <= top or not applied:
                if rung >= len(self.SAFE_LADDER):
                    break
                changed = self._apply_safe_rung(self.SAFE_LADDER[rung], views)
                if changed:
                    applied.append("{}={}".format(self.SAFE_LADDER[rung], changed))
                rung += 1

            trans.Commit()
            level_used = max(0, rung - 1)

            if not applied:
                group.RollBack()
                logger.warning("Safe mode on '{}': nothing could be changed.".format(
                    sheet.SheetNumber))
                return None, level_used

            logger.warning("Safe mode L{} on '{}': {} (rolled back after export).".format(
                level_used, sheet.SheetNumber, ", ".join(applied)))
            # Also in the completion dialog: the file may miss content or
            # hidden-line cleanup.
            self._note_safe_mode(sheet.SheetNumber, fmt,
                                 self.SAFE_LADDER[level_used], applied)
            return group, level_used
        except Exception as ex:
            logger.error("Could not enter safe mode: {}".format(ex))
            if group is not None:
                try:
                    group.RollBack()
                except Exception:
                    pass
            return None, level

    def _apply_safe_rung(self, rung, views):
        """Apply one rung across `views`; returns the change count. Needs an open Transaction."""
        changed = 0

        if rung == 'elements':
            # Surgical: only what the Bad Geometry check named.
            bad_ids = self._bad_geometry_ids()
            if not bad_ids:
                return 0
            for view in views:
                hideable = List[DB.ElementId]()
                for eid in bad_ids:
                    try:
                        element = self.doc.GetElement(eid)
                        if element is None or element.IsHidden(view):
                            continue
                        if not element.CanBeHidden(view):
                            continue
                        hideable.Add(eid)
                    except Exception:
                        continue
                if hideable.Count:
                    try:
                        view.HideElements(hideable)
                        changed += hideable.Count
                    except Exception as ex:
                        logger.debug("HideElements failed on {}: {}".format(view.Id, ex))

        elif rung == 'categories':
            # Reaches geometry inside a RevitLink (linked ids can't go to
            # HideElements; category visibility propagates) and works without a
            # Bad Geometry run. Check's categories first, else SUSPECT_CATEGORIES.
            cat_ids = self._bad_geometry_category_ids()
            if not cat_ids:
                cat_ids = []
                for bic in self.SUSPECT_CATEGORIES:
                    try:
                        category = DB.Category.GetCategory(self.doc, bic)
                        if category is not None:
                            cat_ids.append(category.Id)
                    except Exception:
                        continue
            for view in views:
                for cat_id in cat_ids:
                    try:
                        if not view.CanCategoryBeHidden(cat_id):
                            continue
                        if view.GetCategoryHidden(cat_id):
                            continue
                        view.SetCategoryHidden(cat_id, True)
                        changed += 1
                    except Exception as ex:
                        logger.debug("SetCategoryHidden failed on {}: {}".format(
                            view.Id, ex))

        elif rung == 'links':
            # Every crash journal shows RVT-link graphics regenerating right before death
            for view in views:
                hideable = List[DB.ElementId]()
                try:
                    links = FilteredElementCollector(self.doc, view.Id) \
                        .OfClass(DB.RevitLinkInstance) \
                        .ToElements()
                except Exception:
                    continue
                for link in links:
                    try:
                        if not link.IsHidden(view) and link.CanBeHidden(view):
                            hideable.Add(link.Id)
                    except Exception:
                        continue
                if hideable.Count:
                    try:
                        view.HideElements(hideable)
                        changed += hideable.Count
                    except Exception as ex:
                        logger.debug("Hiding links failed on {}: {}".format(view.Id, ex))

        elif rung == 'coarse':
            # Blunt: simpler family solids (alone insufficient for sheet 1402)
            for view in views:
                try:
                    if view.DetailLevel != DB.ViewDetailLevel.Coarse:
                        view.DetailLevel = DB.ViewDetailLevel.Coarse
                        changed += 1
                except Exception as ex:
                    logger.debug("DetailLevel not settable on {}: {}".format(view.Id, ex))

        elif rung == 'wireframe':
            # Most direct: skips hidden-line removal; hides nothing (see SAFE_LADDER)
            for view in views:
                try:
                    if view.DisplayStyle != DB.DisplayStyle.Wireframe:
                        view.DisplayStyle = DB.DisplayStyle.Wireframe
                        changed += 1
                except Exception as ex:
                    logger.debug("DisplayStyle not settable on {}: {}".format(view.Id, ex))

        return changed

    def _bad_geometry_category_ids(self):
        """Category ids of the elements the Bad Geometry check flagged for THIS model."""
        ids = []
        try:
            if not os.path.exists(self.BAD_GEOMETRY_FILE):
                return ids
            with open(self.BAD_GEOMETRY_FILE, 'r') as f:
                data = json.load(f)
            if data.get('doc') and data.get('doc') != self.doc.Title:
                return ids
            seen = set()
            for entry in data.get('elements', []):
                eid = entry.get('id')
                if eid is None:
                    continue
                try:
                    element = self.doc.GetElement(make_eid(eid))
                    if element is None or element.Category is None:
                        continue
                    cat_id = element.Category.Id
                    key = str(cat_id)
                    if key not in seen:
                        seen.add(key)
                        ids.append(cat_id)
                except Exception:
                    continue
        except Exception as ex:
            logger.debug("Could not read bad-geometry categories: {}".format(ex))
        return ids

    def _end_safe_mode(self, group):
        """Always roll back — safe mode must never leave the model modified."""
        if group is None:
            return
        try:
            group.RollBack()
        except Exception as ex:
            logger.error("Safe-mode rollback FAILED: {} — undo manually".format(ex))

    def _begin_hide_images(self, element, group):
        """Hide the Raster Images category on `element` (sheet or view) and its
        placed views so Revit writes no image file beside the DWG. Joins the
        open safe-mode `group` or opens one; returns the group to roll back
        with `_end_safe_mode` (None if nothing was hidden)."""
        try:
            category = DB.Category.GetCategory(self.doc, DB.BuiltInCategory.OST_RasterImages)
        except Exception:
            category = None
        if category is None:
            return group
        views = [element]
        try:
            for vid in element.GetAllPlacedViews():
                view = self.doc.GetElement(vid)
                if view is not None and not view.IsTemplate:
                    views.append(view)
        except Exception:
            pass  # a view, not a sheet

        own_group = None
        trans = None
        try:
            if group is None:
                own_group = DB.TransactionGroup(self.doc, "BatchOut DWG without images")
                own_group.Start()
            trans = Transaction(self.doc, "Hide images for DWG export")
            trans.Start()
            changed = 0
            for view in views:
                try:
                    # Raises when a view template owns V/G: the post-export
                    # sweep in export_to_dwg still removes the files.
                    if (view.CanCategoryBeHidden(category.Id)
                            and not view.GetCategoryHidden(category.Id)):
                        view.SetCategoryHidden(category.Id, True)
                        changed += 1
                except Exception as ex:
                    logger.debug("Raster Images not hideable on {}: {}".format(view.Id, ex))
            trans.Commit()
            if changed or own_group is None:
                return group or own_group
            own_group.RollBack()
            return None
        except Exception as ex:
            logger.warning("Could not hide images for DWG export: {}".format(ex))
            try:
                if trans is not None and trans.HasStarted() and not trans.HasEnded():
                    trans.RollBack()
            except Exception:
                pass
            if own_group is not None:
                try:
                    own_group.RollBack()
                except Exception:
                    pass
            return group

    def _format_checkboxes(self):
        """(format, checkbox) pairs for the formats crash history tracks."""
        return (("DWG", self.export_dwg), ("PDF", self.export_pdf),
                ("DGN", self.export_dgn), ("DWF", self.export_dwf),
                ("NWC", self.export_nwd), ("IMG", self.export_img))

    def _is_risky_item(self, item):
        """True if this item hard-crashed Revit before in any checked format."""
        label = getattr(item, 'SheetNumber', None)
        if not label:
            return False
        for fmt, box in self._format_checkboxes():
            try:
                if box.IsChecked and self._is_known_fatal(label, fmt):
                    return True
            except Exception:
                continue
        return False

    def _order_risky_last(self, items):
        """Move items with a crash history to the END of the queue (still
        exported), so a crash leaves everything else already on disk."""
        safe, risky = [], []
        for item in items:
            (risky if self._is_risky_item(item) else safe).append(item)
        if risky:
            logger.warning("{} item(s) with a crash history queued last: {}".format(
                len(risky), ", ".join(str(getattr(i, 'SheetNumber', '?')) for i in risky)))
        return safe + risky

    def _ask_safe_mode(self, selected_items):
        """If any selected item+format is known-fatal, ask ONCE how to handle it.
        Only an explicit "Skip them" leaves files out (listed at completion)."""
        self._safe_mode = False
        self._skip_fatal = False
        RUNG_LABEL = {
            'elements':   "hide the exact elements the Bad Geometry check flagged",
            'wireframe':  "switch those views to Wireframe — hides NOTHING, but "
                          "skips the hidden-line pass that crashes",
            'coarse':     "drop those views to Coarse detail level",
            'categories': "hide the categories carrying the bad geometry "
                          "(also reaches geometry inside links) — the drawing "
                          "WILL lose that content",
            'links':      "hide the RVT links in those views",
        }

        recoverable, exhausted, no_safe_mode = [], [], []
        for item in selected_items:
            number = getattr(item, 'SheetNumber', None)
            if not number:
                continue
            for fmt, box in self._format_checkboxes():
                try:
                    if not box.IsChecked or not self._is_known_fatal(number, fmt):
                        continue
                except Exception:
                    continue
                if fmt not in ("DWG", "PDF"):
                    # Safe mode is only wired into these two loops.
                    no_safe_mode.append("{} ({})".format(number, fmt))
                elif self._safe_ladder_exhausted(number, fmt):
                    exhausted.append("{} ({})".format(number, fmt))
                else:
                    rung = self.SAFE_LADDER[self._next_safe_level(number, fmt)]
                    recoverable.append("{} ({}) -> {}".format(
                        number, fmt, RUNG_LABEL.get(rung, rung)))

        if not (recoverable or exhausted or no_safe_mode):
            return

        message = "These crashed Revit's exporter on an earlier run:\n  {}".format(
            "\n  ".join(recoverable + exhausted + no_safe_mode))
        message += ("\n\nThey are queued LAST, so everything else is already on "
                    "disk before Revit gets to them.")
        if exhausted or no_safe_mode:
            message += ("\n\nNo mitigation left for:\n  {}\n  (these go out at "
                        "full fidelity — the geometry itself has to be fixed)".format(
                            "\n  ".join(exhausted + no_safe_mode)))
        if recoverable:
            message += ("\n\nSafe mode applies the change above, exports, then "
                        "rolls it back so your model is left untouched. Each rung "
                        "is tried once; if it still crashes the next run "
                        "escalates automatically.")
        try:
            if self.export_pdf.IsChecked and self.combine_pdf.IsChecked:
                message += ("\n\nNote: Combine PDF puts every sheet in ONE native "
                            "call, so a crash there loses the whole combined file. "
                            "Uncheck Combine to export these individually.")
        except Exception:
            pass

        SAFE_OPT = "Export them with safe mode (recommended)"
        FULL_OPT = "Export them at full fidelity"
        SKIP_OPT = "Skip them (files will be missing)"
        options = ([SAFE_OPT] if recoverable else []) + [FULL_OPT, SKIP_OPT]

        choice = forms.alert(message, title="Items that crashed Revit before",
                             options=options)
        if choice == SKIP_OPT:
            self._skip_fatal = True
        elif choice == FULL_OPT:
            self._safe_mode = False
        else:
            # SAFE_OPT or dismissed: safest choice that still exports everything
            self._safe_mode = bool(recoverable)

    # Rungs that REMOVE content ('wireframe'/'coarse' only change how it is drawn)
    LOSSY_RUNGS = ('categories', 'links')

    def _note_safe_mode(self, sheet_number, fmt, rung, applied):
        """Record a safe-mode export for the completion dialog."""
        label = "{}{} -> {}".format(
            sheet_number, " ({})".format(fmt) if fmt else "", ", ".join(applied))
        if rung in self.LOSSY_RUNGS:
            label += "   [CONTENT HIDDEN - check this file]"
        if label not in self._safe_applied:
            self._safe_applied.append(label)

    def _note_skipped_fatal(self, sheet_number, fmt):
        """Record + surface a sheet skipped because it previously killed Revit."""
        label = "{} ({})".format(sheet_number, fmt)
        if label not in self._skipped_fatal:
            self._skipped_fatal.append(label)
        if not hasattr(self, '_failure_reasons'):
            self._failure_reasons = {}
        self._failure_reasons[(str(sheet_number), fmt)] = \
            "skipped: crashed Revit's {} exporter on an earlier run".format(fmt)
        msg = "Skipped {} — this sheet crashed Revit's {} exporter before.".format(
            sheet_number, fmt)
        logger.warning(msg)
        self._update_progress(
            int(self._overall_counter * 100.0 / max(1, self._overall_total)), msg)
        try:
            self.update_export_item_progress(sheet_number, fmt, 0, QUEUE_SKIPPED)
        except Exception:
            pass

    def _raster_level_for(self, sheet_number):
        """Starting RASTER_LADDER rung for this sheet. Always None (Presentation):
        the crash also hits DWG, which has no raster setting. Kept as the hook
        for per-sheet quality."""
        return None

    def _warn_if_previous_crash(self):
        """If a prior export died mid-sheet, tell the user which sheet+format."""
        pc = self._pending_crash
        if not pc:
            return
        try:
            sheet = pc.get('sheet', '?')
            fmt = pc.get('format', '?')
            same_doc = (not pc.get('doc')) or pc.get('doc') == self.doc.Title
            if same_doc:
                msg = ("Revit crashed last time on sheet '{}' during {} export. "
                       "It will be exported LAST this run — you'll be asked "
                       "whether to use safe mode.").format(sheet, fmt)
            else:
                msg = "Previous export crashed on sheet '{}' ({}, a different model).".format(
                    sheet, fmt)
            # Status bar only: logger.warning would pop the output window on every start
            logger.debug(msg)
            try:
                self.status_text.Text = msg
            except Exception:
                pass
        except Exception as ex:
            logger.debug("Could not surface previous crash: {}".format(ex))

    # ── Nothing exports silently ──
    # 1. filename collision: a Floor Plan and Ceiling Plan both named "Level 1"
    #    collapse onto one file (see _reserve_filename);
    # 2. verification declares failures, so it must not give false negatives.

    def _reset_run_state(self):
        """Clear per-run bookkeeping (filenames used, failures recorded)."""
        self._used_filenames = set()
        self._failed_items = []
        self._verify_misses = 0
        self._failure_reasons = {}  # (item number, format) -> reason, for the report
        self._last_output_file = ''  # last confirmed file; update_export_item_progress reads it

    def _native_output_stamp(self, path):
        """Snapshot one nonempty expected file, never a neighbouring export."""
        try:
            stat = os.stat(path)
            if os.path.isfile(path) and stat.st_size > 0:
                return stat.st_size, stat.st_mtime_ns
        except OSError:
            pass
        return None

    def _confirm_native_output(self, path, before):
        previous_mtime = before[1] / 1000000000.0 if before is not None else None
        self._wait_for_export_file(path, previous_mtime, 3.0)
        after = self._native_output_stamp(path)
        confirmed = after is not None and after != before
        if confirmed:
            self._last_output_file = path
        return confirmed

    def _reserve_filename(self, folder, base, ext):
        """Filename in `folder` unclaimed in this run; " (2)", " (3)", … on collision."""
        if not hasattr(self, '_used_filenames'):
            self._used_filenames = set()
        base = base.strip() or "Unnamed"
        candidate = base
        counter = 1
        while True:
            key = os.path.join(folder, candidate + ext).lower()
            if key not in self._used_filenames:
                self._used_filenames.add(key)
                if candidate != base:
                    logger.warning(
                        "Filename collision: '{}{}' was already written in this "
                        "run — this item goes to '{}{}' instead.".format(
                            base, ext, candidate, ext))
                return candidate
            counter += 1
            candidate = "{} ({})".format(base, counter)

    def _verify_export(self, expected_file, prev_mtime, started_at, ext):
        """True when the export produced a file.

        Falls back to an unreserved file written after `started_at` (some
        exporters decorate the name). After two misses in a row the wait is cut
        short so a failing format doesn't add minutes of polling.
        """
        timeout = 0.5 if getattr(self, '_verify_misses', 0) >= 2 else 3.0
        if self._wait_for_export_file(expected_file, prev_mtime, timeout):
            self._verify_misses = 0
            self._last_output_file = expected_file
            return True
        try:
            import glob as _glob
            folder = os.path.dirname(expected_file)
            used = getattr(self, '_used_filenames', set())
            self._used_filenames = used
            for path in _glob.glob(os.path.join(folder, "*" + ext)):
                if path.lower() == expected_file.lower() or path.lower() in used:
                    continue
                try:
                    if (os.path.isfile(path) and os.path.getsize(path) > 0
                            and os.path.getmtime(path) >= started_at):
                        logger.warning("'{}' landed as '{}' (Revit renamed it).".format(
                            os.path.basename(expected_file), os.path.basename(path)))
                        self._verify_misses = 0
                        used.add(path.lower())
                        self._last_output_file = path
                        return True
                except Exception:
                    continue
        except Exception as ex:
            logger.debug("Export verification fallback failed: {}".format(ex))
        self._verify_misses = getattr(self, '_verify_misses', 0) + 1
        return False

    def _note_failure(self, label, fmt, reason):
        """Record + surface an item that produced no file."""
        reason = str(reason)
        if len(reason) > 160:
            reason = reason[:157] + "..."
        entry = "{} ({}) — {}".format(label, fmt, reason)
        if entry not in self._failed_items:
            self._failed_items.append(entry)
        if not hasattr(self, '_failure_reasons'):
            self._failure_reasons = {}
        self._failure_reasons[(str(label), fmt)] = reason
        logger.error("NOT exported: {}".format(entry))
        try:
            self.update_export_item_progress(label, fmt, 0, QUEUE_FAILED)
        except Exception:
            pass

    def _sync_item_labels(self, items, fmt):
        """Refresh each row's cached sheet number / view name from Revit.

        Per-item try/except: one unreadable row (deleted element, closed
        workset) must not drop every other item.
        """
        for item in items:
            try:
                if hasattr(item, 'Sheet'):
                    item.SheetNumber = item.Sheet.SheetNumber
                    item.SheetName = item.Sheet.Name
                elif hasattr(item, 'View'):
                    item.SheetNumber = item.View.Name
                    item.ViewName = item.View.Name
            except Exception as ex:
                self._note_failure(str(getattr(item, 'SheetNumber', 'Unknown')),
                                   fmt, "could not be read from the model: {}".format(ex))

    def _element_of(self, item):
        """(element, display_name) for a SheetItem or ViewItem; (None, label) otherwise."""
        if hasattr(item, 'Sheet'):
            return item.Sheet, item.Sheet.SheetNumber
        if hasattr(item, 'View'):
            return item.View, item.View.Name
        return None, str(getattr(item, 'SheetNumber', 'Unknown'))

    def _wait_for_export_file(self, path, prev_mtime, timeout=8.0):
        """True once `path` exists nonempty with mtime newer than prev_mtime (or
        new). Export() is synchronous, so usually true on the first check; no
        folder-wide globs (slow on OneDrive)."""
        import time as _t
        deadline = _t.monotonic() + max(0.0, timeout)
        while True:
            try:
                if os.path.isfile(path) and os.path.getsize(path) > 0:
                    if prev_mtime is None or os.path.getmtime(path) > prev_mtime:
                        return True
            except Exception:
                pass
            remaining = deadline - _t.monotonic()
            if remaining <= 0:
                return False
            _t.sleep(min(0.1, remaining))

    def load_sheet_sets_for_filter(self):
        """Populate the multi-select sheet set dropdown with CheckBoxes."""
        try:
            from System.Windows.Controls import CheckBox
            from System.Windows import Thickness

            self.sheet_set_checklist.Children.Clear()

            try:
                t3_cb_style = self.FindResource("T3.CheckBox")
            except Exception:
                t3_cb_style = None

            all_cb = CheckBox()
            all_cb.Content = "All Sheets/Views"
            all_cb.IsChecked = True
            all_cb.Tag = None
            all_cb.Margin = Thickness(4, 3, 4, 3)
            if t3_cb_style:
                all_cb.Style = t3_cb_style
            all_cb.Checked += self._sheet_set_all_checked
            self._sheet_set_all_checkbox = all_cb
            self.sheet_set_checklist.Children.Add(all_cb)

            saved_set_names = self.get_saved_sheet_set_names()
            if saved_set_names:
                from System.Windows.Controls import Separator
                sep = Separator()
                sep.Margin = Thickness(0, 2, 0, 2)
                self.sheet_set_checklist.Children.Add(sep)

                for set_name in sorted(saved_set_names):
                    cb = CheckBox()
                    cb.Content = set_name
                    cb.IsChecked = False
                    cb.Tag = set_name
                    cb.Margin = Thickness(4, 3, 4, 3)
                    if t3_cb_style:
                        cb.Style = t3_cb_style
                    cb.Checked += self._sheet_set_item_changed
                    cb.Unchecked += self._sheet_set_item_changed
                    self.sheet_set_checklist.Children.Add(cb)
            else:
                from System.Windows.Controls import TextBlock as TB
                from System.Windows.Media import SolidColorBrush, Color
                hint = TB()
                hint.Text = "(No saved Sheet Sets in this document)"
                hint.Foreground = SolidColorBrush(Color.FromRgb(0x7F, 0x8C, 0x8D))
                hint.Margin = Thickness(6, 4, 4, 4)
                self.sheet_set_checklist.Children.Add(hint)

        except Exception as ex:
            logger.warning("Could not load sheet sets for filter: {}".format(ex))

    def sheet_set_toggle_clicked(self, sender, e):
        """Toggle the multi-select sheet set popup open/closed."""
        try:
            self.sheet_set_popup.IsOpen = not self.sheet_set_popup.IsOpen
            if not self.sheet_set_popup.IsOpen:
                self.sheet_set_toggle.IsChecked = False
        except Exception as ex:
            logger.debug("Error toggling sheet set popup: {}".format(ex))

    def _sheet_set_all_checked(self, sender, e):
        """When 'All' is checked, uncheck every individual set and clear filter."""
        try:
            for child in self.sheet_set_checklist.Children:
                if hasattr(child, 'Tag') and child.Tag is not None:
                    child.IsChecked = False
            self._update_sheet_set_label()
            self.apply_filters()
            self.update_selection_count()
        except Exception as ex:
            logger.debug("Error in sheet set all-checked: {}".format(ex))

    def _sheet_set_item_changed(self, sender, e):
        """When any individual set checkbox changes, update filter and label."""
        try:
            any_checked = any(
                child.IsChecked
                for child in self.sheet_set_checklist.Children
                if hasattr(child, 'Tag') and child.Tag is not None
            )
            # Set "All" without firing its handler
            self._sheet_set_all_checkbox.Checked -= self._sheet_set_all_checked
            self._sheet_set_all_checkbox.IsChecked = not any_checked
            self._sheet_set_all_checkbox.Checked += self._sheet_set_all_checked

            self._update_sheet_set_label()
            self._apply_sheet_set_filter()
        except Exception as ex:
            logger.debug("Error in sheet set item changed: {}".format(ex))

    def _update_sheet_set_label(self):
        """Update the toggle button text to reflect current selection."""
        try:
            checked = [
                child.Content
                for child in self.sheet_set_checklist.Children
                if hasattr(child, 'Tag') and child.Tag is not None and child.IsChecked
            ]
            if not checked:
                self.sheet_set_summary.Text = "All Sheets/Views"
            elif len(checked) == 1:
                self.sheet_set_summary.Text = checked[0]
            else:
                self.sheet_set_summary.Text = "{} sets selected".format(len(checked))
        except Exception as ex:
            logger.debug("Error updating sheet set label: {}".format(ex))

    def _apply_sheet_set_filter(self):
        """Queue the sheet-set filter — reads ViewSheetSets, needs API context."""
        self._run_in_api_context(self._apply_sheet_set_filter_api)

    def _apply_sheet_set_filter_api(self):
        """Union sheet IDs from all checked sets and apply filter + auto-select."""
        try:
            checked_sets = [
                child.Tag
                for child in self.sheet_set_checklist.Children
                if hasattr(child, 'Tag') and child.Tag is not None and child.IsChecked
            ]
            if not checked_sets:
                self.apply_filters()
                return

            all_ids = set()
            for set_name in checked_sets:
                ids = self.get_sheet_ids_from_set(set_name)
                all_ids.update(ids)

            if not all_ids:
                self.status_text.Text = "No sheets found in selected sets"
                return

            self.apply_filters(sheet_set_ids=list(all_ids))

            selected_count = 0
            for sheet_item in self.all_sheets:
                if sheet_item.Sheet.Id in all_ids:
                    sheet_item.IsSelected = True
                    selected_count += 1
                else:
                    sheet_item.IsSelected = False

            self.sheets_listview.Items.Refresh()
            self.update_selection_count()
            self.status_text.Text = "'{}': {} sheets selected".format(
                self.sheet_set_summary.Text, selected_count)

        except Exception as ex:
            logger.error("Error applying sheet set filter: {}".format(ex))

    def _on_item_property_changed(self, sender, args):
        """Recount when an item's IsSelected changes."""
        if getattr(self, '_batch_updating', False):
            return
        if args and getattr(args, 'PropertyName', None) == "IsSelected":
            self.update_selection_count()

    def load_sheets(self):
        """Phase 1: show sheet names at once; phase 2: load the rest in chunks."""
        try:
            sheets_collector = FilteredElementCollector(self.doc)\
                .OfCategory(BuiltInCategory.OST_Sheets)\
                .WhereElementIsNotElementType()

            sheets = [s for s in sheets_collector if isinstance(s, ViewSheet)]
            sheets.sort(key=lambda x: x.SheetNumber)

            self.all_sheets = []
            for sheet in sheets:
                item = SheetItem(sheet, False, lazy=True)
                item.add_PropertyChanged(self._on_item_property_changed)
                self.all_sheets.append(item)
            self.filtered_sheets = list(self.all_sheets)

            self.update_sheets_list()
            self.status_text.Text = "Loaded {} sheets | Revit {}".format(
                len(self.all_sheets), REVIT_VERSION)

            # Via _run_in_api_context, NOT Dispatcher.BeginInvoke: modeless
            # callbacks would read the API outside context and crash Revit.
            self._lazy_load_index = 0
            self._run_in_api_context(self._lazy_load_chunk)

        except Exception as ex:
            logger.error("Error loading sheets: {}".format(ex))
            # Không exitscript: SystemExit qua biên .NET (window modeless) có thể crash Revit.
            forms.alert("Error loading sheets: {}".format(ex))


    def _lazy_load_chunk(self):
        """Load Revision for a batch of sheets, then schedule the next batch.

        Size/Orientation wait until the Queue is built, so startup never
        queries titleblocks.
        """
        CHUNK_SIZE = 80

        try:
            start = self._lazy_load_index
            chunk = self.all_sheets[start:start + CHUNK_SIZE]

            if not chunk:
                if hasattr(self, 'sheets_listview') and self.sheets_listview:
                    self.sheets_listview.Items.Refresh()
                self.status_text.Text = "Ready | {} sheets | Revit {}".format(
                    len(self.all_sheets), REVIT_VERSION)
                return

            for sheet_item in chunk:
                sheet_item._load_revision_params()

            self._lazy_load_index = start + len(chunk)

            if hasattr(self, 'sheets_listview') and self.sheets_listview:
                self.sheets_listview.Items.Refresh()

            self._run_in_api_context(self._lazy_load_chunk)

        except Exception as ex:
            logger.debug("Error loading sheet chunk: {}".format(ex))

    def _ensure_titleblock_cache(self):
        """Populate titleblock size cache on first call (deferred from startup)."""
        if not getattr(self, '_titleblock_cache_loaded', False):
            self._batch_load_titleblock_sizes()


    def load_views(self):
        """Load all views — phase 1: instant display, phase 2: chunked lazy load.

        OfClass(View), NOT OfCategory(OST_Views): schedules are OST_Schedules and
        some views have no category. No isinstance whitelist: Legends are plain
        `View`s.
        """
        try:
            views_collector = FilteredElementCollector(self.doc)\
                .OfClass(View)\
                .WhereElementIsNotElementType()

            views = []
            excluded = 0
            for v in views_collector:
                try:
                    if v.IsTemplate:
                        continue
                    if isinstance(v, ViewSheet):
                        continue  # Sheets mode
                    if v.ViewType in NON_EXPORTABLE_VIEW_TYPES:
                        excluded += 1
                        continue
                    # Titleblock revision schedules are not standalone views
                    if isinstance(v, ViewSchedule) and \
                            getattr(v, 'IsTitleblockRevisionSchedule', False):
                        excluded += 1
                        continue
                    views.append(v)
                except Exception as view_ex:
                    excluded += 1
                    logger.debug("Skipping unreadable view: {}".format(view_ex))

            views.sort(key=lambda x: x.Name)
            if excluded:
                logger.debug("Views list: {} exportable, {} internal/system views "
                             "excluded".format(len(views), excluded))

            self.all_views = []
            for view in views:
                item = ViewItem(view, False, lazy=True)
                item.add_PropertyChanged(self._on_item_property_changed)
                self.all_views.append(item)
            self.filtered_views = list(self.all_views)

            self._fill_view_type_filter()
            self.update_items_list()
            self.status_text.Text = "Loaded {} views | Revit {}".format(
                len(self.all_views), REVIT_VERSION)

            # Phase 2: chunked Phase/ViewTemplate loading
            self._lazy_view_load_index = 0
            self._run_in_api_context(self._lazy_view_load_chunk)

        except Exception as ex:
            logger.error("Error loading views: {}".format(ex))
            # Không exitscript: SystemExit qua biên .NET (window modeless) có thể crash Revit.
            forms.alert("Error loading views: {}".format(ex))

    def _fill_view_type_filter(self):
        """Fill the view-type dropdown from the loaded views' ViewItem.ViewType
        labels (enum names like "ThreeD" matched no row). Keeps the current pick.
        """
        combo = getattr(self, 'view_type_filter', None)
        if combo is None:
            return
        try:
            from System.Windows.Controls import ComboBoxItem
            current = None
            if combo.SelectedIndex > 0 and combo.SelectedItem is not None:
                current = str(combo.SelectedItem.Content)
            labels = sorted(set(v.ViewType for v in self.all_views), key=lambda s: s.lower())
            self._filling_view_types = True
            try:
                combo.Items.Clear()
                for label in ["All Types"] + labels:
                    item = ComboBoxItem()
                    item.Content = label
                    combo.Items.Add(item)
                combo.SelectedIndex = labels.index(current) + 1 if current in labels else 0
            finally:
                self._filling_view_types = False
        except Exception as ex:
            logger.debug("Could not fill the view type filter: {}".format(ex))

    def _lazy_view_load_chunk(self):
        """Load Phase/ViewTemplate for a batch of views, then schedule the next batch."""
        CHUNK_SIZE = 80
        try:
            start = self._lazy_view_load_index
            chunk = self.all_views[start:start + CHUNK_SIZE]

            if not chunk:
                if hasattr(self, 'sheets_listview') and self.sheets_listview:
                    self.sheets_listview.Items.Refresh()
                self.status_text.Text = "Ready | {} views | Revit {}".format(
                    len(self.all_views), REVIT_VERSION)
                return

            for view_item in chunk:
                view_item._load_extra_data()

            self._lazy_view_load_index = start + len(chunk)

            if hasattr(self, 'sheets_listview') and self.sheets_listview:
                self.sheets_listview.Items.Refresh()

            self._run_in_api_context(self._lazy_view_load_chunk)

        except Exception as ex:
            logger.debug("Error loading view chunk: {}".format(ex))

    def update_sheets_list(self):
        """Update the sheets ListView."""
        self.sheets_listview.ItemsSource = None
        self.sheets_listview.ItemsSource = to_items_source(self.filtered_sheets)

    def update_items_list(self):
        """Update the items ListView based on current selection mode."""
        if self.selection_mode == "sheets":
            self.update_sheets_list()
        else:
            self.update_views_list()

        self.update_selection_count()

    def update_views_list(self):
        """Update the views ListView."""
        self.sheets_listview.ItemsSource = None
        self.sheets_listview.ItemsSource = to_items_source(self.filtered_views)

    def update_selection_count(self):
        """Update the selection count status bar and header checkbox."""
        try:
            if getattr(self, '_batch_updating', False):
                return

            items = self.filtered_sheets if self.selection_mode == "sheets" else self.filtered_views
            total_count = len(items)
            selected_count = sum(1 for s in items if s.IsSelected)
            all_items = self.all_sheets if self.selection_mode == "sheets" else self.all_views
            total_selected = sum(1 for item in all_items if item.IsSelected)
            hidden_selected = total_selected - selected_count

            if hasattr(self, 'selection_count_text'):
                label = "{} {} selected | {} shown".format(
                    total_selected, self.selection_mode, total_count)
                if hidden_selected:
                    label += " | {} selected hidden by filters".format(hidden_selected)
                self.selection_count_text.Text = label

            # Header checkbox: True = all, False = none, None = some
            if hasattr(self, 'header_checkbox') and self.header_checkbox:
                if total_count == 0 or selected_count == 0:
                    self.header_checkbox.IsChecked = False
                elif selected_count == total_count:
                    self.header_checkbox.IsChecked = True
                else:
                    self.header_checkbox.IsChecked = None

        except Exception as ex:
            logger.debug("Error updating selection count: {}".format(ex))

    def selection_mode_changed(self, sender, e):
        """Handle selection mode change (Sheets vs Views radio button)."""
        try:
            if hasattr(self, 'sheets_radio') and self.sheets_radio.IsChecked:
                self.selection_mode = "sheets"
                if not self.all_sheets:
                    self.load_sheets()
                else:
                    self.update_items_list()
                if hasattr(self, 'view_type_filter'):
                    self.view_type_filter.Visibility = Visibility.Collapsed
                if hasattr(self, 'col_number'):
                    self.col_number.Header = "Sheet Number"
                    self.col_name.Header = "Sheet Name"
                    self.col_revision.Header = "Revision"
                    self.col_size.Header = "Size"
                if hasattr(self, 'col_orientation'):
                    self.col_orientation.Header = "Orientation"
            elif hasattr(self, 'views_radio') and self.views_radio.IsChecked:
                self.selection_mode = "views"
                if not self.all_views:
                    # First load queries the document → API context
                    self._run_in_api_context(self.load_views)
                else:
                    self.update_items_list()
                if hasattr(self, 'view_type_filter'):
                    self.view_type_filter.Visibility = Visibility.Visible
                if hasattr(self, 'col_number'):
                    self.col_number.Header = "View Name"
                    self.col_name.Header = "View Type"
                    self.col_revision.Header = "Phase"
                    self.col_size.Header = "Scale"
                if hasattr(self, 'col_orientation'):
                    self.col_orientation.Header = "—"
        except Exception as ex:
            logger.error("Error changing selection mode: {}".format(ex))

    def listview_clicked(self, sender, e):
        """Stop TextBox clicks from toggling the row; selection sync is the
        TwoWay IsSelected binding + sheets_selection_changed."""
        try:
            from System.Windows.Controls import TextBox
            from System.Windows.Media import VisualTreeHelper

            cur = e.OriginalSource
            while cur is not None:
                if isinstance(cur, TextBox):
                    e.Handled = True
                    sender_tb = cur
                    sender_tb.Focus()
                    return
                try:
                    cur = VisualTreeHelper.GetParent(cur)
                except Exception:
                    break
        except Exception as ex:
            logger.debug("Error handling listview click: {}".format(ex))

    def sheets_selection_changed(self, sender, e):
        """WPF selection changed; the TwoWay IsSelected binding already updated the model."""
        try:
            if getattr(self, '_batch_updating', False):
                return
            self.update_selection_count()
            self.update_export_preview_if_needed()
        except Exception as ex:
            logger.debug("Error handling sheets selection changed: {}".format(ex))

    def row_checkbox_clicked(self, sender, e):
        try:
            # The ONLY writer of the model: the row CheckBox is bound OneWay
            # (through the hidden row_selected_text bridge in the XAML).
            if hasattr(sender, 'DataContext') and sender.DataContext is not None:
                sender.DataContext.IsSelected = bool(sender.IsChecked)

            # SheetItem has no INotifyPropertyChanged: re-read every row.
            self.sheets_listview.Items.Refresh()

            self.update_selection_count()
            self.update_export_preview_if_needed()
        except Exception as ex:
            logger.debug("Error handling row checkbox click: {}".format(ex))

    def listview_item_double_clicked(self, sender, e):
        """No-op, so a double-click does not toggle twice."""
        pass

    def on_listview_size_changed(self, sender, e):
        """Size Sheet Name + Custom Filename columns to fill the ListView (zero gap)."""
        try:
            col_num_w = int(self.col_number.Width) if hasattr(self, 'col_number') and self.col_number.Width > 0 else 220
            # checkbox + number + rev + size + orientation + scrollbar/borders
            fixed = 40 + col_num_w + 70 + 70 + 90 + 8
            available = sender.ActualWidth - fixed
            if available <= 0:
                return
            col_name_width = int(available * 0.38)
            self.col_name.Width = col_name_width
            # Last column takes the rest: zero gap at the right edge
            self.col_custom_filename.Width = max(100, int(available - col_name_width))
        except Exception:
            pass

    def window_content_rendered(self, sender, e):
        """Re-run the column fill once layout is final: the first SizeChanged
        (ItemsSource set in __init__) can see a stale width and leave a gap."""
        try:
            self.on_listview_size_changed(self.sheets_listview, None)
        except Exception:
            pass

    def header_checkbox_clicked(self, sender, e):
        """Select / deselect all visible items."""
        try:
            items = self.filtered_sheets if self.selection_mode == "sheets" else self.filtered_views
            if not items:
                return

            all_selected = all(item.IsSelected for item in items)
            target_state = not all_selected

            self._batch_updating = True
            try:
                for item in items:
                    item.IsSelected = target_state
            finally:
                self._batch_updating = False

            self.sheets_listview.Items.Refresh()

            if self.selection_mode == "sheets":
                if target_state:
                    self.status_text.Text = "Selected {} sheets".format(len(items))
                else:
                    self.status_text.Text = "Deselected all sheets"
            else:
                if target_state:
                    self.status_text.Text = "Selected {} views".format(len(items))
                else:
                    self.status_text.Text = "Deselected all views"

            self.update_selection_count()
            self.update_export_preview_if_needed()
        except Exception as ex:
            logger.debug("Error in header_checkbox_clicked: {}".format(ex))


    def get_saved_sheet_set_names(self):
        """Names of saved ViewSheetSets (collector, so PrintManager state is untouched)."""
        try:
            collector = FilteredElementCollector(self.doc).OfClass(ViewSheetSet)
            return [print_set.Name for print_set in collector]
        except Exception as ex:
            logger.debug("Error getting sheet set names: {}".format(ex))
            return []

    def get_sheet_ids_from_set(self, set_name):
        """ElementIds of the views in the named ViewSheetSet ([] if not found)."""
        try:
            collector = FilteredElementCollector(self.doc).OfClass(ViewSheetSet)
            for print_set in collector:
                if print_set.Name == set_name:
                    return [v.Id for v in print_set.Views]
            return []
        except Exception as ex:
            logger.error("Error getting sheets from set '{}': {}".format(set_name, ex))
            return []

    def save_current_selection_as_vs_set(self):
        """Save current selection as a new View/Sheet Set."""
        try:
            if self.selection_mode == "sheets":
                selected_items = [s for s in self.filtered_sheets if s.IsSelected]
                mode_name = "Sheet"
            else:
                selected_items = [v for v in self.filtered_views if v.IsSelected]
                mode_name = "View"

            if not selected_items:
                forms.alert("No {} selected.\n\nPlease select at least one {} first.".format(
                    mode_name.lower() + "s", mode_name.lower()),
                    title="No Selection")
                return

            set_name = forms.ask_for_string(
                prompt="Enter a name for the new {}/Sheet Set:".format(mode_name),
                title="Save {}/Sheet Set".format(mode_name),
                default="Custom Selection {}".format(datetime.now().strftime("%Y-%m-%d %H:%M"))
            )

            if not set_name:
                return

            views = [item.Sheet if self.selection_mode == "sheets" else item.View
                     for item in selected_items]

            # PrintManager.ViewSheetSetting exists only while PrintRange is Select
            # — reading it straight off the manager threw "This property is only
            # available when user choose Select of Print Range" (2026-11-20).
            # ViewSheetSetsService sets the range, assigns the views before SaveAs
            # and puts the document's print settings back.
            service = ViewSheetSetsService(self.doc)
            with disposing(Transaction(self.doc, "Save View/Sheet Set")) as t:
                t.Start()
                service.create_set_from_views(set_name, views)
                t.Commit()

            self.load_sheet_sets_for_filter()

            self.status_text.Text = "Saved '{}' with {} {}s".format(
                set_name, len(selected_items), mode_name.lower())

            forms.alert("Successfully saved '{}' with {} {}s".format(
                set_name, len(selected_items), mode_name.lower()),
                title="Success")

        except Exception as ex:
            logger.error("Error saving View/Sheet Set: {}".format(ex))
            forms.alert("Error saving View/Sheet Set:\n{}".format(str(ex)), title="Error")

    def search_sheets(self, sender, e):
        self.apply_filters()

    def filter_by_size(self, sender, e):
        """Filter sheets by size / views by view type."""
        if getattr(self, '_filling_view_types', False):
            return      # _fill_view_type_filter is rebuilding the dropdown
        self.apply_filters()

    def filter_by_sheet_set(self, sender, e):
        """Legacy handler — delegates to the multi-select implementation."""
        self._apply_sheet_set_filter()


    def save_vs_set_clicked(self, sender, e):
        try:
            # Creates a ViewSheetSet inside a Transaction — needs API context
            self._run_in_api_context(self.save_current_selection_as_vs_set)

        except Exception as ex:
            logger.error("Error handling Save V/S Set: {}".format(ex))

    def refresh_clicked(self, sender, e):
        """Reload sheets/views created or renamed while this modeless window was open."""
        self.status_text.Text = "Refreshing..."
        self._run_in_api_context(self._do_refresh)

    def _do_refresh(self):
        """Re-collect sheets/views/sets from the model, preserving current selection."""
        try:
            selected_sheet_ids = set(eid_value(s.Sheet.Id) for s in self.all_sheets if s.IsSelected)
            selected_view_ids = set(eid_value(v.View.Id) for v in self.all_views if v.IsSelected)
            had_views = bool(self.all_views)

            # Drop stale caches so sheet sizes re-detect against the live model
            self._titleblock_size_cache = {}
            self._paper_size_cache = {}
            self._titleblock_cache_loaded = False

            self.load_cad_export_setups()
            self.load_sheet_sets_for_filter()

            self.load_sheets()
            for item in self.all_sheets:
                if eid_value(item.Sheet.Id) in selected_sheet_ids:
                    item.IsSelected = True

            # Views load lazily on first mode switch — only reload if present
            if had_views or self.selection_mode == "views":
                self.load_views()
                for item in self.all_views:
                    if eid_value(item.View.Id) in selected_view_ids:
                        item.IsSelected = True

            self.apply_filters()
            self.update_selection_count()

            count = len(self.all_sheets) if self.selection_mode == "sheets" else len(self.all_views)
            self.status_text.Text = "Refreshed — {} {} loaded | Revit {}".format(
                count, self.selection_mode, REVIT_VERSION)

        except Exception as ex:
            logger.error("Error refreshing document data: {}".format(ex))
            self.status_text.Text = "Refresh failed — see pyRevit log"

    def apply_filters(self, sheet_set_ids=None):
        """Apply search + filters; sheet_set_ids optionally limits sheets to a ViewSheetSet."""
        # Not yet hydrated (called during XAML loading)
        if not hasattr(self, 'search_textbox'):
            return

        search_text = self.search_textbox.Text.lower() if self.search_textbox.Text else ""

        if self.selection_mode == "sheets":
            self.filtered_sheets = []
            for sheet in self.all_sheets:
                if sheet_set_ids is not None:
                    if sheet.Sheet.Id not in sheet_set_ids:
                        continue

                if search_text:
                    if search_text not in sheet.SheetNumber.lower() and \
                       search_text not in sheet.SheetName.lower():
                        continue

                self.filtered_sheets.append(sheet)

            self.update_items_list()

            if sheet_set_ids is not None:
                # Status is set in filter_by_sheet_set method
                pass
            else:
                self.status_text.Text = "Found {} sheets".format(len(self.filtered_sheets))
        else:
            # Index 0 = "All Types", read by position (2026-10-10: a text test had
            # drifted from the XAML label and filtered every view out).
            view_type_filter = None
            combo = getattr(self, 'view_type_filter', None)
            if combo is not None and combo.SelectedIndex > 0 and combo.SelectedItem is not None:
                view_type_filter = str(combo.SelectedItem.Content)

            self.filtered_views = []
            for view in self.all_views:
                if search_text:
                    if search_text not in view.ViewName.lower() and \
                       search_text not in view.ViewType.lower():
                        continue

                if view_type_filter:
                    if view.ViewType != view_type_filter:
                        continue

                self.filtered_views.append(view)

            self.update_items_list()
            self.status_text.Text = "Found {} views".format(len(self.filtered_views))

    def browse_output_folder(self, sender, e):
        dialog = FolderBrowserDialog()
        dialog.Description = "Select output folder for exports"
        dialog.SelectedPath = self.output_folder.Text

        if dialog.ShowDialog() == DialogResult.OK:
            self.output_folder.Text = dialog.SelectedPath
            self.status_text.Text = "Output folder: {}".format(dialog.SelectedPath)

    # ── Accordion toggle helpers ──
    def _toggle_format_panel(self, body_name, arrow_name, border_name):
        """Expand or collapse a format settings panel."""
        try:
            body = getattr(self, body_name)
            arrow = getattr(self, arrow_name)
            border = getattr(self, border_name)
            from System.Windows import Visibility
            if body.Visibility == Visibility.Collapsed:
                body.Visibility = Visibility.Visible
                arrow.Text = "\uE70E"
                border.BorderBrush = self.FindResource("T3.Ink")
            else:
                body.Visibility = Visibility.Collapsed
                arrow.Text = "\uE70D"
                border.BorderBrush = self.FindResource("T3.Border")
        except Exception as ex:
            logger.debug("Error toggling panel {}: {}".format(body_name, ex))

    def pdf_header_clicked(self, sender, e):
        self._toggle_format_panel("pdf_settings_body", "pdf_expand_arrow",
                                  "pdf_panel_border")

    def dwg_header_clicked(self, sender, e):
        self._toggle_format_panel("dwg_settings_body", "dwg_expand_arrow",
                                  "dwg_panel_border")

    def dgn_header_clicked(self, sender, e):
        self._toggle_format_panel("dgn_settings_body", "dgn_expand_arrow",
                                  "dgn_panel_border")

    def dwf_header_clicked(self, sender, e):
        self._toggle_format_panel("dwf_settings_body", "dwf_expand_arrow",
                                  "dwf_panel_border")

    def nwc_header_clicked(self, sender, e):
        self._toggle_format_panel("nwc_settings_body", "nwc_expand_arrow",
                                  "nwc_panel_border")

    def ifc_header_clicked(self, sender, e):
        self._toggle_format_panel("ifc_settings_body", "ifc_expand_arrow",
                                  "ifc_panel_border")

    def img_header_clicked(self, sender, e):
        self._toggle_format_panel("img_settings_body", "img_expand_arrow",
                                  "img_panel_border")

    def format_changed(self, sender, e):
        formats = []
        if self.export_pdf.IsChecked:
            formats.append("PDF")
        if self.export_dwg.IsChecked:
            formats.append("DWG")
        if self.export_dgn.IsChecked:
            formats.append("DGN")
        if self.export_dwf.IsChecked:
            formats.append("DWF")
        if self.export_nwd.IsChecked:
            formats.append("NWC")
        if self.export_ifc.IsChecked:
            formats.append("IFC")
        if self.export_img.IsChecked:
            formats.append("IMG")

        if formats:
            self.status_text.Text = "Selected formats: {}".format(", ".join(formats))

        self.update_export_preview_if_needed()

    def _pdf_engine(self):
        """"PDF24" when the PDF24 printer engine is picked, else "Revit"."""
        try:
            return "PDF24" if self.pdf_engine.SelectedIndex == 1 else "Revit"
        except Exception:
            return "Revit"

    def pdf_engine_changed(self, sender, e):
        """Say whether PDF24 can be used before the user starts an export."""
        try:
            if self._pdf_engine() != "PDF24":
                self.pdf24_status.Visibility = Visibility.Collapsed
                return
            printer = batchout_pdf24.find_printer()
            if not printer:
                text = ("PDF24 printer not found. Install PDF24 Creator ({}) or pick "
                        "Revit (built-in).".format(batchout_pdf24.INSTALL_URL))
            else:
                blocked = batchout_pdf24.AutoSave().blocked_reason()
                if blocked:
                    text = "PDF24 cannot be used: {}.".format(blocked)
                else:
                    text = "Prints to '{}'. Your PDF24 settings are restored after the export.".format(printer)
                    if self.combine_pdf.IsChecked and not batchout_pdf24.find_doctool():
                        text += " pdf24-DocTool.exe not found: Combine keeps one PDF per sheet."
            self.pdf24_status.Text = text
            self.pdf24_status.Visibility = Visibility.Visible
        except Exception as ex:
            logger.debug("PDF24 status check failed: {}".format(ex))

    def pdf_auto_detect_changed(self, sender, e):
        try:
            is_auto = self.pdf_auto_detect_size.IsChecked

            self.pdf_paper_size.IsEnabled = not is_auto
            self.pdf_landscape.IsEnabled = not is_auto
            self.pdf_portrait.IsEnabled = not is_auto

            if is_auto:
                self.status_text.Text = "Paper size and orientation will be auto-detected from Title Block"
            else:
                self.status_text.Text = "Using manual paper size and orientation settings"

        except Exception as ex:
            logger.error("Error handling auto-detect change: {}".format(ex))


    def edit_filename_clicked(self, sender, e):
        """Queue the filename pattern dialog (reads doc parameters: API context)."""
        self._run_in_api_context(self._edit_filename_api)

    def _edit_filename_api(self):
        """Edit the filename pattern via ParameterSelectorDialog and re-apply it to every item."""
        try:
            try:
                from GUI.ParameterSelectorDialog import ParameterSelectorDialog
            except ImportError:
                sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(__file__)))), 'lib', 'GUI'))
                if 'ParameterSelectorDialog' in sys.modules:
                    del sys.modules['ParameterSelectorDialog']
                from ParameterSelectorDialog import ParameterSelectorDialog

            element_type = 'sheet' if self.selection_mode == 'sheets' else 'view'

            pattern = ParameterSelectorDialog.show_dialog(self.doc, element_type)

            if pattern:
                self.naming_pattern.Text = pattern

                if hasattr(self, 'dwg_filename_sample') and self.dwg_filename_sample:
                    self.dwg_filename_sample.Text = pattern

                items_list = self.all_sheets if self.selection_mode == 'sheets' else self.all_views

                for item in items_list:
                    filename = self.get_export_filename(item)
                    item.CustomFilename = filename

                self.sheets_listview.Items.Refresh()

                self.status_text.Text = "Filename pattern updated: {}".format(pattern)

        except Exception as ex:
            logger.error("Error editing filename pattern: {}".format(ex))
            forms.alert("Error editing filename pattern:\n{}".format(str(ex)))


    def reverse_order_changed(self, sender, e):
        self.filtered_sheets.reverse()
        self.update_sheets_list()

    def nav_toggle_clicked(self, sender, e):
        try:
            target_index = -1
            if sender.Name == "nav_toggle_selection": target_index = 0
            elif sender.Name == "nav_toggle_format": target_index = 1
            elif sender.Name == "nav_toggle_queue": target_index = 2
            if target_index != -1:
                self.switch_to_view(target_index)
        except Exception as ex:
            logger.debug("Error in nav_toggle_clicked: {}".format(ex))

    def switch_to_view(self, index):
        """Switch to the given wizard tab (0=Selection, 1=Format, 2=Queue+Settings)."""
        try:
            # Leaving the Selection tab needs at least one item ticked
            if index > 0 and self.main_tabs.SelectedIndex == 0:
                if self.selection_mode == "sheets":
                    selected_items = [s for s in self.all_sheets if s.IsSelected]
                    if not selected_items:
                        forms.alert("Please select at least one sheet to export.", title="No Sheets Selected")
                        self.sync_navigation_footer(0)
                        return
                else:  # views mode
                    selected_items = [v for v in self.all_views if v.IsSelected]
                    if not selected_items:
                        forms.alert("Please select at least one view to export.", title="No Views Selected")
                        self.sync_navigation_footer(0)
                        return

            self.main_tabs.SelectedIndex = index
            self.update_navigation_buttons(index)
        except Exception as ex:
            logger.debug("Error in switch_to_view: {}".format(ex))

    def sync_navigation_footer(self, index=None):
        """Update the visual state of the sidebar navigation toggles."""
        try:
            if index is None:
                index = self.main_tabs.SelectedIndex

            self.nav_toggle_selection.IsChecked = (index == 0)
            self.nav_toggle_format.IsChecked = (index == 1)
            self.nav_toggle_queue.IsChecked = (index == 2)
        except Exception as ex:
            logger.debug("Error syncing navigation footer: {}".format(ex))

    def update_navigation_buttons(self, index=None):
        """Update action bar buttons and footer highlight for the current tab."""
        if index is None:
            index = self.main_tabs.SelectedIndex

        self.back_button.Visibility = Visibility.Collapsed if index == 0 else Visibility.Visible

        # Last tab (Queue+Settings) → Export; all others → Next
        if index == 2:
            self.next_button_text.Text = "Export"
            self.next_button_icon.Text = ""
        else:
            self.next_button_text.Text = "Next"
            self.next_button_icon.Text = "→"

        self.sync_navigation_footer(index)

    def tab_changed(self, sender, e):
        """Sync nav footer and trigger Queue preview when tab changes via keyboard/programmatic switch."""
        try:
            from System.Windows.Controls import TabControl as _TC
            if not isinstance(e.Source, _TC):
                return
            idx = self.main_tabs.SelectedIndex
            self.update_navigation_buttons(idx)
            if idx == 2:
                self.update_export_preview_if_needed()
        except Exception as ex:
            logger.debug("Error handling tab change: {}".format(ex))

    def update_export_preview_if_needed(self):
        """Refresh the Queue preview list when on tab 2."""
        try:
            if self.main_tabs.SelectedIndex == 2:
                if self.selection_mode == "sheets":
                    selected_items = [s for s in self.all_sheets if s.IsSelected]
                else:
                    selected_items = [v for v in self.all_views if v.IsSelected]

                if selected_items:
                    # Reads live sheet numbers + titleblock sizes → API context
                    self._run_in_api_context(self.build_export_preview)
                else:
                    self.export_items = []
                    self.export_preview_list.ItemsSource = to_items_source(self.export_items)
                    self.progress_text.Text = "No items selected for export"
        except Exception as ex:
            logger.debug("Error updating export preview: {}".format(ex))

    def go_back(self, sender, e):
        idx = self.main_tabs.SelectedIndex
        if idx > 0:
            self.switch_to_view(idx - 1)

    def go_next(self, sender, e):
        """Advance to next tab or start export on the final tab."""
        idx = self.main_tabs.SelectedIndex
        if idx == 0:
            self.switch_to_view(1)
        elif idx == 1:
            self.switch_to_view(2)
        else:
            if self._export_running:
                return
            self._export_running = True
            # Freeze the window: Revit's PDF exporter pumps messages, so a click
            # would re-enter WPF handlers inside Execute (fatal crash).
            self.IsEnabled = False
            # Runs in API context; fires once Revit is idle.
            self.status_text.Text = "Export queued — waiting for Revit..."
            if not self._run_in_api_context(self.start_export):
                self._export_running = False
                self.IsEnabled = True

    def build_export_preview(self):
        """Build the export preview list."""
        try:
            # Deferred from startup: filling it triggers a graphics regeneration.
            self._ensure_titleblock_cache()

            if self.selection_mode == "sheets":
                selected_items = [s for s in self.all_sheets if s.IsSelected]
                # Cached numbers/names may be stale: re-read from Revit
                for sheet_item in selected_items:
                    sheet_item.SheetNumber = sheet_item.Sheet.SheetNumber
                    sheet_item.SheetName = sheet_item.Sheet.Name
            else:
                selected_items = [v for v in self.all_views if v.IsSelected]
                for view_item in selected_items:
                    view_item.ViewName = view_item.View.Name
                    view_item.SheetNumber = view_item.View.Name

            formats = []
            if self.export_pdf.IsChecked:
                formats.append("PDF")
            if self.export_dwg.IsChecked:
                formats.append("DWG")
            if self.export_dgn.IsChecked:
                formats.append("DGN")
            if self.export_dwf.IsChecked:
                formats.append("DWF")
            if self.export_nwd.IsChecked:
                formats.append("NWC")
            if self.export_ifc.IsChecked:
                formats.append("IFC")
            if self.export_img.IsChecked:
                formats.append("IMG")

            is_auto_detect = self.pdf_auto_detect_size.IsChecked if hasattr(self, 'pdf_auto_detect_size') and self.pdf_auto_detect_size.IsChecked is not None else False

            manual_size = "Use Sheet Size"
            if hasattr(self, 'pdf_paper_size') and self.pdf_paper_size.SelectedItem:
                manual_size = self.pdf_paper_size.SelectedItem.Content

            self.export_items = []
            for item in selected_items:
                if (is_auto_detect or manual_size == "Use Sheet Size") and self.selection_mode == "sheets" and hasattr(item, 'Sheet'):
                    detected_size, detected_orientation = self.get_sheet_paper_size_and_orientation(item.Sheet)
                    size = detected_size
                    orientation = detected_orientation
                else:
                    size = manual_size
                    orientation = "Landscape" if self.pdf_landscape.IsChecked else "Portrait"

                for fmt in formats:
                    preview_item = ExportPreviewItem(item, fmt, size, orientation)
                    self.export_items.append(preview_item)

            self.export_preview_list.ItemsSource = to_items_source(self.export_items)
            self.progress_text.Text = "Ready to export {} items".format(len(self.export_items))
        except Exception as ex:
            logger.error("Error building export preview: {}".format(ex))
            if hasattr(self, 'progress_text') and self.progress_text:
                self.progress_text.Text = "Error preparing preview: {}".format(ex)

    def get_export_filename(self, item):
        """Filename from the naming pattern, read live from the sheet/view (SheetItem or
        ViewItem); any element parameter can be a {placeholder}."""
        pattern = self.naming_pattern.Text

        try:
            project_info = self.doc.ProjectInformation
            project_number = project_info.Number or ""
            project_name = project_info.Name or ""
            project_address = project_info.Address or ""
            client_name = project_info.ClientName or ""
            project_status = project_info.Status or ""
        except:
            project_number = ""
            project_name = ""
            project_address = ""
            client_name = ""
            project_status = ""

        element = None
        if hasattr(item, 'Sheet'):
            element = item.Sheet
            sheet_number = element.SheetNumber
            sheet_name = element.Name
        elif hasattr(item, 'View'):
            element = item.View
            sheet_number = element.Name  # view name as "number"
            sheet_name = item.ViewType  # view type as "name"
        else:
            sheet_number = "Unknown"
            sheet_name = "Unknown"

        replacements = {
            "{SheetNumber}": sheet_number,
            "{SheetName}": sheet_name,
            "{ViewName}": sheet_number if hasattr(item, 'View') else "",
            "{ProjectNumber}": project_number,
            "{ProjectName}": project_name,
            "{ProjectAddress}": project_address,
            "{ClientName}": client_name,
            "{ProjectStatus}": project_status,
            "{Date}": datetime.now().strftime("%Y%m%d"),
            "{Time}": datetime.now().strftime("%H%M%S"),
        }

        # Resolve only placeholders in the pattern: "{SheetNumber}" alone skips
        # iterating every parameter (real per-sheet saving on big models).
        import re as _re
        pattern_refs = set(_re.findall(r'\{([^{}]+)\}', pattern))
        standard_keys = set(k[1:-1] for k in replacements.keys())
        needs_dynamic_params = bool(pattern_refs - standard_keys)

        if element and needs_dynamic_params:
            try:
                for param in element.Parameters:
                    try:
                        param_name = param.Definition.Name
                        param_value = ""

                        if param.HasValue:
                            if param.StorageType == DB.StorageType.String:
                                param_value = param.AsString() or ""
                            elif param.StorageType == DB.StorageType.Integer:
                                param_value = str(param.AsInteger())
                            elif param.StorageType == DB.StorageType.Double:
                                param_value = str(param.AsDouble())
                            elif param.StorageType == DB.StorageType.ElementId:
                                elem_id = param.AsElementId()
                                if elem_id and eid_value(elem_id) > 0:
                                    try:
                                        elem = self.doc.GetElement(elem_id)
                                        param_value = elem.Name if elem else ""
                                    except:
                                        param_value = str(eid_value(elem_id))

                        replacements["{" + param_name + "}"] = param_value

                    except Exception as param_ex:
                        logger.debug("Could not read parameter {}: {}".format(
                            param.Definition.Name if hasattr(param, 'Definition') else 'unknown',
                            str(param_ex)
                        ))
                        continue
            except Exception as params_ex:
                logger.warning("Could not iterate parameters: {}".format(str(params_ex)))

        # Sheet built-ins only when needed (~7 get_Parameter calls per sheet)
        if needs_dynamic_params and hasattr(item, 'Sheet'):
            try:
                rev_param = item.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION)
                replacements["{Revision}"] = rev_param.AsString() if rev_param else ""
            except:
                replacements["{Revision}"] = ""

            try:
                rev_date_param = item.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION_DATE)
                replacements["{RevisionDate}"] = rev_date_param.AsString() if rev_date_param else ""
            except:
                replacements["{RevisionDate}"] = ""

            try:
                rev_desc_param = item.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_CURRENT_REVISION_DESCRIPTION)
                replacements["{RevisionDescription}"] = rev_desc_param.AsString() if rev_desc_param else ""
            except:
                replacements["{RevisionDescription}"] = ""

            try:
                drawn_param = item.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_DRAWN_BY)
                replacements["{DrawnBy}"] = drawn_param.AsString() if drawn_param else ""
            except:
                replacements["{DrawnBy}"] = ""

            try:
                checked_param = item.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_CHECKED_BY)
                replacements["{CheckedBy}"] = checked_param.AsString() if checked_param else ""
            except:
                replacements["{CheckedBy}"] = ""

            try:
                approved_param = item.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_APPROVED_BY)
                replacements["{ApprovedBy}"] = approved_param.AsString() if approved_param else ""
            except:
                replacements["{ApprovedBy}"] = ""

            try:
                issue_date_param = item.Sheet.get_Parameter(DB.BuiltInParameter.SHEET_ISSUE_DATE)
                replacements["{IssueDate}"] = issue_date_param.AsString() if issue_date_param else ""
            except:
                replacements["{IssueDate}"] = ""

        elif hasattr(item, 'View'):
            try:
                replacements["{ViewType}"] = item.ViewType
                replacements["{Scale}"] = item.Scale if hasattr(item, 'Scale') else ""
            except:
                pass

            try:
                phase_param = element.get_Parameter(DB.BuiltInParameter.VIEW_PHASE)
                if phase_param:
                    phase_id = phase_param.AsElementId()
                    if phase_id and eid_value(phase_id) > 0:
                        phase = self.doc.GetElement(phase_id)
                        replacements["{Phase}"] = phase.Name if phase else ""
            except:
                replacements["{Phase}"] = ""

            try:
                level_param = element.get_Parameter(DB.BuiltInParameter.VIEW_LEVEL)
                if level_param:
                    level_id = level_param.AsElementId()
                    if level_id and eid_value(level_id) > 0:
                        level = self.doc.GetElement(level_id)
                        replacements["{Level}"] = level.Name if level else ""
            except:
                replacements["{Level}"] = ""

        filename = pattern
        for placeholder, value in replacements.items():
            filename = filename.replace(placeholder, str(value))

        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in invalid_chars:
            filename = filename.replace(char, '_')

        return filename

    def update_export_item_progress(self, sheet_number, format_name, progress, status=""):
        """Update one queue row; the row being exported is scrolled into view."""
        try:
            for item in self.export_items:
                if item.SheetNumber == sheet_number and item.Format == format_name:
                    item.Progress = progress
                    if status:
                        item.Status = status
                    elif progress >= 100:
                        item.Status = QUEUE_EXPORTED
                        item.OutputFile = getattr(self, '_last_output_file', '') or ''
                    elif progress > 0:
                        item.Status = QUEUE_EXPORTING
                    self.export_preview_list.Items.Refresh()
                    if item.Status == QUEUE_EXPORTING:
                        self.export_preview_list.ScrollIntoView(item)
                    break
        except Exception:
            pass
        self._flush_ui()

    def _refresh_queue(self):
        """Redraw every queue row from its Progress/Status, then repaint."""
        try:
            self.export_preview_list.Items.Refresh()
        except Exception:
            pass
        self._flush_ui()

    def _begin_live_progress(self):
        """Mark the whole queue Queued and let the window repaint during the run.

        Modeless: start_export runs inside ExternalEvent.Execute, so repainting
        needs a message pump between exports, and a pump with input live would
        deliver queued clicks mid-API-call. ThreadInputLock disables every window
        of the thread first (what ShowDialog does on the modal path). No lock:
        nothing pumps and the queue updates when Execute returns.
        """
        self._export_started_at = time.time()
        self._live_progress_running = True
        lock = getattr(self, '_input_lock', None)
        if lock is None:
            lock = self._input_lock = ThreadInputLock()
        if not lock.acquire() and self.modeless:
            logger.debug("BatchOut: could not lock Revit input; the queue "
                         "updates when the export finishes.")
        for row in self.export_items:
            row.Progress = 0
            row.Status = QUEUE_QUEUED
        self._refresh_queue()
        self._update_progress(0, "Exporting {} item(s)...".format(len(self.export_items)))

    def _end_live_progress(self):
        """Close out the run: rows never reached say so, Revit gets input back.
        Idempotent (called before the summary and again in start_export's finally)."""
        if not getattr(self, '_live_progress_running', False):
            return
        self._live_progress_running = False
        for row in self.export_items:
            if row.Status in (QUEUE_QUEUED, QUEUE_EXPORTING):
                row.Status = QUEUE_NOT_EXPORTED
        self._refresh_queue()
        lock = getattr(self, '_input_lock', None)
        if lock is not None and lock.active:
            lock.release()
            try:
                self.Activate()
            except Exception:
                pass

    def _flush_ui(self):
        """Let WPF render pending changes so the queue moves while exporting.

        A nested dispatcher frame delivers every waiting Win32 message. Modal:
        ShowDialog disabled Revit's windows, so only painting gets through.
        Modeless: inside ExternalEvent.Execute, with input live, re-entrancy is
        the fatal 0xc0000005 class, so pump only while _begin_live_progress
        holds the input lock.
        """
        if self.modeless:
            lock = getattr(self, '_input_lock', None)
            if lock is None or not lock.active:
                return
        try:
            self.Dispatcher.Invoke(DispatcherPriority.Render, Action(lambda: None))
        except Exception:
            pass

    def _update_progress(self, value, text):
        """Overall bar + status line; while exporting the right label adds
        done/total, elapsed and time-left estimate."""
        try:
            pct = max(0, min(100, int(value)))
            self.overall_progress.Value = pct
            self.progress_text.Text = text
            started = getattr(self, '_export_started_at', None)
            self.progress_percent.Text = format_progress_summary(
                getattr(self, '_overall_counter', 0), getattr(self, '_overall_total', 0), pct,
                elapsed=None if started is None else time.time() - started,
                running=getattr(self, '_live_progress_running', False))
            try:
                self.progress_percent.Foreground = self.FindResource(
                    "T3.Success.Text" if pct >= 100 else "T3.Progress.Fill")
            except Exception:
                pass
        except Exception:
            pass
        self._flush_ui()

    def start_export(self):
        total_exported = 0
        try:
            if not any(getattr(self, name).IsChecked for name in (
                    'export_pdf', 'export_dwg', 'export_dwf', 'export_dgn',
                    'export_nwd', 'export_ifc', 'export_img')):
                self.status_text.Text = "Choose at least one export format"
                forms.alert("Choose at least one format in Settings before exporting.",
                            title="No Export Format Selected")
                return
            self._ensure_titleblock_cache()

            if self.selection_mode == "sheets":
                selected_items = [s for s in self.all_sheets if s.IsSelected]
                if not selected_items:
                    forms.alert("Please select at least one sheet to export.", title="No Sheets Selected")
                    return
                item_type_name = "sheets"
            else:
                selected_items = [v for v in self.all_views if v.IsSelected]
                if not selected_items:
                    forms.alert("Please select at least one view to export.", title="No Views Selected")
                    return
                item_type_name = "views"

            if self.reverse_order.IsChecked:
                selected_items.reverse()

            output_folder = self.output_folder.Text
            if not output_folder:
                forms.alert("Please select an output folder.", title="No Output Folder")
                return

            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            split_by_format = self.save_split_by_format.IsChecked

            # Remember this setup for the next session
            self.save_latest_setup()

            self.next_button.IsEnabled = False
            self.back_button.IsEnabled = False
            self.status_text.Text = "Exporting..."

            # Rebuild every run: old statuses; selection/formats may have changed.
            self.build_export_preview()

            # Overall progress is per item via _overall_counter
            total_exported = 0
            self._overall_total = len(self.export_items)
            self._overall_counter = 0
            self._skipped_fatal = []  # reset: this window can export more than once
            self._safe_applied = []
            self._reset_run_state()

            # Known-fatal sheets: offer safe mode once, before anything runs.
            self._ask_safe_mode(selected_items)

            # Risky items last so a native crash cannot cost the rest of the batch
            selected_items = self._order_risky_last(selected_items)

            # Queue repaints per item from here: user dialogs go before this
            # line or after _end_live_progress.
            self._begin_live_progress()

            def confirmed_count(format_name, count, expected):
                if count < expected:
                    self._failed_items.append(
                        "{}: {} of {} expected output(s) confirmed".format(
                            format_name, count, expected))
                return count

            if self.export_dwg.IsChecked:
                folder = os.path.join(output_folder, "DWG") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_dwg(selected_items, folder)
                total_exported += confirmed_count("DWG", count, len(selected_items))

                # Free DWG-phase memory before the raster-heavy PDF phase.
                GC.Collect()
                GC.WaitForPendingFinalizers()

            if self.export_pdf.IsChecked:
                folder = os.path.join(output_folder, "PDF") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_pdf(selected_items, folder)
                total_exported += confirmed_count(
                    "PDF", count, 1 if self.combine_pdf.IsChecked else len(selected_items))

            if self.export_dwf.IsChecked:
                folder = os.path.join(output_folder, "DWF") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_dwf(selected_items, folder)
                total_exported += confirmed_count("DWF", count, len(selected_items))

            if self.export_dgn.IsChecked:
                folder = os.path.join(output_folder, "DGN") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_dgn(selected_items, folder)
                total_exported += confirmed_count("DGN", count, len(selected_items))

            if self.export_nwd.IsChecked:
                folder = os.path.join(output_folder, "NWC") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_nwd(selected_items, folder)
                total_exported += confirmed_count("NWC", count, len(selected_items))

            if self.export_ifc.IsChecked:
                folder = os.path.join(output_folder, "IFC") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_ifc(selected_items, folder)
                total_exported += confirmed_count("IFC", count, 1)

            if self.export_img.IsChecked:
                folder = os.path.join(output_folder, "Images") if split_by_format else output_folder
                if not os.path.exists(folder):
                    os.makedirs(folder)
                count = self.export_to_images(selected_items, folder)
                total_exported += confirmed_count("Images", count, len(selected_items))

            # Revit gets its input back before the summary dialog asks anything.
            self._end_live_progress()
            report_note = self._write_export_report(output_folder)

            if not total_exported:
                outcome = "No exports confirmed"
            elif self._failed_items or self._skipped_fatal:
                outcome = "Export finished with issues"
            else:
                outcome = "Export complete"
            summary = "{}: {} output(s) confirmed".format(outcome, total_exported)
            if self._failed_items:
                summary += " ({} issue report(s))".format(len(self._failed_items))
            if self._skipped_fatal:
                summary += " ({} skipped)".format(len(self._skipped_fatal))
            self.status_text.Text = summary
            self._update_progress(100, summary)
            self.next_button.IsEnabled = True
            self.back_button.IsEnabled = True

            # Name every skipped/safe-mode/unconfirmed item: a quietly missing or
            # simplified file is the failure the user cannot see.
            skip_note = ""
            if self._skipped_fatal:
                skip_note = (
                    "\n\nSKIPPED — these crashed Revit on an earlier run:\n  {}\n\n"
                    "The crash is inside Revit's own renderer (faceting of the "
                    "views on these sheets), not BatchOut. Fix the offending "
                    "geometry, export them from Revit 2026, or delete\n{}\nto "
                    "let BatchOut try them again.".format(
                        "\n  ".join(self._skipped_fatal), self._crash_history_file))

            safe_note = ""
            if self._safe_applied:
                safe_note = (
                    "\n\nSAFE MODE was used for:\n  {}\n\nThe model itself was "
                    "left untouched (rolled back). Review these files before "
                    "issuing them.".format("\n  ".join(self._safe_applied)))

            fail_note = ""
            if self._failed_items:
                shown = self._failed_items[:20]
                fail_note = "\n\nEXPORT NOT CONFIRMED for:\n  {}".format("\n  ".join(shown))
                if len(self._failed_items) > len(shown):
                    fail_note += "\n  ...and {} more".format(
                        len(self._failed_items) - len(shown))
                fail_note += ("\n\nFull errors are in the pyRevit output log. Some "
                              "view/format pairs simply cannot be exported by Revit "
                              "(e.g. a schedule to DWG) and will always fail here.")

            if forms.alert("{}{}{}{}{}\n\nDo you want to open the output folder?".format(
                              summary, fail_note, safe_note, skip_note, report_note),
                          title=outcome,
                          yes=True, no=True):
                os.startfile(output_folder)

        except Exception as ex:
            logger.error("Export failed: {}".format(ex))
            self._end_live_progress()
            report_note = ""
            if getattr(self, '_export_started_at', None) and self.output_folder.Text:
                report_note = self._write_export_report(self.output_folder.Text)
            forms.alert("Export stopped: {}\n\n{} earlier output(s) were confirmed. "
                        "Existing output files were kept; review the output folder before retrying.{}".format(
                            ex, total_exported, report_note), title="Export Error")
            self.status_text.Text = "Export failed"
            self.next_button.IsEnabled = True
            self.back_button.IsEnabled = True
        finally:
            # Unfreeze (frozen in go_next) on every exit path, early returns too
            try:
                self._end_live_progress()
            finally:
                self._export_running = False
                self.IsEnabled = True

    def _write_export_report(self, output_folder):
        """Write this run's report into the output folder; returns a line for the summary.

        Never raises: a report problem must not turn a finished export into "Export stopped".
        """
        try:
            model = self.doc.Title
        except Exception:
            model = ""
        try:
            if not self.write_report.IsChecked or not self.export_items:
                return ""
            fmt = self._selected_report_format()
            started = getattr(self, '_export_started_at', None)
            when = datetime.fromtimestamp(started) if started else datetime.now()
            rows = batchout_report.build_rows(
                self.export_items, getattr(self, '_failure_reasons', {}), model, when)
            path = batchout_report.write_report(
                batchout_report.report_path(output_folder, fmt, when),
                rows, self._report_columns, fmt)
        except Exception as ex:
            logger.error("Export report not written: {}".format(ex))
            return ("\n\nThe export report was not written ({}). Check that the output "
                    "folder is writable; the exported files are not affected.".format(ex))
        logger.info("Export report: {}".format(path))
        return "\n\nReport ({} rows): {}".format(len(rows), path)

    def export_to_dwg(self, items, output_folder):
        """Export items (sheets or views) to DWG."""
        try:
            self._sync_item_labels(items, "DWG")

            selected_setup = None
            selected_setup_name = None
            if self.cad_export_setup.SelectedIndex > 0:
                selected_item = self.cad_export_setup.SelectedItem
                if hasattr(selected_item, 'Tag') and selected_item.Tag:
                    selected_setup = selected_item.Tag
                    selected_setup_name = selected_item.Content

            # GetPredefinedOptions is the only API that loads a named setup's
            # layers/colours/line weights into an options object.
            if selected_setup:
                try:
                    dwg_options = DWGExportOptions.GetPredefinedOptions(
                        self.doc, selected_setup.Name)
                except Exception as setup_ex:
                    logger.warning("Could not apply export setup '{}': {}".format(
                        selected_setup_name, setup_ex))
                    dwg_options = DWGExportOptions()
            else:
                dwg_options = DWGExportOptions()

            # "Match Revit display" (default on): colours, overrides and linetype
            # scale as the view shows them (an IndexColors + ByLayer setup paints
            # black linework in layer colours). Off: the setup exactly as saved.
            if self.cad_match_revit_display.IsChecked:
                applied = apply_revit_display_fidelity(
                    dwg_options, for_sheets=(self.selection_mode == "sheets"))
                logger.debug("DWG match Revit display: {}".format(", ".join(applied)))

            # Overrides the setup's AutoCAD version
            dwg_version_index = self.dwg_version.SelectedIndex
            if dwg_version_index == 0:
                dwg_options.FileVersion = ACADVersion.R2013
            elif dwg_version_index == 1:
                dwg_options.FileVersion = ACADVersion.R2010
            else:
                dwg_options.FileVersion = ACADVersion.R2007

            # ExportingAreas availability varies by version
            export_views_on_sheets = self.cad_export_views_on_sheets.IsChecked
            try:
                if hasattr(dwg_options, 'ExportingAreas') and hasattr(DB, 'ExportingAreas'):
                    if export_views_on_sheets:
                        dwg_options.ExportingAreas = DB.ExportingAreas.ExportViewsOnSheets
                    else:
                        dwg_options.ExportingAreas = DB.ExportingAreas.DontExportViewsOnSheets
            except Exception as ex:
                logger.debug("ExportingAreas not supported in Revit {}: {}".format(REVIT_VERSION, ex))

            export_links_as_external = self.cad_export_links_as_external.IsChecked
            try:
                if hasattr(dwg_options, 'MergedViews'):
                    dwg_options.MergedViews = not export_links_as_external
            except Exception as ex:
                logger.debug("MergedViews not supported in Revit {}: {}".format(REVIT_VERSION, ex))

            skip_images = bool(self.cad_skip_image_references.IsChecked)
            removed_refs = 0

            exported_count = 0

            for item in items:
                safe_group = None
                element_name = str(getattr(item, 'SheetNumber', 'Unknown'))
                try:
                    element, element_name = self._element_of(item)
                    if element is None:
                        self._note_failure(element_name, "DWG",
                                           "row is neither a sheet nor a view")
                        continue

                    # Known fatal: safe mode first; without it still export (queued
                    # last) — an unannounced missing file is worse.
                    safe_level = None
                    if self._is_known_fatal(item.SheetNumber, "DWG"):
                        if self._skip_fatal:
                            self._note_skipped_fatal(item.SheetNumber, "DWG")
                            continue
                        if self._safe_mode and not self._safe_ladder_exhausted(
                                item.SheetNumber, "DWG"):
                            safe_group, safe_level = self._begin_safe_mode(
                                element, self._next_safe_level(item.SheetNumber, "DWG"),
                                fmt="DWG")
                        if safe_group is None:
                            logger.warning(
                                "'{}' crashed Revit's DWG exporter before and is "
                                "being exported without mitigation.".format(
                                    item.SheetNumber))

                    self.update_export_item_progress(item.SheetNumber, "DWG", 50, "Exporting...")
                    self._update_progress(
                        int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                        "Exporting {} to DWG...".format(element_name)
                    )

                    filename = item.CustomFilename or self.get_export_filename(item)

                    if filename.lower().endswith('.dwg'):
                        filename = filename[:-4]

                    # Unique per run, or two views sharing a name overwrite each other
                    filename = self._reserve_filename(output_folder, filename, ".dwg")

                    # mtime now, so a stale file from an aborted run is not taken as fresh
                    expected_file = os.path.join(output_folder, filename + ".dwg")
                    prev_mtime = os.path.getmtime(expected_file) if os.path.exists(expected_file) else None
                    started_at = time.time()

                    # Export(folder, name, ICollection<ElementId>, DWGExportOptions)
                    view_ids = List[DB.ElementId]()
                    view_ids.Add(element.Id)

                    folder_before = None
                    if skip_images:
                        safe_group = self._begin_hide_images(element, safe_group)
                        folder_before = batchout_refs.snapshot(output_folder)

                    # Breadcrumb: DWG faceting can hard-crash Revit (uncatchable);
                    # the marker names the sheet for the next session.
                    self._write_crash_marker(item.SheetNumber, "DWG",
                                             safe_level=safe_level)

                    if self.api_adapter:
                        self.api_adapter.export_dwg(output_folder, filename, view_ids, dwg_options)
                    else:
                        if REVIT_VERSION >= 2022:
                            self.doc.Export(output_folder, filename, view_ids, dwg_options)
                        else:
                            self.doc.Export(output_folder, filename, view_ids, dwg_options)

                    self._clear_crash_marker()

                    if folder_before is not None:
                        # Images the hide could not reach (e.g. inside a title
                        # block family) still get written: delete them.
                        deleted, failed = batchout_refs.delete_files(
                            batchout_refs.written_reference_files(
                                output_folder, folder_before, keep=[expected_file]))
                        removed_refs += len(deleted)
                        for path in failed:
                            logger.warning("Could not delete image reference file: {}".format(path))

                    if self._verify_export(expected_file, prev_mtime, started_at, ".dwg"):
                        exported_count += 1
                        self._overall_counter += 1
                        self._update_progress(
                            int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                            "Exported {} to DWG".format(element_name)
                        )
                        self.update_export_item_progress(item.SheetNumber, "DWG", 100)
                    else:
                        self._note_failure(item.SheetNumber, "DWG",
                                           "Revit wrote no DWG file")

                except Exception as ex:
                    logger.error("Error exporting {} to DWG: {}".format(element_name, ex))
                    self._note_failure(str(getattr(item, 'SheetNumber', element_name)),
                                       "DWG", ex)
                finally:
                    # Safe mode must NEVER leave the model modified.
                    self._end_safe_mode(safe_group)

            if removed_refs:
                logger.info("DWG export: deleted {} image reference file(s).".format(removed_refs))
            return exported_count

        except Exception as ex:
            logger.error("DWG export failed: {}".format(ex))
            return 0

    def export_to_dgn(self, items, output_folder):
        """Export items (sheets or views) to DGN (MicroStation)."""
        try:
            self._sync_item_labels(items, "DGN")

            dgn_options = DGNExportOptions()

            exported_count = 0

            for item in items:
                element_name = str(getattr(item, 'SheetNumber', 'Unknown'))
                try:
                    element, element_name = self._element_of(item)
                    if element is None:
                        self._note_failure(element_name, "DGN",
                                           "row is neither a sheet nor a view")
                        continue

                    if self._is_known_fatal(item.SheetNumber, "DGN"):
                        if self._skip_fatal:
                            self._note_skipped_fatal(item.SheetNumber, "DGN")
                            continue
                        logger.warning(
                            "'{}' crashed Revit's DGN exporter before — exporting "
                            "anyway (queued last).".format(item.SheetNumber))

                    self.update_export_item_progress(item.SheetNumber, "DGN", 50, "Exporting...")
                    self._update_progress(
                        int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                        "Exporting {} to DGN...".format(element_name)
                    )

                    filename = item.CustomFilename or self.get_export_filename(item)
                    if filename.lower().endswith('.dgn'):
                        filename = filename[:-4]
                    filename = self._reserve_filename(output_folder, filename, ".dgn")

                    view_ids = List[DB.ElementId]()
                    view_ids.Add(element.Id)

                    expected_file = os.path.join(output_folder, filename + ".dgn")
                    prev_mtime = os.path.getmtime(expected_file) if os.path.exists(expected_file) else None
                    started_at = time.time()

                    self._write_crash_marker(item.SheetNumber, "DGN")
                    self.doc.Export(output_folder, filename, view_ids, dgn_options)
                    self._clear_crash_marker()

                    if self._verify_export(expected_file, prev_mtime, started_at, ".dgn"):
                        exported_count += 1
                        self._overall_counter += 1
                        self._update_progress(
                            int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                            "Exported {} to DGN".format(element_name)
                        )
                        self.update_export_item_progress(item.SheetNumber, "DGN", 100)
                    else:
                        self._note_failure(item.SheetNumber, "DGN",
                                           "Revit wrote no DGN file")

                except Exception as ex:
                    logger.error("Error exporting {} to DGN: {}".format(element_name, ex))
                    self._note_failure(str(getattr(item, 'SheetNumber', element_name)),
                                       "DGN", ex)

            return exported_count

        except Exception as ex:
            logger.error("DGN export failed: {}".format(ex))
            return 0

    def export_to_pdf(self, items, output_folder):
        """Export items (sheets or views) to PDF (combined or one file per item)."""
        try:
            import time
            import glob

            self._sync_item_labels(items, "PDF")

            if self._pdf_engine() == "PDF24":
                return self._export_pdf_with_pdf24(items, output_folder)

            combine_pdf = self.combine_pdf.IsChecked

            exported_count = 0

            if combine_pdf:
                included_items = []
                try:
                    # ONE native call: a fatal sheet kills the whole job, but pages
                    # are only dropped if the user chose to skip known-fatal ones.
                    element_ids = List[DB.ElementId]()
                    for item in items:
                        number = getattr(item, 'SheetNumber', None)
                        if number and self._skip_fatal and self._is_known_fatal(number, "PDF"):
                            self._note_skipped_fatal(number, "PDF")
                            continue
                        element, label = self._element_of(item)
                        if element is None:
                            self._note_failure(label, "PDF",
                                               "row is neither a sheet nor a view")
                            continue
                        element_ids.Add(element.Id)
                        included_items.append(item)

                    if element_ids.Count == 0:
                        logger.warning("Combined PDF skipped: nothing left to export.")
                        return 0

                    # Only rows included in the native call can become successful.
                    for item in included_items:
                        self.update_export_item_progress(item.SheetNumber, "PDF", 50, "Exporting...")
                    self._update_progress(
                        int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                        "Exporting combined PDF with {} items...".format(len(included_items)))
                    first_element, first_name = self._element_of(included_items[0])
                    last_element, last_name = self._element_of(included_items[-1])
                    if not hasattr(included_items[0], 'Sheet'):
                        first_name = first_name[:20]
                    if not hasattr(included_items[-1], 'Sheet'):
                        last_name = last_name[:20]
                    filename = "{}-{}_Combined".format(first_name, last_name)
                    filename = self._reserve_filename(output_folder, filename, ".pdf")
                    expected_file = os.path.join(output_folder, filename + ".pdf")
                    before_output = self._native_output_stamp(expected_file)

                    pdf_options = PDFExportOptions()
                    pdf_options.Combine = True
                    pdf_options.FileName = filename

                    # Keep linework under transparent filled regions from dropping
                    self._apply_pdf_quality(pdf_options)

                    if self.api_adapter:
                        pdf_options = self.api_adapter.configure_pdf_options(
                            pdf_options,
                            hide_ref_planes=self.pdf_hide_ref_planes.IsChecked,
                            hide_scope_boxes=self.pdf_hide_scope_boxes.IsChecked,
                            hide_crop_boundaries=self.pdf_hide_crop_boundaries.IsChecked,
                            hide_unreferenced_tags=self.pdf_hide_unreferenced_tags.IsChecked
                        )
                    else:
                        try:
                            if self.pdf_hide_ref_planes.IsChecked:
                                pdf_options.HideReferencePlane = True
                        except Exception as ex:
                            logger.debug("HideReferencePlane not supported in Revit {}: {}".format(REVIT_VERSION, ex))

                        try:
                            if self.pdf_hide_scope_boxes.IsChecked:
                                pdf_options.HideScopeBoxes = True
                        except Exception as ex:
                            logger.debug("HideScopeBoxes not supported in Revit {}: {}".format(REVIT_VERSION, ex))

                        try:
                            if self.pdf_hide_crop_boundaries.IsChecked:
                                pdf_options.HideCropBoundaries = True
                        except Exception as ex:
                            logger.debug("HideCropBoundaries not supported in Revit {}: {}".format(REVIT_VERSION, ex))

                        try:
                            if self.pdf_hide_unreferenced_tags.IsChecked:
                                pdf_options.HideUnreferencedViewTags = True
                        except Exception as ex:
                            logger.debug("HideUnreferencedViewTags not supported in Revit {}: {}".format(REVIT_VERSION, ex))

                    is_auto = self.pdf_auto_detect_size.IsChecked
                    manual_size = self.pdf_paper_size.SelectedItem.Content if self.pdf_paper_size.SelectedItem else "Use Sheet Size"
                    
                    if not is_auto and manual_size != "Use Sheet Size":
                        try:
                            pdf_options.PaperFormat = self.get_paper_format_enum(manual_size)
                        except Exception as ex:
                            logger.debug("Failed to set combined PDF paper format: {}".format(ex))
                            
                        try:
                            pdf_options.PaperOrientation = DB.PageOrientationType.Landscape if self.pdf_landscape.IsChecked else DB.PageOrientationType.Portrait
                        except Exception as ex:
                            logger.debug("Failed to set combined PDF orientation: {}".format(ex))
                            
                        # Zoom only when PaperFormat is not Default
                        try:
                            if self.pdf_fit_to_page.IsChecked:
                                pdf_options.ZoomType = DB.ZoomType.FitToPage
                            else:
                                pdf_options.ZoomType = DB.ZoomType.Zoom
                                pdf_options.ZoomPercentage = 100
                        except Exception as ex:
                            logger.debug("Failed to set combined PDF zoom: {}".format(ex))
                    else:
                        # Default format: Revit uses the sheet size
                        try:
                            pdf_options.PaperFormat = DB.ExportPaperFormat.Default
                        except Exception as ex:
                            logger.debug("Failed to set combined PDF paper format to Default: {}".format(ex))

                    # Breadcrumb names the job, not the sheet (uncheck Combine to pinpoint).
                    self._write_crash_marker(filename, "PDF-combined",
                                             self._raster_level_for(None) or "Presentation")

                    if self.api_adapter:
                        export_result = self.api_adapter.export_pdf(output_folder, filename, element_ids, pdf_options)
                    else:
                        # Export(folder, IList<ElementId>, options): no name argument,
                        # the filename comes from PDFExportOptions.FileName
                        export_result = self.doc.Export(output_folder, element_ids, pdf_options)

                    self._clear_crash_marker()
                    if not export_result:
                        raise RuntimeError(
                            "Revit reported an incomplete combined PDF export. Any output file was kept.")

                    if self._confirm_native_output(expected_file, before_output):
                        exported_count = 1
                        self._overall_counter += len(included_items)
                        self._update_progress(
                            int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                            "PDF combined: {} sheets exported".format(len(included_items))
                        )
                        for item in included_items:
                            self.update_export_item_progress(item.SheetNumber, "PDF", 100)
                    else:
                        # One call for every page: no file means none exported
                        for item in included_items:
                            self._note_failure(getattr(item, 'SheetNumber', 'Unknown'),
                                               "PDF", "combined PDF produced no new or updated nonempty output")

                except Exception as ex:
                    logger.error("Error exporting combined PDF: {}".format(ex))
                    for item in included_items:
                        self._note_failure(getattr(item, 'SheetNumber', 'Unknown'),
                                           "PDF", "combined PDF failed: {}".format(ex))

            else:
                sheets_since_gc = 0
                for item in items:
                    safe_group = None
                    element_name = str(getattr(item, 'SheetNumber', 'Unknown'))
                    try:
                        element, element_name = self._element_of(item)
                        if element is None:
                            self._note_failure(element_name, "PDF",
                                               "row is neither a sheet nor a view")
                            continue

                        # Known fatal: safe mode while a rung is left, else still
                        # exported (queued last) rather than silently dropped.
                        safe_level = None
                        if self._is_known_fatal(item.SheetNumber, "PDF"):
                            if self._skip_fatal:
                                self._note_skipped_fatal(item.SheetNumber, "PDF")
                                continue
                            if self._safe_mode and not self._safe_ladder_exhausted(
                                    item.SheetNumber, "PDF"):
                                safe_group, safe_level = self._begin_safe_mode(
                                    element, self._next_safe_level(item.SheetNumber, "PDF"),
                                    fmt="PDF")
                            if safe_group is None:
                                logger.warning(
                                    "'{}' crashed Revit's PDF exporter before and "
                                    "is being exported without mitigation.".format(
                                        item.SheetNumber))

                        self.update_export_item_progress(item.SheetNumber, "PDF", 50, "Exporting...")
                        self._update_progress(
                            int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                            "Exporting {} to PDF...".format(element_name)
                        )

                        filename = item.CustomFilename or self.get_export_filename(item)

                        if filename.lower().endswith('.pdf'):
                            filename = filename[:-4]

                        # Unique per run: plans may share a name and overwrite each other
                        filename = self._reserve_filename(output_folder, filename, ".pdf")

                        # mtime now (one stat, no folder glob) to detect a fresh write
                        expected_file = os.path.join(output_folder, filename + ".pdf")
                        prev_mtime = os.path.getmtime(expected_file) if os.path.exists(expected_file) else None
                        started_at = time.time()

                        pdf_options = PDFExportOptions()
                        # Combine=True even for one sheet: with False Revit ignores FileName
                        pdf_options.Combine = True
                        pdf_options.FileName = filename

                        # Full quality (keeps linework under transparent filled regions);
                        # only a sheet that crashed the rasteriser starts one rung lower.
                        raster_level = self._raster_level_for(item.SheetNumber)
                        self._apply_pdf_quality(pdf_options, raster_level)

                        if self.api_adapter:
                            pdf_options = self.api_adapter.configure_pdf_options(
                                pdf_options,
                                hide_ref_planes=self.pdf_hide_ref_planes.IsChecked,
                                hide_scope_boxes=self.pdf_hide_scope_boxes.IsChecked,
                                hide_crop_boundaries=self.pdf_hide_crop_boundaries.IsChecked,
                                hide_unreferenced_tags=self.pdf_hide_unreferenced_tags.IsChecked
                            )
                        else:
                            try:
                                if self.pdf_hide_ref_planes.IsChecked:
                                    pdf_options.HideReferencePlane = True
                            except Exception as ex:
                                logger.debug("HideReferencePlane not supported in Revit {}: {}".format(REVIT_VERSION, ex))

                            try:
                                if self.pdf_hide_scope_boxes.IsChecked:
                                    pdf_options.HideScopeBoxes = True
                            except Exception as ex:
                                logger.debug("HideScopeBoxes not supported in Revit {}: {}".format(REVIT_VERSION, ex))

                            try:
                                if self.pdf_hide_crop_boundaries.IsChecked:
                                    pdf_options.HideCropBoundaries = True
                            except Exception as ex:
                                logger.debug("HideCropBoundaries not supported in Revit {}: {}".format(REVIT_VERSION, ex))

                            try:
                                if self.pdf_hide_unreferenced_tags.IsChecked:
                                    pdf_options.HideUnreferencedViewTags = True
                            except Exception as ex:
                                logger.debug("HideUnreferencedViewTags not supported in Revit {}: {}".format(REVIT_VERSION, ex))

                        is_auto = self.pdf_auto_detect_size.IsChecked
                        manual_size = self.pdf_paper_size.SelectedItem.Content if self.pdf_paper_size.SelectedItem else "Use Sheet Size"
                        
                        sheet_size = None
                        sheet_orientation = None
                        
                        if (is_auto or manual_size == "Use Sheet Size") and hasattr(item, 'Sheet'):
                            detected_size, detected_orientation = self.get_sheet_paper_size_and_orientation(item.Sheet)
                            sheet_size = detected_size
                            sheet_orientation = detected_orientation
                        else:
                            sheet_size = manual_size
                            sheet_orientation = "Landscape" if self.pdf_landscape.IsChecked else "Portrait"
                            
                        if sheet_size and sheet_size != "Use Sheet Size":
                            try:
                                pdf_options.PaperFormat = self.get_paper_format_enum(sheet_size)
                            except Exception as ex:
                                logger.debug("Failed to set individual PDF paper format: {}".format(ex))
                                
                            try:
                                pdf_options.PaperOrientation = DB.PageOrientationType.Landscape if sheet_orientation == "Landscape" else DB.PageOrientationType.Portrait
                            except Exception as ex:
                                logger.debug("Failed to set individual PDF orientation: {}".format(ex))
                                
                            try:
                                if self.pdf_fit_to_page.IsChecked:
                                    pdf_options.ZoomType = DB.ZoomType.FitToPage
                                else:
                                    pdf_options.ZoomType = DB.ZoomType.Zoom
                                    pdf_options.ZoomPercentage = 100
                            except Exception as ex:
                                logger.debug("Failed to set individual PDF zoom: {}".format(ex))
                        else:
                            try:
                                pdf_options.PaperFormat = DB.ExportPaperFormat.Default
                            except Exception as ex:
                                logger.debug("Failed to set individual PDF paper format to Default: {}".format(ex))

                        element_ids = List[DB.ElementId]()
                        element_ids.Add(element.Id)

                        # Breadcrumb: names the sheet if the native call hard-crashes Revit
                        self._write_crash_marker(item.SheetNumber, "PDF",
                                                 raster_level or "Presentation",
                                                 safe_level=safe_level)

                        if self.api_adapter:
                            export_result = self.api_adapter.export_pdf(output_folder, filename, element_ids, pdf_options)
                        else:
                            # no name argument: see the combined branch
                            export_result = self.doc.Export(output_folder, element_ids, pdf_options)

                        # Survived: drop the breadcrumb so it is not taken for a crash
                        self._clear_crash_marker()

                        if not export_result:
                            raise RuntimeError("Revit reported PDF export failure; any output file was kept for review")

                        # Polls for the file (no fixed sleep or glob: fast on OneDrive)
                        if self._verify_export(expected_file, prev_mtime, started_at, ".pdf"):
                            exported_count += 1
                            self._overall_counter += 1
                            self._update_progress(
                                int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                                "Exported {} to PDF".format(element_name)
                            )
                            self.update_export_item_progress(item.SheetNumber, "PDF", 100)
                        else:
                            self._note_failure(item.SheetNumber, "PDF",
                                               "Revit wrote no PDF file")

                        # Reclaim raster buffers every 10 sheets: a full GC on Revit's
                        # heap is not free, but memory growth must be capped.
                        sheets_since_gc += 1
                        if sheets_since_gc >= 10:
                            sheets_since_gc = 0
                            GC.Collect()
                            GC.WaitForPendingFinalizers()

                    except Exception as ex:
                        logger.error("Error exporting {} to PDF: {}".format(element_name, ex))
                        self._note_failure(str(getattr(item, 'SheetNumber', element_name)),
                                           "PDF", ex)
                    finally:
                        # Safe mode must NEVER leave the model modified.
                        self._end_safe_mode(safe_group)

            return exported_count

        except Exception as ex:
            logger.error("PDF export failed: {}".format(ex))
            return 0

    PDF24_PAGE_TIMEOUT = 120.0      # s per sheet; PDF24 may start its tray app on the first
    PDF24_MAX_MISSES = 2            # sheets in a row with no file -> stop, PDF24 is not saving

    def _export_pdf_with_pdf24(self, items, output_folder):
        """PDF through the PDF24 Creator printer (PrintManager + PDF24 auto-save).

        One print job per item; PDF24 saves it under the reserved name. Combine
        joins the pages with pdf24-DocTool. Print settings are changed InSession
        inside a TransactionGroup that is rolled back, and the user's printer and
        PDF24 settings are put back. NOT verified in Revit + PDF24 yet.
        """
        import shutil
        import tempfile

        printer = batchout_pdf24.find_printer()
        session = batchout_pdf24.AutoSave()
        problem = None
        if not printer:
            problem = ("PDF24 printer not found — install PDF24 Creator ({}) or set "
                       "PDF Engine to Revit (built-in)".format(batchout_pdf24.INSTALL_URL))
        else:
            blocked = session.blocked_reason()
            if blocked:
                problem = "PDF24 cannot be used: {} — set PDF Engine to Revit (built-in)".format(blocked)
        if problem:
            logger.error(problem)
            for item in items:
                self._note_failure(getattr(item, 'SheetNumber', 'Unknown'), "PDF", problem)
            return 0

        combine = bool(self.combine_pdf.IsChecked)
        doctool = batchout_pdf24.find_doctool() if combine else None
        if combine and not doctool:
            logger.warning("pdf24-DocTool.exe not found: Combine is off for this run, "
                           "one PDF per sheet is written instead.")
            combine = False
        work_folder = tempfile.mkdtemp(prefix="t3lab_pdf24_") if combine else output_folder

        pm = self.doc.PrintManager
        old_printer = None
        group = None
        pages = []          # (item, path) in print order
        exported_count = 0
        misses = 0
        old_print_to_file = None
        try:
            old_printer = pm.PrinterName
            old_print_to_file = bool(pm.PrintToFile)
        except Exception:
            pass
        try:
            group = DB.TransactionGroup(self.doc, "BatchOut PDF24 print setup")
            group.Start()
            session.start(work_folder)
            pm.SelectNewPrintDriver(printer)
            # Revit can class PDF24 as a virtual printer, and then refuses
            # PrintToFile = False (owner, 2026-11-17): print to a file and
            # turn PostScript into PDF with PDF24's Ghostscript.
            virtual = self._pdf24_is_virtual(pm)
            pm.PrintToFile = virtual
            pm.Apply()
            gs = batchout_pdf24.find_ghostscript() if virtual else None

            for index, item in enumerate(items):
                label = str(getattr(item, 'SheetNumber', 'Unknown'))
                if misses >= self.PDF24_MAX_MISSES:
                    self._note_failure(label, "PDF", "skipped: PDF24 saved nothing for the "
                                       "previous {} sheets".format(misses))
                    continue
                try:
                    element, element_name = self._element_of(item)
                    if element is None:
                        self._note_failure(label, "PDF", "row is neither a sheet nor a view")
                        continue
                    # A QUEUE_* state, or _end_live_progress never closes the row out.
                    self.update_export_item_progress(item.SheetNumber, "PDF", 50, QUEUE_EXPORTING)
                    self._update_progress(
                        int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                        "Printing {} to PDF24...".format(element_name))

                    if combine:
                        filename = "{:04d}".format(index + 1)
                    else:
                        filename = item.CustomFilename or self.get_export_filename(item)
                        if filename.lower().endswith('.pdf'):
                            filename = filename[:-4]
                        filename = self._reserve_filename(work_folder, filename, ".pdf")
                    expected_file = os.path.join(work_folder, filename + ".pdf")
                    prev_mtime = os.path.getmtime(expected_file) if os.path.exists(expected_file) else None

                    size, orientation = self._pdf24_paper_for(item)
                    prn = batchout_pdf24.prn_path(work_folder, filename) if virtual else None
                    if prn and os.path.exists(prn):
                        os.remove(prn)
                    trans = Transaction(self.doc, "PDF24 page setup")
                    trans.Start()
                    try:
                        note = self._pdf24_page_setup(pm, size, orientation, prn)
                        trans.Commit()
                    except Exception:
                        trans.RollBack()
                        raise
                    if note:
                        logger.warning("{}: {}".format(element_name, note))

                    session.set_filename(filename)
                    started_at = time.time()
                    if not pm.SubmitPrint(element):
                        raise RuntimeError("Revit did not send the print job to PDF24")

                    taken = getattr(self, '_used_filenames', set())
                    found = batchout_pdf24.wait_for_pdf(
                        expected_file, prev_mtime, started_at, self.PDF24_PAGE_TIMEOUT,
                        taken=[p for p in taken if p != expected_file.lower()]
                        + [page[1] for page in pages], extra=[prn] if prn else ())
                    if found is None:
                        misses += 1
                        self._note_failure(label, "PDF", "PDF24 saved no file within {:.0f} s"
                                           .format(self.PDF24_PAGE_TIMEOUT))
                        continue
                    found = self._pdf24_as_pdf(found, expected_file, gs)
                    misses = 0
                    if found.lower() != expected_file.lower():
                        logger.warning("PDF24 saved '{}' as '{}'; renamed.".format(
                            os.path.basename(expected_file), os.path.basename(found)))
                        if os.path.exists(expected_file):
                            os.remove(expected_file)
                        os.rename(found, expected_file)
                    pages.append((item, expected_file))
                    if not combine:
                        exported_count += 1
                        self._overall_counter += 1
                        self._last_output_file = expected_file
                        self.update_export_item_progress(item.SheetNumber, "PDF", 100)
                except Exception as ex:
                    logger.error("Error printing {} to PDF24: {}".format(label, ex))
                    self._note_failure(label, "PDF", ex)
        except Exception as ex:
            logger.error("PDF24 export failed: {}".format(ex))
            for item in items:
                if not any(item is p[0] for p in pages):
                    self._note_failure(getattr(item, 'SheetNumber', 'Unknown'), "PDF",
                                       "PDF24 export failed: {}".format(ex))
        finally:
            session.restore()
            if old_printer:
                try:
                    pm.SelectNewPrintDriver(old_printer)
                    if old_print_to_file is not None and not self._pdf24_is_virtual(pm):
                        pm.PrintToFile = old_print_to_file
                    pm.Apply()
                except Exception as ex:
                    logger.debug("Could not reselect printer '{}': {}".format(old_printer, ex))
            if group is not None:
                try:
                    group.RollBack()
                except Exception as ex:
                    logger.error("PDF24 print setup rollback failed: {}".format(ex))

        if combine:
            exported_count = self._pdf24_join(doctool, pages, items, output_folder)
            shutil.rmtree(work_folder, ignore_errors=True)
        return exported_count

    @staticmethod
    def _pdf24_is_virtual(pm):
        """True when Revit treats the selected printer as virtual
        (PrintManager.IsVirtual != None): it must then print to a file."""
        try:
            return '{}'.format(pm.IsVirtual) != 'None'
        except Exception:
            return False

    def _pdf24_as_pdf(self, path, expected_file, gs):
        """The printed file as a real PDF at `path` or converted; raises when it
        is neither PDF nor PostScript, or PostScript without Ghostscript."""
        import subprocess
        kind = batchout_pdf24.sniff(path)
        if kind == 'pdf':
            return path
        if kind != 'ps':
            raise RuntimeError("PDF24 printed a file that is neither PDF nor PostScript ({})"
                               .format(os.path.basename(path)))
        if not gs:
            raise RuntimeError("Revit printed PostScript to a file and PDF24's Ghostscript "
                               "(PDF24\\gs\\bin\\gswin64c.exe) was not found to turn it into PDF")
        if os.path.exists(expected_file) and path.lower() != expected_file.lower():
            os.remove(expected_file)
        target = expected_file if path.lower() != expected_file.lower() else expected_file + '.tmp.pdf'
        cmd = batchout_pdf24.ps_to_pdf_command(gs, path, target)
        subprocess.call(cmd, timeout=300, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        if batchout_pdf24.sniff(target) != 'pdf':
            raise RuntimeError("Ghostscript could not turn PDF24's PostScript into PDF")
        try:
            os.remove(path)
        except OSError:
            pass
        if target != expected_file:
            os.rename(target, expected_file)
        return expected_file

    def _pdf24_paper_for(self, item):
        """(paper size, orientation) for one item, same rules as the Revit engine."""
        is_auto = self.pdf_auto_detect_size.IsChecked
        manual = self.pdf_paper_size.SelectedItem.Content if self.pdf_paper_size.SelectedItem else "Use Sheet Size"
        if (is_auto or manual == "Use Sheet Size") and hasattr(item, 'Sheet'):
            return self.get_sheet_paper_size_and_orientation(item.Sheet)
        return manual, ("Landscape" if self.pdf_landscape.IsChecked else "Portrait")

    def _pdf24_page_setup(self, pm, size, orientation, print_file=None):
        """InSession print parameters for one page (open transaction); print_file:
        the print-to-file name for a virtual printer. Returns a note when PDF24
        has no such paper (the current paper is kept)."""
        setup = pm.PrintSetup
        try:
            setup.CurrentPrintSetting = setup.InSession
        except Exception as ex:
            logger.debug("InSession print setting: {}".format(ex))
        params = setup.CurrentPrintSetting.PrintParameters
        note = None
        if size and size != "Use Sheet Size":
            papers = list(pm.PaperSizes)
            match = batchout_pdf24.match_paper(size, [p.Name for p in papers])
            paper = next((p for p in papers if p.Name == match), None)
            if paper is not None:
                params.PaperSize = paper
            else:
                note = "PDF24 has no '{}' paper; the current paper size was used".format(size)
        settings = (
            ('PageOrientation', DB.PageOrientationType.Landscape if orientation == "Landscape"
             else DB.PageOrientationType.Portrait),
            ('PaperPlacement', DB.PaperPlacementType.Center),
            ('ZoomType', DB.ZoomType.FitToPage if self.pdf_fit_to_page.IsChecked else DB.ZoomType.Zoom),
            ('ColorDepth', DB.ColorDepthType.Color),
            ('RasterQuality', DB.RasterQualityType.High),
            ('HiddenLineViews', DB.HiddenLineViewsType.VectorProcessing),
            ('HideReforWorkPlanes', bool(self.pdf_hide_ref_planes.IsChecked)),
            ('HideScopeBoxes', bool(self.pdf_hide_scope_boxes.IsChecked)),
            ('HideCropBoundaries', bool(self.pdf_hide_crop_boundaries.IsChecked)),
            ('HideUnreferencedViewTags', bool(self.pdf_hide_unreferenced_tags.IsChecked)),
            ('ViewLinksinBlue', bool(self.pdf_view_links_blue.IsChecked)),
        )
        for name, value in settings:
            try:
                setattr(params, name, value)
            except Exception as ex:
                logger.debug("Print parameter {} not set: {}".format(name, ex))
        if not self.pdf_fit_to_page.IsChecked:
            try:
                params.Zoom = 100
            except Exception as ex:
                logger.debug("Print zoom not set: {}".format(ex))
        if print_file:
            pm.PrintToFileName = print_file
        pm.Apply()
        return note

    def _pdf24_join(self, doctool, pages, items, output_folder):
        """Join the printed pages into one PDF; 1 on success. On failure the pages
        are kept in the output folder, one file per sheet."""
        import shutil
        import subprocess
        if not pages:
            return 0
        first = self._element_of(pages[0][0])[1]
        last = self._element_of(pages[-1][0])[1]
        name = self._reserve_filename(output_folder, "{}-{}_Combined".format(first, last), ".pdf")
        target = os.path.join(output_folder, name + ".pdf")
        before = self._native_output_stamp(target)
        cmd = batchout_pdf24.join_command(doctool, [p[1] for p in pages], output_folder, name)
        self._update_progress(
            int(self._overall_counter * 100.0 / max(1, self._overall_total)),
            "Joining {} pages with PDF24...".format(len(pages)))
        try:
            subprocess.call(cmd, timeout=300, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
        except Exception as ex:
            logger.error("pdf24-DocTool join failed: {}".format(ex))
        if self._confirm_native_output(target, before):
            self._overall_counter += len(pages)
            for item, _ in pages:
                self.update_export_item_progress(item.SheetNumber, "PDF", 100)
            return 1
        logger.warning("PDF24 could not join the pages; one PDF per sheet was kept instead.")
        kept = 0
        for item, path in pages:
            base = item.CustomFilename or self.get_export_filename(item)
            if base.lower().endswith('.pdf'):
                base = base[:-4]
            dest = os.path.join(output_folder, self._reserve_filename(output_folder, base, ".pdf") + ".pdf")
            try:
                if os.path.exists(dest):
                    os.remove(dest)
                shutil.move(path, dest)     # temp folder may be on another drive
                kept += 1
                self._overall_counter += 1
                self.update_export_item_progress(item.SheetNumber, "PDF", 100)
            except Exception as ex:
                self._note_failure(item.SheetNumber, "PDF", "could not keep the printed page: {}".format(ex))
        return kept

    def export_to_dwf(self, items, output_folder):
        """Export items (sheets or views) to DWF."""
        try:
            self._sync_item_labels(items, "DWF")

            dwf_options = DWFExportOptions()

            exported_count = 0

            for item in items:
                element_name = str(getattr(item, 'SheetNumber', 'Unknown'))
                try:
                    element, element_name = self._element_of(item)
                    if element is None:
                        self._note_failure(element_name, "DWF",
                                           "row is neither a sheet nor a view")
                        continue

                    if self._is_known_fatal(item.SheetNumber, "DWF"):
                        if self._skip_fatal:
                            self._note_skipped_fatal(item.SheetNumber, "DWF")
                            continue
                        logger.warning(
                            "'{}' crashed Revit's DWF exporter before — exporting "
                            "anyway (queued last).".format(item.SheetNumber))

                    self.update_export_item_progress(item.SheetNumber, "DWF", 50, "Exporting...")
                    self._update_progress(
                        int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                        "Exporting {} to DWF...".format(element_name)
                    )

                    filename = item.CustomFilename or self.get_export_filename(item)

                    if filename.lower().endswith('.dwf'):
                        filename = filename[:-4]
                    filename = self._reserve_filename(output_folder, filename, ".dwf")

                    # Export(folder, name, ViewSet, DWFExportOptions)
                    view_set = DB.ViewSet()
                    view_set.Insert(element)

                    expected_file = os.path.join(output_folder, filename + ".dwf")
                    prev_mtime = os.path.getmtime(expected_file) if os.path.exists(expected_file) else None
                    started_at = time.time()

                    self._write_crash_marker(item.SheetNumber, "DWF")
                    self.doc.Export(output_folder, filename, view_set, dwf_options)
                    self._clear_crash_marker()

                    if self._verify_export(expected_file, prev_mtime, started_at, ".dwf"):
                        exported_count += 1
                        self._overall_counter += 1
                        self._update_progress(
                            int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                            "Exported {} to DWF".format(element_name)
                        )
                        self.update_export_item_progress(item.SheetNumber, "DWF", 100)
                    else:
                        self._note_failure(item.SheetNumber, "DWF",
                                           "Revit wrote no DWF file")

                except Exception as ex:
                    logger.error("Error exporting {} to DWF: {}".format(element_name, ex))
                    self._note_failure(str(getattr(item, 'SheetNumber', element_name)),
                                       "DWF", ex)

            return exported_count

        except Exception as ex:
            logger.error("DWF export failed: {}".format(ex))
            return 0

    def export_to_nwd(self, items, output_folder):
        """Export items (sheets or views) to Navisworks NWC."""
        if not HAS_NAVISWORKS:
            return 0

        try:
            self._sync_item_labels(items, "NWC")

            nwd_options = NavisworksExportOptions()

            exported_count = 0

            for item in items:
                element_name = str(getattr(item, 'SheetNumber', 'Unknown'))
                try:
                    element, element_name = self._element_of(item)
                    if element is None:
                        self._note_failure(element_name, "NWC",
                                           "row is neither a sheet nor a view")
                        continue

                    if self._is_known_fatal(item.SheetNumber, "NWC"):
                        if self._skip_fatal:
                            self._note_skipped_fatal(item.SheetNumber, "NWC")
                            continue
                        logger.warning(
                            "'{}' crashed Revit's NWC exporter before — exporting "
                            "anyway (queued last).".format(item.SheetNumber))

                    self.update_export_item_progress(item.SheetNumber, "NWC", 50, "Exporting...")
                    self._update_progress(
                        int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                        "Exporting {} to NWC...".format(element_name)
                    )

                    filename = item.CustomFilename or self.get_export_filename(item)
                    if filename.lower().endswith('.nwc'):
                        filename = filename[:-4]
                    filename = self._reserve_filename(output_folder, filename, ".nwc")

                    expected_file = os.path.join(output_folder, filename + ".nwc")
                    prev_mtime = os.path.getmtime(expected_file) if os.path.exists(expected_file) else None
                    started_at = time.time()

                    nwd_options.ExportScope = DB.NavisworksExportScope.View
                    nwd_options.ViewId = element.Id

                    self._write_crash_marker(item.SheetNumber, "NWC")
                    self.doc.Export(output_folder, filename, nwd_options)
                    self._clear_crash_marker()

                    if self._verify_export(expected_file, prev_mtime, started_at, ".nwc"):
                        exported_count += 1
                        self._overall_counter += 1
                        self._update_progress(
                            int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                            "Exported {} to NWC".format(element_name)
                        )
                        self.update_export_item_progress(item.SheetNumber, "NWC", 100)
                    else:
                        self._note_failure(item.SheetNumber, "NWC",
                                           "Revit wrote no NWC file")

                except Exception as ex:
                    logger.error("Error exporting {} to NWC: {}".format(element_name, ex))
                    self._note_failure(str(getattr(item, 'SheetNumber', element_name)),
                                       "NWC", ex)

            return exported_count

        except Exception as ex:
            logger.error("NWC export failed: {}".format(ex))
            return 0

    def export_to_ifc(self, items, output_folder):
        """Export the whole model to one IFC file."""
        if not HAS_IFC:
            return 0

        try:
            self._sync_item_labels(items, "IFC")

            ifc_options = IFCExportOptions()
            ifc_ver_map = {0: IFCVersion.IFC2x2, 1: IFCVersion.IFC2x3, 2: IFCVersion.IFC4}
            ifc_ver_index = self.ifc_version.SelectedIndex if hasattr(self, 'ifc_version') else 1
            ifc_options.FileVersion = ifc_ver_map.get(ifc_ver_index, IFCVersion.IFC2x3)
            ifc_options.WallAndColumnSplitting = True

            exported_count = 0

            # Unlike the other formats, IFC export needs a transaction
            if len(items) > 0:
                trans = None
                try:
                    for item in items:
                        self.update_export_item_progress(item.SheetNumber, "IFC", 50, "Exporting...")
                    self._update_progress(
                        int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                        "Exporting entire model to IFC..."
                    )

                    # Named from first + last item, like the combined PDF
                    if len(items) > 1:
                        first_item = items[0]
                        last_item = items[-1]
                        if hasattr(first_item, 'Sheet'):
                            first_name = first_item.Sheet.SheetNumber
                            last_name = last_item.Sheet.SheetNumber
                        else:
                            first_name = first_item.View.Name[:20]
                            last_name = last_item.View.Name[:20]
                        filename = "{}-{}_Model_IFC".format(first_name, last_name)
                    elif len(items) == 1:
                        filename = self.get_export_filename(items[0]) + "_IFC"
                    else:
                        filename = "Model_IFC_Export"

                    if filename.lower().endswith('.ifc'):
                        filename = filename[:-4]

                    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
                    for char in invalid_chars:
                        filename = filename.replace(char, '_')
                    filename = filename.strip()
                    filename = self._reserve_filename(output_folder, filename, ".ifc")
                    expected_file = os.path.join(output_folder, filename + ".ifc")
                    before_output = self._native_output_stamp(expected_file)

                    # One call for the whole model: the breadcrumb records the job.
                    self._write_crash_marker(filename, "IFC")
                    trans = Transaction(self.doc, "T3Lab: Export IFC")
                    trans.Start()
                    options = trans.GetFailureHandlingOptions()
                    options.SetForcedModalHandling(True)
                    trans.SetFailureHandlingOptions(options)
                    exported = self.doc.Export(output_folder, filename, ifc_options)
                    self._clear_crash_marker()
                    if not exported:
                        raise RuntimeError("Revit reported that the IFC export failed")
                    status = trans.Commit()
                    if status != DB.TransactionStatus.Committed:
                        raise RuntimeError("IFC transaction was not committed: {}".format(status))
                    if not self._confirm_native_output(expected_file, before_output):
                        raise RuntimeError("No new or updated nonempty IFC output was confirmed")

                    exported_count = 1
                    self._overall_counter += len(items)
                    self._update_progress(
                        int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                        "IFC exported"
                    )
                    for item in items:
                        self.update_export_item_progress(item.SheetNumber, "IFC", 100)
                except Exception as ex:
                    if trans is not None and trans.GetStatus() == DB.TransactionStatus.Pending:
                        raise _PendingExportError(
                            "IFC export is awaiting Revit failure resolution. Later formats were stopped. "
                            "Resolve the Revit dialog before retrying; any output files were kept.")
                    if trans is not None and trans.GetStatus() == DB.TransactionStatus.Started:
                        trans.RollBack()
                    logger.error("Error exporting to IFC: {}".format(ex))
                    for item in items:
                        self._note_failure(getattr(item, 'SheetNumber', 'Unknown'),
                                           "IFC", "model IFC export failed: {}. Any output files were kept.".format(ex))

            return exported_count

        except _PendingExportError:
            raise
        except Exception as ex:
            logger.error("IFC export failed: {}".format(ex))
            return 0

    def export_to_images(self, items, output_folder):
        """Export items (sheets or views) to PNG/JPEG images."""
        try:
            import time
            import glob

            self._sync_item_labels(items, "IMG")

            exported_count = 0

            for item in items:
                element_name = str(getattr(item, 'SheetNumber', 'Unknown'))
                try:
                    element, element_name = self._element_of(item)
                    if element is None:
                        self._note_failure(element_name, "IMG",
                                           "row is neither a sheet nor a view")
                        continue

                    if self._is_known_fatal(item.SheetNumber, "IMG"):
                        if self._skip_fatal:
                            self._note_skipped_fatal(item.SheetNumber, "IMG")
                            continue
                        logger.warning(
                            "'{}' crashed Revit's image exporter before — exporting "
                            "anyway (queued last).".format(item.SheetNumber))

                    self.update_export_item_progress(item.SheetNumber, "IMG", 50, "Exporting...")
                    self._update_progress(
                        int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                        "Exporting {} to Image...".format(element_name)
                    )

                    filename = item.CustomFilename or self.get_export_filename(item)

                    if filename.lower().endswith(('.png', '.jpg')):
                        filename = filename.rsplit('.', 1)[0]

                    invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
                    for char in invalid_chars:
                        filename = filename.replace(char, '_')
                    filename = filename.strip()

                    use_fit_to_page = not (hasattr(self, 'img_zoom_to') and self.img_zoom_to.IsChecked)
                    fit_pixels = 1080
                    zoom_percent = 50
                    try:
                        if hasattr(self, 'img_fit_pixels'):
                            fit_pixels = int(self.img_fit_pixels.Text)
                    except:
                        pass
                    try:
                        if hasattr(self, 'img_zoom_percent'):
                            zoom_percent = int(self.img_zoom_percent.Text)
                    except:
                        pass
                    use_horizontal = not (hasattr(self, 'img_dir_vertical') and self.img_dir_vertical.IsChecked)
                    dpi_map = {
                        0: ImageResolution.DPI_72,
                        1: ImageResolution.DPI_96,
                        2: ImageResolution.DPI_150,
                        3: ImageResolution.DPI_300,
                        4: ImageResolution.DPI_600,
                    }
                    dpi_index = self.img_dpi.SelectedIndex if hasattr(self, 'img_dpi') else 2
                    img_resolution = dpi_map.get(dpi_index, ImageResolution.DPI_150)
                    shaded_idx = self.img_shaded_format.SelectedIndex if hasattr(self, 'img_shaded_format') else 0
                    nonshaded_idx = self.img_nonshaded_format.SelectedIndex if hasattr(self, 'img_nonshaded_format') else 0
                    shaded_fmt = ImageFileType.JPEGLossless if shaded_idx == 1 else ImageFileType.PNG
                    nonshaded_fmt = ImageFileType.JPEGLossless if nonshaded_idx == 1 else ImageFileType.PNG

                    # Revit picks the shaded or non-shaded type per view: watch both
                    # extensions (a "*.png"-only glob reported every JPEG as missing).
                    img_exts = set()
                    for fmt in (shaded_fmt, nonshaded_fmt):
                        img_exts.add('.jpg' if fmt == ImageFileType.JPEGLossless else '.png')

                    reserved = filename
                    for ext in sorted(img_exts):
                        reserved = self._reserve_filename(output_folder, reserved, ext)
                    filename = reserved

                    existing_images = set()
                    for ext in img_exts:
                        existing_images |= set(glob.glob(
                            os.path.join(output_folder, "*" + ext)))

                    img_options = ImageExportOptions()
                    if use_fit_to_page:
                        img_options.ZoomType = DB.ZoomFitType.FitToPage
                        img_options.PixelSize = fit_pixels
                        img_options.FitDirection = DB.FitDirectionType.Horizontal if use_horizontal else DB.FitDirectionType.Vertical
                    else:
                        img_options.ZoomType = DB.ZoomFitType.Zoom
                        img_options.Zoom = zoom_percent
                    img_options.ImageResolution = img_resolution
                    img_options.FilePath = os.path.join(output_folder, filename)
                    img_options.HLRandWFViewsFileType = nonshaded_fmt
                    img_options.ShadowViewsFileType = shaded_fmt
                    img_options.ExportRange = DB.ExportRange.SetOfViews

                    view_ids = List[DB.ElementId]()
                    view_ids.Add(element.Id)
                    img_options.SetViewsAndSheets(view_ids)

                    self._write_crash_marker(item.SheetNumber, "IMG")
                    self.doc.ExportImage(img_options)
                    self._clear_crash_marker()

                    time.sleep(0.3)

                    current_images = set()
                    for ext in img_exts:
                        current_images |= set(glob.glob(
                            os.path.join(output_folder, "*" + ext)))
                    new_images = current_images - existing_images

                    # Revit appends " - Sheet - A101 - ..." to the name: rename back,
                    # keeping the extension Revit chose.
                    expected_file = None
                    if new_images:
                        actual_file = sorted(new_images)[0]
                        actual_ext = os.path.splitext(actual_file)[1].lower()
                        expected_file = os.path.join(output_folder, filename + actual_ext)
                        if os.path.normcase(actual_file) != os.path.normcase(expected_file):
                            try:
                                # Name reserved by this run: overwrite intended
                                if os.path.exists(expected_file):
                                    os.remove(expected_file)
                                os.rename(actual_file, expected_file)
                            except Exception as rename_ex:
                                logger.warning("Could not rename {} to {}: {}".format(
                                    os.path.basename(actual_file),
                                    os.path.basename(expected_file),
                                    rename_ex
                                ))
                                expected_file = actual_file

                    if expected_file and os.path.exists(expected_file):
                        exported_count += 1
                        self._overall_counter += 1
                        self._update_progress(
                            int(self._overall_counter * 100.0 / max(1, self._overall_total)),
                            "Exported {} to Image".format(element_name)
                        )
                        self._last_output_file = expected_file
                        self.update_export_item_progress(item.SheetNumber, "IMG", 100)
                    else:
                        self._note_failure(item.SheetNumber, "IMG",
                                           "Revit wrote no image file")

                except Exception as ex:
                    logger.error("Error exporting {} to Image: {}".format(element_name, ex))
                    self._note_failure(str(getattr(item, 'SheetNumber', element_name)),
                                       "IMG", ex)

            return exported_count

        except Exception as ex:
            logger.error("Image export failed: {}".format(ex))
            return 0


    def minimize_button_clicked(self, sender, e):
        self.WindowState = WindowState.Minimized

    def maximize_button_clicked(self, sender, e):
        if self.WindowState == WindowState.Maximized:
            self.WindowState = WindowState.Normal
            self.btn_maximize.ToolTip = "Maximize"
        else:
            self.WindowState = WindowState.Maximized
            self.btn_maximize.ToolTip = "Restore"

    def close_button_clicked(self, sender, e):
        self.Close()


# ── MAIN SCRIPT ──


def show_batchout_dialog():
    """Show the window modeless, or modal when the persistent engine is not active."""
    if not revit.doc:
        forms.alert("Please open a Revit document first.", exitscript=True)
        return None

    window = ExportManagerWindow()
    if window.modeless:
        window.show(modal=False)
    else:
        logger.info("Persistent engine not active — running modal. "
                    "Reload pyRevit to enable modeless mode.")
        if hasattr(window, 'status_text') and window.status_text is not None:
            window.status_text.Text = ("Modal mode — run pyRevit Reload once to "
                                       "unlock working in Revit alongside this window")
        window.show(modal=True)
    return window
