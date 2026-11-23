# -*- coding: utf-8 -*-
"""
BVBSExportDialog.py
===================
WPF dialog for BVBS Export: write ``.abs`` files (BVBS BF2D) for bending
machines from shape-driven rebar.

A wrong leg length is a wrong bar on the machine, so the window works in this
order and never skips a step:

1. **Self-test** when it opens - the writer must reproduce the BVBS guideline's
   reference lines byte for byte, otherwise Export stays disabled (E-BVBS-01).
2. **Preview** - every rebar set is read (centre-line -> legs -> outer legs, see
   ``Snippets/_rebar.py``) and listed per mark with the legs that WILL be
   written, so they can be compared with the bending schedule. Bars that cannot
   be written (free-form 3D, curved legs, no diameter, no mark) are listed as
   ``Skipped`` with the reason, never dropped silently.
3. **Write + verify** - each file goes to a temporary name, is read back and
   checked (checksum, structure, identical to the record it came from), and only
   then moved into place. "Verified" is only reported for files read back clean.

Read-only with respect to the model: no transaction is ever opened. The writer,
checksum, weight chain and file naming are pure Python in ``Snippets/_bvbs.py``
(unit-tested outside Revit); this module is the Revit and WPF glue.

Part of T3Lab Extension.
"""

import json
import os

import clr
clr.AddReference('System')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')

from System.Windows import Visibility, RoutedEventHandler
from System.Windows.Controls import CheckBox
from System.Windows.Media.Imaging import BitmapImage, BitmapCacheOption
from System import Uri, UriKind

from GUI.WPF_Base import T3WPFWindow
from GUI.T3Dialog import confirm as t3_confirm, show_error, show_info, show_warning

from Snippets import _assembly, _bvbs, _bvbs_read, _rebar
from Snippets._bvbs_read import (DEFAULT_GRADE, DEFAULT_PLAN, DEFAULT_REVISION, SEV_FAIL,
                                 SEV_OK, SEV_WARN, natural_key as _natural_key)
from Snippets._compat import eid_value, short_error
from Snippets._host import resolve_uidoc

GUI_DIR = os.path.dirname(__file__)
XAML_FILE = os.path.join(GUI_DIR, 'Tools', 'BVBSExport.xaml')
DIALOG_TITLE = "BVBS Export"

SCOPE_MODEL = "model"
SCOPE_SELECTION = "selection"
SCOPE_ASSEMBLY = "assembly"
SCOPE_PARTITION = "partition"


def _plural(count, word):
    return u"%d %s%s" % (count, word, u"" if count == 1 else u"s")


def _settings_path():
    from core.paths import user_data_path
    return user_data_path("bvbs_export.json")


# ── ROW ──────────────────────────────────────────────────────────────────────

class BarRow(object):
    """One preview line: a distinct leg chain under one mark, or a skipped set.

    Plain attributes only - the grid reads them through WPF bindings and the
    checkbox column through the string bridge (rule 24).
    """

    def __init__(self, mark, status, severity, record=None, diameter_mm=0.0, quantity=0,
                 length_mm=None, legs=u"", detail=u"", group=u"", ids=(), reason=u""):
        self.is_selected = False
        self.mark = mark
        self.dia_text = u"%d" % int(round(diameter_mm)) if diameter_mm else u"—"
        self.qty_text = u"%d" % quantity
        self.length_text = u"%d" % int(round(length_mm)) if length_mm else u"—"
        self.legs_text = legs or u"—"
        self.StatusText = status
        self.Severity = severity
        self.record = record
        self.quantity = quantity
        self.detail = detail
        self.group = group
        self.rebar_ids = list(ids)
        self.reason = reason

    @property
    def exportable(self):
        return self.record is not None and self.Severity != SEV_FAIL

    @classmethod
    def from_item(cls, item):
        row = cls(item.mark, item.status, item.severity, record=item.record,
                  diameter_mm=item.diameter_mm, quantity=item.quantity,
                  length_mm=item.length_mm, legs=item.legs, detail=item.detail,
                  group=item.group, ids=item.ids, reason=item.reason)
        row.is_selected = item.selected
        return row


# ── DIALOG ───────────────────────────────────────────────────────────────────

class BVBSExportDialog(T3WPFWindow):
    """Main window class for BVBS Export."""

    def __init__(self, doc=None):
        T3WPFWindow.__init__(self, XAML_FILE)
        self.doc = doc
        self._loading = True
        self._is_busy = False
        self._loaded_once = False
        self._partial = False             # the last preview was stopped half-way
        self._index = None                # RebarIndex, built once per window
        self._assemblies = []
        self._asm_choices = []            # AssemblyRecord behind each cb_assembly item
        self._partition_choices = []      # partition text behind each cb_partition item
        self._asm_marks = {}
        self._by_assembly = {}            # assembly id -> [RebarRecord]
        self._rows = []
        self._notes = []
        self._weight_counts = {}
        self._weight_table = _bvbs.load_weight_table()
        self._selftest_ok = True
        self._settings = self._read_settings()

        self._load_logo()
        self._wire_row_events()
        self._run_selftest()
        self._init_fields()
        self._loading = False
        self._update_summary()
        self._arm_first_load()

    # ── SETUP ────────────────────────────────────────────────────────────────

    def _load_logo(self):
        """Load and bind the T3Lab logo to title bar and window icon."""
        try:
            logo_path = os.path.join(GUI_DIR, 'T3Lab_logo.png')
            if os.path.exists(logo_path):
                bitmap = BitmapImage()
                bitmap.BeginInit()
                bitmap.CacheOption = BitmapCacheOption.OnLoad
                bitmap.UriSource = Uri(logo_path, UriKind.Absolute)
                bitmap.EndInit()
                bitmap.Freeze()
                if hasattr(self, 'logo_image') and self.logo_image:
                    self.logo_image.Source = bitmap
                self.Icon = bitmap
        except Exception:
            pass

    def _wire_row_events(self):
        """Row checkbox clicks bubble to the grid; a handler inside the
        DataTemplate would never be wired (own namescope)."""
        try:
            self.grid_preview.AddHandler(
                CheckBox.ClickEvent, RoutedEventHandler(self.grid_preview_checkbox_clicked), True)
        except Exception:
            pass

    def _run_selftest(self):
        """E-BVBS-01: refuse to export when the writer drifted from the guideline."""
        ok, message = _bvbs.self_test()
        self._selftest_ok = ok
        if not ok:
            self.txt_selftest.Text = message
            self.selftest_box.Visibility = Visibility.Visible
            self.btn_export.IsEnabled = False
            self.btn_preview.IsEnabled = False
            self._set_status(u"Export is disabled — the BVBS self-test failed.", SEV_FAIL)

    def _read_settings(self):
        try:
            with open(_settings_path(), "r", encoding="utf-8") as handle:
                data = json.load(handle)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}

    def _save_settings(self):
        """Per-user convenience only: a failure here must never stop an export."""
        try:
            data = {"folder": self.tb_folder.Text.strip(),
                    "grade": self.tb_grade.Text.strip(),
                    "revision": self.tb_revision.Text.strip()}
            with open(_settings_path(), "w", encoding="utf-8") as handle:
                json.dump(data, handle, indent=2)
        except Exception:
            pass

    def _default_folder(self):
        saved = self._settings.get("folder")
        if saved and os.path.isdir(saved):
            return saved
        try:
            model_path = self.doc.PathName
            if model_path and os.path.isdir(os.path.dirname(model_path)):
                return os.path.dirname(model_path)
        except Exception:
            pass
        documents = os.path.join(os.path.expanduser("~"), "Documents")
        return documents if os.path.isdir(documents) else os.path.expanduser("~")

    def _init_fields(self):
        project = u""
        try:
            project = self.doc.ProjectInformation.Number or u""
        except Exception:
            pass
        self.tb_project.Text = project
        self.tb_plan.Text = DEFAULT_PLAN
        self.tb_revision.Text = self._settings.get("revision") or DEFAULT_REVISION
        self.tb_grade.Text = self._settings.get("grade") or DEFAULT_GRADE
        self.tb_folder.Text = self._default_folder()
        self.chk_only_shape_driven.IsChecked = True
        self.chk_per_assembly.IsChecked = False
        self.set_items_source(self.grid_preview, [])

    def _arm_first_load(self):
        """Show the window first, then read the model: a progress bar nobody can
        see is no progress bar. Falls back to loading straight away."""
        if not self._selftest_ok:
            return
        try:
            self.ContentRendered += self._on_content_rendered
        except Exception:
            self._first_load()

    def _on_content_rendered(self, sender, e):
        self._first_load()

    # ── SHARED HELPERS ───────────────────────────────────────────────────────

    def _set_status(self, text, level=SEV_OK):
        self.status_text.Text = text
        key = {SEV_OK: "T3.Success.Accent", SEV_WARN: "T3.Warning.Accent",
               SEV_FAIL: "T3.Danger.Accent"}.get(level, "T3.Success.Accent")
        try:
            self.dot_status.Fill = self.FindResource(key)
        except Exception:
            pass

    def _scope(self):
        if self.rb_scope_selection.IsChecked:
            return SCOPE_SELECTION
        if self.rb_scope_assembly.IsChecked:
            return SCOPE_ASSEMBLY
        if self.rb_scope_partition.IsChecked:
            return SCOPE_PARTITION
        return SCOPE_MODEL

    def _busy_controls(self):
        return [self.btn_export, self.btn_preview, self.btn_folder, self.cb_assembly,
                self.cb_partition, self.rb_scope_model, self.rb_scope_selection,
                self.rb_scope_assembly, self.rb_scope_partition,
                self.chk_only_shape_driven, self.chk_per_assembly]

    def _progress_callback(self, label):
        """progress(index, total, text) for build_rebar_index; False stops it."""
        state = {"total": None}

        def callback(index, total, text=None):
            if state["total"] != total:
                state["total"] = total
                self._update_progress(0, maximum=max(total, 1))
            return self.step_progress(index, u"%s — %d / %d" % (label, index, total))
        return callback

    # ── FIRST LOAD: INDEX, ASSEMBLIES, PARTITIONS ────────────────────────────

    def _first_load(self):
        if self._loaded_once or self._is_busy:
            return
        self._loaded_once = True
        if self.doc is None:
            show_warning(u"Open a Revit project before running BVBS Export.", title=DIALOG_TITLE)
            return
        self._is_busy = True
        self.begin_progress(maximum=1, disable=self._busy_controls())
        try:
            self._index = _rebar.build_rebar_index(
                self.doc, progress=self._progress_callback(u"Reading reinforcement"))
            try:
                self._assemblies = _assembly.collect_assemblies(self.doc)
            except Exception as exc:
                self._assemblies = []
                show_warning(u"Assemblies could not be read, so the Assembly scope is empty. %s"
                             % short_error(exc), title=DIALOG_TITLE)
            self._build_maps()
        except Exception as exc:
            self._index = None
            show_error(u"The reinforcement could not be read from the model. %s Close this "
                       u"window and try again." % short_error(exc), title=DIALOG_TITLE)
        finally:
            self.end_progress()
            self._is_busy = False
        if self._index is None:
            self._update_summary()
            return
        if self._index.stopped:
            self._partial = True
            self._set_status(u"Reading was stopped — refresh the preview to read the whole "
                             u"model.", SEV_WARN)
        self._fill_scope_combos()
        self._pick_initial_scope()
        self._rebuild_preview()

    def _build_maps(self):
        self._asm_marks, self._by_assembly = _bvbs_read.assembly_maps(self._index,
                                                                      self._assemblies)

    def _fill_scope_combos(self):
        self._loading = True
        try:
            self._asm_choices = sorted(self._assemblies,
                                       key=lambda a: (_natural_key(a.mark), a.id))
            labels = [u"%s (%d)" % (a.mark or u"(no mark)", a.id) for a in self._asm_choices]
            self.set_items_source(self.cb_assembly, labels)
            if labels:
                self.cb_assembly.SelectedIndex = 0

            counts = _bvbs_read.partition_counts(self._index)
            self._partition_choices = sorted(counts, key=_natural_key)
            self.set_items_source(self.cb_partition, [u"%s (%d)" % (p, counts[p])
                                                      for p in self._partition_choices])
            if self._partition_choices:
                self.cb_partition.SelectedIndex = 0
        finally:
            self._loading = False

    def _selection_ids(self):
        try:
            uidoc = resolve_uidoc()
            return [eid_value(item) for item in uidoc.Selection.GetElementIds()]
        except Exception:
            return []

    def _pick_initial_scope(self):
        """Open on the selection when something rebar-related is selected."""
        if self._selection_rows():
            self._loading = True
            try:
                self.rb_scope_selection.IsChecked = True
            finally:
                self._loading = False
        self._update_scope_controls()

    def _update_scope_controls(self):
        scope = self._scope()
        self.cb_assembly.IsEnabled = scope == SCOPE_ASSEMBLY and bool(self._asm_choices)
        self.cb_partition.IsEnabled = scope == SCOPE_PARTITION and bool(self._partition_choices)

    # ── SCOPE ────────────────────────────────────────────────────────────────

    def _assembly_rows(self, assembly):
        return _bvbs_read.assembly_rows(self._index, self._by_assembly, assembly)

    def _selection_rows(self):
        """Rebar for the selection: rebar itself, the rebar of selected hosts, the
        rebar of selected assemblies. Unique, in selection order."""
        if self._index is None:
            return []
        return _bvbs_read.selection_rows(self._index, self._assemblies, self._by_assembly,
                                         self._selection_ids())

    def _scope_rows(self):
        """(rows in scope, group label forced for all of them or None)."""
        if self._index is None:
            return [], None
        scope = self._scope()
        if scope == SCOPE_SELECTION:
            return self._selection_rows(), None
        if scope == SCOPE_ASSEMBLY:
            at = self.cb_assembly.SelectedIndex
            if 0 <= at < len(self._asm_choices):
                assembly = self._asm_choices[at]
                return self._assembly_rows(assembly), (assembly.mark or u"")
            return [], None
        if scope == SCOPE_PARTITION:
            at = self.cb_partition.SelectedIndex
            if 0 <= at < len(self._partition_choices):
                wanted = self._partition_choices[at]
                return [r for r in self._index.rows if r.partition == wanted], None
            return [], None
        return list(self._index.rows), None

    # ── PREVIEW: READ EVERY BAR (Snippets/_bvbs_read.py) ──────────────────────

    def _rebuild_preview(self):
        """Read every bar in the current scope and fill the grid. Never raises."""
        if self._index is None or self._is_busy or not self._selftest_ok:
            return
        self._is_busy = True
        self._partial = False
        self._weight_counts = {}
        self._notes = []
        try:
            self._rows = self._read_scope()
        except Exception as exc:
            self._rows = []
            show_error(u"The preview could not be built. %s Choose another scope or refresh and "
                       u"try again." % short_error(exc), title=DIALOG_TITLE)
        finally:
            self._is_busy = False
        self._show_rows()

    def _read_scope(self):
        rows_in_scope, forced_group = self._scope_rows()
        sets, self._notes = _bvbs_read.bar_sets(rows_in_scope,
                                                bool(self.chk_only_shape_driven.IsChecked))

        def progress(number, total):
            return bool(self.step_progress(
                number, u"Reading centre-lines — %d / %d" % (number, total)))

        self.begin_progress(maximum=max(len(sets), 1), disable=self._busy_controls())
        try:
            items, stopped = _bvbs_read.read_sets(self.doc, sets, forced_group, self._asm_marks,
                                                  self._weight_table, self._weight_counts,
                                                  progress=progress)
        finally:
            self.end_progress()
        if stopped:
            self._partial = True
        return [BarRow.from_item(item) for item in items]

    # ── GRID & SUMMARY ───────────────────────────────────────────────────────

    def _show_rows(self):
        self.set_items_source(self.grid_preview, self._rows)
        self.grid_preview_empty.Visibility = Visibility.Collapsed if self._rows \
            else Visibility.Visible
        self.txt_weight_source.Text = _bvbs.weight_caption(self._weight_counts)
        notes = list(self._notes)
        if self._partial:
            notes.insert(0, u"The preview was stopped before every bar was read — refresh it "
                            u"before exporting.")
        self.txt_note.Text = u" ".join(notes)
        self.txt_note.Visibility = Visibility.Visible if notes else Visibility.Collapsed
        self.txt_row_detail.Text = u"Select a row to see where its numbers come from."
        self._update_summary()

    def _ticked(self):
        return [r for r in self._rows if r.is_selected and r.exportable]

    def _update_summary(self, keep_status=False):
        """Counts, button label and status line from the current rows.

        ``keep_status`` leaves the status line alone - used right after an
        export, whose result must not be overwritten by "Ready".
        """
        ticked = self._ticked()
        bars = sum(r.quantity for r in ticked)
        exportable = [r for r in self._rows if r.exportable]
        skipped = sum(r.quantity for r in self._rows if not r.exportable)
        to_check = sum(1 for r in exportable if not r.is_selected)
        self.btn_export_label.Text = u"Export %s" % _plural(bars, u"bar")
        self.btn_export.IsEnabled = (bool(ticked) and self._selftest_ok and not self._partial
                                     and not self._is_busy)
        self.btn_preview.IsEnabled = self._selftest_ok and not self._is_busy
        self._sync_header()
        if self._index is None or not self._selftest_ok or self._is_busy:
            return

        if self._rows:
            parts = [u"%s in %s" % (_plural(sum(r.quantity for r in exportable), u"bar"),
                                    _plural(len(exportable), u"mark"))]
            if skipped:
                parts.append(u"%d skipped" % skipped)
            if to_check:
                parts.append(u"%d to check" % to_check)
            self.txt_preview_count.Text = u" · ".join(parts)
        else:
            self.txt_preview_count.Text = u"Nothing to preview"

        if keep_status:
            return
        if self._partial:
            self._set_status(u"Preview incomplete — refresh it before exporting.", SEV_WARN)
        elif not self._rows:
            self._set_status(u"Ready — no rebar in this scope.", SEV_WARN)
        else:
            text = u"Ready — %s in %s ticked" % (_plural(bars, u"bar"),
                                                  _plural(len(ticked), u"mark"))
            if skipped:
                text += u" · %d skipped" % skipped
            self._set_status(text, SEV_WARN if (skipped or to_check) else SEV_OK)

    def _sync_header(self):
        """Header checkbox: ticked when every writable row is, unticked when none is."""
        try:
            writable = [r for r in self._rows if r.exportable]
            ticked = sum(1 for r in writable if r.is_selected)
            if not writable or ticked == 0:
                state = False
            elif ticked == len(writable):
                state = True
            else:
                state = None
            header = getattr(self, "chk_all_grid_preview", None) \
                or self.FindName("chk_all_grid_preview")
            header.IsChecked = state
        except Exception:
            pass

    # ── EVENT HANDLERS: WINDOW ───────────────────────────────────────────────

    def close_button_clicked(self, sender=None, e=None):
        # The title-bar X (and Esc) is the only close control: it must not close
        # the window under a running scan or export.
        if self._is_busy:
            return
        self.Close()

    # ── EVENT HANDLERS: SCOPE & OPTIONS ──────────────────────────────────────

    def scope_changed(self, sender, e):
        if self._loading:
            return
        self._update_scope_controls()
        self._rebuild_preview()

    def assembly_changed(self, sender, e):
        if self._loading:
            return
        if self._scope() == SCOPE_ASSEMBLY:
            at = self.cb_assembly.SelectedIndex
            if 0 <= at < len(self._asm_choices) and self._plan_is_default_text():
                self.tb_plan.Text = self._asm_choices[at].mark or DEFAULT_PLAN
            self._rebuild_preview()

    def partition_changed(self, sender, e):
        if self._loading:
            return
        if self._scope() == SCOPE_PARTITION:
            self._rebuild_preview()

    def _plan_is_default_text(self):
        """True while tb_plan still holds the default or an assembly mark we put there."""
        text = self.tb_plan.Text.strip()
        return text in ("", DEFAULT_PLAN) or text in {a.mark for a in self._assemblies}

    def shape_filter_changed(self, sender, e):
        if self._loading:
            return
        self._rebuild_preview()

    def option_changed(self, sender, e):
        # One file per assembly only changes file naming; the preview stays valid.
        self._update_summary()

    def preview_clicked(self, sender, e):
        if self._is_busy:
            return
        self._rebuild_preview()

    # ── EVENT HANDLERS: GRID ─────────────────────────────────────────────────

    def grid_preview_checkbox_clicked(self, sender, e):
        # A row that cannot be written must never stay ticked.
        stale = [row for row in self._rows if not row.exportable and row.is_selected]
        for row in stale:
            row.is_selected = False
        if stale:
            try:
                self.grid_preview.Items.Refresh()
            except Exception:
                pass
        self._update_summary()

    def select_all_grid_preview_clicked(self, sender, e):
        self.toggle_all_rows(self.grid_preview, "is_selected", sender.IsChecked)
        for row in self._rows:
            if not row.exportable:
                row.is_selected = False
        try:
            self.grid_preview.Items.Refresh()
        except Exception:
            pass
        self._update_summary()

    def grid_preview_selection_changed(self, sender, e):
        row = self.grid_preview.SelectedItem
        if row is None:
            self.txt_row_detail.Text = u"Select a row to see where its numbers come from."
            return
        try:
            self.txt_row_detail.Text = row.detail
        except Exception:
            pass

    # ── EVENT HANDLERS: FOLDER & EXPORT ──────────────────────────────────────

    def folder_clicked(self, sender, e):
        try:
            import _cpython_bootstrap
            _cpython_bootstrap.install_forms_shim()
        except Exception:
            pass
        try:
            from pyrevit import forms
            folder = forms.pick_folder(title=u"Choose the folder for the .abs files")
        except Exception as exc:
            show_warning(u"The folder picker could not open. %s Type the folder path into the "
                         u"Output folder box instead." % short_error(exc), title=DIALOG_TITLE)
            return
        if folder:
            self.tb_folder.Text = folder

    def _apply_fields(self, records):
        return _bvbs_read.apply_fields(records, self.tb_project.Text.strip(),
                                       self.tb_plan.Text.strip(), self.tb_revision.Text.strip(),
                                       self.tb_grade.Text.strip())

    def _skipped_summary(self):
        """' · 4 skipped (free-form 3D, curved leg)' or '' when nothing was skipped."""
        counts = {}
        for row in self._rows:
            if not row.exportable:
                reason = row.reason or u"skipped"
                counts[reason] = counts.get(reason, 0) + row.quantity
        if not counts:
            return u""
        return u" · %d skipped (%s)" % (sum(counts.values()), u", ".join(sorted(counts)))

    def export_clicked(self, sender, e):
        if self._is_busy or not self._selftest_ok:
            return
        rows = self._ticked()
        if not rows:
            show_warning(u"Nothing to export. Tick at least one ready row in the preview, or "
                         u"choose another scope.", title=DIALOG_TITLE)
            return
        if self._partial:
            show_warning(u"The preview was stopped before every bar was read. Refresh the preview "
                         u"first so the file holds the whole scope.", title=DIALOG_TITLE)
            return

        folder = self.tb_folder.Text.strip()
        if not folder or not os.path.isdir(folder):
            show_warning(u"The output folder does not exist: %s. Choose an existing folder with "
                         u"“Choose folder…”." % (folder or u"(empty)"),
                         title=DIALOG_TITLE)
            return

        records = [r.record for r in rows]
        plan = self._apply_fields(records)
        files = _bvbs.plan_files(records, bool(self.chk_per_assembly.IsChecked),
                                 plan or DEFAULT_PLAN)
        paths = [(os.path.join(folder, name), recs) for name, recs in files]
        existing = [path for path, _ in paths if os.path.exists(path)]
        if existing:
            names = u"\n".join(os.path.basename(p) for p in existing[:8])
            if len(existing) > 8:
                names += u"\n… and %d more" % (len(existing) - 8)
            if not t3_confirm(
                    u"Overwrite %s in %s?" % (_plural(len(existing), u"existing file"), folder),
                    title=DIALOG_TITLE, details=names,
                    ok_text=u"Overwrite %s" % _plural(len(existing), u"file"), danger=True):
                return
        self._write_all(paths, folder)

    def _write_all(self, paths, folder):
        """Write each file (tmp -> verify -> move) and report what really happened."""
        self._is_busy = True
        self.begin_progress(maximum=max(len(paths), 1), disable=self._busy_controls())
        written_files, failed_files, not_written = [], [], []
        bars = 0
        try:
            for number, (path, recs) in enumerate(paths, 1):
                if not self.step_progress(
                        number - 1, u"Writing %s — %d / %d"
                        % (os.path.basename(path), number, len(paths))):
                    not_written = [os.path.basename(p) for p, _ in paths[number - 1:]]
                    break
                try:
                    _, _, problems = _bvbs.write_verified(path, recs)
                except Exception as exc:
                    problems = [u"unexpected error: %s" % short_error(exc)]
                if problems:
                    failed_files.append((os.path.basename(path), problems))
                else:
                    written_files.append(path)
                    bars += sum(r.quantity for r in recs)
        finally:
            self.end_progress()
            self._is_busy = False

        if written_files:
            self._save_settings()
        self._report_export(written_files, failed_files, not_written, bars, folder)
        self._update_summary(keep_status=True)

    def _report_export(self, written_files, failed_files, not_written, bars, folder):
        if failed_files:
            lines = []
            for name, problems in failed_files:
                lines.append(u"%s:" % name)
                lines.extend(u"  • %s" % problem for problem in problems[:4])
            count = _plural(len(failed_files), u"file")
            self._set_status(u"%s failed read-back verification — not written." % count, SEV_FAIL)
            show_error(
                u"%s failed the read-back check, so it was not written and any earlier file of "
                u"the same name was left untouched. Do not send these files to a machine — "
                u"report this to T3Lab with the details below." % count,
                title=DIALOG_TITLE, details=u"\n".join(lines))
            return
        if not written_files:
            self._set_status(u"Nothing was written — export stopped.", SEV_WARN)
            return
        skipped = self._skipped_summary()
        status = u"Wrote %s to %s%s · verified" % (
            _plural(bars, u"bar"), _plural(len(written_files), u"file"), skipped)
        if not_written:
            status += u" · stopped before %s" % _plural(len(not_written), u"file")
        self._set_status(status, SEV_WARN if (skipped or not_written) else SEV_OK)
        details = u"\n".join(written_files)
        if not_written:
            details += u"\n\nNot written (stopped): " + u", ".join(not_written)
        show_info(
            u"%s. Every file was read back from %s and its checksums verified." % (status, folder),
            title=DIALOG_TITLE, details=details)


def show_bvbs_export(doc):
    """Entry point used by the pushbutton."""
    BVBSExportDialog(doc).ShowDialog()
