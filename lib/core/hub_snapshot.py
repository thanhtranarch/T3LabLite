# -*- coding: utf-8 -*-
"""T3Lab Data Hub, Revit side: write one snapshot of a document's elements,
types and warnings to the hub inbox (docs/data-hub.md).

Revit only writes the file; core/hub_store.py (bridge process, outside Revit)
ingests it into SQLite, so Revit never holds a database lock and is blocked
only for the export itself. One snapshot = one complete pass: when
`max_seconds` runs out the file is dropped (a partial snapshot would read as
deleted elements).

File: <hub>/inbox/<model_key>-<UTC stamp>.ndjson, written as .part and renamed
when complete. One JSON object per line, ASCII-safe (core.jsonsafe):
  {"t":"header", ...}  {"t":"e", ...} per element/type  {"t":"w", ...} per
  warning  {"t":"end", "complete": true, counts, units}
Element keys: u UniqueId, i id, k "i"|"t", c category, f family, y type,
l level, ws workset, ph phase, n name, p {parameter: value}.
Doubles are numbers in the project's display units (units map in "end").

Runs on Revit's main thread under IronPython 2.7 (MCP server) or CPython 3.
Pure helpers are tested by lite_guard/test_data_hub.py; the Revit reads are NOT
verified in Revit yet.
"""
from __future__ import unicode_literals

import hashlib
import io
import os
import time

from core import jsonsafe

FORMAT = 'hub/1'
DEFAULT_MAX_SECONDS = 90
MAX_SECONDS_CAP = 110           # the MCP server waits 120 s for a tool
CHECK_EVERY = 200               # elements between clock checks
# Non-model categories worth having by default (BuiltInCategory names).
EXTRA_CATEGORIES = ('OST_Sheets', 'OST_Levels', 'OST_Grids', 'OST_ProjectInformation')


# ─── Pure helpers ──────────────────────────────────────────────────────────────

def hub_dir():
    """%LOCALAPPDATA%\\T3Lab\\hub (large files: never the roaming profile).
    T3LAB_HUB_DIR overrides it. Same rule as core/hub_store.py."""
    override = os.environ.get('T3LAB_HUB_DIR')
    if override:
        return override
    base = os.environ.get('LOCALAPPDATA') or os.path.join(os.path.expanduser('~'), '.t3lab')
    return os.path.join(base, 'T3Lab', 'hub')


def inbox_dir(root=None):
    return os.path.join(root or hub_dir(), 'inbox')


def model_key(identity):
    """12 hex of the model's identity (central / cloud / file path, lowercased)."""
    text = (identity or '').strip().lower()
    return hashlib.sha256(text.encode('utf-8')).hexdigest()[:12]


def utc_stamp(now=None):
    return time.strftime('%Y%m%dT%H%M%SZ', time.gmtime(now))


def utc_iso(now=None):
    return time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime(now))


def snapshot_name(key, now=None):
    return '{}-{}'.format(key, utc_stamp(now))


def clean_names(value):
    """['Walls', ' doors '] -> ['walls', 'doors']; None / [] -> None (= default)."""
    if not value:
        return None
    if not isinstance(value, (list, tuple)):
        value = [value]
    out = []
    for item in value:
        text = '{}'.format(item).strip().lower()
        if text and text not in out:
            out.append(text)
    return out or None


def clamp_seconds(value):
    try:
        seconds = float(value)
    except (TypeError, ValueError):
        return DEFAULT_MAX_SECONDS
    return max(5.0, min(float(MAX_SECONDS_CAP), seconds))


def number(value, digits=6):
    """Rounded float; an integer-valued one stays a float so the hash is stable."""
    try:
        return round(float(value), digits)
    except (TypeError, ValueError):
        return None


def int_or_text(text, fallback):
    """AsValueString of an Integer parameter: 'Yes' stays text, '3' -> 3."""
    text = (text or '').strip()
    if not text:
        return fallback
    try:
        return int(text)
    except ValueError:
        return text


class SnapshotWriter(object):
    """Lines to <inbox>/<name>.ndjson.part; commit() renames, abort() removes."""

    def __init__(self, folder, name):
        if not os.path.isdir(folder):
            os.makedirs(folder)
        self.name = name
        self.path = os.path.join(folder, name + '.ndjson')
        self._part = self.path + '.part'
        self._file = io.open(self._part, 'w', encoding='utf-8')
        self.counts = {'e': 0, 't': 0, 'w': 0}

    def write(self, record):
        self._file.write(jsonsafe.dumps(record, separators=(',', ':')))
        self._file.write('\n')

    def element(self, record):
        self.counts['t' if record.get('k') == 't' else 'e'] += 1
        self.write(record)

    def warning(self, record):
        self.counts['w'] += 1
        self.write(record)

    def commit(self, end):
        end = dict(end, t='end', complete=True, elements=self.counts['e'],
                   types=self.counts['t'], warnings=self.counts['w'])
        self.write(end)
        self._file.close()
        if os.path.exists(self.path):
            os.remove(self.path)
        os.rename(self._part, self.path)
        return self.path

    def abort(self):
        try:
            self._file.close()
        except Exception:
            pass
        try:
            os.remove(self._part)
        except OSError:
            pass


# ─── Revit reads ───────────────────────────────────────────────────────────────

def identity(doc):
    """(kind, identity text): cloud GUIDs, central path, file path, else title."""
    try:
        if doc.IsModelInCloud:
            path = doc.GetCloudModelPath()
            return 'cloud', 'cloud:{}/{}'.format(path.GetProjectGUID(), path.GetModelGUID())
    except Exception:
        pass
    try:
        if doc.IsWorkshared:
            from Autodesk.Revit.DB import ModelPathUtils
            central = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                doc.GetWorksharingCentralModelPath())
            if central:
                return 'central', '{}'.format(central)
    except Exception:
        pass
    path = '{}'.format(getattr(doc, 'PathName', '') or '')
    if path:
        return 'file', path
    return 'untitled', 'untitled:{}'.format(doc.Title)


def _header(doc, key, kind, now, categories, parameters):
    out = {'t': 'header', 'format': FORMAT, 'model_key': key, 'identity_kind': kind,
           'title': '{}'.format(doc.Title), 'taken_at': utc_iso(now),
           'scope': {'categories': categories, 'parameters': parameters}}
    try:
        app = doc.Application
        out['revit'] = '{}'.format(app.VersionNumber)
        out['revit_build'] = '{}'.format(app.VersionBuild)
    except Exception:
        pass
    for name, read in (('workshared', lambda: bool(doc.IsWorkshared)),
                       ('modified', lambda: bool(doc.IsModified))):
        try:
            out[name] = read()
        except Exception:
            pass
    try:
        from Autodesk.Revit.DB import Document
        version = Document.GetDocumentVersion(doc)
        out['version_guid'] = '{}'.format(version.VersionGUID)
        out['saves'] = int(version.NumberOfSaves)
    except Exception:
        pass
    try:
        path = '{}'.format(doc.PathName or '')
        if path:        # = the MCP audit's doc_sha, to join hub and audit later
            out['doc_path_sha'] = hashlib.sha256(path.lower().encode('utf-8')).hexdigest()[:12]
    except Exception:
        pass
    return out


class _Reader(object):
    """Per-export caches and the parameter/field reads of one element."""

    def __init__(self, doc, parameters):
        from Autodesk.Revit import DB
        self.DB = DB
        self.doc = doc
        self.parameters = parameters
        self.units = {}
        self._names = {}        # ElementId value -> name (levels, phases)
        self._types = {}        # type id value -> (family, type)
        self._worksets = {}
        self._workshared = False
        try:
            self._workshared = bool(doc.IsWorkshared)
            self._table = doc.GetWorksetTable() if self._workshared else None
        except Exception:
            self._table = None
        from Snippets._compat import eid_value
        self._eid = eid_value
        self._invalid = eid_value(DB.ElementId.InvalidElementId)
        from Autodesk.Revit.DB import BuiltInParameter
        self._level_bips = [getattr(BuiltInParameter, n) for n in (
            'FAMILY_LEVEL_PARAM', 'SCHEDULE_LEVEL_PARAM', 'INSTANCE_REFERENCE_LEVEL_PARAM',
            'WALL_BASE_CONSTRAINT') if hasattr(BuiltInParameter, n)]

    def _name_of(self, element_id):
        value = self._eid(element_id)
        if value == self._invalid:
            return None
        if value not in self._names:
            try:
                self._names[value] = '{}'.format(self.doc.GetElement(element_id).Name)
            except Exception:
                self._names[value] = None
        return self._names[value]

    def family_type(self, elem):
        try:
            type_id = elem.GetTypeId()
        except Exception:
            return None, None
        value = self._eid(type_id)
        if value == self._invalid:
            return None, None
        if value not in self._types:
            family = name = None
            try:
                etype = self.doc.GetElement(type_id)
                family = '{}'.format(etype.FamilyName or '') or None
                name = self._safe_name(etype)
            except Exception:
                pass
            self._types[value] = (family, name)
        return self._types[value]

    def _safe_name(self, elem):
        try:
            return '{}'.format(elem.Name) or None
        except Exception:
            try:        # pythonnet: ElementType.Name has no usable getter
                return '{}'.format(self.DB.Element.Name.GetValue(elem)) or None
            except Exception:
                return None

    def level(self, elem):
        try:
            name = self._name_of(elem.LevelId)
            if name:
                return name
        except Exception:
            pass
        for bip in self._level_bips:
            try:
                p = elem.get_Parameter(bip)
                if p is not None and p.HasValue:
                    name = self._name_of(p.AsElementId())
                    if name:
                        return name
            except Exception:
                continue
        return None

    def workset(self, elem):
        if self._table is None:
            return None
        try:
            wid = self._eid(elem.WorksetId)
            if wid not in self._worksets:
                self._worksets[wid] = '{}'.format(self._table.GetWorkset(elem.WorksetId).Name)
            return self._worksets[wid]
        except Exception:
            return None

    def phase(self, elem):
        try:
            return self._name_of(elem.CreatedPhaseId)
        except Exception:
            return None

    def value(self, p):
        """JSON value of one parameter, or None when it has none."""
        DB = self.DB
        try:
            if not p.HasValue:
                return None
            storage = p.StorageType
            if storage == DB.StorageType.String:
                text = p.AsString()
                return '{}'.format(text) if text else None
            if storage == DB.StorageType.Integer:
                return int_or_text(p.AsValueString(), p.AsInteger())
            if storage == DB.StorageType.Double:
                raw = p.AsDouble()
                try:
                    unit = p.GetUnitTypeId()
                    name = p.Definition.Name
                    if name not in self.units:
                        self.units[name] = '{}'.format(unit.TypeId).split('-')[0].split(':')[-1]
                    return number(DB.UnitUtils.ConvertFromInternalUnits(raw, unit))
                except Exception:
                    return number(raw)      # no unit (plain number)
            if storage == DB.StorageType.ElementId:
                if self._eid(p.AsElementId()) == self._invalid:
                    return None
                text = p.AsValueString()
                return '{}'.format(text) if text else self._eid(p.AsElementId())
        except Exception:
            return None
        return None

    def params(self, elem):
        out = {}
        if self.parameters:
            for wanted in self.parameters:
                try:
                    for p in elem.GetParameters(wanted) or []:
                        v = self.value(p)
                        if v is not None:
                            out.setdefault(p.Definition.Name, v)
                except Exception:
                    continue
            return out
        try:
            ordered = elem.GetOrderedParameters()
        except Exception:
            ordered = []
        for p in ordered:
            try:
                name = p.Definition.Name
            except Exception:
                continue
            if name in out:
                continue
            v = self.value(p)
            if v is not None:
                out[name] = v
        return out

    def record(self, elem, category, kind):
        if kind == 't':
            family = None
            try:
                family = '{}'.format(elem.FamilyName or '') or None
            except Exception:
                pass
            type_name = self._safe_name(elem)
        else:
            family, type_name = self.family_type(elem)
        rec = {'t': 'e', 'u': '{}'.format(elem.UniqueId), 'i': self._eid(elem.Id),
               'k': kind, 'c': category}
        fields = (('f', family), ('y', type_name), ('n', self._safe_name(elem)))
        if kind == 'i':
            fields += (('l', self.level(elem)), ('ws', self.workset(elem)),
                       ('ph', self.phase(elem)))
        for key, value in fields:
            if value:
                rec[key] = value
        params = self.params(elem)
        if params:
            rec['p'] = params
        return rec


def _category_filter(DB, eid_value, categories):
    """f(elem) -> category name to export, or None to skip."""
    from Autodesk.Revit.DB import BuiltInCategory
    extra = set()
    for name in EXTRA_CATEGORIES:
        if hasattr(BuiltInCategory, name):
            extra.add(int(getattr(BuiltInCategory, name)))
    model = DB.CategoryType.Model

    def pick(elem):
        try:
            cat = elem.Category
            if cat is None:
                return None
            name = '{}'.format(cat.Name)
            if not name:
                return None
            if categories:
                return name if name.lower() in categories else None
            if name.startswith('<'):            # <Sketch>, <Room Separation>, ...
                return None
            if cat.CategoryType != model and eid_value(cat.Id) not in extra:
                return None
            if getattr(elem, 'ViewSpecific', False):
                return None
            return name
        except Exception:
            return None
    return pick


def publish(doc, arguments=None, clock=time.time, folder=None):
    """Write one snapshot of `doc`; the result dict of the hub_publish tool."""
    arguments = arguments or {}
    categories = clean_names(arguments.get('categories'))
    parameters = arguments.get('parameters') or None
    if parameters and not isinstance(parameters, (list, tuple)):
        parameters = [parameters]
    parameters = ['{}'.format(p).strip() for p in parameters or [] if '{}'.format(p).strip()] or None
    budget = clamp_seconds(arguments.get('max_seconds', DEFAULT_MAX_SECONDS))

    from Autodesk.Revit import DB
    from Snippets._compat import eid_value
    started = clock()
    kind, ident = identity(doc)
    key = model_key(ident)
    writer = SnapshotWriter(folder or inbox_dir(), snapshot_name(key, started))
    reader = _Reader(doc, parameters)
    pick = _category_filter(DB, eid_value, categories)
    seen_categories = {}
    try:
        writer.write(_header(doc, key, kind, started, categories, parameters))
        passes = (('i', DB.FilteredElementCollector(doc).WhereElementIsNotElementType()),
                  ('t', DB.FilteredElementCollector(doc).WhereElementIsElementType()))
        scanned = 0
        for pass_kind, collector in passes:
            for elem in collector:
                scanned += 1
                if scanned % CHECK_EVERY == 0 and clock() - started > budget:
                    writer.abort()
                    done = writer.counts['e'] + writer.counts['t']
                    return {
                        'error': ('Snapshot not finished in {:.0f} s ({} elements written, '
                                  'nothing was published).'.format(budget, done)),
                        'error_code': 'hub_publish_timeout', 'nothing_ran': True,
                        'hint': ('Pass categories=[...] and/or parameters=[...] to publish '
                                 'less, or a larger max_seconds (max {}).'.format(MAX_SECONDS_CAP))}
                category = pick(elem)
                if category is None:
                    continue
                writer.element(reader.record(elem, category, pass_kind))
                if pass_kind == 'i':
                    seen_categories[category] = seen_categories.get(category, 0) + 1
        for warning in doc.GetWarnings():
            try:
                uids = []
                for eid in warning.GetFailingElements():
                    failing = doc.GetElement(eid)
                    if failing is not None:
                        uids.append('{}'.format(failing.UniqueId))
                writer.warning({'t': 'w', 'd': '{}'.format(warning.GetDescriptionText()),
                                's': '{}'.format(warning.GetSeverity()), 'u': sorted(uids)})
            except Exception:
                continue
        seconds = round(clock() - started, 1)
        path = writer.commit({'seconds': seconds, 'units': reader.units})
    except Exception:
        writer.abort()
        raise
    total = writer.counts['e'] + writer.counts['t']
    return {'published': True, 'snapshot': writer.name, 'model_key': key,
            'document': '{}'.format(doc.Title), 'elements': writer.counts['e'],
            'types': writer.counts['t'], 'warnings': writer.counts['w'],
            'categories': len(seen_categories), 'seconds': seconds,
            'elements_per_second': int(total / seconds) if seconds else total,
            'scope': {'categories': categories or 'default (model categories, sheets, '
                                                 'levels, grids)',
                      'parameters': parameters or 'all'},
            'file': path}
