# -*- coding: utf-8 -*-
"""
_cloud_links.py
===============
Autodesk Docs (ACC / BIM 360) models for Batch Link: link them from the cloud,
never from their Desktop Connector copy on disk.

Revit has no API to browse Autodesk Docs. A cloud link needs the region,
project GUID and model GUID (ModelPathUtils.ConvertCloudGUIDsToCloudPath), so
the GUIDs come from what Revit already knows on this PC (collect_known_models)
and a Desktop Connector file is matched to one of them by project, folder and
file name. Revit links a cloud model only into a cloud model: Create / LoadFrom
raise "Revit cannot link a cloud model to non-cloud model".

Pure Python down to REVIT PART (tested by dev/test_batch_link.py); the Revit
part imports the API inside each function.

Part of T3Lab Extension.
"""

import os
import re

CLOUD_PREFIX = u"Autodesk Docs://"
# "BIM 360://" is how older releases wrote the same paths.
_CLOUD_SCHEMES = ("autodesk docs://", "bim 360://")
# Desktop Connector roots: Autodesk Docs (ACC) and the older BIM 360 connector.
_DC_ROOTS = ("accdocs", "bim 360")
_PROJECT_FILES = "project files"
_GUID = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$", re.I)
_SPLIT = re.compile(r"[\\/]+")

# Where a model's GUIDs were read, best source first: on a conflict the
# first one added wins (collect_known_models adds them in this order).
SOURCE_HOST = "this model"
SOURCE_LINKED = "a link in this model"
SOURCE_IN_LINK = "a link inside a linked model"
SOURCE_OPEN = "a cloud model open in Revit"
SOURCE_CACHE = "Revit's cloud cache on this PC"

STATUS_READY = "Ready"
STATUS_LINKED = "Already Linked"
STATUS_HOST = "Host Model"
STATUS_AS_FILE = "Linked as file"
STATUS_UNKNOWN = "Cloud ID unknown"

HOW_TO_KNOW = (u"Link or open it once from Autodesk Docs in Revit "
               u"(Insert \u203a Link Revit \u203a Autodesk Docs), then press Refresh.")
REASON_UNKNOWN = u"Revit has no Autodesk Docs ID for this model on this PC yet. " + HOW_TO_KNOW
REASON_NOT_DC = (u"Not inside a Desktop Connector folder "
                 u"(\u2026\\ACCDocs\\<hub>\\<project>\\Project Files\\\u2026).")


def is_guid(text):
    return bool(_GUID.match(text or ""))


def guid_text(value):
    """Lower-case GUID text of a Guid / "{...}" / "b.<guid>" value, or ""."""
    text = (u"{}".format(value) if value is not None else u"").strip().strip("{}").lower()
    if text.startswith("b."):
        text = text[2:]
    return text if is_guid(text) else ""


def is_cloud_path(path):
    return (path or "").strip().lower().startswith(_CLOUD_SCHEMES)


def split_cloud_path(path):
    """(project, folders, file name) of "Autodesk Docs://Project/A/B/M.rvt", or None."""
    if not is_cloud_path(path):
        return None
    rest = path.strip().split("://", 1)[1]
    parts = [p for p in _SPLIT.split(rest) if p.strip()]
    if len(parts) < 2:
        return None
    return parts[0], tuple(parts[1:-1]), parts[-1]


def desktop_connector_parts(path):
    """(project, segments under Project Files) of a Desktop Connector path, or None.

    ``E:\\DC\\ACCDocs\\<hub>\\IR2\\Project Files\\A\\M.rvt`` -> ("IR2", ("A", "M.rvt")).
    A folder at project level gives (project, ()); anything outside Project
    Files (e.g. Plans) is not a model folder and gives None.
    """
    parts = [p for p in _SPLIT.split(path or "") if p]
    for index, part in enumerate(parts):
        if part.strip().lower() not in _DC_ROOTS or len(parts) <= index + 2:
            continue
        rest = parts[index + 3:]
        if not rest:
            return parts[index + 2], ()
        if rest[0].strip().lower() != _PROJECT_FILES:
            return None
        return parts[index + 2], tuple(rest[1:])
    return None


def cloud_location(project, folders):
    """"Autodesk Docs://Project/A/B" ("" without a project)."""
    if not project:
        return u""
    return CLOUD_PREFIX + u"/".join((project,) + tuple(folders or ()))


def _low(items):
    return tuple(i.strip().lower() for i in items)


def parse_cache_path(path):
    """(project GUID, model GUID) of a file in Revit's cloud cache, or None.

    CollaborationCache\\<account>\\<project GUID>\\[...\\]<model GUID>.rvt
    """
    parts = [p for p in _SPLIT.split(path or "") if p]
    stem, ext = os.path.splitext(parts[-1] if parts else "")
    if ext.lower() != ".rvt" or not is_guid(stem):
        return None
    lows = [p.lower() for p in parts]
    if "collaborationcache" not in lows:
        return None
    start = lows.index("collaborationcache")
    folders = parts[start + 1:-1]
    # Documented layout first: the project folder sits right under the account.
    candidates = folders[1:2] + folders
    for part in candidates:
        if is_guid(part) and part.lower() != stem.lower():
            return part.lower(), stem.lower()
    return None


class CloudModel(object):
    """One Autodesk Docs model: where it is and the IDs that link it."""

    def __init__(self, name, project, folders, project_guid, model_guid,
                 region="", source="", model_path=None):
        self.name = name or u""
        self.project = project or u""
        self.folders = tuple(folders or ())
        self.project_guid = guid_text(project_guid)
        self.model_guid = guid_text(model_guid)
        self.region = region or ""
        self.source = source or ""
        self.model_path = model_path    # the Revit ModelPath, when one was read

    @classmethod
    def from_path(cls, visible_path, project_guid, model_guid, region="",
                  source="", model_path=None):
        parts = split_cloud_path(visible_path)
        if parts is None:
            name = [p for p in _SPLIT.split((visible_path or u"").split("://")[-1]) if p]
            return cls(name[-1] if name else u"", u"", (), project_guid, model_guid,
                       region, source, model_path)
        project, folders, name = parts
        return cls(name, project, folders, project_guid, model_guid, region,
                   source, model_path)

    @property
    def location(self):
        return cloud_location(self.project, self.folders)

    @property
    def display_path(self):
        location = self.location
        return location + u"/" + self.name if location else self.name

    @property
    def has_name(self):
        """False for a cached file Revit could only name by its GUID."""
        return bool(self.name) and not is_guid(os.path.splitext(self.name)[0])


class CloudCatalog(object):
    """Every cloud model known on this PC, one entry per model GUID."""

    def __init__(self):
        self._models = {}
        self.linked_guids = set()     # already linked into the host (nested too)
        self.local_link_names = set()  # file names of the host's disk / DC links
        self.host_guid = ""
        self.host_project_guid = ""
        self.host_region = ""

    def add(self, model):
        if model is None or not model.model_guid or not model.project_guid:
            return None
        known = self._models.get(model.model_guid)
        if known is None:
            self._models[model.model_guid] = model
            return model
        if not known.has_name and model.has_name:
            known.name = model.name
        if not known.project and model.project:
            known.project, known.folders = model.project, model.folders
        elif (not known.folders and model.folders
              and known.project.lower() == model.project.lower()):
            known.folders = model.folders
        if not known.region:
            known.region = model.region
        if known.model_path is None:
            known.model_path = model.model_path
        return known

    def __len__(self):
        return len(self._models)

    def get(self, model_guid):
        return self._models.get(guid_text(model_guid))

    def models(self):
        found = [m for m in self._models.values() if m.has_name]
        found.sort(key=lambda m: m.display_path.lower())
        return found

    def fill_projects(self):
        """Give a model read without its path the project name its GUID has elsewhere."""
        for model in self._models.values():
            if not model.project:
                model.project = self.project_name(model.project_guid)

    def project_name(self, project_guid):
        project_guid = guid_text(project_guid)
        for model in self.models():
            if model.project_guid == project_guid and model.project:
                return model.project
        return u""

    def region_for(self, model):
        """Region to build `model`'s path with: its own, its project's, the host's."""
        if model.region:
            return model.region
        for other in self._models.values():
            if other.project_guid == model.project_guid and other.region:
                return other.region
        return self.host_region

    def resolve(self, project, folders, name):
        """(model, "") for a file at project/folders/name, or (None, reason).

        The folder must match: Autodesk Docs keeps separate copies with the same
        file name in WIP / SHARED / PUBLISHED, and linking the wrong one is worse
        than not linking. Only a cached path that carries no folders (older
        "BIM 360://Project/Model.rvt") matches on project and name alone.
        """
        same = [m for m in self.models()
                if m.name.lower() == (name or u"").lower()
                and m.project.lower() == (project or u"").strip().lower()]
        want = _low(folders or ())
        exact = [m for m in same if _low(m.folders) == want]
        if exact:
            return exact[0], u""
        loose = [m for m in same if not m.folders]
        if len(loose) == 1:
            return loose[0], u""
        if len(loose) > 1:
            return None, (u"{} cloud models are called {} in project {}. "
                          u"Link this one from Autodesk Docs in Revit.").format(
                              len(loose), name, project)
        if same:
            return None, (u"The known cloud model {} is in another folder ({}). "
                          + HOW_TO_KNOW).format(name, same[0].location)
        return None, REASON_UNKNOWN

    def resolve_name(self, project_guid, name):
        """(model, "") for the only model called `name` in a project, or (None, reason)."""
        project_guid = guid_text(project_guid)
        same = [m for m in self.models()
                if m.project_guid == project_guid
                and m.name.lower() == (name or u"").lower()]
        if len(same) == 1:
            return same[0], u""
        if same:
            return None, (u"{} cloud models are called {} in this project. "
                          u"Use Reload From on Revit's Manage Links.").format(len(same), name)
        return None, REASON_UNKNOWN

    def in_folder(self, project, folders, recursive=True):
        """Known models inside an Autodesk Docs folder (and below it when recursive)."""
        want = _low(folders or ())
        found = []
        for model in self.models():
            if model.project.lower() != (project or u"").strip().lower():
                continue
            have = _low(model.folders)
            if not model.folders:
                # Location unknown: only a project-level folder may claim it.
                if not want:
                    found.append(model)
            elif have == want or (recursive and have[:len(want)] == want):
                found.append(model)
        return found


class CloudCandidate(object):
    """One row of the Link Models grid in Autodesk Docs mode."""

    def __init__(self, name, location, model=None, local_path=u"",
                 status=STATUS_READY, severity="Success", selectable=True,
                 checked=True, tip=u""):
        self.name = name
        self.location = location
        self.model = model
        self.local_path = local_path
        self.status = status
        self.severity = severity
        self.selectable = selectable
        self.checked = checked
        self.tip = tip


def _model_candidate(catalog, model, local_path=u""):
    tip = u"{}\nLinks from Autodesk Docs (path type Cloud). ID found in {}.".format(
        model.display_path, model.source or "this PC")
    if local_path:
        tip += u"\nDesktop Connector copy (not linked): " + local_path
    if model.model_guid == catalog.host_guid:
        return CloudCandidate(model.name, model.location, model, local_path,
                              STATUS_HOST, "Warning", False, False,
                              u"This is the model you are working in.")
    if model.model_guid in catalog.linked_guids:
        return CloudCandidate(model.name, model.location, model, local_path,
                              STATUS_LINKED, "Warning", True, False, tip)
    if model.name.lower() in catalog.local_link_names:
        return CloudCandidate(model.name, model.location, model, local_path,
                              STATUS_AS_FILE, "Warning", True, False,
                              tip + u"\nAlready linked from a disk / Desktop Connector copy: "
                              u"Manage Links \u203a Switch to Cloud re-points that link "
                              u"instead of adding a second one.")
    return CloudCandidate(model.name, model.location, model, local_path, tip=tip)


def cloud_candidates(catalog, dc_files=(), folder=None, recursive=True):
    """Rows of the Link Models grid in Autodesk Docs mode, sorted by location.

    `folder` is desktop_connector_parts() of the Desktop Connector folder that
    narrows the list (None lists every known model); `dc_files` are the .rvt
    files found in it. A file and a known model with the same GUID are one row.
    """
    rows = {}
    for path in dc_files or ():
        parts = desktop_connector_parts(path)
        name = [p for p in _SPLIT.split(path or u"") if p][-1:] or [u""]
        if parts is None or not parts[1]:
            rows[u"file:" + path] = CloudCandidate(
                name[0], u"", None, path, STATUS_UNKNOWN, "Warning", False, False,
                REASON_NOT_DC)
            continue
        project, segments = parts
        model, reason = catalog.resolve(project, segments[:-1], segments[-1])
        if model is None:
            rows[u"file:" + path] = CloudCandidate(
                segments[-1], cloud_location(project, segments[:-1]), None, path,
                STATUS_UNKNOWN, "Warning", False, False, reason)
        else:
            rows[model.model_guid] = _model_candidate(catalog, model, path)

    known = (catalog.in_folder(folder[0], folder[1], recursive) if folder
             else catalog.models())
    for model in known:
        if model.model_guid not in rows:
            rows[model.model_guid] = _model_candidate(catalog, model)

    found = list(rows.values())
    found.sort(key=lambda r: (r.location.lower(), r.name.lower()))
    return found


def match_link(catalog, link_path):
    """(model, "") for the cloud model behind a disk / Desktop Connector link,
    or (None, reason). Used by Manage Links › Switch to Cloud."""
    if is_cloud_path(link_path):
        return None, u"Already a cloud link."
    name = [p for p in _SPLIT.split(link_path or u"") if p][-1:]
    if not name:
        return None, u"The link has no saved path."
    parts = desktop_connector_parts(link_path)
    if parts is not None and parts[1]:
        project, segments = parts
        return catalog.resolve(project, segments[:-1], segments[-1])
    # A plain disk / server copy: only the one model of that name in this project.
    return catalog.resolve_name(catalog.host_project_guid, name[0])


# -- REVIT PART --------------------------------------------------------------

_CENTRAL_PATHS = {}     # (path, mtime) -> BasicFileInfo.CentralPath of a cached file


def is_cloud_doc(doc):
    try:
        return bool(doc is not None and doc.IsModelInCloud)
    except Exception:
        return False


def model_from_model_path(model_path, visible_path=u"", source=""):
    """CloudModel of a cloud ModelPath (None for a disk / server path)."""
    try:
        if model_path is None or not model_path.CloudPath:
            return None
        project_guid = guid_text(model_path.GetProjectGUID())
        model_guid = guid_text(model_path.GetModelGUID())
        region = getattr(model_path, "Region", "") or ""
    except Exception:
        return None
    if not visible_path:
        try:
            from Autodesk.Revit.DB import ModelPathUtils
            visible_path = ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path) or u""
        except Exception:
            visible_path = u""
    return CloudModel.from_path(visible_path, project_guid, model_guid, region,
                                source, model_path)


def _doc_model(doc, source):
    if not is_cloud_doc(doc):
        return None
    try:
        model_path = doc.GetCloudModelPath()
    except Exception:
        return None
    try:
        visible = doc.PathName or u""
    except Exception:
        visible = u""
    return model_from_model_path(model_path, visible, source)


def _info_guid(info, key):
    try:
        if info is not None and info.ContainsKey(key):
            return guid_text(info[key])
    except Exception:
        pass
    return ""


def link_type_model(link_type, source=SOURCE_LINKED):
    """CloudModel a RevitLinkType points at, or None for a disk / server link."""
    from Autodesk.Revit.DB import ExternalResourceTypes
    visible = u""
    info = None
    try:
        ref = link_type.GetExternalResourceReference(
            ExternalResourceTypes.BuiltInExternalResourceTypes.RevitLink)
        if ref is not None:
            visible = ref.InSessionPath or u""
            info = ref.GetReferenceInformation()
    except Exception:
        pass
    try:
        ext = link_type.GetExternalFileReference()
        model = model_from_model_path(ext.GetPath() if ext is not None else None,
                                      visible, source)
        if model is not None:
            return model
    except Exception:
        pass
    # An unloaded link can lack a cloud ModelPath; its resource reference keeps the IDs.
    project_guid = _info_guid(info, "LinkedModelProjectId")
    model_guid = _info_guid(info, "LinkedModelModelId")
    if not (project_guid and model_guid):
        return None
    if not visible:
        from Snippets._links import element_name
        visible = element_name(link_type).split(" : ")[0].strip()
    return CloudModel.from_path(visible, project_guid, model_guid, "", source)


def _doc_link_models(doc, source, local_names=None):
    """Cloud models of doc's link types; the file names of the other links go
    into `local_names` when a set is given."""
    from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkType
    from Snippets._links import element_name
    found = []
    try:
        for link_type in FilteredElementCollector(doc).OfClass(RevitLinkType):
            model = link_type_model(link_type, source)
            if model is not None:
                found.append(model)
            elif local_names is not None:
                name = element_name(link_type).split(" : ")[0].strip()
                if name:
                    local_names.add(name.lower())
    except Exception:
        pass
    return found


def _loaded_link_docs(doc):
    from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkInstance
    docs, seen = [], set()
    try:
        for instance in FilteredElementCollector(doc).OfClass(RevitLinkInstance):
            link_doc = instance.GetLinkDocument()
            if link_doc is None or link_doc.Title in seen:
                continue
            seen.add(link_doc.Title)
            docs.append(link_doc)
    except Exception:
        pass
    return docs


def _other_open_docs(doc):
    try:
        return [d for d in doc.Application.Documents
                if not d.Equals(doc) and not d.IsFamilyDocument]
    except Exception:
        return []


def cache_root(doc):
    """Revit's default cloud cache folder for this release ("" when unknown).

    Revit 2025+ can move it (Options › Cloud Model); a moved cache is not read.
    """
    base = os.environ.get("LOCALAPPDATA", "")
    try:
        version = str(doc.Application.VersionNumber)
    except Exception:
        version = ""
    if not base or not version:
        return ""
    return os.path.join(base, "Autodesk", "Revit", "Autodesk Revit " + version,
                        "CollaborationCache")


def cache_files(root, project_guids):
    """[(path, project GUID, model GUID)] of cached cloud models in these projects."""
    wanted = set(guid_text(g) for g in project_guids or ())
    found = []
    try:
        accounts = os.listdir(root)
    except (OSError, TypeError):
        return found
    for account in accounts:
        account_dir = os.path.join(root, account)
        try:
            projects = os.listdir(account_dir)
        except OSError:
            continue
        for project in projects:
            if not is_guid(project) or (wanted and project.lower() not in wanted):
                continue
            project_dir = os.path.join(account_dir, project)
            for current, dirs, files in os.walk(project_dir):
                if current[len(project_dir):].count(os.sep) >= 2:
                    dirs[:] = []
                for file_name in files:
                    path = os.path.join(current, file_name)
                    ids = parse_cache_path(path)
                    if ids is not None:
                        found.append((path,) + ids)
    return found


def _central_path(path):
    """BasicFileInfo.CentralPath of a cached file, memoised per file version."""
    try:
        key = (path, os.path.getmtime(path))
    except OSError:
        return u""
    if key not in _CENTRAL_PATHS:
        text = u""
        info = None
        try:
            from Autodesk.Revit.DB import BasicFileInfo
            info = BasicFileInfo.Extract(path)
            text = info.CentralPath or u""
        except Exception:
            text = u""
        finally:
            try:
                if info is not None:
                    info.Dispose()
            except Exception:
                pass
        _CENTRAL_PATHS[key] = text
    return _CENTRAL_PATHS[key]


def scan_cache(root, project_guids, progress=None):
    """CloudModels named from the cached files of these projects."""
    files = cache_files(root, project_guids)
    models, seen = [], set()
    for index, (path, project_guid, model_guid) in enumerate(files, 1):
        if model_guid in seen:
            continue
        if progress is not None:
            progress(index, len(files))
        central = _central_path(path)
        if not central:
            continue
        seen.add(model_guid)
        models.append(CloudModel.from_path(central, project_guid, model_guid, "",
                                           SOURCE_CACHE))
    return models


def collect_known_models(doc, include_cache=True, progress=None):
    """CloudCatalog of every cloud model Revit knows on this PC.

    Sources, best first: this model, its link types (nested too), the link
    types inside each loaded link, the other models open in Revit and their
    links, then the cloud cache of this model's project (one BasicFileInfo
    read per new cached file). `progress(index, total)` follows the cache.
    """
    catalog = CloudCatalog()
    if doc is None:
        return catalog
    host = _doc_model(doc, SOURCE_HOST)
    if host is not None:
        catalog.host_guid = host.model_guid
        catalog.host_project_guid = host.project_guid
        catalog.host_region = host.region
        catalog.add(host)
    for model in _doc_link_models(doc, SOURCE_LINKED, catalog.local_link_names):
        catalog.add(model)
        catalog.linked_guids.add(model.model_guid)
    for link_doc in _loaded_link_docs(doc):
        for model in _doc_link_models(link_doc, SOURCE_IN_LINK):
            catalog.add(model)
    for other in _other_open_docs(doc):
        catalog.add(_doc_model(other, SOURCE_OPEN))
        for model in _doc_link_models(other, SOURCE_OPEN):
            catalog.add(model)
    if include_cache and catalog.host_project_guid:
        for model in scan_cache(cache_root(doc), [catalog.host_project_guid], progress):
            catalog.add(model)
    catalog.fill_projects()
    return catalog


def cloud_model_path(model, region=""):
    """Revit ModelPath that links `model` from Autodesk Docs."""
    if model.model_path is not None:
        return model.model_path
    import System
    from Autodesk.Revit.DB import ModelPathUtils
    region = model.region or region or getattr(ModelPathUtils, "CloudRegionUS", "US")
    return ModelPathUtils.ConvertCloudGUIDsToCloudPath(
        region, System.Guid.Parse(model.project_guid), System.Guid.Parse(model.model_guid))
