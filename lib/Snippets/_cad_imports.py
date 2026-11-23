# -*- coding: utf-8 -*-
"""CAD imports and CAD links (ImportInstance): list, delete, owner view.

Shared by the ManaDWG window (GUI/ManaDWGDialog.py) and the `cad_imports` tool
(core/tools/links.py), so no UI here. Tests: dev/test_dwg_delete.py,
dev/test_dwg_view_navigation.py, dev/test_tools_data_domain.py.
"""

import logging

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ImportInstance,
    SubTransaction,
    Transaction,
    TransactionStatus,
    ElementId,
    View,
)

from Snippets._compat import eid_value, elem_name

logger = logging.getLogger('T3Lab.cad_imports')


class DWGItem(object):
    """View-model for one CAD import or CAD link instance."""
    def __init__(self, instance, link_type, is_link, name, view_names, file_path, owner_view_id=None):
        self.Instance   = instance
        self.LinkType   = link_type
        self.IsLink     = is_link
        self.DWGType    = "Link" if is_link else "Import"
        self.Name       = name
        self.ViewNames  = view_names
        self.OwnerViewId = owner_view_id
        self.ViewCount  = 1
        self.FilePath   = file_path
        self.IsSelected = False


def collect_dwg_items(document):
    """Return a list of DWGItem for all ImportInstance elements in the document."""
    items = []
    if not document:
        return items
    try:
        instances = FilteredElementCollector(document)\
            .OfClass(ImportInstance)\
            .ToElements()
    except Exception as ex:
        logger.warning("Could not collect ImportInstances: {}".format(ex))
        return items

    for inst in instances:
        try:
            is_link   = inst.IsLinked
            type_id   = inst.GetTypeId()
            link_type = document.GetElement(type_id) if type_id != ElementId.InvalidElementId else None

            if link_type is not None:
                try:
                    name = elem_name(link_type)
                except Exception:
                    name = "Unknown"
            else:
                try:
                    name = inst.Name
                except Exception:
                    name = "Unknown"

            file_path = ""
            if is_link and link_type is not None:
                try:
                    ext_ref = link_type.GetExternalFileReference()
                    if ext_ref is not None:
                        file_path = str(ext_ref.GetAbsolutePath())
                except Exception:
                    pass

            owner_view_id = None
            try:
                owner_view_id = inst.OwnerViewId
                if owner_view_id == ElementId.InvalidElementId:
                    view_names = "All Views (3D)"
                else:
                    view_elem  = document.GetElement(owner_view_id)
                    view_names = view_elem.Name if view_elem is not None else "Unknown View"
            except Exception:
                view_names = "Unknown View"

            items.append(DWGItem(inst, link_type, is_link, name, view_names, file_path, owner_view_id))

        except Exception as item_ex:
            logger.warning("Skipped one ImportInstance: {}".format(item_ex))

    return items


def _delete_cad_instance(document, instance):
    """Delete one CAD instance, unpinning it first; a failure restores its pin.

    Document.Delete(ElementId) refuses a pinned element with "ElementId cannot be deleted"
    while CanDeleteElement says True (Revit 2026.5, 2026-11-02). Revit pins Origin-to-Origin
    imports itself. Test: dev/test_dwg_delete.py.
    """
    sub = SubTransaction(document)
    sub.Start()
    try:
        if instance.Pinned:
            instance.Pinned = False
        document.Delete(instance.Id)
        sub.Commit()
    except Exception:
        sub.RollBack()
        raise


def delete_cad_items(document, items):
    """Delete the items' instances and any link type left unused; return (deleted, failures).

    failures = [(name, reason)]. Raises when the transaction as a whole does not commit.
    """
    deleted = 0
    failures = []
    link_types = {}
    t = Transaction(document, "T3Lab: Delete CAD Items")
    t.Start()
    try:
        for item in items:
            if item.IsLink and item.LinkType is not None:
                link_types[eid_value(item.LinkType.Id)] = item
            if item.Instance is None:
                continue
            eid = item.Instance.Id
            try:
                _delete_cad_instance(document, item.Instance)
                deleted += 1
            except Exception as ex:
                logger.warning("Could not delete element {}: {}".format(eid_value(eid), ex))
                failures.append((item.Name, ex))

        used_type_ids = set(
            eid_value(inst.GetTypeId())
            for inst in FilteredElementCollector(document).OfClass(ImportInstance).ToElements()
        )
        for type_value, item in link_types.items():
            if type_value in used_type_ids:
                continue
            try:
                document.Delete(item.LinkType.Id)
            except Exception as ex:
                logger.warning("Could not delete link type {}: {}".format(type_value, ex))
                failures.append(("{} (link type)".format(item.Name), ex))

        if t.Commit() != TransactionStatus.Committed:
            raise RuntimeError("Revit rolled the delete back when committing it.")
    except Exception:
        if not t.HasEnded():
            t.RollBack()
        raise
    return deleted, failures


def resolve_dwg_view(document, view_id):
    """Resolve only the stored owner view, never a possibly ambiguous view name."""
    if document is None or not document.IsValidObject:
        raise ValueError("The Revit document is no longer available.")
    if view_id is None or view_id == ElementId.InvalidElementId:
        raise ValueError("This CAD item has no specific owner view (All Views / 3D).")
    view = document.GetElement(view_id)
    if view is None or not isinstance(view, View) or not view.IsValidObject:
        raise ValueError("The owner view no longer exists. Refresh the CAD list.")
    if view.IsTemplate:
        raise ValueError("View templates cannot be opened.")
    return view
