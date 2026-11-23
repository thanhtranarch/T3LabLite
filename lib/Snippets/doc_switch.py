# -*- coding: utf-8 -*-
"""Find an open document and make it the active one in this Revit window, without a
click (switch_active_document, list_open_documents).

UIApplication.OpenAndActivateDocument activates a document that is already open;
it takes a file path for a local file and a cloud ModelPath for an Autodesk Docs /
BIM 360 model (whose PathName is not a file path). A document never saved has
neither, so the API cannot activate it. Must run in a Revit API context
(ExternalEvent), outside any transaction. Shared by T3Lab and T3Lab Lite;
IronPython 2.7 + CPython 3, Revit types imported inside functions.
Test: dev/test_doc_switch.py (lite_guard/test_doc_switch.py).
"""
from __future__ import unicode_literals

import ntpath          # Revit paths are Windows paths, on every engine and in tests

UNSAVED = ('"{}" has never been saved, and Revit can only activate a document by its '
           'file or cloud path. Save it once (then switch works), or click its tab.')


def _norm(path):
    return ntpath.normcase(ntpath.normpath(path)) if path else ''


def is_cloud(doc):
    try:
        return bool(getattr(doc, 'IsModelInCloud', False))
    except Exception:
        return False


def activation(doc):
    """{'can_activate', 'kind': 'file' | 'cloud' | 'unsaved'} for list_open_documents."""
    if is_cloud(doc):
        return {'can_activate': True, 'kind': 'cloud'}
    if getattr(doc, 'PathName', None):
        return {'can_activate': True, 'kind': 'file'}
    return {'can_activate': False, 'kind': 'unsaved'}


def match(docs, query):
    """(document, None) or (None, {'error', 'candidates'?}). Precedence: exact title,
    exact path, file name with or without extension, unique title substring."""
    q = (query or '').strip()
    if not q:
        return None, {'error': 'path_or_title is required.'}
    low, qpath = q.lower(), _norm(q)
    for d in docs:
        if d.Title.lower() == low:
            return d, None
    for d in docs:
        if d.PathName and not is_cloud(d) and _norm(d.PathName) == qpath:
            return d, None
    for d in docs:
        base = ntpath.basename(d.PathName.replace('/', '\\').rstrip('\\')) if d.PathName else ''
        if base and (base.lower() == low or ntpath.splitext(base)[0].lower() == low):
            return d, None
    partial = [d for d in docs if low in d.Title.lower()]
    if len(partial) == 1:
        return partial[0], None
    if partial:
        return None, {'error': 'Ambiguous document "{}": several open documents match.'.format(q),
                      'candidates': [d.Title for d in partial]}
    return None, None


def activate(uiapp, doc):
    """Bring `doc`'s window to the front. Returns (True, kind) or (False, error text)."""
    if is_cloud(doc):
        from Autodesk.Revit.DB import OpenOptions
        uiapp.OpenAndActivateDocument(doc.GetCloudModelPath(), OpenOptions(), False)
        return True, 'cloud'
    path = getattr(doc, 'PathName', None)
    if not path:
        return False, UNSAVED.format(doc.Title)
    uiapp.OpenAndActivateDocument(path)
    return True, 'file'
