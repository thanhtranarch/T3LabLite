# -*- coding: utf-8 -*-
"""BatchOut DWG export: find and delete the image reference files Revit writes
next to a DWG (title-block logos etc. — DWG cannot embed rasters).

Revit-free; tested by `lite_guard/test_batchout_refs.py`.
"""

import os

# What Revit writes for a raster / PDF image placed in a view or title block.
REFERENCE_EXTENSIONS = ('.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.gif', '.pdf')


def snapshot(folder):
    """{lowercase file name: mtime} of `folder` (non-recursive); {} if unreadable."""
    state = {}
    try:
        names = os.listdir(folder)
    except OSError:
        return state
    for name in names:
        path = os.path.join(folder, name)
        try:
            if os.path.isfile(path):
                state[name.lower()] = os.path.getmtime(path)
        except OSError:
            continue
    return state


def written_reference_files(folder, before, keep=()):
    """Paths of image/PDF files created or rewritten in `folder` since `before`
    (a `snapshot`). `keep` = paths that are real export outputs, never returned.
    Untouched files are never returned, so a PDF exported earlier is safe."""
    # snapshot() lower-cases names, so compare lower-case on every OS.
    kept = set(os.path.normcase(os.path.abspath(p)).lower() for p in keep if p)
    found = []
    after = snapshot(folder)
    for name in sorted(after):
        if os.path.splitext(name)[1] not in REFERENCE_EXTENSIONS:
            continue
        if name in before and before[name] == after[name]:
            continue
        path = os.path.join(folder, name)
        if os.path.normcase(os.path.abspath(path)).lower() in kept:
            continue
        found.append(path)
    return found


def delete_files(paths):
    """Delete `paths`; returns (deleted, failed) path lists."""
    deleted, failed = [], []
    for path in paths:
        try:
            os.remove(path)
            deleted.append(path)
        except OSError:
            failed.append(path)
    return deleted, failed
