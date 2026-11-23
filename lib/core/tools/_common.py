# -*- coding: utf-8 -*-
"""Helpers shared by core/tools/<domain>.py (not a domain: no SPECS)."""
from __future__ import unicode_literals

MAX_ROWS = 300          # rows per list in a tool result (result_budget pages the rest)


def norm(text):
    return ' '.join('{}'.format(text or '').split()).lower()


def names(value):
    if value is None:
        return []
    if not isinstance(value, (list, tuple)):
        value = [value]
    return ['{}'.format(v) for v in value if '{}'.format(v).strip()]


def int_ids(value):
    """(ids, bad) from ints or digit strings ("@set" is expanded by the server)."""
    if value is None:
        return [], []
    if not isinstance(value, (list, tuple)):
        value = [value]
    ids, bad = [], []
    for v in value:
        try:
            i = int('{}'.format(v).strip())
        except (TypeError, ValueError):
            bad.append('{}'.format(v))
            continue
        if i not in ids:
            ids.append(i)
    return ids, bad
