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
