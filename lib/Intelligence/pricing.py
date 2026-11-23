# -*- coding: utf-8 -*-
"""
USD estimate of one model call from its token usage.

Prices are Anthropic first-party list prices per million tokens
(input, output, cache read, cache write) as of 2026-11 — checked against
the Claude API pricing reference at that date. They are keyed by MODEL FAMILY SUBSTRING
(the provider resolves model ids live from /v1/models, so there is no
hardcoded id anywhere else to key on); the first matching row wins, most
specific first. An unknown model, or a local one, estimates to None —
never to 0, which would read as "free".

Known simplification: Claude Haiku 5.5 bills prompts over 100K tokens at a
higher tier; that tier is ignored here (the agent's prompts are far below).

Author: Tran Tien Thanh
"""
from __future__ import unicode_literals

__author__ = "Tran Tien Thanh"
__title__ = "Model Pricing"

# (family substring, input, output, cache_read, cache_write) — USD per MTok.
# Cache write = 125% of input (5-minute TTL). Cache read is 10% of input
# except where a row says otherwise.
PRICES = (
    ("fable-5-2",   8.00, 40.00, 0.20, 10.00),   # cache read 0.025x
    ("mythos-5-2",  8.00, 40.00, 0.20, 10.00),
    ("fable-5-1",  10.00, 50.00, 0.25, 12.50),   # cache read 0.025x
    ("mythos-5-1", 10.00, 50.00, 0.25, 12.50),
    ("fable-5",    10.00, 50.00, 1.00, 12.50),
    ("mythos-5",   10.00, 50.00, 1.00, 12.50),
    ("opus-5-5",    4.00, 20.00, 0.20,  5.00),   # cache read 0.05x
    ("opus-5",      5.00, 25.00, 0.50,  6.25),
    ("opus-4",      5.00, 25.00, 0.50,  6.25),   # 4.6 / 4.7 / 4.8 / 4.5
    ("sonnet-5-5",  2.00, 10.00, 0.10,  2.50),   # cache read 0.05x
    ("sonnet-5",    2.00, 10.00, 0.20,  2.50),
    ("sonnet-4",    3.00, 15.00, 0.30,  3.75),   # 4.6 / 4.5
    ("haiku-5-5",   0.10,  0.50, 0.01,  0.125),
    ("haiku-4",     1.00,  5.00, 0.10,  1.25),
)


def price_row(model):
    """The PRICES row for `model`, or None when the model is unknown."""
    m = u"{}".format(model or u"").lower()
    if not m:
        return None
    for row in PRICES:
        if row[0] in m:
            return row
    return None


def estimate_usd(model, in_tok=0, out_tok=0, cache_read=0, cache_write=0):
    """USD for one call (or one turn's accumulated counts). None if unknown.

    `in_tok` is the UNCACHED input count the API reports as input_tokens;
    cached reads and writes are priced on their own rows.
    """
    row = price_row(model)
    if row is None:
        return None
    try:
        usd = (float(in_tok or 0) * row[1] + float(out_tok or 0) * row[2]
               + float(cache_read or 0) * row[3]
               + float(cache_write or 0) * row[4]) / 1e6
    except Exception:
        return None
    return round(usd, 6)


def format_usd(usd):
    """'$0.0123' style, '–' for None."""
    if usd is None:
        return u"–"
    if usd < 0.01:
        return u"${:.4f}".format(usd)
    return u"${:.2f}".format(usd)
