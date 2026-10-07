# -*- coding: utf-8 -*-
"""When the automatic update check is due -- pure logic, no .NET, no Revit.

``core/updater.py`` imports ``clr`` and ``System.Net`` and can only run inside
Revit. The schedule lives here instead so it runs under plain Python and is
covered by ``lite_guard/test_update_schedule.py``. It must stay importable by
IronPython 2.7 (``startup.py`` runs there): no f-strings, no annotations.

The check runs once per *period*: a week by default (ISO weeks, Monday to
Sunday), or a day. Two stamps live in ``%APPDATA%\\T3LabAI\\mcp_paths.json``:

    last_update_check    period of the last COMPLETED check, i.e. GitHub was
                         reached ('2026-W41' weekly, '2026-10-07' daily). Old
                         installs wrote a date here; it is read as the week or
                         day that date falls in.
    last_update_attempt  'YYYY-MM-DD' of the last attempt, whatever happened,
                         so an offline or blocked machine retries on the next
                         day's first Revit start, not on every start.

A week is only ticked off when the check got an answer from GitHub. A machine
that was offline on Monday therefore checks again on Tuesday instead of going
a whole week without one.
"""

from __future__ import unicode_literals

import re
from datetime import date

INTERVAL_KEY = 'auto_update_interval'
STAMP_KEY = 'last_update_check'
ATTEMPT_KEY = 'last_update_attempt'

WEEKLY = 'weekly'
DAILY = 'daily'
DEFAULT_INTERVAL = WEEKLY

_DAY_RE = re.compile(r'^(\d{4})-(\d{2})-(\d{2})$')
_WEEK_RE = re.compile(r'^\d{4}-W\d{2}$')


def normalize_interval(value):
    """'daily' stays daily; anything else (missing, typo, 'weekly') is weekly."""
    text = ('%s' % value).strip().lower() if value else ''
    return DAILY if text == DAILY else WEEKLY


def day_id(today):
    """'2026-10-07'."""
    return today.strftime('%Y-%m-%d')


def week_id(today):
    """ISO week of `today`: '2026-W41'. Weeks run Monday to Sunday.

    The ISO year can differ from the calendar year around New Year, so
    2027-01-01 is '2026-W53' and 2027-01-04 (the Monday after) is '2027-W01'.
    """
    iso = today.isocalendar()
    return '%04d-W%02d' % (iso[0], iso[1])


def period_id(today, interval):
    """The period `today` falls in, for `interval`."""
    return day_id(today) if normalize_interval(interval) == DAILY else week_id(today)


def stamp_period(stamp, interval):
    """The period a stored stamp belongs to, in units of `interval`.

    Returns '' when the stamp is empty or unreadable, or when it cannot say
    anything about the current unit (a week stamp does not name a day).
    """
    text = ('%s' % stamp).strip() if stamp else ''
    if _WEEK_RE.match(text):
        return text if normalize_interval(interval) == WEEKLY else ''
    match = _DAY_RE.match(text)
    if match:
        try:
            stamped = date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            return ''
        return period_id(stamped, interval)
    return ''


def is_due(settings, today):
    """True when the automatic check should run now.

    settings: the mcp_paths.json dict. today: a datetime.date.
    Due unless this period's check already completed, or an attempt was
    already made today.
    """
    interval = normalize_interval(settings.get(INTERVAL_KEY))
    done = stamp_period(settings.get(STAMP_KEY), interval)
    if done and done == period_id(today, interval):
        return False
    if settings.get(ATTEMPT_KEY) == day_id(today):
        return False
    return True
