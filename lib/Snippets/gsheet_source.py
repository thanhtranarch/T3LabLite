# -*- coding: utf-8 -*-
"""
gsheet_source — use a Google Sheet as the workbook of an Excel table.

Google exports any spreadsheet as .xlsx with its formatting (fills, fonts,
borders, merged cells, column widths), so a Google Sheet goes through exactly
the same reader as a local workbook: this module only turns "something the
user pasted" into a downloaded .xlsx in a per-user cache.

Accepted sources:
  https://docs.google.com/spreadsheets/d/<id>/edit#gid=0     normal link
  https://docs.google.com/spreadsheets/d/e/<pub-id>/pubhtml   File > Share >
                                                             Publish to web
  <id>                                                       bare id (44 ch)
  X:\\My Drive\\Doors.gsheet                                  Google Drive for
                                                             desktop shortcut

No sign-in: the sheet must be shared "Anyone with the link — Viewer" or
published to the web. A private sheet answers with Google's sign-in page;
that is reported with the two ways out (share it, or File > Download >
Microsoft Excel and pick the .xlsx).

Download: urllib first (CPython std lib); if that fails before any HTTP answer
(TLS / proxy inside Revit), .NET WebClient, which uses the Windows proxy and
certificate store. Tested by dev/test_excel_table.py (no network needed).

Author: Tran Tien Thanh
"""

import json
import os
import re
import tempfile

_ID_RE = re.compile(r'/spreadsheets/d/([A-Za-z0-9_-]{20,})')
_PUB_RE = re.compile(r'/spreadsheets/d/e/([A-Za-z0-9_-]{20,})')
_BARE_RE = re.compile(r'^[A-Za-z0-9_-]{40,}$')
_GID_RE = re.compile(r'[#&?]gid=(\d+)')

TIMEOUT = 30


class GSheetError(Exception):
    """User-facing: what went wrong and what to do."""


NOT_SHARED = ("Google answered with its sign-in page: the sheet is private. In Google "
              "Sheets use Share > General access > 'Anyone with the link' (Viewer), or "
              "File > Download > Microsoft Excel (.xlsx) and pick that file instead.")


def _read_gsheet_file(path):
    """Google Drive for desktop '.gsheet' shortcut -> its URL / doc id."""
    try:
        with open(path, encoding='utf-8') as fh:
            data = json.load(fh)
    except (IOError, OSError, ValueError):
        raise GSheetError("Cannot read the Google Drive shortcut {}.".format(
            os.path.basename(path)))
    return data.get('url') or data.get('doc_id') or ''


def parse(text):
    """Pasted text / .gsheet path -> {'id', 'published', 'gid', 'url'} or None."""
    text = (text or '').strip().strip('"')
    if not text:
        return None
    if text.lower().endswith('.gsheet') and os.path.isfile(text):
        text = _read_gsheet_file(text)
    gid = _GID_RE.search(text)
    gid = gid.group(1) if gid else None
    m = _PUB_RE.search(text)
    if m:
        return {'id': m.group(1), 'published': True, 'gid': gid,
                'url': 'https://docs.google.com/spreadsheets/d/e/{}/pubhtml'.format(m.group(1))}
    m = _ID_RE.search(text)
    if m:
        sid = m.group(1)
    elif _BARE_RE.match(text):
        sid = text
    else:
        return None
    return {'id': sid, 'published': False, 'gid': gid,
            'url': 'https://docs.google.com/spreadsheets/d/{}/edit'.format(sid)}


def is_google(text):
    try:
        return parse(text) is not None
    except GSheetError:
        return True                    # a .gsheet file, even an unreadable one


def export_url(src):
    if src['published']:
        return 'https://docs.google.com/spreadsheets/d/e/{}/pub?output=xlsx'.format(src['id'])
    return 'https://docs.google.com/spreadsheets/d/{}/export?format=xlsx'.format(src['id'])


def cache_path(src, folder=None):
    if folder is None:
        base = os.environ.get('APPDATA') or tempfile.gettempdir()
        folder = os.path.join(base, 'T3LabAI', 'gsheets')
    if not os.path.isdir(folder):
        os.makedirs(folder)
    return os.path.join(folder, '{}.xlsx'.format(src['id']))


def check_payload(data):
    """Raise GSheetError unless `data` is an .xlsx (zip) file."""
    if data[:2] == b'PK':
        return
    head = data[:2048].lower()
    if b'<html' in head or b'<!doctype' in head:
        raise GSheetError(NOT_SHARED)
    raise GSheetError("Google did not send an Excel file ({} bytes). Check the link.".format(
        len(data)))


def _fetch_urllib(url):
    import urllib.request
    import urllib.error
    req = urllib.request.Request(url, headers={'User-Agent': 'T3Lab-ManaSched'})
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            return resp.read()
    except urllib.error.HTTPError as ex:
        if ex.code in (401, 403):
            raise GSheetError(NOT_SHARED)
        if ex.code == 404:
            raise GSheetError("Google cannot find that sheet (404). Check the link.")
        raise GSheetError("Google answered HTTP {} for the export.".format(ex.code))


def _fetch_dotnet(url):
    import clr
    clr.AddReference('System')
    from System.Net import WebClient, ServicePointManager, SecurityProtocolType
    try:
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
    except Exception:
        pass
    client = WebClient()
    client.Headers.Add('User-Agent', 'T3Lab-ManaSched')
    return bytes(bytearray(client.DownloadData(url)))


def download(text, folder=None, fetch=None):
    """Download the sheet as .xlsx; returns (local path, parsed source).

    `fetch(url) -> bytes` is injectable for tests.
    """
    src = parse(text)
    if src is None:
        raise GSheetError("'{}' is not a Google Sheets link.".format(text))
    url = export_url(src)
    if fetch is not None:
        data = fetch(url)
    else:
        try:
            data = _fetch_urllib(url)
        except GSheetError:
            raise
        except Exception as first:
            try:
                data = _fetch_dotnet(url)
            except Exception as second:
                raise GSheetError("Cannot reach Google Sheets ({}; {}). Check the internet "
                                  "connection or proxy.".format(first, second))
    check_payload(data)
    path = cache_path(src, folder)
    tmp = path + '.part'
    with open(tmp, 'wb') as fh:
        fh.write(data)
    os.replace(tmp, path)
    return path, src
