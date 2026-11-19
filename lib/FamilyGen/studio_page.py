# -*- coding: utf-8 -*-
"""FamiGen Studio page and pyRevit version gate (pure Python, no Revit).

``inline_page()`` joins ``studio/studio.html`` with ``studio.css``, three.js r158
(``studio/vendor/three.min.js``) and ``studio.js`` into ONE HTML string;
``build_page()`` writes it to ``%APPDATA%\\T3LabAI\\famigen_studio``. One file
means the page never depends on how WebView2 resolves relative file URLs, and
the same page runs in a test browser (``lite_guard/test_famigen_studio.py``).
"""
import base64
import io
import os
import traceback

from FamilyGen.studio_controller import MIN_PYREVIT

STUDIO_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'studio')
PAGE_NAME = 'famigen_studio.html'
# The same logo every T3 window shows in its title bar.
LOGO_PATH = os.path.join(os.path.dirname(STUDIO_DIR), os.pardir, 'GUI', 'T3Lab_logo.png')
LOGO_MARKER = '__T3_LOGO_DATA__'
MARKERS = (('/*__STUDIO_CSS__*/', ('studio.css',)),
           ('/*__THREE_JS__*/', ('vendor', 'three.min.js')),
           ('/*__STUDIO_JS__*/', ('studio.js',)))


# ── pyRevit version gate ─────────────────────────────────────────────────────

def pyrevit_version():
    """(major, minor, patch) of the running pyRevit, or None if unknown."""
    try:
        from pyrevit.versionmgr import get_pyrevit_version
        ver = get_pyrevit_version()
        return (int(ver.major), int(ver.minor), int(ver.patch))
    except Exception:
        return None


def version_text(ver):
    return u'.'.join(str(v) for v in ver) if ver else u'unknown'


def requirement(ver=None):
    """(ok, message): FamiGen Studio needs pyRevit 7.0 or later."""
    ver = ver if ver is not None else pyrevit_version()
    if ver is None:
        return False, (u'FamiGen Studio needs pyRevit 7.0 or later, and the pyRevit version '
                       u'could not be read. Update pyRevit from github.com/pyrevitlabs/pyRevit/'
                       u'releases, then restart Revit.')
    if tuple(ver[:2]) < MIN_PYREVIT:
        return False, (u'FamiGen Studio needs pyRevit 7.0 or later; this Revit runs pyRevit {}. '
                       u'pyRevit 7 renders tool pages with WebView2, which the Studio\'s 3D '
                       u'view needs. Update pyRevit from github.com/pyrevitlabs/pyRevit/releases, '
                       u'then restart Revit.'.format(version_text(ver)))
    return True, u''


MESSAGING_MESSAGE = (
    u'FamiGen Studio opened, but this pyRevit ({}) cannot pass messages between the page and '
    u'Revit (output.on_message, added in pyRevit 7.1), or the WebView2 runtime is missing. '
    u'Update pyRevit to 7.1 or later and install the Microsoft Edge WebView2 Runtime, then '
    u'restart Revit.')


class Outbox(object):
    """Posts to the page in order; holds them while the page cannot take them.

    pyRevit's ``post_message`` returns False until the page has finished
    loading (``IsWindowReady``), and the page says ``ready`` before that: the
    ``init`` answer (models, categories, AI state) was dropped, the badge
    stayed "AI …" and Send locked after the first reply (2026-11-19).
    ``wake`` is called when something is held, so the host retries soon.
    """
    LIMIT = 500

    def __init__(self, send, wake=None):
        self._send = send
        self._wake = wake
        self._queue = []

    @property
    def pending(self):
        return len(self._queue)

    def post(self, message):
        self._queue.append(message)
        if len(self._queue) > self.LIMIT:
            del self._queue[0]
        if not self.flush() and self._wake is not None:
            self._wake()

    def flush(self):
        """Send what is held, oldest first. True when nothing is left."""
        while self._queue:
            try:
                delivered = self._send(self._queue[0])
            except Exception:
                traceback.print_exc()      # a message that cannot be sent must not block the rest
                delivered = True
            if delivered is False:
                return False
            del self._queue[0]
        return True


# ── the page ─────────────────────────────────────────────────────────────────

def _read(*parts):
    with io.open(os.path.join(STUDIO_DIR, *parts), 'r', encoding='utf-8') as handle:
        return handle.read()


def page_folder():
    base = os.environ.get('APPDATA') or os.path.expanduser('~')
    return os.path.join(base, 'T3LabAI', 'famigen_studio')


def inline_page():
    """The Studio as one HTML string (CSS, three.js and the app script inlined)."""
    html = _read('studio.html')
    for marker, parts in MARKERS:
        text = _read(*parts)
        if marker not in html:
            raise RuntimeError('studio.html lost its {} marker'.format(marker))
        if '</script' in text.lower():
            raise RuntimeError('inlined text must not contain a closing script tag')
        html = html.replace(marker, text, 1)
    with io.open(os.path.normpath(LOGO_PATH), 'rb') as handle:
        logo = base64.b64encode(handle.read()).decode('ascii')
    if LOGO_MARKER not in html:
        raise RuntimeError('studio.html lost its logo marker')
    return html.replace(LOGO_MARKER, 'data:image/png;base64,' + logo, 1)


def build_page(folder=None):
    folder = folder or page_folder()
    if not os.path.isdir(folder):
        os.makedirs(folder)
    path = os.path.join(folder, PAGE_NAME)
    with io.open(path, 'w', encoding='utf-8') as handle:
        handle.write(inline_page())
    return path
