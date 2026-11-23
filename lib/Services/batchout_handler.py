# -*- coding: utf-8 -*-
"""
BatchOut External Event Handler
===============================
Provides an IExternalEventHandler implementation for BatchOut (ExportManagerWindow).
Satisfies IExternalEventHandler on both CPython 3 and IronPython 2.

The .NET type name gets a per-load suffix (clr_namespace, rule S15). A fixed name only
works while this module stays in sys.modules; once it is imported a second
time in the same Revit session, the .NET type from the first import still
exists and pythonnet raises "Duplicate type name within an assembly". That
happens when a run drops the module, the engine restarts, or the
BatchOutDialog fallback defines the handler again. BatchOut then failed on
every click until Revit was restarted.
"""

import sys
import clr
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.UI import IExternalEventHandler
from Snippets._compat import clr_namespace   # per-load .NET type name (CLAUDE.md rule 5)


class BatchOutEventHandler(IExternalEventHandler):
    __namespace__ = clr_namespace('BatchOut')
    """Runs queued window actions inside a valid Revit API context.

    Required because the window is modeless (Show, not ShowDialog): ALL Revit
    API access — reads included, not just transactions — must happen in here.
    Touching the API from WPF events or Dispatcher callbacks after the command
    has returned runs outside API context and hard-crashes Revit on workshared
    models.
    """

    def __init__(self, logger):
        self._logger = logger
        self._queue = []

    def add(self, action):
        self._queue.append(action)

    def remove(self, action):
        """Discard only the request whose Raise failed."""
        for index in range(len(self._queue) - 1, -1, -1):
            if self._queue[index] is action:
                del self._queue[index]
                break

    def clear(self):
        self._queue = []

    def Execute(self, uiapp):
        try:
            actions = self._queue
            self._queue = []
            for action in actions:
                try:
                    action()
                except Exception as ex:
                    try:
                        self._logger.error(
                            "BatchOut external event failed: {}".format(ex))
                    except Exception:
                        pass
        except Exception:
            pass

    def GetName(self):
        return "T3Lab BatchOut Handler"
