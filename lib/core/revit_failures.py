# -*- coding: utf-8 -*-
"""Revit failures during a headless MCP tool: an error rolls back with no dialog,
warnings are dismissed, both are reported.

Several shared tool helpers force modal failure handling; called from the MCP
server, a modal failure dialog would hold Revit until someone clicks it. Hooked
around one spec write tool on Revit's main thread (core/server.py handler).
IronPython 2.7 + CPython 3; Revit types imported inside guard() so the session
is testable with fakes (lite_guard/test_bridge_contract.py).
"""
from __future__ import unicode_literals

import contextlib

TEXT_CHARS = 300
MAX_ITEMS = 50


class FailureSession(object):
    """What Revit reported during one call."""

    def __init__(self, rollback_result=None):
        self.rollback_result = rollback_result     # FailureProcessingResult.ProceedWithRollBack
        self.warnings = 0
        self.warning_texts = []
        self.errors = []

    def on_failures(self, sender, args):
        try:
            fa = args.GetFailuresAccessor()
            rolled = False
            for msg in list(fa.GetFailureMessages()):
                text = '{}'.format(msg.GetDescriptionText() or '')[:TEXT_CHARS]
                if '{}'.format(msg.GetSeverity()) == 'Warning':
                    fa.DeleteWarning(msg)
                    self.warnings += 1
                    if len(self.warning_texts) < MAX_ITEMS and text not in self.warning_texts:
                        self.warning_texts.append(text)
                    continue
                rolled = True
                if len(self.errors) < MAX_ITEMS:
                    self.errors.append(text)
            if rolled and self.rollback_result is not None:
                try:
                    opts = fa.GetFailureHandlingOptions()
                    opts.SetClearAfterRollback(True)       # no error dialog after it
                    fa.SetFailureHandlingOptions(opts)
                except Exception:
                    pass
                args.SetProcessingResult(self.rollback_result)
        except Exception:
            pass

    def report(self):
        out = {'warnings_dismissed': self.warnings}
        if self.warning_texts:
            out['warning_texts'] = self.warning_texts
        if self.errors:
            out['errors_rolled_back'] = self.errors
            out['hint'] = ('Revit refused these changes; the transaction was rolled back. '
                           'Change the approach and retry.')
        return out


def annotate(result, session):
    """Fold the session into a tool result (in place); a rollback is a failure."""
    if isinstance(result, dict) and (session.errors or session.warnings):
        result['revit_failures'] = session.report()
        if session.errors:
            result['success'] = False
            result.setdefault('error', 'Revit rolled the change back; see revit_failures.')
    return result


@contextlib.contextmanager
def guard(uiapp):
    """Hook Application.FailuresProcessing around one call."""
    from System import EventHandler
    from Autodesk.Revit.DB import FailureProcessingResult
    from Autodesk.Revit.DB.Events import FailuresProcessingEventArgs
    session = FailureSession(rollback_result=FailureProcessingResult.ProceedWithRollBack)
    on_fail = EventHandler[FailuresProcessingEventArgs](session.on_failures)
    app = uiapp.Application
    try:
        app.FailuresProcessing += on_fail
        yield session
    finally:
        try:
            app.FailuresProcessing -= on_fail
        except Exception:
            pass
