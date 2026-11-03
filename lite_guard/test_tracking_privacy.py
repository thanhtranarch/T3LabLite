#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Usage tracking privacy: consent first, minimum data (lib/tracking/).

    python3 lite_guard/test_tracking_privacy.py

Checks what docs/compliance/2026-11-02-tracking-pdpl.md promises:
  - nothing is sent without an explicit, current "Allow" (Luat 91/2025/QH15
    Art. 9: silence is not consent); closing the dialog stores nothing;
    "tracking_enabled": false always wins; bumping CONSENT_VERSION asks again
  - the user ID is random, not built from the Windows user / computer name
  - the model title never leaves the machine (only a salted code)
  - free text, personal fields, e-mails, phone numbers and user folders are
    not sent, also inside nested arguments and error messages

Revit is not available here: Idling is a fake event, the dialog a stub and
sending is captured instead of posted. Plain CPython 3, standard library.
"""
import json
import os
import shutil
import socket
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True

_TMP = tempfile.mkdtemp(prefix='t3lab_tracking_test_')
os.environ['APPDATA'] = _TMP

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB = os.path.join(ROOT, 'lib')
if LIB not in sys.path:
    sys.path.insert(0, LIB)

from core import paths                                   # noqa: E402
from tracking import consent, hooks                      # noqa: E402
from tracking.service import (TelemetryService, redact_arguments,  # noqa: E402
                              scrub_text, INSTALL_ID_KEY)


def _read(*parts):
    with open(os.path.join(LIB, *parts), encoding='utf-8') as fh:
        return fh.read()


class FakeEvent(object):
    def __init__(self):
        self.handlers = []

    def __iadd__(self, fn):
        self.handlers.append(fn)
        return self

    def __isub__(self, fn):
        self.handlers.remove(fn)
        return self

    def fire(self):
        for fn in list(self.handlers):
            fn(None, None)


class FakeApp(object):
    def __init__(self):
        self.Idling = FakeEvent()


class Base(unittest.TestCase):
    def setUp(self):
        paths.save_settings({})
        TelemetryService._cached_user_hash = None
        self.sent = []
        self._orig_dispatch = TelemetryService._dispatch_async
        TelemetryService._dispatch_async = classmethod(
            lambda cls, payload: self.sent.append(payload))
        self._orig_show = consent.show_dialog

    def tearDown(self):
        TelemetryService._dispatch_async = self._orig_dispatch
        consent.show_dialog = self._orig_show

    def allow(self):
        consent.record_answer(True, via='test')

    def start_revit(self, answer):
        """startup.py -> track_session_start(handle); then the first Idling."""
        app = FakeApp()
        consent.show_dialog = lambda: answer
        hooks.track_session_start(app)
        app.Idling.fire()
        return app


class TestConsent(Base):
    def test_nothing_sent_without_an_answer(self):
        TelemetryService.record_tool_usage('ManaViews')
        TelemetryService.record_mcp_call('get_views', {}, {}, 1)
        self.assertEqual(self.sent, [])
        self.assertIsNone(consent.current_answer())

    def test_closing_the_dialog_stores_nothing_and_sends_nothing(self):
        self.start_revit(consent.NO_ANSWER)
        self.assertEqual(self.sent, [])
        self.assertNotIn(consent.CONSENT_KEY, paths.load_settings())

    def test_allow_is_recorded_with_time_and_version_then_sends(self):
        self.start_revit(consent.ALLOW)
        record = paths.load_settings()[consent.CONSENT_KEY]
        self.assertIs(record['value'], True)
        self.assertEqual(record['version'], consent.CONSENT_VERSION)
        self.assertRegex(record['at'], r'^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\dZ$')
        self.assertEqual([p['tool_name'] for p in self.sent], ['session_start'])

    def test_deny_never_sends_and_is_not_asked_again(self):
        self.start_revit(consent.DENY)
        self.assertEqual(self.sent, [])
        asked = []
        consent.show_dialog = lambda: asked.append(1)
        app = FakeApp()
        hooks.track_session_start(app)
        app.Idling.fire()
        self.assertEqual(asked, [])
        TelemetryService.record_tool_usage('ManaViews')
        self.assertEqual(self.sent, [])

    def test_enter_key_is_dont_allow(self):
        src = _read('tracking', 'consent.py')
        self.assertIn('dialog.DefaultButton = TaskDialogResult.CommandLink2', src)

    def test_old_switch_off_always_wins(self):
        self.allow()
        paths.set_setting('tracking_enabled', False)
        TelemetryService.record_tool_usage('ManaViews')
        self.assertEqual(self.sent, [])

    def test_old_switch_on_is_not_consent(self):
        paths.set_setting('tracking_enabled', True)
        TelemetryService.record_tool_usage('ManaViews')
        self.assertEqual(self.sent, [])

    def test_new_consent_version_asks_again(self):
        paths.set_setting(consent.CONSENT_KEY, {
            'value': True, 'version': consent.CONSENT_VERSION - 1})
        self.assertIsNone(consent.current_answer())
        paths.set_setting(consent.CONSENT_KEY, {
            'value': False, 'version': consent.CONSENT_VERSION - 1})
        self.assertIs(consent.current_answer(), False)

    def test_reload_before_idling_asks_once(self):
        asked = []
        consent.show_dialog = lambda: asked.append(1) or consent.ALLOW
        app = FakeApp()
        hooks.track_session_start(app)
        hooks.track_session_start(app)      # pyRevit Reload
        app.Idling.fire()
        self.assertEqual(asked, [1])


class TestMinimumData(Base):
    def setUp(self):
        Base.setUp(self)
        self.allow()

    def test_user_id_is_random_not_from_machine(self):
        uid = TelemetryService.get_user_hash()
        self.assertRegex(uid, r'^[0-9a-f]{16}$')
        self.assertEqual(paths.load_settings()[INSTALL_ID_KEY], uid)
        import hashlib
        user = os.environ.get('USERNAME') or os.environ.get('USER') or ''
        old = hashlib.sha256('{}_{}'.format(socket.gethostname(), user)
                             .encode('utf-8')).hexdigest()[:16]
        self.assertNotEqual(uid, old)
        TelemetryService._cached_user_hash = None
        self.assertEqual(TelemetryService.get_user_hash(), uid)   # stable

    def test_model_title_never_sent(self):
        code = TelemetryService.doc_code('ClientX_Tower_Central_nguyenvana')
        self.assertRegex(code, r'^doc_[0-9a-f]{12}$')
        self.assertEqual(code, TelemetryService.doc_code('ClientX_Tower_Central_nguyenvana'))
        self.assertEqual(TelemetryService.doc_code(''), '')
        src = _read('tracking', 'service.py')
        self.assertIn('payload["doc_name"] = cls.doc_code(doc_title)', src)

    def test_free_text_and_personal_fields_not_sent(self):
        out = redact_arguments({
            'description': 'Rename sheets for Nguyen Van A, a.nguyen@client.vn',
            'query': 'walls of Mr Binh',
            'prompt': 'x',
            'owner_email': 'a@b.com',
            'user_name': 'nguyenvana',
            'phone': '0912345678',
            'view_name': 'Level 1',
            'element_ids': [101, 102],
        })
        self.assertEqual(out['description'], '(text not sent, 50 chars)')
        self.assertEqual(out['query'], '(text not sent, 16 chars)')
        self.assertEqual(out['owner_email'], '(not sent)')
        self.assertEqual(out['user_name'], '(not sent)')
        self.assertEqual(out['phone'], '(not sent)')
        self.assertEqual(out['view_name'], 'Level 1')
        self.assertEqual(out['element_ids'], [101, 102])

    def test_scrubbed_inside_other_values_and_nested(self):
        out = redact_arguments({
            'value': 'Checked by a.nguyen@client.vn, tel 0912 345 678',
            'options': {'note': 'secret note', 'mark': 'C:\\Users\\nguyenvana\\x'},
            'steps': [{'name': 'x', 'arguments': {'comment': 'hi', 'value': 'b@c.io'}}],
        })
        self.assertEqual(out['value'], 'Checked by (email), tel (phone)')
        self.assertEqual(out['options']['note'], '(text not sent, 11 chars)')
        self.assertEqual(out['options']['mark'], 'C:\\Users\\(user)\\x')
        self.assertEqual(out['steps'][0]['arguments'],
                         {'comment': '(text not sent, 2 chars)', 'value': '(email)'})

    def test_error_message_scrubbed(self):
        TelemetryService.record_mcp_call('export_sheets_pdf', {}, {
            'isError': True,
            'error': "Cannot write C:\\Users\\nguyenvana\\Desktop\\A.pdf for /home/binh/x"}, 3)
        msg = self.sent[0]['error_message']
        self.assertNotIn('nguyenvana', msg)
        self.assertNotIn('binh', msg)
        TelemetryService.record_tool_usage('X', status='error',
                                           error_message='mail me: x@y.vn')
        self.assertEqual(self.sent[1]['error_message'], 'mail me: (email)')

    def test_purpose_has_no_free_text(self):
        TelemetryService.record_mcp_call('send_code_to_revit', {
            'code': 'SECRET = 1', 'description': 'Delete rooms of Client Y'}, {}, 2)
        TelemetryService.record_mcp_call('rename_views', {
            'goal': 'Rename for Mr Binh'}, {}, 2)
        dumped = json.dumps(self.sent)
        for marker in ('SECRET', 'Client Y', 'Binh'):
            self.assertNotIn(marker, dumped)

    def test_numbers_and_short_ids_survive(self):
        self.assertEqual(scrub_text('Level 1 at 3000 mm, id 123456'),
                         'Level 1 at 3000 mm, id 123456')
        self.assertEqual(len(scrub_text('a' * 1000)), 300)


def tearDownModule():
    shutil.rmtree(_TMP, ignore_errors=True)


if __name__ == '__main__':
    unittest.main(verbosity=1)
