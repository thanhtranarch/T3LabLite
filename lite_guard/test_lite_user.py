# -*- coding: utf-8 -*-
"""T3Lab Lite: the Assistant's user name (lib/config/lite_user.py).

Run: python3 lite_guard/test_lite_user.py   Exit 0 = all pass.
"""
from __future__ import unicode_literals

import io
import json
import os
import re
import shutil
import sys
import tempfile
import unittest

LIB = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'lib')
sys.path.insert(0, LIB)

from config import lite_user  # noqa: E402


def _write(name, data):
    with io.open(os.path.join(lite_user._dir(), name), 'w', encoding='utf-8') as f:
        f.write(json.dumps(data, ensure_ascii=False))


class LiteUserTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix='t3lite_user_')
        self.old = os.environ.get('APPDATA')
        os.environ['APPDATA'] = self.tmp
        os.makedirs(lite_user._dir())

    def tearDown(self):
        if self.old is None:
            os.environ.pop('APPDATA', None)
        else:
            os.environ['APPDATA'] = self.old
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_fresh_install_has_no_name_and_asks(self):
        self.assertEqual(lite_user.get_name(), '')
        self.assertTrue(lite_user.needs_name())

    def test_shipped_default_is_never_a_name(self):
        _write('user_profile.json', {'user': {'name': 'Thạnh'}})
        _write('settings.json', {'username': 'Thạnh'})
        self.assertEqual(lite_user.get_name(), '')
        self.assertTrue(lite_user.needs_name())
        self.assertEqual(lite_user.set_name('Thạnh'), '')

    def test_name_from_an_earlier_install_is_kept(self):
        _write('user_profile.json', {'user': {'name': 'Minh'}})
        self.assertEqual(lite_user.get_name(), 'Minh')
        self.assertFalse(lite_user.needs_name())
        _write('user_profile.json', {'user': {'name': 'Someone else'}})
        self.assertEqual(lite_user.get_name(), 'Minh', 'copied once, then its own file')

    def test_settings_username_is_the_second_source(self):
        _write('settings.json', {'username': 'Lan'})
        self.assertEqual(lite_user.get_name(), 'Lan')

    def test_typed_name_goes_to_its_own_file_only(self):
        _write('user_profile.json', {'user': {'name': 'Owner'}})
        self.assertEqual(lite_user.set_name('  Hoa  '), 'Hoa')
        self.assertEqual(lite_user.get_name(), 'Hoa')
        with io.open(os.path.join(lite_user._dir(), 'user_profile.json'), encoding='utf-8') as f:
            self.assertEqual(json.load(f)['user']['name'], 'Owner', 'T3Lab AI data untouched')
        self.assertTrue(lite_user.path().endswith('t3lablite_user.json'))

    def test_cancel_asks_only_once(self):
        lite_user.mark_asked()
        self.assertFalse(lite_user.needs_name())
        self.assertEqual(lite_user.get_name(), '')


class WiringTests(unittest.TestCase):
    def read(self, rel):
        with io.open(os.path.join(LIB, rel), encoding='utf-8') as f:
            return f.read()

    def test_no_built_in_name_left(self):
        for rel in ('config/user_profile.py', 'config/settings.py',
                    'GUI/T3LabAssistantDialog.py', 'GUI/LLMSettingDialog.py'):
            code = '\n'.join(l for l in self.read(rel).splitlines()
                             if not l.strip().startswith('#'))
            self.assertIsNone(re.search(r"""['"]Thạnh['"]""", code), rel)

    def test_profile_delegates_to_lite_user(self):
        src = self.read('config/user_profile.py')
        self.assertIn('return lite_user.get_name()', src)
        self.assertIn('return lite_user.set_name(name)', src)

    def test_assistant_asks_once_before_opening(self):
        src = self.read('GUI/T3LabAssistantDialog.py')
        body = src.split('def show_assistant_dialog():', 1)[1]
        self.assertLess(body.index('_ask_name_once()'), body.index('show_docked('))

    def test_settings_box_is_not_prefilled_with_a_fake_name(self):
        self.assertIn('get_name(fallback=False)', self.read('GUI/LLMSettingDialog.py'))


if __name__ == '__main__':
    unittest.main()
