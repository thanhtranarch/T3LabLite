#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""lite_guard/guard.py itself -- no Revit needed.

    python3 lite_guard/test_guard.py

Covers marker matching on CRLF (Windows checkout, core.autocrlf) and LF
(index, Linux CI) copies of the same file.

Plain CPython 3, standard library only.
"""
import os
import sys
import unittest

sys.dont_write_bytecode = True
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import guard  # noqa: E402


class _Tree(object):

    def __init__(self, files):
        self._files = files

    def files(self):
        return set(self._files)

    def read(self, path):
        return self._files.get(path)


class MarkerTests(unittest.TestCase):

    MARKER = {'file': 'skill.md', 'pattern': r'^name: drawing-rules$',
              'why': 'test'}

    def _errors(self, raw):
        report = guard.Report()
        guard.check_markers(_Tree({'skill.md': raw}),
                            {'markers': [self.MARKER]}, report)
        return report.errors

    def test_lf(self):
        self.assertEqual(self._errors(b'---\nname: drawing-rules\n---\n'), [])

    def test_crlf(self):
        self.assertEqual(
            self._errors(b'---\r\nname: drawing-rules\r\n---\r\n'), [])

    def test_lost_marker_still_fails(self):
        self.assertEqual(len(self._errors(b'---\r\nname: other\r\n---\r\n')), 1)

    def test_missing_file_still_fails(self):
        self.assertEqual(len(self._errors(None)), 1)


if __name__ == '__main__':
    unittest.main()
