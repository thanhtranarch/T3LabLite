# -*- coding: utf-8 -*-
"""
Tools that died with a raw pyRevit traceback ("cannot import name 'revit'
from 'pyrevit'" after an update or pyRevit Reload without a Revit restart)
now load their dialog inside GUI.ErrorGuard.run_tool, which shows "restart
Revit" and logs to %APPDATA%\\T3LabAI\\errors.log.

Run:  python3 lite_guard/test_tool_error_guard.py   (exit 0 = all pass)
"""
import io
import os
import sys
import unittest

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(REPO, 'lib'))

GUARDED = {
    'Model Auditor': ('T3Lab.tab/Standards & Families.panel/Standards.stack/'
                      'ModelAuditor.pushbutton/script.py', 'GUI.ModelAuditorDialog'),
    'Mana Sheets': ('T3Lab.tab/Views & Sheets.panel/ManaSheets.pushbutton/script.py',
                    'GUI.ManaSheetsDialog'),
}


class GuardedScriptTests(unittest.TestCase):
    def test_dialog_import_runs_inside_run_tool(self):
        for tool, (rel, module) in GUARDED.items():
            with self.subTest(tool=tool):
                with io.open(os.path.join(REPO, rel), encoding='utf-8') as f:
                    src = f.read()
                self.assertTrue(src.startswith('#! python3'))
                self.assertIn('run_tool("{}", _main)'.format(tool), src)
                main_body = src[src.index('def _main():'):src.index("if __name__ == '__main__':")]
                self.assertIn('from {} import'.format(module), main_body)
                top = src[:src.index('def _main():')]
                self.assertNotIn('from {} import'.format(module), top)

    def test_pyrevit_import_error_says_run_the_tool_again(self):
        # A restart never cured it: pyrevit.revit had failed to load earlier in
        # the same click (lite_guard/test_pyrevit_revit_thread.py).
        from GUI import ErrorGuard
        step = ErrorGuard.next_step(
            "ImportError: cannot import name 'revit' from 'pyrevit' "
            "(C:\\pyRevit-Master\\pyrevitlib\\pyrevit\\__init__.py)")
        self.assertIn('Run the tool again', step)
        self.assertNotIn('restart', step.lower())


if __name__ == '__main__':
    unittest.main()
