# -*- coding: utf-8 -*-
"""
GUI.assistant — the T3Lab Assistant chat window, split by concern.

GUI/T3LabAssistantDialog.py keeps the WPF window class (it must inherit
ONLY from T3WPFWindow: PythonNet cannot mix a CLR base with a Python mixin)
and the module-level names other code imports. This package holds what was
split out of it:

  module-level helpers   theme, deps, text, lang, launchers, history,
                         constants
  window helpers         plain classes built with the window
                         (`self._x = Helper(self)`), holding a group of
                         its methods; all state stays on the window

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""

__author__  = "Tran Tien Thanh"
__title__   = "T3Lab Assistant package"
