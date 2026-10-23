# -*- coding: utf-8 -*-
"""
The logger every GUI.assistant module shares.

pyRevit names its logger after the running command, so this is the same
logger GUI.T3LabAssistantDialog gets from script.get_logger(). Plain CPython
(tests) gets a standard-library logger instead.

Author: Tran Tien Thanh
Mail: trantienthanh909@gmail.com
"""
from __future__ import unicode_literals

__author__  = "Tran Tien Thanh"
__title__   = "Assistant Logger"

try:
    from pyrevit import script as _script
    logger = _script.get_logger()
except Exception:                                    # plain CPython (tests)
    import logging
    logger = logging.getLogger("T3LabAssistant")
