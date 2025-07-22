# -*- coding: utf-8 -*-
"""
logging_util
============
Version   : 2.0
Date      : 2025-05-15
Author    : AO

A tiny helper that gives each script a **TimedRotatingFileHandler** pointing to
`%APPDATA%\MyPyRevitExtension\logs`.
Every log entry:

* is UTF-8
* includes the calling script name in the `src` column
* rolls over at midnight and keeps 90 days of history
* prepends the current project-exceptions JSON path the first time the log
  file is created.

IronPython 2.7 / Revit 2024 compatible.
"""

from __future__ import unicode_literals

import os
import logging
import atexit
import datetime
from logging.handlers import TimedRotatingFileHandler

# -----------------------------------------------------------------------------
# Constants
# -----------------------------------------------------------------------------
APPDATA  = os.environ.get('APPDATA', '')
LOG_DIR  = r"I:\BLU - Service Delivery\04 Building Information Management\07 The Button\logs"
VERSION  = '2.0'

# Ensure log directory exists
if not os.path.isdir(LOG_DIR):
    try:
        os.makedirs(LOG_DIR)
    except OSError:
        pass

# -----------------------------------------------------------------------------
# Helper classes / functions
# -----------------------------------------------------------------------------
class _SrcFilter(logging.Filter):
    """Inject a default `src` attribute if the record does not have one."""

    def __init__(self, default_src):
        super(_SrcFilter, self).__init__()
        self.default_src = default_src

    def filter(self, record):
        if not hasattr(record, 'src'):
            record.src = self.default_src
        return True


def get_logger(name, filename_override=None):
    """
    Return an initialised :pyclass:`logging.Logger`.

    Multiple calls with the same *name* return the already-configured instance
    (“init-once” behaviour).
    """
    logger = logging.getLogger(name)
    if getattr(logger, '_initialized', False):
        return logger

    # -----------------------------------------------------------------
    # Work out a human-readable model name (or fallback)
    # -----------------------------------------------------------------
    model = 'unsaved_doc'
    try:
        doc  = __revit__.ActiveUIDocument.Document           # noqa: F821 (pyRevit)
        path = doc.PathName
        if path:
            model = os.path.splitext(os.path.basename(path))[0]
    except Exception:
        pass

    # -----------------------------------------------------------------
    # Build log-file name:  <base>_<model>_v<ver>_<YYYY_MM>.log
    # -----------------------------------------------------------------
    ym        = datetime.datetime.now().strftime('%Y_%m')
    base      = filename_override or name
    log_fname = '{0}_{1}_v{2}_{3}.log'.format(base, model, VERSION, ym)
    log_path  = os.path.join(LOG_DIR, log_fname)

    # Write a one-off header with the project exception JSON path
    if not os.path.isfile(log_path):
        try:
            from exception_manager import ExceptionManager
            proj = ExceptionManager._get_project_path()
            with open(log_path, 'a') as fp:
                fp.write('project_exceptions_path: {0}\n'.format(proj))
        except Exception:
            pass

    # -----------------------------------------------------------------
    # Handler / formatter
    # -----------------------------------------------------------------
    handler = TimedRotatingFileHandler(
        log_path, when='midnight', interval=1,
        backupCount=90, encoding='utf-8'
    )

    fmt = '%(asctime)s [v' + VERSION + '] %(src)-14s %(levelname)s: %(message)s'
    handler.setFormatter(logging.Formatter(fmt))

    # -----------------------------------------------------------------
    # Logger configuration (once only)
    # -----------------------------------------------------------------
    logger.setLevel(logging.INFO)
    logger.addHandler(handler)
    logger.addFilter(_SrcFilter(name))
    logger.propagate = False     # prevent double-logging to root

    atexit.register(logging.shutdown)
    logger._initialized = True
    return logger
