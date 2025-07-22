# -*- coding: utf-8 -*-
from __future__ import unicode_literals  # Ensure all string literals are Unicode (IronPython 2.7)
__title__   = "Configure Exceptions"
__doc__     = """Version = 1.7.2
Date    = July 2025
========================================
Description:
Interactive UI tool for managing **project-specific exception**
strings used by the Change Register workflow. It lets you add,
edit or delete literals that must retain their exact capitalization
when applying sentence-case globally.
========================================
How-To:
1. Click the button on the ribbon.
2. Choose Add, Edit or Delete in the list picker.
3. Follow the prompts to modify your exception list.
4. Confirm changes in the popup messages.
5. Run "Convert to Lowercase" to apply the changes. 
========================================
TODO:
[FEATURE] – Bulk import/export of exception lists
========================================
Author: AO"""

"""
Configure Exceptions
===============
Version  : 1.7.2
Date     : 2025-05-15
Author   : AO

A lightweight UI helper for maintaining **project-specific exception
strings** used by the Change Register tool. It lets users add, edit or
remove literals that must keep their exact capitalisation/spelling when
sentence-case is applied.

Key features
------------
* Detects the correct *project* JSON path via ``ExceptionManager`` and
  edits it in-place.
* Presents a simple pyRevit list-picker workflow (Add / Edit / Delete).
* Saves atomically on NTFS (``System.IO.File.Replace``) and falls back to
  ``shutil.move`` where atomic replace is not supported.
* Writes every mutating action to the shared rotating log created by
  ``logging_util.get_logger`` (one file per user & model).
* IronPython 2.7 compatible – all string literals are Unicode by default.

This script modifies **data only**; no CAD model changes are made.

Compatible with IronPython 2.7 and Revit 2024.
"""



import datetime
import getpass
import json
import os
import shutil
import sys
from collections import OrderedDict  # noqa: F401 – retained for possible future use

import clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import (  # noqa: E402 (IronPython import style)
    DialogResult,
    MessageBox,
    MessageBoxButtons,
)
from System.IO import File, IOException

from pyrevit import forms
from exception_manager import ExceptionManager
from logging_util import get_logger

# ---------------------------------------------------------------------------
# Project‑relative paths & run‑time constants
# ---------------------------------------------------------------------------

SCRIPT_DIR   = os.path.dirname(__file__)
LIB_DIR      = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "..", "..", "..","..","lib"))
if LIB_DIR not in sys.path:
    sys.path.append(LIB_DIR)

PROJECT_JSON = ExceptionManager._get_project_path()
PROJECT_DIR  = os.path.dirname(PROJECT_JSON)
MODEL_NAME   = os.path.splitext(os.path.basename(PROJECT_JSON))[0].replace("_exceptions", "")
USER         = getpass.getuser()
VERSION      = "1.7.2"

# initialise logger (shared with other scripts in the extension) -------------
logger = get_logger("ExceptionsManager", filename_override="ExceptionsManager")
SRC    = "ManageExUI"

# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def _load_json(path):
    """Return ``list`` loaded from *path* or an empty list on failure."""
    try:
        with open(path, "r") as fp:
            data = json.load(fp)
            if isinstance(data, list):
                return data
            raise ValueError("JSON root node is not a list")
    except Exception as exc:  # noqa: B902 – broad on purpose, UI feedback required
        logger.error("load_failed path=%s error=%s", path, exc, extra={"src": SRC, "version": VERSION, "user": USER})
        MessageBox.Show(
            "Failed to load exceptions file:\n{0}\n\nDefaulting to an empty list.".format(exc),
            "Manage Exceptions – Warning",
            MessageBoxButtons.OK,
        )
        return []


def _atomic_replace(src_tmp, dst):
    """Replace *dst* by *src_tmp* atomically if possible, else fall back."""
    try:
        File.Replace(src_tmp, dst, None, True)  # NTFS‑only, fast & atomic
    except IOException:
        shutil.move(src_tmp, dst)  # DFS / non‑NTFS fallback – not atomic but safe enough


def _save_json(path, flat_list, log_msg=None):
    """Write *flat_list* to *path* atomically and log *log_msg* if given."""
    try:
        tmp_path = path + ".tmp"
        with open(tmp_path, "w") as tmp_fp:
            json.dump(flat_list, tmp_fp, indent=2)
        _atomic_replace(tmp_path, path)

        if log_msg:
            logger.info(log_msg, extra={"src": SRC, "version": VERSION, "user": USER})
            for handler in logger.handlers:
                handler.flush()

    except Exception as exc:  # noqa: B902 – broad on purpose
        logger.error("save_failed path=%s error=%s", path, exc, extra={"src": SRC, "version": VERSION, "user": USER})
        MessageBox.Show(
            "Failed to save project exceptions:\n{0}".format(exc),
            "Manage Exceptions – Error",
            MessageBoxButtons.OK,
        )

# ---------------------------------------------------------------------------
# Core UI routine
# ---------------------------------------------------------------------------

def manage_exceptions():
    """pyRevit list‑picker loop for adding / editing / deleting exceptions."""
    ExceptionManager.clear_cache()
    title = "Project Exceptions"

    actions = OrderedDict([
        ("Add Exception", "add"),
        ("Edit Exception", "edit"),
        ("Delete Exception", "delete"),
        ("Exit", None),
    ])

    while True:
        choice = forms.SelectFromList.show(list(actions.keys()), title=title, multiselect=False)
        if not choice or actions[choice] is None:
            break  # Exit selected or dialog cancelled

        flat_list = _load_json(PROJECT_JSON)
        action = actions[choice]

        # ------------------------------------------------------------------
        # ADD
        # ------------------------------------------------------------------
        if action == "add":
            new_exc = forms.ask_for_string("Enter new exception text:")
            if not new_exc:
                continue

            if any(e.lower() == new_exc.lower() for e in flat_list):
                MessageBox.Show("\"{0}\" already exists.".format(new_exc), title, MessageBoxButtons.OK)
                continue

            flat_list.append(new_exc)
            _save_json(PROJECT_JSON, flat_list, log_msg='add value="{0}"'.format(new_exc))
            ExceptionManager.clear_cache()

        # ------------------------------------------------------------------
        # EDIT
        # ------------------------------------------------------------------
        elif action == "edit":
            if not flat_list:
                MessageBox.Show("No exceptions to edit.", title, MessageBoxButtons.OK)
                continue

            old_val = forms.SelectFromList.show(flat_list, "Select exception to edit", multiselect=False)
            if old_val is None:
                continue

            new_val = forms.ask_for_string("Enter new value:", old_val)
            if not new_val:
                continue

            if any(e.lower() == new_val.lower() and e.lower() != old_val.lower() for e in flat_list):
                MessageBox.Show("That exception already exists.", title, MessageBoxButtons.OK)
                continue

            flat_list[flat_list.index(old_val)] = new_val
            _save_json(PROJECT_JSON, flat_list, log_msg='edit old="{0}" new="{1}"'.format(old_val, new_val))
            ExceptionManager.clear_cache()

        # ------------------------------------------------------------------
        # DELETE
        # ------------------------------------------------------------------
        elif action == "delete":
            if not flat_list:
                MessageBox.Show("No exceptions to delete.", title, MessageBoxButtons.OK)
                continue

            to_delete = forms.SelectFromList.show(flat_list, "Select exception to delete", multiselect=False)
            if to_delete is None:
                continue

            confirm = MessageBox.Show("Delete \"{0}\"?".format(to_delete), title, MessageBoxButtons.YesNo)
            if confirm != DialogResult.Yes:
                continue

            flat_list.remove(to_delete)
            _save_json(PROJECT_JSON, flat_list, log_msg='delete value="{0}"'.format(to_delete))
            ExceptionManager.clear_cache()

# ---------------------------------------------------------------------------
if __name__ == "__main__":
    try:
        manage_exceptions()
    except Exception as exc:  # noqa: B902 – broad catch for UI feedback
        logger.error("fatal_error=%s", exc, extra={"src": SRC, "version": VERSION, "user": USER})
        MessageBox.Show("Fatal error: {0}".format(exc), "Manage Exceptions – Error", MessageBoxButtons.OK)



#from Snippets._customprint import kit_button_clicked    # Import Reusable Function from 'lib/Snippets/_customprint.py'
#kit_button_clicked(btn_name=__title__)                  # Display Default Print Message