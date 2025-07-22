# -*- coding: utf-8 -*-
from __future__ import unicode_literals     # IronPython 2.7 compatibility
__title__   = "Issue Sheet Generator with Project No."
__doc__     = """Version = 1.0.0
Date    = July 2025
========================================
Description:
Generates Issue Sheets for projects using the BIM Number
sheet numbering scheme. Collects all sheets marked
"Appears In Sheet List" and populates an Excel template
========================================
How-To:
1. Click the button on the ribbon.
2. In the Save File dialog, select where to save the Issue Sheet.
3. Wait for the script to finish — its path will be printed in the console.
========================================
**Important:**
Make sure every sheet you want to include has the
“Appears In Sheet List” flag enabled under Properties → Identity Data.
========================================
Author: AO"""

import os
import shutil
import clr
import re
import sys
import datetime
from collections import OrderedDict
from System.Runtime.InteropServices import Marshal

# Add necessary .NET and Revit references
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('Microsoft.Office.Interop.Excel')
clr.AddReference('System.Windows.Forms')

import Microsoft.Office.Interop.Excel as Excel
from System.Windows.Forms import SaveFileDialog, DialogResult

from Autodesk.Revit.DB import (
    FilteredElementCollector, BuiltInCategory,
    ViewSheet, RevisionCloud, Revision,
    RevisionNumberType
)
from Autodesk.Revit.UI import TaskDialog

# -- Helper functions --

def current_date():
    """Return current date as 'DDMMYY'."""
    from System import DateTime
    return DateTime.Now.ToString("ddMMyy")


def get_rev_number(revision, sheet=None):
    """Retrieve the revision number on a sheet or sequence number."""
    if sheet and isinstance(sheet, ViewSheet):
        return sheet.GetRevisionNumberOnSheet(revision.Id)
    if hasattr(revision, 'RevisionNumber'):
        return revision.RevisionNumber
    return revision.SequenceNumber


def excel_col_name(n):
    """Convert zero-based index to Excel column name."""
    name = ''
    while n >= 0:
        n, r = divmod(n, 26)
        name = chr(65 + r) + name
        n -= 1
    return name


def save_file_dialog(init_dir):
    """Show a standard Save File dialog and return chosen file path or None."""
    dialog = SaveFileDialog()
    dialog.InitialDirectory = init_dir
    # Use .NET DateTime or fallback to Python datetime
    try:
        timestamp = datetime.datetime.Now.ToString("ddMMyy")
    except AttributeError:
        timestamp = datetime.datetime.now().strftime("%d%m%y")
    dialog.FileName = "Issue Sheet_{0}.xlsx".format(timestamp)
    dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
    dialog.Title = "Save Issue Sheet"
    result = dialog.ShowDialog()
    if result == DialogResult.OK:
        return dialog.FileName
    return None

# -- Revit context --
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# -- Collect all sheets and revisions --
all_sheets = sorted(
    FilteredElementCollector(doc)
        .OfCategory(BuiltInCategory.OST_Sheets)
        .WhereElementIsNotElementType()
        .ToElements(),
    key=lambda s: s.SheetNumber
)
all_clouds = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_RevisionClouds) \
    .WhereElementIsNotElementType() \
    .ToElements()
all_revisions = FilteredElementCollector(doc) \
    .OfCategory(BuiltInCategory.OST_Revisions) \
    .WhereElementIsNotElementType() \
    .ToElements()

# -- Gather revision metadata --
rev_data = []
date_pattern = re.compile(r'^\d{1,2}[./]\d{1,2}[./]\d{2}$')
for rev in all_revisions:
    num = get_rev_number(rev)
    raw_date = rev.RevisionDate
    try:
        d = raw_date.ToShortDateString()
    except AttributeError:
        d = str(raw_date).strip()
    if not date_pattern.match(d):
        continue
    rev_data.append((num, d, rev.Description))
rev_data.sort(key=lambda x: x[0])

# -- Class to represent a sheet with revisions --
class RevisedSheet(object):
    def __init__(self, sheet):
        self._sheet = sheet
        self._clouds = []
        self._rev_ids = set()
        self._find_clouds()
        self._find_revisions()

    def _find_clouds(self):
        """Collect all revision clouds visible on the sheet."""
        view_ids = [self._sheet.Id] + \
            [doc.GetElement(vp).ViewId for vp in self._sheet.GetAllViewports()]
        for c in all_clouds:
            if c.OwnerViewId in view_ids:
                self._clouds.append(c)

    def _find_revisions(self):
        """Collect all revision IDs associated with the sheet."""
        self._rev_ids.update(c.RevisionId for c in self._clouds)
        self._rev_ids.update(self._sheet.GetAdditionalRevisionIds())

    @property
    def sheet_number(self):
        return self._sheet.SheetNumber

    @property
    def sheet_name(self):
        return self._sheet.Name

    @property
    def rev_count(self):
        return len(self._rev_ids)

    def get_drawing_number(self):
        """Construct drawing number from project and sheet parameters."""
        proj_info = doc.ProjectInformation
        proj_fields = ['Project Number', 'EWP_Project_Originator Code', 'EWP_Project_Role Code']
        parts = [proj_info.LookupParameter(name).AsString().strip()
                 for name in proj_fields if proj_info.LookupParameter(name)]
        sheet_fields = ['EWP_Sheet_Zone Code', 'EWP_Sheet_Level Code',
                        'EWP_Sheet_Type Code', 'Sheet Number']
        parts += [self._sheet.LookupParameter(name).AsString().strip()
                  for name in sheet_fields if self._sheet.LookupParameter(name)]
        return "-".join(filter(None, parts))

# -- Filter valid sheets for the revision report --
revised_sheets = []
for s in all_sheets:
    p = s.LookupParameter("Appears In Sheet List")
    if not p or p.AsInteger() != 1:
        continue
    rs = RevisedSheet(s)
    if s.GetAllViewports() and rs.rev_count > 0:
        revised_sheets.append(rs)

# -- Prepare Excel report --
template_path = r"I:\\BLU - Service Delivery\\04 Building Information Management\\07 The Button\\Project No_Issue Sheets\\Document Issue Sheet.xlsx"
save_path = save_file_dialog(os.path.dirname(template_path))
if not save_path:
    sys.exit()

shutil.copy(template_path, save_path)

# Instantiate Excel using ApplicationClass to avoid abstract interface error
excel = Excel.ApplicationClass()
excel.Visible = False
excel.DisplayAlerts = False
wb = excel.Workbooks.Open(save_path)

# -- Fill revision dates in template --
for sheet_idx in range(1, wb.Sheets.Count + 1):
    ws = wb.Sheets.Item[sheet_idx]
    for idx, (_n, date_str, _c) in enumerate(rev_data):
        d, m, y = [int(x) for x in re.findall(r'\\d+', date_str)]
        col = excel_col_name(3 + idx)
        ws.Range[col + "6"].Value2 = d
        ws.Range[col + "7"].Value2 = m
        ws.Range[col + "8"].Value2 = y

# -- Fill sheet data in chunks of 27 rows per sheet --
chunk_size = 27
for chunk_idx in range(0, len(revised_sheets), chunk_size):
    ws = wb.Sheets.Item[chunk_idx // chunk_size + 1]
    block = revised_sheets[chunk_idx: chunk_idx + chunk_size]
    for i, rs in enumerate(block):
        row = 10 + i
        ws.Range["A{0}".format(row)].Value2 = rs.get_drawing_number()
        ws.Range["B{0}".format(row)].Value2 = rs.sheet_name
        revs = [doc.GetElement(rid) for rid in rs._rev_ids]
        revs.sort(key=lambda r: r.SequenceNumber)
        seq_groups = OrderedDict()
        for rev in revs:
            seq_groups.setdefault(rev.RevisionNumberingSequenceId, []).append(rev)
        for group in seq_groups.values():
            for idx, rev in enumerate(group, start=1):
                seq = doc.GetElement(rev.RevisionNumberingSequenceId)
                prefix = suffix = ''
                if seq and seq.NumberType == RevisionNumberType.Numeric:
                    settings = seq.GetNumericRevisionSettings()
                    prefix = settings.Prefix or ''
                    suffix = settings.Suffix or ''
                    start = settings.StartNumber
                    min_digits = settings.MinimumDigits
                    number = start + idx - 1
                    label = prefix + str(number).zfill(min_digits) + suffix
                rev_num = get_rev_number(rev)
                for j, (n, _, _) in enumerate(rev_data):
                    if n == rev_num:
                        col = excel_col_name(3 + j)
                        # Use .format() instead of f-string for IronPython compatibility
                        ws.Range["{0}{1}".format(col, row)].Value2 = label
                        break

# -- Save, close and fully release Excel to unlock the file --
wb.Save()
wb.Close(False)
excel.Quit()

# Release COM objects
Marshal.ReleaseComObject(wb)
Marshal.ReleaseComObject(excel)
wb = None
excel = None

import System
# Force .NET garbage collection to finalize COM releases
System.GC.Collect()
System.GC.WaitForPendingFinalizers()
System.GC.Collect()
System.GC.WaitForPendingFinalizers()

print "Revision report saved to: {0}".format(save_path)
