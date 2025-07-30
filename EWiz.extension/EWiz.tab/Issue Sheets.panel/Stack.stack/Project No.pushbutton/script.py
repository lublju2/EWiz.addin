# -*- coding: utf-8 -*-
from __future__ import unicode_literals     # IronPython 2.7 compatibility
__title__   = "Issue Sheet Generator with Project No."
__doc__     = """Version = 1.0.1
Date    = July 2025
========================================
Description:
Generates Issue Sheets for projects using the BIM Number
sheet numbering scheme. Collects all sheets marked
"Appears In Sheet List" and populates a macro‑enabled Excel template.
========================================
How‑To:
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
import sys
import shutil
import clr
import re
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
    ViewSheet, RevisionNumberType
)
from Autodesk.Revit.UI import TaskDialog

# -- Helper functions ------------------------------------------------------

def current_date():
    from System import DateTime
    try:
        return DateTime.Now.ToString("ddMMyy")
    except:
        return datetime.datetime.now().strftime("%d%m%y")

def get_rev_number(revision, sheet=None):
    if sheet and isinstance(sheet, ViewSheet):
        return sheet.GetRevisionNumberOnSheet(revision.Id)
    if hasattr(revision, 'RevisionNumber'):
        return revision.RevisionNumber
    return revision.SequenceNumber

def excel_col_name(n):
    name = ''
    while n >= 0:
        n, r = divmod(n, 26)
        name = chr(65 + r) + name
        n -= 1
    return name

def save_file_dialog(init_dir):
    dlg = SaveFileDialog()
    dlg.InitialDirectory = init_dir
    dlg.FileName = "Issue Sheet_{0}.xlsm".format(current_date())
    dlg.Filter = "Excel Macro‑Enabled Workbook (*.xlsm)|*.xlsm|All Files (*.*)|*.*"
    dlg.Title = "Save Issue Sheet"
    if dlg.ShowDialog() == DialogResult.OK:
        return dlg.FileName
    return None

def filter_valid_sheets_and_show_count(all_sheets, all_revisions):
    """
    Keep only sheets marked "Appears In Sheet List" with at least one revision.
    Show total number of revisions in the model.
    """
    revised_sheets = []
    for sheet in all_sheets:
        p = sheet.LookupParameter("Appears In Sheet List")
        if not p or p.AsInteger() != 1:
            continue
        rs = RevisedSheet(sheet)
        if sheet.GetAllViewports() and rs.rev_count > 0:
            revised_sheets.append(rs)
    total_revisions = len(all_revisions)
    #TaskDialog.Show(
        #"Revision Summary",
        #"Total revisions in model: {0}".format(total_revisions)
    #)
    return revised_sheets, total_revisions

# -- Revision filtering ----------------------------------------------------

def build_filtered_rev_data(all_revisions, revised_sheets):
    """
    Return a list of tuples (revId, num, date_str, desc) for only those revisions
    actually assigned to one of the revised_sheets, sorted by num.
    """
    # collect all used revision IDs
    assigned = set()
    for rs in revised_sheets:
        assigned |= rs._rev_ids

    date_rx = re.compile(r'^\d{1,2}[./]\d{1,2}[./]\d{2,4}$')
    filtered = []
    for rev in all_revisions:
        if rev.Id not in assigned:
            continue
        num = get_rev_number(rev)
        raw = rev.RevisionDate
        try:
            dstr = raw.ToShortDateString()
        except AttributeError:
            dstr = str(raw).strip()
        if not date_rx.match(dstr):
            continue
        filtered.append((rev.Id, num, dstr, rev.Description))

    filtered.sort(key=lambda x: x[1])
    return filtered

# -- Chunked fill functions ------------------------------------------------

def fill_revision_header_chunk(ws, rev_chunk):
    """
    Fill the revision date header (rows 6–8, columns D onward) for a slice of revisions.
    rev_chunk: list of (revId, num, date_str, desc) tuples.
    """
    for j, (_id, _num, date_str, _desc) in enumerate(rev_chunk):
        parts = re.findall(r'\d+', date_str)
        if len(parts) != 3:
            continue
        d, m, y = map(int, parts)
        col = excel_col_name(3 + j)  # 0→A,1→B,2→C so 3→D
        ws.Range["{0}6".format(col)].Value2 = d
        ws.Range["{0}7".format(col)].Value2 = m
        ws.Range["{0}8".format(col)].Value2 = y

def fill_sheet_block_chunk(ws, sheet_block, rev_chunk):
    """
    Fill rows 10–36 for a block of up to 27 sheets, writing drawing numbers,
    sheet names, and revision labels only for revs in rev_chunk.
    rev_chunk: list of (revId, num, date_str, desc).
    """
    # map revId→index in this chunk
    rev_index = {rid: idx for idx, (rid, _, _, _) in enumerate(rev_chunk)}

    for i, rs in enumerate(sheet_block):
        row = 10 + i
        ws.Range["A{0}".format(row)].Value2 = rs.get_drawing_number()
        ws.Range["B{0}".format(row)].Value2 = rs.sheet_name

        # gather this sheet's revisions, filter to this chunk
        revs = [doc.GetElement(rid) for rid in rs._rev_ids if rid in rev_index]
        revs.sort(key=lambda r: r.SequenceNumber)

        groups = OrderedDict()
        for rev in revs:
            groups.setdefault(rev.RevisionNumberingSequenceId, []).append(rev)

        # write each label in the correct column
        for grp in groups.values():
            for seq_idx, rev in enumerate(grp, start=1):
                rid = rev.Id
                j = rev_index.get(rid)
                if j is None:
                    continue
                seq = doc.GetElement(rev.RevisionNumberingSequenceId)
                if seq and seq.NumberType == RevisionNumberType.Numeric:
                    s = seq.GetNumericRevisionSettings()
                    prefix = s.Prefix or ''
                    suffix = s.Suffix or ''
                    num_val = s.StartNumber + seq_idx - 1
                    label = prefix + str(num_val).zfill(s.MinimumDigits) + suffix
                else:
                    label = get_rev_number(rev)
                col = excel_col_name(3 + j)
                ws.Range["{0}{1}".format(col, row)].Value2 = label

def fill_all_chunks(wb, all_revisions, revised_sheets,
                    rev_chunk_size=50, sheet_chunk_size=27):
    """
    Orchestrate filling of the workbook by splitting revisions into rev_chunk_size
    and sheets into sheet_chunk_size, then filling each worksheet accordingly.
    """
    rev_data = build_filtered_rev_data(all_revisions, revised_sheets)
    total_rev = len(rev_data)
    total_sheets = len(revised_sheets)
    num_rev_chunks = (total_rev + rev_chunk_size - 1) // rev_chunk_size
    num_sheet_chunks = (total_sheets + sheet_chunk_size - 1) // sheet_chunk_size

    for rc in range(num_rev_chunks):
        rev_chunk = rev_data[rc*rev_chunk_size : (rc+1)*rev_chunk_size]
        for sc in range(num_sheet_chunks):
            sheet_block = revised_sheets[sc*sheet_chunk_size : (sc+1)*sheet_chunk_size]
            ws_index = rc * num_sheet_chunks + sc + 1
            ws = wb.Sheets.Item[ws_index]
            fill_revision_header_chunk(ws, rev_chunk)
            fill_sheet_block_chunk(ws, sheet_block, rev_chunk)

# -- Class to represent a sheet with its revisions ------------------------

class RevisedSheet(object):
    def __init__(self, sheet):
        self._sheet = sheet
        self._clouds = []
        self._rev_ids = set()
        self._find_clouds()
        self._find_revisions()

    def _find_clouds(self):
        view_ids = [self._sheet.Id]
        view_ids += [doc.GetElement(vp).ViewId for vp in self._sheet.GetAllViewports()]
        for c in all_clouds:
            if c.OwnerViewId in view_ids:
                self._clouds.append(c)

    def _find_revisions(self):
        for c in self._clouds:
            self._rev_ids.add(c.RevisionId)
        for rid in self._sheet.GetAdditionalRevisionIds():
            self._rev_ids.add(rid)

    @property
    def sheet_name(self):
        return self._sheet.Name

    @property
    def rev_count(self):
        return len(self._rev_ids)

    def get_drawing_number(self):
        parts = []
        proj_info = doc.ProjectInformation
        params = [
            'Project Number',
            'EWP_Project_Originator Code',
            'EWP_Sheet_Zone Code',
            'EWP_Sheet_Level Code',
            'EWP_Sheet_Type Code',
            'EWP_Project_Role Code',
            'Sheet Number'
        ]
        for pname in params:
            if 'Project' in pname:
                p = proj_info.LookupParameter(pname)
            else:
                p = self._sheet.LookupParameter(pname)
            if p and p.AsString():
                parts.append(p.AsString().strip())
        return "-".join(parts)

# -- Main execution --------------------------------------------------------

def main():
    global doc, all_clouds

    uidoc = __revit__.ActiveUIDocument
    doc = uidoc.Document

    # collect sheets, clouds, revisions
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

    # filter sheets & show total count
    revised_sheets, total_revisions = filter_valid_sheets_and_show_count(
        all_sheets, all_revisions
    )

    # prepare Excel template
    template_path = r"I:\BLU - Service Delivery\04 Building Information Management\07 EWiz\Document Issue Sheet.xlsm"
    if not os.path.exists(template_path):
        TaskDialog.Show("Error", "Template not found:\n" + template_path)
        sys.exit()

    save_path = save_file_dialog(os.path.dirname(template_path))
    if not save_path:
        sys.exit()
    shutil.copy(template_path, save_path)

    # launch Excel
    excel = Excel.ApplicationClass()
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(save_path)

    # fill job info on each sheet
    proj_info = doc.ProjectInformation
    nm = proj_info.LookupParameter('Project Name')
    num = proj_info.LookupParameter('Project Number')
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets.Item[i]
        name = nm.AsString().strip() if nm else ""
        number = num.AsString().strip() if num else ""
        ws.Range["A6"].Value2 = "Job name: " + name
        ws.Range["A6"].Font.Bold = True
        ws.Range["B7"].Value2 = "Job no: " + number
        ws.Range["B7"].Font.Bold = True

    # fill revisions & sheets in chunks, filtering out unused revisions
    fill_all_chunks(wb, all_revisions, revised_sheets)

    # save & clean up COM
    wb.Save()
    wb.Close(False)
    excel.Quit()
    Marshal.ReleaseComObject(wb)
    Marshal.ReleaseComObject(excel)
    wb, excel = None, None

    import System
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()
    System.GC.Collect()
    System.GC.WaitForPendingFinalizers()

    print "Revision report saved to: {0}".format(save_path)

if __name__ == '__main__':
    main()
