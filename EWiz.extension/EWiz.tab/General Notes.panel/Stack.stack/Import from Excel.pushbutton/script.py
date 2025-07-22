# -*- coding: utf-8 -*-
from __future__ import unicode_literals  # IronPython 2.7 compatibility
__title__   = "Import from Excel"
__doc__     = """Version = 1.2.0
Date    = July 2025
========================================
Description:
Reads the General Notes Excel file (make sure to use the template) and recreates all TextNotes in the active Revit view.
========================================
How‑To:
1. Open the General Notes Drafting View or activate it on the General Notes Sheet.
2. Click the “Import from Excel” button on the ribbon.
3. In the file dialog, select your .xlsm workbook.
4. The script will try to locate the default text‑note styles. If they are unavailable, you will be prompted to choose one from the list of available styles.
5. Watch the console for progress updates, skipped tabs, and added sections. The notes should now update in the Drafting View.
6. Review the updates on the General Notes Sheet.

Note: If you need more than one sheet, split the Excel file into two and run the script separately in two Drafting Views.
========================================
Author: PBM"""

import clr
import sys
import os
import re


# ─────────────────────────────────────────────────────────────────────────────
# Simple console logger
# ─────────────────────────────────────────────────────────────────────────────
def log(msg):
    sys.stdout.write(u"{0}\n".format(msg))


# ─────────────────────────────────────────────────────────────────────────────
# Helpers: Excel reading
# ─────────────────────────────────────────────────────────────────────────────
clr.AddReference('Microsoft.Office.Interop.Excel')
import Microsoft.Office.Interop.Excel as Excel
from System import Type, Activator, Array
from System.Runtime.InteropServices import Marshal
import System.Reflection


def read_excel_worksheets(path):
    """Read all worksheets and their data from Excel."""
    log(u"📂  Opening Excel workbook: {0}".format(path))

    app = None
    wb = None
    worksheets_data = []

    try:
        # Create Excel application using COM
        excel_type = Type.GetTypeFromProgID("Excel.Application")
        app = Activator.CreateInstance(excel_type)

        # Set properties using reflection for COM objects
        app.GetType().InvokeMember("Visible",
                                   System.Reflection.BindingFlags.SetProperty,
                                   None, app, Array[object]([False]))
        app.GetType().InvokeMember("DisplayAlerts",
                                   System.Reflection.BindingFlags.SetProperty,
                                   None, app, Array[object]([False]))

        # Get Workbooks collection and open file
        workbooks = app.GetType().InvokeMember("Workbooks",
                                               System.Reflection.BindingFlags.GetProperty,
                                               None, app, None)
        wb = workbooks.GetType().InvokeMember("Open",
                                              System.Reflection.BindingFlags.InvokeMethod,
                                              None, workbooks, Array[object]([path]))

        # Get Worksheets collection
        worksheets = wb.GetType().InvokeMember("Worksheets",
                                               System.Reflection.BindingFlags.GetProperty,
                                               None, wb, None)
        count = worksheets.GetType().InvokeMember("Count",
                                                  System.Reflection.BindingFlags.GetProperty,
                                                  None, worksheets, None)

        for i in range(1, count + 1):
            ws = worksheets.GetType().InvokeMember("Item",
                                                   System.Reflection.BindingFlags.GetProperty,
                                                   None, worksheets, Array[object]([i]))
            sheet_name = ws.GetType().InvokeMember("Name",
                                                   System.Reflection.BindingFlags.GetProperty,
                                                   None, ws, None)

            # Skip Splash Screen tab
            if sheet_name == 'Splash Screen':
                log(u"⏭️  Skipping '{0}' tab".format(sheet_name))
                continue

            # Check H1 cell
            h1_range = ws.GetType().InvokeMember("Range",
                                                 System.Reflection.BindingFlags.GetProperty,
                                                 None, ws, Array[object](["H1"]))
            h1_value = h1_range.GetType().InvokeMember("Value2",
                                                       System.Reflection.BindingFlags.GetProperty,
                                                       None, h1_range, None)
            if h1_value != 'Yes':
                log(u"⏭️  Skipping '{0}' tab (H1 = {1})".format(sheet_name, h1_value))
                continue

            # Read J3 cell content
            j3_range = ws.GetType().InvokeMember("Range",
                                                 System.Reflection.BindingFlags.GetProperty,
                                                 None, ws, Array[object](["J3"]))
            j3_value = j3_range.GetType().InvokeMember("Value2",
                                                       System.Reflection.BindingFlags.GetProperty,
                                                       None, j3_range, None)
            if j3_value:
                worksheets_data.append({
                    'title': sheet_name,
                    'content': unicode(j3_value)
                })
                log(u"✅  Added '{0}' tab for processing".format(sheet_name))
            else:
                log(u"⚠️  '{0}' tab has empty J3 cell".format(sheet_name))

    except Exception as ex:
        log(u"❌  Excel read error: {0}".format(ex))
    finally:
        # Clean up COM objects
        try:
            if wb is not None:
                wb.GetType().InvokeMember("Close",
                                          System.Reflection.BindingFlags.InvokeMethod,
                                          None, wb, Array[object]([False]))
                Marshal.ReleaseComObject(wb)
            if app is not None:
                app.GetType().InvokeMember("Quit",
                                           System.Reflection.BindingFlags.InvokeMethod,
                                           None, app, None)
                Marshal.ReleaseComObject(app)
        except:
            pass

    return worksheets_data


# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

# Layout settings
START_X = 8.883660091  # Starting X position (top left corner)
START_Y = 4.440310784  # Starting Y position (top left corner)
BOTTOM_Y = 2.794031111  # Bottom Y position of the page
SECTION_SPACING = 0.002  # Small gap between title and content
INTER_SECTION_SPACING = 0.05  # Gap between sections
COLUMN_WIDTH = 0.37  # Width of each column
PAGE_HEIGHT = START_Y - BOTTOM_Y  # Calculate page height from coordinates
TEXT_WIDTH = 0.3  # Width constraint for text notes

# Abbreviations section fixed coordinates
ABBREV_X = 11.181472982
ABBREV_Y = 4.4015

# ─────────────────────────────────────────────────────────────────────────────
# Load Revit API
# ─────────────────────────────────────────────────────────────────────────────
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')
from Autodesk.Revit.DB import (
    TextNote,
    Transaction,
    XYZ,
    TextNoteOptions,
    FilteredElementCollector,
    TextNoteType,
    BuiltInParameter
)
from Autodesk.Revit.UI import TaskDialog
from System.Windows.Forms import OpenFileDialog, DialogResult
from System.Windows.Forms import Form, Label, ComboBox, Button, DialogResult as WinFormsDialogResult, FormStartPosition, \
    ComboBoxStyle
from System.Drawing import Size, Point


# ─────────────────────────────────────────────────────────────────────────────
# Helper functions
# ─────────────────────────────────────────────────────────────────────────────

def find_excel_file(doc):
    return prompt_for_excel_file()


def prompt_for_excel_file():
    """Prompt user to select the Excel file."""
    try:
        dialog = OpenFileDialog()
        dialog.Title = "Select GenNotes Excel File"
        dialog.Filter = "Excel Files (*.xlsm)|*.xlsm|All Files (*.*)|*.*"
        dialog.FilterIndex = 1
        dialog.Multiselect = False

        if dialog.ShowDialog() == DialogResult.OK:
            excel_path = dialog.FileName
            log(u"📁  User selected Excel file: {0}".format(excel_path))

            # Validate the selected file
            if not os.path.exists(excel_path):
                TaskDialog.Show("Error", "Selected file does not exist.")
                return None

            if not excel_path.lower().endswith('.xlsm'):
                TaskDialog.Show("Error", "Please select an Excel file with .xlsm extension.")
                return None

            return excel_path
        else:
            log(u"❌  User cancelled file selection")
            return None

    except Exception as ex:
        log(u"❌  Error prompting for Excel file: {0}".format(ex))
        TaskDialog.Show("Error", "Error opening file dialog: {0}".format(ex))
        return None


def delete_all_textnotes(doc):
    """Delete all TextNotes from the current view."""
    try:
        # Get all TextNotes in the current view
        text_notes = FilteredElementCollector(doc, doc.ActiveView.Id) \
            .OfClass(TextNote) \
            .ToElements()

        deleted_count = 0
        if text_notes:
            log(u"🗑️  Found {0} existing TextNotes to delete".format(len(text_notes)))
            for note in text_notes:
                doc.Delete(note.Id)
                deleted_count += 1
            log(u"✅  Deleted {0} existing TextNotes".format(deleted_count))
        else:
            log(u"ℹ️  No existing TextNotes found to delete")

        return deleted_count

    except Exception as ex:
        log(u"⚠️  Error deleting TextNotes: {0}".format(ex))
        return 0


def select_text_note_type(available_types, type_purpose):
    """Show a dialog to let user select a text note type."""
    try:
        # Create form
        form = Form()
        form.Text = "Select {0} Text Note Type".format(type_purpose)
        form.Size = Size(450, 150)
        form.StartPosition = FormStartPosition.CenterParent

        # Create label
        label = Label()
        label.Text = "Select {0} text note type:".format(type_purpose)
        label.Location = Point(10, 10)
        label.Size = Size(400, 20)

        # Create combo box
        combo = ComboBox()
        combo.Location = Point(10, 35)
        combo.Size = Size(415, 25)
        combo.DropDownStyle = ComboBoxStyle.DropDownList

        # Populate combo box with type names
        type_names = []
        for t in available_types:
            try:
                name = t.Name
            except AttributeError:
                param = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
                if param:
                    name = param.AsString()
                else:
                    name = u"<Unnamed TextNoteType>"
            type_names.append(name)

        for name in type_names:
            combo.Items.Add(name)

        if combo.Items.Count > 0:
            combo.SelectedIndex = 0

        # Create OK button
        ok_button = Button()
        ok_button.Text = "OK"
        ok_button.Location = Point(270, 70)
        ok_button.Size = Size(75, 25)
        ok_button.DialogResult = WinFormsDialogResult.OK

        # Create Cancel button
        cancel_button = Button()
        cancel_button.Text = "Cancel"
        cancel_button.Location = Point(350, 70)
        cancel_button.Size = Size(75, 25)
        cancel_button.DialogResult = WinFormsDialogResult.Cancel

        # Add controls to form
        form.Controls.Add(label)
        form.Controls.Add(combo)
        form.Controls.Add(ok_button)
        form.Controls.Add(cancel_button)
        form.AcceptButton = ok_button
        form.CancelButton = cancel_button

        # Show dialog
        result = form.ShowDialog()

        if result == WinFormsDialogResult.OK and combo.SelectedIndex >= 0:
            selected_name = combo.SelectedItem.ToString()
            # Find the corresponding TextNoteType
            for i, t in enumerate(available_types):
                if type_names[i] == selected_name:
                    return t

        return None

    except Exception as ex:
        log(u"❌  Error in text note type selection dialog: {0}".format(ex))
        return None


def calculate_text_note_height(text_note_type, text_content, text_width):
    """Calculate the height of a TextNote based on its type and content, processing paragraph by paragraph."""
    # Cannot simply be extracted from the TextNote because that information is not available until the transaction is committed. Therefore we need to estimate it based on the TextNoteType parameters and the content.
    try:
        # Get text size from the TextNoteType
        text_size_param = text_note_type.get_Parameter(BuiltInParameter.TEXT_SIZE)
        if text_size_param:
            text_size = text_size_param.AsDouble()
        else:
            log(u"⚠️  No TEXT_SIZE parameter found for TextNoteType '{0}'".format(text_note_type.Name))
            text_size = 0.1  # fallback text size

        # Estimate line height including spacing
        line_height = text_size * 1.6

        # Estimated width of each character
        avg_char_width = text_size * 0.55

        # Calculate approximate characters per line
        chars_per_line = max(1, int(text_width / avg_char_width))

        # Split text into paragraphs
        if text_content:
            paragraphs = text_content.split('\n')
        else:
            paragraphs = ['']

        total_lines = 0

        # Calculate lines for each paragraph
        for paragraph in paragraphs:
            if paragraph.strip():  # Non-empty paragraph
                paragraph_length = len(paragraph.strip())
                paragraph_lines = max(1, (paragraph_length + chars_per_line - 1) // chars_per_line)
                total_lines += paragraph_lines
                # log(u"📏  Paragraph '{0}...' length: {1}, lines: {2}".format(
                #     paragraph[:20], paragraph_length, paragraph_lines))
            else:  # Empty paragraph (line break)
                total_lines += 1
                # log(u"📏  Empty paragraph (line break): 1 line")

        # Calculate total height
        total_height = total_lines * line_height

        # log(u"📏  Text size: {0}, Line height: {1}, Chars per line: {2}, Total lines: {3}, Total height: {4}".format(
        #     text_size, line_height, chars_per_line, total_lines, total_height))

        return total_height

    except Exception as ex:
        log(u"⚠️  Error calculating TextNote height: {0}".format(ex))
        return 0.2  # fallback height


def check_text_note_fits(current_y, text_height, bottom_boundary):
    """Check if a TextNote would fit within the page boundaries."""
    return (current_y - text_height) >= bottom_boundary


def is_abbreviations_section(title):
    """Check if a section is an abbreviations section based on title."""
    abbrev_keywords = ['abbreviation', 'abbrev', 'acronym']
    title_lower = title.lower()
    return any(keyword in title_lower for keyword in abbrev_keywords)


# ─────────────────────────────────────────────────────────────────────────────
# Main logic
# ─────────────────────────────────────────────────────────────────────────────

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

# Find Excel file
EXCEL_PATH = find_excel_file(doc)
if not EXCEL_PATH:
    TaskDialog.Show("Error", "No Excel file selected. Script will exit.")
    sys.exit()

# Read Excel data
worksheets_data = read_excel_worksheets(EXCEL_PATH)
if not worksheets_data:
    TaskDialog.Show("No Data", "No valid worksheets found to process.")
    sys.exit()

tx = Transaction(doc, "Create TextNotes from Excel")
tx.Start()

try:
    # Delete all existing TextNotes first
    delete_all_textnotes(doc)

    # Get text note types (styles)
    title_type = None
    content_type = None

    text_note_types = FilteredElementCollector(doc) \
        .OfClass(TextNoteType) \
        .ToElements()

    log(u"📋  Found {0} text note types".format(len(text_note_types)))

    # Debug: list all available text note types safely
    log(u"📋  Available text note types:")
    for t in text_note_types:
        # Try .Name, fallback to SYMBOL_NAME_PARAM, else placeholder
        try:
            name = t.Name
        except AttributeError:
            param = t.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
            if param:
                name = param.AsString()
            else:
                name = u"<Unnamed TextNoteType>"

        # log(u"   - {0}".format(name))

        # Match by that name
        if name == 'EWP_3.5mm Arrow Masking':
            title_type = t
        elif name == 'EWP_2.5mm Arrow':
            content_type = t

    # If types not found, prompt user to select
    if not title_type:
        log(u"⚠️  'EWP_3.5mm Arrow' not found. Prompting user to select title type.")
        title_type = select_text_note_type(text_note_types, "Title")
        if not title_type:
            log(u"❌  No title type selected.")
            tx.RollBack()
            TaskDialog.Show("Error", "No title text note type selected.")
            sys.exit()

    if not content_type:
        log(u"⚠️  'EWP_2.5mm Arrow' not found. Prompting user to select content type.")
        content_type = select_text_note_type(text_note_types, "Content")
        if not content_type:
            log(u"❌  No content type selected.")
            tx.RollBack()
            TaskDialog.Show("Error", "No content text note type selected.")
            sys.exit()

    # Validate and adjust text width for both types
    title_min_width = TextNote.GetMinimumAllowedWidth(doc, title_type.Id)
    title_max_width = TextNote.GetMaximumAllowedWidth(doc, title_type.Id)
    title_width = max(title_min_width, min(TEXT_WIDTH, title_max_width))

    content_min_width = TextNote.GetMinimumAllowedWidth(doc, content_type.Id)
    content_max_width = TextNote.GetMaximumAllowedWidth(doc, content_type.Id)
    content_width = max(content_min_width, min(TEXT_WIDTH, content_max_width))

    # log(u"📏  Title width: {0} (min: {1}, max: {2})".format(title_width, title_min_width, title_max_width))
    # log(u"📏  Content width: {0} (min: {1}, max: {2})".format(content_width, content_min_width, content_max_width))

    current_y = START_Y
    current_column = 0
    created_notes = 0

    # Separate abbreviations from regular sections
    regular_sections = []
    abbreviations_sections = []

    for data in worksheets_data:
        if is_abbreviations_section(data['title']):
            abbreviations_sections.append(data)
        else:
            regular_sections.append(data)

    # Process regular sections in columns
    for data in regular_sections:
        # Calculate heights before placing
        title_height = calculate_text_note_height(title_type, data['title'], title_width)
        content_height = calculate_text_note_height(content_type, data['content'], content_width)

        # Calculate total height needed for this section
        total_section_height = title_height + SECTION_SPACING + content_height + INTER_SECTION_SPACING

        # Check if the entire section would fit, if not move to next column
        if not check_text_note_fits(current_y, total_section_height, BOTTOM_Y):
            current_column += 1
            current_y = START_Y
            # log(u"📄  Moving to column {0} for section '{1}'".format(current_column + 1, data['title']))

        current_x = START_X + (current_column * COLUMN_WIDTH)

        # Create title TextNote
        title_point = XYZ(current_x, current_y, 0)
        title_options = TextNoteOptions()
        title_options.TypeId = title_type.Id

        title_note = TextNote.Create(doc, doc.ActiveView.Id, title_point, title_width, data['title'], title_options)
        created_notes += 1

        # Adjust Y position after title
        current_y -= (title_height + SECTION_SPACING)
        # log(u"📏  Title '{0}' height: {1}".format(data['title'], title_height))

        # Create content TextNote
        content_point = XYZ(current_x, current_y, 0)
        content_options = TextNoteOptions()
        content_options.TypeId = content_type.Id

        content_note = TextNote.Create(doc, doc.ActiveView.Id, content_point, content_width, data['content'],
                                       content_options)
        created_notes += 1

        # Adjust Y position for next section
        current_y -= (content_height + INTER_SECTION_SPACING)
        # log(u"📏  Content height: {0}".format(content_height))

        log(u"📝  Created notes for '{0}' at X: {1} Y: {2}".format(data['title'], current_x, current_y))

    # Process abbreviations sections at fixed coordinates
    abbrev_y = ABBREV_Y
    for data in abbreviations_sections:
        # Calculate heights for abbreviations
        title_height = calculate_text_note_height(title_type, data['title'], title_width)
        content_height = calculate_text_note_height(content_type, data['content'], content_width)

        # Create title TextNote at fixed position
        title_point = XYZ(ABBREV_X, abbrev_y, 0)
        title_options = TextNoteOptions()
        title_options.TypeId = title_type.Id

        title_note = TextNote.Create(doc, doc.ActiveView.Id, title_point, title_width, data['title'], title_options)
        created_notes += 1

        # Adjust Y position after title
        abbrev_y -= (title_height + SECTION_SPACING)

        # Create content TextNote
        content_point = XYZ(ABBREV_X, abbrev_y, 0)
        content_options = TextNoteOptions()
        content_options.TypeId = content_type.Id

        content_note = TextNote.Create(doc, doc.ActiveView.Id, content_point, content_width, data['content'],
                                       content_options)
        created_notes += 1

        # Adjust Y position for next abbreviations section
        abbrev_y -= (content_height + INTER_SECTION_SPACING)

        log(u"📝  Created abbreviations '{0}' at fixed position X: {1} Y: {2}".format(data['title'], ABBREV_X, abbrev_y))

    tx.Commit()
    log(u"✅  Done. Created {0} TextNotes ({1} regular, {2} abbreviations).".format(
        created_notes, len(regular_sections) * 2, len(abbreviations_sections) * 2))
    TaskDialog.Show(
        "Finished",
        u"Created {0} TextNotes from {1} Excel tabs ({2} regular sections, {3} abbreviations sections).".format(
            created_notes, len(worksheets_data), len(regular_sections), len(abbreviations_sections))
    )

except Exception as ex:
    # Roll back the transaction
    tx.RollBack()

    # Get full traceback
    import traceback

    tb = traceback.format_exc()

    log(u"❌  Error creating TextNotes: {0}".format(tb))

    # Show full traceback in a dialog
    TaskDialog.Show(
        "Error creating TextNotes – full traceback",
        tb
    )
    sys.exit()
