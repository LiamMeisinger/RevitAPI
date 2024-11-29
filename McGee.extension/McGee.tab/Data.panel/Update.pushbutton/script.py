# -*- coding: utf-8 -*-
__doc__ = "This will update the Selected Excel with current Revit revision information"
__title__ = "Update Excel"  #Title of the extension
__author__ = "Liam Meisinger"

# IMPORTS
from Autodesk.Revit.DB import *
import os
import xlsxwriter


#.NET Imports
import clr
from Autodesk.Revit.UI.Selection import Selection


clr.AddReference('System')

# VARIABLES
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document  #type:Document
selection = uidoc.Selection  #type: Selection
active_view = doc.ActiveView
active_level = doc.ActiveView.GenLevel

"""Sheet Params"""
global dest


# Function to get all sheets
def get_all_sheets():
    return FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_Sheets) \
        .WhereElementIsNotElementType() \
        .ToElements()


# Function to get revision information for a sheet
def get_sheet_revisions(sheet):
    revisions = []
    revision_ids = sheet.GetAllRevisionIds()

    for rev_id in revision_ids:
        rev = doc.GetElement(rev_id)
        rev_date = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_DATE).AsString()
        rev_desc = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_DESCRIPTION).AsString()
        rev_num = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_SEQUENCE_NUM).AsString()
        rev_d_by = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_ISSUED_TO).AsString()
        rev_i_by = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_ISSUED_BY).AsString()

        revisions.append({
            "Sheet Number": sheet.SheetNumber,
            "Sheet Name": sheet.Name,
            "Revision Number": rev_num,
            "Revision Description": rev_desc,
            "Revision Date": rev_date,
            "Drawn By": rev_d_by,
            "Issued By": rev_i_by
        })

    return revisions


# Function to export data to Excel
def export_to_excel(data, file_path):
    workbook = xlsxwriter.Workbook(file_path)
    worksheet = workbook.add_worksheet("Revisions")

    # Write the header
    headers = ["Sheet Number", "Sheet Name", "Revision Number", "Revision Description", "Revision Date", "Drawn By",
               "Issued By"]
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    # Write the data
    for row_num, row_data in enumerate(data, start=1):
        worksheet.write(row_num, 0, row_data["Sheet Number"])
        worksheet.write(row_num, 1, row_data["Sheet Name"])
        worksheet.write(row_num, 2, row_data["Revision Number"])
        worksheet.write(row_num, 3, row_data["Revision Description"])
        worksheet.write(row_num, 4, row_data["Revision Date"])
        worksheet.write(row_num, 5, row_data["Drawn By"])
        worksheet.write(row_num, 6, row_data["Issued By"])

    # Auto-fit columns
    for col_num, header in enumerate(headers):
        # Determine the maximum length of the data in each column
        max_length = len(header)
        for row_data in data:
            max_length = max(max_length, len(str(row_data[header.replace(" ", " ")])))

        # Set the column width based on the maximum length (+ 2 for some padding)
        worksheet.set_column(col_num, col_num, max_length + 2)

    # Close the workbook to save the file
    workbook.close()


# Get all sheets in the project
sheets = get_all_sheets()

# Collect revision information for each sheet
all_revisions = []
for sheet in sheets:
    sheet_revisions = get_sheet_revisions(sheet)
    all_revisions.extend(sheet_revisions)

# Specify the file path where the Excel file will be saved
output_dir = os.path.expanduser("~/Documents")
output_file = os.path.join(output_dir, "Sheet_Revisions.xlsx")

# Export the collected revision information to an Excel file
export_to_excel(all_revisions, output_file)

print("Export completed. The file is saved at: " + output_file)