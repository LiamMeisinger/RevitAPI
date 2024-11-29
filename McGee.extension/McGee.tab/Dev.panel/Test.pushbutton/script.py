# -*- coding: utf-8 -*-
__doc__ = "This is my testing script, currently testing Splitting Lines"
__title__ = "Testing"  #Title of the extension
__author__ = "Liam Meisinger"

# IMPORTS
from Autodesk.Revit.DB import *
from pyrevit import forms
import os
import xlsxwriter


#.NET Imports
import clr
from Autodesk.Revit.UI.Selection import Selection
import sys

clr.AddReference('System')

# VARIABLES
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document  #type:Document
selection = uidoc.Selection  #type: Selection
active_view = doc.ActiveView
active_level = doc.ActiveView.GenLevel

"""Testing Sheet Params"""
global dest


# Function to get all sheets
def get_all_sheets():
    return FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_Sheets) \
        .WhereElementIsNotElementType() \
        .ToElements()


# Function to extract parameters from a sheet
def get_sheet_parameters(sheet):
    parameters = {}
    for param in sheet.Parameters:  # Iterate over all parameters
        param_name = param.Definition.Name
        if param.StorageType == StorageType.String:
            param_value = param.AsString()
        elif param.StorageType == StorageType.Integer:
            param_value = param.AsInteger()
        elif param.StorageType == StorageType.Double:
            param_value = param.AsDouble()
        elif param.StorageType == StorageType.ElementId:
            param_value = param.AsElementId().IntegerValue
        else:
            param_value = "Unknown"

        parameters[param_name] = param_value

    return parameters


# Collect all sheets in the project
sheets = get_all_sheets()

# List to store all sheet parameters
all_sheet_parameters = []

for sheet in sheets:
    sheet_params = get_sheet_parameters(sheet)
    all_sheet_parameters.append(sheet_params)


# Now all_sheet_parameters contains a list of dictionaries, each containing the parameters of a sheet
def get_sheet_revisions(sheet):
    revisions = []
    revision_ids = sheet.GetAllRevisionIds()

    for rev_id in revision_ids:
        rev = doc.GetElement(rev_id)
        rev_date = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_DATE).AsString()
        rev_desc = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_DESCRIPTION).AsString()
        rev_num = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_NUM).AsString()
        rev_d_by = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_ISSUED_TO).AsString()
        rev_i_by = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_ISSUED_BY).AsString()

        revisions.append({
            "Sheet Number": sheet.LookupParameter("MCG_Project Code").AsString()+"-" +
                            sheet.LookupParameter("MCG_Organisation").AsString()+"-" +
                            sheet.LookupParameter("MCG_Volume").AsString()+"-" +
                            sheet.LookupParameter("MCG_Level Ref").AsString()+"-" +
                            sheet.LookupParameter("MCG_Document Type").AsString()+"-" +
                            sheet.LookupParameter("MCG_Discipline").AsString()+"-" +
                            sheet.SheetNumber,
            "Sheet Name": sheet.Name.title(),
            "Current Revision": sheet.LookupParameter("Current Revision").AsString(),
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
    worksheet = workbook.add_worksheet("Sheet Data")

    # Write the header
    headers = ["Sheet Number", "Sheet Name", "Current Revision", "Revision Number", "Revision Description",
               "Revision Date", "Drawn By", "Issued By"]
    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    # Write the data
    for row_num, row_data in enumerate(data, start=1):
        worksheet.write(row_num, 0, row_data["Sheet Number"])
        worksheet.write(row_num, 1, row_data["Sheet Name"])
        worksheet.write(row_num, 2, row_data["Current Revision"])
        worksheet.write(row_num, 3, row_data["Revision Number"])
        worksheet.write(row_num, 4, row_data["Revision Description"])
        worksheet.write(row_num, 5, row_data["Revision Date"])
        worksheet.write(row_num, 6, row_data["Drawn By"])
        worksheet.write(row_num, 7, row_data["Issued By"])

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
output_dir = forms.pick_folder()
filename = forms.ask_for_string(
    default='ProjectName/Number - Sheet Revisions.xlsx',
    prompt='Enter Project Number',
    title='Revision Manager'
)
output_file = os.path.join(output_dir, filename)

# Export the collected revision information to an Excel file
export_to_excel(all_revisions, output_file)

print("Export completed. The file is saved at: " + output_file)
