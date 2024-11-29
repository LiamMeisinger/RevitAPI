__doc__ = "Creates and excel of all current sheets in the model"
__title__ = "Sheet To\nExcel"  #Title of the extension
__author__ = "Liam Meisinger"

# IMPORTS
from Autodesk.Revit.DB import *
import os
import xlsxwriter
import xlrd

#.NET Imports
import clr
from Autodesk.Revit.UI.Selection import Selection
import sys

clr.AddReference('System')

app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document  #type:Document
selection = uidoc.Selection  #type: Selection
active_view = doc.ActiveView
active_level = doc.ActiveView.GenLevel


# Function to get all sheets
def get_all_sheets():
    return FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_Sheets) \
        .WhereElementIsNotElementType() \
        .ToElements()


# Function to get revision details for a sheet
def get_sheet_revisions(sheet):
    revisions = []
    revision_ids = sheet.GetAllRevisionIds()

    for rev_id in revision_ids:
        rev = doc.GetElement(rev_id)

        # Retrieve revision number as a string
        rev_number_param = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_NUM)
        print(rev_number_param.AsString())
        if rev_number_param:
            rev_number = rev_number_param.AsString()
        else:
            rev_number = "No Revision Number"

        # Retrieve revision date
        rev_date_param = rev.get_Parameter(BuiltInParameter.PROJECT_REVISION_REVISION_DATE)
        if rev_date_param:
            rev_date = rev_date_param.AsString()
        else:
            rev_date = "No Revision Date"

        # Store the revision number and date as a tuple
        revisions.append((rev_number, rev_date))

    return revisions


# Collect all sheets in the project
sheets = get_all_sheets()

# List to store all sheet revision information
all_sheet_revisions = []

for sheet in sheets:
    sheet_info = {
        "Sheet Number": sheet.SheetNumber,
        "Sheet Name": sheet.Name,
        "Revisions": get_sheet_revisions(sheet)
    }
    all_sheet_revisions.append(sheet_info)

# Example: Print out revision information for each sheet

print(all_sheet_revisions)


# Function to inspect all parameters of a revision
def inspect_revision_parameters(revision):
    for param in revision.Parameters:
        param_name = param.Definition.Name
        param_value = param.AsString() or param.AsValueString()
        print(param_name, param_value)

def get_sheet_by_number(sheet_number):
    sheets = FilteredElementCollector(doc) \
        .OfCategory(BuiltInCategory.OST_Sheets) \
        .WhereElementIsNotElementType() \
        .ToElements()

    for sheet in sheets:
        if sheet.SheetNumber == sheet_number:
            return sheet
    return None

# Example: Isolate a sheet with a specific sheet number
sheet_number = "3300"  # Replace with the sheet number you want to isolate
isolated_sheet = get_sheet_by_number(sheet_number)



# Example: Use on a single revision element
revision_ids = isolated_sheet.GetAllRevisionIds()  # Assuming some_sheet is a sheet you've selected
if revision_ids:
    first_revision = doc.GetElement(revision_ids[1])
    inspect_revision_parameters(first_revision)


