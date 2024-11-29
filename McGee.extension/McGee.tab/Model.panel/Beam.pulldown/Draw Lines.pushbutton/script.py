# -*- coding: utf-8 -*-
__doc__ = "Creating Beams"
__title__ = "Draw"  #Title of the extension
__author__ = "Liam Meisinger"

# IMPORTS
from Autodesk.Revit.DB import *

#.NET Imports
import clr
from Autodesk.Revit.DB.Structure import StructuralType
from Autodesk.Revit.UI.Selection import ISelectionFilter, Selection, ObjectType

clr.AddReference('System')
from System.Collections.Generic import List

# VARIABLES
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document  #type:Document
selection = uidoc.Selection  #type: Selection
active_view = doc.ActiveView
active_level = doc.ActiveView.GenLevel

"""CREATE BEAM"""
# Select Lines
def select_model_lines():
    references = selection.PickObjects(ObjectType.Element, "Select a model line to define the beam path")
    model_lines = []
    for reference in references:
        element = doc.GetElement(reference)
        if isinstance(element, CurveElement):
            model_lines.append(element.GeometryCurve)
        else:
            raise ValueError("One of the selected elements is not a model line")
    return model_lines


lines = select_model_lines()


# Get Default Beam Type
beam_type_id = doc.GetDefaultFamilyTypeId(ElementId(BuiltInCategory.OST_StructuralFraming))
beam_type = doc.GetElement(beam_type_id)
level = FilteredElementCollector(doc).OfClass(Level).FirstElement()

# Create Beam
t = Transaction(doc, 'Create Beam')
t.Start()
for line in lines:
    beam = doc.Create.NewFamilyInstance(line, beam_type, level, StructuralType.Beam)
    print('Created Beams: {}'.format(beam.Id))
t.Commit()


