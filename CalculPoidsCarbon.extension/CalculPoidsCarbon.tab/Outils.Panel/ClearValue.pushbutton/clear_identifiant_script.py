# -*- coding: utf-8 -*-
import math
import pickle
import copy
import urllib2 as ur
import collections, functools, operator
from itertools import groupby
import json
import os
from System.Net import *
from collections import defaultdict, Counter, OrderedDict
import sys
import clr
from Autodesk.Revit.DB import *
clr.AddReference("ProtoGeometry")
from Autodesk.DesignScript.Geometry import *
clr.AddReference("RevitNodes")
import Revit
clr.ImportExtensions(Revit.Elements)
clr.ImportExtensions(Revit.GeometryConversion)
from Revit.Elements import*
clr.AddReference("RevitServices")
import RevitServices
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
import System
clr.AddReference("DSCoreNodes")
import DSCore
from DSCore import*
doc = __revit__.ActiveUIDocument.Document
clr.AddReference("RevitApi")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import subprocess
# Start a transaction
doc = __revit__.ActiveUIDocument.Document

all_elements_v2 = []
filtered_elements = []
has_mat_volume = []
no_material = []
lst1 = []
# Get the parameter by name (change this to your parameter name)
param_name = "_POIDS_CARBONE"

all_elements = list(FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements())

for e in all_elements:
	if e.LookupParameter("Famille") is not None and e.LookupParameter("Phase de création") is not None:
		all_elements_v2.append(e)

for el in all_elements_v2:
	phase = el.get_Parameter(BuiltInParameter.PHASE_CREATED).AsValueString()
	if "Exi" not in phase and "EXI" not in phase and el.CreatedPhaseId != None:
		filtered_elements.append(el)

# Append the correct elements to their lists based on if they have dependent elements, layers or neither
for item in filtered_elements:
    materials = item.GetMaterialIds(False)
    volume = item.LookupParameter("Volume")
    famille = item.LookupParameter("Famille").AsValueString()
    for mat in materials:
        if item.GetMaterialVolume(mat) >= 0.000000001:
            has_mat_volume.append(item)
    if len(materials) == 0:
    	no_material.append(item)
    	dep_ids = item.GetDependentElements(None)
    	for dep in dep_ids:
    		dep = doc.GetElement(dep)
    		materials = dep.GetMaterialIds(False)
    		for mat in materials:
    			if dep.GetMaterialVolume(mat) >= 0.000000001:
    				has_mat_volume.append(item)
    			elif dep not in no_material:
    				no_material.append(dep)

for x in has_mat_volume:
	lst1.append(str(x.Id))
	has_stairs = hasattr(x, "GetStairs")
	if has_stairs:
		has_mat_volume.Remove(x)

param_name = "_POIDS_CARBONE"
t = Transaction(doc, "Clear _POIDS_CARBONE Parameter")
t.Start()		
# Clear the parameter value for each element
try:
    for element in has_mat_volume:
        param = element.LookupParameter(param_name)
        if param is not None:
            if not param.IsReadOnly:
                param.Set("")  # Clear the parameter value
    # Commit the transaction
except Exception as e:
    # Handle exceptions or errors here
    print("An error occurred:", str(e))
t.Commit()

# End the transaction
t.Dispose()
