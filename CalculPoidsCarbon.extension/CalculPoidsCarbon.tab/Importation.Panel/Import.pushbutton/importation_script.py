# -*- coding: utf-8 -*-
import clr

clr.AddReference("RevitApi")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

import pickle
import copy
import json
from collections import OrderedDict
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import ElementId, Transaction, FilteredElementCollector, View, View3D, ParameterFilterElement, ParameterFilterRuleFactory, ElementParameterFilter, ParameterValueProvider, FilterElement, BuiltInParameter, OverrideGraphicSettings, Color, FillPatternElement, BuiltInCategory, AnnotationSymbol, GraphicsStyleType, CategoryType, PhaseFilter, ExternalDefinitionCreationOptions, ParameterType, BuiltInParameterGroup, Category, CategorySet, SharedParameterElement 
from System.Collections.Generic import List
clr.AddReference("DSCoreNodes")
import datetime
from pyrevit import forms
import os
import Autodesk

doc = __revit__.ActiveUIDocument.Document

lst = []

# Define variables for directories at given paths
directory = "carbon_data"
versions_directory = "versions"
new_dir_path = os.path.normpath(os.path.expanduser("~/Desktop"))
path = os.path.join(new_dir_path, directory) 
version_path = os.path.join(new_dir_path, directory, versions_directory)

with open(os.path.join(path, "compressed_dico.pickle"), "rb") as h:
    compressed_dico = pickle.load(h)

with open(os.path.join(path, "import_calc2.pickle"), "rb") as g:
    carbon_results = pickle.load(g)

with open(os.path.join(path, "element_ids.json"), "r") as f:
    element_ids = json.load(f)
has_material_volume = [doc.GetElement(ElementId(e)) for e in element_ids]

list_of_results = carbon_results

categories = []

for el in has_material_volume:
    if el is not None and el.Category is not None:
        cat = el.Category
        if cat not in categories:
            categories.append(cat)

def SetInstanceParamVaryBetweenGroupsBehaviour(doc, guid, allowVaryBetweenGroups=True):

    sp = SharedParameterElement.Lookup(doc, guid)

    # Should never happen as we will call 
    # this only for *existing* shared param.
    if sp is None:
        return

    defn = sp.GetDefinition()

    if defn.VariesAcrossGroups != allowVaryBetweenGroups:
        # Must be within an outer transaction!
        t = Transaction(doc, "Vary between groups")
            # Start the transaction
        t.Start()
    
        # Set the 'Allow Vary Between Groups' option
        defn.SetAllowVaryBetweenGroups(doc, allowVaryBetweenGroups)
        # Commit the transaction
        t.Commit()
        # Rollback the transaction in case of an exception
        t.RollBack()
                

# def create_icc_parameter(doc):
#     """
#     Create a new instance parameter in a Revit document.

#     Parameters:
#         doc (Document): The Revit document in which the parameter will be created.

#     Returns:
#         str: A message indicating the result of the operation. 
#               Returns "Parameter created successfully." if successful.
#               Returns an error message if an exception occurs or if the shared parameter file is not found.
#     """
    
#     # Start a transaction
#     t = Transaction(doc, "Create Instance Parameter")
#     t.Start()
 
#     # # Get or create the shared parameter file
#     app = doc.Application
#     sharedParameterFile = app.OpenSharedParameterFile()
 
#     # Define the parameter group and name
#     param_group = "EMC2B"  # Change the parameter group as per your requirement
#     param_name = '_POIDS_CARBONE'
 
#     # Create an ExternalDefinitionCreationOptions object
#     options = ExternalDefinitionCreationOptions(param_name, ParameterType.Text)
 
#     # Create or get the external definition
#     definition_group = sharedParameterFile.Groups.Create(param_group)
#     definition = definition_group.Definitions.Create(options)

#     # Bind the parameter as an instance parameter
#     category_set = CategorySet()
#     for item in categories:
#         category_set.Insert(item)

#     instanceBinding = app.Create.NewInstanceBinding(category_set)
#     try:
#         doc.ParameterBindings.Insert(definition, instanceBinding, BuiltInParameterGroup.PG_IDENTITY_DATA)
#     except Exception as e:
#         print(e)
#     # Commit the transaction
#     t.Commit()
#     return "Parameter created successfully."

shared_param_collector = FilteredElementCollector(doc).OfClass(SharedParameterElement).ToElements()

for shared_param_el in shared_param_collector:
    if shared_param_el.Name == "_POIDS_CARBONE":
        poids_carbone_guid = shared_param_el.GuidValue
    
# shared_parameters = [param.Name for param in shared_param_collector]
# if not '_POIDS_CARBONE' in shared_parameters:
#     create_icc_parameter(doc)
SetInstanceParamVaryBetweenGroupsBehaviour(doc, poids_carbone_guid, allowVaryBetweenGroups=True)

# Loop through the calculette results to replace "é" with "e".
for dictt in list_of_results:
    updated_dict = {}
    for key, value in dictt.items():
        new_key = key.replace("é", "e")
        updated_dict[new_key] = value
    list_of_results[list_of_results.index(dictt)] = updated_dict

# Loop through compressed_dico and list_of_results simultaneously to calculate the carbon weight per element using the percentages previously calculated    
for dico in compressed_dico:
    for result in list_of_results:
        if dico['quantity'] == result['Quantite']:
            for i in range(len(dico["individual_quantities"])):
                r = dico["individual_quantities"][i]
                element_carbon_weight = r * result["Impact sur le changement climatique (kgCO2e)"]
                dico["individual_quantities"][i] = element_carbon_weight

element_quantity_dict = {}

for item in compressed_dico:
    elements = item["elements"]
    individual_quantities = item["individual_quantities"]

    # Iterate through elements and individual_quantities lists simultaneously
    for element_id, quantity in zip(elements, individual_quantities):
        element_id = str(element_id)  # Convert to string to ensure consistent keys

        # Check if the element ID is already in the dictionary
        if element_id in element_quantity_dict:
            # Add the new quantity to the existing value
            element_quantity_dict[element_id] += quantity
        else:
            # If not, create a new entry with the element ID as the key
            element_quantity_dict[element_id] = quantity
element_quantity_dict = OrderedDict(sorted(element_quantity_dict.items()))
element_quantity_dict_copy = copy.copy(element_quantity_dict)
for key, value in element_quantity_dict_copy.items():
    if value < 1:
        lst.append({value: key})
        del element_quantity_dict[key]
# Append the category ids of all elements calculated to material_volume_collection
material_volume_collection = [element.Category.Id for element in has_material_volume if element is not None and element.Category is not None]

# Make an ICollection list for the api
material_volume_collection_net = List[ElementId](material_volume_collection)


# Create values for each filter
parameter_name = "_POIDS_CARBONE"
no_value = "0"
very_small_value = "0.000001"
small_value = "5000"
medium_value = "35000"
large_value = "250000"
very_large_value = "1000000"

#Declare filters_added as False
filters_added = False

#Collect filters before creating new ones
pre_filters_collector = list(FilteredElementCollector(doc).OfClass(FilterElement))

# Make filters_added true if filters already in project
for filterr in pre_filters_collector:
    if "POIDS_CARBON" in filterr.Name:
        filters_added = True
        
# Create filters if not in project        
if not filters_added:
    for elem in has_material_volume:
        if elem is not None and elem.LookupParameter("_POIDS_CARBONE") is not None:
            parameter_id = elem.LookupParameter("_POIDS_CARBONE").Id
            
            very_large_rule = ParameterFilterRuleFactory.CreateGreaterRule(parameter_id, very_large_value, False)
            large_rule = ParameterFilterRuleFactory.CreateGreaterRule(parameter_id, large_value, False)
            medium_rule = ParameterFilterRuleFactory.CreateGreaterRule(parameter_id, medium_value, False)
            small_rule = ParameterFilterRuleFactory.CreateGreaterRule(parameter_id, small_value, False)
            very_small_rule = ParameterFilterRuleFactory.CreateGreaterRule(parameter_id, very_small_value, False)
            
            transaction = Transaction(doc, "Create Filters")
            transaction.Start()
            # Create each filter for different sizes and apply the correct rule
            very_large_filter = ParameterFilterElement.Create(doc, "C_05_TRES_GRAND_POIDS_CARBON", material_volume_collection_net)
            apply_very_large_rule = ElementParameterFilter(very_large_rule)
            very_large_filter.SetElementFilter(apply_very_large_rule)
            
            large_filter = ParameterFilterElement.Create(doc, "C_04_GRAND_POIDS_CARBON", material_volume_collection_net)
            apply_large_rule = ElementParameterFilter(large_rule)
            large_filter.SetElementFilter(apply_large_rule)
            
            medium_filter = ParameterFilterElement.Create(doc, "C_03_MOYEN_POIDS_CARBON", material_volume_collection_net)
            apply_medium_rule = ElementParameterFilter(medium_rule)
            medium_filter.SetElementFilter(apply_medium_rule)
            
            small_filter = ParameterFilterElement.Create(doc, "C_02_PETIT_POIDS_CARBON", material_volume_collection_net)
            apply_small_rule = ElementParameterFilter(small_rule)
            small_filter.SetElementFilter(apply_small_rule)
            
            very_small_filter = ParameterFilterElement.Create(doc, "C_01_TRES_PETIT_POIDS_CARBON", material_volume_collection_net)
            apply_very_small_rule = ElementParameterFilter(very_small_rule)
            very_small_filter.SetElementFilter(apply_very_small_rule)

            transaction.Commit()
            break
        
# Collect filters after creation
filters_collector = list(FilteredElementCollector(doc).OfClass(FilterElement))        
filters = []
carbon_filter_ids = []
carbon_filter_names = []

# Extract filter IDs and names
for filterr in filters_collector:
    filters.append({filterr.Name: filterr.Id})
for filter_dict in filters:
    for kee, valu in filter_dict.items():
        if "CARBON" in kee:
            carbon_filter_ids.append(valu)
            carbon_filter_names.append(kee)
            
# Create view and view template
view_collector = FilteredElementCollector(doc).OfClass(View3D)
view_names = []

# Ensure that only the appropriate views are available for selection
for vu in view_collector:
    if "{" in vu.Name:
        view_3dee = vu
        break

# Create new isometric 3Dview
transaction = Transaction(doc, "Create view")
transaction.Start()
default_3d_view = View3D.CreateIsometric(doc, view_3dee.GetTypeId())
transaction.Commit()


# Collect phase filters
phase_filters = FilteredElementCollector(doc).OfClass(PhaseFilter).ToElements()
phase_filter_names = []

# Get filter names
for flter in phase_filters:
    phase_filter_names.append(flter.Name)

# Allow user to choose which phase to present in the view template
phase_selected = forms.SelectFromList.show(phase_filter_names, button_name='Choissisez le filtre', multiselect=False)
view_templates = FilteredElementCollector(doc).OfClass(View3D).WhereElementIsNotElementType().ToElements()

# Declare new_temp_name
new_temp_name = None

# Change new_temp_name if user chooses a view template phase not already calculated
if phase_selected:
    for temp in view_templates:
        if temp.Name == "CLC_CARBON_" + phase_selected.upper():
            break
    else:
        new_temp_name = "CLC_CARBON_" + phase_selected.upper()
    
    # If user chooses a new view template phase
    if new_temp_name:
        transaction = Transaction(doc, "Create view template")
        transaction.Start()
        
        # Create view template
        calc_carbon = default_3d_view.CreateViewTemplate()
        calc_carbon.Name = "CLC_CARBON_" + phase_selected.upper()
        
    
        template_id = calc_carbon.Id
        template = doc.GetElement(template_id)
        template_parameter_list = template.GetTemplateParameterIds()
        
        # Add and enable carbon filters
        for filt in carbon_filter_ids:
            template.AddFilter(filt)
            template.SetIsFilterEnabled(filt, True)
        transaction.Commit()
        
        # Get visibility patterns
        all_patterns  = FilteredElementCollector(doc).OfClass(FillPatternElement).ToElements()
        solid_pattern = [i for i in all_patterns if i.GetFillPattern().IsSolidFill][0]
        
        # Set visibility colours and patterns depending on carbon weight
        for filt_id in template.GetFilters():
            transaction = Transaction(doc, "Create colour visibility")
            transaction.Start()
            filt_name = doc.GetElement(filt_id).Name
            if "1" in filt_name:
                ogs = OverrideGraphicSettings()
                green = Color(80, 197, 58)
                colour = ogs.SetSurfaceForegroundPatternColor(green)
                pattern = ogs.SetSurfaceForegroundPatternId(solid_pattern.Id)
                template.SetFilterOverrides(filt_id, colour)
            elif "2" in filt_name:
                ogs = OverrideGraphicSettings()
                yellow_green = Color(169, 197, 58)
                colour = ogs.SetSurfaceForegroundPatternColor(yellow_green)
                pattern = ogs.SetSurfaceForegroundPatternId(solid_pattern.Id)
                template.SetFilterOverrides(filt_id, colour)
            elif "3" in filt_name:
                ogs = OverrideGraphicSettings()
                yellow = Color(197, 193, 58)
                colour = ogs.SetSurfaceForegroundPatternColor(yellow)
                pattern = ogs.SetSurfaceForegroundPatternId(solid_pattern.Id)
                template.SetFilterOverrides(filt_id, colour)
            elif "4" in filt_name:
                ogs = OverrideGraphicSettings()
                orange = Color(197, 128, 58)
                colour = ogs.SetSurfaceForegroundPatternColor(orange)
                pattern = ogs.SetSurfaceForegroundPatternId(solid_pattern.Id)
                template.SetFilterOverrides(filt_id, colour)
            elif "5" in filt_name:
                ogs = OverrideGraphicSettings()
                red = Color(197, 66, 58)
                colour = ogs.SetSurfaceForegroundPatternColor(red)
                pattern = ogs.SetSurfaceForegroundPatternId(solid_pattern.Id)
                template.SetFilterOverrides(filt_id, colour)
                lines = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).ToElements()
            transaction.Commit()
        # Get and hide category annotation 
        transaction = Transaction(doc, "Hide annotation")
        transaction.Start()
        categorys = doc.Settings.Categories
        for cat in categorys:
            if cat.CategoryType == CategoryType.Annotation or cat.Name == "Liens RVT":
                try:
                    template.SetCategoryHidden(cat.Id, True)
                except:
                    pass
        # Set the filter picked by the user previously    
        for fil in template.Parameters:
            if str(fil.Definition.Name) == "Filtre des phases":
                for flter in phase_filters:
                    if str(phase_selected) == flter.Name:                   
                        fil.Set(flter.Id)
                    
        transaction.Commit()
    else:
        pass
    
    # def save_current_state():
    #     with open("current_state.pickle", "wb") as f:
    #         pickle.dump(carbon_results, f)
    # try:
    #     with open("current_state.pickle", "rb") as f:
    #             previous_state = pickle.load(f)
    # except:
    #     previous_state = None
    
    # Save the current state
    # save_current_state()

    # if previous_state is not None and previous_state != carbon_results:
        t = Transaction(doc, "Update _POIDS_CARBONE Parameters")
        t.Start()

        # Iterate through elements in has_material_volume
        for element in has_material_volume:
            has_id = hasattr(element, "Id")
            if has_id:
                element_id_str = str(element.Id)
                if element_id_str in element_quantity_dict:
                    # Get the summed individual quantity from the dictionary
                    summed_quantity = str(element_quantity_dict[element_id_str])
                    
                    # Update the "_POIDS_CARBONE" parameter of the element
                    try:
                        parameter_name = "_POIDS_CARBONE"
                        param = element.LookupParameter(parameter_name)

                        if param is not None and not param.IsReadOnly:
                            param.Set(summed_quantity)

                    except Autodesk.Revit.Exceptions.ReadOnlyException:
                        # Handle read-only parameter (skip the element)
                        print("Skipping read-only element: {element.Id}")

                    except Exception as e:
                        # Handle other exceptions
                        print("Erreur lors de la mise à jour de l'élément {element.Id} : {str(e)}. Veuillez réessayer ultérieurement.")
        t.Commit()

        
        if default_3d_view:
            current_datetime = datetime.datetime.now().strftime("%y%m%d_%Hh%M")
            # Desired name for the new 3D view
            desired_name = "CALCUL_CARBON_{}".format(current_datetime)
        
            # Check if a 3D view with the desired name already exists
            existing_3d_views = FilteredElementCollector(doc).OfClass(View3D).ToElements()
            existing_3d_view_names = [view.Name for view in existing_3d_views]
        
            new_3d_view_name = desired_name
            counter = 1
        
            while new_3d_view_name in existing_3d_view_names:
                # If a 3D view with the same name exists, add a number to the name
                new_3d_view_name = "{0} {1}".format(desired_name, counter)
                counter += 1
        
            # Start a new transaction to duplicate the 3D view and set the unique name
            with Transaction(doc, "Duplicate and Activate 3D View") as transaction:
                transaction.Start()
        
                # Duplicate the default 3D view
                new_3d_view = View3D.CreatePerspective(doc, default_3d_view.GetTypeId())
        
                # Set the unique name for the new 3D view
                new_3d_view.Name = new_3d_view_name
        
                # Commit the transaction to save the new 3D view
                transaction.Commit()
            with Transaction(doc, "Activate template") as transaction:
                transaction.Start()            
                template_name = "CLC_CARBON_" + phase_selected.upper()  # Replace with the name of your view template
    
                # Find the view template by name
                template = None
                for view_template in FilteredElementCollector(doc).OfClass(View):
                    if view_template.IsTemplate and view_template.Name == template_name:
                        template = view_template
                        break
                
                if template is not None:
                    new_3d_view.ViewTemplateId = template.Id
                else:
                    for view_template in FilteredElementCollector(doc).OfClass(View):
                        if view_template.IsTemplate:
                            pass
                transaction.Commit()
            # Open the new 3D view
            __revit__.ActiveUIDocument.ActiveView = new_3d_view
        else:
            TaskDialog.Show("Aucune vue 3D trouvée", "Aucune vue 3D trouvée.")
    # else:
    #     # Display a simplified message box to ask the user to recalculate the values
    #     alert = TaskDialog.Show("La data n'a pas changée", "Veuillez recalculer le poids de vos éléments avant de vouloir importer vos changements.")
else:
    TaskDialog.Show("Filter de phase manquant", "Veuillez choisir un filtre de phase pour completer l'importation des valeurs.")
