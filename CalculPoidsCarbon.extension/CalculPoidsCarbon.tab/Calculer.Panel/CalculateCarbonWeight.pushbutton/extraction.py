#! python3
# -*- coding: utf-8 -*-
from outil import set_path, volume_conv, round_3_decimals, remove_duplicate_dicts

set_path()

import os
import json
import copy
from collections import defaultdict, OrderedDict
import clr
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import (
    BuiltInParameter,
    FilteredElementCollector,
    Phase,
    Transaction,
    ParameterType,
    BuiltInParameterGroup,
    ProjectInfo,
    BuiltInCategory,
    ExternalDefinitionCreationOptions,
    ElementId,
)

import openpyxl

clr.AddReference("System.Windows.Forms")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument.Document

# Get title of revit project document
doc_title = doc.Title

# Create nested dict for material class mapping w/ product code and density
material_mapping = {
    "carton": ("458E62J7FpygLovJB7mRva", 700),
    "verre": ("a7Q7YHJeDQcPyCvYKpqMZW", 2200),
    "metaux": ("AGUFGLX9zzHhVzgNEgpBmL", 7800),
    "plastiques": ("EdUijQTYPFcgrBCqFDzmj4", 1400),
    "acier": ("LoNxZDMDkjP4sdurKhSJSS", 7800),
    "laines_minerales": ("mdnyL8jE8jrCt7840RqzjBfYqU", 800),
    "autres_materiaux": ("PfvLPMfxBtvFJYzVGhYjD6", 2000),
    "aluminium": ("PSQCZvJS3W2NJFtt8e7A2E", 2700),
    "peinture": ("5HWwybW3AFrAaMibHrBQZw", 1200),
    "autres_materiaux_biosources": ("7tippkadbes2pGUdySoHj6", 1200),
    "bois": ("Ac489R6BYsxswScw5ptVUM", 7600),
    "beton": ("aJrbbWGEPxNgEbRSKt9YDR", 2400),
    "platre": ("eKTdoAMBud4QkJURvh5MKK", 800),
    "terre": ("QU3AAosYiLqtJGHzx66DFa", 1400),
    "sable_granulats_roches": ("TBhiVjsFwaaqgc9tPYP92G", 2700),
    "laine_de_bois": ("U6fvT6RaLE3Kwjwm9WBe59", 700),
}

# Material class and material mapping
carton = []
verre = ["Verre", "Glass", "Cristal", "Vidrio"]
metaux = ["Métal", "Metal", "Mtal"]
plastiques = ["Plastique", "Plastic", "Plástico", "Plstico"]
acier = []
laines_minerales = []
autres_materiaux = []
aluminium = []
peinture = ["Peinture/revêtement", "Paint/Coating", "Peindre", "Peinture"]
autres_materiaux_biosources = [
    "Gaz",
    "Liquide",
    "Terre",
    "Gas",
    "Plante",
    "Soil",
    "Terreno",
]
bois = ["Bois", "Madera", "Wood"]
beton = ["Béton", "Concreto", "Concrete", "Hormigón"]
platre = []
terre = []
sable_granulats_roches = [
    "Maçonnerie",
    "Pierre",
    "Masonry",
    "Stone",
    "Céramique",
    "Ceramic",
    "Cerámica",
    "Maconnerie",
]
laine_de_bois = []
materiaux_inconnus = [
    "Pas d'attribution",
    "Lot 11",
    "Lot 09",
    "Lot 07",
    "Lot 04",
    "Générique",
    "Divers",
    "Textile",
    "System",
    "Système",
    "Generic",
    "Miscellaneous",
    "Non attribuée",
    "Unassigned",
    "Genérico",
    "Sin asignar",
    "Sistema",
    "Varios",
]


unit = "kg"
quantity = 0

filtered_elements = []
def compressed_dico_extraction():
    # Filter out those whose family and created phase parameters are None
    for element in FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
        family_param = element.LookupParameter("Famille")
        phase_param = element.LookupParameter("Phase de création")
        
        # Check if both family and phase parameters are not None
        if family_param is not None and phase_param is not None:
            phase = element.get_Parameter(BuiltInParameter.PHASE_CREATED).AsValueString()
            
            # Check if the phase meets the conditions
            if "Exi" not in phase and "EXI" not in phase and element.CreatedPhaseId is not None:
                # Add the element to the filtered list
                filtered_elements.append(element)
                
    has_material_volume = []            
    very_small_number = 0.0001

    # Loop through filtered_elements to find those with material volumes
    for filtered_element in filtered_elements:
        materials = filtered_element.GetMaterialIds(False)
        volume = filtered_element.LookupParameter("Volume")
        for mat in materials:
            if filtered_element.GetMaterialVolume(mat) >= very_small_number:
                has_material_volume.append(filtered_element)
        # Retrieve materials of elements with tangible dependent elements such as Mullions and Panels.
        if len(materials) == 0:
            dep_ids = filtered_element.GetDependentElements(None)
            for dependent_id in dep_ids:
                dependent_id = doc.GetElement(dependent_id)
                materials = dependent_id.GetMaterialIds(False)
                for mat in materials:
                    if dependent_id.GetMaterialVolume(mat) >= very_small_number:
                        has_material_volume.append(dependent_id)

    # Create a new list to discard hand rails and bars as can't calculate for now
    filtered_list = [
        a
        for a in has_material_volume
        if a is not None
        and a.Category is not None
        and a.Category.Name != "Garde-corps"
        and a.Category.Name != "Barreaux"
    ]

    # Update the original list with the filtered list
    has_material_volume = filtered_list

    # Define list to keep contrat cadre elements
    to_remove = []

    # Define list for elements with unknown material classes
    elements_with_unknown_materials = []
    unit = "kg"
    elem_mats = []
    elements_with_unknown_materials = []
    mat_not_in_db = []
    dico = []

    # Remove all contrat cadre elements
    for element in has_material_volume:
        has_symbol = hasattr(element, "Symbol")
        if has_symbol and element.Symbol is not None:
            contrat_cadre_param = element.Symbol.LookupParameter("_CONTRAT_CADRE")
            if (
                contrat_cadre_param is not None
                and contrat_cadre_param.AsValueString() == "Oui"
            ):
                to_remove.append(element)
        # Loop through final filtered list to retrieve data and update dico
        materials = element.GetMaterialIds(False)
        element_id = element.Id
        lot = "N/A"
        for material in materials:
            elem_mat = doc.GetElement(material)
            volume = volume_conv(element.GetMaterialVolume(material))
            sous_projet = element.LookupParameter("Sous-projet").AsValueString()
            if elem_mat.MaterialClass not in materiaux_inconnus:
                try:
                    for material_name, (component_id, density) in material_mapping.items():
                        if elem_mat.MaterialClass in globals()[material_name]:
                            elem_mats.append(component_id)
                            quantity = density * volume
                            break
                    else:
                        mat_not_in_db.append(elem_mat.MaterialClass)
                        quantity = "Material not in database"
                        # Collect the element triggering an alert
                        elements_with_unknown_materials.append(element)
                    dico.append({
                        "sous-projet": sous_projet,
                        "lot": str(lot),
                        "volume": round_3_decimals(volume),
                        "unit": str(unit),
                        "component_id": str(elem_mats[-1]),
                        "element_id": str(element_id),
                        "category": element.Category.Name,
                        "quantity": round_3_decimals(quantity)
                    })
                except:
                    pass
    for item in to_remove:
        has_material_volume.remove(item)
        
    # Display TaskDialog with the list of elements triggering alerts
    if len(elements_with_unknown_materials) > 0:
        TaskDialog.Show(
            "Classe de materiau inconnu",
            "Attention, les matériaux des éléments listés ci-dessous ont des classes inconnues de la base de données de la Calculette Carbone. Par conséquent, ils ne seront pas calculés. Pour éviter des calculs incorrects, veuillez attribuer les classes correctes aux matériaux utilisés.",
        )
        workbook = openpyxl.Workbook()
        worksheet = workbook.active

        # Add headers to the worksheet
        worksheet.append(["Element ID", "Category"])

        # Write elements with unknown materials to the worksheet
        for element in elements_with_unknown_materials:
            worksheet.append([str(element.Id), element.Name])

        # Save the workbook to a file
        workbook.save("materiau.x_classe.s_inconnu.s.xlsx")

    cw_agg_dict = {}
    compressed_dico = []



    dico = remove_duplicate_dicts(dico)
    # Iterate over each dictionary in the dico
    for dictionary in dico:
        sous_projet = dictionary["sous-projet"]
        lot = dictionary["lot"]
        component_id = dictionary["component_id"]
        unit = dictionary["unit"]
        volume = dictionary["volume"]
        element_id = dictionary["element_id"]
        category = dictionary["category"]
        quantity = dictionary["quantity"]
        # Check if there is an existing nested dictionary with the same component_id and element_id
        cw_key = (component_id, element_id)
        # If the key already exists in the aggregated dictionary, add the volume and quantity
        if cw_key in cw_agg_dict:
            cw_agg_dict[cw_key]["volume"] += volume
            cw_agg_dict[cw_key]["quantity"] += quantity
        else:
            # Otherwise, create a new entry in the aggregated dictionary
            cw_agg_dict[cw_key] = {
                "sous-projet": sous_projet,
                "lot": lot,
                "component_id": str(component_id),
                "volume": volume,
                "unit": str(unit),
                "element_id": element_id,
                "category": category,
                "quantity": quantity,
            }
    # Convert the aggregated dictionary back into a list of dictionaries
    result = list(cw_agg_dict.values())
    result = sorted(result, key=lambda x: x["element_id"])
    # Group all elements by component_id and sum up the quantities
    # Create two lists, elements and individual quantities
    # If element is added to a component_id group then append the host_id and quantity simultaneously
    # This simplifies the importation process when we must assign the correct quantity to the correct revit element

    # Initialise projet_dict to organise items with same sous-projet and component_id
    projet_dict = {}

    for item in result:
        projet_key = item["sous-projet"]
        if projet_key not in projet_dict:
            projet_dict[projet_key] = defaultdict(list)
        component_id = item["component_id"]
        projet_dict[projet_key][component_id].append(item)

    compressed_dico = []
    # Compress data into compressed_dico
    for projet_key, component_dict in projet_dict.items():
        for component_id, component_items in component_dict.items():
            combined_item = {
                "sous-projet": projet_key,
                "category": [item["category"] for item in component_items],
                "component_id": component_id,
                "quantity": sum(item["quantity"] for item in component_items),
                "unit": component_items[0]["unit"],
                "elements": [item["element_id"] for item in component_items],
                "individual_quantities": [item["quantity"] for item in component_items],
                "lot": [item["lot"] for item in component_items],
            }

            compressed_dico.append(combined_item)
            
    # Define variables for directories at given paths
    directory = "carbon_data"
    versions_directory = "versions"
    new_dir_path = os.path.normpath(os.path.expanduser("~/Desktop"))
    path = os.path.join(new_dir_path, directory)
    version_path = os.path.join(new_dir_path, directory, versions_directory)

    # Declare boolean variables to check if paths exist
    path_exists = os.path.exists(path)
    version_path_exists = os.path.exists(version_path)

    # If paths do not exist, then create
    if not path_exists:
        os.mkdir(path)
    if not version_path_exists:
        os.mkdir(version_path)
        
    element_ids = [element.Id.IntegerValue for element in has_material_volume]

    with open(os.path.join(path, "element_ids.json"), "w") as f:
        json.dump(element_ids, f)
        
    return compressed_dico

compressed_dico = compressed_dico_extraction()

def create_element_id_dict(compressed_dico):
    # Define dictionaries for code
    element_quantity_dict = {}
    category_quantity_dict = {}
    # Share out the quantities to their respective elements
    for item in compressed_dico:
        elements = item["elements"]
        individual_quantities = item["individual_quantities"]
        categories = item["category"]
        # Iterate through elements and individual_quantities lists simultaneously
        for element_id, quantity, category in zip(
            elements, individual_quantities, categories
        ):
            element_id = str(element_id)  # Convert to string to ensure consistent keys

            # Check if the element ID is already in the dictionary
            if element_id in element_quantity_dict:
                # Add the new quantity to the existing value
                element_quantity_dict[element_id] += quantity
            else:
                # If not, create a new entry with the element ID as the key
                element_quantity_dict[element_id] = quantity
            if category in category_quantity_dict:
                # Add the new quantity to the existing value
                category_quantity_dict[category] += quantity
            else:
                # If not, create a new entry with the element ID as the key
                category_quantity_dict[category] = quantity
    element_quantity_dict = OrderedDict(sorted(element_quantity_dict.items()))
    element_quantity_dict_copy = copy.copy(element_quantity_dict)
    return element_quantity_dict_copy, category_quantity_dict
    