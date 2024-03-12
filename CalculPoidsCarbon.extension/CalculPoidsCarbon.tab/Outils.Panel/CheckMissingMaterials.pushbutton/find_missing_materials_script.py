#! python3
# -*- coding: utf-8 -*-
import clr
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInParameter
import os
doc = __revit__.ActiveUIDocument.Document
clr.AddReference("RevitApi")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import openpyxl

# has_material = IN[0]
# filtered_elements = IN[1]
# Create lists of element families and lists for mapping materials
no_material_search = [
    "Mur de base",
    "Sol",
    "Toit de base",
    "Plafond composé",
    "Plafond de base",
]
material_mapping = {
    "carton": ("458E62J7FpygLovJB7mRva", 700),
    "verre": ("a7Q7YHJeDQcPyCvYKpqMZW", 2179.97),
    "metaux": ("AGUFGLX9zzHhVzgNEgpBmL", 7800),
    "plastiques": ("EdUijQTYPFcgrBCqFDzmj4", 1350),
    "acier": ("LoNxZDMDkjP4sdurKhSJSS", 7849.88),
    "laines_minerales": ("mdnyL8jE8jrCt7840RqzjBfYqU", 800),
    "autres_materiaux": ("PfvLPMfxBtvFJYzVGhYjD6", 2000),
    "aluminium": ("PSQCZvJS3W2NJFtt8e7A2E", 2700),
    "peinture": ("5HWwybW3AFrAaMibHrBQZw", 1190),
    "autres_materiaux_biosources": ("7tippkadbes2pGUdySoHj6", 1200),
    "bois": ("Ac489R6BYsxswScw5ptVUM", 759.99),
    "beton": ("aJrbbWGEPxNgEbRSKt9YDR", 2407.27),
    "platre": ("eKTdoAMBud4QkJURvh5MKK", 800),
    "terre": ("QU3AAosYiLqtJGHzx66DFa", 1400),
    "sable_granulats_roches": ("TBhiVjsFwaaqgc9tPYP92G", 2699.96),
    "laine_de_bois": ("U6fvT6RaLE3Kwjwm9WBe59", 700),
}
no_material_check = [
    "Meneaux de murs-rideaux",
    "Ossature",
    "Plafonds",
    "Poteaux",
    "Sols",
    "Toits",
    "Poteaux porteurs",
    "Equipement spécialisé",
]
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
    "Céramique",
    "Gaz",
    "Liquide",
    "Terre",
    "Ceramic",
    "Gas",
    "Plante",
    "Cerámica",
    "Soil",
    "Terreno",
]
bois = ["Bois", "Madera", "Wood"]
beton = ["Béton", "Concreto", "Concrete", "Hormigón"]
platre = []
terre = []
sable_granulats_roches = ["Maçonnerie", "Pierre", "Masonry", "Stone"]
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
# Create empty lists and define functions needed for wall code
all_elements_v2 = []
filtered_elements = []
mat_not_in_db = []
dico = []
has_mat_volume = []
no_volume = []
cat_mats = []
no_mats_or_has_dep_els = []
lst1 = []
no_material = []
unit = "kg"
quantity = 0

# Define variables for directories at given paths
directory = "carbon_data"
new_dir_path = os.path.normpath(os.path.expanduser("~/Desktop"))
path = os.path.join(new_dir_path, directory)

# Declare boolean variables to check if paths exist
path_exists = os.path.exists(path)

# If paths do not exist, then create
if not path_exists:
    os.mkdir(path)

def volume_conv(volume_in_cubic_foot):
    return volume_in_cubic_foot / 35.3147


def width_conv(width_in_feet):
    return width_in_feet / 3.2808


def round_3_decimals(number):
    return round(number, 3)


all_elements = list(
    FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
)

for e in all_elements:
    if (
        e.LookupParameter("Famille") is not None
        and e.LookupParameter("Phase de création") is not None
    ):
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
        elem_mat = doc.GetElement(mat)
        if item.GetMaterialVolume(mat) >= 0.000000001:
            has_mat_volume.append(item)
            mat_class = elem_mat.MaterialClass
        if (len(materials) == 0 or mat_class in materiaux_inconnus) and hasattr(
            volume, "AsDouble"
        ):
            no_material.append({str(item.Id): str(item.Category.Name)})

updated_material = []

if len(no_material) > 0:
    for dict in no_material:
        updated_dict = (
            {}
        )  # Create a new dictionary to store the updated key-value pairs
        for key, value in dict.items():
            new_value = value.replace("\xe9", "e")
            updated_dict[key] = (
                new_value  # Store the updated key-value pair in the new dictionary
            )
        updated_material.append(
            updated_dict
        )  # Append the updated dictionary to the new list

    # Initialize an empty string to store the missing elements
    missing_elements_str = ""
    missing_materials_list = []
    for a in no_material:
        updated_dict = (
            {}
        )  # Create a new dictionary to store the updated key-value pairs
        for key, value in a.items():
            new_value = value.replace("\xe9", "e")
            updated_dict[key] = (
                new_value  # Store the updated key-value pair in the new dictionary
            )
        missing_elements_str += str(updated_dict) + "\n"
        missing_materials_list.append(list([key, value]))

    # Create a new workbook and select the active worksheet
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    # Add headers to the worksheet
    worksheet.append(["Element ID", "Category"])

    # Iterate through the missing_materials_list and append each item to the worksheet
    for item in missing_materials_list:
        worksheet.append(item)

    file_name = os.path.join(path, "materiau.x_manquant.s.xlsx")

    # Save the workbook to a file
    workbook.save(file_name)

    # Create an instance of TaskDialog
    dialog = TaskDialog("Missing Materials")

    # Set the content of the dialog box
    dialog.MainContent = missing_elements_str

    # Show the dialog box
    dialog.Show()
else:
    TaskDialog.Show(
        "Aucun matériau manquant trouvé", "Vous n'avez aucun élément sans matériaux"
    )
