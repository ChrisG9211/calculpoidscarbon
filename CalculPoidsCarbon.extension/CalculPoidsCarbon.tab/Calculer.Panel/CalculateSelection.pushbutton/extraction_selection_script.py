#! python3
# -*- coding: utf-8 -*-
from outil import install_packages

install_packages()
import sys
import site
site.addsitedir(site.getusersitepackages(), known_paths=None)
# Now, you can use site.getsitepackages()
sys.path.extend(site.getusersitepackages())
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import urllib.request as ur
import json
import pickle
from collections import defaultdict, OrderedDict
import copy
import clr
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import BuiltInParameter
from pypac import get_pac, PACSession
import datetime
import openpyxl
import os
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import Form, TextBox, Label, Button, Application
import System
from System.Drawing import Point
import requests


clr.AddReference("RevitApi")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

# Create lists of element families and lists for mapping materials
no_material_search = ["Mur de base", "Sol", "Toit de base", "Plafond composé", "Plafond de base"]

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
    "laine_de_bois": ("U6fvT6RaLE3Kwjwm9WBe59", 700)
}

# Create list of element category names to check if no materials
no_material_check = ["Meneaux de murs-rideaux", "Ossature", "Plafonds", "Poteaux", "Sols", "Toits", "Poteaux porteurs", "Equipement spécialisé"]

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
autres_materiaux_biosources = ["Gaz", "Liquide", "Terre", "Gas", "Plante", "Soil", "Terreno"]
bois = ["Bois", "Madera", "Wood"]
beton = ["Béton", "Concreto", "Concrete", "Hormigón"]
platre = []
terre = []
sable_granulats_roches = ["Maçonnerie", "Pierre", "Masonry", "Stone", "Céramique", "Ceramic", "Cerámica", "Maconnerie"]
laine_de_bois = []
materiaux_inconnus = ["Pas d'attribution", "Lot 11", "Lot 09", "Lot 07", "Lot 04", "Générique", "Divers", "Textile", "System", "Système", "Generic", "Miscellaneous", "Non attribuée", "Unassigned", "Genérico", "Sin asignar", "Sistema", "Varios"] 

# Create empty lists and define functions needed for code
filtered_by_family = []
filtered_elements = []
mat_not_in_db = []
dico = []
lst = []
has_material_volume = []
elem_mats = []
no_material = []
unit = "kg"
quantity = 0

def volume_conv(volume_in_cubic_foot):
    """
    converts cubic feet to cubic meters

    Parameters
    ----------
    volume_in_cubic_foot : int
        to be converted to cubic meters.

    Returns
    -------
    int
        number after conversion from cubic feet to cubic meters.

    """
    return volume_in_cubic_foot / 35.3147

def round_3_decimals(number):
    """
    

    Parameters
    ----------
    number : TYPE
        DESCRIPTION.

    Returns
    -------
    TYPE
        DESCRIPTION.

    """
    
    """ rounds number to 3 decimals
    :param number: int
    :return: number rounded to 3 decimal points   
    """
    return round(number, 3)

# Get selected elements in the project
selection = uidoc.Selection
selected_elements = [doc.GetElement(id) for id in selection.GetElementIds()]

# Filter out those whose family and created phase parameters are None
for selected_element in selected_elements:
    if selected_element.LookupParameter("Famille") is not None and selected_element.LookupParameter("Phase de création") is not None:
        filtered_by_family.append(selected_element)

# Filter out existing
message_shown = False
for element in filtered_by_family:
    phase = element.get_Parameter(BuiltInParameter.PHASE_CREATED).AsValueString()
    if "Exi" not in phase and "EXI" not in phase and element.CreatedPhaseId is not None:
        filtered_elements.append(element)
    elif not message_shown:
        TaskDialog.Show("Attention !", "Attention, un ou plusieurs éléments que vous avez sélectionnés se trouvent dans une phase existante et ne seront donc pas calculés.")
        message_shown = True

# Loop through filtered_elements to find those with material volumes
for filtered_element in filtered_elements:
    materials = filtered_element.GetMaterialIds(False)
    volume = filtered_element.LookupParameter("Volume")
    famille = filtered_element.LookupParameter("Famille").AsValueString()
    for mat in materials:
        if filtered_element.GetMaterialVolume(mat) >= 0.000000001:
            has_material_volume.append(filtered_element)
    if len(materials) == 0 and filtered_element.Category.Name in no_material_check and hasattr(volume, "AsDouble"):
        no_material.append(str(filtered_element.Id))
    # Retrieve materials of elements with tangible dependent elements such as Mullions and Panels.
    elif len(materials) == 0:
        dep_ids = filtered_element.GetDependentElements(None)
        for dependent_id in dep_ids:
            dependent_id = doc.GetElement(dependent_id)
            materials = dependent_id.GetMaterialIds(False)
            for mat in materials:
                if dependent_id.GetMaterialVolume(mat) >= 0.000000001:
                    has_material_volume.append(dependent_id)

# has_material_volume = [x for x in has_material_volume if not (hasattr(x, "GetStairs") or x.Category.Name == "Luminaires")]

to_remove = []

for contrat_cadre in has_material_volume:
    has_symbol = hasattr(contrat_cadre, "Symbol")
    if has_symbol and contrat_cadre.Symbol is not None:
        contrat_cadre_param = contrat_cadre.Symbol.LookupParameter("_CONTRAT_CADRE")
        if contrat_cadre_param is not None and contrat_cadre_param.AsValueString() == "Oui":
            to_remove.append(contrat_cadre)

for item in to_remove:
    has_material_volume.remove(item)

elements_with_unknown_materials = []

# Loop through final filtered list to retrieve data and update dico

for elem in has_material_volume:
    materials = elem.GetMaterialIds(False)
    element_id = elem.Id
    lot = "N/A"
    for material in materials:
        elem_mat = doc.GetElement(material)
        volume = volume_conv(elem.GetMaterialVolume(material))
        sous_projet = elem.LookupParameter("Sous-projet").AsValueString()
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
                    elements_with_unknown_materials.append(elem)
                dico.append({
                    "sous-projet": sous_projet,
                    "lot": str(lot),
                    "volume": round_3_decimals(volume),
                    "unit": str(unit),
                    "component_id": str(elem_mats[-1]),
                    "element_id": str(element_id),
                    "category": elem.Category.Name,
                    "quantity": round_3_decimals(quantity)
                })
            except Exception as e:
                # Handle exceptions appropriately
                # You might want to log the exception details for debugging
                pass

# Display TaskDialog with the list of elements triggering alerts
if len(elements_with_unknown_materials) > 0:
    TaskDialog.Show("Classe de materiau inconnu", "Attention, les matériaux des éléments listés ci-dessous ont des classes inconnues de la base de données de la Calculette Carbone. Par conséquent, ils ne seront pas calculés. Pour éviter des calculs incorrects, veuillez attribuer les classes correctes aux matériaux utilisés.")
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    # Add headers to the worksheet
    worksheet.append(["Element ID", "Category"])

    # Write elements with unknown materials to the worksheet
    for element in elements_with_unknown_materials:
        worksheet.append([str(element.Id), element.Category.Name])

    # Save the workbook to a file
    workbook.save("materiau.x_classe.s_inconnu.s.xlsx")

alert_executed = False  # Flag to track if the alert has been executed

for component in has_material_volume:
    materials = component.GetMaterialIds(False)
    for solid in materials:
        elem_mat = doc.GetElement(solid)
        if elem_mat.MaterialClass in materiaux_inconnus and not alert_executed:
            TaskDialog.Show("Matériel Inconnu", "Impossible d'attribuer une classe de matériau à un ou plusieurs éléments sélectionnés. Assurez-vous que les classes de matériaux attribuées aux matériaux sont aussi précises que possible.")
            alert_executed = True  # Set the alert flag to True
    if alert_executed:
        break

cw_agg_dict = {}

# Function to remove duplicates
def remove_duplicate_dicts(lst):
    """
    removes duplicate dictionaries from param lst

    Parameters
    ----------
    lst : list
        a list in need of duplicate removal.

    Returns
    -------
    unique_dicts : a list with duplicates removed.

    """
    unique_dicts = []
    seen_dicts = set()
    for d in lst:
        dict_items = tuple(sorted(d.items()))
        if dict_items not in seen_dicts:
            unique_dicts.append(d)
            seen_dicts.add(dict_items)

    return unique_dicts

dico = remove_duplicate_dicts(dico)
# Iterate over each dictionary in the dico
for i, dictionary in enumerate(dico):
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
            "quantity": quantity
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

# Calculate percentage of individual quantities relative to total quantity for each item in compressed_dico
# If quantity is 0, remove from compressed_dico
components = {}
counter = 10
for i, q in enumerate(compressed_dico):
    if q["quantity"] != 0:
        for j in range(len(q["individual_quantities"])):
            r = q["individual_quantities"][j]
            percentage = r / q["quantity"]
            q["individual_quantities"][j] = percentage
    elif q["quantity"] == 0.0:
        compressed_dico.remove(q)
    data = {"{} {}".format(q["sous-projet"], counter): {"Hello": {"unit": q["unit"], "quantity": q["quantity"], "product_id": q["component_id"]}}}
    components.update(data)
    counter += 1

# Define variables for directories at given paths
directory = "carbon_data"
new_dir_path = os.path.normpath(os.path.expanduser("~/Desktop"))
path = os.path.join(new_dir_path, directory) 


# Declare boolean variables to check if paths exist
path_exists = os.path.exists(path)

# If paths do not exist, then create
if not path_exists:
    os.mkdir(path)

response_data = None
# Loop through components
req_params = {'persist':False, 'project_lifetime': 50,'method': 'dynamic', 'components': components}
# Set up proxy handler and opener
url = 'http://192.168.96.212:81/api/v2/get_component_impacts'
r = requests.post(url, json=req_params)
r.raise_for_status()  # Raise an HTTPError for bad responses (4xx or 5xx)
# Retrieve data
results = r.json()

try:
    response_data = results["impacts"]
except Exception as e:
    TaskDialog.Show("Fonctionnalité maintenance.", "Cette fonction est actuellement indisponible. Veuillez réessayer ultérieurement.")

if response_data:
    # Dump results so that final_extraction.py can complete the importation process
    with open(os.path.join(path, "import_calc2.pickle"), "wb") as g:
        pickle.dump(response_data, g, protocol=0)
        
    with open(os.path.join(path, "compressed_dico.pickle"), "wb") as h:
        pickle.dump(compressed_dico, h, protocol=0)
        
    element_ids = [element.Id.IntegerValue for element in has_material_volume]
    
    with open(os.path.join(path, "element_ids.json"), "w") as f:
        json.dump(element_ids, f)
    
    # Define variables for directories at given paths
    directory = "carbon_data"
    new_dir_path = os.path.normpath(os.path.expanduser("~/Desktop"))
    path = os.path.join(new_dir_path, directory)
    
    # Declare boolean variables to check if paths exist
    path_exists = os.path.exists(path)
    
    # If paths do not exist, then create
    if not path_exists:
        os.mkdir(path)
    
    list_of_results = results["impacts"]
    
    if len(selected_elements) > 0 and len(dico) > 0:
        # Loop through the calculette results to replace "é" with "e".
        for dictionary in list_of_results:
            updated_dict = {}
            for key, value in dictionary.items():
                new_key = key.replace("é", "e")
                updated_dict[new_key] = value
            list_of_results[list_of_results.index(dictionary)] = updated_dict
        
        # Loop through compressed_dico and list_of_results simultaneously to calculate the carbon weight per element using the percentages previously calculated    
        for dico in compressed_dico:
            for result in list_of_results:
                if dico['quantity'] == result['Quantite']:
                    for i in range(len(dico["individual_quantities"])):
                        percentage = dico["individual_quantities"][i]
                        element_carbon_weight = percentage * result["Impact sur le changement climatique (kgCO2e)"]
                        dico["individual_quantities"][i] = element_carbon_weight
        
        element_quantity_dict = {}
        category_quantity_dict = {}
        for item in compressed_dico:
            elements = item["elements"]
            individual_quantities = item["individual_quantities"]
            categories = item["category"]
            # Iterate through elements and individual_quantities lists simultaneously
            for element_id, quantity, category in zip(elements, individual_quantities, categories):
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
        for a, b in element_quantity_dict_copy.items():
            if b < 1:
                lst.append({b: a})
                del element_quantity_dict[a]
               
        total = 0
        mcrcom = None
        prod = None
        impact = None
        sous_projet_totals = {}
        material_totals = {}
        material_par_sp = {}
        user_answer = None
        
        # Function to create a simple form with TextBox
        class QuestionForm(Form):
            def __init__(self):
                self.Text = "Nom de selection"

                # Label
                self.label = Label()
                self.label.Text = "Veuillez donner un nom à cette selection:"
                self.label.Location = System.Drawing.Point(25, 30)
                self.label.AutoSize = True
                self.Controls.Add(self.label)

                # TextBox
                self.textBox = TextBox()
                self.textBox.Location = System.Drawing.Point(40, 60)
                self.textBox.Width = 200
                self.Controls.Add(self.textBox)

                # Button
                self.button = Button()
                self.button.Text = "Submit"
                self.button.Location = System.Drawing.Point(105, 90)
                self.button.Click += self.button_click
                self.Controls.Add(self.button)
                self.AcceptButton = self.button

            def button_click(self, sender, event):
                global user_answer
                user_answer = self.textBox.Text
                if user_answer != "":
                    self.Close()

        if __name__ == "__main__":
            form = QuestionForm()
            Application.Run(form)

        if user_answer != None:
            for result in results["impacts"]:
                total_value = float(result['Impact sur le changement climatique (kgCO2e)'])
            
                # Convert total to a string if it's an integer
                if isinstance(total, int):
                    total = str(total)
                
                # Remove commas
                total = total.replace(',', '')
                
                # Format the result with commas
                total = '{:,.3f}'.format(float(total) + total_value)
                mcrcom = result["Macro-composant de niveau 1"]
                mcrcom = mcrcom[:-3]
                prod = result["Produit"]
                impact = round_3_decimals(result['Impact sur le changement climatique (kgCO2e)'])
                
                if mcrcom not in sous_projet_totals:
                    sous_projet_totals.update({mcrcom: round_3_decimals(result['Impact sur le changement climatique (kgCO2e)'])})
                else:
                    sous_projet_totals[mcrcom] += round_3_decimals(result['Impact sur le changement climatique (kgCO2e)'])
                
                if prod not in material_totals:
                    material_totals.update({prod: impact})
                else:
                    material_totals[prod] += impact
                
                if mcrcom not in material_par_sp:
                    material_par_sp[mcrcom] = {prod: impact}
                else:
                    material_par_sp[mcrcom][prod] = impact
            
            fig = make_subplots(rows=2, cols=2, subplot_titles=('Sous Projet Totals', 'Matériaux Totals', 'Matériel par Sous Projet', 'Familles totals'))
            
            # Create bar charts for Sous Projet Totals
            fig.add_trace(go.Bar(x=list(sous_projet_totals.keys()), y=list(sous_projet_totals.values()), name='Sous Projet Totals'), row=1, col=1)
            
            # Create bar charts for Material Totals
            fig.add_trace(go.Bar(x=list(material_totals.keys()), y=list(material_totals.values()), name='Matériaux Totals'), row=1, col=2)
            
            # Create bar charts for Material by Sous Projet
            for sp, material_data in material_par_sp.items():
                fig.add_trace(go.Bar(x=list(material_data.keys()), y=list(material_data.values()), name=f'{sp}'), row=2, col=1)
            
            # Create bar charts for Material by Category
            fig.add_trace(go.Bar(x=list(category_quantity_dict.keys()), y=list(category_quantity_dict.values()), name='Familles totals'), row=2, col=2)
            
            # Update x-axis labels for all subplots
            fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=1, col=1)
            fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=1, col=2)
            fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=2, col=1)
            fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=2, col=2)
            
            moyen_total = total.replace(',', '')
            Moyen_poids_carbone_par_element = '{:,.0f}'.format(float(moyen_total) / len(element_quantity_dict))

            # Update layout
            fig.update_layout(barmode='group', title_text=f"Resultats poids carbone(kgCO²e): {total}\n    Selection: {user_answer} \n    Nombre d'élements calculés: {'{:,.0f}'.format(len(element_quantity_dict))}    Moyen poids carbone par élément(kgCO²e): {Moyen_poids_carbone_par_element}")
            # Save the plot to an HTML file using the open function with 'utf-8' encoding
            current_datetime = datetime.datetime.now().strftime("%y%m%d %Hh%M")
            html_file = "carbon_data {}.html".format(current_datetime)

            with open(path + "\\" + html_file, 'w', encoding='utf-8') as file:
                file.write(fig.to_html())
            
            alert = TaskDialog.Show("Calcul réussi", "L'extraction et le processus de calcul ont réussi. Veuillez importer les calculs dans les éléments du projet ou consultez les résultats dans le fichier 'carbon_data' au format html enregistré dans le répertoire du projet.")
        else:
            TaskDialog.Show("Calcul abandonné","Calcul abandonné")            
    else:
        TaskDialog.Show("Calcul unréalisable", "Veuillez sélectionner un ou plusieurs éléments avec des correctes classes de materiaux pour effectuer un calcul d'une sélection.")
else:
    pass