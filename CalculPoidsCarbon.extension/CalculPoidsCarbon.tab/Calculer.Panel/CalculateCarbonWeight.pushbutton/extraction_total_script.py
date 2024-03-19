#! python3
# -*- coding: utf-8 -*-
from outil import install_packages

install_packages()
import socket
import sys
import site
import os

path_env = os.environ.get("PATH")
path_list = path_env.split(os.pathsep)
path_appended = False
for path in path_list:
    if path.endswith("Python38\Lib\site-packages"):
        site.USER_SITE = path
        sys.path.insert(0, path)
    if not path_appended:
        path_list.insert(0, site.getusersitepackages())
        sys.path.insert(0, site.getusersitepackages())
        path_appended = True
site.USER_SITE = site.getusersitepackages()
site.addsitedir(site.getusersitepackages(), known_paths=None)
sys.path.extend(site.getusersitepackages())

# Import necessary modules
import plotly.graph_objs as go
from plotly.subplots import make_subplots
import urllib.request as ur
import json
import pickle
from collections import defaultdict, OrderedDict
import copy
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

clr.AddReference("System.Windows.Forms")
import System
from System.Drawing import Size
from System.Windows.Forms import (
    Form,
    FormStartPosition,
    CheckBox,
    ComboBox,
    Button,
    DialogResult,
    MessageBox,
    MessageBoxButtons,
    MessageBoxIcon,
)
import csv
import psycopg2
import sqlalchemy
from sqlalchemy import create_engine, Column, Integer, String, Date, ForeignKey, URL
from sqlalchemy.orm import Session, relationship, declarative_base
from itertools import zip_longest
import requests
from requests.exceptions import RequestException, HTTPError, ConnectionError, Timeout
import datetime
from datetime import date
import openpyxl


clr.AddReference("RevitApi")
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

# Create empty lists and define functions needed for code
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
    return volume_in_cubic_foot / 35.314667


def round_3_decimals(number):
    """
    rounds number to 3 decimals.

    Parameters
    ----------
    number : int

    Returns
    -------
    int
        number rounded to 3 decimal points.

    """
    return round(number, 3)


# Collect phases
phases = FilteredElementCollector(doc).OfClass(Phase).ToElements()
phase_names = ["Tous constructions nouveaux"]

# Collect phases for when user chooses which phase to viusalise at importation
for phas in phases:
    phase_names.append(phas.Name)

filtered_elements = []

# Filter out those whose family and created phase parameters are None
for element in FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements():
    family_param = element.LookupParameter("Famille")
    phase_param = element.LookupParameter("Phase de création")

    if family_param is not None and phase_param is not None:
        phase = element.get_Parameter(BuiltInParameter.PHASE_CREATED).AsValueString()

        if (
            "Exi" not in phase
            and "EXI" not in phase
            and element.CreatedPhaseId is not None
        ):
            filtered_elements.append(element)

has_material_volume = []

# Loop through filtered_elements to find those with material volumes
for filtered_element in filtered_elements:
    materials = filtered_element.GetMaterialIds(False)
    volume = filtered_element.LookupParameter("Volume")
    famille = filtered_element.LookupParameter("Famille").AsValueString()
    for mat in materials:
        if filtered_element.GetMaterialVolume(mat) >= 0.000000001:
            has_material_volume.append(filtered_element)
    # Retrieve materials of elements with tangible dependent elements such as Mullions and Panels.
    if len(materials) == 0:
        dep_ids = filtered_element.GetDependentElements(None)
        for dependent_id in dep_ids:
            dependent_id = doc.GetElement(dependent_id)
            materials = dependent_id.GetMaterialIds(False)
            for mat in materials:
                if dependent_id.GetMaterialVolume(mat) >= 0.000000001:
                    has_material_volume.append(dependent_id)

# Create a new list to discard hand rails and bars as can't calculate for now
has_material_volume = [
    a
    for a in has_material_volume
    if a is not None
    and a.Category is not None
    and a.Category.Name != "Garde-corps"
    and a.Category.Name != "Barreaux"
]

# Define list to keep contrat cadre elements
to_remove = []

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
for item in to_remove:
    has_material_volume.remove(item)

# Define list for elements with unknown material classes
elements_with_unknown_materials = []

mat_not_in_db = []
dico = []
elem_mats = []

# Loop through final filtered list to retrieve data and update dico
for elem in has_material_volume:
    materials = elem.GetMaterialIds(False)
    element_id = elem.Id
    lot = "N/A"
    for material in materials:
        elem_mat = doc.GetElement(material)
        volume = volume_conv(elem.GetMaterialVolume(material))
        sous_projet = elem.LookupParameter("Sous-projet").AsValueString()
        # If material class is not "Generic", "Miscellaneous", "Non attribuée", etc. then continue
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
                dico.append(
                    {
                        "sous-projet": sous_projet,
                        "lot": str(lot),
                        "volume": round_3_decimals(volume),
                        "unit": str(unit),
                        "component_id": str(elem_mats[-1]),
                        "element_id": str(element_id),
                        "category": elem.Category.Name,
                        "quantity": round_3_decimals(quantity),
                    }
                )
            except:
                pass

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
        worksheet.append([str(element.Id), element.Category.Name])

    # Save the workbook to a file
    workbook.save("materiau.x_classe.s_inconnu.s.xlsx")

cw_agg_dict = {}
compressed_dico = []


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

# Calculate percentage of individual quantities relative to total quantity for each item in compressed_dico
components = {}

for q in compressed_dico:
    sous_projet = q["sous-projet"]
    material_id = [
        key
        for key, (value, _) in material_mapping.items()
        if q["component_id"] == value
    ]
    material_id = "".join(material_id)

    if q["quantity"] != 0:
        percentage_quantities = [r / q["quantity"] for r in q["individual_quantities"]]
        q["individual_quantities"] = percentage_quantities

    if sous_projet not in components:
        components[sous_projet] = {}

    components[sous_projet][material_id] = {
        "unit": q["unit"],
        "quantity": q["quantity"],
        "product_id": q["component_id"],
    }
# To remove items with quantity 0 from compressed_dico
compressed_dico = [q for q in compressed_dico if q["quantity"] != 0]


# Define variable to represent Project_id parameter
project_parameter_name = "Project_Id"

# Define variables for windows form
project_id = None
new_version = None
update_version = None
selected_version = None
local_calculation = None
req_params = None
final_phase_calcul = None

# Define project_info and get parameters for when we're ready to assign param project_id received from calculette
project_info = FilteredElementCollector(doc).OfClass(ProjectInfo).FirstElement()
parameters = project_info.Parameters

# Get Project_id from revit param. If no param, first calcul and project_id remains None
for info in parameters:
    if info.Definition.Name == project_parameter_name:
        project_id = info.AsString()
    if info.Definition.Name == "Etat du projet":
        phase_nom = info.AsString()
for info in parameters:
    if info.Definition.Name == "_INF_PRJ_NOM":
        project_name = info.AsString()
        break
    if info.Definition.Name == "Nom du bâtiment":
        project_name = info.AsString()

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

# Create simpleform if Project_id param
if project_id:

    class SimpleForm(Form):
        def __init__(self):
            self.Text = "Choisissez une option"
            self.StartPosition = (
                FormStartPosition.CenterScreen
            )  # Center the form on the screen
            self.Size = Size(400, 375)

            self.local_calculation_checkbox = CheckBox()
            self.local_calculation_checkbox.Text = (
                "Faire un calcul local (pas envoyé au Calculette Carbon)"
            )
            self.local_calculation_checkbox.Location = System.Drawing.Point(20, 30)
            self.local_calculation_checkbox.Size = Size(350, 30)

            self.new_version_checkbox = CheckBox()
            self.new_version_checkbox.Text = "Créer une nouvelle version"
            self.new_version_checkbox.Location = System.Drawing.Point(20, 80)
            self.new_version_checkbox.Size = Size(350, 30)

            self.update_version_checkbox = CheckBox()
            self.update_version_checkbox.Text = "Mettre à jour une version"
            self.update_version_checkbox.Location = System.Drawing.Point(20, 130)
            self.update_version_checkbox.Size = Size(350, 30)

            self.version_dropdown = ComboBox()
            self.version_dropdown.Location = System.Drawing.Point(20, 180)
            self.version_dropdown.Size = Size(350, 30)

            self.final_phase_calcul_checkbox = CheckBox()
            self.final_phase_calcul_checkbox.Text = "Faire un calcul final de la phase (à faire UNIQUEMENT si la phase est terminée)"
            self.final_phase_calcul_checkbox.Location = System.Drawing.Point(20, 230)
            self.final_phase_calcul_checkbox.Size = Size(350, 30)

            self.validate_button = Button()
            self.validate_button.Text = "Valider"
            self.validate_button.Location = System.Drawing.Point(150, 270)
            self.validate_button.Size = Size(100, 30)

            self.validate_button.Click += self.validate_button_clicked

            self.Controls.Add(self.new_version_checkbox)
            self.Controls.Add(self.update_version_checkbox)
            self.Controls.Add(self.version_dropdown)
            self.Controls.Add(self.local_calculation_checkbox)
            self.Controls.Add(self.final_phase_calcul_checkbox)
            self.Controls.Add(self.validate_button)

        def validate_button_clicked(self, sender, args):
            # Check the number of selected options
            selected_options = [
                self.new_version_checkbox.Checked,
                self.update_version_checkbox.Checked,
                self.local_calculation_checkbox.Checked,
                self.final_phase_calcul_checkbox.Checked,
            ]

            # Ensure that only one option is selected
            if sum(selected_options) != 1:
                MessageBox.Show(
                    "Veuillez cocher une case.",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                )
            elif (
                self.update_version_checkbox.Checked
                and not self.version_dropdown.SelectedItem
            ):
                MessageBox.Show(
                    "Veuillez choisir une version",
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                )
            else:
                self.DialogResult = DialogResult.OK
                self.Close()

    def get_version_numbers():
        csv_filename = version_path + f"\\versions {doc_title}.csv"
        # Replace this with your actual code to read version numbers from a CSV file
        try:
            with open(csv_filename, newline="") as csvfile:
                reader = csv.reader(csvfile)
                # Skip the header row
                next(reader, None)
                # Read version numbers from the specified columns (A2, B2, C2, etc.)
                version_numbers = [row[0] for row in reader]
                return version_numbers
        except FileNotFoundError:
            return []

    def show_dialog():
        form = SimpleForm()

        # Populate dropdown with version numbers from a CSV file
        form.version_dropdown.Items.AddRange(get_version_numbers())

        result = form.ShowDialog()

        if result == DialogResult.OK:
            global new_version
            global update_version
            global selected_version
            global local_calculation
            global final_phase_calcul
            # Check the selected options
            new_version = form.new_version_checkbox.Checked
            update_version = form.update_version_checkbox.Checked
            selected_version = (
                form.version_dropdown.SelectedItem if update_version else None
            )
            local_calculation = form.local_calculation_checkbox.Checked
            final_phase_calcul = form.final_phase_calcul_checkbox.Checked


# Call the show_dialog() function to display the dialog but only if Project_id revit param exists
if project_id:
    show_dialog()

# Get version id chosen by user
if selected_version:
    version_id = selected_version.split(";")[1]

# Get username associated with revit version
application = doc.Application
username = application.Username


def send_request(req_params):
    try:
        url = "http://192.168.96.212:81/api/v2/get_component_impacts"
        r = requests.post(url, json=req_params)
        r.raise_for_status()  # Raise an HTTPError for bad responses (4xx or 5xx)

        # Try to decode JSON
        return r.json()

    except RequestException as e:
        print(f"Error sending request: {e}")
        return None  # Handle the error case appropriately in your code


# Common parameters
common_params = {"project_lifetime": 50, "method": "static", "components": components}

# Do the necessary condition depending on whether, project_id is None, new version, update a version, or local
if project_id is None:
    req_params = {**common_params, "user": username, "persist": True}
elif local_calculation:
    req_params = {**common_params, "persist": False}
elif project_id and new_version:
    req_params = {
        **common_params,
        "user": username,
        "persist": True,
        "project_id": project_id,
    }
elif project_id and update_version:
    req_params = {
        **common_params,
        "user": username,
        "persist": True,
        "project_id": project_id,
        "version_id": version_id,
    }
elif project_id and final_phase_calcul:
    req_params = {
        **common_params,
        "user": username,
        "persist": True,
        "project_id": project_id,
    }


if req_params:
    # Send request using the refactored function
    results = send_request(req_params)

    # Remove duplicates if they occur
    response_data = remove_duplicate_dicts(results["impacts"])

    # Dump results so that importation_script.py can complete the importation process
    with open(os.path.join(path, "import_calc2.pickle"), "wb") as g:
        pickle.dump(response_data, g, protocol=0)

    with open(os.path.join(path, "compressed_dico.pickle"), "wb") as h:
        pickle.dump(compressed_dico, h, protocol=0)

    element_ids = [element.Id.IntegerValue for element in has_material_volume]

    with open(os.path.join(path, "element_ids.json"), "w") as f:
        json.dump(element_ids, f)

    list_of_results = response_data

    # Collect project info params
    project_info = FilteredElementCollector(doc).OfClass(ProjectInfo).FirstElement()
    parameters = project_info.Parameters
    proj_info_param_names = []

    # Get parameter names
    for params in parameters:
        proj_info_param_names.append(params.Definition.Name)

    def create_project_parameter(doc):
        """
        Create a project parameter in a Revit document.

        Parameters:
            doc (Document): The Revit document in which the parameter will be created.

        Returns:
            str: A message indicating the result of the operation.
                  Returns "Parameter created successfully." if successful.
                  Returns an error message if an exception occurs or if the shared parameter file is not found.
        """

        # Start a transaction
        t = Transaction(doc, "Create Project Parameter")
        t.Start()

        # Get or create the shared parameter file
        app = doc.Application
        sharedParameterFile = app.OpenSharedParameterFile()

        # Define the parameter name
        param_name = "Project_Id"

        # Create an ExternalDefinitionCreationOptions object
        options = ExternalDefinitionCreationOptions(param_name, ParameterType.Text)

        # Check if the parameter already exists
        definition = None
        for d in sharedParameterFile.Groups.get_Item("EMC2B").Definitions:
            if d.Name == param_name:
                definition = d
                break

        # If the definition doesn't exist, create it
        if definition is None:
            group = sharedParameterFile.Groups.get_Item("EMC2B")
            definition = group.Definitions.Create(options)

        # Define the categories to which the parameter will be bound
        categories = doc.Settings.Categories
        target_category = categories.get_Item(BuiltInCategory.OST_ProjectInformation)

        # Create a CategorySet and insert the target category
        category_set = app.Create.NewCategorySet()
        category_set.Insert(target_category)

        # Bind the parameter
        instanceBinding = app.Create.NewInstanceBinding(category_set)
        doc.ParameterBindings.Insert(
            definition, instanceBinding, BuiltInParameterGroup.PG_DATA
        )

        # Commit the transaction
        t.Commit()
        return "Parameter created successfully."

    row_list = [["Version name", "Version ID"], ["Version 1", results["version_id"]]]

    def create_version_csv():
        """
        Creates a CSV file containing version information.

        The function opens a CSV file with the name 'versions ' + str(doc_title) + '.csv'
        and writes the data from the provided row_list. The file is created or overwritten
        if it already exists.

        Parameters:
        - None

        Returns:
        - None
        """
        file_path = version_path + "/versions " + str(doc_title) + ".csv"

        with open(file_path, "w", newline="") as file:
            writer = csv.writer(file, delimiter=";")
            writer.writerows(row_list)

    def add_version():
        """
        Adds a new version entry to the CSV file.

        The function reads the existing CSV file to determine the current number of versions,
        increments the version counter, and appends a new row with the updated version information.
        The new row is also written to the CSV file.

        Global Variables:
        - row_list: A list containing rows of data for the CSV file.

        Parameters:
        - None

        Returns:
        - None
        """
        global row_list  # Access the global variable

        file_path = version_path + "/versions " + str(doc_title) + ".csv"

        counter = 0

        for row in open(file_path):
            counter += 1

        if new_version:
            row_list.append(["Version " + str(counter), results["version_id"]])

            # Update the CSV file with the new row
            file_path = version_path + "/versions " + str(doc_title) + ".csv"

            with open(file_path, "a", newline="") as file:
                writer = csv.writer(file, delimiter=";")
                writer.writerow(["Version " + str(counter), results["version_id"]])

    # Check if project_id is None
    if project_id is None:
        create_version_csv()

    if new_version:
        add_version()

    # Assign the parameter Project_Id the project_id given by the calculette
    if not "Project_Id" in proj_info_param_names:
        create_project_parameter(doc)
        project_parameter_name = "Project_Id"
        # Start a transaction
        transaction = Transaction(doc, "Set Shared Parameter Value")
        transaction.Start()

        # Collect elements of a specific category (e.g., walls)
        collector = FilteredElementCollector(doc)
        project_info = (
            collector.OfCategory(BuiltInCategory.OST_ProjectInformation)
            .WhereElementIsNotElementType()
            .ToElements()
        )

        # Loop through elements and set the shared parameter value
        for info in project_info:
            # Get the shared parameter by name
            project_shared_param = info.LookupParameter(project_parameter_name)

            # Check if the parameter exists
            if project_shared_param:
                # Set the parameter value
                project_shared_param.Set(
                    results["project_id"]
                )  # Set the value to 10.0 for example

        # Commit the transaction
        transaction.Commit()

    # Loop through the calculette results to replace "é" with "e".
    for dict in list_of_results:
        updated_dict = {}
        for key, value in dict.items():
            new_key = key.replace("é", "e")
            updated_dict[new_key] = value
        list_of_results[list_of_results.index(dict)] = updated_dict

    # Loop through compressed_dico and list_of_results simultaneously to calculate the carbon weight per element using the percentages previously calculated
    for dico in compressed_dico:
        for result in list_of_results:
            if dico["quantity"] == result["Quantite"]:
                for i in range(len(dico["individual_quantities"])):
                    percentage = dico["individual_quantities"][i]
                    element_carbon_weight = (
                        percentage
                        * result["Impact sur le changement climatique (kgCO2e)"]
                    )
                    dico["individual_quantities"][i] = element_carbon_weight

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

    # Define variables for plotly code
    total = 0
    mcrcom = None
    prod = None
    impact = None
    sous_projet_totals = {}
    material_totals = {}
    material_par_sp = {}

    for result in list_of_results:
        total_value = float(result["Impact sur le changement climatique (kgCO2e)"])

        # Convert total to a string if it's an integer
        if isinstance(total, int):
            total = str(total)

        # Remove commas
        total = total.replace(",", "")

        # Format the result with commas
        total = "{:,.3f}".format(float(total) + total_value)
        mcrcom = result["Macro-composant de niveau 1"]
        prod = result["Produit"]
        impact = round_3_decimals(
            result["Impact sur le changement climatique (kgCO2e)"]
        )

        if mcrcom not in sous_projet_totals:
            sous_projet_totals.update(
                {
                    mcrcom: round_3_decimals(
                        result["Impact sur le changement climatique (kgCO2e)"]
                    )
                }
            )
        else:
            sous_projet_totals[mcrcom] += round_3_decimals(
                result["Impact sur le changement climatique (kgCO2e)"]
            )

        if prod not in material_totals:
            material_totals.update({prod: impact})
        else:
            material_totals[prod] += impact

        if mcrcom not in material_par_sp:
            material_par_sp[mcrcom] = {prod: impact}
        else:
            material_par_sp[mcrcom][prod] = impact

    fig = make_subplots(
        rows=2,
        cols=2,
        subplot_titles=(
            "Sous Projet Totals",
            "Matériaux Totals",
            "Matériel par Sous Projet",
            "Familles totals",
        ),
    )

    # Create bar charts for Sous Projet Totals
    fig.add_trace(
        go.Bar(
            x=list(sous_projet_totals.keys()),
            y=list(sous_projet_totals.values()),
            name="Sous Projet Totals",
        ),
        row=1,
        col=1,
    )

    # Create bar charts for Material Totals
    fig.add_trace(
        go.Bar(
            x=list(material_totals.keys()),
            y=list(material_totals.values()),
            name="Matériaux Totals",
        ),
        row=1,
        col=2,
    )

    # Create bar charts for Material by Sous Projet
    for sp, material_data in material_par_sp.items():
        fig.add_trace(
            go.Bar(
                x=list(material_data.keys()),
                y=list(material_data.values()),
                name=f"{sp}",
            ),
            row=2,
            col=1,
        )

    # Create bar charts for Material by Category
    fig.add_trace(
        go.Bar(
            x=list(category_quantity_dict.keys()),
            y=list(category_quantity_dict.values()),
            name="Familles totals",
        ),
        row=2,
        col=2,
    )

    # Update x-axis labels for all subplots
    fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=1, col=1)
    fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=1, col=2)
    fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=2, col=1)
    fig.update_yaxes(title_text="Impact climat[kgCO2e]", row=2, col=2)

    moyen_total = total.replace(",", "")
    Moyen_poids_carbone_par_element = "{:,.0f}".format(
        float(moyen_total) / len(element_quantity_dict)
    )

    # Update layout
    fig.update_layout(
        barmode="group",
        title_text=f"Resultats poids carbone(kgCO²e): {total}\n    Projet: {doc.Title}\n    Nombre d'élements calculés: {'{:,.0f}'.format(len(element_quantity_dict))}    Moyen poids carbone par élément(kgCO²e): {Moyen_poids_carbone_par_element}",
    )
    # Save the plot to an HTML file using the open function with 'utf-8' encoding
    current_datetime = datetime.datetime.now().strftime("%y%m%d %Hh%M")

    # Create file
    html_file = "carbon_data {}.html".format(current_datetime)

    # Write file at given directory recently created
    with open(os.path.join(path, html_file), "w", encoding="utf-8") as file:
        file.write(fig.to_html())
        
    lst = []
    
    for key, value in element_quantity_dict_copy.items():
        if value < 1:
            lst.append({value: key})
            del element_quantity_dict[key]

    material_class_data = []
    sous_projet_data = []
    family_data = []
    type_data = []
    instance_data = []

    for key, value in element_quantity_dict.items():
        element_idee = ElementId(int(key))
        element = doc.GetElement(element_idee)
        elements_calculated_for_phase = len(element_quantity_dict)
        materials = element.GetMaterialIds(False)

        try:
            sous_projet_name = element.LookupParameter("Sous-projet").AsValueString()
            if all(sous_projet_name not in d.keys() for d in sous_projet_data):
                sous_projet_data.append(
                    {
                        sous_projet_name: {
                            "carbon_weight": 0,
                            "instance_calculated": 0,
                            "material_class_calculated": [],
                            "family_calculated": [],
                            "type_calculated": [],
                        }
                    }
                )
            for sous_projet in sous_projet_data:
                if sous_projet_name in sous_projet.keys():
                    # Update "carbon_weight" for the matching sous_projet_name
                    sous_projet[sous_projet_name]["carbon_weight"] += value
                    sous_projet[sous_projet_name]["instance_calculated"] += 1
                    if (
                        mat_class
                        not in sous_projet[sous_projet_name][
                            "material_class_calculated"
                        ]
                    ):
                        sous_projet[sous_projet_name][
                            "material_class_calculated"
                        ].append(mat_class)
                    if (
                        type_name
                        not in sous_projet[sous_projet_name]["type_calculated"]
                    ):
                        sous_projet[sous_projet_name]["type_calculated"].append(
                            type_name
                        )
                    if (
                        family_name
                        not in sous_projet[sous_projet_name]["family_calculated"]
                    ):
                        sous_projet[sous_projet_name]["family_calculated"].append(
                            family_name
                        )

                    # Update material_classes_calculated with the count of material classes
                    sous_projet[sous_projet_name]["material_class_count"] = len(
                        set(sous_projet[sous_projet_name]["material_class_calculated"])
                    )
                    sous_projet[sous_projet_name]["type_count"] = len(
                        set(sous_projet[sous_projet_name]["type_calculated"])
                    )
                    sous_projet[sous_projet_name]["family_count"] = len(
                        set(sous_projet[sous_projet_name]["family_calculated"])
                    )

        except:
            pass

        try:
            family_name = element.Category.Name
            if all(family_name not in d.keys() for d in family_data):
                family_data.append(
                    {
                        family_name: {
                            "carbon_weight": 0,
                            "instance_calculated": 0,
                            "material_class_calculated": [],
                            "type_calculated": [],
                        }
                    }
                )
            for element_category in family_data:
                if family_name in element_category.keys():
                    element_category[family_name]["carbon_weight"] += value
                    element_category[family_name]["instance_calculated"] += 1
                    if (
                        mat_class
                        not in element_category[family_name][
                            "material_class_calculated"
                        ]
                    ):
                        element_category[family_name][
                            "material_class_calculated"
                        ].append(mat_class)
                    if (
                        type_name
                        not in element_category[family_name]["type_calculated"]
                    ):
                        element_category[family_name]["type_calculated"].append(
                            type_name
                        )

                    # Update material_classes_calculated with the count of material classes
                    element_category[family_name]["material_class_count"] = len(
                        set(element_category[family_name]["material_class_calculated"])
                    )
                    element_category[family_name]["type_count"] = len(
                        set(element_category[family_name]["type_calculated"])
                    )
        except:
            pass

        try:
            type_name = element.LookupParameter("Type").AsValueString()
            if all(type_name not in d.keys() for d in type_data):
                type_data.append(
                    {
                        type_name: {
                            "carbon_weight": 0,
                            "instance_calculated": 0,
                            "material_class_calculated": [],
                            "family": family_name,
                        }
                    }
                )
            for type in type_data:
                if type_name in type.keys():
                    # Update "carbon_weight" for the matching type_name
                    type[type_name]["carbon_weight"] += value
                    type[type_name]["instance_calculated"] += 1
                    if mat_class not in type[type_name]["material_class_calculated"]:
                        type[type_name]["material_class_calculated"].append(mat_class)

                    # Update material_classes_calculated with the count of material classes
                    type[type_name]["material_class_count"] = len(
                        set(type[type_name]["material_class_calculated"])
                    )
        except:
            pass
        multiple_materials = ""
        try:
            type_name = element.LookupParameter("Type").AsValueString()
            family_name = element.Category.Name
            sous_projet_name = element.LookupParameter("Sous-projet").AsValueString()
            mat_class = elem_mat.MaterialClass
            # if len(materials) > 1:
            #     for m in materials:
            #         elem_m = doc.GetElement(m)
            #         m_class = elem_m.MaterialClass
            #         multiple_materials += str(m_class + ", ")
            if all(key not in d.keys() for d in instance_data) and len(materials) == 1:
                instance_data.append(
                    {
                        key: {
                            "carbon_weight": value,
                            "family": family_name,
                            "sous_projet": sous_projet_name,
                            "type": type_name,
                            "material_class": mat_class,
                        }
                    }
                )
            else:
                instance_data.append(
                    {
                        key: {
                            "carbon_weight": value,
                            "family": family_name,
                            "sous_projet": sous_projet_name,
                            "type": type_name,
                            "material_class": "multiple",
                        }
                    }
                )
        except Exception as e:
            print(e)
        for material in materials:
            elem_mat = doc.GetElement(material)
            mat_class = elem_mat.MaterialClass
            if (
                all(mat_class not in d.keys() for d in material_class_data)
                and mat_class not in materiaux_inconnus
            ):
                material_class_data.append(
                    {
                        mat_class: {
                            "carbon_weight": 0,
                            "instance_calculated": 0,
                        }
                    }
                )
            for mat in material_class_data:
                if mat_class in mat.keys():
                    # Update "carbon_weight" for the matching type_name
                    mat[mat_class]["carbon_weight"] += value
                    mat[mat_class]["instance_calculated"] += 1
        material_classes_for_phase = len(material_class_data)

    family_data_for_phase = len(family_data)
    type_data_for_phase = len(type_data)
    project_data = project_name
    phase_data = (
        phase_nom,
        total.replace(",", ""),
        date.today(),
        elements_calculated_for_phase,
        material_classes_for_phase,
        family_data_for_phase,
        type_data_for_phase,
    )

    ip_address = socket.gethostbyname("PO-ARP-6155")

    if final_phase_calcul:
        Base = declarative_base()

        class Project(Base):
            __tablename__ = "project"

            project_id = Column(Integer, primary_key=True)
            project_name = Column(String)

        class Phase(Base):
            __tablename__ = "phase"

            phase_id = Column(Integer, primary_key=True)
            phase_name = Column(String)
            carbon_weight_kg = Column(Integer)
            date_calculated = Column(Date)
            project_id = Column(Integer, ForeignKey("project.project_id"))
            instance_calculated = Column(Integer)
            material_class_calculated = Column(Integer)
            family_calculated = Column(Integer)
            type_calculated = Column(Integer)

            # Define a relationship to the Project class
            project = relationship("Project", back_populates="phase")

        Project.phase = relationship(
            "Phase", order_by=Phase.phase_id, back_populates="project"
        )

        class SousProjet(Base):
            __tablename__ = "sous_projet"

            sous_projet_id = Column(Integer, primary_key=True)
            sous_projet_name = Column(String)
            carbon_weight_kg = Column(Integer)
            project_id = Column(Integer, ForeignKey("project.project_id"))
            phase_id = Column(Integer, ForeignKey("phase.phase_id"))
            instance_calculated = Column(Integer)
            material_class_calculated = Column(Integer)
            family_calculated = Column(Integer)
            type_calculated = Column(Integer)

            # Define relationships to the Project and Phase classes
            project = relationship("Project", back_populates="sous_projet")
            phase = relationship("Phase", back_populates="sous_projet")

        # Add the sous_projets relationship to the Project and Phase classes
        Project.sous_projet = relationship(
            "SousProjet", order_by=SousProjet.sous_projet_id, back_populates="project"
        )
        Phase.sous_projet = relationship(
            "SousProjet", order_by=SousProjet.sous_projet_id, back_populates="phase"
        )

        class Family(Base):
            __tablename__ = "family"

            family_id = Column(Integer, primary_key=True)
            family_name = Column(String)
            carbon_weight_kg = Column(Integer)
            project_id = Column(Integer, ForeignKey("project.project_id"))
            phase_id = Column(Integer, ForeignKey("phase.phase_id"))
            instance_calculated = Column(Integer)
            material_class_calculated = Column(Integer)
            type_calculated = Column(Integer)

            # Define relationships to the Project, Phase, and Sous_projet classes
            project = relationship("Project", back_populates="family")
            phase = relationship("Phase", back_populates="family")

        # Add the material_classes relationship to the Project, Phase, and Sous_projet classes
        Project.family = relationship(
            "Family", order_by=Family.family_id, back_populates="project"
        )
        Phase.family = relationship(
            "Family", order_by=Family.family_id, back_populates="phase"
        )

        class Type(Base):
            __tablename__ = "type"

            type_id = Column(Integer, primary_key=True)
            type_name = Column(String)
            carbon_weight_kg = Column(Integer)
            project_id = Column(Integer, ForeignKey("project.project_id"))
            phase_id = Column(Integer, ForeignKey("phase.phase_id"))
            family = Column(String)
            instance_calculated = Column(Integer)
            material_class_calculated = Column(Integer)

            # Define relationships to the Project, Phase, and Sous_projet classes
            project = relationship("Project", back_populates="type")
            phase = relationship("Phase", back_populates="type")

        # Add the material_classes relationship to the Project, Phase, and Sous_projet classes
        Project.type = relationship(
            "Type", order_by=Type.type_id, back_populates="project"
        )
        Phase.type = relationship("Type", order_by=Type.type_id, back_populates="phase")

        class Instance(Base):
            __tablename__ = "instance"

            instance_id = Column(Integer, primary_key=True)
            carbon_weight_kg = Column(Integer)
            project_id = Column(Integer, ForeignKey("project.project_id"))
            phase_id = Column(Integer, ForeignKey("phase.phase_id"))
            material_class = Column(String)
            sous_projet = Column(String)
            family = Column(String)
            type = Column(String)

            # Define relationships to the Project, Phase, and Sous_projet classes
            project = relationship("Project", back_populates="instance")
            phase = relationship("Phase", back_populates="instance")

        # Add the material_classes relationship to the Project, Phase, and Sous_projet classes
        Project.instance = relationship(
            "Instance", order_by=Instance.instance_id, back_populates="project"
        )
        Phase.instance = relationship(
            "Instance", order_by=Instance.instance_id, back_populates="phase"
        )

        class MaterialClass(Base):
            __tablename__ = "material_class"

            material_class_id = Column(Integer, primary_key=True)
            material_class_name = Column(String)
            carbon_weight_kg = Column(Integer)
            instance_calculated = Column(Integer)
            project_id = Column(Integer, ForeignKey("project.project_id"))
            phase_id = Column(Integer, ForeignKey("phase.phase_id"))

            # Define relationships to the Project, Phase, and Sous_projet classes
            project = relationship("Project", back_populates="material_class")
            phase = relationship("Phase", back_populates="material_class")

        # Add the material_classes relationship to the Project, Phase, and Sous_projet classes
        Project.material_class = relationship(
            "MaterialClass",
            order_by=MaterialClass.material_class_id,
            back_populates="project",
        )
        Phase.material_class = relationship(
            "MaterialClass",
            order_by=MaterialClass.material_class_id,
            back_populates="phase",
        )

        def insert_data_into_database(
            project_data,
            phase_data,
            sous_projet_data,
            family_data,
            type_data,
            instance_data,
            material_class_data,
        ):

            # Set your database connection details
            url_object = URL.create(
                "postgresql+psycopg2",
                username="postgres",
                password="postgres",  # plain (unescaped) text
                host=ip_address,
                database="carbonbim",
            )
            engine = create_engine(url_object)

            # Create a session to interact with the database
            session = Session(engine)

            # try:
            # Insert data into the 'project' table
            project = Project(
                project_name=project_data,
            )
            session.add(project)

            # Insert data into the 'phases' table
            phase = Phase(
                phase_name=phase_data[0],
                carbon_weight_kg=phase_data[1],
                date_calculated=date.today(),
                instance_calculated=phase_data[3],
                material_class_calculated=phase_data[4],
                family_calculated=phase_data[5],
                type_calculated=phase_data[6],
            )
            project.phase.append(phase)

            session.add(phase)

            # Iterate through sous_projet_data, material_class_data, and family_data, etc. simultaneously
            for (
                instance_entry,
                sous_projet_entry,
                family_entry,
                type_entry,
                material_class_entry,
            ) in zip_longest(
                instance_data,
                sous_projet_data,
                family_data,
                type_data,
                material_class_data,
                fillvalue=None,
            ):
                # Use the unique identifier as sous_projet_name, material_class_name, and element_name
                if instance_entry is not None:
                    key = list(instance_entry.keys())[0]
                if sous_projet_entry is not None:
                    sous_projet_name = list(sous_projet_entry.keys())[0]
                if family_entry is not None:
                    family_name = list(family_entry.keys())[0]
                if type_entry is not None:
                    type_name = list(type_entry.keys())[0]
                if material_class_entry is not None:
                    material_class_name = list(material_class_entry.keys())[0]

                # Insert data into the 'sous_projet' table
                if sous_projet_entry is not None:
                    sous_projet = SousProjet(
                        sous_projet_name=sous_projet_name,
                        carbon_weight_kg=sous_projet_entry[sous_projet_name][
                            "carbon_weight"
                        ],
                        instance_calculated=sous_projet_entry[sous_projet_name][
                            "instance_calculated"
                        ],
                        material_class_calculated=sous_projet_entry[sous_projet_name][
                            "material_class_count"
                        ],
                        family_calculated=sous_projet_entry[sous_projet_name][
                            "family_count"
                        ],
                        type_calculated=sous_projet_entry[sous_projet_name][
                            "type_count"
                        ],
                        phase=phase,
                        project=project,
                    )
                    session.add(sous_projet)

                # Insert data into the 'family' table
                if family_entry is not None:
                    family = Family(
                        family_name=family_name,
                        carbon_weight_kg=family_entry[family_name]["carbon_weight"],
                        instance_calculated=family_entry[family_name][
                            "instance_calculated"
                        ],
                        material_class_calculated=family_entry[family_name][
                            "material_class_count"
                        ],
                        type_calculated=family_entry[family_name]["type_count"],
                        phase=phase,
                        project=project,
                    )
                    session.add(family)

                # Insert data into the 'type' table
                if type_entry is not None:
                    type = Type(
                        type_name=type_name,
                        carbon_weight_kg=type_entry[type_name]["carbon_weight"],
                        instance_calculated=type_entry[type_name][
                            "instance_calculated"
                        ],
                        material_class_calculated=type_entry[type_name][
                            "material_class_count"
                        ],
                        family=type_entry[type_name]["family"],
                        phase=phase,
                        project=project,
                    )
                    session.add(type)

                # Insert data into the 'instance' table
                if instance_entry is not None:
                    instance = Instance(
                        carbon_weight_kg=instance_entry[key]["carbon_weight"],
                        family=instance_entry[key]["family"],
                        sous_projet=instance_entry[key]["sous_projet"],
                        type=instance_entry[key]["type"],
                        material_class=instance_entry[key]["material_class"],
                        phase=phase,
                        project=project,
                    )
                session.add(instance)
                # Insert data into the 'material_class' table
                if material_class_entry is not None:
                    material_class = MaterialClass(
                        material_class_name=material_class_name,
                        carbon_weight_kg=material_class_entry[material_class_name][
                            "carbon_weight"
                        ],
                        instance_calculated=material_class_entry[material_class_name][
                            "instance_calculated"
                        ],
                        phase=phase,
                        project=project,
                    )
                    session.add(material_class)

            # Commit the transaction
            session.commit()

            print("Data inserted successfully.")

            # except Exception as e:
            #     print(f"Error: {str(e)}")

            # finally:
            #     # Close the session
            session.close()

        insert_data_into_database(
            project_data,
            phase_data,
            sous_projet_data,
            family_data,
            type_data,
            instance_data,
            material_class_data,
        )
    alert = TaskDialog.Show(
        "Calcul réussi",
        "L'extraction et le processus de calcul ont réussi. Veuillez importer les calculs dans les éléments du projet ou consultez les résultats dans le fichier 'carbon_data' au format html enregistré dans un dossier appelé 'carbon_data' sur votre bureau.",
    )
else:
    TaskDialog.Show("Calcul abandonné", "Calcul abandonné")
