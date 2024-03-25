#! python3
# -*- coding: utf-8 -*-
from outil import install_packages, set_path, round_3_decimals, remove_duplicate_dicts, create_project_parameter
from extraction import compressed_dico_extraction, create_element_id_dict

install_packages()
set_path()

# Import necessary modules
import os
import pickle
import clr
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    ProjectInfo,
    BuiltInCategory,
)
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
import requests
from requests.exceptions import RequestException

clr.AddReference("System.Windows.Forms")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument.Document
__location__ = os.path.realpath(
    os.path.join(os.getcwd(), os.path.dirname(__file__)))
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

components = {}
compressed_dico = compressed_dico_extraction()
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

    list_of_results = response_data

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

    # Collect project info params
    project_info = FilteredElementCollector(doc).OfClass(ProjectInfo).FirstElement()
    parameters = project_info.Parameters
    proj_info_param_names = []

    # Get parameter names
    for params in parameters:
        proj_info_param_names.append(params.Definition.Name)

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
    
    element_quantity_dict, category_quantity_dict = create_element_id_dict(compressed_dico)

    if final_phase_calcul:
        with open(os.path.join(__location__, 'insert_db.py'), 'r') as file:
            exec(file.read())
    with open(os.path.join(__location__, 'graphs.py'), 'r') as file:
        exec(file.read())
    alert = TaskDialog.Show(
        "Calcul réussi",
        "L'extraction et le processus de calcul ont réussi. Veuillez importer les calculs dans les éléments du projet ou consultez les résultats dans le fichier 'carbon_data' au format html enregistré dans un dossier appelé 'carbon_data' sur votre bureau.",
    )
else:
    TaskDialog.Show("Calcul abandonné", "Calcul abandonné")
