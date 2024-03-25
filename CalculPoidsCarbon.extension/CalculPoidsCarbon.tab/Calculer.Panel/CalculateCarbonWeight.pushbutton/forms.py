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
import clr
import csv
import os
clr.AddReference("System.Windows.Forms")

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument.Document

# Get title of revit project document
doc_title = doc.Title

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

# Define variables for windows form
new_version = None
update_version = None
selected_version = None
local_calculation = None
req_params = None
final_phase_calcul = None

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

    return new_version, update_version, selected_version, local_calculation, final_phase_calcul