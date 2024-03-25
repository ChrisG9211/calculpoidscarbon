#! python3
import subprocess
import sys
import site
import os
import csv

from Autodesk.Revit.DB import (
    Transaction,
    ParameterType,
    BuiltInParameterGroup,
    BuiltInCategory,
    ExternalDefinitionCreationOptions,
)

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument.Document

# Get title of revit project document
doc_title = doc.Title

# List of required packages
required_packages = [
    "plotly==5.17.0",
    "openpyxl==3.1.2",
    "sqlalchemy==2.0.27",
    "psycopg2==2.9.9",
    "datetime==5.4"
]

# Form the pip install command
pip_command = ["pip", "install"] + required_packages

# Define variables for directories at given paths
directory = "carbon_data"
versions_directory = "versions"
new_dir_path = os.path.normpath(os.path.expanduser("~/Desktop"))
path = os.path.join(new_dir_path, directory)
version_path = os.path.join(new_dir_path, directory, versions_directory)

# Declare boolean variables to check if paths exist
path_exists = os.path.exists(path)
version_path_exists = os.path.exists(version_path)

# Execute the command
def install_packages():
    """
    Install packages using pip command.

    This function executes the pip command specified in the global variable `pip_command`
    and checks if the installation was successful.

    Returns:
        CompletedProcess: A subprocess.CompletedProcess instance representing the
        result of the pip command execution.
    """
    return subprocess.run(pip_command, check=True)

def set_path():
    """
    Set Python site-packages path.

    This function retrieves the current PATH environment variable and splits it into a list of paths.
    It then iterates through each path, looking for the directory containing Python packages for Python 3.8.
    When found, it sets the user-specific site-packages directory and inserts it into the sys.path list.
    Finally, it adds the directory to the known site-packages directories.
    """
    path_env = os.environ.get("PATH")
    path_list = path_env.split(os.pathsep)
    for path in path_list:
        if path.endswith("Python38\Lib\site-packages"):
            site.USER_SITE = path
            sys.path.insert(0, path)
            site.addsitedir(path, known_paths=None)
    
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



