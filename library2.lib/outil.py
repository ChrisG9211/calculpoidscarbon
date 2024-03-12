#! python3
import subprocess

# List of required packages
required_packages = [
    "plotly==5.17.0",
    "openpyxl==3.1.2",
    "sqlalchemy==2.0.27",
    "psycopg2==2.9.9",
    "datetime==5.4"
]

# Form the pip install command
pip_command = ["python3.8.5", "-m" "pip3", "install"] + required_packages

# Execute the command
def install_packages():
    return subprocess.run(pip_command, check=True)

