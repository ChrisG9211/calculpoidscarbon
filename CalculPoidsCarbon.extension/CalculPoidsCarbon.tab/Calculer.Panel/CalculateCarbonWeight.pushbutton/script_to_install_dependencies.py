import subprocess

# List of required packages
required_packages = [
    "plotly==5.17.0",
    "pypac==0.16.4",
    "openpyxl==3.1.2"
]

# Form the pip install command
pip_command = ["pip", "install"] + required_packages

# Execute the command
subprocess.run(pip_command, check=True)
