# VueZigZagBE

This tool automates the generation of **zig-zag filled PowerPoint presentations** for use in **UCI MIND** research. It is designed to streamline patient record creation and standardize how data is presented across studies.

The system combines a **Visual Basic Script (VBS)** for advanced PowerPoint manipulation with a **Python FastAPI** backend to manage patient data, trigger generation tasks, and provide an interface for integration or automation.

## Features

- Automatically generates zig-zag visual layouts in PowerPoint
- Accepts patient data via REST API
- Triggers and controls VBS execution from the FastAPI backend
- Maintains logs/records of generated presentations
- Customizable templates for consistent research formatting

## Use Cases

- Streamlined presentation generation for cognitive or memory trials
- Automatic documentation of patient sessions or study outcomes
- Standardized slide creation for multi-participant research studies
## connection_string.txt set up
## Dev environment set up

1. Create a virtual environment for developing this app (example with PowerShell on Windows 11):
```
# Navigate to this repository's root folder

# Create a Python virtual environment (only need to do this once):
python -m venv .venv

# Activate the virtual environment:
.\.venv\Scripts\Activate.ps1
# On Windows: .ps1 for PowerShell or .bat for Command Prompt

# If using PowerShell and "running scripts is disabled on this system", need to
# enable running external scripts. Open PowerShell as admin and use this command:
set-executionpolicy remotesigned
# (only need to do this once)

# While in the virtual env, update pip and install packages (only need to do this once):
python -m pip install --upgrade pip
pip install -r requirements.txt

# Run the script: develop, debug, etc.
python main.py

# Press CTRL+C to close the development server

# Deactivate when done
deactivate
```

## connection_string.txt set up
In development, if your UCI MIND account has access to the database, you can use this string instead.
```
"DRIVER={SQL Server};Provider=SQLOLEDB;Server=____;Database=____;Integrated Security=SSPI;DataTypeCompatibility=80;MARS Connection=True;"
```
In production, please replace the underscores with the correct information of the database server and user.
```
"DRIVER={SQL Server};Provider=SQLOLEDB;Server=____.uci.edu;Database=____;User Id=____;Password=____;DataTypeCompatibility=80;MARS Connection=True;"
```

*For research use at UCI MIND only.*
