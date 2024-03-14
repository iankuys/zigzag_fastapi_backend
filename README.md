# VueZigZagBE

Written with Python 3.12.0.

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
