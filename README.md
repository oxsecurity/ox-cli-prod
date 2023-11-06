# oxcli
The OX CLI is a .NET command line interface application that will be expanded over time for the purpose of satisfying customer use case POCs.

Object processes designed in VB.NET are easily replicated in other languages, although this example utilizes a Python CLI engine for engine call processing.

------

To use, download and unzip contents of OXcli.zip into local folder. Windows recommended at this time.

OXCLI commands utilizing OX GraphQL API require a simple Python script located in the \python folder.
Modify the .env file with your OX API KEY.

In order for Python to run, test using:  python python_examp.py help

Likely will require:

pip install -r requirements
pip install python-getenv
pip install jsonpickle
