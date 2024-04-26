# oxcli
The OX CLI is a .NET command line interface application that will be expanded over time for the purpose of satisfying customer use case POCs.

Object processes designed in VB.NET are easily replicated in other languages, although this example utilizes a Python CLI engine for engine call processing.

To use, download and unzip contents of OXcli.zip into local folder. Tested on both Windows and Mac.

Preparations require DOTNET6 RUNTIME and PYTHON in the environment and accessible via %PATH%.

OXCLI commands utilizing OX GraphQL API require a simple Python script.

1) Download the *OXcli.zip file from GitHub
2) Unzip the ZIP contents into a folder that you will run the application from
3) Ensure both PYTHON (or PYTHON3/PIP3 if MAC) and DOTNET are in the path
       Open Terminal/Shell and type "python3" & "dotnet"
       If "file not found" you will need to insert into PATH
       Install DOTNET 6.025 RUNTIMETo put DOTNET into path use command: _ln -s /usr/local/share/dotnet/dotnet /usr/local/bin/_ (may need SUDO)
4) From terminal/shell run dotnet oxcli.dll checkme
5) Assuming good, run dotnet _oxcli.dll setenv --key ox_api_key_ to set OX environment
6) Prepare Python
       python -m pip install -r requirements.txt    (allows Python to install all dependencies defined in file)
       Try it out: _python getApps_ should produce file getapps_response.json

OX CLI should now be ready to go
