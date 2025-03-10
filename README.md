# Case-Management-script
A simple python script to generate nested folders and copy template files into the structure based on user input.

The script operates under the structure of a unique case number which contains one or more item numbers under that case. The script will create a parent directory for the case number, which contains subdirectories for each item number as well as a "Reports" subdirectory. The script will generate a simple worklog txt file in the parent folder which is modified to contain headers for each declared item number. The script then copies a report template from a declared location in the code into the Reports folder.

Two versions of the script exist in this project, a simple command line script as well as a tkinter GUI version. The versions operate slightly differently but the primary concepts are the same.

A request was made to include the report template into each item folder as well, the GUI version was updated to include this feature, which can be commented out if unwanted.

EDIT March 6th, 2025:
This script is currently being worked on to include some features from the warrant builder application as far as making changes internally to a template report file. Updates should be forthcoming soon.
