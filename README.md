# Case-Management-script
A simple python script to generate nested folders and copy template files into the structure based on user input.

The script operates under the structure of a unique case number which contains one or more item numbers under that case. The script will create a parent directory for the case number, which contains subdirectories for each item number as well as a "Reports" subdirectory. The script will generate a simple worklog txt file in the parent folder which is modified to contain headers for each declared item number. The script then copies a report template from a declared location in the code into the Reports folder. The report template will aslo iterate a table for each item entered. This requires the docxtpl library.

Two versions of the script exist in this project, a simple command line script as well as a tkinter GUI version. The versions operate slightly differently but the primary concepts are the same.

EDIT March 9th, 2025:

A new version of the GUI script was added. This version leverages TKinter for the GUI and docxtpl for template interaction. The GUI version also leverages a small json file to store "default" values to be inserted into each report template that is produced. The user can house their template wherever they like and choose an output directory vs. the hard coded initial CLI version. Based on the item numbers entered, the script will iterate item number folders, headings in the .txt worklog file, and iterate item number tables within the template. 

Some more variables need to be established before this is "done", but it works in its current state. Error handling and edge cases should be addressed, duplicate case numbers have not yet been addressed.

The CLI version has not been updated yet, the functionality is not the same between the two versions right now.
