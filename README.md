# Case-Management-script
A simple python script to generate nested folders and copy template files into the structure based on user input.

The script creates a directory structure based on user input where parent case directory will contain one or more item numbers under that case. The script will also generate an "exports" folder within each item number for derivative evidence as well as a "Reports" subdirectory. The script will generate a simple worklog txt file in the parent folder which is modified to contain headers for each declared item number. The script then copies a report template from a declared location in the code into the Reports folder. The report template will be modified according to user input values. This requires the docxtpl library. If you would like to modify the template, read up on the docxtpl usage to avoid breaking the syntax:

https://docxtpl.readthedocs.io/en/latest/

This version leverages TKinter for the GUI and docxtpl for template interaction. The GUI version also leverages a small json file to store "default" values to be inserted into each report template that is produced. The user can house their template wherever they like and choose an output directory vs. the hard coded initial CLI version. Based on the item numbers entered, the script will iterate item number folders, headings in the .txt worklog file, and iterate item number tables within the template. 

The CLI version has not been updated yet, the functionality is not the same between the two versions right now.
