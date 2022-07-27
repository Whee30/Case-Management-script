# Case-Management-script
A simple python script to generate nested folders and copy template files into the structure based on user input.

Currently, the script is set up based on a consistent naming convention of 9 character case numbers with the option to add additional flavor text. 

The script will create a parent directory at a location specified in the initialPath variable. Within this directory will be a named folder for each evidence item as entered by the user (strings separated by spaces). A reports folder will be created within the parent directory and a template report document will be copied from a set location with the 9 character case number inserted into the filename. Finally, a worklog text file will be generated in the parent directory named for the case number. The content of the worklog will contain the case number and new lines for each item number entered.

The newly added GUI version is tweaked just a bit. The parent directory is now named "case number, assigned detective". The file paths were simplified a bit from the initial version as well.
