# Python script to create folder structures upon case creation

# importing os module 
import os 

# Initial path where your parent level case files will be stored
initialPath = "F:/"

# User input asking for a case number and short description "V12345678 Fraud Case". This will be used elsewhere.
caseNumber = input("Enter the case number and description: ")
 
# Uses base path and adds on your case number as a directory
# Path F:/V12345678 descriptor/
filePath = os.path.join(initialPath, caseNumber) 

# In this example, my case numbers are 9 digits and then a descriptor.
# truncated path takes the first 9 digits for use later
truncatedCase = caseNumber[0:9]

# Create the main directory 
os.mkdir(filePath) 

# Declare where the template docx lives
reportSource = "C:/Users/<username>/.../template.docx"

# Create a "reports" directory inside the parent case directory
reportDest = os.path.join(filePath, "Reports")
os.mkdir(reportDest)

# Put a copy of the template document into the new reports folder
copiedReportFile = os.path.join(reportDest, truncatedCase)
from shutil import copyfile
copyfile(reportSource, copiedReportFile + " Examination Report.docx")

# take a second user input and create an array from the items entered
# example: "1A 2B 3C"
evidenceItems = input("Input evidence items separated by a space:")
arrayvars = evidenceItems.split()

# create folders for each evidence item entered
for i in arrayvars:
	os.mkdir(os.path.join(filePath, i))

# create and edit the worklog txt file within the parent case directory
worklogFileName = str(truncatedCase + " worklog.txt")
worklogFilePath = os.path.join(filePath, worklogFileName)
worklog_file = open(worklogFilePath, "w+")

# put the case number in the worklog txt file
worklog_file.write(caseNumber + "\n\n")

# make headers for the item numbers in the worklog txt file
for i in arrayvars:
    worklog_file.write(i + ":\n\n")
worklog_file.close()

# holds the window open to confirm all steps are complete. If the window auto closes, something is failing above.
input("Press ENTER to exit")