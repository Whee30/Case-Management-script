# Python script to create folder structures upon case creation
import os 
import json
from shutil import copyfile
from docxtpl import DocxTemplate
import argparse
import sys





label = {}
entry = {}
content = {}
stored_data = {}

if os.path.exists('stored_values.json') is  False:
    default_values = {
        'name':'Ofc. J. Smith #123',
        'agency':'anytown police department',
        'template':'my_template.docx',
        'destination':'D:/'
    }
    with open('stored_values.json','w') as file:
        json.dump(default_values,file,indent=4)
else:
    with open('stored_values.json', 'r') as file:
        default_values = json.load(file)

def apply_changes():
    print("Applying changes")

parser = argparse.ArgumentParser(description="Case Creation Utility")

# Add arguments (flags)
parser.add_argument('-d', '--defaults', action='store_true', help="Display current defaults.")
parser.add_argument('-n', '--name', action='store_true', help="Change default name.")
parser.add_argument('-a', '--agency', action='store_true', help="Change default agency.")
parser.add_argument('-i', '--input', action='store_true', help="Choose a template.")
parser.add_argument('-o', '--output', action='store_true', help="Choose an output directory.")
parser.add_argument('-c', '--changeall', action='store_true', help="Change all default values in succession.")

#parser.add_argument('-h', '--help', action='help', help="Show this help message and exit.")

# Parse the arguments
args = parser.parse_args()



if len(sys.argv) > 1:
    if args.defaults:
        print(f"The default name is: {default_values['name']}")
        print(f"The default agency is: {default_values['agency']}")
        print(f"The default template path is: {default_values['template']}")
        print(f"The default output path is: {default_values['destination']}")
    
    if args.changeall:
        default_values['name'] = input("Set default name:")
        default_values['agency'] = input("Set default agency:")
        default_values['template'] = input("Set .docx template:")
        default_values['output'] = input("Set output path:")
        apply_changes
    if args.name:
        default_values['name'] = input("Set default name:")
    if args.agency:
        default_values['agency'] = input("Set default agency:")
    if args.input:
        default_values['template'] = input("Set .docx template:")
    if args.output:
        default_values['output'] = input("Set output path:")
else:
    '''


    #####################
    ##### Functions #####
    #####################

    # change/save default values
    def change_default_data():
        for key in stored_data:
            stored_data[key] = entry[key].get()
        with open('stored_values.json', 'w') as file:
            json.dump(stored_data, file, indent=4)

    # Choose output directory

    # choose template file  

    def generate_case():
        # Remove leading and trailing whitespace, remove tabs, remove empty lines
        # This preserves whitespace insid of a line, for items like "cellphone 1"
        item_content = entry['items'].get(1.0, END).splitlines()
        item_list = [item.strip().replace("\t","") for item in item_content if item.strip()]
        
        # get the nine digit case number from the beginning of the input
        casefile_path = os.path.join(stored_data['destination'], entry['case_num'].get().strip())
        truncated_case_number = entry['case_num'].get()[0:9]

        # Check if path already exists, if not, create case folders
        if os.path.exists(casefile_path):
            print("This case path already exists, please choose another case number/description.")
            return
        else:
            os.mkdir(casefile_path) 

        # Create a "reports" directory inside the parent case directory
        report_directory = os.path.join(casefile_path, "Reports")
        os.mkdir(report_directory)

        # create folders for each evidence item entered
        for i in item_list:
            os.mkdir(os.path.join(casefile_path, i))

        # create and edit the worklog txt file within the parent case directory
        worklog_name = str(truncated_case_number + " worklog.txt")
        worklog_path = os.path.join(casefile_path, worklog_name)
        worklog_file = open(worklog_path, "w+")

        worklog_file.write(entry['case_num'].get() + "\n\n")

        for i in item_list:
            worklog_file.write(i + ":\n\n")
        worklog_file.close()

                    # Actually build the .docx
        edit_template = DocxTemplate(entry['template'].get())

        for key, value in entry.items():
            if key == "items":
                content['items'] = entry['items'].get(1.0, END)
            else:
                content[key] = value.get()
        
        content['item_list'] = item_list

        edit_template.render(content, autoescape=True)
        output_filename = f"{entry['case_num'].get()} Examination Report.docx"
        output_path = os.path.join(report_directory, output_filename)
        edit_template.save(output_path)
    '''










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
    reportSource = "skeleton.docx"

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