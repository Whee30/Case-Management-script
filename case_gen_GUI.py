import os
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog, messagebox
from shutil import copyfile
import json
import subprocess
from docxtpl import DocxTemplate

# Create compatible json if not already present
if os.path.exists('stored_values.json') is  False:
    # establish initial default values for json
    default_values = {
        'name':'Ofc. J. Smith #123',
        'agency':'anytown police department',
        'template':'my_template.docx',
        'destination':'D:/'
    }

    # apply defaults to json
    with open('stored_values.json','w') as file:
        json.dump(default_values,file,indent=4)
    
    # Make the newly generated json a hidden file... comment out if this behavior is not desired.
    # Hidden files present permission problems. This feature not functional currently.
    # subprocess.check_call(["attrib","+H","stored_values.json"])
    
# create root window
root = Tk()
main_frame = Frame(root)
main_frame.pack(fill="both", expand=1)

style = Style()
# print(style.theme_names())
style.theme_use("clam")

label = {}
entry = {}
content = {}

# load json values for the "stored values" fields
with open('stored_values.json', 'r') as file:
    stored_data = json.load(file)

#####################
##### Functions #####
#####################

# reset stored values into stored data fields
def reset_stored_data():
    for item in stored_data:
        entry[item].delete(0, END)
        entry[item].insert(0, stored_data[item])

# change/save default values
def change_default_data():
    for key in stored_data:
        stored_data[key] = entry[key].get()
    with open('stored_values.json', 'w') as file:
        json.dump(stored_data, file, indent=4)

# Choose output directory
def output_select():
    selected_folder = filedialog.askdirectory()
    if selected_folder:
        entry['destination'].delete(0, END)
        entry['destination'].insert(0, selected_folder)

# choose template file  
def template_select():
    selected_file = filedialog.askopenfilename()
    if selected_file:
        entry['template'].delete(0, END)
        entry['template'].insert(0, selected_file)

def reset_form():
    if messagebox.askyesno("Warning", "Are you sure you want to reset the form?"):
        entry['case_num'].delete(0, END)
        entry['items'].delete(1.0, END)

def generate_case():
    # Remove leading and trailing whitespace, remove tabs, remove empty lines
    # This preserves whitespace insid of a line, for items like "cellphone 1"
    item_content = entry['items'].get(1.0, END).splitlines()
    item_list = [item.strip().replace("\t","") for item in item_content if item.strip()]
    
    # get the nine digit case number from the beginning of the input
    casefile_path = os.path.join(stored_data['destination'], entry['case_num'].get().strip())
    truncated_case_number = entry['case_num'].get()[0:9]

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

    root.destroy()



#####################
##### Begin GUI #####
#####################

root.title("Case Creation Utility v0.2")
root.geometry('800x400')

left_frame = Frame(main_frame)
left_frame.pack(side="left", fill="both", expand="1", padx=10, pady=10)

right_frame = Frame(main_frame, relief="groove")
right_frame.pack(side="top", fill="none", expand="0", ipadx=10, ipady=10, padx=10, pady=10)

r_button_frame = Frame(right_frame)
r_button_frame.grid(column=0, row=8, columnspan=2)

########################
##### LEFT WIDGETS #####
########################

label['new_info'] = Label(left_frame, text="New Case Information:", font=("TkDefaultFont", 14))
label['case'] = Label(left_frame, text="Case Number/Description:")
entry['case_num'] = Entry(left_frame, width=25, font=("TkDefaultFont", 12))
label['items'] = Label(left_frame, text="List Items one per line:")
entry['items'] = Text(left_frame, width=25, height=10, font=("TkDefaultFont", 12))
spacer = Label(left_frame, text="")
submit_form = Button(left_frame, text="Create Case", command=generate_case)
clear_form = Button(left_frame, text="Clear Form", command=reset_form)

#####################
##### LEFT GRID #####
#####################

label['new_info'].grid(column=0, row=0, columnspan=2, sticky="nw", pady=5)
label['case'].grid(column=0, row=2, sticky="nw", padx=10, pady=5)
entry['case_num'].grid(column=1, row=2, columnspan=2, sticky="nw", pady=5)
label['items'].grid(column=0, row=3, sticky="nw", padx=10, pady=5)
entry['items'].grid(column=1, row=3, columnspan=2, sticky="nw", pady=5)
submit_form.grid(column=1, row=4, sticky="ne", pady=5, padx=5)
clear_form.grid(column=2, row=4, sticky="nw", pady=5, padx=5)

#########################
##### RIGHT WIDGETS #####
#########################

label['stored_info'] = Label(right_frame, text="Stored Defaults:", font=("TkDefaultFont", 14))
label['name'] = Label(right_frame, text="Name:")
entry['name'] = Entry(right_frame, textvariable=stored_data['name'], width=40)
label['agency'] = Label(right_frame, text="Agency:")
entry['agency'] = Entry(right_frame, textvariable=stored_data['agency'], width=40)
label['destination'] = Label(right_frame, text="Output Dir:")
entry['destination'] = Entry(right_frame, textvariable=stored_data['destination'], width=40)
output_browse = Button(right_frame, text="browse", command=output_select)
label['template'] = Label(right_frame, text="Template Path:")
entry['template'] = Entry(right_frame, textvariable=stored_data['template'], width=40)
template_browse = Button(right_frame, text="browse", command=template_select)
r_spacer = Label(right_frame, text="")
change_defaults = Button(r_button_frame, text="Submit Defaults", command=change_default_data)
reset_defaults = Button(r_button_frame, text="Reset Defaults", command=reset_stored_data)

######################
##### RIGHT GRID #####
######################

label['stored_info'].grid(column=0, row=0, pady=5, padx=5, columnspan=2, sticky="nw")
label['name'].grid(column=0, row=1, sticky="nw", padx=10, pady=5)
entry['name'].grid(column=1, row=1, sticky="nw", pady=5)
label['agency'].grid(column=0, row=2, sticky="nw", padx=10, pady=5)
entry['agency'].grid(column=1, row=2, sticky="nw", pady=5)
label['destination'].grid(column=0, row=3, sticky="nw", padx=10, pady=5)
entry['destination'].grid(column=1, row=3, sticky="e", pady=5)
output_browse.grid(column=1, row=4, sticky="e", pady=5)
label['template'].grid(column=0, row=5, sticky="nw", padx=10, pady=5)
entry['template'].grid(column=1, row=5, sticky="nw", pady=5)
template_browse.grid(column=1, row=6, sticky="e", pady=5)

r_spacer.grid(column=0, row=7, columnspan=2)

change_defaults.grid(column=0, row=0, sticky="s", padx=5)
reset_defaults.grid(column=1, row=0, sticky="s", padx=5)

# After application loads, fill stored values
reset_stored_data()

root.mainloop()