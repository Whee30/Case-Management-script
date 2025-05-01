import os
import re
import json
from shutil import copyfile
from tkinter.ttk import Style
from tkinter import filedialog, messagebox, Text, END, Tk, ttk
from docxtpl import DocxTemplate
from datetime import datetime
#import pyi_splash


stored_data = {
    'name':'',
    'unit':'',
    'agency':'',
    'template':'',
    'destination':''
}

# Create compatible json if not already present
if os.path.exists('stored_values.json') is  False:
    # establish initial default values for json
    default_values = {
        'name':'Ofc. J. Smith #123',
        'unit':'Criminal Investigations Unit',
        'agency':'anytown police department',
        'template':'my_template.docx',
        'destination':'D:/'
    }

    # apply defaults to json
    with open('stored_values.json','w') as file:
        json.dump(default_values,file,indent=4)
        
root = Tk()
main_frame = ttk.Frame(root)
main_frame.pack(fill="both", expand=1)

# Style options in TKinter are: 'winnative', 'clam', 'alt', 'default', 'classic', 'vista', 'xpnative'
# print(style.theme_names()) # Shows available styles
style = Style()
style.theme_use("clam")

label = {}
entry = {}
content = {}
illegal_characters = r'[<>:"/|\\?*]'    

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
    messagebox.showinfo("Attention","The stored values have been updated.")

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
    # Set time variable
    start_date = (datetime.now().strftime('%m/%d/%Y'))

    # Remove leading and trailing whitespace, remove tabs, remove empty lines
    # This preserves whitespace insid of a line, for items like "cellphone 1"
    item_content = entry['items'].get(1.0, END).splitlines()
    item_list = [item.strip().replace("\t","") for item in item_content if item.strip()]
    

    clean_case_num = re.sub(illegal_characters, '', entry['case_num'].get())
    clean_case_desc = re.sub(illegal_characters, '', entry['desc'].get())
    case_and_description = clean_case_num + ' ' + clean_case_desc

    casefile_path = os.path.join(stored_data['destination'], case_and_description.strip())

    if entry['case_num'].get() == '':
        messagebox.showerror("Attention","You didn't enter a case number.")
        return

    if len(item_list) == 0:
        messagebox.showerror("Attention","You didn't enter any items, please enter at least one item.")
        return

    if os.path.exists(casefile_path.strip()):
        messagebox.showerror("Attention","This case path already exists, please choose another case number/description.")
        return
    
    if len(item_list) != len(set(item_list)):
        messagebox.showerror("Attention","Two or more of your items are the same value, please change them so each line is unique.")
        return

    os.mkdir(casefile_path) 

    # Create a "reports" directory inside the parent case directory
    report_directory = os.path.join(casefile_path, "Reports")
    os.mkdir(report_directory)

    # create item folders as well as nested export folders for each evidence item entered
    for i in item_list:
        clean_i = re.sub(illegal_characters, '', i)
        joined_path = os.path.join(casefile_path, clean_i)
        os.mkdir(joined_path)
        os.mkdir(os.path.join(joined_path, 'Exports'))

    # create and edit the worklog txt file within the parent case directory
    worklog_name = clean_case_num + " worklog.txt"
    worklog_path = os.path.join(casefile_path, worklog_name)
    worklog_file = open(worklog_path, "w+")

    worklog_file.write(case_and_description + "\n\n")
    worklog_file.write(f"Date Started: {start_date}\n\n")

    for i in item_list:
        worklog_file.write(i + ":\n\n")
    worklog_file.close()

    # Build the .docx
    edit_template = DocxTemplate(entry['template'].get())

    for key, value in entry.items():
        if key == "items":
            content['items'] = entry['items'].get(1.0, END)
        else:
            content[key] = value.get()
    
    content['item_list'] = item_list
    content['start_date'] = start_date

    edit_template.render(content, autoescape=True)
    output_filename = f"{clean_case_num} Examination Report.docx"
    output_path = os.path.join(report_directory, output_filename)
    edit_template.save(output_path)

    case_complete = messagebox.askyesno("Success!",f"The case structure has been generated at: \n\n'{stored_data['destination']}{case_and_description}/'\n\nDo you want to close the program now?")
    
    if case_complete:
        os.startfile(casefile_path)
        root.destroy()

    os.startfile(casefile_path)

#####################
##### Begin GUI #####
#####################

root.title("Case Creation Utility v1.0")
root.minsize(800,360)

left_frame = ttk.Frame(main_frame)
left_frame.pack(side="left", fill="both", expand="1", padx=10, pady=10)

right_frame = ttk.Frame(main_frame, relief="groove")
right_frame.pack(side="top", fill="none", expand="0", ipadx=10, ipady=10, padx=10, pady=10)

r_button_frame = ttk.Frame(right_frame)
r_button_frame.grid(column=0, row=8, columnspan=2)

########################
##### LEFT WIDGETS #####
########################

label['new_info'] = ttk.Label(left_frame, text="New Case Information:", font=("TkDefaultFont", 14))
label['case'] = ttk.Label(left_frame, text="Case Number:")
entry['case_num'] = ttk.Entry(left_frame, width=25, font=("TkDefaultFont", 12))
label['desc'] = ttk.Label(left_frame, text="Case Description:\n(optional)")
entry['desc'] = ttk.Entry(left_frame, width=25, font=("TkDefaultFont", 12))
label['items'] = ttk.Label(left_frame, text="List Items one per line:")
entry['items'] = Text(left_frame, width=25, height=10, font=("TkDefaultFont", 12), relief="sunken")
spacer = ttk.Label(left_frame, text="")
submit_form = ttk.Button(left_frame, text="Create Case", command=generate_case)
clear_form = ttk.Button(left_frame, text="Clear Form", command=reset_form)

#####################
##### LEFT GRID #####
#####################

label['new_info'].grid(column=0, row=0, columnspan=2, sticky="nw", pady=5)
label['case'].grid(column=0, row=2, sticky="nw", padx=10, pady=5)
entry['case_num'].grid(column=1, row=2, columnspan=2, sticky="nw", pady=5)
label['desc'].grid(column=0, row=3, sticky="nw", padx=10, pady=5)
entry['desc'].grid(column=1, row=3, columnspan=2, sticky="nw", pady=5)
label['items'].grid(column=0, row=4, sticky="nw", padx=10, pady=5)
entry['items'].grid(column=1, row=4, columnspan=2, sticky="nw", pady=5)
submit_form.grid(column=1, row=5, sticky="ne", pady=5, padx=5)
clear_form.grid(column=2, row=5, sticky="nw", pady=5, padx=5)

#########################
##### RIGHT WIDGETS #####
#########################

label['stored_info'] = ttk.Label(right_frame, text="Stored Defaults:", font=("TkDefaultFont", 14))
label['name'] = ttk.Label(right_frame, text="Name:")
entry['name'] = ttk.Entry(right_frame, textvariable=stored_data['name'], width=40)
label['unit'] = ttk.Label(right_frame, text="Unit:")
entry['unit'] = ttk.Entry(right_frame, textvariable=stored_data['unit'], width=40)
label['agency'] = ttk.Label(right_frame, text="Agency:")
entry['agency'] = ttk.Entry(right_frame, textvariable=stored_data['agency'], width=40)
label['destination'] = ttk.Label(right_frame, text="Output Dir:")
entry['destination'] = ttk.Entry(right_frame, textvariable=stored_data['destination'], width=40)
output_browse = ttk.Button(right_frame, text="browse", command=output_select)
label['template'] = ttk.Label(right_frame, text="Template:")
entry['template'] = ttk.Entry(right_frame, textvariable=stored_data['template'], width=40)
template_browse = ttk.Button(right_frame, text="browse", command=template_select)
r_spacer = ttk.Label(right_frame, text="")
change_defaults = ttk.Button(r_button_frame, text="Submit Defaults", command=change_default_data)
reset_defaults = ttk.Button(r_button_frame, text="Reset Defaults", command=reset_stored_data)

######################
##### RIGHT GRID #####
######################

label['stored_info'].grid(column=0, row=0, pady=5, padx=5, columnspan=2, sticky="nw")
label['name'].grid(column=0, row=1, sticky="nw", padx=10, pady=5)
entry['name'].grid(column=1, row=1, sticky="nw", pady=5)
label['unit'].grid(column=0, row=2, sticky="nw", padx=10, pady=5)
entry['unit'].grid(column=1, row=2, sticky="nw", pady=5)
label['agency'].grid(column=0, row=3, sticky="nw", padx=10, pady=5)
entry['agency'].grid(column=1, row=3, sticky="nw", pady=5)
label['template'].grid(column=0, row=4, sticky="nw", padx=10, pady=5)
entry['template'].grid(column=1, row=4, sticky="nw", pady=5)
template_browse.grid(column=1, row=5, sticky="e", pady=5)
label['destination'].grid(column=0, row=6, sticky="nw", padx=10, pady=5)
entry['destination'].grid(column=1, row=6, sticky="e", pady=5)
output_browse.grid(column=1, row=7, sticky="e", pady=5)


r_spacer.grid(column=0, row=7, columnspan=2)

change_defaults.grid(column=0, row=0, sticky="s", padx=5)
reset_defaults.grid(column=1, row=0, sticky="s", padx=5)

# After application loads, fill stored values
reset_stored_data()

#pyi_splash.close()
root.mainloop()