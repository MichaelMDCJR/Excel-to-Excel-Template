import pandas as pd
from openpyxl import load_workbook
import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from openpyxl.utils.dataframe import dataframe_to_rows
import json

''' 
X Save settings
X directions at beginning
X comments
error checking
make it look pretty
'''

# Function to read Excel file
def read_excel(file_path):
    try:
        return pd.read_excel(file_path)
    except:
        err = tk.Tk()
        err.title("File Error")
        err.minsize(300, 100)
        err.attributes("-topmost", True)
        directions = Label(
            text="Error Opening File. Make sure the selected file is an Excel file(.xlsx).", wraplength=260, padx=20, pady=10)
        directions.grid(row=0, column=0)

        close_button = Button(text="Okay", command=lambda: [prompt.destroy, exit()])
        close_button.grid(row=1, column=0, pady=10)
        err.mainloop()

# Function to merge data based on mappings
def merge_data(template_df, source_df, column_mappings):
    # Loops through the column mapping and assigns source columns to template
    for target_col, source_col in column_mappings.items():
        template_df[target_col] = source_df[source_col]
    return template_df


# Function to write DataFrame to Excel file with the same formatting
def write_excel_with_template(template_path, df, output_path):
    # Ensure the directory exists
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    # Load the template workbook
    book = load_workbook(template_path)
    sheet = book.active

    # Goes through every row and every cell in the row, copying the format over
    for row_index, row in enumerate(dataframe_to_rows(df, index=False, header=False), 2):
        for col_index, value in enumerate(row, 1):
            # Gets a specific cell and its formating
            cell = sheet.cell(row=row_index, column=col_index, value=value)

    # Save the workbook
    if not output_path.endswith(".xlsx"):
        output_path += ".xlsx"
    book.save(output_path)
    root.destroy()

# Write Excel file
def end_script():
    merged_df = merge_data(template_df, source_df, mapping)
    write_excel_with_template(template_file, merged_df, output_path)
    print(f"Data successfully written to {output_path}")
    return

# Add selected listbox to entry
def add_entry():
    # Clear entry boxes
    template_name_entry.delete(0, tk.END)
    source_index_entry.delete(0, tk.END)
    # Add to entry without index
    temp_sel = template_listbox.curselection()
    sour_sel = source_listbox.curselection()
    if not temp_sel and not sour_sel:
        return
    template_name_entry.insert(0, template_listbox.get(temp_sel)[3:])
    source_index_entry.insert(0, source_listbox.get(sour_sel)[3:])
    # Add mapping to dict
    user_template_string = template_name_entry.get()
    user_source_string = source_index_entry.get()
    mapping[user_template_string] = user_source_string
    # Add mapping to listbox
    mappings_listbox.insert(mappings_listbox.size(), str(mappings_listbox.size()) + ": " + user_source_string + " mapped to " + user_template_string)

# Delete selected mapping
def delete_map():
    # Delete mapping from listbox
    selection = mappings_listbox.curselection()
    if not selection:
        return
    to_be_del = mappings_listbox.get(selection)
    to_be_del_index = int(to_be_del[0])
    mappings_listbox.delete(to_be_del_index)

    # Get template column name
    template_del = (to_be_del.split(" mapped to ", 1)[1])

    # Delete mapping from dict
    del mapping[template_del]
    # Clear listbox and put the rest of the mappings back in
    mappings_listbox.delete(0, END)
    temp_counter = 0
    for i in mapping:
        mappings_listbox.insert(temp_counter,str(temp_counter) + ": " + str({mapping[i]})[2:-2] + " mapped to " + str({i})[2:-2])
        temp_counter = temp_counter + 1

# Prompts the user to choose a name to save the file
def save_prompt():
    save_window = tk.Tk()
    save_window.title("Enter File Name")
    save_window.minsize(320, 100)
    save_window.attributes("-topmost", True)
    save_window.config(bg="#CE1126")
    # Save entry
    save_label = Label(save_window, text="Enter the name you wish to save your mapping to:", font=("Arial", 12), bg="#CE1126")
    save_label.grid(row=0, column=0)
    save_entry = tk.Entry(save_window)
    save_entry.grid(row=1, column=0, pady=10, padx=10)

    # Save button
    save_real_button = tk.Button(save_window, text="Save", font=("Arial", 12), command=lambda: [save_file(save_entry.get()), save_window.destroy()])
    save_real_button.grid(row=2, column=0)
    save_window.mainloop()

# Saves the file
def save_file(name):
    # Creates json file and dumps info into it
    json_dict = json.dumps(mapping)
    with open(name + ".json", "w") as saved_json:
        json.dump(json_dict, saved_json)

# Loads the selected json file
def load_file():
    global mapping
    # Loads selected json
    selected_json = filedialog.askopenfilename(initialdir=os.getcwd(), title="Load Mapping", filetypes=[("Json", "*.json")])
    # If 'cancel' was selected, json will be empty
    if not selected_json:
        return
    # Opens json and reads it to data
    with open(selected_json, "r") as dict_saved:
        data = json.load(dict_saved)
        # Turns json back to dict
        dict_data = json.loads(data)
        mapping = dict_data
        # Puts the dict into the listbox
        temp_counter = 0
        for i in mapping:
            mappings_listbox.insert(temp_counter,
                                    str(temp_counter) + ": " + str({mapping[i]})[2:-2] + " mapped to " + str({i})[2:-2])
            temp_counter = temp_counter + 1

# Initial prompt telling users how to  use the program
prompt = tk.Tk()
prompt.title("Directions")
prompt.minsize(500,220)
directions = Label(text="A set of file explorers will pop up after this. In the first window, select your source file. "
                        "This will be the file where information is drawn from. In the second window, select your "
                        "template Excel file. This will be the file your information is placed in, according to the "
                        "template. Finally, you will have to select the name of the new file being created and where "
                        "you want to store it. You can write over an existing file if needed.", font=("Arial", 14),wraplength=560, padx=20, pady=10)
directions.grid(row=0, column=0)

close_button = Button(text="Proceed", command=prompt.destroy)
close_button.grid(row=1, column=0, pady=10)

prompt.mainloop()

# Dictionary to store mapping of columns
mapping = {}

# Prompt user for source file, template, and output location and name
try:
    source_file = filedialog.askopenfilename(title="Source File")
    template_file = filedialog.askopenfilename(title="Template File")
    output_path = filedialog.asksaveasfilename(title="Output File Name and Location")

except:
    prompt = tk.Tk()
    prompt.title("File Error")
    prompt.minsize(300, 150)
    prompt.attributes("-topmost", True)
    directions = Label(
        text="Error Opening File", wraplength=260, padx=20, pady=10)
    directions.grid(row=0, column=0)

    close_button = Button(text="Okay", command=prompt.destroy)
    close_button.grid(row=1, column=0, pady=10)

# Read excel files to df
source_df = read_excel(source_file)
template_df = read_excel(template_file)

# Create tkinter window
root = tk.Tk()
root.title("Excel to Excel Converter")
root.minsize(1200, 600)
root.attributes("-topmost", True)
root.config(bg="#CE1126")

# List box labels
template_label = Label(root, text="Template Columns", bg="#CE1126")
template_label.grid(row=0, column=0, pady=10)
template_label.config(font=("Arial",20,"bold"))
source_label = Label(root, text="Source Columns", bg="#CE1126")
source_label.grid(row=0, column=1, pady=10)
source_label.config(font=("Arial",20,"bold"))
map_label = Label(root, text="Mapping", bg="#CE1126")
map_label.grid(row=0, column=2, pady=10)
map_label.config(font=("Arial",20,"bold"))

# create the template listbox
template_listbox = Listbox(root, selectmode=SINGLE, exportselection=False, width=40, height=20, bg="#FFFFFF", relief=GROOVE)
template_listbox.config(font=("Arial",12))

# Put df columns in list box
for i in range(template_df.columns.size):
    template_listbox.insert(i, str(i) + ": " + template_df.columns[i])

template_listbox.grid(row=1, column=0, padx=30)

# create the source index listbox
source_listbox = Listbox(root, selectmode=SINGLE, exportselection=False, width=40, height=20, bg="#FFFFFF", relief=GROOVE)
source_listbox.config(font=("Arial",12))

# Put df columns in list box
for i in range(source_df.columns.size):
    source_listbox.insert(i, str(i) + ": " + source_df.columns[i])

source_listbox.grid(row=1, column=1)

# Create the mappings list box
mappings_listbox = Listbox(root, selectmode=SINGLE, exportselection=False, width=40, height=20, bg="#FFFFFF", relief=GROOVE)
mappings_listbox.grid(row=1, column=2, padx=30)
mappings_listbox.config(font=("Arial",12))

# Entry boxes
template_name_entry = tk.Entry(root)
source_index_entry = tk.Entry(root)

# Save button
save_button = tk.Button(root, text="Save Current Mapping", command=save_prompt)
save_button.config(font=("Arial",12))
save_button.grid(row=4, column=0)

# Load button
load_button = tk.Button(root, text="Load Mapping", command=load_file)
load_button.config(font=("Arial",12))
load_button.grid(row=5, column=0)

# Add selected to entry box button
entry_button = tk.Button(root, text="Add", command=add_entry)
entry_button.config(font=("Arial",12))
entry_button.grid(row=4, column=1, pady=20)

# End script button
end_button = tk.Button(root, text="Create Excel File",command=end_script)
end_button.config(font=("Arial",12))
end_button.grid(row=5, column=1, pady=30)

# Delete button
delete_button = tk.Button(root, text="Delete mapping", command=delete_map)
delete_button.config(font=("Arial",12))
delete_button.grid(row=4, column=2, pady=10)

root.mainloop()

