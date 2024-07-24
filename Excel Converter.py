import pandas as pd
from openpyxl import load_workbook
import os
import tkinter as tk
from tkinter import *
from tkinter import filedialog

# Function to read Excel file
def read_excel(file_path):
    return pd.read_excel(file_path)

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

    # Get the template
    book = load_workbook(template_path)

    # Use the workbook with pandas ExcelWriter and add .xlsx if the file doesn't have it already
    if output_path.endswith(".xlsx"):
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=book.sheetnames[0])
        root.destroy()
    else:
        with pd.ExcelWriter(output_path + ".xlsx", engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name=book.sheetnames[0])
        root.destroy()

# When the button is pressed add the info to the mapping unless 'done' was submitted, then merge the df and write to excel
def pressed():
    user_template_string = template_name_entry.get()
    # if user_template_string.lower() == 'done':
    #     merged_df = merge_data(template_df, source_df, mapping)
    #     write_excel_with_template(template_file, merged_df, output_path)
    #     print(f"Data successfully written to {output_path}")
    #     return
    user_source_index = source_index_entry.get()
    mapping[user_template_string] = user_source_index
    template_name_entry.delete(0, tk.END)
    source_index_entry.delete(0, tk.END)

def end_script():
    merged_df = merge_data(template_df, source_df, mapping)
    write_excel_with_template(template_file, merged_df, output_path)
    print(f"Data successfully written to {output_path}")
    return

# Dictionary to store mapping of columns
mapping = {}

# Prompt user for source file, template, and output location and name
source_file = filedialog.askopenfilename(title="Source File")
template_file = filedialog.askopenfilename(title="Template File")
output_path = filedialog.asksaveasfilename(title="Output File Name and Location")

# Read excel files to df
source_df = read_excel(source_file)
template_df = read_excel(template_file)

# Create tkinter window
root = tk.Tk()
root.title("Excel to Excel Converter")
root.minsize(480, 300)

# List box labels
template_label = Label(root, text="Template Columns").grid(row=0, column=0)
source_label = Label(root, text="Source Columns").grid(row=0, column=1)

# create the template listbox
template_listbox = Listbox(root, selectmode=SINGLE, exportselection=False)

# Put df columns in list box
for i in range(template_df.columns.size):
    template_listbox.insert(i, str(i) + ": " + template_df.columns[i])

template_listbox.grid(row=1, column=0)

# create the source index listbox
source_listbox = Listbox(root, selectmode=SINGLE, exportselection=False)

# Put df columns in list box
for i in range(source_df.columns.size):
    source_listbox.insert(i, str(i) + ": " + source_df.columns[i])

source_listbox.grid(row=1, column=1)

# Label for entry boxes
template_entrybox_label = Label(root, text="Enter the template column name (or 'done' to finish): ")
template_entrybox_label.grid(row=2, column=0)
source_entrybox_label = Label(root, text="Enter the source column index: ")
source_entrybox_label.grid(row=3, column=0)

# Entry boxes
template_name_entry = tk.Entry(root)
template_name_entry.grid(row=2, column=1)
source_index_entry = tk.Entry(root)
source_index_entry.grid(row=3, column=1)

# Add selected listbox to entry
def add_entry():
    template_name_entry.delete(0, tk.END)
    source_index_entry.delete(0, tk.END)
    # Add to entry without index
    template_name_entry.insert(0, template_listbox.get(template_listbox.curselection())[3:])
    source_index_entry.insert(0, source_listbox.get(source_listbox.curselection())[3:])

# Add selected to entry box button
entry_button = tk.Button(root, text="Add Selected", command=add_entry)
entry_button.grid(row=4, column=0)

# Submit button
submit_button = tk.Button(root, text="Submit", command=pressed)
submit_button.grid(row=4, column=1)

# End script button
end_button = tk.Button(root, text="Create Excel File",command=end_script)
end_button.grid(row=6, column=0, pady=10)

root.mainloop()

