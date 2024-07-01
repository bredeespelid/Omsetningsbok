import os
import pandas as pd
from glob import glob
import re
import datetime
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# Create the Tkinter root window
root = tk.Tk()
root.withdraw()  # Hide the root window

# Prompt the user to select an Excel file
input_file_path = filedialog.askopenfilename(
    title="Select an Excel file",
    filetypes=[("Excel files", "*.xlsx")]
)

if not input_file_path:
    messagebox.showerror("Error", "No file selected.")
    exit()

# Read the selected Excel file, skipping the first 7 rows
df = pd.read_excel(input_file_path, skiprows=7)

# Add a new column to the left of "Navn"
df.insert(0, 'Avd', '')

# Function to extract numeric values from "Navn"
def extract_number(name):
    match = re.match(r'(\d+)', str(name))
    return int(match.group(1)) if match else None

# Fill in the "Avd" column with numeric values from "Navn"
current_number = None
for index, row in df.iterrows():
    number = extract_number(row['Navn'])
    if number is not None:
        current_number = number
    df.at[index, 'Avd'] = current_number

# Fill any remaining NaN values with the last valid values
df['Avd'] = df['Avd'].fillna(method='ffill')
df['Avd'] = df['Avd'].fillna(0).astype(int)

# Function to filter out rows based on "Konto" criteria
def filter_konto(konto):
    if pd.isna(konto) or konto == '' or konto == ' ':
        return False
    if isinstance(konto, str) and len(konto) == 1 and konto.isalpha():
        return False
    if isinstance(konto, str) and konto.strip() == '1529':
        return False
    return True

# Filter the DataFrame based on the "Konto" column
df = df[df['Konto'].apply(filter_konto)]

# Convert cells below and to the right of D1 to integers where possible
for col in df.columns[3:]:  # Column D is the fourth column (index 3)
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

# Convert column names from D1 onwards to yyyy/mm/dd format
new_columns = {}
for col in df.columns[3:]:
    try:
        # Use the format 'dd.mm.yyyy' and convert to 'yyyy/mm/dd'
        date_val = pd.to_datetime(col, format='%d.%m.%Y')
        # Save in the correct format
        new_columns[col] = date_val.strftime("%Y/%m/%d")
    except ValueError:
        pass  # Skip columns that cannot be converted

# Update the column names in the DataFrame
df.rename(columns=new_columns, inplace=True)

# Filter out rows with numbers 10 and 90 in the "Avd" column
df = df[~df['Avd'].isin([10, 90])]

# Prompt the user to select a save location for the new Excel file
output_file_path = filedialog.asksaveasfilename(
    title="Save Excel file",
    defaultextension=".xlsx",
    filetypes=[("Excel files", "*.xlsx")]
)

if not output_file_path:
    messagebox.showerror("Error", "No save location specified.")
    exit()

# Save the DataFrame to a new Excel file with formulas
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Sheet1')
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    
    # Add DATEVALUE formulas in the column headers
    for col_num, value in enumerate(df.columns.values):
        if isinstance(value, str) and re.match(r'^\d{4}/\d{2}/\d{2}$', value):
            worksheet.write_formula(0, col_num, f'=DATEVALUE("{value}")')

messagebox.showinfo("Success", "The Excel file has been saved successfully.")
