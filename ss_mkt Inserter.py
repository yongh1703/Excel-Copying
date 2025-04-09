import xlwings as xw
import tkinter as tk
import os
import re
import sys
from tkinter import filedialog

# File path for the target workbook
target_file = 'File Path'

# Select a source file using a file dialog
def select_source_file(title="Select the Source Excel File"):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel Files", "*.xlsx;*.xlsm")],
    )
    return file_path

# Extract date in DD-MM-YYYY format from the filename
def extract_date_from_filename(file_path):
    filename = os.path.basename(file_path)
    match = re.search(r'(\d{2}-\d{2}-\d{4})', filename)
    if match:
        return match.group(1)
    else:
        raise ValueError("Date not found in the source file name")

# === First Source File Operation === #
source_file_1 = select_source_file("Select the Archived Global Supply")
source_sheet_1 = 'Supply by Source & Status'
target_sheet = 'ss_mkt'

# === Second Source File Operation === #
source_file_2 = select_source_file("Select the Archived Monthly Supply")
source_sheet_2 = 'Monthly- Project kt'  

app = xw.App(visible=False)
try:
    wb_source1 = app.books.open(source_file_1, read_only=True)
    wb_source2 = app.books.open(source_file_2, read_only=True)
    wb_target = app.books.open(target_file)

    # Validate sheets
    if source_sheet_1 not in [s.name for s in wb_source1.sheets]:
        print(f"Sheet '{source_sheet_1}' not found in {source_file_1}")
        sys.exit()

    if source_sheet_2 not in [s.name for s in wb_source2.sheets]:
        print(f"Sheet '{source_sheet_2}' not found in {source_file_2}")
        sys.exit()

    ws_source1 = wb_source1.sheets[source_sheet_1]
    ws_source2 = wb_source2.sheets[source_sheet_2]
    ws_target = wb_target.sheets[target_sheet]

    # === Copy from first file === #
    data1 = ws_source1.range('A113:AK141').value
    ws_target.range('A37').value = data1
    date_str1 = extract_date_from_filename(source_file_1)
    ws_target.range('A35').value = f"Global Supply {date_str1}"

    # === Copy from second file === #
    data2 = ws_source2.range('BS110:DN138').value  
    ws_target.range('CV37').value = data2       
    date_str2 = extract_date_from_filename(source_file_2)
    ws_target.range('AM35').value = f"Monthly Supply {date_str2}"

    # Save and close everything
    wb_target.save()
    wb_source1.close()
    wb_source2.close()
    wb_target.close()

finally:
    print(f"Extracted From {os.path.basename(source_file_1)} and {os.path.basename(source_file_2)}")
    app.quit()
