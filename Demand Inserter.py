import xlwings as xw
import tkinter as tk
import os
import re
import sys
from tkinter import filedialog

# File path
target_file = 'File Path'  

# Set the location to open the file dialog
def select_source_file():
    # Initialize tkinter
    root = tk.Tk()
    root.withdraw()  # Don't need a full GUI, just the file dialog

    # Open the file dialog to select the source file
    file_path = filedialog.askopenfilename(
        title="Select the Source Excel File",
        filetypes=[("Excel Files", "*.xlsx;*.xlsm")],  # Allow only Excel files
    )
    return file_path

# Extract the date from the file name using regex (assuming format is 'Global LNG Demand YYYY-MM-DD.xlsx')
def extract_date_from_filename(file_path):
    filename = os.path.basename(file_path)  # Extracts the filename from the full path
    match = re.search(r'(\d{4}-\d{2}-\d{2})', filename)  # Look for the date pattern in the filename
    if match:
        return match.group(1)  # Return the matched date string
    else:
        raise ValueError("Date not found in the source file name")

# Ask the user to select the source file
source_file = select_source_file("Select the Archived Global Demand")

# Excel sheet names (update if needed)
source_sheet = 'dd_fcast'
target_sheet = 'dd_mkt'

# Open both workbooks
app = xw.App(visible=False)
try:
    # Open source and target workbooks
    wb_source = app.books.open(source_file, read_only=True)

    # ✅ Check if the required sheet exists
    if source_sheet not in [s.name for s in wb_source.sheets]:
        print("Wrong File: Sheet 'dd_fcast' not found.")
        wb_source.close()
        app.quit()
        sys.exit()  # Stop execution

    wb_target = app.books.open(target_file)

    ws_source = wb_source.sheets[source_sheet]
    ws_target = wb_target.sheets[target_sheet]

    # Read data from source range G9:SD77
    data = ws_source.range('G9:SD77').value  # 2D list
    data_year = ws_source.range('SF9:TT77').value

    # Paste into target starting at E78
    ws_target.range('E78').value = data
    ws_target.range('SD78').value = data_year

    # Extract date from the selected file name
    date_str = extract_date_from_filename(source_file)

    # Update C77 with the date string
    ws_target.range('C77').value = date_str

    # Save and close
    wb_target.save()
    wb_source.close()
    wb_target.close()
finally:
    print(f"Extracted From {os.path.basename(source_file)}")
    app.quit()
