import os
import glob
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# Path to the folder containing the Excel files
folder_path = '/Users/ethan/Library/CloudStorage/OneDrive-Personal/CampusArtistry/Product Creation/Erank Data'

# Get the most recent Excel file from the folder
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))
most_recent_file = max(excel_files, key=os.path.getctime)

# Load the most recent Excel file
workbook = load_workbook(most_recent_file)

# Get the first sheet of the workbook
ws = workbook.active

# Create a new sheet within the same workbook
ns = workbook.create_sheet(title='Data')

# Copy the data from the source sheet to the new sheet
for row in ws.iter_rows(values_only=True):
    ns.append(row)

# Get the range of cells containing the copied data
num_rows = ns.max_row
num_columns = ns.max_column
start_cell = ns.cell(row=1, column=1)
end_cell = ns.cell(row=num_rows, column=num_columns)

# Create a table from the range of cells
table = Table(displayName="Table1", ref=f"{start_cell.coordinate}:{end_cell.coordinate}")

# Apply table style
table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9")

# Add the table to the new sheet
ns.add_table(table)

# Save the modified workbook
workbook.save(most_recent_file)