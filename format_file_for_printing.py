import tkinter as tk
from tkinter import filedialog
import pandas as pd
from openpyxl import load_workbook
from openpyxl.worksheet.page import PageMargins


#open file explorer to select file to format

root = tk.Tk()
root.withdraw()

input_file = filedialog.askopenfilename(
    title="Select Excel File", 
    filetypes=[("Excel files", "*.xlsx")]
)

if not input_file:
    print("No file selected. Exiting...")
    exit()

#load workbook using openpyxl
wb = load_workbook(input_file)
ws = wb.active

#set page orientation to landscape
ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE

#set margins to narrow
ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.5, bottom=0.5)

#scale the sheet ot fit all columns on the page
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0



#auto adjust width for all cols
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter  # Get the column name
    for cell in col:
        try:
            if cell.value:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width

#select fixed width of first col (emplids)
ws.column_dimensions['A'].width = 10 

#save workbook with the formatting applied
wb.save(input_file)


