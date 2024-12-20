import tkinter as tk
from tkinter import filedialog
import pandas as pd 

#function 
def remove_top_rows(file_path, rows_to_remove):
    #load excel file
    df = pd.read_excel(file_path)

    #remove rows by index
    df_cleaned = df.drop(rows_to_remove)

    #Reset indexing and drop the old index col
    df_cleaned = df_cleaned.reset_index(drop=True)

    #Save the modified DataFrame to a new excel file 
    df_cleaned.to_excel(file_path, index=False)
    print(f"Rows removed successfully. Output saved to {file_path}")

root = tk.Tk()
root.withdraw()

#open file dialog to select file
input_file = filedialog.askopenfilename(
    title="Select an Excel File", 
    filetypes=[("Excel files", "*.xlsx")]
)

#check in case input file not selected
if not input_file:
    print("No file selcted. Existing...")
    exit()

#-------------- Modifications to the excel file

#remove top rows, just leave data
rows_to_remove = list(range(1,14))

remove_top_rows(input_file, rows_to_remove)

print("This is the end")
input("Press Enter to exit")