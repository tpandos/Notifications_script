import tkinter as tk
from tkinter import filedialog
import pandas as pd 

#get the promt window to open to choose file

root = tk.Tk()
root.withdraw()

input_file = filedialog.askopenfilename(
    title="Select Excel File", 
    filetypes=[("Excel files", "*.xlsx")]
)

if not input_file: 
    print("No file selected. Exiting...")
    exit()

##----------------------------------
#read the file and set it in a dataframe without setting a header
excel_data = pd.read_excel(input_file, header=None)

#remove the first 14 rows 
data_clean_headers = excel_data.iloc[14:].reset_index(drop=True)

#make the 15th row, the new header
data_clean_headers.columns = data_clean_headers.iloc[0]

#reset index for final excel data 
data_clean_headers = data_clean_headers.drop(0).reset_index(drop=True)

#remove blank cols (C, F), remove Business Unit (E), Training Date (M)
#0-based index, columns 2, 4, 5, 12 

cols_to_remove = [2, 4, 5, 12]
#df = df.drop(specify cols to drop df.columns[cols_to_remove], axis=1)
final_data = data_clean_headers.drop(data_clean_headers.columns[cols_to_remove], axis=1)

#save the modified file
final_data.to_excel(input_file, index=False)

print("This is the end")
input("Press Enter to exit")