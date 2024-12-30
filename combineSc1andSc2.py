import tkinter as tk
from tkinter import filedialog
import pandas as pd 
from tkinter import messagebox

#get the promt window to open to choose file

root = tk.Tk()
root.withdraw()

response = messagebox.askyesno("Confirm Action ", "Are you sure you want to continue?")

if response:
    print("Action confirmed, Executing script")
else:
    print("Action cancelled. Exiting...")
    exit()

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

#filter column I (8) completed values and remove
#filter final_data in col 8 to the values NOT "Completed" now final_data only has other than "not completed"
final_data = final_data[final_data.iloc[:,8] != "Completed"]

#save the modified file
final_data.to_excel(input_file, index=False)

column_to_filter = "Supervisor Name"

unique_values = final_data[column_to_filter].unique()

save_directory = filedialog.askdirectory(
    title="Select the directory to save the files"
)

if not save_directory:
    print("No directory selected. Exiting...")
    exit()

for value in unique_values:
    split_df = final_data[final_data[column_to_filter] == value]
    output_file = f"{save_directory}/{str(value)}.xlsx"
    split_df.to_excel(output_file, index=False)

print(f"Filtered file for '{value}' saved as: {output_file}")

