import tkinter as tk
from tkinter import filedialog
import pandas as pd 

#tkinter window file explorer
root = tk.Tk()
root.withdraw()  #hide main window

#open file dialog to select file
input_file = filedialog.askopenfilename(
    title="Select an Excel File", 
    filetypes=[("Excel files", "*.xlsx")]
)

if not input_file: #if no file selected, exit and print 
    print("No file selected. Exiting...")
    exit()

#load file to pandas DataFrame
df = pd.read_excel(input_file)

################ MODIFICATIONS TO THE EXEL FILE####################################################################

#----------------------- Sort spreadsheet by the values on the "Supervisor Name" column in ascending order
#sort_column = "Supervisor Name" # name of the column to sort
#df_sorted = df.sort_values(by=sort_column, ascending=True) #method to sort values by column df_sorted is the data to be saved in excel in line: 

#--------------------------------------------Filter by unique value in col and extract those values in separate spreadsheets

#specify col to filter
column_to_filter = "Supervisor Name"

#get all unique values from the col we want to filter
unique_values = df[column_to_filter].unique()



###############################################

#select save location
save_directory = filedialog.askdirectory(
    title="Select the directory to save the files"
)

#if not file path is selected
if not save_directory: 
    print("No directory selected. Exiting")
    exit()

#filter and save df for each unique value entries 
for value in unique_values:
    #filter df by current value
    filtered_df = df[df[column_to_filter] == value]
    #construct the filename using the filtred value
    output_file = f"{save_directory}/{str(value)}.xlsx"
    #save the filtered df to the specific file
    filtered_df.to_excel(output_file, index=False)

    print(f"Filtered file for '{value}' saved as: {output_file}")

