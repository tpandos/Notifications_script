import os
import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog


#ask user to select the folder
root = tk.Tk()
root.withdraw()
directory = filedialog.askdirectory(
    title="Select folder containing Exel files"
)

if not directory:
    print("No directory selected. Exiting...")
    exit()

#initializing outlook application
outlook = win32.Dispatch('Outlook.Application')

#loop through every file in the directory
for filename in os.listdir(directory):
    if filename.endswith(".xlsx"):
        file_path = os.path.join(directory,filename)

        try: 
            df = pd.read_excel(file_path)

            supervisor_name = df.iloc[0,6]
            supervisor_email = df.iloc[0,7]

            mail = outlook.CreateItem(0)
            mail.Subject = f"Subject for {supervisor_name}"
            mail.To = supervisor_email
            mail.Body = f"Dear {supervisor_name},\n find the attached file"

            #attach file
            mail.Attachments.Add(file_path)

            mail.Save()

            print(f"draft email created for {supervisor_name} ({supervisor_email}) with attachment: {filename}")

        except Exception as e:
            print(f"Failed to process file {filename}: {e}")