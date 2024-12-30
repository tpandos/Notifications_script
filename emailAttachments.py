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
for filename in os.listdir(directory):    #for all files in a folder directory
    if filename.endswith(".xlsx"):          #make sure it only takes excel files
        file_path = os.path.join(directory,filename)    #file path using the directory path and the filename joined

        try: 
            df = pd.read_excel(file_path)       #for each file in the directory read the excel file

            supervisor_name = df.iloc[0,6]      #save the supervisor's name from col 6 in a variable
            supervisor_email = df.iloc[0,7]     #save the supervisor's email from col 7 in a variable

            mail = outlook.CreateItem(0)        #CreateItem(0) to create an Mail Item in outlook. 0 = Mail item, 1=Appointment item, 2=Contact Item...
            mail.Subject = f"Subject for {supervisor_name}"     #Subject of the email
            mail.To = supervisor_email                      #address of the recepient, here we are using the super's email
            mail.Body = f"Dear {supervisor_name},\n find the attached file"     #Body of the email

            #attach file
            mail.Attachments.Add(file_path)     #attach the corresponding file

            mail.Save()     #Save as draft

            print(f"draft email created for {supervisor_name} ({supervisor_email}) with attachment: {filename}") #cmd message

        except Exception as e:
            print(f"Failed to process file {filename}: {e}")  #exception handling