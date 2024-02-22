import pandas as pd
import shutil
import os
from tkinter import Tk, filedialog
from tkinter import messagebox

# Function to ask the user to select the parent folder
def select_parent_folder():
    root = Tk()
    root.withdraw()
    parent_folder = filedialog.askdirectory(title="Select Parent Folder")
    root.destroy()
    return parent_folder

# Directory paths
GoodClassPath = './1_GoodClass'
NGClassPath = './2_NGClass'

# Path to the Excel file containing the list of files to move
FileListPath = '0_File_List/FileList.xlsx'
FileList_df = pd.read_excel(FileListPath)

# Get the parent folder directory from the user
ParentFolder = select_parent_folder()

# Check whether directories already exist
if not os.path.exists(GoodClassPath):
    os.mkdir(GoodClassPath)
    print("Folder %s created!" % GoodClassPath)
else:
    print("Folder %s already exists" % GoodClassPath)

if not os.path.exists(NGClassPath):
    os.mkdir(NGClassPath)
    print("Folder %s created!" % NGClassPath)
else:
    print("Folder %s already exists" % NGClassPath)

# Read the Excel file into a pandas DataFrame
try:
    print("Processing files, please wait...")
    for foldername, subfolders, filenames in os.walk(ParentFolder):
        for filename in filenames:
            if filename.endswith(".xlsx"):
                file_path = os.path.join(foldername, filename)
                try:
                    # Move the file to the GoodClassPath if it exists in the FileList_df
                    if filename in FileList_df['File Name'].values:
                        destination_file_path = os.path.join(GoodClassPath, filename)
                        shutil.move(file_path, destination_file_path)
                    # Otherwise, move it to the NGClassPath
                    else:
                        destination_file_path = os.path.join(NGClassPath, filename)
                        shutil.move(file_path, destination_file_path)
                except:
                    print(f"Error moving file: {file_path}")

    messagebox.showinfo("Success", "Successfully moved Good Classes and NG Classes to the respective folders!")
    print("Successfully moved Good Classes and NG Classes to the respective folders!")

# Throw a useful error message for missing path
except:
    messagebox.showerror("Missing files or folder", "0_File_List folder or FileList.xlsx is missing! Please create the relevant folder/files before running!")
    print("0_File_List folder or FileList.xlsx is missing! Please create the relevant folder/files before running!")
