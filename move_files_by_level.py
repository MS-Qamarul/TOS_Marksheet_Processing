import os
import shutil

def move_excel_files_by_level(directory, new_folder_name):
    
    # Create the new folder if it doesn't exist
    new_folder_path = os.path.join(directory, new_folder_name)
    if not os.path.exists(new_folder_path):
        os.makedirs(new_folder_path)

    # List all the files in the directory
    files = os.listdir(directory)

    # Move Excel files to the new folder
    for file in files:
        if file.endswith(".xlsx") and "-" + new_folder_name in file:
            source_path = os.path.join(directory, file)
            destination_path = os.path.join(new_folder_path, file)
            shutil.move(source_path, destination_path)
            print("Moved " + file + " to " + new_folder_name + " folder")

if __name__ == "__main__":
    new_folder_name = input(">>> Enter level: ")
    # Replace 'directory_path' with the path to the directory containing the Excel files
    directory_path = "./1_GoodClass"
    move_excel_files_by_level(directory_path, new_folder_name)
