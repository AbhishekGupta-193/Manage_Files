import os
import shutil
import pandas as pd

# Define the paths
desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
source_folder = os.path.join(desktop_path, "old")
destination_folder = os.path.join(desktop_path, "new")
excel_file_path = os.path.join(desktop_path, "file_names.xlsx")

# Create the destination folder if it doesn't exist
if not os.path.exists(destination_folder):
    os.makedirs(destination_folder)

# Load the Excel file
df = pd.read_excel(excel_file_path)

# Iterate through the DataFrame and move/rename files
for index, row in df.iterrows():
    old_name = row['old_filename']
    new_name = row['new_filename']
    print(old_name)
    print(new_name)
    source_file = os.path.join(source_folder, old_name)
    destination_file = os.path.join(destination_folder, new_name)
    print(source_file)

    if os.path.exists(source_file):
        shutil.move(source_file, destination_file)
        print(f"Moved and renamed: {old_name} to {new_name}")
    else:
        print(f"File not found: {old_name}")

print("All files have been moved and renamed successfully.")
