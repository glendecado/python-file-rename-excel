from pathlib import Path
import pandas as pd
import tkinter as tk
from tkinter import filedialog



# GUI Function
def gui():
    # Create the main window
    root = tk.Tk()
    root.title("Directory Selector")
    root.geometry("300x100")

    # Create a button to trigger directory selection
    select_button = tk.Button(root, text="Select Directory", command=select_directory)
    select_button.pack(pady=20)
    root.mainloop()


# Function to rename files and save them to an Excel file
def rename_subfiles_and_save_to_excel(directory, prefix, suffix, excel_file):
    # Create a Path object for the directory
    path = Path(directory)
    renamed_files = []  # List to store renamed file names

    # Iterate through each item in the directory
    for item in path.iterdir():
        if item.is_file():  # Check if it's a file (not a directory)
            # Construct the new name by adding the prefix and suffix
            new_name = prefix + item.name + suffix
            new_path = item.with_name(new_name)  # Create a new path with the new name

            # Rename the item and add to the list
            try:
                item.rename(new_path)
                renamed_files.append(new_name)
                print(f'Renamed: {item.name} to {new_name}')
            except Exception as e:
                print(f'Error renaming {item.name}: {e}')

    # Save the renamed file names to an Excel file
    df = pd.DataFrame(renamed_files,)
    df.to_excel(excel_file, index=False)
    print(f'Saved renamed files to {excel_file}')



# Directory selection and renaming function
def select_directory():
    # Open a dialog to select a directory
    directory = filedialog.askdirectory()

    # Print the selected directory path
    if directory:
        directory_path = directory  # Replace with your directory path
        excel_file_path = Path(directory) / 'renamed_files.xlsx'  # Path for the output Excel file
        rename_subfiles_and_save_to_excel(directory_path, 'LMSCD_', '_exe.pdf', excel_file_path)



# Main function to run the GUI
def main():
    gui()

# Run the script
if __name__ == "__main__":
    main()
