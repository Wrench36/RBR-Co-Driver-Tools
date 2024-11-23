import tkinter as tk
from tkinter import filedialog
import os

def select_files():
    """Allow the user to select multiple files."""
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_paths = filedialog.askopenfilenames(title="Select Files")
    return file_paths

def get_txt_lines(txt_file):
    """Read the lines from a .txt file and return as a list."""
    with open(txt_file, 'r') as file:
        return [line.strip() for line in file.readlines()]

def rename_files(files, txt_lines):
    """Rename the files to match the lines from the .txt file."""
    if len(files) != len(txt_lines):
        print(f"Error: The number of files ({len(files)}) does not match the number of lines in the .txt file ({len(txt_lines)}).")
        return
    
    for i, file in enumerate(files):
        new_name = os.path.join(os.path.dirname(file), f"{txt_lines[i]}{os.path.splitext(file)[1]}")
        try:
            os.rename(file, new_name)
        except FileExistsError:
            print('')
        print(f"Renamed: {file} -> {new_name}")
    print('done')

def main():
    # Select files using tkinter
    selected_files = select_files()

    # Ask for the .txt file containing names
    txt_file = filedialog.askopenfilename(title="Select Text File with Names", filetypes=[("text file", "*.txt")])

    # Get lines from the text file
    txt_lines = get_txt_lines(txt_file)

    # Rename files if the counts match
    rename_files(selected_files, txt_lines)

if __name__ == "__main__":
    main()
