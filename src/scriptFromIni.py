import pandas as pd
import tkinter as tk
from tkinter import filedialog
import configparser
import os

def process_ini(ini_file,script_file):
    config = configparser.ConfigParser()
    config.read(ini_file)
    data = {}
    for section in config.sections():
        section = section.replace("PACENOTE::", "")
        section = section.replace("RANGE::", "")
        with open(script_file, 'a') as file:
            if pd.notna(section):
                file.write(f"{section}\n")
    remove_duplicate_lines(script_file)
    return data

# Main function to handle file selection and sheet processing
def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    parent_dir = os.path.dirname(os.path.abspath(__file__)) 
    docs_dir = os.path.abspath(os.path.join(parent_dir, "../"))
    
    # Ask the user to select the recording Script file
    script_file = filedialog.askopenfilename(initialdir=docs_dir,title="Select A Script File", filetypes=[("Text Files", "*.txt")])
    
    loop(script_file)

def remove_duplicate_lines(filename):
    with open(filename, 'r') as f:
        lines = f.readlines()

    unique_lines = []
    seen_lines = set()

    for line in lines:
        if line not in seen_lines:
            unique_lines.append(line)
            seen_lines.add(line)

    with open(filename, 'w') as f:
        f.writelines(unique_lines)
    
    
def loop(script_file):
    parent_dir = os.path.dirname(os.path.abspath(__file__)) 
    packages_dir = os.path.abspath(os.path.join(parent_dir, "../../config/pacenotes/packages"))
    # Ask the user to select the Excel file
    ini_file = filedialog.askopenfilename(initialdir=packages_dir,title="Select An ini File", filetypes=[("Config File", "*.ini")])

    if not ini_file:
        print("No file selected, exiting.")
        
        return
    
    data = process_ini(ini_file,script_file)
    loop(script_file)
    
# Run the main function
if __name__ == "__main__":
    main()
