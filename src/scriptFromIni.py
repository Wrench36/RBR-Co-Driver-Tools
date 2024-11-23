import pandas as pd
import tkinter as tk
from tkinter import filedialog
import configparser

def process_ini(ini_file,script_file):
    config = configparser.ConfigParser()
    config.read(ini_file)
    data = {}
    for section in config.sections():
        if section.startswith("PACENOTE::"):
            name = section.replace("PACENOTE::", "")
            with open(script_file, 'a') as file:
                file.write(f"{name}\n")
    return data

# Main function to handle file selection and sheet processing
def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Ask the user to select the recording Script file
    script_file = filedialog.askopenfilename(title="Select A Script File", filetypes=[("Text Files", "*.txt")])
    
    loop(script_file)
    
    
    
def loop(script_file):
    # Ask the user to select the Excel file
    ini_file = filedialog.askopenfilename(title="Select An ini File", filetypes=[("Config File", "*.ini")])

    if not ini_file:
        print("No file selected, exiting.")
        return
    
    data = process_ini(ini_file,script_file)
    loop(script_file)
    
# Run the main function
if __name__ == "__main__":
    main()
