import tkinter as tk
import subprocess

def run_generateINI():
    subprocess.run(["python", "src/generateINI.py"])

def run_iniToExcel():
    subprocess.run(["python", "src/iniToExcel.py"])

def run_iniToExcelFromPackage():
    subprocess.run(["python", "src/iniToExcelFromPackage.py"])

def run_scriptGenerator():
    subprocess.run(["python", "src/scriptGenerator.py"])

def run_renameRecodings():
    subprocess.run(["python", "src/renameRecodings.py"])

def run_autorecord():
    subprocess.run(["python", "src/recordAll.py"])

# Set up the main window
root = tk.Tk()
root.title("Co-Driver Tools")

# Add buttons for each action
button1 = tk.Button(root, text="Generate ini from Excel", command=run_generateINI)
button1.pack(pady=10)

button2 = tk.Button(root, text="Generate excel sheet from ini", command=run_iniToExcel)
button2.pack(pady=10)

button3 = tk.Button(root, text="Generate excel from packages ini", command=run_iniToExcelFromPackage)
button3.pack(pady=10)

button4 = tk.Button(root, text="Generate recording script", command=run_scriptGenerator)
button4.pack(pady=10)

button5 = tk.Button(root, text="Auto record from script", command=run_autorecord)
button5.pack(pady=10)

button6 = tk.Button(root, text="Rename recordings", command=run_renameRecodings)
button6.pack(pady=10)

# Start the GUI event loop
root.mainloop()
