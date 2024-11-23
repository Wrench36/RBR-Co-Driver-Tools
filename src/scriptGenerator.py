import tkinter as tk
import subprocess

def run_scriptFromExcel():
    subprocess.run(["python", "src/scriptFromExcel.py"])

def run_scriptFromIni():
    subprocess.run(["python", "src/scriptFromIni.py"])

# Set up the main window
root = tk.Tk()
root.title("Recording Script Generator")

# Add buttons for each action
button1 = tk.Button(root, text="Generate recording script from excel", command=run_scriptFromExcel)
button1.pack(pady=10)

button2 = tk.Button(root, text="Generate recording script from ini", command=run_scriptFromIni)
button2.pack(pady=10)

# Start the GUI event loop
root.mainloop()
