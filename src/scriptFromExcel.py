import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Function to get the list of sheets from the Excel file
def get_sheet_names(excel_file):
    xls = pd.ExcelFile(excel_file)
    return xls.sheet_names

# Function to select the sheet using a tkinter Listbox, keep the window open
def select_sheet(sheet_names, excel_file,script_file):
    def on_select(event):
        global selected_sheet
        # Get the selected item from the Listbox
        widget = event.widget
        selection = widget.curselection()
        if selection:
            selected_sheet = widget.get(selection[0])
            process_sheet(excel_file, selected_sheet,script_file)
            # Update the listbox after processing
            sheet_names.remove(selected_sheet)
            if not sheet_names:  # If no more sheets are left, close
                root.quit()
            else:
                listbox.delete(0, tk.END)  # Clear current listbox items
                for sheet in sheet_names:  # Repopulate listbox with remaining sheets
                    listbox.insert(tk.END, sheet)

    root = tk.Tk()
    root.title("Select a sheet")

    listbox = tk.Listbox(root, height=len(sheet_names))
    listbox.pack()

    for sheet in sheet_names:
        listbox.insert(tk.END, sheet)

    # Bind the Listbox selection to the on_select function
    listbox.bind('<<ListboxSelect>>', on_select)

    root.mainloop()

# Function to process the Excel sheet and save as .ini
def process_sheet(excel_file, sheet_name, script_file):
    # Read the data from the sheet
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

    # Read the header from row 1 (columns A-G)
    header = df.iloc[0, :7].values.tolist()  # First row (A1 to G1) for definitions header

    # The definitions now start at A2, covering columns A to G
    definitions = df.iloc[1:, :7]  # Columns A-G, starting from row 2

    # The organization table now starts at J2, with up to 9 rows and extending from column J onwards
    organization_table = df.iloc[1:10, 9:]  # Up to 9 rows, starting from column J

    # Create a dictionary for the strings and their additional values from the definitions
    data_dict = definitions.set_index(0).T.to_dict('list')

    # Save the output file with the sheet name
    output_file = script_file
    with open(output_file, 'a') as file:
        for row_idx, row in definitions.iloc[0:].iterrows():
            file.write(f"{row[0]}\n")
        


# Main function to handle file selection and sheet processing
def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    

    # Ask the user to select the recording Script file
    script_file = filedialog.askopenfilename(title="Select A Script File", filetypes=[("Text Files", "*.txt")])
    
    # Ask the user to select the Excel file
    excel_file = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])

    if not excel_file:
        print("No file selected, exiting.")
        return

    # Get the list of sheet names
    sheet_names = get_sheet_names(excel_file)

    # Keep the sheet selection window open and process sheets until all are done
    select_sheet(sheet_names, excel_file,script_file)

# Run the main function
if __name__ == "__main__":
    main()
