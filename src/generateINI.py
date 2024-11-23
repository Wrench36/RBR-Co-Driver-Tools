import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime

# Function to get the list of sheets from the Excel file
def get_sheet_names(excel_file):
    xls = pd.ExcelFile(excel_file)
    return xls.sheet_names

# Function to select the sheet using a tkinter Listbox, keep the window open
def select_sheet(sheet_names, excel_file):
    def on_select(event):
        global selected_sheet
        # Get the selected item from the Listbox
        widget = event.widget
        selection = widget.curselection()
        if selection:
            selected_sheet = widget.get(selection[0])
            process_sheet(excel_file, selected_sheet)
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
def process_sheet(excel_file, sheet_name):
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

    # Get current date and time
    current_datetime = datetime.now()
    current_datetime_str = current_datetime.strftime("%Y-%m-%d %H:%M:%S")

    # Save the output file with the sheet name
    output_file = f"{sheet_name}.ini"
    with open(output_file, 'w') as file:
        # Write the current date and time at the top of the .ini file
        file.write(f"; {sheet_name}\n; generateINI.py\n; By: Wrench\n; Date: {current_datetime_str}\n\n")

        # Iterate over the organization table, skipping the first row (header)
        for row_idx, row in organization_table.iloc[0:].iterrows():
            print(row_idx)
            for col_idx, string in row.items():
                if pd.notna(string):  # If the cell is not empty
                    # Retrieve the corresponding data for the string from the data_dict
                    string_data = data_dict.get(string, [])

                    # Write the PACENOTE section header
                    file.write(f"[PACENOTE::{string}]\n")

                    # Write the key-value pairs dynamically based on the headers from row 1
                    for header_idx, header_value in enumerate(header[1:], start=1):  # Skip first column (id)
                        value = df.iloc[row_idx, header_idx]  # Access corresponding value from row
                        if pd.notna(value):  # Only write non-empty values
                            if header_value != "Column":
                                file.write(f"{header_value.lower()}={value}\n")  # Use the header value as the key

                    # Write the 'column' based on the organization table position
                    file.write(f"column={col_idx - 9}\n")  # Adjust column index offset

                    # Write additional "Link" or other data from string_data
                    for key, value in zip(header[6:], string_data[6:]):
                        if pd.notna(value):
                            file.write(f"{key}={value}\n")

                    file.write("\n")  # Add a newline between sections

    print(f"File saved as {output_file}")


# Main function to handle file selection and sheet processing
def main():
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    

    # Ask the user to select the Excel file
    excel_file = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])

    if not excel_file:
        print("No file selected, exiting.")
        return

    # Get the list of sheet names
    sheet_names = get_sheet_names(excel_file)

    # Keep the sheet selection window open and process sheets until all are done
    select_sheet(sheet_names, excel_file)

# Run the main function
if __name__ == "__main__":
    main()
