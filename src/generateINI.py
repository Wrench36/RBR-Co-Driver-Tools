import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
import os

def get_sheet_names(excel_file):
    xls = pd.ExcelFile(excel_file)
    return xls.sheet_names

def select_sheet(sheet_names, excel_file):
    def on_select(event):
        global selected_sheet
        widget = event.widget
        selection = widget.curselection()
        if selection:
            selected_sheet = widget.get(selection[0])
            process_sheet(excel_file, selected_sheet)
            sheet_names.remove(selected_sheet)
            if not sheet_names:
                root.quit()
            else:
                listbox.delete(0, tk.END)
                for sheet in sheet_names:
                    listbox.insert(tk.END, sheet)

    root = tk.Tk()
    root.title("Select a sheet")

    listbox = tk.Listbox(root, height=len(sheet_names))
    listbox.pack()

    for sheet in sheet_names:
        listbox.insert(tk.END, sheet)

    listbox.bind('<<ListboxSelect>>', on_select)

    root.mainloop()

def process_sheet(excel_file, sheet_name):
    df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
    header = df.iloc[0, :7].values.tolist()  # First row (A1 to G1) for definitions header
    definitions = df.iloc[1:, :7]  # Columns A-G, starting from row 2
    # The organization table starts at J2, with up to 9 rows and extending from column J onwards
    organization_table = df.iloc[1:10, 9:]  # Up to 9 rows, starting from column J
    data_dict = definitions.set_index(0).T.to_dict('list')

    current_datetime = datetime.now()
    current_datetime_str = current_datetime.strftime("%Y-%m-%d %H:%M:%S")

    output_file = f"{sheet_name}.ini"
    with open(output_file, 'w') as file:
        file.write(f"; {sheet_name}\n; generateINI.py\n; By: Wrench\n; Date: {current_datetime_str}\n\n")

        for row_idx, row in organization_table.iloc[0:].iterrows():
            for col_idx, string in row.items():
                if pd.notna(string):  # If the cell is not empty
                    # Retrieve the corresponding data for the string from the data_dict
                    string_data = data_dict.get(string, [])
                    file.write(f"[PACENOTE::{string}]\n")
                    
                    for index, item in enumerate(string_data):
                        if pd.notna(item) and header[index+1] != "Column":
                            file.write(f"{header[index+1].lower()}={item}\n")

                    # Write the 'column' based on the organization table position
                    file.write(f"column={col_idx - 9}\n")  # Adjust column index offset

                    file.write("\n")

    print(f"File saved as {output_file}")


def main():
    root = tk.Tk()
    root.withdraw()
    parent_dir = os.path.dirname(os.path.abspath(__file__)) 
    docs_dir = os.path.abspath(os.path.join(parent_dir, "../"))
    excel_file = filedialog.askopenfilename(initialdir=docs_dir,title="Select Excel File", filetypes=[("Excel files", "*.xlsx")])

    if not excel_file:
        print("No file selected, exiting.")
        return
    
    sheet_names = get_sheet_names(excel_file)

    select_sheet(sheet_names, excel_file)

# Run the main function
if __name__ == "__main__":
    main()
