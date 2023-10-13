# user_input.py
import pandas as pd
from tkinter import filedialog, simpledialog, Tk

def get_user_input():
    root = Tk()
    root.withdraw()

    input_file = filedialog.askopenfilename(title="Select the input Excel file")
    if not input_file:
        return None, None, None

    output_file = simpledialog.askstring("Output File Name", "Please enter the name for the output file:")
    if not output_file:
        return None, None, None
    
    if not output_file.endswith('.xlsx'):
        output_file += '.xlsx'

    all_sheet_names = pd.ExcelFile(input_file).sheet_names
    sheet_names_input = simpledialog.askstring("Sheet Names", f"Enter the names of the sheets you want to read, separated by commas:\n\nAvailable Sheets:\n{', '.join(all_sheet_names)}")
    if not sheet_names_input:
        return None, None, None

    sheet_names = [name.strip() for name in sheet_names_input.split(',')]
    
    return input_file, output_file, sheet_names