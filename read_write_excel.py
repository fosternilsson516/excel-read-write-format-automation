import pandas as pd
import json
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from user_input import get_user_input
from excel_formatting import apply_formatting


input_file, output_file, sheet_names = get_user_input()

# If any of the returned values are None, it means the user canceled one of the dialog boxes.
if None in [input_file, output_file, sheet_names]:
    print("Operation cancelled by user.")
    exit()


output_sheet_suffix = "_Script"

def process_sheet(sheet_name, writer):
    # Read the excel sheet into a DataFrame
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    # Convert the DataFrame to a JSON string
    json_str = df.to_json(orient='records', lines=True)

    data_list = [json.loads(line) for line in json_str.splitlines()]

    # Prepare data for new DataFrame
    output_data = []
    for record in data_list:
        output_data.append({
            "Test Case/Script Name": record["Test Case Name"],
            "Test Case/Script Description": record["Description"],
            "Application": record["Interface"],
            "Test Type": record["Artifact Type"],
            "Test Case Precondition": record["Test Data Requirements"],
            "Test Case Expected Result": record["Expected Result"],
        })

    # Create a DataFrame with the prepared data
    output_df = pd.DataFrame(output_data)

    output_df["Test Script Step: User Action"] = ""
    output_df["Test Script Step: Expected Result"] = ""
    
    # Write the new DataFrame to a respective sheet in Excel
    output_df.to_excel(writer, sheet_name=sheet_name + output_sheet_suffix, index=False)
    # Apply formatting using openpyxl
    worksheet = writer.sheets[sheet_name + output_sheet_suffix]
    worksheet = apply_formatting(worksheet)
    # Set header font to bold
    
   
            

# Use pd.ExcelWriter to write to the new Excel file
with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
    for sheet in sheet_names:
        print(f"Processing data from sheet: {sheet}")
        process_sheet(sheet, writer)