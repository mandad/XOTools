import smartsheet
import os
from template_smartsheet_secrets import *

# Set up the Smartsheet client
os.environ['SMARTSHEET_ACCESS_TOKEN'] = API_TOKEN
ss_client = smartsheet.Smartsheet(api_base=smartsheet.__gov_base__)
# Make sure we don't miss any error
ss_client.errors_as_exceptions(True)

def set_new_cruise(complement_sheet_id, summary_sheet_id):
    # Prompt user for the value to be added
    new_cruise = input("Enter the value to be added to 'Cruises Sailed': ")
    
    update_dropdown_based_on_checkbox(complement_sheet_id, new_cruise)
    update_cruise_name(summary_sheet_id, new_cruise)

def update_dropdown_based_on_checkbox(sheet_id, new_cruise_name):
    # Get the sheet
    sheet = ss_client.Sheets.get_sheet(sheet_id, level=2, include="objectValue")

    # Get the column Ids for "Cruises Sailed" and "Current Cruise"
    cruises_sailed_col_id = None
    current_cruise_col_id = None
    for column in sheet.columns:
        if column.title == "Cruises Sailed":
            cruises_sailed_col_id = column.id
        elif column.title == "Current Cruise":
            current_cruise_col_id = column.id

    # Ensure both columns were found
    if not cruises_sailed_col_id or not current_cruise_col_id:
        print("Columns not found.")
        return

    # Iterate over rows to check "Current Cruise" and update "Cruises Sailed"
    for row in sheet.rows:
        current_cruise_cell = row.get_column(current_cruise_col_id)
        cruises_sailed_cell = row.get_column(cruises_sailed_col_id)

        # Check if "Current Cruise" is checked
        if current_cruise_cell.value:
            # If the cell has existing values, split them using a semicolon (;) and check if the new value is already present
            existing_values = cruises_sailed_cell.object_value.values if cruises_sailed_cell.object_value else []
            
            # If the new value is not present, append it
            if new_cruise_name not in existing_values:
                existing_values.append(new_cruise_name)
                #updated_value = ", ".join(existing_values)
                #ss_client.Cells.update_cell(sheet_id, cruises_sailed_cell.id, updated_value)
                
                # Create a Cell object to update
                new_cell = smartsheet.models.Cell()
                new_cell.column_id = cruises_sailed_col_id
                #new_cell.object_value.object_type = MULTI_PICKLIST #8
                #new_cell.object_value.values = existing_values
                new_cell.object_value = {"objectType": "MULTI_PICKLIST", "values":existing_values}
                new_cell.strict = False
                
                # Update the row
                updated_row = smartsheet.models.Row()
                updated_row.id = row.id
                updated_row.cells.append(new_cell)
                ss_client.Sheets.update_rows(sheet_id, [updated_row])
    print("Cruise added to list for all sailing personnel.")
                

def update_cruise_name(target_sheet_id, new_cruise_name):
    # Get the target sheet
    target_sheet = ss_client.Sheets.get_sheet(target_sheet_id)
    
    # Get the column Ids for "Value" and "Key" in the target sheet
    target_value_col_id = None
    target_key_col_id = None
    for column in target_sheet.columns:
        if column.title == "Value":
            target_value_col_id = column.id
        elif column.title == "Key":
            target_key_col_id = column.id

    # Ensure both columns were found
    if not target_value_col_id or not target_key_col_id:
        print("Columns not found in target sheet.")
        return

    # Update the "Value" cell in the row where "Key" is "Current Cruise"
    for row in target_sheet.rows:
        key_cell = row.get_column(target_key_col_id)
        if key_cell.value == "Current Cruise":
            # Create a Cell object to update
            new_cell = smartsheet.models.Cell()
            new_cell.column_id = target_value_col_id
            new_cell.value = new_cruise_name
            new_cell.strict = False
            
            # Update the row
            updated_row = smartsheet.models.Row()
            updated_row.id = row.id
            updated_row.cells.append(new_cell)
            ss_client.Sheets.update_rows(target_sheet_id, [updated_row])
            print("Updated target sheet successfully.")
            return

    print("No row with 'Key' as 'Current Cruise' found in target sheet.")

# Run the function
set_new_cruise(SHEET_ID, SUMM_SHEET_ID)
