import pandas as pd
from pathlib import Path
from openpyxl import load_workbook

# Paths
folder_path = "C:/All Sales/Sales/JM"
file_path = "C:/All Sales/Sales - 2-12-2024/Daily Sales Report 02-12-24.xlsx"
output_file = "C:/All Sales/Processed Sales Data.xlsx"

# Extract the folder name
folder_name = Path(folder_path).name
print("Folder Name:", folder_name)

# Define the mapping dictionary
sheet_name_mapping = {
    "KC": "KC Product Sales",
    "NW": "North West Product Sales",
    "JM": "J M LIMITED",
    "SP": "Spetra Product Sales"
}

# Look up the sheet name using the folder name
matched_sheet_name = sheet_name_mapping.get(folder_name)
if matched_sheet_name:
    # Load the Excel sheet into pandas with engine specified
    data = pd.read_excel(file_path, sheet_name=matched_sheet_name, engine='openpyxl')

    # Define country sections and start/end markers
    country_sections = ["J M LIMITED (UK)", "J M LIMITED (GERMANY)", "J M LIMITED (FRANCE)", "J M LIMITED (ITALY)"]
    country_data = {}

    # Extract country-specific data
    for i, section in enumerate(country_sections):
        print(f"Processing section: {section}")

        # Strip any leading or trailing spaces from the column with country names
        data.iloc[:, 1] = data.iloc[:, 1].str.strip()

        # Find the row index where the country section starts
        section_indices = data[data.iloc[:, 1] == section].index

        if len(section_indices) == 0:
            print(f"Section '{section}' not found in the data.")
            continue

        start_row = section_indices[0] + 1
        print(f"Start row: {start_row}")
        # Determine the end row (next section or end of file)
        if i < len(country_sections) - 1:
            next_section_indices = data[data.iloc[:, 1] == country_sections[i + 1]].index
            end_row = next_section_indices[0] if len(next_section_indices) > 0 else len(data)
        else:
            end_row = len(data)

        print(f"End row: {end_row}")
        print(matched_sheet_name)
        workbook = load_workbook(file_path)
        sheet = workbook[matched_sheet_name]  # Replace with the correct sheet name if needed

        # Specify the row and column to update. Row and column are 1-indexed in openpyxl.
        # Example: updating cell at row 3, column 2 (B3)
        sheet.cell(row=5, column=6).value = '10'

        workbook.save(file_path)

        print("Excel file updated successfully.")
        # # Extract the country-specific data
        # country_df = data.iloc[start_row-1:end_row].reset_index(drop=True)

        # # print(country_df)
        # # Add column headers based on the first row
        # country_df.columns = country_df.iloc[0]
        # country_df = country_df[1:]  # Remove the header row from the data

        # # Add data to the "NO. OF SOLD" column
        # if "NO. OF SOLD" in country_df.columns:
        #     country_df["NO. OF SOLD"] = 10  # Example: Adding '10' to each entry
        # else:
        #     print(f"Column 'NO. OF SOLD' not found in section {section}")

        # # Store in dictionary
        # country_data[section] = country_df
        # print(country_data)

    # # Load the workbook to update the same Excel sheet
    # with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
    #     for country, df in country_data.items():
    #         # Write each DataFrame back to the corresponding sheet
    #         df.to_excel(writer, sheet_name=matched_sheet_name, index=False, header=True)
        
    print(f"Processed data saved to the same file: {file_path}")

else:
    print("No matching sheet name found.")












# # Use glob to find all Excel files (with .xls and .xlsx extensions)
# excel_files = glob.glob(f"{folder_path}/*.csv")
# df=[]
# for data in excel_files:
#    df.append(pd.read_csv(data))

# # Step 2: Initialize result_list as an empty list
# result_list = []
# # Iterate over the rows in the sheet (assuming headers in the first row)
# for row in sheet.iter_rows(min_row=2, values_only=True):
#     code = row[0]  # Column A (code)
#     name = row[1]  # Column B (name)
#     value = row[2]  # Column C (value)

#  # Check if any of the values in the row are null
#     if code is None or name is None or value is None:
#         # print(f"Row contains null values: {row}")
#         continue  # Skip processing this row if you want to ignore it

#     # Append data in the desired format
#     result_list.append({
#         "code": code,
#          "name": name,
#          "value": 0
#     })

# # Print the result list to see the data
# # print(result_list)

# # Step 4: Compare result_list codes with df2 codes and perform operations
# for dff in df[0:]:
#     for _, row2 in dff.iterrows():
#         code_found = False

#         # Check if code from df2 exists in result_list
#         for item in result_list:
#             if item["code"] == row2['(Child) ASIN']:
#                 # If code matches, perform addition
#                 # if item["code"]=="B00128WK4I":
#                 #     print(row2['Units ordered'])
#                 item["value"] += row2['Units ordered']
#                 # print('match found', row2['(Child) ASIN'],item["value"])
#                 code_found = True
#                 break

#         # If code not found, add the new code and value to result_list
#         # if not code_found:
#         #     result_list.append({"code": row2['(Child) ASIN'], "value": row2['Units ordered']})


# # print(result_list)


# def get_value_by_code(code):
#     # Iterate over the result_list and check for the code
#     for item in result_list:
#         if item["code"] == code:
#             return item["value"]
#     return None  # Return None if code is not found

# for code in result_list:
#     # Define the value to search for
#     search_value = code["code"]
#     # Initialize a variable to store the cell address
#     cell_address = None


#     # Iterate through the rows and columns to find the value
#     for row in sheet.iter_rows():  # Iterate over all rows
#         for cell in row:  # Iterate over each cell in the row
#             if cell.value == search_value:  # Match the target value
#                 cell_address = cell.coordinate  # Get the cell address
#                 break  # Stop the inner loop if found
#         if cell_address:
#             break  # Stop the outer loop if found

#     # Output the result
#     if cell_address:
#         # print(f"The value '{search_value}' is located at: {cell_address}")
#         # Extract row number from the original address
#         row_number = int(cell_address[1:])  # Extract everything after the first character (row)
#         new_column = "C"  # Specify the new column

#         # Construct the new cell address
#         new_cell_address = f"{new_column}{row_number}"

#         # Copy value from original cell to new cell
#         sheet[new_cell_address] = get_value_by_code(search_value)
#         workbook.save(file_path)
#     else:
#         print(f"The value '{search_value}' was not found in the sheet.")


# #change file name like master inventory.xlsx to SP master inventory 2044-06-15.xlsx
# if os.path.exists(file_path):
#     # Rename the file
#     today = datetime.today()
#     today_str = today.strftime("%Y-%m-%d")

#     #get folder name
#     folder_name = Path(folder_path).name

#     #rename file name
#     new_file_path = f"C:/Sales/{folder_name} Master Inventory {today_str}.xlsx"
#     os.rename(file_path,new_file_path)
#     # print(f"File renamed from '{file_path}")
# else:
#     print(f"The file '{file_path}' does not exist.")

# # ..................................MIELLE OLD AND NEW CLASSIFICATION........................

# # Define variables to store MIELLE OLD and NEW values
# regions = {"AUS": {"old": 0, "new": 0}, 
#            "FRA": {"old": 0, "new": 0}, 
#            "GER": {"old": 0, "new": 0}, 
#            "IT": {"old": 0, "new": 0}, 
#            "UK": {"old": 0, "new": 0}}

# # MIELLE OLD and NEW codes
# MIELLE_OLD_CODE = {"B07N7PK9QK", "B09LN2XKKQ", "B07QLHFSFP"}
# MIELLE_NEW_CODE = {"B0DHVLFR2V"}

# # Map file names to region keys
# region_mapping = {
#     "KC AUS": "AUS",
#     "KC FRA": "FRA",
#     "KC GER": "GER",
#     "KC SPA": "GER",
#     "KC SWE": "GER",
#     "KC BEL": "GER",
#     "KC NL": "GER",
#     "KC POL": "GER",
#     "KC IT": "IT",
#     "KC UK": "UK",
# }

# # Process each file and update region counts
# for file, dff in zip(excel_files, df):
#     file_name = os.path.basename(os.path.splitext(file)[0])
#     region_key = region_mapping.get(file_name)
    
#     if region_key:
#         for _, row in dff.iterrows():
#             if row["(Child) ASIN"] in MIELLE_OLD_CODE:
#                 regions[region_key]["old"] += row["Units ordered"]
#             elif row["(Child) ASIN"] in MIELLE_NEW_CODE:
#                 regions[region_key]["new"] += row["Units ordered"]

        
# for reg in regions:
#     print(f"{reg} - OLD: {regions[reg]['old']}, NEW: {regions[reg]['new']}")               