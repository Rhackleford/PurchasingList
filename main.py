import pandas as pd
import numpy as np

# Load spreadsheet
xl = pd.read_excel('testjob.xls', engine='xlrd')

# Find the index of the row containing "HARDWARE PARTS"
start_index = xl[xl.apply(lambda row: row.astype(str).str.contains('HARDWARE PARTS').any(), axis=1)].index[0]

# Find the index of the row containing "Packsize Program"
end_index = xl[xl.apply(lambda row: row.astype(str).str.contains('Packsize Program').any(), axis=1)].index[0]

# Remove all rows above the start row and below the end row
xl = xl.loc[start_index:end_index-1]

# Define the list of sections
sections = ['HARDWARE PARTS', 'ACCESSORY PARTS for BUYOUT',
            'ADDITIONAL ACCESSORY PARTS', 'RECESSED HARDWARE - Install Prior to Shipping',
            'Berenson INTEGRATED PULL PARTS']

# Add a new column "Section" initialized with NaN
xl['Section'] = np.nan

# Assign the section names to the 'Section' column and set the original cells to blank
for sec in sections:
    sec_indices = xl[xl.apply(lambda row: row.astype(str).str.contains(sec).any(), axis=1)].index
    xl.loc[sec_indices, 'Section'] = sec
    xl.loc[sec_indices, xl.columns[0]] = ""  # Clear only the first column of the identified rows

# Find the index of the row containing "Metal Parts - Cut Length and Qty"
metal_parts_index = xl[xl.apply(lambda row: row.astype(str).str.contains('Metal Parts - Cut Length and Qty').any(), axis=1)].index[0]

# Find the index of the row containing "ACCESSORY PARTS for BUYOUT"
accessory_buyout_index = xl[xl.apply(lambda row: row.astype(str).str.contains('ACCESSORY PARTS for BUYOUT').any(), axis=1)].index[0]

# Remove all rows from "Metal Parts - Cut Length and Qty" to "ACCESSORY PARTS for BUYOUT"
xl = xl.drop(list(range(metal_parts_index, accessory_buyout_index)))

# Get the index of the start and end of "ACCESSORY PARTS for BUYOUT" section
accessory_start_index = xl[xl['Section'] == 'ACCESSORY PARTS for BUYOUT'].index[0]
accessory_end_index = xl[xl['Section'] == 'ADDITIONAL ACCESSORY PARTS'].index[0]

# # Re-arrange the cells for "ACCESSORY PARTS for BUYOUT" section and create a new section "GLASS PARTS"
for idx in range(accessory_start_index + 1, accessory_end_index):

        xl.loc[idx, 'Temp1'] = xl.loc[idx, xl.columns[0]]
        xl.loc[idx, 'Temp2'] = xl.loc[idx, xl.columns[1]]
        xl.loc[idx, 'Temp3'] = xl.loc[idx, xl.columns[4]]

        xl.loc[idx, xl.columns[0]] = xl.loc[idx, 'Temp2']
        xl.loc[idx, xl.columns[1]] = np.nan
        xl.loc[idx, xl.columns[5]] = xl.loc[idx, 'Temp1']

        xl.loc[idx, xl.columns[4]] = np.nan
        xl.loc[idx, xl.columns[3]] = xl.loc[idx, 'Temp3']


# Get the index of the start and end of "ADDITIONAL ACCESSORY PARTS" section
additional_start_index = xl[xl['Section'] == 'ADDITIONAL ACCESSORY PARTS'].index[0]
additional_end_index = xl[xl['Section'] == 'RECESSED HARDWARE - Install Prior to Shipping'].index[0]

# Re-arrange the cells for "ADDITIONAL ACCESSORY PARTS" section
for idx in range(additional_start_index + 1, additional_end_index):
    xl.loc[idx, 'Temp1'] = xl.loc[idx, xl.columns[0]]
    xl.loc[idx, 'Temp2'] = xl.loc[idx, xl.columns[2]]
    xl.loc[idx, 'Temp3'] = xl.loc[idx, xl.columns[3]]
    xl.loc[idx, 'Temp4'] = xl.loc[idx, xl.columns[7]]

    xl.loc[idx, xl.columns[0]] = xl.loc[idx, 'Temp3']
    xl.loc[idx, xl.columns[3]] = np.nan
    xl.loc[idx, xl.columns[9]] = xl.loc[idx, 'Temp1']

    xl.loc[idx, xl.columns[2]] = np.nan
    xl.loc[idx, xl.columns[5]] = xl.loc[idx, 'Temp2']

    xl.loc[idx, xl.columns[7]] = np.nan
    xl.loc[idx, xl.columns[6]] = xl.loc[idx, 'Temp4']

# # Get the index of the start and end of "GLASS PARTS" section
# glass_start_index = xl[xl['Section'] == 'GLASS PARTS'].index[0]
# glass_end_index = xl[xl['Section'] == 'ADDITIONAL ACCESSORY PARTS'].index[0]


# # Re-arrange the cells for "GLASS PARTS" section
# for idx in range(glass_start_index + 1, glass_end_index):
#     xl.loc[idx, xl.columns[10]] = xl.loc[idx, xl.columns[8]]
#     xl.loc[idx, xl.columns[11]] = xl.loc[idx, xl.columns[9]]
#
#     xl.loc[idx, xl.columns[8]] = np.nan
#     xl.loc[idx, xl.columns[9]] = np.nan


# def final_cleanup(df):
#     # Remove rows containing 'PART' in column 0, 'BUY' in column 7 and 'QTY' in column 5
#     df = df[~((df[df.columns[0]].str.contains('PART', na=False)) & (df[df.columns[7]].str.contains('BUY', na=False)) & (
#         df[df.columns[5]].str.contains('QTY', na=False)))]
#
#     # Also remove rows containing 'PART#', 'PART #', or 'PART  #' in column 0
#     df = df[~df[df.columns[0]].str.contains('PART\s*#', regex=True, na=False)]
#
#     # Drop the specified columns
#     df = df.drop(df.columns[[2, 4, 6, 7, 8]], axis=1)
#
#     # Move the contents of the last column to column 7
#     df[df.columns[7]] = df[df.columns[-1]]  # -1 refers to the last column
#     df = df.drop(df.columns[-1], axis=1)
#
#     # If there is data in column 0 and no data in column 3 and column 7, move the contents of column 0 to column 7
#     mask = df[df.columns[0]].notna() & df[df.columns[3]].isna() & df[df.columns[7]].isna()
#     df.loc[mask, df.columns[7]] = df.loc[mask, df.columns[0]]
#
#     # After moving the contents from column 0 to 7, make column 0 blank where the content has been moved
#     df.loc[mask, df.columns[0]] = np.nan
#
#     # If column 0 is empty and column 4 is not empty, move the contents of column 4 to column 7
#     mask_2 = df[df.columns[0]].isna() & df[df.columns[4]].notna()
#     df.loc[mask_2, df.columns[7]] = df.loc[mask_2, df.columns[4]]
#
#     # After moving the contents from column 4 to 7, make column 4 blank where the content has been moved
#     df.loc[mask_2, df.columns[4]] = np.nan
#
#     # Replace cells containing 'BUY' or 'PICKED' with NaN
#     df.replace('BUY', np.nan, inplace=True)
#     df.replace('PICKED', np.nan, inplace=True)
#
#     return df


# Drop the temporary columns
xl = xl.drop(['Temp1', 'Temp2', 'Temp3', 'Temp4'], axis=1)

# Remove all blank rows
xl = xl.dropna(how='all')

# Print the DataFrame to console
print(xl)

# Before saving the Excel file, call the final_cleanup function
# xl = final_cleanup(xl)


# Reset the index
xl = xl.reset_index(drop=True)



# Save to new .xlsx file (in Excel 2007+ format)
xl.to_excel('cleaned_data.xlsx', index=False)

