import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

def is_number(n):
    try:
        float(n)
        return True
    except ValueError:
        return False

# Load spreadsheet
xl = pd.read_excel('testjob.xls', engine='xlrd')

# Initial processing
start_index = xl[xl.apply(lambda row: row.astype(str).str.contains('HARDWARE PARTS').any(), axis=1)].index[0]
end_index = xl[xl.apply(lambda row: row.astype(str).str.contains('Packsize Program').any(), axis=1)].index[0]
metal_parts_index = xl[xl.apply(lambda row: row.astype(str).str.contains('Metal Parts - Cut Length and Qty').any(), axis=1)].index[0]
accessory_buyout_index = xl[xl.apply(lambda row: row.astype(str).str.contains('ACCESSORY PARTS for BUYOUT').any(), axis=1)].index[0]

xl = xl.loc[start_index:end_index-1]
xl = xl.drop(list(range(metal_parts_index, accessory_buyout_index)))

xl.loc[xl[xl.columns[3]].isna() | (xl[xl.columns[3]] == '') | (xl[xl.columns[3]].astype(str).str.isspace()), xl.columns[3]] = xl.loc[xl[xl.columns[3]].isna() | (xl[xl.columns[3]] == '') | (xl[xl.columns[3]].astype(str).str.isspace()), xl.columns[2]]
xl.loc[xl[xl.columns[3]].isna() | (xl[xl.columns[3]] == '') | (xl[xl.columns[3]].astype(str).str.isspace()), xl.columns[2]] = ""
numeric_mask = xl[xl.columns[6]].apply(is_number) & (xl[xl.columns[5]].isna() | (xl[xl.columns[5]] == '') | (xl[xl.columns[5]].astype(str).str.isspace()))
xl.loc[numeric_mask, xl.columns[5]] = xl.loc[numeric_mask, xl.columns[6]]
xl.loc[numeric_mask, xl.columns[6]] = ""

# Section handling
sections = ['HARDWARE PARTS', 'Hinges & Mounting Plates', 'Legrabox & Antaro', 'Metabox, Tandem & Accuride', 'Blum Metal Parts',
            'Closet', 'Other', 'ACCESSORY PARTS for BUYOUT',
            'ADDITIONAL ACCESSORY PARTS', 'Decorative Hardware', 'RECESSED HARDWARE - Install Prior to Shipping',
            'Berenson INTEGRATED PULL PARTS']

xl.insert(0, 'Section', np.nan)

for section in sections:
    section_rows = xl[xl.iloc[:, 1] == section].index
    xl.loc[section_rows, 'Section'] = section
    xl.loc[section_rows, xl.columns[1]] = ""
    print(section)

xl = xl.dropna(how='all')

sections = ['HARDWARE PARTS', 'ACCESSORY PARTS for BUYOUT', 'ADDITIONAL ACCESSORY PARTS',
            'RECESSED HARDWARE - Install Prior to Shipping', 'Berenson INTEGRATED PULL PARTS']

section_indexes = {}
for section in sections:
    section_indexes[section] = xl[xl[xl.columns[0]] == section].index.tolist()

for i, (section, indexes) in enumerate(section_indexes.items()):
    if not indexes:
        continue
    start_index = indexes[0]
    if i+1 < len(sections):
        next_section = sections[i+1]
        if next_section in section_indexes:
            end_index = section_indexes[next_section][0]
        else:
            end_index = xl.index[-1]
    else:
        end_index = xl.index[-1]

    print(f"Processing section: {section}")

    if section == 'HARDWARE PARTS':
        pass
    elif section == 'ACCESSORY PARTS for BUYOUT':
        xl.loc[start_index + 1 : end_index - 1, xl.columns[6]] = xl.loc[start_index + 1 : end_index - 1, xl.columns[1]]
        xl.loc[start_index + 1 : end_index - 1, xl.columns[1]] = xl.loc[start_index + 1 : end_index - 1, xl.columns[2]]
        xl.loc[start_index + 1 : end_index - 1, xl.columns[2]] = ""
        xl.loc[start_index + 1: end_index - 1, xl.columns[5]] = xl.loc[start_index + 1: end_index - 1, xl.columns[4]]
        xl.loc[start_index + 1: end_index - 1, xl.columns[4]] = ""
    elif section == 'ADDITIONAL ACCESSORY PARTS':
        # Apply logic for "ADDITIONAL ACCESSORY PARTS" section here
        xl.loc[start_index + 1: end_index - 1, xl.columns[0]] = xl.loc[start_index + 1: end_index - 1, xl.columns[1]]

        xl.loc[start_index + 1: end_index - 1, xl.columns[6]] = xl.loc[start_index + 1: end_index - 1, xl.columns[3]]
        xl.loc[start_index + 1: end_index - 1, xl.columns[3]] = ''
        xl.loc[start_index + 1: end_index - 1, xl.columns[1]] = xl.loc[start_index + 1: end_index - 1, xl.columns[4]]
        xl.loc[start_index + 1: end_index - 1, xl.columns[4]] = ''
    elif section == 'RECESSED HARDWARE - Install Prior to Shipping':
        pass

xl = xl.dropna(how='all')

print(xl)

xl.replace('', np.nan, inplace=True)

rows_to_drop_partname = xl[xl.iloc[:, 1] == 'PART NAME'].index
xl.drop(rows_to_drop_partname, inplace=True)
rows_to_drop_qty = xl[xl.iloc[:, 0] == 'QTY'].index
xl.drop(rows_to_drop_qty, inplace=True)
rows_to_drop_qty2 = xl[xl.iloc[:, 3] == 'QTY'].index
xl.drop(rows_to_drop_qty2, inplace=True)
rows_to_drop_description = xl[xl.iloc[:, 2] == 'Description'].index
xl.drop(rows_to_drop_description, inplace=True)

words = ['BUY', 'PICKED', 'ID#', 'PART #', 'PART  #', 'QTY', 'ID  #']
rows_to_remove = xl[xl.iloc[:, 7:].isin(words).any(axis=1)].index
xl = xl.drop(rows_to_remove)
xl = xl.replace(words, '')

xl.dropna(how='all', inplace=True)
xl.dropna(axis=1, how='all', inplace=True)
xl = xl.reset_index(drop=True)

xl.to_excel('cleaned_data.xlsx', index=False)

# Create a new workbook and select the active worksheet
wb = Workbook()
ws = wb.active

# Transfer data from pandas DataFrame to the worksheet
for r in dataframe_to_rows(xl, index=False, header=True):
    ws.append(r)

# Adjust column widths
for column in ws.columns:
    max_length = 0
    column = [cell for cell in column]
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[get_column_letter(column[0].column)].width = adjusted_width

# Save the workbook
wb.save('cleaned_data.xlsx')
