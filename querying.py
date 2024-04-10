import pandas as pd
import tkinter as tk
from tkinter import filedialog


def select_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    return file_path


def export_to_excel(df):
    root = tk.Tk()
    root.withdraw()  # Hide the main window

    # Open a dialog to choose the folder
    folder_selected = filedialog.askdirectory()

    if folder_selected:  # If a folder was selected
        # Construct the full file path
        output_name = input("What would you like to name your exported file? ")
        file_path = folder_selected + f'/{output_name}.xlsx'
        
        # Export the DataFrame to Excel
        df.to_excel(file_path, index=False)
        print(f'DataFrame is exported successfully to {file_path}')
    else:
        print('No folder was selected.')


file = select_file()

cont = False
while cont == False:
    sheet_name = input('Do you have more than one sheet in this workbook? (Y/N) ')
    if sheet_name == "Y" or sheet_name== "y":
        sheet_name = input('What is the exact name of the sheet as it appears in the workbook? ')
        mb = pd.read_excel(file, dtype={'NDC11': str, 'ndc11': str}, sheet_name=sheet_name)
        cont = True

    elif sheet_name == "N" or sheet_name == "n":
        mb = pd.read_excel(file, dtype={'NDC11': str, 'ndc11': str})
        cont = True

    else:
        print('Not a valid entry...')

year = input('Which year would you like data for? ')

if year == '2023':
    data = pd.read_csv('2023_data.csv', dtype={'ndc': str, 'year': str, 'quarter': str})

elif year == '2022':
    data = pd.read_csv('2022_data.csv', dtype={'ndc': str, 'year': str, 'quarter': str})

elif year == '2021':
    data = pd.read_csv('2021_data.csv', dtype={'ndc': str, 'year': str, 'quarter': str})

elif year == '2020':
    data = pd.read_csv('2020_data.csv', dtype={'ndc': str, 'year': str, 'quarter': str})

elif year == '2019':
    data = pd.read_csv('2019_data.csv', dtype={'ndc': str, 'year': str, 'quarter': str})

if 'NDC11' in mb.columns:
    mb.rename(columns={'NDC11': 'ndc11'}, inplace=True)

elif 'NDC' in mb.columns:
    mb.rename(columns={'NDC': 'ndc11'}, inplace=True)

elif 'ndc' in mb.columns:
    mb.rename(columns={'ndc': 'ndc11'}, inplace=True)

if 'ndc' in data.columns:
    data.rename(columns={'ndc': 'ndc11'}, inplace=True)

df = pd.merge(data, mb[['ndc11', 'ProductNameLong', 'ProductName2']], on='ndc11', how='left')

pools = pd.read_csv('pools.csv')

df = pd.merge(df, pools[['state','Pool','MCO Included']], on='state', how='left')

ndcs = mb['ndc11'].unique()
df = df[df['ndc11'].isin(ndcs)].reset_index(drop=True)

df = df.drop(columns=['Unnamed: 0', 'labeler_code', 'product_code', 'package_size', 'suppression_used', 'medicaid_amount_reimbursed', 'non_medicaid_amount_reimbursed', 'product_name']).reset_index(drop=True)

quarters = input('Which quarters would you like? (e.g., 1,2) ')
quarters = quarters.split(',')

df = df[df['quarter'].isin(quarters)].reset_index(drop=True)

df.rename(columns={'utilization_type': 'ID', 'state': 'ST', 'ndc': 'ndc11', 'year': 'Year', 'quarter': 'Quarter', 'units_reimbursed': 'Units', 'number_of_prescriptions': 'Scripts', 'total_amount_reimbursed': 'Total Amount'}, inplace=True)

cols = list(df.columns)

cols.remove('ProductNameLong')
cols.remove('ProductName2')
cols.remove('Pool')
cols.remove('MCO Included')

cols.insert(2, 'Pool')
cols.insert(3, 'MCO Included')
cols.insert(7, 'ProductNameLong')
cols.insert(8, 'ProductName2')

df = df[cols]

print(df)

export_to_excel(df)
