import os
import xlwings as xw
import pandas as pd
from glob import glob

# # Initialize an empty DataFrame to hold the stacked data
master_df = pd.DataFrame()
#
path = fr"P:\PACS\Finance\Budgets\2024 Q2\2024 Q2 PropCo Forecasts\1-2024 Q2 PropCo Forecast Template.xlsx"

# Read the "entity name" sheet into a DataFrame
entity_df = pd.read_excel(path, sheet_name='Entity_Name')

wb = xw.Book(path)
xw.App(visible=False)

for sheet in wb.sheets:
    print(sheet)
    # Read data range from Excel
    name = sheet.name

    other_rev = wb.sheets[sheet].range('O18:Z18').value
    pro_fees = wb.sheets[sheet].range('O106:Z106').value
    dep = wb.sheets[sheet].range('O118:Z118').value
    interest = wb.sheets[sheet].range('O122:Z122').value
    months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October',
              'November', 'December']

    # Define ledger amounts for each ledger
    ledger_other_rev = ['5990000'] * 12
    ledger_pro_fees = ['6900490'] * 12
    ledger_dep = ['7120000'] * 12
    ledger_interest = ['9000000'] * 12

    # Concatenate ledger amounts with corresponding months
    ledgers = pd.Series(ledger_other_rev + ledger_pro_fees + ledger_dep + ledger_interest)

    # Create a DataFrame with the combined data
    df = pd.DataFrame({
        'Sheet': name,
        'Ledger': ledgers,
        'Month': months * 4,  # Repeat the months for each set of amounts
        'Amount': other_rev + pro_fees + dep + interest  # Concatenate all amounts
    })

    if name in entity_df['Company'].values:
        entity_name = entity_df.loc[entity_df['Company'] == name, 'Reference ID'].iloc[0]
    else:
        entity_name = None

    # Append the entity name to the DataFrame
    df['Entity Name'] = entity_name

    # Append the DataFrame to the master DataFrame
    master_df = pd.concat([master_df, df], ignore_index=True)
    master_df.drop(columns=['Sheet'], inplace=True)
    master_df.dropna(subset=['Entity Name'], inplace=True)

wb.close()

# Save the master DataFrame to a new Excel file
master_df.to_excel("master_template.xlsx", index=False)
