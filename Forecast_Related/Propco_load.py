import os
import xlwings as xw
import pandas as pd
from glob import glob

print("*"*12)
print("Did you update the Entity list? ")
print("*"*12)
print("Did you update the file path? ")

print("*"*12)
year = int(input("Year?: "))

# File path
path = fr"P:\PACS\Finance\Budgets\2024 Q3\2024 Q3 PropCo Forecasts\1-2024 Q3 PropCo Forecast Template v2.xlsx"
new_path = directory_path = os.path.dirname(path)

# EIB file
eib_temp = fr"Virtual Machine Upload WD_upload_budget_main.xlsx"

#######################################################################
#######################################################################
# Initialize an empty DataFrame to hold the stacked data
master_df = pd.DataFrame()

# while True:
#     # File path
#     path = input("Enter the file path: ")

if '\\PACS' in path or '\PACS' in path:
    print(r"Remove \PACS or \\PACS from the file path")


# Check if the file exists
if os.path.exists(path):
    try:
        # Read the "entity name" sheet into a DataFrame
        entity_df = pd.read_excel(path, sheet_name='Entity_Name')
    except Exception as e:
        print(f"Failed to read the Excel file: {e}")
else:
    print("File path does not exist. Please check and try again.")


wb = xw.Book(path)
xw.App(visible=False)
app = xw.apps.active

# Extract data
for sheet in wb.sheets:
    print(sheet)
    # Read data range from Excel
    name = sheet.name

    other_rev = wb.sheets[sheet].range('O18:Z18').value
    pro_fees = wb.sheets[sheet].range('O106:Z106').value
    dep = wb.sheets[sheet].range('O118:Z118').value
    tax = wb.sheets[sheet].range('O120:Z120').value
    insurance = wb.sheets[sheet].range('O121:Z121').value
    interest = wb.sheets[sheet].range('O122:Z122').value
    non_op_rev = wb.sheets[sheet].range('O142:Z142').value
    months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct',
              'Nov', 'Dec']


    # Define ledger amounts for each ledger
    ledger_other_rev = ['5990000'] * 12
    ledger_pro_fees = ['6900490'] * 12 # _ADM
    ledger_dep = ['7120000'] * 12 # _PROP
    ledger_tax = ['7300000'] * 12 #_PROP
    ledger_ins = ['7400000'] * 12 #_PROP
    ledger_interest = ['7500000'] * 12 #_NONOP
    ledger_non_op = ['9000000'] * 12 #_NONOP

    # Concatenate ledger amounts with corresponding months
    ledgers = pd.Series(ledger_other_rev + ledger_pro_fees + ledger_dep + ledger_tax + ledger_ins + ledger_interest + ledger_non_op)

    # Create a DataFrame with the combined data
    df = pd.DataFrame({
        'Sheet': name,
        'Ledger': ledgers,
        'Month': months * 7,  # Repeat the months for each set of amounts
        'Amount': other_rev + pro_fees + dep + tax + insurance + interest + non_op_rev # Concatenate all amounts
    })


    if name in entity_df['Company'].values:
        entity_name = entity_df.loc[entity_df['Company'] == name, 'Reference ID'].iloc[0]
        prop_id = entity_df.loc[entity_df['Company'] == name, 'ID'].iloc[0]
    else:
        entity_name = None
        prop_id = None

    df['Entity Name'] = entity_name
    df['Prop_id'] = prop_id

    # Append the DataFrame to the master DataFrame
    master_df = pd.concat([master_df, df], ignore_index=True)

master_df.drop(columns=['Sheet'], inplace=True)
master_df.dropna(subset=['Entity Name'], inplace=True)

ledger_mapping = {
    '6900490': '_ADM',
    '7120000': '_PROP',
    '7300000': '_PROP',
    '7400000': '_PROP',
    '7500000': '_NONOP',
    '9000000': '_NONOP'
}

# Section for adding and revising columns
master_df['Cost_Center'] = master_df['Ledger'].map(ledger_mapping)
master_df['Cost_Center'] = master_df['Prop_id'] + master_df['Cost_Center']
master_df['Amount'] = pd.to_numeric(master_df['Amount'], errors='coerce')
master_df['Amount'] = master_df['Amount'].round(2)
master_df.drop(columns=['Prop_id'], inplace=True)
master_df['Cost_Center'].fillna("", inplace=True)
master_df['Year'] = year
master_df['Account Set'] = 'Standard_Child'
master_df['Debit'] = 0
master_df['Ledger'] = pd.to_numeric(master_df['Ledger'], errors='coerce')
master_df['Amount'] = pd.to_numeric(master_df['Amount'], errors='coerce')
master_df['Debit'] = master_df['Amount'].where(master_df['Ledger'] > 5990000, master_df['Debit'])
master_df.loc[master_df['Ledger'] > 5990000, 'Amount'] = 0.0
master_df['Index0'] = ""
master_df['Index1'] = 1
last_row_index = len(master_df)
master_df['Index2'] = range(1, last_row_index + 1)
master_df['Index3'] = ""
master_df['Index6'] = ""
# Convert values
master_df['Debit'] = pd.to_numeric(master_df['Debit'], errors='coerce').fillna(0)
master_df['Credit'] = pd.to_numeric(master_df['Credit'], errors='coerce').fillna(0)
master_df['Debit'] = master_df['Debit'].astype(int)
master_df['Credit'] = master_df['Credit'].astype(int)

# Prep for xlsx eib
order_col = ['Index0', 'Index1', 'Index2', 'Index3', 'Entity Name', 'Year', 'Month', 'Ledger', 'Account Set', 'Debit', 'Amount', 'Index6', 'Cost_Center']
master_df = master_df[order_col]
wb.close()
master_df.to_excel("propco_data.xlsx", index=False)

wb_m = xw.Book("propco_data.xlsx")

values = wb_m.sheets[0].range('A2:M9999').value

try:
    with open(eib_temp, 'r') as file:
        # If the file exists, proceed with your operations
        pass
except FileNotFoundError:
    # If the file is not found, prompt the user to enter the file path
    eib_temp = input("Enter the EIB template file path: ")

############################################
############################################
wb_e = xw.Book(eib_temp)
sheet = wb_e.sheets['Budget Lines Data']
sheet.range('A6').value = values  # Assuming you want to start from cell A6
wb_e.save(fr"{new_path}\Propco_{year}_eib.xlsx")

wb_e.close()
