from openpyxl import load_workbook

wb = load_workbook(filename='example-data.xlsx')

# List sheets available
sheets = wb.get_sheet_names()
print(sheets)

# Load active sheet or named sheet
sheet = wb.active
# sheet = wb['User Information']

# Read a specific cell
print(sheet['D8'].value)
