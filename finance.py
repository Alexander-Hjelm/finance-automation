import sys
from openpyxl import load_workbook

class CostEntry:

    def __init__(self, datetime, cost, comment):
        self.datetime = datetime
        self.cost = cost
        self.comment = comment

def generate_header(sheet):
    pass

def delete_footer(sheet):
    pass

def generate_footer(sheet, offset_y):
    pass

# TODO: Remove duplicate cost entries
# TODO: Generate output file
# TODO: Month/sheet management

# Read data file names
input_filenames = []
for i in range(1, len(sys.argv)-1):
    input_filenames.append(sys.argv[i])

output_filename = sys.argv[-1]
wb_output = load_workbook(output_filename)
sheet_out = wb_output.active

generate_header(sheet_out)
delete_footer(sheet_out)

cost_entries_summed = []
r = 7
while sheet_out['B'+str(r)].value != None:

    # Sum cost, TODO: make leaner, fix signs
    cost = 0
    if sheet_out['C'+str(r)].value != None:
        cost += max(sheet_out['C'+str(r)].value, 0)
    if sheet_out['D'+str(r)].value != None:
        cost += max(sheet_out['D'+str(r)].value, 0)
    if sheet_out['E'+str(r)].value != None:
        cost += max(sheet_out['E'+str(r)].value, 0)
    if sheet_out['F'+str(r)].value != None:
        cost += max(sheet_out['F'+str(r)].value, 0)
    if sheet_out['G'+str(r)].value != None:
        cost += max(sheet_out['G'+str(r)].value, 0)
    if sheet_out['H'+str(r)].value != None:
        cost += max(sheet_out['H'+str(r)].value, 0)
    if sheet_out['I'+str(r)].value != None:
        cost += max(sheet_out['I'+str(r)].value, 0)
    if sheet_out['J'+str(r)].value != None:
        cost += max(sheet_out['J'+str(r)].value, 0)
    if sheet_out['K'+str(r)].value != None:
        cost += max(sheet_out['K'+str(r)].value, 0)
    if sheet_out['L'+str(r)].value != None:
        cost += max(sheet_out['L'+str(r)].value, 0)
    if sheet_out['M'+str(r)].value != None:
        cost += max(sheet_out['M'+str(r)].value, 0)
    if sheet_out['N'+str(r)].value != None:
        cost += max(sheet_out['N'+str(r)].value, 0)

    # Read a specific cell
    cost_entry = CostEntry(
        sheet_out['B'+str(r)].value,
        cost,
        sheet_out['O'+str(r)].value
    )
    cost_entries_summed.append(cost_entry)
    r = r+1

for filename in input_filenames:
    # Load xlsx workbook
    wb_input = load_workbook(filename)

    # List sheets available
    #sheets = wb.get_sheet_names()
    #print(sheets)

    # Load active sheet or named sheet
    sheet_in = wb_input.active
    # sheet = wb['User Information']

    # Iterate over rows
    r = 9
    while sheet_in['A'+str(r)].value != None:

        # Read a specific cell
        cost_entry = CostEntry(
            sheet_in['C'+str(r)].value,
            sheet_in['G'+str(r)].value,
            sheet_in['E'+str(r)].value
        )
        cost_entries_summed.append(cost_entry)
        r = r+1

for cost_entry in cost_entries_summed:
    print("*****************")
    print(cost_entry.comment)
    print(cost_entry.cost)
    print(cost_entry.datetime)

generate_footer(sheet_out, 100)

