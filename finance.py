import sys
from openpyxl import load_workbook

class CostEntry:

    def __init__(self, datetime, cost, comment):
        self.datetime = datetime
        self.cost = cost
        self.comment = comment


# Read data file names
input_filenames = []
for i in range(1, len(sys.argv)-1):
    input_filenames.append(sys.argv[i])

output_filename = sys.argv[-1]
wb_output = load_workbook(output_filename)
sheet_out = wb_output.active

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
    r = 1
    while sheet_in['A'+str(r)].value != None:

        # Read a specific cell
        cost_entry = CostEntry(
            sheet_in['A'+str(r)].value,
            sheet_in['B'+str(r)].value,
            sheet_in['C'+str(r)].value
        )

        r = r+1
