import sys
from enum import Enum
from openpyxl import load_workbook

class CostEntry:
    def __init__(self, datetime, cost, comment):
        self.datetime = datetime
        self.cost = cost
        self.comment = comment

    def __eq__(self, other):
        return self.datetime == other.datetime and self.cost == other.cost and self.comment == other.comment

class CostTypeRule:
    def __init__(self, from_field, to_field):
        self.from_field = from_field
        self.to_field = to_field

def generate_header(sheet):
    sheet["C2"].value = "Inkomster"
    sheet["D2"].value = "Konton"
    sheet["G2"].value = "Utgifter"
    #TODO: Implement

def clear_sheet(sheet):
    sheet_out.delete_cols(1, 1000)

def generate_footer(sheet, offset_y):
    sheet['C'+str(offset_y)]="=SUM(C7:C"+str(offset_y-2)+")"
    #TODO: Implement

def put_cost_entries(sheet, cost_entries):
    i=7
    for cost_entry in cost_entries:
        if cost_entry.comment == None:
            print("WARNING: Found cost entry comment that was None")
            continue
        if not cost_entry.comment in cost_type_translation_table:
            print("WARNING: Did not find comment: " + str(cost_entry.comment) + " in cost rules, please add it. Skipping for now...")
            continue

        cost_type = cost_type_translation_table[cost_entry.comment]
        cost_rule = cost_type_rules[cost_type]
        from_field = cost_rule.from_field
        to_field = cost_rule.to_field
        cost = cost_entry.cost
        sheet[from_field + str(i)].value = -cost
        sheet[to_field + str(i)].value = cost
        i+=1

# Config

class CostType(Enum):
    EXPENSE_CLOTHING=1
    EXPENSE_FOOD=2
    EXPENSE_FUN=3
    EXPENSE_GIFTS=4
    EXPENSE_HOBBY=5
    EXPENSE_HOUSEHOLD=6
    EXPENSE_MISC=7
    EXPENSE_ROUTINE=8
    EXPENSE_TRAVEL=9
    INCOME=10
    TRANSFER_SAVINGS_TO_CARD=11
    TRANSFER_SAVINGS_TO_STOCK=12

cost_type_rules = {
    CostType.EXPENSE_CLOTHING: CostTypeRule('F', 'M'),
    CostType.EXPENSE_FOOD: CostTypeRule('F', 'H'),
    CostType.EXPENSE_FUN: CostTypeRule('F', 'K'),
    CostType.EXPENSE_GIFTS: CostTypeRule('F', ''),
    CostType.EXPENSE_HOBBY: CostTypeRule('F', 'J'),
    CostType.EXPENSE_HOUSEHOLD: CostTypeRule('F', 'I'),
    CostType.EXPENSE_MISC: CostTypeRule('F', 'N'),
    CostType.EXPENSE_ROUTINE: CostTypeRule('F', 'G'),
    CostType.EXPENSE_TRAVEL: CostTypeRule('F', 'L'),
    CostType.INCOME: CostTypeRule('C', 'D'),
    CostType.TRANSFER_SAVINGS_TO_CARD: CostTypeRule('D', 'F'),
    CostType.TRANSFER_SAVINGS_TO_STOCK: CostTypeRule('D', 'E')
}

cost_type_translation_table = {
    "EMMAUS": CostType.EXPENSE_CLOTHING,
    "STADIUM DROTTNI": CostType.EXPENSE_CLOTHING,
    "AB STORSTOCKHOL": CostType.EXPENSE_FOOD,
    "BURGER KING ODE": CostType.EXPENSE_FOOD,
    "HEMKÖP DJURGÅRDS": CostType.EXPENSE_FOOD,
    "HEMKÖP SOLNA MAL": CostType.EXPENSE_FOOD,
    "ICA LAPPKARRSBER": CostType.EXPENSE_FOOD,
    "PROFESSORN RESTA": CostType.EXPENSE_FUN,
    "CLAS OHLSON": CostType.EXPENSE_HOUSEHOLD,
    "FOLKTANDVÅRD": CostType.EXPENSE_ROUTINE,
    "CLAS OHLSON 218": CostType.EXPENSE_HOUSEHOLD,
    "84319530719301": CostType.TRANSFER_SAVINGS_TO_CARD
}

# TODO: Month/sheet management

# Read data file names
input_filenames = []
for i in range(1, len(sys.argv)-1):
    input_filenames.append(sys.argv[i])

output_filename = sys.argv[-1]
wb_output = load_workbook(output_filename)
sheet_out = wb_output.active

cost_entries_summed = []
r = 7
while sheet_out['B'+str(r)].value != None:

    # Sum cost, TODO: make leaner
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
            abs(sheet_in['G'+str(r)].value),
            sheet_in['E'+str(r)].value
        )

        # Duplicate check
        duplicate_found = False
        for cost_entry_2 in cost_entries_summed:
            if cost_entry == cost_entry_2:
                duplicate_found = True
                break

        if not duplicate_found:
            cost_entries_summed.append(cost_entry)
        r = r+1

for cost_entry in cost_entries_summed:
    print("*****************")
    print(cost_entry.comment)
    print(cost_entry.cost)
    print(cost_entry.datetime)

clear_sheet(sheet_out)
generate_header(sheet_out)
put_cost_entries(sheet_out, cost_entries_summed)
generate_footer(sheet_out, 100)

wb_output.save(output_filename)
