import sys
import string
from enum import Enum
from openpyxl import load_workbook
from openpyxl import Workbook

alphabet_uppercase = string.ascii_uppercase

class Payment:
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

def generate_header(sheet, initial_balances):
    sheet["C2"].value = "Inkomster"
    sheet["D2"].value = "Konton"
    sheet["G2"].value = "Utgifter"

    sheet["C3"].value = "Rutin"
    sheet["D3"].value = "E-sparkonto"
    sheet["E3"].value = "Investerarkonto"
    sheet["F3"].value = "Kortkonto"
    sheet["G3"].value = "Rutin"
    sheet["H3"].value = "Mat"
    sheet["I3"].value = "Hushåll"
    sheet["J3"].value = "Hobby"
    sheet["K3"].value = "Nöje"
    sheet["L3"].value = "Resa"
    sheet["M3"].value = "Kläder"
    sheet["N3"].value = "Övrigt"
    sheet["O3"].value = "Kommentar"

    sheet["A4"].value = "Budgetering"
    sheet["D4"].value = 0
    sheet["E4"].value = 0
    sheet["G4"].value = 0
    sheet["H4"].value = 0
    sheet["I4"].value = 0
    sheet["J4"].value = 0
    sheet["K4"].value = 0
    sheet["L4"].value = 0
    sheet["M4"].value = 0
    sheet["N4"].value = 0

    sheet["A5"].value = "Ingående balans"
    sheet["D5"].value = initial_balances["Transaktioner Privatkonto"]
    sheet["E5"].value = initial_balances["Transaktioner Sparkonto"]
    sheet["F5"].value = initial_balances["Transaktioner Aktiekonto"]

def generate_footer(sheet, offset_y):
    sheet['B'+str(offset_y)]="Differens"
    for c in alphabet_uppercase[2:14:1] : 
        sheet[c+str(offset_y)]="=SUM("+c+"7:"+c+str(offset_y-2)+")"

    sheet['B'+str(offset_y+1)]="Utgående saldo"
    sheet['D'+str(offset_y+1)]="=SUM(D5,D"+str(offset_y)+")"
    sheet['E'+str(offset_y+1)]="=SUM(E5,E"+str(offset_y)+")"
    sheet['F'+str(offset_y+1)]="=SUM(F5,F"+str(offset_y)+")"

    outgoing_balances = {
        "Transaktioner Privatkonto": 0,
        "Transaktioner Sparkonto": 0,
        "Transaktioner Aktiekonto": 0
    }
    for i in range(6, offset_y):
        if sheet['D'+str(i)].value is not None:
            outgoing_balances["Transaktioner Privatkonto"] += sheet['D'+str(i)].value
        if sheet['E'+str(i)].value is not None:
            outgoing_balances["Transaktioner Sparkonto"] += sheet['E'+str(i)].value
        if sheet['F'+str(i)].value is not None:
            outgoing_balances["Transaktioner Aktiekonto"] += sheet['F'+str(i)].value


    sheet["D"+str(offset_y+3)].value = "Sparkonto"
    sheet["E"+str(offset_y+3)].value = "Aktiekonto"

    sheet["G"+str(offset_y+3)].value = "Rutin"
    sheet["H"+str(offset_y+3)].value = "Mat"
    sheet["I"+str(offset_y+3)].value = "Hushåll"
    sheet["J"+str(offset_y+3)].value = "Hobby"
    sheet["K"+str(offset_y+3)].value = "Nöje"
    sheet["L"+str(offset_y+3)].value = "Resa"
    sheet["M"+str(offset_y+3)].value = "Kläder"
    sheet["N"+str(offset_y+3)].value = "Övrigt"
    sheet["O"+str(offset_y+3)].value = "Kommentar"

    sheet["B"+str(offset_y+4)].value = "Budgetdifferens"
    sheet["B"+str(offset_y+5)].value = "Budgetdifferens, carryover"
    sheet["B"+str(offset_y+6)].value = "Budgetdifferens, manuella ändringar till nästa månad"
    sheet["B"+str(offset_y+7)].value = "Budget, ackumulerad nästa månad"

    sheet['D'+str(offset_y+4)]="=SUM(D4,D"+str(offset_y)+")"
    sheet['E'+str(offset_y+4)]="=SUM(E4,E"+str(offset_y)+")"

    for c in alphabet_uppercase[6:14:1] : 
        sheet[c+str(offset_y+4)]="=SUM("+c+"4,-"+c+str(offset_y)+")"
        sheet[c+str(offset_y+5)]=0
        sheet[c+str(offset_y+6)]=0
        sheet[c+str(offset_y+7)]="=SUM("+c+str(offset_y+4)+","+c+str(offset_y+4)+",-"+c+str(offset_y+6)+")"

    sheet["D"+str(offset_y+7)]="=SUM(D"+str(offset_y+4)+",D"+str(offset_y+4)+",-D"+str(offset_y+6)+")"
    sheet["E"+str(offset_y+7)]="=SUM(E"+str(offset_y+4)+",E"+str(offset_y+4)+",-E"+str(offset_y+6)+")"

    sheet["B"+str(offset_y+9)]="Kontroll"
    sheet["B"+str(offset_y+10)]="Utgående saldo"
    sheet["B"+str(offset_y+11)]="Ingående saldo"
    sheet["B"+str(offset_y+12)]="Differens"

    sheet["C"+str(offset_y+9)]="=SUM(C"+str(offset_y)+":N"+str(offset_y)+")"
    sheet["C"+str(offset_y+10)]="=SUM(D"+str(offset_y+1)+":N"+str(offset_y+1)+")"
    sheet["C"+str(offset_y+11)]="=SUM(D6:F6)"
    sheet["C"+str(offset_y+12)]="=C"+str(offset_y+10)+"-C"+str(offset_y+11)

    return outgoing_balances

def put_payments(sheet, payments):
    i=7
    for payment in payments:
        if payment.comment == None:
            print("WARNING: Found payment comment that was None")
            continue
        if not payment.comment in cost_type_translation_table:
            print("WARNING: Did not find comment: " + str(payment.comment) + " in cost rules, please add it. Skipping for now...")
            continue

        cost_type = cost_type_translation_table[payment.comment]
        cost_rule = cost_type_rules[cost_type]
        from_field = cost_rule.from_field
        to_field = cost_rule.to_field
        cost = payment.cost
        sheet[from_field + str(i)].value = -cost
        sheet[to_field + str(i)].value = cost
        sheet['B' + str(i)].value = payment.datetime
        sheet['O' + str(i)].value = payment.comment
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

# Read data file names
input_filenames = []
for i in range(1, len(sys.argv)-1):
    input_filenames.append(sys.argv[i])

output_filename = sys.argv[-1]
wb_output = load_workbook(output_filename)

payments_summed = {}
r = 7

    #TODO: Modify payment structure to take multiple to-accounts
    #TODO: Modify output file routine to read multiple to-accounts per payment
    #TODO: Fix all places where the code has broken
    #TODO: Match transactions from input fils to the output file, proritize the ones in the saved sheet over the input data

# Collect payments for all sheets in the ouput file
for sheet_name in wb_output.sheetnames:
    sheet_out = wb_output.get_sheet_by_name(sheet_name)
    while sheet_out['B'+str(r)].value != None:

        # Sum cost
        cost = 0
        for c in alphabet_uppercase[2:14:1] : 
            if sheet_out[c+str(r)].value != None:
                cost += max(sheet_out[c+str(r)].value, 0)

        # Read a specific cell
        payment = Payment(
            sheet_out['B'+str(r)].value,
            cost,
            sheet_out['O'+str(r)].value
        )

        month_identifier = str.rsplit(payment.datetime, '-', 1)[0]
        if not month_identifier in payments_summed:
            payments_summed[month_identifier] = []
        payments_summed[month_identifier].append(payment)
        r = r+1

initial_balances = {
    "Transaktioner Privatkonto": 0,
    "Transaktioner Sparkonto": 0,
    "Transaktioner Aktiekonto": 0
}

for filename in input_filenames:
    # Load xlsx workbook
    wb_input = load_workbook(filename)

    # Load active sheet or named sheet
    sheet_in = wb_input.active

    # Iterate over rows
    r = 9
    while sheet_in['A'+str(r)].value != None:

        # Read a specific cell
        payment = Payment(
            sheet_in['C'+str(r)].value,
            abs(sheet_in['G'+str(r)].value),
            sheet_in['E'+str(r)].value
        )

        month_identifier = str.rsplit(payment.datetime, '-', 1)[0]
        if not month_identifier in payments_summed:
            payments_summed[month_identifier] = []

        # Duplicate check
        duplicate_found = False
        for payment_2 in payments_summed[month_identifier]:
            if payment == payment_2:
                duplicate_found = True
                break

        if not duplicate_found:
            payments_summed[month_identifier].append(payment)
        r = r+1

    # Discern inital account balance
    initial_balance = 0
    i=9
    while sheet_in['H'+str(i)].value != None:
        initial_balance = sheet_in['H'+str(i)].value
        i+=1
    initial_balances[sheet_in["A1"].value] = initial_balance

for month in payments_summed.keys():
    for payment in payments_summed[month]:
        print("*****************")
        print(payment.comment)
        print(payment.cost)
        print(payment.datetime)

print(initial_balances)

# Sort the months
months_to_iterate = []
for month_identifier in payments_summed.keys():
    months_to_iterate.append(month_identifier)
months_to_iterate.sort();

# Save initial buget
saved_data = {}
saved_data["initial_budget"] = []
saved_data["carryover"] = {}
saved_data["manual_changes"] = {}

first_sheet = wb_output.get_sheet_by_name(months_to_iterate[0])
for c in alphabet_uppercase[6:14:1]: 
    saved_data["initial_budget"].append(first_sheet[c+"4"].value)

for month_identifier in months_to_iterate:
    sheet = wb_output.get_sheet_by_name(month_identifier)
    saved_data["carryover"][month_identifier] = []
    saved_data["manual_changes"][month_identifier] = []

    footer_y = 1
    while sheet['B'+str(footer_y)].value != "Budgetdifferens, carryover":
        footer_y+=1

    for c in alphabet_uppercase[6:14:1]: 
        saved_data["carryover"][month_identifier].append(sheet[c+str(footer_y)].value)
        saved_data["manual_changes"][month_identifier].append(sheet[c+str(footer_y+1)].value)

#Refer to data in other sheets:
#='2020-12'!K23

print(saved_data)

# Clear output workbook
wb_output = Workbook()

for month_identifier in months_to_iterate:
    wb_output.create_sheet(month_identifier)
    sheet_out = wb_output.get_sheet_by_name(month_identifier)
    generate_header(sheet_out, initial_balances)
    put_payments(sheet_out, payments_summed[month_identifier])
    initial_balances = generate_footer(sheet_out, 8+len(payments_summed[month_identifier]))

    # Load saved data
    i=0
    for c in alphabet_uppercase[6:14:1]: 
        sheet_out[c+"4"].value = saved_data["initial_budget"][i]
        i+=1

    footer_y = 1
    while sheet_out['B'+str(footer_y)].value != "Budgetdifferens, carryover":
        footer_y+=1

    i=0
    for c in alphabet_uppercase[6:14:1]: 
        sheet_out[c+str(footer_y)].value = saved_data["carryover"][month_identifier][i]
        sheet_out[c+str(footer_y+1)].value = saved_data["manual_changes"][month_identifier][i]
        i+=1

wb_output.remove_sheet(wb_output.active)

# Set budget reference to previous sheet
for month_index in range(0, len(months_to_iterate)-1):
    sheet_1 = wb_output.get_sheet_by_name(months_to_iterate[month_index])
    sheet_2 = wb_output.get_sheet_by_name(months_to_iterate[month_index+1])

    footer_y_1 = 1
    while sheet_1['B'+str(footer_y_1)].value != "Budget, ackumulerad nästa månad":
        footer_y_1+=1

    for c in alphabet_uppercase[3:14:1] : 
        sheet_2[c+str(4)].value="='"+months_to_iterate[month_index]+"'!"+c+str(footer_y_1)

wb_output.save(output_filename)
