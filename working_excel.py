from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook("MRI_update.xlsx")
ws = wb.active

my_wb = Workbook()
my_ms = my_wb.active

for row in range(1, 101):
    for col in range(1, 6):
        char = get_column_letter(col)
        if ws[char + str(row)] == "BL":
            my_ws.append(list(ws["A"+str(row)].value), ws["B"+str(row)].value, str(ws["C"+str(row)].value), str(ws["D"+str(row)].value))
my_wb.save("BL.xlsx")