from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook("mouse.xlsx")
ws = wb["TBI"]


def get_row_data(row):
    # define a function to get data from a certain row here, so that it is easier to append data in the for loop.
    row_data = [ws["A" + str(row)].value, ws["B" + str(row)].value, ws["C" + str(row)].value, ws["D" + str(row)].value,
                ws["E" + str(row)].value, ws["F" + str(row)].value, ws["G" + str(row)].value, ws["H" + str(row)].value,
                ws["I" + str(row)].value, ws["J" + str(row)].value, ws["K" + str(row)].value, ws["L" + str(row)].value,
                ws["M" + str(row)].value, ws["N" + str(row)].value, ws["O" + str(row)].value]
    return row_data


wb.create_sheet("temp")
ws_1 = wb["temp"]

for row in range(1, ws.max_row + 1):
    for col in range(1, 16):
        char = get_column_letter(col)
        if ws[char + str(row)].value == "WT":
            ws_1.append(get_row_data(row))

wb.save("mouse.xlsx")

wb = load_workbook("mouse.xlsx")
ws = wb["temp"]

wb.create_sheet("temp_sham")
wb.create_sheet("temp_TBI")
ws_2 = wb["temp_sham"]
ws_3 = wb["temp_TBI"]

for row in range(1, ws.max_row + 1):
    for col in range(1, 16):
        char = get_column_letter(col)
        if ws[char + str(row)].value == "s-sham":
            ws_2.append(get_row_data(row))

for row in range(1, ws.max_row + 1):
    for col in range(1, 16):
        char = get_column_letter(col)
        if ws[char + str(row)].value == "s-TBI":
            ws_3.append(get_row_data(row))

del wb["temp"]
wb.save("mouse.xlsx")

wb = load_workbook("mouse.xlsx")
ws = wb["temp_sham"]

wb.create_sheet("sorted_WT")
ws_4 = wb["sorted_WT"]
for row in range(1, ws.max_row + 1):
    for col in range(1, 16):
        char = get_column_letter(col)
        if ws[char + str(row)].value == "F":
            ws_4.append(get_row_data(row))

for row in range(1, ws.max_row + 1):
    for col in range(1, 16):
        char = get_column_letter(col)
        if ws[char + str(row)].value == "M":
            ws_4.append(get_row_data(row))

del wb["temp_sham"]
wb.save("mouse.xlsx")

wb = load_workbook("mouse.xlsx")
ws = wb["temp_TBI"]
ws_4 = wb["sorted_WT"]

for row in range(1, ws.max_row + 1):
    for col in range(1, 16):
        char = get_column_letter(col)
        if ws[char + str(row)].value == "F":
            ws_4.append(get_row_data(row))

for row in range(1, ws.max_row + 1):
    for col in range(1, 16):
        char = get_column_letter(col)
        if ws[char + str(row)].value == "M":
            ws_4.append(get_row_data(row))

del wb["temp_TBI"]
wb.save("mouse.xlsx")