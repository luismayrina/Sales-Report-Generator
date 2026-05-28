import openpyxl

wb = openpyxl.load_workbook("Sample Reports.xlsx")
ws = wb["Offtake Report Summary"]

for i in range(1, 35):
    v = ws.cell(i, 1).value
    if v is not None:
        print(f"Row {i}: '{str(v).strip()}'")
