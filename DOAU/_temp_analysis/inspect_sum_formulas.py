import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=False)
ws_sum = wb['Summary']

for r in range(1, ws_sum.max_row+1):
    name = ws_sum.cell(r, 2).value
    f = ws_sum.cell(r, 4).value
    if name in ['Bonus', 'Salaries', 'Payroll Taxes']:
        print(f"Summary R{r:3d} | {name:<20} | {f}")
