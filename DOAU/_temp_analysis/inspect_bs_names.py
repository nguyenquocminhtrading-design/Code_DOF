import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=True)
ws_bs = wb['BS']

for r in range(1, ws_bs.max_row+1):
    lvl = ws_bs.cell(r, 16).value
    name = ws_bs.cell(r, 17).value
    k_val = ws_bs.cell(r, 11).value
    if name and any(k in str(name).lower() for k in ['receivable', 'payable', 'inventory']):
        print(f"BS R{r:2d} | Lvl {lvl} | {name:<35} | K: {k_val}")
