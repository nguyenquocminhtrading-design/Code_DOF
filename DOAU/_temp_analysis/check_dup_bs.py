import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=True)
ws_bs = wb['BS']

counts = {}
for r in range(3, ws_bs.max_row+1):
    name = ws_bs.cell(r, 17).value
    k = ws_bs.cell(r, 11).value
    if name:
        counts[name] = counts.get(name, []) + [(r, ws_bs.cell(r, 16).value, k)]

for name, rows in counts.items():
    if len(rows) > 1:
        print(f"Name '{name}' appears {len(rows)} times in BS:")
        for r, lvl, k in rows:
            print(f"  Row {r}, Lvl {lvl}, K={k}")
