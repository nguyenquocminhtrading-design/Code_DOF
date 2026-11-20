import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=False)
ws_007 = wb['007-DOAU data']

print("=== PARSING 007 DATA STRUCTURE ===")
# Let's extract all rows from row 4 to 244
rows_007 = []
for r in range(4, ws_007.max_row + 1):
    lvl = None
    name = None
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v is not None and str(v).strip() != '':
            lvl = c
            name = str(v).strip()
            break
    if name is not None:
        val_g = ws_007.cell(r, 7).value
        rows_007.append((r, lvl, name, str(val_g) if val_g is not None else ''))

with open('007_parsed_items.txt', 'w', encoding='utf-8') as f:
    for r, lvl, name, formula in rows_007:
        f.write(f"R{r:3d} | Lvl {lvl} | {name:<45} | {formula}\n")

print(f"Total extracted items: {len(rows_007)}")
