import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=True)
ws_007 = wb['007-DOAU data']
ws_pl = wb['PL']

pl_names = {}
for r in range(3, ws_pl.max_row+1):
    q_name = ws_pl.cell(r, 17).value
    if q_name:
        pl_names[str(q_name).strip().lower()] = (r, str(q_name).strip())

print(f"Total PL sheet rows with names: {len(pl_names)}")

# Extract PL items from 007 (row 89 to 234)
matched = 0
unmatched = []
for r in range(89, 235):
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v:
            name = str(v).strip()
            val_g = ws_007.cell(r, 7).value
            is_subtotal = str(val_g).startswith('=')
            key = name.lower()
            if key in pl_names:
                matched += 1
            else:
                unmatched.append((r, c, name, is_subtotal))
            break

print(f"PL Matched: {matched}, Unmatched: {len(unmatched)}")
print("\nUnmatched PL items:")
for r, lvl, name, is_sub in unmatched:
    if not is_sub:
        print(f"R{r:3d} (L{lvl}) | {name}")
