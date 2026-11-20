import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=True)
ws_007 = wb['007-DOAU data']
ws_bs = wb['BS']

# Let's read all BS items from 007
bs_items = []
for r in range(4, 86):
    lvl = None
    name = None
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v is not None and str(v).strip() != '':
            lvl = c
            name = str(v).strip()
            break
    if name is not None:
        f = ws_007.cell(r, 7).value
        has_val = (f is not None)
        formula = str(f) if f is not None else ""
        bs_items.append((r, lvl, name, has_val, formula))

# Build lookup from BS sheet
# Map (lvl, clean_name) -> K value, and clean_name -> list of (lvl, K)
bs_data = []
for r in range(3, ws_bs.max_row+1):
    lvl = ws_bs.cell(r, 16).value
    name = ws_bs.cell(r, 17).value
    k = ws_bs.cell(r, 11).value
    if name:
        bs_data.append((r, lvl, str(name).strip().lower(), k if k is not None else 0.0))

print("=== TESTING BS MATCHING WITH AND WITHOUT LEVEL ===")
for idx, (r_007, lvl_007, name, has_val, form) in enumerate(bs_items, 1):
    if not has_val or form.startswith('='):
        continue
    
    key = name.lower()
    # Match with level
    sum_with_lvl = sum(k for r, l, n, k in bs_data if n == key and l == lvl_007 and isinstance(k, (int, float)))
    # Match without level
    sum_no_lvl = sum(k for r, l, n, k in bs_data if n == key and isinstance(k, (int, float)))
    
    if sum_with_lvl != sum_no_lvl:
        print(f"BS {idx:2d} (007_R{r_007:2d}) '{name}' (Lvl {lvl_007}): With Lvl = {sum_with_lvl}, No Lvl = {sum_no_lvl}")
