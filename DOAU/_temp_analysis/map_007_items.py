import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=False)
ws_007 = wb['007-DOAU data']

print("=== PARSING EXACT BS ITEMS FROM 007 ===")
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

print(f"Total BS items: {len(bs_items)}")
for idx, (r, lvl, name, has_val, formula) in enumerate(bs_items, 1):
    print(f"BS {idx:2d} | 007_R{r:2d} | L{lvl} | {name:<35} | hasVal={has_val!s:<5} | form={formula[:30]}")

print("\n=== PARSING EXACT PL ITEMS FROM 007 ===")
pl_items = []
for r in range(89, 235):
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
        pl_items.append((r, lvl, name, has_val, formula))

print(f"Total PL items: {len(pl_items)}")
for idx, (r, lvl, name, has_val, formula) in enumerate(pl_items, 1):
    print(f"PL {idx:3d} | 007_R{r:3d} | L{lvl} | {name:<35} | hasVal={has_val!s:<5} | form={formula[:30]}")
