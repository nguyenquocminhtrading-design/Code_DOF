import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=True)
ws_007 = wb['007-DOAU data']
ws_bs = wb['BS']
ws_pl = wb['PL']

bs_lookup = {}
for r in range(3, ws_bs.max_row+1):
    lvl = ws_bs.cell(r, 16).value
    q_name = ws_bs.cell(r, 17).value
    if q_name:
        bs_lookup[str(q_name).strip().lower()] = (r, lvl, str(q_name).strip())

print("=== CHECK BS LEVELS ===")
for r in range(4, 86):
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v:
            name = str(v).strip()
            key = name.lower()
            if key in bs_lookup:
                bs_r, bs_lvl, bs_q = bs_lookup[key]
                if c != bs_lvl:
                    print(f"LEVEL MISMATCH: 007 R{r:2d} '{name}' has Lvl {c}, but BS R{bs_r:2d} has Lvl {bs_lvl}!")
            break

pl_lookup = {}
for r in range(3, ws_pl.max_row+1):
    lvl = ws_pl.cell(r, 16).value
    q_name = ws_pl.cell(r, 17).value
    if q_name:
        pl_lookup[str(q_name).strip().lower()] = (r, lvl, str(q_name).strip())

print("\n=== CHECK PL LEVELS ===")
for r in range(89, 235):
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v:
            name = str(v).strip()
            key = name.lower()
            if key in pl_lookup:
                pl_r, pl_lvl, pl_q = pl_lookup[key]
                if c != pl_lvl:
                    print(f"LEVEL MISMATCH: 007 R{r:2d} '{name}' has Lvl {c}, but PL R{pl_r:2d} has Lvl {pl_lvl}!")
            break
