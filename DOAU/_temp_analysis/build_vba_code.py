import openpyxl
import re

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=False)
ws_007 = wb['007-DOAU data']

# Let's inspect each row 1 to 234
rows_data = []

for r in range(1, 235):
    txt = ""
    lvl = 0
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v and str(v).strip():
            txt = str(v).strip()
            lvl = c
            break
    
    val_g = ws_007.cell(r, 7).value
    g_str = str(val_g).strip() if val_g is not None else ""
    
    # Determine type
    row_type = ""
    formula_spec = ""
    
    if r == 1:
        row_type = "TITLE"
        txt = "FINANCIAL STATEMENT SUMMARY"
    elif r == 2:
        row_type = "TITLE"
        txt = "BALANCE SHEET"
    elif r == 3:
        row_type = "HEADER_MONTH"
    elif r in [87, 88, 232]:
        row_type = "BLANK"
    elif r == 89:
        row_type = "TITLE"
        txt = "PROFIT AND LOSS"
    elif g_str.startswith('='):
        row_type = "FORMULA"
        # translate column letters in formula: in 007 it refers to column G (or G..G)
        # replace any G(\d+) with D\1
        f_trans = re.sub(r'\bG(\d+)\b', r'D\1', g_str)
        formula_spec = f_trans
    elif g_str != "":
        # Has value in G -> it's a leaf item
        if r < 89:
            row_type = "BS_SUMIFS"
        else:
            row_type = "PL_SUMIFS"
    else:
        # No value in G -> check if it's a leaf item that happens to be empty in Jan, or a header
        # Let's check if any month has value for this row
        has_any_val = False
        for c in range(7, 19):
            mv = ws_007.cell(r, c).value
            if mv is not None and str(mv).strip() != "":
                has_any_val = True
                break
        if has_any_val:
            if r < 89:
                row_type = "BS_SUMIFS"
            else:
                row_type = "PL_SUMIFS"
        else:
            row_type = "HEADER"

    rows_data.append((r, lvl, txt, row_type, formula_spec))

print(f"Total processed rows: {len(rows_data)}")

# Print breakdown of row types
from collections import Counter
counts = Counter(rt for r, lvl, txt, rt, fs in rows_data)
print("Row types breakdown:", counts)

# Print rows that are FORMULA
print("\nFormulas:")
for r, lvl, txt, rt, fs in rows_data:
    if rt == "FORMULA":
        print(f"R{r:3d} | {txt:<35} | {fs}")
