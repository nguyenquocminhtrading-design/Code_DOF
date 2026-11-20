import openpyxl
import re

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=False)
ws_007 = wb['007-DOAU data']

true_headers = {
    # BS
    4, 5, 6, 12, 16, 17, 25, 39, 44, 45, 46, 48, 53, 55, 60, 66, 79,
    # PL
    90, 91, 92, 105, 106, 112, 119, 136, 158, 159, 161, 175, 186, 205, 220, 221, 225
}

rows_spec = []

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
        formula_spec = re.sub(r'\bG(\d+)\b', r'D\1', g_str)
    elif r in true_headers:
        row_type = "HEADER"
    else:
        if r < 89:
            row_type = "BS_SUMIFS"
        else:
            row_type = "PL_SUMIFS"
            
    rows_spec.append((r, lvl, txt, row_type, formula_spec))

print("Total rows:", len(rows_spec))
from collections import Counter
print(Counter(rt for r, lvl, txt, rt, fs in rows_spec))

with open('_temp_analysis/rows_spec.txt', 'w', encoding='utf-8') as f:
    for r, lvl, txt, rt, fs in rows_spec:
        f.write(f"r={r:3d} | lvl={lvl} | {txt:<35} | {rt:<12} | {fs}\n")
print("Exported rows_spec.txt")
