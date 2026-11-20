import openpyxl
import re

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=False)
ws_007 = wb['007-DOAU data']

print("=== BS ITEMS WITH FORMULAS IN 007 (Col G & Col O) ===")
# Map 007 row -> BS item index (1-based)
bs_row_to_idx = {}
idx = 1
for r in range(4, 86):
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v and str(v).strip():
            bs_row_to_idx[r] = idx
            idx += 1
            break

for r, i in bs_row_to_idx.items():
    name = ws_007.cell(r, 1).value or ws_007.cell(r, 2).value or ws_007.cell(r, 3).value or ws_007.cell(r, 4).value or ws_007.cell(r, 5).value or ws_007.cell(r, 6).value
    g_form = ws_007.cell(r, 7).value
    o_form = ws_007.cell(r, 15).value
    if (g_form and str(g_form).startswith('=')) or (o_form and str(o_form).startswith('=')):
        print(f"Item {i:2d} (R{r:2d}) '{name}': G='{g_form}' | O='{o_form}'")

print("\n=== PL ITEMS WITH FORMULAS IN 007 (Col G & Col O) ===")
pl_row_to_idx = {}
idx = 1
for r in range(89, 235):
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v and str(v).strip():
            pl_row_to_idx[r] = idx
            idx += 1
            break

for r, i in pl_row_to_idx.items():
    name = ws_007.cell(r, 1).value or ws_007.cell(r, 2).value or ws_007.cell(r, 3).value or ws_007.cell(r, 4).value or ws_007.cell(r, 5).value or ws_007.cell(r, 6).value
    g_form = ws_007.cell(r, 7).value
    o_form = ws_007.cell(r, 15).value
    if (g_form and str(g_form).startswith('=')) or (o_form and str(o_form).startswith('=')):
        print(f"Item {i:3d} (R{r:3d}) '{name}': G='{g_form}' | O='{o_form}'")
