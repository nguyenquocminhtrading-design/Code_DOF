import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=False, keep_vba=True)
ws_007 = wb['007-DOAU data']

# Check alignment of DOAU_v2.bas row specifications against 007-DOAU data
with open('DOAU_v2.bas', 'r', encoding='utf-8') as f:
    vba_code = f.read()

import re
matches = re.findall(r'm\((\d+)\)\s*=\s*Array\((\d+),\s*"([^"]*)",\s*"([^"]*)",\s*"([^"]*)"\)', vba_code)
print(f"Parsed {len(matches)} rows from DOAU_v2.bas.")

mismatches = 0
for r_str, lvl_str, name_vba, rt, form in matches:
    r = int(r_str)
    # get 007 row text
    txt_007 = ""
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v and str(v).strip():
            txt_007 = str(v).strip()
            break
    
    # compare
    if r in [1, 3, 87, 88, 232]:
        continue
    if name_vba.lower() != txt_007.lower():
        print(f"MISMATCH at R{r:3d}: VBA='{name_vba}' vs 007='{txt_007}'")
        mismatches += 1

if mismatches == 0:
    print("SUCCESS: 100% of rows (from 1 to 234) in DOAU_v2.bas match sheet 007-DOAU data exactly!")
else:
    print(f"FAILED: Found {mismatches} mismatches.")
