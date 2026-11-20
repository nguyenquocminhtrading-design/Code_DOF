import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=True)
ws_007 = wb['007-DOAU data']
ws_bs = wb['BS']

bs_names = {}
for r in range(3, ws_bs.max_row+1):
    lvl = ws_bs.cell(r, 16).value
    q_name = ws_bs.cell(r, 17).value
    if q_name:
        bs_names[str(q_name).strip().lower()] = (r, lvl, str(q_name).strip())

print(f'Total BS rows with names: {len(bs_names)}')

matched = 0
unmatched = 0
for r in range(4, 87):
    for c in range(1, 7):
        v = ws_007.cell(r, c).value
        if v:
            name = str(v).strip()
            val_g = ws_007.cell(r, 7).value
            is_subtotal = str(val_g).startswith('=')
            
            key = name.lower()
            if key in bs_names:
                bs_r, bs_lvl, bs_q = bs_names[key]
                print(f'MATCH: 007 R{r:2d} "{name}" -> BS R{bs_r:2d} (Q{bs_r})')
                matched += 1
            else:
                print(f'NOT IN BS: 007 R{r:2d} (Lvl {c}) "{name}" (subtotal={is_subtotal})')
                unmatched += 1
            break

print(f'\nTotal matched: {matched}, unmatched: {unmatched}')
