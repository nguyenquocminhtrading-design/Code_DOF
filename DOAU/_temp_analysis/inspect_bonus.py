import openpyxl

wb = openpyxl.load_workbook('FS DOAU Sep 26 - Adjusted.xlsm', data_only=False)
ws_007 = wb['007-DOAU data']

print("Row 114 (Bonus in DL):", ws_007.cell(114, 7).value, ws_007.cell(114, 15).value)
print("Row 121 (Bonus in OH):", ws_007.cell(121, 7).value, ws_007.cell(121, 15).value)
print("Row 187 (Bonus in Admin):", ws_007.cell(187, 7).value, ws_007.cell(187, 15).value)
