with open('DOAU_standard.vb', 'r', encoding='utf-8', errors='ignore') as f:
    text = f.read()

lines = text.splitlines()
for i, l in enumerate(lines):
    if 'Write Profit & Loss section' in l:
        for j in range(i, min(i+60, len(lines))):
            print(lines[j])
        break
