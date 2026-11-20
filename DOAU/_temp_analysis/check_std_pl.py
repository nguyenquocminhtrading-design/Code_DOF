with open('DOAU_standard.vb', 'r', encoding='utf-8', errors='ignore') as f:
    text = f.read()

for line in text.splitlines():
    if 'Bonus' in line or 'Array(' in line and 'plItems' in line:
        if 'Bonus' in line or 'Override' in line or len(line.split(',')) >= 5:
            print(line.strip())
