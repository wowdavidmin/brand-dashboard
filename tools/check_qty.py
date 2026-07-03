#!/usr/bin/env python3
"""Check quantity data from file #2"""
import openpyxl
from collections import defaultdict

f2 = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\2) 영업DATA_Raw_v1.3_20260608_openclaw 작업용.xlsx'
wb = openpyxl.load_workbook(f2, read_only=True, data_only=True)
ws = wb['Raw_20260608']

seasons = {'26 SPR','26 SUM','26 FAL','26 HOL'}
buyer_qty = defaultdict(lambda: defaultdict(int))
brand_qty = defaultdict(lambda: defaultdict(int))
cnt = 0

for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row); cnt += 1
    q = v[26] or 0
    season = str(v[12] or '').strip()
    if not (q and season in seasons): continue
    buyer = str(v[3] or '').strip()
    brand = str(v[4] or '').strip()
    buyer_qty[buyer][season] += q
    brand_qty[brand][season] += q

wb.close()

print(f'Total rows: {cnt}')
print()

# Buyers containing GAP
print('=== GAP-related buyers ===')
for b in sorted(b for b in buyer_qty if 'GAP' in b.upper()):
    for s in seasons:
        print(f'  {b:30s} {s}: {buyer_qty[b][s]:>12,}')
    t = sum(buyer_qty[b].values())
    print(f'  {b:30s} TOTAL: {t:>12,}')
    print()

# All buyers totals
print('=== ALL BUYERS season totals ===')
all_buyer_totals = defaultdict(int)
for b in buyer_qty:
    for s in seasons:
        all_buyer_totals[s] += buyer_qty[b][s]
for s in seasons:
    print(f'  {s}: {all_buyer_totals[s]:>12,}')
print(f'  GRAND TOTAL: {sum(all_buyer_totals.values()):>12,}')

# Compare: what was in file #3 sheet 2?
print()
print('=== File #3 reference data (BU1_D1 total from old report) ===')
print('  From 3) file sheet 2.21~26시즌별:')
print('  BU1_D1 SPR: 3,615,218')
print('  BU1_D1 SUM: 6,521,759')
print('  BU1_D1 FAL: 10,298,563')
print('  BU1_D1 HOL: 5,996,010')
print('  BU1_D1 Total: 26,431,550')
