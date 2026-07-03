#!/usr/bin/env python3
"""Compare Brand(D) vs Division(K) columns for GAP/IPSS GAP"""
import openpyxl
from collections import defaultdict

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
f2 = BASE + r'\2) 영업DATA_Raw_v1.3_20260608_openclaw 작업용.xlsx'
wb = openpyxl.load_workbook(f2, read_only=True, data_only=True)
ws = wb['Raw_20260608']

seasons = {'26 SPR','26 SUM','26 FAL','26 HOL'}
by_brand = defaultdict(lambda: [0, 0])  # brand -> [qty, amt]
by_division = defaultdict(lambda: [0, 0])  # division -> [qty, amt]
gap_outlet_division = defaultdict(int)  # what Division does Gap OUTLET brand use?
gap_outlet_brand_count = 0

cnt = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row); cnt += 1
    q = v[26] or 0; a = v[27] or 0
    season = str(v[12] or '').strip()
    if not (q and season in seasons): continue
    brand = str(v[4] or '').strip()
    div = str(v[10] or '').strip()
    by_brand[brand][0] += q; by_brand[brand][1] += a
    by_division[div][0] += q; by_division[div][1] += a
    if brand == 'Gap OUTLET':
        gap_outlet_division[div] += q
        gap_outlet_brand_count += 1

wb.close()
print(f'Total rows: {cnt}')

print('\n=== GAP/IPSS GAP by BRAND(D열) ===')
for b in ['Gap', 'Gap OUTLET', 'BRFS', 'Madhappy', 'Mott & Bow', 'IPSS GAP']:
    if b in by_brand:
        print(f'  {b:20s}: qty={int(by_brand[b][0]):>10,}  amt=${int(by_brand[b][1]):>10,}')

print('\n=== GAP/IPSS GAP by DIVISION(K열) ===')
for d in ['GAP', 'IPSS GAP', 'Gap OUTLET', 'BRFS', 'MADHAPPY', 'Mott & Bow']:
    if d in by_division:
        print(f'  {d:20s}: qty={int(by_division[d][0]):>10,}  amt=${int(by_division[d][1]):>10,}')

print(f'\n=== Gap OUTLET brand -> Division mapping ===')
for d, q in sorted(gap_outlet_division.items(), key=lambda x: -x[1]):
    print(f'  Division={d:20s} qty={int(q):>10,}')

print(f'\n=== All Division values for GAP-related ===')
for d in sorted(by_division.keys()):
    if 'GAP' in d.upper() or 'BRFS' in d or 'MOTT' in d or 'BOW' in d or 'MADHAPPY' in d:
        print(f'  {d:25s}: qty={int(by_division[d][0]):>10,}')

print(f'\n=== Gap OUTLET data: {gap_outlet_brand_count} rows ===')
