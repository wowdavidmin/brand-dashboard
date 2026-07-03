#!/usr/bin/env python3
"""Analyze quantity distribution and check for anomalies"""
import openpyxl
from collections import Counter, defaultdict
import statistics

f2 = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\2) 영업DATA_Raw_v1.3_20260608_openclaw 작업용.xlsx'
wb = openpyxl.load_workbook(f2, read_only=True, data_only=True)
ws = wb['Raw_20260608']

seasons = {'26 SPR','26 SUM','26 FAL','26 HOL'}

# Distribution of qty per row
qty_dist = Counter()
# Dept breakdown for GAP
gap_dept_qty = defaultdict(lambda: defaultdict(int))
# Rows with very large qty
large_rows = []
# Qty by brand for GAP INC
gap_brand_qty = defaultdict(lambda: defaultdict(int))

cnt = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row); cnt += 1
    q = v[26] or 0
    season = str(v[12] or '').strip()
    if not (q and season in seasons): continue
    
    buyer = str(v[3] or '').strip()
    brand = str(v[4] or '').strip()
    dept = str(v[5] or '').strip()
    
    # Distribution bucket
    if q < 100: qty_dist['1-99'] += 1
    elif q < 500: qty_dist['100-499'] += 1
    elif q < 1000: qty_dist['500-999'] += 1
    elif q < 5000: qty_dist['1000-4999'] += 1
    elif q < 10000: qty_dist['5000-9999'] += 1
    elif q < 50000: qty_dist['10000-49999'] += 1
    elif q < 100000: qty_dist['50000-99999'] += 1
    else: qty_dist['100000+'] += 1
    
    if q >= 100000:
        large_rows.append((q, buyer, brand, season, dept))
    
    if 'GAP' in buyer.upper() and 'IPSS' not in buyer.upper():
        gap_dept_qty[dept][season] += q
        gap_brand_qty[brand][season] += q

wb.close()
print(f'Total rows: {cnt}')
print()

print('=== Qty distribution per row ===')
for bucket in ['1-99','100-499','500-999','1000-4999','5000-9999','10000-49999','50000-99999','100000+']:
    print(f'  {bucket:>12}: {qty_dist[bucket]:>6} rows')

print()
print(f'=== Large qty rows (>100,000) ===')
large_rows.sort(key=lambda x: -x[0])
for q, buyer, brand, season, dept in large_rows[:10]:
    print(f'  qty={q:>10,} | buyer={buyer:25s} | brand={brand:20s} | season={season} | dept={dept}')

print()
print('=== GAP INC by department ===')
for dept in sorted(gap_dept_qty.keys()):
    t = sum(gap_dept_qty[dept].values())
    print(f'  {dept:15s}: total={t:>10,}')
    for s in seasons:
        print(f'    {s}: {gap_dept_qty[dept][s]:>10,}')

print()
print('=== GAP INC by brand (top 15) ===')
brand_totals = [(b, sum(gap_brand_qty[b].values())) for b in gap_brand_qty]
brand_totals.sort(key=lambda x: -x[1])
for b, t in brand_totals[:15]:
    print(f'  {b:30s}: {t:>12,}')
