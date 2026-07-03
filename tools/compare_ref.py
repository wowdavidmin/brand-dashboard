#!/usr/bin/env python3
"""Compare new report vs reference (2-week-old) data"""
import openpyxl
from collections import defaultdict

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
ref_file = BASE + r'\참고_백업.xlsx'
new_report = BASE + r'\영업DATA_5Sheet_v3_20260630.xlsx'

print('=== Reading reference file (2 weeks ago) ===')
wb_ref = openpyxl.load_workbook(ref_file, read_only=True, data_only=True)

# Sheet 2 (2.21~26시즌별) has brand-level data
ws_ref = wb_ref['2.21~26시즌별']

# Parse: rows are brands, columns are seasons (SPR qty/total/net, SUM qty/total/net, ...)
# This sheet has complex layout. Let me extract what we can.
ref_brand_data = {}
r = 0
season_order = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
for row in ws_ref.iter_rows(values_only=True):
    r += 1
    if r < 7: continue  # skip headers
    vals = list(row)
    brand_name = str(vals[0] or '').strip()
    if not brand_name or brand_name.startswith('[') or brand_name.startswith('※'):
        continue
    # Check if it's a data row or a subtotal
    if '요약' in brand_name or brand_name in season_order:
        continue
    
    # Extract qty data: columns 1,4,7,10,13 (SPR qty, SUM qty, FAL qty, HOL qty, Total qty)
    qtys = []
    for ci in [1, 4, 7, 10, 13]:
        q = vals[ci] if vals[ci] else 0
        try: q = float(q)
        except: q = 0
        qtys.append(int(q))
    
    ref_brand_data[brand_name] = {
        'qty': {s: qtys[i] for i, s in enumerate(season_order)},
        'total_qty': qtys[4],
    }

wb_ref.close()
print(f'  Parsed {len(ref_brand_data)} brands from reference')

# Compute reference totals
ref_dept_totals = defaultdict(int)
dept_map = {
    'GAP KIDS': 'BU1', 'GAP MEN\'S': 'BU1', "GAP WMN'S": 'BU1',
    'IPSS GAP': 'BU1', 'BRFS': 'BU1', 'Mott & Bow': 'BU1',
    'Madhappy': 'BU1',
    'OLD NAVY': 'BU2-1', "ON  MEN'S": 'BU2-1', "ON  WMN'S": 'BU2-1', "ON KID'S & BABY": 'BU2-1',
    'WALMART CANADA': 'BU2-2', "WALMART D29 (INTIMATE)": 'BU2-2', 'QUINCE': 'BU2-2',
    'BR': 'BU3-1', 'IPSS OLD NAVY': 'BU3-1',
    "Kohl's": 'BU3-2', 'URBAN': 'BU3-2',
    'AERIE WOMEN/S KNIT': 'BU3-3', "AERIE WOMEN'S ACTIVE": 'BU3-3', 'offline- Apparel': 'BU3-3',
    'MLB': 'BU3-4',
    'NUUDS': 'BU4', 'ONER ACTIVE': 'BU4', 'ROSS': 'BU4',
    "AEO WOMEN'S ACTIVE": 'BU4', "AEO WOMEN'S C&S": 'BU4', "AEO MEN'S": 'BU4',
}

for brand, data in ref_brand_data.items():
    for ref_name, dept in dept_map.items():
        if brand == ref_name or brand.startswith(ref_name.split("'")[0]):
            ref_dept_totals[dept] += data['total_qty']
            break
    else:
        ref_dept_totals['기타'] += data['total_qty']

print(f'\n=== Reference BU1 Total: {ref_dept_totals.get("BU1", 0):>10,}')

# Now read new report
print('\n=== Reading new report ===')
wb_new = openpyxl.load_workbook(new_report, read_only=True)
ws_new = wb_new['1.시즌현황']

# Find Total row
new_totals = {}
r = 0
for row in ws_new.iter_rows(values_only=True):
    r += 1
    if r == 4:  # brand header row
        brand_cols = {}
        for c, v in enumerate(row):
            if v and isinstance(v, str) and v not in ('시즌', ''):
                brand_cols[c] = v.strip()
    if row[0] == 'Total':
        for c, v in enumerate(row):
            if v and isinstance(v, (int, float)) and c in brand_cols:
                new_totals[brand_cols[c]] = v

wb_new.close()

# Map new report brands to departments
new_dept_totals = defaultdict(int)
for brand, qty in new_totals.items():
    b = brand.upper()
    if 'GAP' in b and 'IPSS' not in b and 'OLD NAVY' not in b: new_dept_totals['BU1'] += qty
    elif 'IPSS GAP' in b: new_dept_totals['BU1'] += qty
    elif 'BRFS' in b: new_dept_totals['BU1'] += qty
    elif 'MADHAPPY' in b: new_dept_totals['BU1'] += qty
    elif 'MOTT' in b or 'BOW' in b: new_dept_totals['BU1'] += qty
    elif 'OLD NAVY' in b: new_dept_totals['BU2-1'] += qty
    elif 'ON ' in b: new_dept_totals['BU2-1'] += qty
    elif 'WAL' in b: new_dept_totals['BU2-2'] += qty
    elif 'QUINCE' in b: new_dept_totals['BU2-2'] += qty
    elif 'BANANA' in b or 'KOHL' in b: new_dept_totals['BU3-1'] += qty
    elif 'URBAN' in b: new_dept_totals['BU3-2'] += qty
    elif 'AERIE' in b or 'OFFLINE' in b: new_dept_totals['BU3-3'] += qty
    elif 'MLB' in b: new_dept_totals['BU3-4'] += qty
    else: new_dept_totals['BU4'] += qty

print('\n============= COMPARISON =============')
print(f'{"Dept":20s} {"2주전":>12s} {"현재":>12s} {"증감":>12s} {"증감률":>10s}')
print('-' * 66)
for dept in ['BU1', 'BU2-1', 'BU2-2', 'BU3-1', 'BU3-2', 'BU3-3', 'BU3-4', 'BU4']:
    ref_q = ref_dept_totals.get(dept, 0)
    new_q = new_dept_totals.get(dept, 0)
    diff = new_q - ref_q
    pct = f'{diff/ref_q*100:+.1f}%' if ref_q else 'N/A'
    print(f'{dept:20s} {ref_q:>12,} {new_q:>12,} {diff:>+12,} {pct:>10s}')

total_ref = sum(ref_dept_totals.values())
total_new = sum(new_dept_totals.values())
print('-' * 66)
print(f'{"GRAND TOTAL":20s} {total_ref:>12,} {total_new:>12,} {total_new-total_ref:>+12,} {(total_new-total_ref)/total_ref*100:+.1f}%')

# Compare specific brands
print('\n=== Brand-level comparison (Top brands) ===')
brand_compare = []
for brand, data in ref_brand_data.items():
    ref_q = data['total_qty']
    new_q = new_totals.get(brand, 0)
    # Also try matching with alternative names
    if new_q == 0:
        for nb, nq in new_totals.items():
            if nb.upper()[:4] == brand.upper()[:4] or brand.upper()[:4] == nb.upper()[:4]:
                new_q = nq
                break
    if ref_q > 100000:  # Only significant brands
        brand_compare.append((brand, ref_q, new_q))
brand_compare.sort(key=lambda x: -x[1])
for brand, ref_q, new_q in brand_compare[:15]:
    diff = new_q - ref_q
    pct = f'{diff/ref_q*100:+.1f}%' if ref_q else 'N/A'
    print(f'  {brand:25s} {ref_q:>10,} -> {new_q:>10,} ({diff:>+8,} / {pct:>8s})')
