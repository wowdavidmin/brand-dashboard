#!/usr/bin/env python3
"""Precise comparison of old vs new report data"""
import openpyxl
from collections import defaultdict

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
ref_file = BASE + r'\참고_백업.xlsx'
new_report = BASE + r'\영업DATA_5Sheet_v3_20260630.xlsx'

# --- Read reference (old report) ---
print('Reading reference...')
wb = openpyxl.load_workbook(ref_file, read_only=True, data_only=True)
ws = wb['2.21~26시즌별']

# The sheet has 4 season groups + TOTAL
# Reading layout: Row 5 = header (seasons), Row 6 = sub-header (qty/total/net)
# Row 7+ = data: brand, SPR_qty, SPR_total, SPR_net, SUM_qty, SUM_total, SUM_net, ...
ref_old = {}
r = 0
for row in ws.iter_rows(values_only=True):
    r += 1
    vals = list(row)
    if r < 7: continue
    name = str(vals[0] or '').strip()
    if not name or '요약' in name or name in ('','26 SPR','26 SUM','26 FAL','26 HOL'):
        continue
    if name.startswith('[') or name.startswith('※'): continue
    
    # Get total qty from column 14 (0-indexed: col 13 = TOTAL qty)
    # Structure: col0=brand, col1=SPR_qty, col4=SUM_qty, col7=FAL_qty, col10=HOL_qty, col13=TOTAL_qty
    try:
        total_qty = int(float(vals[13])) if vals[13] else 0
        if total_qty > 0:
            ref_old[name] = total_qty
    except (ValueError, TypeError, IndexError):
        pass

wb.close()
print(f'  Parsed {len(ref_old)} brands')

# Aggregate by dept for reference
def old_dept(name):
    n = name.upper()
    if 'GAP' in n: return 'BU1'
    if 'IPSS GAP' in n: return 'BU1'
    if 'BRFS' in n: return 'BU1'
    if 'MOTT' in n or 'BOW' in n: return 'BU1'
    if 'MADHAPPY' in n: return 'BU1'
    if 'OLD NAVY' in n: return 'BU2-1'
    if 'ON ' in n: return 'BU2-1'
    if 'WALMART' in n: return 'BU2-2'
    if 'QUINCE' in n: return 'BU2-2'
    if n in ('BR',): return 'BU3-1'
    if 'IPSS OLD' in n or 'IPSS ON' in n: return 'BU3-1'
    if 'KOHL' in n: return 'BU3-2'
    if 'URBAN' in n: return 'BU3-2'
    if 'AERIE' in n: return 'BU3-3'
    if 'OFFLINE' in n: return 'BU3-3'
    if 'MLB' in n: return 'BU3-4'
    if '영업1부' in n or '영업2부' in n or '영업3부' in n or '영업4부' in n or '영업5부' in n or '영업6부' in n:
        return 'SUBTOTAL'
    if '총합계' in n: return 'GRAND'
    return 'BU4'

ref_dept = defaultdict(int)
for name, qty in ref_old.items():
    d = old_dept(name)
    if d not in ('SUBTOTAL', 'GRAND'):
        ref_dept[d] += qty

# --- Read new report ---
print('Reading new report...')
wb2 = openpyxl.load_workbook(new_report, read_only=True)
ws2 = wb2['1.시즌현황']
new_brands = {}
for row in ws2.iter_rows(values_only=True):
    if row[0] == 'Total':
        for c, v in enumerate(row):
            if c >= 1 and v is not None and isinstance(v, (int, float)):
                # Find brand name from row 4
                pass

# Get brand names from row 4
brand_names = {}
for row2 in ws2.iter_rows(values_only=True):
    if row2[0] == '시즌':
        for c, v in enumerate(row2):
            if c >= 1 and v and isinstance(v, str):
                brand_names[c] = v.strip()
    elif row2[0] == 'Total':
        new_brands = {brand_names.get(c, f'col{c}'): v for c, v in enumerate(row2) if c >= 1 and v and isinstance(v, (int,float))}
        break
wb2.close()

print(f'  Parsed {len(new_brands)} brands')

def new_dept(name):
    n = name.upper()
    if 'GAP' in n and 'IPSS' not in n: return 'BU1'
    if 'IPSS GAP' in n: return 'BU1'
    if 'BRFS' in n: return 'BU1'
    if 'MADHAPPY' in n: return 'BU1'
    if 'MOTT' in n or 'BOW' in n: return 'BU1'
    if 'OLD NAVY' in n: return 'BU2-1'
    if 'ON ' in n: return 'BU2-1'
    if 'WAL' in n: return 'BU2-2'
    if 'QUINCE' in n: return 'BU2-2'
    if 'BANANA' in n or 'KOHL' in n: return 'BU3-1'
    if 'URBAN' in n: return 'BU3-2'
    if 'AERIE' in n or 'OFFLINE' in n or '신원' in n: return 'BU3-3'
    if 'MLB' in n: return 'BU3-4'
    return 'BU4'

new_dept_qty = defaultdict(int)
for name, qty in new_brands.items():
    new_dept_qty[new_dept(name)] += qty

# --- Print comparison ---
print('\n============= DEPARTMENT COMPARISON =============')
print(f'{"Dept":12s} {"2주전(참고)":>14s} {"현재":>12s} {"증감":>12s} {"증감률":>10s}')
print('-' * 62)
for dept in ['BU1', 'BU2-1', 'BU2-2', 'BU3-1', 'BU3-2', 'BU3-3', 'BU3-4', 'BU4']:
    v1 = ref_dept.get(dept, 0)
    v2 = new_dept_qty.get(dept, 0)
    diff = v2 - v1
    pct = f'{diff/v1*100:+.1f}%' if v1 else 'NEW'
    if dept == 'BU1':
        print(f'{"━"*62}')
        print(f'  ★ 영업1부')
    print(f'{dept:12s} {v1:>12,}  {v2:>12,}  {diff:>+10,}  {pct:>8s}')

# --- Brand matching ---
print('\n============= BRAND MATCHING =============')
print(f'{"참고(2주전) 브랜드":30s} {"수량":>10s}  |  {"현재 브랜드":25s} {"수량":>10s}  {"차이":>10s}')
print('-' * 95)

# Manual brand matching
brand_pairs = [
    # Old brand -> New brand(s)
    ('GAP KIDS', ['Gap', 'Gap OUTLET']),
    ("GAP MEN'S", ['Gap', 'Gap OUTLET']),
    ("GAP WMN'S", ['Gap', 'Gap OUTLET']),
    ('IPSS GAP', ['ON IPSS']),
    ('BRFS', ['BRFS']),
    ('Mott & Bow', ['Mott & Bow']),
    ('OLD NAVY', ['OLD NAVY']),
    ("ON  MEN'S", ['ON IPSS']),
    ("ON  WMN'S", ['ON IPSS']),
    ("ON KID'S & BABY", ['ON IPSS']),
    ('WALMART CANADA', ['WAL-MART']),
    ("WALMART D29 (INTIMATE)", ['WAL-MART']),
    ("WALMART D26 (BABY)", ['WAL-MART']),
    ('QUINCE', ['QUINCE']),
    ('BR', ['BANANA REPUBLIC']),
    ('IPSS OLD NAVY', ['BANANA REPUBLIC']),
    ("Kohl's", ["KOHL`S"]),
    ('URBAN', ['URBAN OUTFITTERS']),
    ('AERIE WOMEN/S KNIT', ['AERIE']),
    ("AERIE WOMEN'S ACTIVE", ['AERIE']),
    ('offline- Apparel', ['Offline']),
    ('MLB', ['MLB']),
    ('NUUDS', ['NUUDS']),
    ('ONER ACTIVE', ['ONER ACTIVE']),
    ('ROSS', ['ROSS']),
    ("AEO WOMEN'S ACTIVE", ['AMERICAN EAGLE']),
    ("AEO WOMEN'S C&S", ['AMERICAN EAGLE']),
    ("AEO MEN'S", ['AMERICAN EAGLE']),
]

print(f'\n{"BU1 - 영업1부":-^95s}')
bu1_ref = 0; bu1_new = 0
for old_name, new_names in brand_pairs:
    old_q = ref_old.get(old_name, 0)
    if old_q == 0: continue
    if new_dept(old_name) != 'BU1': continue
    
    new_q_total = sum(new_brands.get(n, 0) for n in new_names)
    diff = new_q_total - old_q
    pct = f'{diff/old_q*100:+.1f}%' if old_q else ''
    
    new_str = '+'.join(f'{n}({new_brands.get(n,0):,})' for n in new_names if new_brands.get(n,0)>0)
    print(f'{old_name:30s} {old_q:>10,}  |  {new_str:40s}  {diff:>+10,}  {pct:>8s}')
    bu1_ref += old_q; bu1_new += new_q_total

print(f'{"="*60}')
print(f'{"BU1 TOTAL":30s} {bu1_ref:>10,}  ->  {bu1_new:>10,}  ({bu1_new-bu1_ref:>+10,} / {(bu1_new-bu1_ref)/bu1_ref*100:+.1f}%)')

# Remaining unmatched brands
matched_new = set()
for old_name, new_names in brand_pairs:
    for n in new_names: matched_new.add(n)
unmatched = [f'{n}({q:,})' for n,q in sorted(new_brands.items()) if n not in matched_new and new_dept(n) == 'BU1' and q > 0]
if unmatched:
    print(f'  NEW brands in BU1: {", ".join(unmatched)}')
