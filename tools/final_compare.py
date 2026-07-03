#!/usr/bin/env python3
"""Final comparison: reference vs new report - fixed dept mapping"""
import openpyxl
from collections import defaultdict

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'

# Read reference
wb = openpyxl.load_workbook(BASE + r'\참고_★ 2021~2026 영업 DATA FILE_전달용_v1.2_20260609의 백업.xlsx', read_only=True, data_only=True)
ws = wb['2.21~26시즌별']

old_brands = {}
r = 0
for row in ws.iter_rows(values_only=True):
    r += 1
    if r < 7: continue
    vals = list(row)
    name = str(vals[0] or '').strip()
    if not name or '요약' in name: continue
    if name.startswith('[') or name.startswith('*'): continue
    if name in ('26 SPR','26 SUM','26 FAL','26 HOL',''): continue
    try:
        qty = int(float(vals[13])) if vals[13] else 0
        if qty > 0: old_brands[name] = qty
    except: pass
wb.close()

def ref_dept(name):
    n = name.upper()
    if 'GAP' in n or 'BRFS' in n or 'MOTT' in n or 'BOW' in n or 'MADHAPPY' in n: return 'BU1-영업1부'
    if 'OLD NAVY' in n: return 'BU2-영업1부'
    if 'ON ' in n: return 'BU2-영업1부'
    if 'WALMART' in n: return 'BU2-영업2부'
    if 'QUINCE' in n: return 'BU2-영업2부'
    if n in ('BR',) or 'IPSS' in n: return 'BU3-영업1부'
    if 'KOHL' in n: return 'BU3-영업2부'
    if 'URBAN' in n: return 'BU3-영업2부'
    if 'AERIE' in n or 'OFFLINE' in n: return 'BU3-영업3부'
    if 'MLB' in n: return 'BU3-영업4부'
    return 'BU4부'

ref_dept_qty = defaultdict(int)
for name, qty in old_brands.items():
    ref_dept_qty[ref_dept(name)] += qty

# Read new report
wb2 = openpyxl.load_workbook(BASE + r'\영업DATA_5Sheet_v3_20260630.xlsx', read_only=True)
ws2 = wb2['1.시즌현황']
brand_cols = {}
new_qty = {}
for row in ws2.iter_rows(values_only=True):
    if row[0] == '시즌':
        for c, v in enumerate(row):
            if c >= 1 and v and isinstance(v, str) and '시즌' not in v:
                brand_cols[c] = v.strip()
    elif row[0] == 'Total':
        for c, v in enumerate(row):
            if c >= 1 and v and isinstance(v, (int, float)):
                brand = brand_cols.get(c, '')
                if brand: new_qty[brand] = v
wb2.close()

def new_dept(name):
    n = name.upper()
    if 'GAP' in n and 'IPSS' not in n: return 'BU1-영업1부'
    if 'BRFS' in n: return 'BU1-영업1부'
    if 'MADHAPPY' in n or 'MOTT' in n or 'BOW' in n: return 'BU1-영업1부'
    if 'OLD NAVY' in n or ('ON ' in n): return 'BU2-영업1부'
    if 'WAL' in n or 'QUINCE' in n: return 'BU2-영업2부'
    if 'BANANA' in n: return 'BU3-영업1부'
    if 'KOHL' in n: return 'BU3-영업2부'
    if 'URBAN' in n: return 'BU3-영업2부'
    if 'AERIE' in n or 'OFFLINE' in n or '신원' in n: return 'BU3-영업3부'
    if 'MLB' in n: return 'BU3-영업4부'
    return 'BU4부'

new_dept_qty = defaultdict(int)
for name, qty in new_qty.items():
    new_dept_qty[new_dept(name)] += qty

# Comparison
DEPT_ORDER = ['BU1-영업1부','BU2-영업1부','BU2-영업2부','BU3-영업1부','BU3-영업2부','BU3-영업3부','BU3-영업4부','BU4부']

print('=' * 70)
print('         2주전 vs 현재 부서별 수량 비교')
print('=' * 70)
print(f'{"부서":20s} {"2주전":>12s} {"현재":>12s} {"증감":>10s} {"%":>8s}')
print('-' * 62)

ok = True
for dept in DEPT_ORDER:
    v1 = int(ref_dept_qty.get(dept, 0))
    v2 = int(new_dept_qty.get(dept, 0))
    diff = v2 - v1
    pct = diff / v1 * 100 if v1 else 0
    flag = ''
    if abs(pct) > 50 and v1 > 100000:
        flag = ' <<< 검증필요'
        ok = False
    print(f'{dept:20s} {v1:>12,} {v2:>12,} {diff:>+10,} {pct:>+7.1f}%{flag}')

print('-' * 62)
total_ref = int(sum(ref_dept_qty.values()))
total_new = int(sum(new_dept_qty.values()))
print(f'{"전체합계":20s} {total_ref:>12,} {total_new:>12,} {total_new-total_ref:>+10,} {(total_new-total_ref)/total_ref*100:>+7.1f}%')

print()
if ok:
    print('모든 부서 증가율이 50% 미만입니다 - 정상 범위')
else:
    print('증감률이 큰 부서가 있습니다. 아래 브랜드 매칭 확인 필요')

# BU1 brand detail
print()
print('--- BU1-영업1부 브랜드 상세 ---')
print(f'{"참고(2주전)":25s} {"수량":>10s}   {"현재 브랜드":20s} {"수량":>10s}')
for old_name in sorted([n for n in old_brands if ref_dept(n)=='BU1-영업1부'], key=lambda n: -old_brands[n]):
    oq = old_brands[old_name]
    new_found = [(nn, nq) for nn, nq in sorted(new_qty.items()) if new_dept(nn)=='BU1-영업1부' and 
                 (old_name[:4].upper() in nn.upper() or nn[:4].upper() in old_name[:4].upper())]
    if not new_found:
        # show gap brands
        new_found = [(nn, nq) for nn, nq in sorted(new_qty.items()) if new_dept(nn)=='BU1-영업1부']
    new_str = ' | '.join(f'{n[:18]:18s}{q:>8,}' for n,q in new_found[:3])
    print(f'{old_name:25s} {oq:>10,}   {new_str}')

# BU3 issues
print()
print('--- BU3-영업1부 vs BU3-영업2부 ---')
print(f'BU3-1 참고: {int(ref_dept_qty.get("BU3-영업1부", 0)):>10,}')
print(f'BU3-1 현재: {int(new_dept_qty.get("BU3-영업1부", 0)):>10,}')
for nn, nq in sorted(new_qty.items()):
    if new_dept(nn) in ('BU3-영업1부', 'BU3-영업2부'):
        print(f'  {new_dept(nn):15s} {nn:20s} {int(nq):>10,}')

# Check KOHL
print()
print('--- KOHL brand in new report ---')
for nn, nq in sorted(new_qty.items()):
    if 'KOHL' in nn.upper():
        print(f'  {nn:25s} {int(nq):>10,} mapped to {new_dept(nn)}')
