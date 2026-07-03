#!/usr/bin/env python3
"""Rebuild all 5 sheets using BRAND (not buyer) as column dimension"""
import openpyxl, os, glob
from openpyxl.styles import Font, Border, Side, Alignment
from collections import defaultdict
from datetime import datetime

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
files = [f for f in glob.glob(os.path.join(BASE, '2)*')) if not os.path.basename(f).startswith('~$')]
F2 = max(files, key=os.path.getmtime)
OUT = os.path.join(BASE, f'영업DATA_5Sheet_v3_{datetime.now().strftime("%Y%m%d")}.xlsx')

print(f'Reading: {F2}')
wb2 = openpyxl.load_workbook(F2, read_only=True, data_only=True)
ws = wb2['Raw_20260608']

seasons_list = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
months_list = [f'{m}월' for m in range(1, 13)]

# Aggregations by BRAND instead of buyer
S1 = defaultdict(lambda: [0,0,0,0,0,0])  # (season, brand) -> [qty, amt, net, dc, profit, em]
S2 = defaultdict(lambda: [0,0,0])        # (brand, season) -> [qty, amt, net]
S3 = defaultdict(lambda: [0,0,0])        # (month, brand) -> [qty, amt, net]
S4 = defaultdict(lambda: 0)              # (dept, brand, factory, season) -> qty
S5 = defaultdict(lambda: 0)              # (factory, brand, season) -> qty
all_brands = set()
cnt = 0

for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row); cnt += 1
    q = v[26] or 0; a = v[27] or 0; n = v[29] or 0
    dc = v[48] or 0; prof = v[50] or 0; em = v[56] or 0
    brand = str(v[4] or '').strip()
    season = str(v[12] or '').strip()
    dept = str(v[5] or '').strip()
    factory = str(v[2] or '').strip()
    month = str(v[21] or '').strip()

    if not (q and a and season in seasons_list): continue

    S1[(season, brand)] = [S1[(season, brand)][i] + x for i, x in enumerate([q, a, n, dc, prof, em])]
    all_brands.add(brand)
    S2[(brand, season)] = [S2[(brand, season)][i] + x for i, x in enumerate([q, a, n])]
    if month: S3[(month, brand)] = [S3[(month, brand)][i] + x for i, x in enumerate([q, a, n])]
    S4[(dept, brand, factory, season)] += q
    S5[(factory, brand, season)] += q

    if cnt % 30000 == 0: print(f'  {cnt}...')

wb2.close()
print(f'Aggregated: {cnt} rows, {len(all_brands)} brands')

# Map brand to department
def brand_dept(b):
    x = b.upper()
    if 'GAP' in x and 'IPSS' not in x and 'OLD NAVY' not in x: return 'BU1-영업1부'
    if 'IPSS GAP' in x: return 'BU1-영업1부'
    if 'BRFS' in x: return 'BU1-영업1부'
    if 'MADHAPPY' in x: return 'BU1-영업1부'
    if 'MOTT' in x or 'BOW' in x: return 'BU1-영업1부'
    if 'OLD NAVY' in x: return 'BU2-영업1부'
    if x.startswith('ON ') or 'ON MEN' in x or 'ON WOM' in x or 'ON KID' in x: return 'BU2-영업1부'
    if 'WALMART' in x or 'WAL-MART' in x: return 'BU2-영업2부'
    if 'QUINCE' in x: return 'BU2-영업2부'
    if x in ('BR',) or 'BANANA' in x: return 'BU3-영업1부'
    if 'IPSS ON' in x: return 'BU3-영업1부'
    if 'KOHL' in x: return 'BU3-영업2부'
    if 'URBAN' in x: return 'BU3-영업2부'
    if 'AERIE' in x: return 'BU3-영업3부'
    if 'OFFLINE' in x or '신원' in x or '인터텍스' in x: return 'BU3-영업3부'
    if 'MLB' in x: return 'BU3-영업4부'
    if 'NUUDS' in x or 'ONER' in x or 'ROSS' in x or 'AEO' in x or 'MAKALOT' in x: return 'BU4부'
    if 'GAP OUTLET' in x or 'GAP' in x: return 'BU1-영업1부'
    return 'BU4부'

DEPT_ORDER = ['BU1-영업1부','BU2-영업1부','BU2-영업2부',
              'BU3-영업1부','BU3-영업2부','BU3-영업3부','BU3-영업4부','BU4부']

dept_brands = defaultdict(list)
for b in sorted(all_brands):
    dept_brands[brand_dept(b)].append(b)

# Styling
FMT_INT = '#,##0'; FMT_DEC2 = '#,##0.00'; FMT_PCT = '0.00'; FMT_RATIO = '0.0000'
thin = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
hdr = Font(bold=True, size=10, name='Calibri')
ttl = Font(bold=True, size=13, name='Calibri')
numf = Font(size=10, name='Calibri')

def k(v): return int(round(v / 1000))

wb = openpyxl.Workbook()

# ======================================================
# SHEET 1: 시즌현황 (Season x Brand Matrix)
# ======================================================
ws1 = wb.active; ws1.title = '1.시즌현황'
ws1.cell(row=1, column=1, value='[약진통상] 2026 Season 부서별 브랜드별 현황 (단위: $K, 수량: PCS)').font = ttl
ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=80)

pair_col1 = {}; col = 2
for dg in DEPT_ORDER:
    brands = dept_brands.get(dg, [])
    if not brands: continue
    cs = col
    for b in brands:
        pair_col1[(dg, b)] = col; col += 6
    ce = col - 1
    if ce > cs:
        ws1.merge_cells(start_row=3, start_column=cs, end_row=3, end_column=ce)
        ws1.cell(row=3, column=cs, value=dg).font = hdr
        ws1.cell(row=3, column=cs).alignment = Alignment(horizontal='center')

ws1.cell(row=4, column=1, value='시즌').font = hdr
mlabels = ['수량','총매출','순매출','평균단가','마진율%','CM평균%']
for (dg, b), c in pair_col1.items():
    ws1.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c+5)
    ws1.cell(row=4, column=c, value=b).font = Font(bold=True, size=8, name='Calibri')
    ws1.cell(row=4, column=c).alignment = Alignment(horizontal='center', wrap_text=True)
    for mi, ml in enumerate(mlabels):
        cell = ws1.cell(row=5, column=c+mi, value=ml)
        cell.font = Font(bold=True, size=7, name='Calibri')
        cell.alignment = Alignment(horizontal='center'); cell.border = thin

r = 6
for season in seasons_list + ['Total']:
    ws1.cell(row=r, column=1, value=season).font = Font(bold=(season=='Total'), size=10, name='Calibri')
    ws1.cell(row=r, column=1).border = thin
    for (dg, b), c in pair_col1.items():
        if season == 'Total':
            sd = [sum(S1[(s, b)][i] for s in seasons_list) for i in range(6)]
        else:
            sd = S1.get((season, b), [0,0,0,0,0,0])
        tq, ta, tn, tdc, tp, tem = sd
        if tq == 0:
            for mi in range(6):
                cell = ws1.cell(row=r, column=c+mi, value=0)
                cell.font = numf; cell.border = thin; cell.alignment = Alignment(horizontal='right')
            continue
        vals = [int(tq), k(ta), k(tn), round(ta/tq, 2) if tq else 0,
                round(tp/ta*100, 2) if ta else 0, round(tem/ta*100, 2) if ta else 0]
        fmts = [FMT_INT, FMT_INT, FMT_INT, FMT_DEC2, FMT_PCT, FMT_PCT]
        for mi, (v, fmt) in enumerate(zip(vals, fmts)):
            cell = ws1.cell(row=r, column=c+mi, value=v)
            cell.number_format = fmt
            cell.font = Font(bold=(season=='Total'), size=10, name='Calibri')
            cell.border = thin; cell.alignment = Alignment(horizontal='right')
    r += 1
print('Sheet 1 done')

# ======================================================
# SHEET 2: 시즌별 (Brand x Season)
# ======================================================
ws2 = wb.create_sheet('2.시즌별')
ws2.cell(row=1, column=1, value='2026 Season별 브랜드별 매출현황 (단위: $K)').font = ttl
ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)

c = 2; col2 = {}
for s in seasons_list:
    ws2.cell(row=3, column=c, value=s).font = hdr
    ws2.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+2)
    ws2.cell(row=4, column=c, value='수량').font = hdr
    ws2.cell(row=4, column=c+1, value='총매출').font = hdr
    ws2.cell(row=4, column=c+2, value='순매출').font = hdr
    col2[s] = (c, c+1, c+2); c += 3
ws2.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+3)
ws2.cell(row=3, column=c, value='TOTAL').font = hdr
ws2.cell(row=4, column=c, value='수량').font = hdr
ws2.cell(row=4, column=c+1, value='총매출').font = hdr
ws2.cell(row=4, column=c+2, value='순매출').font = hdr
ws2.cell(row=4, column=c+3, value='마진율%').font = hdr; tc2 = c
ws2.cell(row=3, column=1, value='브랜드/부서').font = hdr

r = 5
for dn in DEPT_ORDER:
    brands = sorted(dept_brands.get(dn, set()))
    if not brands: continue
    for b in brands:
        ws2.cell(row=r, column=1, value=b)
        for s in seasons_list:
            d = S2.get((b, s), [0,0,0])
            cq, ca, cn_ = col2[s]
            ws2.cell(row=r, column=cq, value=int(d[0])).number_format = FMT_INT
            ws2.cell(row=r, column=ca, value=k(d[1])).number_format = FMT_INT
            ws2.cell(row=r, column=cn_, value=k(d[2])).number_format = FMT_INT
        tq = sum(S2.get((b, s), [0,0,0])[0] for s in seasons_list)
        ta = sum(S2.get((b, s), [0,0,0])[1] for s in seasons_list)
        tn = sum(S2.get((b, s), [0,0,0])[2] for s in seasons_list)
        ws2.cell(row=r, column=tc2, value=int(tq)).number_format = FMT_INT
        ws2.cell(row=r, column=tc2+1, value=k(ta)).number_format = FMT_INT
        ws2.cell(row=r, column=tc2+2, value=k(tn)).number_format = FMT_INT
        ws2.cell(row=r, column=tc2+3, value=round((ta-tn)/ta*100, 2) if ta else 0).number_format = FMT_PCT
        for ci in range(1, tc2+4):
            ws2.cell(row=r, column=ci).border = thin; ws2.cell(row=r, column=ci).font = numf
            ws2.cell(row=r, column=ci).alignment = Alignment(horizontal='right')
        ws2.cell(row=r, column=1).alignment = Alignment(horizontal='left'); ws2.cell(row=r, column=1).font = numf
        r += 1
    # Subtotal
    ws2.cell(row=r, column=1, value=f'{dn} 요약').font = Font(bold=True, size=10, name='Calibri')
    for s in seasons_list:
        dq = sum(S2.get((b, s), [0,0,0])[0] for b in brands)
        da = sum(S2.get((b, s), [0,0,0])[1] for b in brands)
        dn_ = sum(S2.get((b, s), [0,0,0])[2] for b in brands)
        cq, ca, cn_ = col2[s]
        ws2.cell(row=r, column=cq, value=int(dq)).number_format = FMT_INT; ws2.cell(row=r, column=cq).font = Font(bold=True)
        ws2.cell(row=r, column=ca, value=k(da)).number_format = FMT_INT; ws2.cell(row=r, column=ca).font = Font(bold=True)
        ws2.cell(row=r, column=cn_, value=k(dn_)).number_format = FMT_INT; ws2.cell(row=r, column=cn_).font = Font(bold=True)
    tq = sum(sum(S2.get((b, s), [0,0,0])[0] for s in seasons_list) for b in brands)
    ta = sum(sum(S2.get((b, s), [0,0,0])[1] for s in seasons_list) for b in brands)
    tn = sum(sum(S2.get((b, s), [0,0,0])[2] for s in seasons_list) for b in brands)
    ws2.cell(row=r, column=tc2, value=int(tq)).number_format = FMT_INT; ws2.cell(row=r, column=tc2).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+1, value=k(ta)).number_format = FMT_INT; ws2.cell(row=r, column=tc2+1).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+2, value=k(tn)).number_format = FMT_INT; ws2.cell(row=r, column=tc2+2).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+3, value=round((ta-tn)/ta*100, 2) if ta else 0).number_format = FMT_PCT
    ws2.cell(row=r, column=tc2+3).font = Font(bold=True)
    for ci in range(1, tc2+4): ws2.cell(row=r, column=ci).border = thin
    r += 1
print('Sheet 2 done')

# ======================================================
# SHEET 3: 월별 (Month x Brand)
# ======================================================
ws3 = wb.create_sheet('3.월별')
ws3.cell(row=1, column=1, value='2026년 월별 부서별 브랜드별 매출현황 (단위: $K)').font = ttl
ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=80)

pair_col3 = {}; col = 2
for dg in DEPT_ORDER:
    brands = dept_brands.get(dg, [])
    if not brands: continue
    cs = col
    for b in brands:
        pair_col3[(dg, b)] = col; col += 3
    ce = col - 1
    if ce > cs:
        ws3.merge_cells(start_row=3, start_column=cs, end_row=3, end_column=ce)
        ws3.cell(row=3, column=cs, value=dg).font = hdr; ws3.cell(row=3, column=cs).alignment = Alignment(horizontal='center')

ws3.cell(row=4, column=1, value='월').font = hdr
for (dg, b), c in pair_col3.items():
    ws3.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c+2)
    ws3.cell(row=4, column=c, value=b).font = Font(bold=True, size=8, name='Calibri')
    ws3.cell(row=4, column=c).alignment = Alignment(horizontal='center', wrap_text=True)
    for mi, ml in enumerate(['수량','매출','순매출']):
        cell = ws3.cell(row=5, column=c+mi, value=ml)
        cell.font = Font(bold=True, size=7, name='Calibri')
        cell.alignment = Alignment(horizontal='center'); cell.border = thin

r = 6
for mn in months_list:
    ws3.cell(row=r, column=1, value=mn).font = hdr; ws3.cell(row=r, column=1).border = thin
    for (dg, b), c in pair_col3.items():
        sd = S3.get((mn, b), [0,0,0])
        tq, ta, tn = sd
        cell_q = ws3.cell(row=r, column=c, value=int(tq)); cell_q.number_format = FMT_INT
        cell_a = ws3.cell(row=r, column=c+1, value=k(ta)); cell_a.number_format = FMT_INT
        cell_n = ws3.cell(row=r, column=c+2, value=k(tn)); cell_n.number_format = FMT_INT
        for ci in range(c, c+3):
            ws3.cell(row=r, column=ci).font = numf; ws3.cell(row=r, column=ci).border = thin
            ws3.cell(row=r, column=ci).alignment = Alignment(horizontal='right')
    r += 1
print('Sheet 3 done')

# ======================================================
# SHEET 4: 시즌부서공장 (Season > Factory x Dept-Brand)
# ======================================================
ws4 = wb.create_sheet('4.시즌부서공장')
ws4.cell(row=1, column=1, value='2026 Season 시즌별 Factory x 부서-브랜드 배정 수량').font = ttl
ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=80)

ws4.cell(row=3, column=1, value='시즌').font = hdr; ws4.cell(row=3, column=2, value='Factory').font = hdr

pair_col4 = {}; col = 3
for dg in DEPT_ORDER:
    brands = dept_brands.get(dg, [])
    if not brands: continue
    cs = col
    for b in brands:
        if any(S4[(d,b,f,s)] for (d,b2,f,s) in S4 if b2==b for _ in [1]):  # brand has data
            pair_col4[(dg, b)] = col; col += 1
    ce = col - 1
    if ce > cs:
        ws4.merge_cells(start_row=3, start_column=cs, end_row=3, end_column=ce)
        ws4.cell(row=3, column=cs, value=dg).font = hdr; ws4.cell(row=3, column=cs).alignment = Alignment(horizontal='center')

for (dg, b), c in pair_col4.items():
    cell = ws4.cell(row=4, column=c, value=b)
    cell.font = Font(bold=True, size=7, name='Calibri')
    cell.alignment = Alignment(horizontal='center', wrap_text=True, text_rotation=90)
    cell.border = thin

r = 5
for season in seasons_list:
    factories = set()
    for (d, b, f, s), q in S4.items():
        if s == season and q > 0: factories.add(f)
    factories = sorted(factories)
    if not factories: continue
    
    start_r = r
    for fi, factory in enumerate(factories):
        ws4.cell(row=r+fi, column=2, value=factory).font = numf
        ws4.cell(row=r+fi, column=2).border = thin; ws4.cell(row=r+fi, column=1).border = thin
        for (dg, b), c in pair_col4.items():
            tot = 0
            for (d2, b2, f2, s2), q2 in S4.items():
                if b2 == b and f2 == factory and s2 == season and brand_dept(b2) == dg:
                    tot += q2
            cell = ws4.cell(row=r+fi, column=c, value=int(tot) if tot else 0)
            cell.number_format = FMT_INT
            cell.font = numf; cell.border = thin; cell.alignment = Alignment(horizontal='right')
    
    if factories:
        ws4.merge_cells(start_row=start_r, start_column=1, end_row=start_r+len(factories)-1, end_column=1)
        cell = ws4.cell(row=start_r, column=1, value=season)
        cell.font = Font(bold=True, size=10, name='Calibri')
        cell.alignment = Alignment(horizontal='center', vertical='center')
    r += len(factories) + 1

ws4.column_dimensions['A'].width = 12; ws4.column_dimensions['B'].width = 32
print('Sheet 4 done')

# ======================================================
# SHEET 5: 시즌공장 (Factory x Brand x Season)
# ======================================================
ws5 = wb.create_sheet('5.시즌공장')
ws5.cell(row=1, column=1, value='2026 Season 공장별 브랜드별 시즌별 배정 수량 및 비중').font = ttl
ws5.merge_cells(start_row=1, start_column=1, end_row=1, end_column=15)
for ci, h in enumerate(['Factory', '브랜드'], 1): ws5.cell(row=3, column=ci, value=h).font = hdr
col5 = {}; c = 3
for s in seasons_list:
    ws5.cell(row=3, column=c, value=s).font = hdr; ws5.cell(row=3, column=c+1, value='비중').font = hdr
    col5[s] = c; c += 2
ws5.cell(row=3, column=c, value='Total').font = hdr; ws5.cell(row=3, column=c+1, value='비중').font = hdr; tc5 = c

fg = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
for (fac, br, s), q in S5.items():
    fg[fac][br][s] += q

r = 4
for factory in sorted(fg.keys()):
    brands = sorted(fg[factory].keys())
    for bi, brand in enumerate(brands):
        ws5.cell(row=r, column=1, value=factory if bi == 0 else '')
        ws5.cell(row=r, column=2, value=brand)
        fbt = sum(fg[factory][brand].get(s, 0) for s in seasons_list)
        for s in seasons_list:
            q = fg[factory][brand].get(s, 0)
            cell_q = ws5.cell(row=r, column=col5[s], value=int(q))
            cell_q.number_format = FMT_INT
            cell_p = ws5.cell(row=r, column=col5[s]+1, value=round(q/fbt, 4) if fbt else 0)
            cell_p.number_format = FMT_RATIO
        ws5.cell(row=r, column=tc5, value=int(fbt)).number_format = FMT_INT
        ws5.cell(row=r, column=tc5+1, value=1.0).number_format = FMT_RATIO
        for ci in range(1, tc5+2):
            ws5.cell(row=r, column=ci).border = thin; ws5.cell(row=r, column=ci).font = numf
            ws5.cell(row=r, column=ci).alignment = Alignment(horizontal='right')
        ws5.cell(row=r, column=1).alignment = Alignment(horizontal='left')
        ws5.cell(row=r, column=2).alignment = Alignment(horizontal='left')
        r += 1
print('Sheet 5 done')

wb.save(OUT)
print(f'Saved: {OUT}')
import os; print(f'Size: {os.path.getsize(OUT)/1024:.0f} KB')
