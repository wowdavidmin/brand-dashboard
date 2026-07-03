#!/usr/bin/env python3
"""Rebuild report matching reference brand structure exactly"""
import openpyxl, os, glob
from openpyxl.styles import Font, Border, Side, Alignment
from collections import defaultdict, OrderedDict
from datetime import datetime

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
files = [f for f in glob.glob(os.path.join(BASE, '2)*')) if not os.path.basename(f).startswith('~$')]
F2 = max(files, key=os.path.getmtime)
OUT = os.path.join(BASE, f'영업DATA_5Sheet_FINAL_{datetime.now().strftime("%Y%m%d")}.xlsx')

print(f'Reading: {F2}')
wb2 = openpyxl.load_workbook(F2, read_only=True, data_only=True)
ws = wb2['Raw_20260608']

seasons = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
months = [f'{m}월' for m in range(1, 13)]

# Aggregations
S1 = defaultdict(lambda: [0,0,0,0,0,0])  # (season, brand) -> [qty, amt, net, dc, profit, em]
S2 = defaultdict(lambda: [0,0,0])         # (brand, season)
S3 = defaultdict(lambda: [0,0,0])         # (month, brand)
S4 = defaultdict(lambda: 0)               # (dept, brand, factory, season)
S5 = defaultdict(lambda: 0)               # (factory, brand, season)

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
    if not (q and a and season in seasons): continue
    S1[(season, brand)] = [S1[(season, brand)][i] + x for i, x in enumerate([q, a, n, dc, prof, em])]
    S2[(brand, season)] = [S2[(brand, season)][i] + x for i, x in enumerate([q, a, n])]
    if month: S3[(month, brand)] = [S3[(month, brand)][i] + x for i, x in enumerate([q, a, n])]
    S4[(dept, brand, factory, season)] += q
    S5[(factory, brand, season)] += q
    if cnt % 30000 == 0: print(f'  {cnt}...')
wb2.close()
print(f'Aggregated: {cnt} rows')

# ===== BRAND TO DEPARTMENT MAPPING (matching reference structure) =====
# Reference structure:
# BU1: GAP, IPSS GAP, BRFS, Madhappy, Mott&Bow
# BU2-1: OLD NAVY, ON MEN'S
# BU2-2: WALMART, QUINCE
# BU3-1: BR, IPSS ON
# BU3-2: KOHL'S, URBAN
# BU3-3: AERIE, Offline
# BU3-4: MLB
# BU4: NUUDS, ONER ACTIVE, ROSS, AEO (=AMERICAN EAGLE)
# BU5: Justice, INGRID & ISABEL
# GTBU: OLDNAVY(HYBRID)
# Extra from file #2: TOMS

# File #2 brands mapped to reference brand names:
# "Gap" + "Gap OUTLET" → "GAP" (merge)
# "ON IPSS" → "IPSS ON" (BU3-1) 
# "WAL-MART" → "WALMART" (BU2-2)
# "BANANA REPUBLIC" → "BR" (BU3-1)
# "AMERICAN EAGLE" → "AEO" (BU4)
# "KOHL`S" → "KOHL'S" (BU3-2)
# "URBAN OUTFITTERS" → "URBAN" (BU3-2)

# Reference brand list in order
REF_BRANDS = [
    ('BU1-영업1부', ['GAP', 'IPSS GAP', 'BRFS', 'Madhappy', 'Mott&Bow', 'Gap OUTLET']),
    ('BU2-영업1부', ['OLD NAVY', "ON MEN'S"]),
    ('BU2-영업2부', ['WALMART', 'QUINCE']),
    ('BU3-영업1부', ['BR', 'IPSS ON']),
    ('BU3-영업2부', ["KOHL'S", 'URBAN']),
    ('BU3-영업3부', ['AERIE', 'Offline']),
    ('BU3-영업4부', ['MLB']),
    ('BU4부', ['NUUDS', 'ONER ACTIVE', 'ROSS', 'AEO']),
    ('BU5-영업5부', ['Justice', 'INGRID & ISABEL']),
    ('GTBU-과테말라', ['OLDNAVY(HYBRID)']),
]

# Map file #2 brand names -> reference brand names
BRAND_ALIAS = {
    'Gap': 'GAP',
    'Gap OUTLET': 'Gap OUTLET',
    'BRFS': 'BRFS',
    'Madhappy': 'Madhappy',
    'Mott & Bow': 'Mott&Bow',
    'OLD NAVY': 'OLD NAVY',
    'ON IPSS': "ON MEN'S",
    'WAL-MART': 'WALMART',
    'QUINCE': 'QUINCE',
    'BANANA REPUBLIC': 'BR',
    "KOHL`S": "KOHL'S",
    'URBAN OUTFITTERS': 'URBAN',
    'AERIE': 'AERIE',
    'Offline': 'Offline',
    'MLB': 'MLB',
    'NUUDS': 'NUUDS',
    'ONER ACTIVE': 'ONER ACTIVE',
    'ROSS': 'ROSS',
    'AMERICAN EAGLE': 'AEO',
    'JUSTICE': 'Justice',
    'INGRID & ISABEL': 'INGRID & ISABEL',
    'OLDNAVY(HYBRID)': 'OLDNAVY(HYBRID)',
    'TOMS': 'TOMS',  # extra brand from file #2
}

# Build S1, S2, S3 aggregated by reference brand name
S1_ref = defaultdict(lambda: [0,0,0,0,0,0])
S2_ref = defaultdict(lambda: [0,0,0])
S3_ref = defaultdict(lambda: [0,0,0])

for (season, brand), d in S1.items():
    ref_name = BRAND_ALIAS.get(brand, brand)
    for i in range(6): S1_ref[(season, ref_name)][i] += d[i]

for (brand, season), d in S2.items():
    ref_name = BRAND_ALIAS.get(brand, brand)
    for i in range(3): S2_ref[(ref_name, season)][i] += d[i]

for (month, brand), d in S3.items():
    ref_name = BRAND_ALIAS.get(brand, brand)
    for i in range(3): S3_ref[(ref_name, month)][i] += d[i]

# Rebuild S4/S5 too
S4_ref = defaultdict(lambda: 0)
S5_ref = defaultdict(lambda: 0)
for (dept, brand, factory, season), q in S4.items():
    ref_name = BRAND_ALIAS.get(brand, brand)
    S4_ref[(dept, ref_name, factory, season)] += q
for (factory, brand, season), q in S5.items():
    ref_name = BRAND_ALIAS.get(brand, brand)
    S5_ref[(factory, ref_name, season)] += q

# Collect all reference brands that have data
all_ref_brands = set()
for (season, b) in S1_ref: all_ref_brands.add(b)

print(f'Reference brands with data: {len(all_ref_brands)}')
for dn, brands in REF_BRANDS:
    for b in brands:
        has = 'Y' if b in all_ref_brands else '-'
        print(f'  {dn:15s} {b:20s} {has}')

# Also add TOMS to BU4
REF_BRANDS_extra = list(REF_BRANDS)
# Add TOMS to BU4
for i, (dn, brands) in enumerate(REF_BRANDS_extra):
    if dn == 'BU4부':
        REF_BRANDS_extra[i] = (dn, brands + ['TOMS'])
        break

# Styling
FMT_INT = '#,##0'; FMT_DEC2 = '#,##0.00'; FMT_PCT = '0.00'; FMT_RATIO = '0.0000'
thin = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
hdr = Font(bold=True, size=10, name='Calibri')
ttl = Font(bold=True, size=13, name='Calibri')
numf = Font(size=10, name='Calibri')
def k(v): return int(round(v / 1000))

wb = openpyxl.Workbook()

# =============================================
# SHEET 1: 시즌현황
# =============================================
ws1 = wb.active; ws1.title = '1.시즌현황'
ws1.cell(row=1, column=1, value='[약진통상] 2026 Season 부서별 브랜드별 현황 (단위: $K)').font = ttl
ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=80)

pair1 = {}; col = 2
for dn, brands in REF_BRANDS_extra:
    if not brands: continue
    cs = col
    for b in brands:
        pair1[(dn, b)] = col; col += 6
    if col-1 > cs:
        ws1.merge_cells(start_row=3, start_column=cs, end_row=3, end_column=col-1)
        ws1.cell(row=3, column=cs, value=dn).font = hdr; ws1.cell(row=3, column=cs).alignment = Alignment(horizontal='center')

ws1.cell(row=4, column=1, value='시즌').font = hdr
ml = ['수량','총매출','순매출','평균단가','마진율%','CM%']
for (dn, b), c in pair1.items():
    ws1.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c+5)
    ws1.cell(row=4, column=c, value=b).font = Font(bold=True, size=8, name='Calibri')
    ws1.cell(row=4, column=c).alignment = Alignment(horizontal='center', wrap_text=True)
    for mi, m in enumerate(ml):
        cell = ws1.cell(row=5, column=c+mi, value=m)
        cell.font = Font(bold=True, size=7, name='Calibri'); cell.alignment = Alignment(horizontal='center'); cell.border = thin

r = 6
for season in seasons + ['Total']:
    ws1.cell(row=r, column=1, value=season).font = Font(bold=(season=='Total'), size=10, name='Calibri')
    ws1.cell(row=r, column=1).border = thin
    for (dn, b), c in pair1.items():
        if season == 'Total':
            sd = [sum(S1_ref[(s, b)][i] for s in seasons) for i in range(6)]
        else:
            sd = S1_ref.get((season, b), [0,0,0,0,0,0])
        tq, ta, tn, tdc, tp, tem = sd
        if not tq:
            for mi in range(6):
                cell = ws1.cell(row=r, column=c+mi, value=0)
                cell.font = numf; cell.border = thin; cell.alignment = Alignment(horizontal='right')
            continue
        vals = [int(tq), k(ta), k(tn),
                round(ta/tq, 2) if tq else 0,
                round(tp/ta*100, 2) if ta else 0,
                round(tem/ta*100, 2) if ta else 0]
        fmts = [FMT_INT, FMT_INT, FMT_INT, FMT_DEC2, FMT_PCT, FMT_PCT]
        for mi, (v, fmt) in enumerate(zip(vals, fmts)):
            cell = ws1.cell(row=r, column=c+mi, value=v)
            cell.number_format = fmt
            cell.font = Font(bold=(season=='Total'), size=10, name='Calibri')
            cell.border = thin; cell.alignment = Alignment(horizontal='right')
    r += 1

# Add NOTE about IPSS GAP
ws1.cell(row=r+1, column=1, value='* IPSS GAP: 파일 #2에 별도 브랜드 없음. Gap OUTLET(${3.53}/unit)은 별도 할인채널.').font = Font(italic=True, size=9, color='666666')
ws1.merge_cells(start_row=r+1, start_column=1, end_row=r+1, end_column=30)
print('Sheet 1 done')

# =============================================
# SHEET 2: 시즌별
# =============================================
ws2 = wb.create_sheet('2.시즌별')
ws2.cell(row=1, column=1, value='2026 Season별 브랜드별 매출현황 (단위: $K)').font = ttl
ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)
ws2.cell(row=3, column=1, value='브랜드/부서').font = hdr

c = 2; col2 = {}
for s in seasons:
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

r = 5
for dn, brands in REF_BRANDS_extra:
    for b in brands:
        ws2.cell(row=r, column=1, value=b)
        for s in seasons:
            d = S2_ref.get((b, s), [0,0,0])
            cq, ca, cn_ = col2[s]
            ws2.cell(row=r, column=cq, value=int(d[0])).number_format = FMT_INT
            ws2.cell(row=r, column=ca, value=k(d[1])).number_format = FMT_INT
            ws2.cell(row=r, column=cn_, value=k(d[2])).number_format = FMT_INT
        tq = sum(S2_ref.get((b, s), [0,0,0])[0] for s in seasons)
        ta = sum(S2_ref.get((b, s), [0,0,0])[1] for s in seasons)
        tn = sum(S2_ref.get((b, s), [0,0,0])[2] for s in seasons)
        ws2.cell(row=r, column=tc2, value=int(tq)).number_format = FMT_INT
        ws2.cell(row=r, column=tc2+1, value=k(ta)).number_format = FMT_INT
        ws2.cell(row=r, column=tc2+2, value=k(tn)).number_format = FMT_INT
        ws2.cell(row=r, column=tc2+3, value=round((ta-tn)/ta*100, 2) if ta else 0).number_format = FMT_PCT
        for ci in range(1, tc2+4):
            ws2.cell(row=r, column=ci).border = thin; ws2.cell(row=r, column=ci).font = numf
            ws2.cell(row=r, column=ci).alignment = Alignment(horizontal='right')
        ws2.cell(row=r, column=1).alignment = Alignment(horizontal='left')
        r += 1
    # Subtotal
    ws2.cell(row=r, column=1, value=f'{dn} 요약').font = Font(bold=True)
    for s in seasons:
        dq = sum(S2_ref.get((b, s), [0,0,0])[0] for b in brands)
        da = sum(S2_ref.get((b, s), [0,0,0])[1] for b in brands)
        dn_ = sum(S2_ref.get((b, s), [0,0,0])[2] for b in brands)
        cq, ca, cn_ = col2[s]
        ws2.cell(row=r, column=cq, value=int(dq)).font = Font(bold=True)
        ws2.cell(row=r, column=ca, value=k(da)).font = Font(bold=True)
        ws2.cell(row=r, column=cn_, value=k(dn_)).font = Font(bold=True)
    tq = sum(sum(S2_ref.get((b, s), [0,0,0])[0] for s in seasons) for b in brands)
    ta = sum(sum(S2_ref.get((b, s), [0,0,0])[1] for s in seasons) for b in brands)
    tn = sum(sum(S2_ref.get((b, s), [0,0,0])[2] for s in seasons) for b in brands)
    ws2.cell(row=r, column=tc2, value=int(tq)).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+1, value=k(ta)).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+2, value=k(tn)).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+3, value=round((ta-tn)/ta*100, 2) if ta else 0).font = Font(bold=True)
    for ci in range(1, tc2+4): ws2.cell(row=r, column=ci).border = thin
    r += 1
print('Sheet 2 done')

# =============================================
# SHEET 3: 월별
# =============================================
ws3 = wb.create_sheet('3.월별')
ws3.cell(row=1, column=1, value='2026년 월별 부서별 브랜드별 매출현황 (단위: $K)').font = ttl
ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=80)

pair3 = {}; col = 2
for dn, brands in REF_BRANDS_extra:
    if not brands: continue
    cs = col
    for b in brands:
        pair3[(dn, b)] = col; col += 3
    if col-1 > cs:
        ws3.merge_cells(start_row=3, start_column=cs, end_row=3, end_column=col-1)
        ws3.cell(row=3, column=cs, value=dn).font = hdr; ws3.cell(row=3, column=cs).alignment = Alignment(horizontal='center')

ws3.cell(row=4, column=1, value='월').font = hdr
for (dn, b), c in pair3.items():
    ws3.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c+2)
    ws3.cell(row=4, column=c, value=b).font = Font(bold=True, size=8, name='Calibri')
    ws3.cell(row=4, column=c).alignment = Alignment(horizontal='center', wrap_text=True)
    for mi, m in enumerate(['수량','매출','순매출']):
        cell = ws3.cell(row=5, column=c+mi, value=m)
        cell.font = Font(bold=True, size=7, name='Calibri'); cell.alignment = Alignment(horizontal='center'); cell.border = thin

r = 6
for mn in months:
    ws3.cell(row=r, column=1, value=mn).font = hdr; ws3.cell(row=r, column=1).border = thin
    for (dn, b), c in pair3.items():
        sd = S3_ref.get((b, mn), [0,0,0])
        vals = [int(sd[0]), k(sd[1]), k(sd[2])]
        for mi, v in enumerate(vals):
            cell = ws3.cell(row=r, column=c+mi, value=v)
            cell.number_format = FMT_INT; cell.font = numf; cell.border = thin; cell.alignment = Alignment(horizontal='right')
    r += 1

# Total
ws3.cell(row=r, column=1, value='총합계').font = Font(bold=True)
for (dn, b), c in pair3.items():
    dq = sum(S3_ref.get((b, mn), [0,0,0])[0] for mn in months)
    da = sum(S3_ref.get((b, mn), [0,0,0])[1] for mn in months)
    dn_ = sum(S3_ref.get((b, mn), [0,0,0])[2] for mn in months)
    for mi, v in enumerate([int(dq), k(da), k(dn_)]):
        cell = ws3.cell(row=r, column=c+mi, value=v)
        cell.number_format = FMT_INT; cell.font = Font(bold=True); cell.border = thin; cell.alignment = Alignment(horizontal='right')
print('Sheet 3 done')

# =============================================
# SHEET 4: 시즌부서공장
# =============================================
ws4 = wb.create_sheet('4.시즌부서공장')
ws4.cell(row=1, column=1, value='2026 Season 시즌별 Factory x 부서-브랜드 배정 수량').font = ttl
ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=80)
ws4.cell(row=3, column=1, value='시즌').font = hdr; ws4.cell(row=3, column=2, value='Factory').font = hdr

pair4 = {}; col = 3
for dn, brands in REF_BRANDS_extra:
    active = [b for b in brands if any(S4_ref[(d, b, f, s)] for (d, b2, f, s) in S4_ref if b2 == b)]
    if not active: continue
    cs = col
    for b in active:
        pair4[(dn, b)] = col; col += 1
    if col-1 > cs:
        ws4.merge_cells(start_row=3, start_column=cs, end_row=3, end_column=col-1)
        ws4.cell(row=3, column=cs, value=dn).font = hdr; ws4.cell(row=3, column=cs).alignment = Alignment(horizontal='center')

for (dn, b), c in pair4.items():
    cell = ws4.cell(row=4, column=c, value=b)
    cell.font = Font(bold=True, size=7, name='Calibri')
    cell.alignment = Alignment(horizontal='center', wrap_text=True, text_rotation=90); cell.border = thin

r = 5
for season in seasons:
    factories = set()
    for (d, b, f, s), q in S4_ref.items():
        if s == season and q > 0: factories.add(f)
    factories = sorted(factories)
    if not factories: continue
    start_r = r
    for fi, factory in enumerate(factories):
        ws4.cell(row=r+fi, column=2, value=factory).font = numf
        ws4.cell(row=r+fi, column=2).border = thin; ws4.cell(row=r+fi, column=1).border = thin
        for (dn, b), c in pair4.items():
            tot = sum(q for (d, b2, f, s), q in S4_ref.items() if b2 == b and f == factory and s == season)
            cell = ws4.cell(row=r+fi, column=c, value=int(tot) if tot else 0)
            cell.number_format = FMT_INT; cell.font = numf; cell.border = thin; cell.alignment = Alignment(horizontal='right')
    if factories:
        ws4.merge_cells(start_row=start_r, start_column=1, end_row=start_r+len(factories)-1, end_column=1)
        cell = ws4.cell(row=start_r, column=1, value=season)
        cell.font = Font(bold=True, size=10, name='Calibri'); cell.alignment = Alignment(horizontal='center', vertical='center')
    r += len(factories) + 1
ws4.column_dimensions['A'].width = 12; ws4.column_dimensions['B'].width = 32
print('Sheet 4 done')

# =============================================
# SHEET 5: 시즌공장
# =============================================
ws5 = wb.create_sheet('5.시즌공장')
ws5.cell(row=1, column=1, value='2026 Season 공장별 브랜드별 시즌별 배정 수량 및 비중').font = ttl
ws5.merge_cells(start_row=1, start_column=1, end_row=1, end_column=15)
for ci, h in enumerate(['Factory', '브랜드'], 1): ws5.cell(row=3, column=ci, value=h).font = hdr
col5 = {}; c = 3
for s in seasons:
    ws5.cell(row=3, column=c, value=s).font = hdr; ws5.cell(row=3, column=c+1, value='비중').font = hdr
    col5[s] = c; c += 2
ws5.cell(row=3, column=c, value='Total').font = hdr; ws5.cell(row=3, column=c+1, value='비중').font = hdr; tc5 = c

fg = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
for (fac, br, s), q in S5_ref.items():
    fg[fac][br][s] += q

r = 4
for factory in sorted(fg.keys()):
    brands = sorted(fg[factory].keys())
    for bi, brand in enumerate(brands):
        ws5.cell(row=r, column=1, value=factory if bi == 0 else '')
        ws5.cell(row=r, column=2, value=brand)
        fbt = sum(fg[factory][brand].get(s, 0) for s in seasons)
        for s in seasons:
            q = fg[factory][brand].get(s, 0)
            ws5.cell(row=r, column=col5[s], value=int(q)).number_format = FMT_INT
            ws5.cell(row=r, column=col5[s]+1, value=round(q/fbt, 4) if fbt else 0).number_format = FMT_RATIO
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
print(f'\nSaved: {OUT}')
import os; print(f'Size: {os.path.getsize(OUT)/1024:.0f} KB')
