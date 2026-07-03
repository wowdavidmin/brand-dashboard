#!/usr/bin/env python3
"""Add S.TOTAL columns after each department group in Sheet 1 & 3"""
import openpyxl, os
from openpyxl.styles import Font, Border, Side, Alignment
from collections import defaultdict

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
ENH = os.path.join(BASE, '영업DATA_Enhanced_v1.0_20260629.xlsx')
TEMPLATE = os.path.join(BASE, '리포트_5Sheet_최종_v8_20260630.xlsx')
OUT = os.path.join(BASE, '리포트_5Sheet_최종_v9_20260630.xlsx')

# Sheet 2-5 from template
import shutil
shutil.copy2(TEMPLATE, OUT)

# ===== Read enhanced data for Sheet 1 and 3 =====
print('Reading enhanced data...')
wb_enh = openpyxl.load_workbook(ENH, read_only=True, data_only=True)
ws_enh = wb_enh['Raw_20260629']
seasons = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
months_list = [f'{m}월' for m in range(1, 13)]

def to_num(x):
    if x is None: return 0
    if isinstance(x, (int, float)): return x
    try: return float(str(x).replace(',',''))
    except: return 0

S = defaultdict(lambda: [0,0,0,0,0,0,0])  # (season, division) -> [qty, amt, net, dc, profit, em, cmt]
Sm = defaultdict(lambda: [0,0,0])          # (month, division) -> [qty, amt, net]

for i, row in enumerate(ws_enh.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row)
    s = str(v[12] or '').strip()
    d = str(v[10] or '').strip()
    m = str(v[21] or '').strip()
    q = to_num(v[26]); a = to_num(v[27]); n = to_num(v[29])
    dc = to_num(v[44]); p = to_num(v[46]); em = to_num(v[52]); cmt = to_num(v[44])
    if not (q and a): continue
    if s in seasons: S[(s, d)] = [S[(s, d)][i]+x for i,x in enumerate([q,a,n,dc,p,em,cmt])]
    if m in months_list: Sm[(m, d)] = [Sm[(m, d)][i]+x for i,x in enumerate([q,a,n])]
wb_enh.close()
print(f'  Aggregated: {len(S)} season entries, {len(Sm)} month entries')

# Reference brand structure
REF = [
    ('BU1-영업1부', ['GAP', 'IPSS GAP', 'BRFS', 'Madhappy', 'Mott & Bow']),
    ('BU2-영업1부', ['OLD NAVY', "ON  MEN'S"]),
    ('BU2-영업2부', ['WALMART', 'QUINCE']),
    ('BU3-영업1부', ['BR', 'IPSS ON']),
    ('BU3-영업2부', ["KOHL'S", 'URBAN']),
    ('BU3-영업3부', ['AERIE', 'Offline']),
    ('BU3-영업4부', ['MLB']),
    ('BU4부', ['NUUDS', 'ONER ACTIVE', 'ROSS', 'AEO']),
    ('BU5-영업5부', ['Justice', 'INGRID & ISABEL']),
    ('GTBU-과테말라', ['OLDNAVY(HYBRID)']),
]

# Active brands
active = []
for dn, brs in REF:
    for b in brs:
        if sum(S[(s,b)][0] for s in seasons) > 0:
            active.append((dn, b))

FMT = '#,##0'; FMT2 = '#,##0.00'; FMTP = '0.0'
thin = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
hdr = Font(bold=True, size=10, name='Calibri')

wb = openpyxl.load_workbook(OUT)

# ===== SHEET 1: with S.TOTAL =====
print('Building Sheet 1 with S.TOTAL...')
ws1 = wb['1.시즌현황']
del wb['1.시즌현황']
ws1 = wb.create_sheet('1.시즌현황', 0)

ws1.cell(row=1, column=1, value='[약진통상] 2026 Season 부서별 브랜드별 현황 (단위: $K)').font = Font(bold=True, size=13, name='Calibri')
ws1.merge_cells('A1:ZZ1')

# Build column structure with S.TOTAL
# active_with_total: list of (dept, brand_or_total) for columns
active_wt = []
for dn in [d for d, brs in REF]:
    these = [(d, b) for b in [br for br in [brs for d2, brs in REF if d2 == dn][0]] if sum(S[(s,b)][0] for s in seasons) > 0]
    if not these: continue
    active_wt.extend(these)
    active_wt.append((dn, 'S.TOTAL'))

def write_section(ws1, sr, label, midx, fmt, is_pct=False):
    r = sr
    ws1.cell(row=r, column=1, value=label).font = Font(bold=True, size=12, name='Calibri')
    r += 1
    # Dept headers
    col = 2
    for dn, brs in REF:
        these = [(d, b) for (d, b) in active_wt if d == dn]
        if not these: continue
        cs = col
        for _ in these: col += 1
        if col-1 > cs:
            ws1.merge_cells(start_row=r, start_column=cs, end_row=r, end_column=col-1)
            ws1.cell(row=r, column=cs, value=dn).font = hdr
            ws1.cell(row=r, column=cs).alignment = Alignment(horizontal='center')
    r += 1
    # Brand headers
    ws1.cell(row=r, column=1, value='시즌').font = hdr; ws1.cell(row=r, column=1).border = thin
    col = 2
    for dn, item in active_wt:
        cell = ws1.cell(row=r, column=col, value=item)
        if item == 'S.TOTAL':
            cell.font = Font(bold=True, size=8, name='Calibri', color='333333')
        else:
            cell.font = Font(bold=True, size=8, name='Calibri')
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin; col += 1
    r += 1
    # Data
    for season in seasons + ['Total']:
        ws1.cell(row=r, column=1, value=season).font = Font(bold=(season=='Total'), size=10, name='Calibri')
        ws1.cell(row=r, column=1).border = thin; col = 2
        for dn, item in active_wt:
            if item == 'S.TOTAL':
                dept_brands = [b for (d, b) in active_wt if d == dn and b != 'S.TOTAL']
                # Precompute brand totals
                brand_data = {}
                for b in dept_brands:
                    if season == 'Total':
                        brand_data[b] = [sum(S[(s, b)][i] for s in seasons) for i in range(7)]
                    else:
                        brand_data[b] = S.get((season, b), [0]*7)
                
                tq = sum(bd[0] for bd in brand_data.values())
                ta = sum(bd[1] for bd in brand_data.values())
                tn = sum(bd[2] for bd in brand_data.values())
                tp = sum(bd[4] for bd in brand_data.values())
                tcmt = sum(bd[6] for bd in brand_data.values())
                
                if midx == 0: v = int(tq)
                elif midx == 1: v = int(round(ta/1000))
                elif midx == 2: v = int(round(tn/1000))
                elif midx == 3: v = round(ta/tq, 2) if tq else 0
                elif midx == 4: v = round(tp/ta*100, 1) if ta else 0
                elif midx == 5: v = round(tcmt/tq, 2) if tq else 0
            else:
                b = item
                if season == 'Total':
                    sv = [sum(S[(s, b)][i] for s in seasons) for i in range(7)]
                else:
                    sv = S.get((season, b), [0]*7)
                if midx == 0: v = int(sv[0])
                elif midx == 1: v = int(round(sv[1]/1000))
                elif midx == 2: v = int(round(sv[2]/1000))
                elif midx == 3: v = round(sv[1]/sv[0], 2) if sv[0] else 0
                elif midx == 4: v = round(sv[4]/sv[1]*100, 1) if sv[1] else 0
                elif midx == 5: v = round(sv[6]/sv[0], 2) if sv[0] else 0
            
            cell = ws1.cell(row=r, column=col, value=v if v else '-')
            if isinstance(v, (int, float)) and v != '-':
                if item == 'S.TOTAL': cell.font = Font(bold=True, size=10, name='Calibri')
                else: cell.font = Font(size=10, name='Calibri')
                if midx == 4: cell.number_format = FMTP
                elif midx in (3, 5): cell.number_format = FMT2
                else: cell.number_format = FMT
            else:
                cell.font = Font(size=10, name='Calibri')
            cell.border = thin; cell.alignment = Alignment(horizontal='right'); col += 1
        r += 1
    return r + 1

r = 3
for label, midx, fmt in [
    ('1. ORDER 수량', 0, FMT), ('2. 총매출액', 1, FMT), ('3. 순매출액', 2, FMT),
    ('4. 평균단가', 3, FMT2), ('5. 마진율', 4, FMTP), ('6. CM', 5, FMT2),
]:
    r = write_section(ws1, r, label, midx, fmt)
print('  Sheet 1 done')

# ===== SHEET 3: with S.TOTAL =====
print('Building Sheet 3 with S.TOTAL...')
del wb['3.월별']
ws3 = wb.create_sheet('3.월별')

ws3.cell(row=1, column=1, value='2026년 월별 부서별 브랜드별 매출현황 (단위: $K)').font = Font(bold=True, size=13, name='Calibri')
ws3.merge_cells('A1:ZZ1')

col = 2
mc = {}
for mn in months_list:
    ws3.cell(row=3, column=col, value=mn).font = hdr
    ws3.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col+2)
    ws3.cell(row=4, column=col, value='수량').font = hdr
    ws3.cell(row=4, column=col+1, value='매출').font = hdr
    ws3.cell(row=4, column=col+2, value='순매출').font = hdr
    mc[mn] = col; col += 3

ws3.cell(row=3, column=1, value='브랜드/부서').font = hdr

r = 5
for dn in [d for d, brs in REF]:
    these = [(d, b) for b in [br for br in [brs for d2, brs in REF if d2 == dn][0]] if sum(Sm.get((mn, b), [0]*3)[0] for mn in months_list) > 0]
    if not these: continue
    brands = [b for (d, b) in these]
    
    for b in brands:
        ws3.cell(row=r, column=1, value=b)
        for mn in months_list:
            d = Sm.get((mn, b), [0, 0, 0])
            for ci, v in enumerate([int(d[0]), int(round(d[1]/1000)), int(round(d[2]/1000))]):
                cell = ws3.cell(row=r, column=mc[mn]+ci, value=v if v else '-')
                cell.number_format = FMT
                cell.font = Font(size=10); cell.border = thin; cell.alignment = Alignment(horizontal='right')
        ws3.cell(row=r, column=1).font = Font(size=10)
        ws3.cell(row=r, column=1).alignment = Alignment(horizontal='left')
        r += 1
    
    # S.TOTAL row
    ws3.cell(row=r, column=1, value='S.TOTAL').font = Font(bold=True, size=10, name='Calibri')
    for mn in months_list:
        dq = sum(int(Sm.get((mn, b), [0]*3)[0]) for b in brands)
        da = sum(int(round(Sm.get((mn, b), [0]*3)[1]/1000)) for b in brands)
        dn_ = sum(int(round(Sm.get((mn, b), [0]*3)[2]/1000)) for b in brands)
        for ci, v in enumerate([dq, da, dn_]):
            cell = ws3.cell(row=r, column=mc[mn]+ci, value=v if v else '-')
            cell.number_format = FMT
            cell.font = Font(bold=True, size=10, name='Calibri')
            cell.border = thin; cell.alignment = Alignment(horizontal='right')
    r += 1

print('  Sheet 3 done')

# Reorder sheets
order = ['1.시즌현황','2.시즌별','3.월별','4.시즌부서공장','5.시즌공장']
for i, name in enumerate(order):
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        if idx != i: wb.move_sheet(name, offset=i-idx)

wb.save(OUT)
print(f'Saved: {OUT} ({os.path.getsize(OUT)/1024:.0f} KB)')
wb.close()
