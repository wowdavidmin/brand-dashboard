#!/usr/bin/env python3
"""Rebuild Sheet 1 with section-separated layout"""
import openpyxl, os
from openpyxl.styles import Font, Border, Side, Alignment
from collections import defaultdict

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
ENH = os.path.join(BASE, '영업DATA_Enhanced_v1.0_20260629.xlsx')

wb = openpyxl.load_workbook(ENH, read_only=True, data_only=True)
ws = wb['Raw_20260629']
seasons = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']

def to_num(x):
    if x is None: return 0
    if isinstance(x, (int, float)): return x
    try: return float(str(x).replace(',',''))
    except: return 0

S = defaultdict(lambda: [0,0,0,0,0,0,0])

for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row)
    s = str(v[12] or '').strip()
    if s not in set(seasons): continue
    d = str(v[10] or '').strip()
    q = to_num(v[26]); a = to_num(v[27]); n = to_num(v[29])
    dc = to_num(v[44]); p = to_num(v[46]); em = to_num(v[52]); cmt = to_num(v[44])
    if not (q and a): continue
    S[(s, d)] = [S[(s, d)][i]+x for i,x in enumerate([q,a,n,dc,p,em,cmt])]
wb.close()
print(f'Aggregated: {len(S)} entries')

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

FMT='#,##0'; FMT2='#,##0.00'; FMTP='0.00'
thin=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
hdr=Font(bold=True,size=10,name='Calibri')

OUT = os.path.join(BASE, '리포트_5Sheet_최종_20260630.xlsx')

# Build active brands list
active = []
for dn, brs in REF:
    for b in brs:
        if sum(S[(s,b)][0] for s in seasons) > 0:
            active.append((dn, b))

def write_metric_section(ws1, start_row, label, metric_idx, fmt):
    r = start_row
    # Section title
    ws1.cell(row=r, column=1, value=label).font = Font(bold=True, size=12, name='Calibri')
    r += 1
    
    # Dept headers (in REF order)
    col = 2
    for dn, brs in REF:
        these = [(d,b) for (d,b) in active if d == dn]
        if not these: continue
        brand_count = len(these)
        cs = col
        col += brand_count
        ws1.merge_cells(start_row=r, start_column=cs, end_row=r, end_column=col-1)
        ws1.cell(row=r, column=cs, value=dn).font = hdr
        ws1.cell(row=r, column=cs).alignment = Alignment(horizontal='center')
    r += 1
    
    # Brand headers
    ws1.cell(row=r, column=1, value='시즌').font = hdr
    ws1.cell(row=r, column=1).border = thin
    col = 2
    for dn, b in active:
        cell = ws1.cell(row=r, column=col, value=b)
        cell.font = Font(bold=True, size=8, name='Calibri')
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin
        col += 1
    r += 1
    
    # Data rows
    for season in seasons + ['Total']:
        ws1.cell(row=r, column=1, value=season).font = Font(bold=(season=='Total'), size=10, name='Calibri')
        ws1.cell(row=r, column=1).border = thin
        col = 2
        for dn, b in active:
            if season == 'Total':
                vals = [sum(S[(s,b)][i] for s in seasons) for i in range(7)]
            else:
                vals = S.get((season,b), [0]*7)
            
            if metric_idx == 0:  # Qty
                v = int(vals[0])
            elif metric_idx == 1:  # Total Sales ($K)
                v = int(round(vals[1] / 1000))
            elif metric_idx == 2:  # Net Sales ($K)
                v = int(round(vals[2] / 1000))
            elif metric_idx == 3:  # Avg Unit Price
                v = round(vals[1] / vals[0], 2) if vals[0] else 0
            elif metric_idx == 4:  # Margin%
                v = round(vals[4] / vals[1] * 100, 2) if vals[1] else 0
            elif metric_idx == 5:  # CM = CMT/qty
                v = round(vals[6] / vals[0], 2) if vals[0] else 0
            else:
                v = 0
            
            cell = ws1.cell(row=r, column=col, value=v)
            cell.number_format = fmt if metric_idx != 4 else FMTP
            cell.font = Font(bold=(season=='Total'), size=10, name='Calibri')
            cell.border = thin
            cell.alignment = Alignment(horizontal='right')
            col += 1
        r += 1
    return r + 1  # blank separator

wb2 = openpyxl.Workbook()
ws1 = wb2.active
ws1.title = '1.시즌현황'
ws1.cell(row=1, column=1, value='[약진통상] 2026 Season 부서별 브랜드별 현황 (단위: $K)').font = Font(bold=True, size=13, name='Calibri')
ws1.merge_cells('A1:ZZ1')

metrics = [
    ('1. ORDER 수량', 0, FMT),
    ('2. 총매출액', 1, FMT),
    ('3. 순매출액', 2, FMT),
    ('4. 평균단가', 3, FMT2),
    ('5. 마진율', 4, FMTP),
    ('6. CM', 5, FMT2),
]

r = 3
for label, idx, fmt in metrics:
    r = write_metric_section(ws1, r, label, idx, fmt)

wb2.save(OUT)
print(f'Saved: {OUT}')
import os; print(f'Size: {os.path.getsize(OUT)/1024:.0f} KB')
