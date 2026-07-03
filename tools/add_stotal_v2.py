#!/usr/bin/env python3
"""Rebuild with S.TOTAL - simpler approach"""
import openpyxl, os, shutil
from openpyxl.styles import Font, Border, Side, Alignment
from collections import defaultdict

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
ENH = os.path.join(BASE, '영업DATA_Enhanced_v1.0_20260629.xlsx')
TEMPLATE = os.path.join(BASE, '리포트_5Sheet_최종_v8_20260630.xlsx')
OUT = os.path.join(BASE, '리포트_5Sheet_최종_v10_20260630.xlsx')
shutil.copy2(TEMPLATE, OUT)

# Read enhanced data
print('Reading data...')
wb_enh = openpyxl.load_workbook(ENH, read_only=True, data_only=True)
ws_enh = wb_enh['Raw_20260629']
seasons = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
months_list = [f'{m}월' for m in range(1, 13)]

def to_num(x):
    if x is None: return 0
    if isinstance(x, (int, float)): return x
    try: return float(str(x).replace(',',''))
    except: return 0

S = defaultdict(lambda: [0]*7)
Sm = defaultdict(lambda: [0]*3)
for i, row in enumerate(ws_enh.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row)
    s = str(v[12] or '').strip()
    d = str(v[10] or '').strip()
    m = str(v[21] or '').strip()
    q = to_num(v[26]); a = to_num(v[27]); n = to_num(v[29])
    dc = to_num(v[44]); p = to_num(v[46]); em = to_num(v[52]); cmt = to_num(v[44])
    if not (q and a): continue
    if s in seasons:
        x = S[(s, d)]; S[(s, d)] = [x[i]+v for i,v in enumerate([q,a,n,dc,p,em,cmt])]
    if m in months_list:
        x = Sm[(m, d)]; Sm[(m, d)] = [x[i]+v for i,v in enumerate([q,a,n])]
wb_enh.close()

# Pre-compute totals per brand
def brand_total(b):
    return [sum(S[(s,b)][i] for s in seasons) for i in range(7)]

REF = [
    ('BU1-영업1부', ['GAP','IPSS GAP','BRFS','Madhappy','Mott & Bow']),
    ('BU2-영업1부', ['OLD NAVY',"ON  MEN'S"]),
    ('BU2-영업2부', ['WALMART','QUINCE']),
    ('BU3-영업1부', ['BR','IPSS ON']),
    ('BU3-영업2부', ["KOHL'S",'URBAN']),
    ('BU3-영업3부', ['AERIE','Offline']),
    ('BU3-영업4부', ['MLB']),
    ('BU4부', ['NUUDS','ONER ACTIVE','ROSS','AEO']),
    ('BU5-영업5부', ['Justice','INGRID & ISABEL']),
    ('GTBU-과테말라', ['OLDNAVY(HYBRID)']),
]

thin=Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
hdr=Font(bold=True,size=10,name='Calibri')

wb = openpyxl.load_workbook(OUT)

# ===== SHEET 1 =====
print('Sheet 1...')
del wb['1.시즌현황']
ws1 = wb.create_sheet('1.시즌현황', 0)
ws1.cell(row=1,column=1,value='[약진통상] 2026 Season 부서별 브랜드별 현황 (단위: $K)').font=Font(bold=True,size=13,name='Calibri')
ws1.merge_cells('A1:ZZ1')

# Build column list with S.TOTAL
cols = []
for dn, brs in REF:
    active = [b for b in brs if sum(S[(s,b)][0] for s in seasons) > 0]
    if not active: continue
    for b in active: cols.append((dn, b, False))
    cols.append((dn, 'S.TOTAL', True))

def sec1(ws1, sr, label, midx):
    r = sr
    ws1.cell(row=r,column=1,value=label).font=Font(bold=True,size=12,name='Calibri'); r+=1
    # Dept headers
    col = 2; prev_dn = None
    for dn, item, is_total in cols:
        if dn != prev_dn:
            if prev_dn is not None and col_start < col-1:
                ws1.merge_cells(start_row=r,start_column=col_start,end_row=r,end_column=col-1)
                ws1.cell(row=r,column=col_start,value=prev_dn).font=hdr
                ws1.cell(row=r,column=col_start).alignment=Alignment(horizontal='center')
            col_start = col; prev_dn = dn
        col += 1
    if col_start < col-1:
        ws1.merge_cells(start_row=r,start_column=col_start,end_row=r,end_column=col-1)
        ws1.cell(row=r,column=col_start,value=prev_dn).font=hdr
        ws1.cell(row=r,column=col_start).alignment=Alignment(horizontal='center')
    r += 1
    # Brand headers
    ws1.cell(row=r,column=1,value='시즌').font=hdr; ws1.cell(row=r,column=1).border=thin
    col=2
    for dn, item, ist in cols:
        cell=ws1.cell(row=r,column=col,value=item)
        cell.font=Font(bold=True,size=8,name='Calibri'); cell.alignment=Alignment(horizontal='center',wrap_text=True); cell.border=thin; col+=1
    r += 1
    # Data
    for season in seasons + ['Total']:
        ws1.cell(row=r,column=1,value=season).font=Font(bold=(season=='Total'),size=10,name='Calibri')
        ws1.cell(row=r,column=1).border=thin; col=2
        for dn, item, ist in cols:
            vals = [0]*7
            if ist:  # S.TOTAL
                dept_brs = [b for (d2, b2, is2) in cols if d2 == dn and not is2]
                for b in dept_brs:
                    sv = brand_total(b) if season == 'Total' else S.get((season, b), [0]*7)
                    for i in range(7): vals[i] += sv[i]
            else:
                vals = brand_total(item) if season == 'Total' else S.get((season, item), [0]*7)
            
            if midx == 0: v = int(vals[0])
            elif midx == 1: v = int(round(vals[1]/1000))
            elif midx == 2: v = int(round(vals[2]/1000))
            elif midx == 3: v = round(vals[1]/vals[0], 2) if vals[0] else 0
            elif midx == 4: v = round(vals[4]/vals[1]*100, 1) if vals[1] else 0
            elif midx == 5: v = round(vals[6]/vals[0], 2) if vals[0] else 0
            
            cell = ws1.cell(row=r, column=col, value=v if v else '-')
            if isinstance(v, (int, float)):
                cell.font = Font(bold=(ist or season=='Total'), size=10, name='Calibri')
                if midx in (3,5): cell.number_format = '#,##0.00'
                elif midx == 4: cell.number_format = '0.0'
                else: cell.number_format = '#,##0'
            else:
                cell.font = Font(size=10, name='Calibri')
            cell.border = thin; cell.alignment = Alignment(horizontal='right'); col += 1
        r += 1
    return r + 1

r = 3
for label, midx in [('1. ORDER 수량',0),('2. 총매출액',1),('3. 순매출액',2),('4. 평균단가',3),('5. 마진율',4),('6. CM',5)]:
    r = sec1(ws1, r, label, midx)
print('  Done')

# ===== SHEET 3 =====
print('Sheet 3...')
del wb['3.월별']
ws3 = wb.create_sheet('3.월별')
ws3.cell(row=1,column=1,value='2026년 월별 부서별 브랜드별 매출현황 (단위: $K)').font=Font(bold=True,size=13,name='Calibri')
ws3.merge_cells('A1:ZZ1')

col = 2; mc = {}
for mn in months_list:
    ws3.cell(row=3,column=col,value=mn).font=hdr; ws3.merge_cells(start_row=3,start_column=col,end_row=3,end_column=col+2)
    ws3.cell(row=4,column=col,value='수량').font=hdr; ws3.cell(row=4,column=col+1,value='매출').font=hdr; ws3.cell(row=4,column=col+2,value='순매출').font=hdr
    mc[mn]=col; col+=3
ws3.cell(row=3,column=1,value='브랜드/부서').font=hdr

r = 5
for dn, brs in REF:
    active = [b for b in brs if sum(Sm.get((mn,b),[0]*3)[0] for mn in months_list) > 0]
    if not active: continue
    for b in active:
        ws3.cell(row=r,column=1,value=b).font=Font(size=10)
        for mn in months_list:
            d=Sm.get((mn,b),[0]*3)
            for ci,v in enumerate([int(d[0]),int(round(d[1]/1000)),int(round(d[2]/1000))]):
                cell=ws3.cell(row=r,column=mc[mn]+ci,value=v if v else '-')
                cell.number_format='#,##0' if v else None; cell.font=Font(size=10); cell.border=thin; cell.alignment=Alignment(horizontal='right')
        ws3.cell(row=r,column=1).border=thin; r+=1
    # S.TOTAL
    ws3.cell(row=r,column=1,value='S.TOTAL').font=Font(bold=True,size=10)
    for mn in months_list:
        dq=sum(int(Sm.get((mn,b),[0]*3)[0]) for b in active)
        da=sum(int(round(Sm.get((mn,b),[0]*3)[1]/1000)) for b in active)
        dn_=sum(int(round(Sm.get((mn,b),[0]*3)[2]/1000)) for b in active)
        for ci,v in enumerate([dq,da,dn_]):
            cell=ws3.cell(row=r,column=mc[mn]+ci,value=v if v else '-')
            cell.number_format='#,##0' if v else None; cell.font=Font(bold=True,size=10,name='Calibri'); cell.border=thin; cell.alignment=Alignment(horizontal='right')
    r+=1
print('  Done')

# Reorder
for i, name in enumerate(['1.시즌현황','2.시즌별','3.월별','4.시즌부서공장','5.시즌공장']):
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        if idx != i: wb.move_sheet(name, offset=i-idx)

wb.save(OUT)
print(f'Saved: {os.path.getsize(OUT)/1024:.0f} KB')
wb.close()
