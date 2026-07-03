#!/usr/bin/env python3
"""Clean rebuild with S.TOTAL columns"""
import openpyxl, os, shutil
from openpyxl.styles import Font, Border, Side, Alignment
from collections import defaultdict

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
ENH = os.path.join(BASE, '영업DATA_Enhanced_v1.0_20260629.xlsx')
TEMPLATE = os.path.join(BASE, '리포트_5Sheet_최종_v8_20260630.xlsx')
OUT = os.path.join(BASE, '리포트_5Sheet_최종_v11_20260630.xlsx')
shutil.copy2(TEMPLATE, OUT)

# Read data
print('Reading enhanced data...')
wb_enh = openpyxl.load_workbook(ENH, read_only=True, data_only=True)
ws_enh = wb_enh['Raw_20260629']
seasons = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
months = [f'{m}월' for m in range(1, 13)]

def n2(x):
    if x is None: return 0
    if isinstance(x, (int, float)): return x
    try: return float(str(x).replace(',',''))
    except: return 0

SD = defaultdict(lambda: [0]*7)  # (season, div) -> [q,a,n,dc,p,em,cmt]
MD = defaultdict(lambda: [0]*3)  # (month, div) -> [q,a,n]

for i, row in enumerate(ws_enh.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row)
    s = str(v[12] or '').strip()
    d = str(v[10] or '').strip()
    m = str(v[21] or '').strip()
    q = n2(v[26]); a = n2(v[27]); n = n2(v[29])
    dc = n2(v[44]); p = n2(v[46]); em = n2(v[52]); cmt = n2(v[44])
    if not (q and a): continue
    if s in seasons:
        t = SD[(s,d)]; SD[(s,d)] = [t[i]+x for i,x in enumerate([q,a,n,dc,p,em,cmt])]
    if m in months:
        t = MD[(m,d)]; MD[(m,d)] = [t[i]+x for i,x in enumerate([q,a,n])]
wb_enh.close()
print('  Aggregated')

def bt(b):  # brand total across seasons
    return [sum(SD[(s,b)][i] for s in seasons) for i in range(7)]

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
FMT='#,##0'; FMT2='#,##0.00'; FMTP='0.0'

wb = openpyxl.load_workbook(OUT)

# Build col list with S.TOTAL
CL = []  # (dept, name, is_total)
for dn, brs in REF:
    act = [b for b in brs if sum(SD[(s,b)][0] for s in seasons) > 0]
    if not act: continue
    for b in act: CL.append((dn, b, False))
    CL.append((dn, 'S.TOTAL', True))

def get_val(season, item, is_total, midx, dn):
    vlist = [0]*7
    if is_total:
        for d2, b2, is2 in CL:
            if d2 == dn and not is2:
                sv = bt(b2) if season == 'Total' else SD.get((season, b2), [0]*7)
                for i in range(7): vlist[i] += sv[i]
    else:
        vlist = bt(item) if season == 'Total' else SD.get((season, item), [0]*7)
    if midx == 0: return int(vlist[0])
    elif midx == 1: return int(round(vlist[1]/1000))
    elif midx == 2: return int(round(vlist[2]/1000))
    elif midx == 3: return round(vlist[1]/vlist[0], 2) if vlist[0] else 0
    elif midx == 4: return round(vlist[4]/vlist[1]*100, 1) if vlist[1] else 0
    elif midx == 5: return round(vlist[6]/vlist[0], 2) if vlist[0] else 0

def numfmt(midx):
    if midx in (3,5): return FMT2
    elif midx == 4: return FMTP
    return FMT

# ===== SHEET 1 =====
print('Sheet 1...')
del wb['1.시즌현황']
ws1 = wb.create_sheet('1.시즌현황', 0)
ws1.cell(row=1,column=1,value='[약진통상] 2026 Season 부서별 브랜드별 현황 (단위: $K)').font=Font(bold=True,size=13,name='Calibri')
ws1.merge_cells('A1:ZZ1')

r = 3
for sec_name, midx in [('1. ORDER 수량',0),('2. 총매출액',1),('3. 순매출액',2),('4. 평균단가',3),('5. 마진율',4),('6. CM',5)]:
    # Section title
    ws1.cell(row=r,column=1,value=sec_name).font=Font(bold=True,size=12,name='Calibri'); r+=1
    # Dept headers
    col=2; pdn=None; cs=2
    for dn, name, ist in CL:
        if dn != pdn:
            if pdn is not None and cs < col-1:
                ws1.merge_cells(start_row=r,start_column=cs,end_row=r,end_column=col-1)
                ws1.cell(row=r,column=cs,value=pdn).font=hdr; ws1.cell(row=r,column=cs).alignment=Alignment(horizontal='center')
            cs=col; pdn=dn
        col+=1
    if cs < col-1:
        ws1.merge_cells(start_row=r,start_column=cs,end_row=r,end_column=col-1)
        ws1.cell(row=r,column=cs,value=pdn).font=hdr; ws1.cell(row=r,column=cs).alignment=Alignment(horizontal='center')
    r+=1
    # Brand/header
    ws1.cell(row=r,column=1,value='시즌').font=hdr; ws1.cell(row=r,column=1).border=thin
    col=2
    for dn, name, ist in CL:
        cell=ws1.cell(row=r,column=col,value=name)
        cell.font=Font(bold=True,size=8,name='Calibri'); cell.alignment=Alignment(horizontal='center',wrap_text=True); cell.border=thin; col+=1
    r+=1
    # Data
    for season in seasons+['Total']:
        ws1.cell(row=r,column=1,value=season).font=Font(bold=(season=='Total'),size=10); ws1.cell(row=r,column=1).border=thin
        col=2
        for dn, name, ist in CL:
            v = get_val(season, name, ist, midx, dn)
            cell = ws1.cell(row=r,column=col,value=v if v else '-')
            if v:
                cell.number_format = numfmt(midx)
                cell.font = Font(bold=(ist or season=='Total'), size=10, name='Calibri')
            else:
                cell.font = Font(size=10, name='Calibri')
            cell.border=thin; cell.alignment=Alignment(horizontal='right'); col+=1
        r+=1
    r+=1

print('  Done')

# ===== SHEET 3 =====
print('Sheet 3...')
del wb['3.월별']
ws3 = wb.create_sheet('3.월별')
ws3.cell(row=1,column=1,value='2026년 월별 부서별 브랜드별 매출현황 (단위: $K)').font=Font(bold=True,size=13,name='Calibri')
ws3.merge_cells('A1:ZZ1')
col=2; mc={}
for mn in months:
    ws3.cell(row=3,column=col,value=mn).font=hdr; ws3.merge_cells(start_row=3,start_column=col,end_row=3,end_column=col+2)
    ws3.cell(row=4,column=col,value='수량').font=hdr; ws3.cell(row=4,column=col+1,value='매출').font=hdr; ws3.cell(row=4,column=col+2,value='순매출').font=hdr
    mc[mn]=col; col+=3
ws3.cell(row=3,column=1,value='브랜드/부서').font=hdr
r=5
for dn, brs in REF:
    act = [b for b in brs if sum(MD[(mn,b)][0] for mn in months) > 0]
    if not act: continue
    for b in act:
        ws3.cell(row=r,column=1,value=b).font=Font(size=10); ws3.cell(row=r,column=1).border=thin
        for mn in months:
            d=MD[(mn,b)]; vals=[int(d[0]),int(round(d[1]/1000)),int(round(d[2]/1000))]
            for ci,v in enumerate(vals):
                cell=ws3.cell(row=r,column=mc[mn]+ci,value=v if v else '-')
                cell.number_format=FMT; cell.font=Font(size=10); cell.border=thin; cell.alignment=Alignment(horizontal='right')
        r+=1
    # S.TOTAL
    ws3.cell(row=r,column=1,value='S.TOTAL').font=Font(bold=True,size=10,name='Calibri')
    for mn in months:
        dq=sum(int(MD[(mn,b)][0]) for b in act)
        da=sum(int(round(MD[(mn,b)][1]/1000)) for b in act)
        dn_=sum(int(round(MD[(mn,b)][2]/1000)) for b in act)
        for ci,v in enumerate([dq,da,dn_]):
            cell=ws3.cell(row=r,column=mc[mn]+ci,value=v if v else '-')
            cell.number_format=FMT; cell.font=Font(bold=True,size=10,name='Calibri'); cell.border=thin; cell.alignment=Alignment(horizontal='right')
    r+=1
print('  Done')

# Reorder
for i, name in enumerate(['1.시즌현황','2.시즌별','3.월별','4.시즌부서공장','5.시즌공장']):
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        if idx != i: wb.move_sheet(name, offset=i-idx)

wb.save(OUT)
import os; print(f'Saved: {os.path.getsize(OUT)/1024:.0f} KB')
wb.close()
