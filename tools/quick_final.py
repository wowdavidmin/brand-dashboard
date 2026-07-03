#!/usr/bin/env python3
"""Final report v2 using Division(K) column - quick version"""
import openpyxl, os, glob
from openpyxl.styles import Font, Border, Side, Alignment
from collections import defaultdict

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
files = [f for f in glob.glob(os.path.join(BASE, '2)*')) if not os.path.basename(f).startswith('~$')]
F2 = max(files, key=os.path.getmtime)
OUT = os.path.join(BASE, '영업DATA_5Sheet_FINALv2_20260630.xlsx')
print(f'Output: {OUT}')

wb2 = openpyxl.load_workbook(F2, read_only=True, data_only=True)
ws = wb2['Raw_20260608']

seasons = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
S1 = defaultdict(lambda: [0,0,0,0,0,0])
S2 = defaultdict(lambda: [0,0,0])

cnt = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row); cnt += 1
    q = v[26] or 0; a = v[27] or 0; n = v[29] or 0
    dc=v[48]or 0; prof=v[50]or 0; em=v[56]or 0
    ref = str(v[10] or '').strip() or str(v[4] or '').strip()
    s = str(v[12] or '').strip()
    if not (q and a and s in seasons): continue
    S1[(s, ref)] = [S1[(s, ref)][i]+x for i,x in enumerate([q,a,n,dc,prof,em])]
    S2[(ref, s)] = [S2[(ref, s)][i]+x for i,x in enumerate([q,a,n])]
    if cnt%30000==0: print(f'  {cnt}...')
wb2.close()
print(f'Aggregated: {cnt} rows')

# Reference brand structure
REF = [
    ('BU1-영업1부', ['GAP', 'IPSS GAP', 'BRFS', 'Madhappy', 'Mott & Bow']),
    ('BU2-영업1부', ['OLD NAVY', "ON  MEN'S"]),
    ('BU2-영업2부', ['WALMART', 'QUINCE']),
    ('BU3-영업1부', ['BR', 'IPSS ON']),
    ('BU3-영업2부', ["KOHL'S", 'URBAN']),
    ('BU3-영업3부', ['AERIE', 'Offline']),
    ('BU3-영업4부', ['MLB']),
    ('BU4부', ['NUUDS', 'ONER ACTIVE', 'ROSS', 'AEO', 'TOMS']),
    ('BU5-영업5부', ['Justice', 'INGRID & ISABEL']),
    ('GTBU-과테말라', ['OLDNAVY(HYBRID)']),
]

# Print coverage
print('Coverage:')
for dn, brs in REF:
    for b in brs:
        qty = sum(S1[(s,b)][0] for s in seasons)
        if qty > 0:
            print(f'  {dn:15s} {b:20s}: {int(qty):>10,}')
        else:
            print(f'  {dn:15s} {b:20s}: 0')

# Build simple Excel
FMT = '#,##0'; FMT2 = '#,##0.00'; FMTP = '0.00'
thin = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
hdr = Font(bold=True, size=10, name='Calibri')
def k(v): return int(round(v/1000))

wb = openpyxl.Workbook()
ws1 = wb.active; ws1.title = '1.시즌현황'
ws1.cell(row=1, column=1, value='[약진통상] 2026 Season (단위: $K, Division 기준)').font = Font(bold=True, size=13, name='Calibri')
ws1.merge_cells('A1:ZZ1')

p1 = {}; col = 2
for dn, brs in REF:
    act = [b for b in brs if sum(S1[(s,b)][0] for s in seasons) > 0]
    if not act: continue
    cs = col
    for b in act: p1[(dn,b)] = col; col += 6
    if col-1 > cs:
        ws1.merge_cells(start_row=3, start_column=cs, end_row=3, end_column=col-1)
        ws1.cell(row=3, column=cs, value=dn).font = hdr
        ws1.cell(row=3, column=cs).alignment = Alignment(horizontal='center')

ws1.cell(row=4, column=1, value='시즌').font = hdr
ml = ['수량','총매출','순매출','평균단가','마진율%','CM%']
for (dn,b), c in p1.items():
    ws1.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c+5)
    ws1.cell(row=4, column=c, value=b).font = Font(bold=True, size=8, name='Calibri')
    ws1.cell(row=4, column=c).alignment = Alignment(horizontal='center', wrap_text=True)
    for mi, m in enumerate(ml):
        cell = ws1.cell(row=5, column=c+mi, value=m)
        cell.font = Font(bold=True, size=7, name='Calibri')
        cell.alignment = Alignment(horizontal='center'); cell.border = thin

r = 6
for season in seasons + ['Total']:
    ws1.cell(row=r, column=1, value=season).font = Font(bold=(season=='Total'), size=10, name='Calibri')
    ws1.cell(row=r, column=1).border = thin
    for (dn,b), c in p1.items():
        if season == 'Total':
            sd = [sum(S1[(s,b)][i] for s in seasons) for i in range(6)]
        else:
            sd = S1.get((season,b), [0]*6)
        tq, ta, tn, _, tp, tem = sd
        if not tq:
            for mi in range(6):
                cell = ws1.cell(row=r, column=c+mi, value=0)
                cell.font = Font(size=10); cell.border = thin; cell.alignment = Alignment(horizontal='right')
            continue
        vals = [int(tq), k(ta), k(tn), round(ta/tq,2) if tq else 0,
                round(tp/ta*100,2) if ta else 0, round(tem/ta*100,2) if ta else 0]
        fmts = [FMT, FMT, FMT, FMT2, FMTP, FMTP]
        for mi, (v, fmt) in enumerate(zip(vals, fmts)):
            cell = ws1.cell(row=r, column=c+mi, value=v)
            cell.number_format = fmt
            cell.font = Font(bold=(season=='Total'), size=10, name='Calibri')
            cell.border = thin; cell.alignment = Alignment(horizontal='right')
    r += 1
print('Sheet 1 done')

# Sheet 2
ws2 = wb.create_sheet('2.시즌별')
ws2.cell(row=1, column=1, value='2026 Season별 (단위: $K)').font = Font(bold=True, size=13, name='Calibri')
ws2.merge_cells('A1:T1')
ws2.cell(row=3, column=1, value='브랜드/부서').font = hdr
c=2; co2={}
for s in seasons:
    ws2.cell(row=3, column=c, value=s).font = hdr
    ws2.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+2)
    ws2.cell(row=4, column=c, value='수량').font = hdr
    ws2.cell(row=4, column=c+1, value='총매출').font = hdr
    ws2.cell(row=4, column=c+2, value='순매출').font = hdr
    co2[s] = (c,c+1,c+2); c+=3
ws2.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+3)
ws2.cell(row=3, column=c, value='TOTAL').font = hdr
ws2.cell(row=4, column=c, value='수량').font = hdr
ws2.cell(row=4, column=c+1, value='총매출').font = hdr
ws2.cell(row=4, column=c+2, value='순매출').font = hdr
ws2.cell(row=4, column=c+3, value='마진율%').font = hdr; tc2=c

r=5
for dn,brs in REF:
    act = [b for b in brs if sum(S2[(b,s)][0] for s in seasons) > 0]
    if not act: continue
    for b in act:
        ws2.cell(row=r, column=1, value=b)
        for s in seasons:
            d = S2.get((b,s), [0,0,0])
            cq,ca,cn_ = co2[s]
            ws2.cell(row=r, column=cq, value=int(d[0])).number_format = FMT
            ws2.cell(row=r, column=ca, value=k(d[1])).number_format = FMT
            ws2.cell(row=r, column=cn_, value=k(d[2])).number_format = FMT
        tq = sum(S2.get((b,s),[0,0,0])[0] for s in seasons)
        ta = sum(S2.get((b,s),[0,0,0])[1] for s in seasons)
        tn = sum(S2.get((b,s),[0,0,0])[2] for s in seasons)
        ws2.cell(row=r, column=tc2, value=int(tq)).number_format = FMT
        ws2.cell(row=r, column=tc2+1, value=k(ta)).number_format = FMT
        ws2.cell(row=r, column=tc2+2, value=k(tn)).number_format = FMT
        ws2.cell(row=r, column=tc2+3, value=round((ta-tn)/ta*100,2) if ta else 0).number_format = FMTP
        for ci in range(1, tc2+4):
            ws2.cell(row=r, column=ci).border = thin
            ws2.cell(row=r, column=ci).font = Font(size=10, name='Calibri')
            ws2.cell(row=r, column=ci).alignment = Alignment(horizontal='right')
        ws2.cell(row=r, column=1).alignment = Alignment(horizontal='left')
        r+=1
    ws2.cell(row=r, column=1, value=f'{dn} 요약').font = Font(bold=True)
    for s in seasons:
        dq = sum(S2.get((b,s),[0,0,0])[0] for b in act)
        da = sum(S2.get((b,s),[0,0,0])[1] for b in act)
        dn_ = sum(S2.get((b,s),[0,0,0])[2] for b in act)
        cq,ca,cn_ = co2[s]
        ws2.cell(row=r, column=cq, value=int(dq)).font = Font(bold=True)
        ws2.cell(row=r, column=ca, value=k(da)).font = Font(bold=True)
        ws2.cell(row=r, column=cn_, value=k(dn_)).font = Font(bold=True)
    tq = sum(sum(S2.get((b,s),[0,0,0])[0] for s in seasons) for b in act)
    ta = sum(sum(S2.get((b,s),[0,0,0])[1] for s in seasons) for b in act)
    tn = sum(sum(S2.get((b,s),[0,0,0])[2] for s in seasons) for b in act)
    ws2.cell(row=r, column=tc2, value=int(tq)).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+1, value=k(ta)).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+2, value=k(tn)).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+3, value=round((ta-tn)/ta*100,2) if ta else 0).font = Font(bold=True)
    for ci in range(1,tc2+4): ws2.cell(row=r, column=ci).border = thin
    r+=1
print('Sheet 2 done')

wb.save(OUT)
print(f'Saved: {OUT}')
