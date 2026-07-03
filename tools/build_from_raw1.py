#!/usr/bin/env python3
"""Build report from file #1, enhanced with mappings from file #2"""
import openpyxl, os, glob
from openpyxl.styles import Font, Border, Side, Alignment
from collections import defaultdict
from datetime import datetime

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'

# Step 1: Extract mappings from file #2 (reference)
print('Step 1: Extracting mappings from file #2...')
f2 = os.path.join(BASE, '2) 영업DATA_Raw_v1.3_20260608_openclaw 작업용.xlsx')
wb2 = openpyxl.load_workbook(f2, read_only=True, data_only=True)
ws2 = wb2['Raw_20260608']

factory_map = {}  # Contract Factory -> Factory
division_map = {}  # Brand -> Division (most common Division per Brand)
month_map = {}  # date -> month string

for i, row in enumerate(ws2.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row)
    cf = str(v[1] or '').strip()
    fac = str(v[2] or '').strip()
    brand = str(v[4] or '').strip()
    div = str(v[10] or '').strip()
    month_str = str(v[21] or '').strip()
    
    if cf and fac:
        factory_map[cf] = fac
    if brand and div:
        key = brand.upper()
        if key not in division_map:
            division_map[key] = div
    if cf and fac:
        key = cf.upper().replace(' ','').replace('.','').replace(',','')
        if key not in factory_map:
            factory_map[key] = fac
    # Also store original key
    if cf and fac:
        factory_map[cf.upper()] = fac

wb2.close()
print(f'  Factory mappings: {len(factory_map)}')
print(f'  Division mappings: {len(division_map)}')

for cf, fac in sorted(factory_map.items())[:5]:
    print(f'    {cf[:40]:40s} -> {fac}')
print('  ...')

for b, d in sorted(division_map.items())[:5]:
    print(f'    {b:25s} -> {d}')
print('  ...')

# Step 2: Read & enhance file #1 (the actual raw data)
print(f'\nStep 2: Reading & enhancing file #1...')
f1 = os.path.join(BASE, '1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx')
wb1 = openpyxl.load_workbook(f1, read_only=True, data_only=True)
ws1 = wb1['Raw_20260629']

seasons = ['26 SPR','26 SUM','26 FAL','26 HOL']
months_kr = {1:'1월',2:'2월',3:'3월',4:'4월',5:'5월',6:'6월',7:'7월',8:'8월',9:'9월',10:'10월',11:'11월',12:'12월'}

# Aggregations using enhanced data
S1 = defaultdict(lambda: [0,0,0,0,0,0])  # (season, division) -> [qty, amt, net, dc, profit, em]
S2 = defaultdict(lambda: [0,0,0])         # (division, season)
S3 = defaultdict(lambda: [0,0,0])         # (division, month)
S4 = defaultdict(lambda: 0)               # (dept, division, factory, season)
S5 = defaultdict(lambda: 0)               # (factory, division, season)

cnt = 0
unmapped_cf = set()
unmapped_brand = set()

for i, row in enumerate(ws1.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row); cnt += 1
    q = v[23] or 0; a = v[24] or 0; disc = v[25] or 0
    net_sales = a - disc
    dc = v[44] or 0; prof = v[46] or 0; em = v[52] or 0
    
    contract_factory = str(v[1] or '').strip()
    brand = str(v[3] or '').strip()
    dept = str(v[4] or '').strip()
    season = str(v[10] or '').strip()
    
    # Derive enhanced columns
    # Case-insensitive lookup
    cf_key = contract_factory.upper()
    # Try exact, then normalized
    factory = factory_map.get(cf_key)
    if not factory:
        cf_norm = cf_key.replace(' ','').replace('.','').replace(',','').replace('(','').replace(')','')
        for fk, fv in factory_map.items():
            if cf_norm in fk or fk in cf_norm:
                factory = fv
                break
    if not factory:
        factory = contract_factory
        unmapped_cf.add(contract_factory)
    
    # Use Grouped Div (I열, index 8 in file #1) as the Division source
    # This matches how file #2's Division(K) was derived from Grouped Div(J)
    grouped_div = str(v[8] or '').strip()
    
    # Map Grouped Div -> Division (following file #2 pattern)
    gd_upper = grouped_div.upper()
    if 'IPSS GAP' in gd_upper:
        division = 'IPSS GAP'
    elif 'IPSS OLD NAVY' in gd_upper:
        division = 'IPSS ON'
    elif 'GAP' in gd_upper:
        division = 'GAP'
    else:
        # Fallback: use brand-level mapping from file #2
        brand_key = brand.upper()
        division = division_map.get(brand_key, brand)
        if brand_key not in division_map:
            unmapped_brand.add(brand)
    
    # Month from ship date (col 18)
    sd = v[18]
    month_str = ''
    if isinstance(sd, datetime):
        month_str = months_kr.get(sd.month, '')
    elif isinstance(sd, str) and len(sd) >= 10:
        try:
            d = datetime.strptime(sd[:10], '%Y-%m-%d')
            month_str = months_kr.get(d.month, '')
        except: pass
    
    if not (q and a and season in set(seasons)): continue
    
    S1[(season, division)] = [S1[(season, division)][i] + x for i, x in enumerate([q, a, net_sales, dc, prof, em])]
    S2[(division, season)] = [S2[(division, season)][i] + x for i, x in enumerate([q, a, net_sales])]
    if month_str: S3[(division, month_str)] = [S3[(division, month_str)][i] + x for i, x in enumerate([q, a, net_sales])]
    S4[(dept, division, factory, season)] += q
    S5[(factory, division, season)] += q
    
    if cnt % 20000 == 0: print(f'  {cnt}...')

wb1.close()
print(f'Aggregated: {cnt} rows')

if unmapped_cf:
    print(f'\nUnmapped Contract Factories ({len(unmapped_cf)}):')
    for cf in sorted(unmapped_cf)[:10]:
        print(f'  {cf[:50]}')
if unmapped_brand:
    print(f'Unmapped Brands ({len(unmapped_brand)}):')
    for b in sorted(unmapped_brand)[:10]:
        print(f'  {b}')

# Step 3: Build report
print('\nStep 3: Building report...')

REF_BRANDS = [
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

# Print coverage
print('Coverage (Division 기준):')
for dn, brs in REF_BRANDS:
    for b in brs:
        qty = sum(S1[(s,b)][0] for s in seasons)
        if qty > 0:
            print(f'  {dn:15s} {b:20s}: {int(qty):>10,}')
        else:
            print(f'  {dn:15s} {b:20s}: 0')

FMT = '#,##0'; FMT2 = '#,##0.00'; FMTP = '0.00'
thin = Border(left=Side(style='thin'),right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
hdr = Font(bold=True, size=10, name='Calibri')
def k(v): return int(round(v/1000))

OUT = os.path.join(BASE, '영업DATA_5Sheet_RAWv1_20260630.xlsx')
wb = openpyxl.Workbook()

# ===== SHEET 1 =====
ws1 = wb.active; ws1.title = '1.시즌현황'
ws1.cell(row=1, column=1, value='[약진통상] 2026 Season (1번 raw 기준, 단위: $K)').font = Font(bold=True, size=13, name='Calibri')
ws1.merge_cells('A1:ZZ1')

p1 = {}; col = 2
for dn, brs in REF_BRANDS:
    act = [b for b in brs if sum(S1[(s,b)][0] for s in seasons) > 0]
    if not act: continue
    cs = col
    for b in act: p1[(dn,b)] = col; col += 6
    if col-1 > cs:
        ws1.merge_cells(start_row=3, start_column=cs, end_row=3, end_column=col-1)
        ws1.cell(row=3, column=cs, value=dn).font = hdr; ws1.cell(row=3, column=cs).alignment = Alignment(horizontal='center')

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

# ===== SHEET 2 =====
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
for dn,brs in REF_BRANDS:
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

# ===== SHEET 3 =====
ws3 = wb.create_sheet('3.월별')
ws3.cell(row=1, column=1, value='2026년 월별 (단위: $K)').font = Font(bold=True, size=13, name='Calibri')
ws3.merge_cells('A1:ZZ1')
p3={}; col=2
months_list = [f'{m}월' for m in range(1,13)]
for dn,brs in REF_BRANDS:
    act = [b for b in brs if sum(S3.get((b,mn),[0,0,0])[0] for mn in months_list) > 0]
    if not act: continue
    cs=col
    for b in act: p3[(dn,b)]=col; col+=3
    if col-1>cs: ws3.merge_cells(start_row=3, start_column=cs, end_row=3, end_column=col-1); ws3.cell(row=3,column=cs,value=dn).font=hdr; ws3.cell(row=3,column=cs).alignment=Alignment(horizontal='center')
ws3.cell(row=4,column=1,value='월').font=hdr
for (dn,b),c in p3.items():
    ws3.merge_cells(start_row=4,start_column=c,end_row=4,end_column=c+2)
    ws3.cell(row=4,column=c,value=b).font=Font(bold=True,size=8,name='Calibri')
    ws3.cell(row=4,column=c).alignment=Alignment(horizontal='center',wrap_text=True)
    for mi,m in enumerate(['수량','매출','순매출']):
        cell=ws3.cell(row=5,column=c+mi,value=m); cell.font=Font(bold=True,size=7,name='Calibri'); cell.alignment=Alignment(horizontal='center'); cell.border=thin
r=6
for mn in months_list:
    ws3.cell(row=r,column=1,value=mn).font=hdr; ws3.cell(row=r,column=1).border=thin
    for (dn,b),c in p3.items():
        sd=S3.get((b,mn),[0,0,0])
        vals=[int(sd[0]),k(sd[1]),k(sd[2])]
        for mi,v in enumerate(vals):
            cell=ws3.cell(row=r,column=c+mi,value=v); cell.number_format=FMT; cell.font=Font(size=10,name='Calibri'); cell.border=thin; cell.alignment=Alignment(horizontal='right')
    r+=1
print('Sheet 3 done')

# ===== SHEET 4 =====
ws4 = wb.create_sheet('4.시즌부서공장')
ws4.cell(row=1, column=1, value='2026 Season 시즌별 Factory x 부서-브랜드 배정 수량').font = Font(bold=True, size=13, name='Calibri')
ws4.merge_cells('A1:ZZ1')
ws4.cell(row=3, column=1, value='시즌').font = hdr; ws4.cell(row=3, column=2, value='Factory').font = hdr
p4={}; col=3
for dn,brs in REF_BRANDS:
    act=[b for b in brs if any(S4[(d,b,f,s)] for (d,b2,f,s) in S4 if b2==b)]
    if not act: continue; cs=col
    for b in act: p4[(dn,b)]=col; col+=1
    if col-1>cs: ws4.merge_cells(start_row=3,start_column=cs,end_row=3,end_column=col-1); ws4.cell(row=3,column=cs,value=dn).font=hdr; ws4.cell(row=3,column=cs).alignment=Alignment(horizontal='center')
for (dn,b),c in p4.items():
    cell=ws4.cell(row=4,column=c,value=b); cell.font=Font(bold=True,size=7,name='Calibri'); cell.alignment=Alignment(horizontal='center',wrap_text=True,text_rotation=90); cell.border=thin
r=5
for season in seasons:
    factories=set(); 
    for (d,b,f,s),q in S4.items():
        if s==season and q>0: factories.add(f)
    factories=sorted(factories)
    if not factories: continue
    sr=r
    for fi,fac in enumerate(factories):
        ws4.cell(row=r+fi,column=2,value=fac).font=Font(size=10,name='Calibri')
        ws4.cell(row=r+fi,column=2).border=thin; ws4.cell(row=r+fi,column=1).border=thin
        for (dn,b),c in p4.items():
            tot=sum(q for (d,b2,f,s),q in S4.items() if b2==b and f==fac and s==season)
            cell=ws4.cell(row=r+fi,column=c,value=int(tot)if tot else 0); cell.number_format=FMT; cell.font=Font(size=10,name='Calibri'); cell.border=thin; cell.alignment=Alignment(horizontal='right')
    if factories:
        ws4.merge_cells(start_row=sr,start_column=1,end_row=sr+len(factories)-1,end_column=1)
        cell=ws4.cell(row=sr,column=1,value=season); cell.font=Font(bold=True,size=10,name='Calibri'); cell.alignment=Alignment(horizontal='center',vertical='center')
    r+=len(factories)+1
ws4.column_dimensions['A'].width=12; ws4.column_dimensions['B'].width=32
print('Sheet 4 done')

# ===== SHEET 5 =====
ws5 = wb.create_sheet('5.시즌공장')
ws5.cell(row=1,column=1,value='2026 Season 공장별 브랜드별 시즌별 배정 수량 및 비중').font=Font(bold=True,size=13,name='Calibri')
ws5.merge_cells('A1:O1')
for ci,h in enumerate(['Factory','브랜드'],1): ws5.cell(row=3,column=ci,value=h).font=hdr
col5={}; c=3
for s in seasons:
    ws5.cell(row=3,column=c,value=s).font=hdr; ws5.cell(row=3,column=c+1,value='비중').font=hdr; col5[s]=c; c+=2
ws5.cell(row=3,column=c,value='Total').font=hdr; ws5.cell(row=3,column=c+1,value='비중').font=hdr; tc5=c
fg=defaultdict(lambda:defaultdict(lambda:defaultdict(int)))
for (fac,br,s),q in S5.items(): fg[fac][br][s]+=q
r=4
for factory in sorted(fg.keys()):
    brands=sorted(fg[factory].keys())
    for bi,brand in enumerate(brands):
        ws5.cell(row=r,column=1,value=factory if bi==0 else '')
        ws5.cell(row=r,column=2,value=brand)
        fbt=sum(fg[factory][brand].get(s,0) for s in seasons)
        for s in seasons:
            q=fg[factory][brand].get(s,0)
            ws5.cell(row=r,column=col5[s],value=int(q)).number_format=FMT
            ws5.cell(row=r,column=col5[s]+1,value=round(q/fbt,4)if fbt else 0).number_format='0.0000'
        ws5.cell(row=r,column=tc5,value=int(fbt)).number_format=FMT; ws5.cell(row=r,column=tc5+1,value=1.0).number_format='0.0000'
        for ci in range(1,tc5+2): ws5.cell(row=r,column=ci).border=thin; ws5.cell(row=r,column=ci).font=Font(size=10,name='Calibri'); ws5.cell(row=r,column=ci).alignment=Alignment(horizontal='right')
        ws5.cell(row=r,column=1).alignment=Alignment(horizontal='left'); ws5.cell(row=r,column=2).alignment=Alignment(horizontal='left')
        r+=1
print('Sheet 5 done')

wb.save(OUT)
print(f'\nSaved: {OUT}')
import os; print(f'Size: {os.path.getsize(OUT)/1024:.0f} KB')
