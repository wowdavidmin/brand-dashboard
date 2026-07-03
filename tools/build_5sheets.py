#!/usr/bin/env python3
"""Build 5-sheet report from file #2 data. All monetary values in $K (thousands, integer)."""
import openpyxl, os, glob
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
from collections import defaultdict
from datetime import datetime

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
files = [f for f in glob.glob(os.path.join(BASE, '2)*')) if not os.path.basename(f).startswith('~$')]
if not files:
    files = [f for f in glob.glob(os.path.join(BASE, '*Raw*')) if not os.path.basename(f).startswith('~$')]
F2 = max(files, key=os.path.getmtime)
OUT = os.path.join(BASE, f'영업DATA_5Sheet_{datetime.now().strftime("%Y%m%d")}.xlsx')

print(f'Reading: {F2}')
wb2 = openpyxl.load_workbook(F2, read_only=True, data_only=True)
ws = wb2['Raw_20260608']

seasons_list = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']

# Aggregation stores
S1 = defaultdict(lambda: [0,0,0,0,0,0])  # (season, buyer) -> [qty, amt, net, dc, profit, exec_margin]
S2 = defaultdict(lambda: [0,0,0])        # (brand, season) -> [qty, amt, net]
S3 = defaultdict(lambda: [0,0,0])        # (month, buyer) -> [qty, amt, net]
S4 = defaultdict(lambda: 0)              # (dept, brand, factory, season) -> qty
S5 = defaultdict(lambda: 0)              # (factory, brand, season) -> qty
all_buyers = set()
all_brands = set()
cnt = 0

for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row); cnt += 1
    q = v[26] or 0; a = v[27] or 0; n = v[29] or 0
    dc = v[48] or 0; prof = v[50] or 0; em = v[56] or 0
    buyer = str(v[3] or '').strip()
    brand = str(v[4] or '').strip()
    season = str(v[12] or '').strip()
    dept = str(v[5] or '').strip()
    factory = str(v[2] or '').strip()
    month = str(v[21] or '').strip()

    if not (q and a and season in seasons_list): continue

    S1[(season, buyer)] = [S1[(season, buyer)][i] + x for i, x in enumerate([q, a, n, dc, prof, em])]
    all_buyers.add(buyer)
    S2[(brand, season)] = [S2[(brand, season)][i] + x for i, x in enumerate([q, a, n])]
    all_brands.add(brand)
    if month: S3[(month, buyer)] = [S3[(month, buyer)][i] + x for i, x in enumerate([q, a, n])]
    S4[(dept, brand, factory, season)] += q
    S5[(factory, brand, season)] += q

    if cnt % 30000 == 0: print(f'  {cnt}...')

wb2.close()
print(f'Aggregated: {cnt} rows, {len(all_buyers)} buyers, {len(all_brands)} brands')

# === Department/Brand mapping ===
def buyer_to_dept(b):
    x = b.upper()
    if 'OLD NAVY' in x: return 'BU2-영업1부'
    if x.startswith('ON ') or 'ON MEN' in x or 'ON WOM' in x or 'ON KID' in x: return 'BU2-영업1부'
    if 'WALMART' in x: return 'BU2-영업2부'
    if 'QUINCE' in x: return 'BU2-영업2부'
    if x in ('BR',): return 'BU3-영업1부'
    if 'IPSS ON' in x: return 'BU3-영업1부'
    if 'KOHL' in x: return 'BU3-영업2부'
    if 'URBAN' in x: return 'BU3-영업2부'
    if 'AERIE' in x: return 'BU3-영업3부'
    if 'OFFLINE' in x: return 'BU3-영업3부'
    if 'MLB' in x: return 'BU3-영업4부'
    if 'GAP' in x and 'IPSS' not in x: return 'BU1-영업1부'
    if 'IPSS GAP' in x: return 'BU1-영업1부'
    if x in ('BRFS',): return 'BU1-영업1부'
    if x in ('MADHAPPY',): return 'BU1-영업1부'
    if 'MOTT' in x or 'BOW' in x: return 'BU1-영업1부'
    return 'BU4부'

def brand_to_dept(b):
    return buyer_to_dept(b)

# Department order for columns
DEPT_ORDER = ['BU1-영업1부', 'BU2-영업1부', 'BU2-영업2부', 'BU3-영업1부', 'BU3-영업2부', 'BU3-영업3부', 'BU3-영업4부', 'BU4부']

# Collect buyers by dept
dept_buyers = defaultdict(list)
for b in sorted(all_buyers):
    dept_buyers[buyer_to_dept(b)].append(b)

# Collection of brand lists for sheets 2
dept_brands = defaultdict(set)
for b in all_brands:
    dept_brands[brand_to_dept(b)].add(b)

months_list = [f'{m}월' for m in range(1, 13)]

# ===== BUILD EXCEL =====
thin = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
hdr_font = Font(bold=True, size=10, name='Calibri')
title_font = Font(bold=True, size=13, name='Calibri')
subj_font = Font(bold=True, size=11, name='Calibri')
num_font = Font(size=10, name='Calibri')
gray_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')

def k(v):
    return int(round(v / 1000))

def unit_price(qty, amt):
    """Unit price in actual dollars (not K)"""
    return round(amt / qty, 2) if qty else 0

def margin_rate(amt, net):
    """Profit margin as percentage"""
    # Using PROFIT / total sales based on available data
    return round((amt - net) / amt * 100, 2) if amt else 0

wb = openpyxl.Workbook()

# ======================================================
# SHEET 1: 시즌현황 (Season x Org-Brand Matrix)
# ======================================================
ws1 = wb.active
ws1.title = '1.시즌현황'
ws1.cell(row=1, column=1, value='[약진통상] 2026 Season 부서별 브랜드별 현황 (단위: $K, 수량: PCS)').font = title_font
ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=60)

# Build column structure: Department header row
# Row 3: Department headers
# Row 4: Brand headers
# Row 5: Metric headers (수량, 총매출, 순매출, 평균단가, 마진율, CM평균)

# Column layout per brand: 6 columns (수량|총매출|순매출|평균단가|마진율|CM평균)
# Brand section header will merge 6 cells

col = 2  # Start from column B (A is season label)
col_map = {}  # (dept, buyer) -> (col_start, 6 cols)
buyer_order = []  # Ordered list of (dept, buyer) for columns

for dept in DEPT_ORDER:
    buyers = dept_buyers.get(dept, [])
    if not buyers: continue
    for buyer in buyers:
        buyer_order.append((dept, buyer))
        col_map[(dept, buyer)] = col
        col += 6  # 6 metrics per buyer
    # S.TOTAL
    buyer_order.append((dept, 'S.TOTAL'))
    col_map[(dept, 'S.TOTAL')] = col
    col += 6

num_cols = col - 1  # Total data columns

# Row 3: Department headers (merged across their buyers)
ws1.cell(row=3, column=1, value='').font = hdr_font
for dept in DEPT_ORDER:
    buyers = dept_buyers.get(dept, [])
    if not buyers: continue
    # Find first and last col for this dept (including S.TOTAL)
    first_buyer = buyers[0]
    last_total = ('S.TOTAL')
    # Get the first col of first buyer
    c_start = col_map[(dept, first_buyer)]
    # S.TOTAL is the last item for this dept in buyer_order
    last_item_col = col_map.get((dept, 'S.TOTAL'))
    # Actually find the last column in buyer_order for this dept
    c_end = col_map[(dept, 'S.TOTAL')] + 5
    if c_end > c_start:
        ws1.merge_cells(start_row=3, start_column=c_start, end_row=3, end_column=c_end)
        ws1.cell(row=3, column=c_start, value=dept).font = Font(bold=True, size=10, name='Calibri')
        ws1.cell(row=3, column=c_start).alignment = Alignment(horizontal='center')

# Row 4: Brand names
ws1.cell(row=4, column=1, value='시즌').font = hdr_font
for (dept, buyer) in buyer_order:
    c = col_map[(dept, buyer)]
    ws1.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c+5)
    ws1.cell(row=4, column=c, value=buyer).font = Font(bold=True, size=9, name='Calibri')
    ws1.cell(row=4, column=c).alignment = Alignment(horizontal='center', wrap_text=True)
    metric_labels = ['수량', '총매출', '순매출', '평균단가', '마진율%', 'CM평균%']
    for mi, ml in enumerate(metric_labels):
        cell = ws1.cell(row=5, column=c+mi, value=ml)
        cell.font = Font(bold=True, size=8, name='Calibri')
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = thin
    ws1.cell(row=4, column=c).border = thin

# Data rows
r = 6
for season in seasons_list + ['Total']:
    ws1.cell(row=r, column=1, value=season).font = Font(bold=(season=='Total'), size=10, name='Calibri')
    ws1.cell(row=r, column=1).border = thin
    
    for (dept, buyer) in buyer_order:
        c = col_map[(dept, buyer)]
        
        if buyer == 'S.TOTAL':
            # Sum all buyers in this dept
            dept_buyers_list = dept_buyers.get(dept, [])
            if season == 'Total':
                tq = sum(S1[(s, b)][0] for s in seasons_list for b in dept_buyers_list)
                ta = sum(S1[(s, b)][1] for s in seasons_list for b in dept_buyers_list)
                tn = sum(S1[(s, b)][2] for s in seasons_list for b in dept_buyers_list)
                tdc = sum(S1[(s, b)][3] for s in seasons_list for b in dept_buyers_list)
                tp = sum(S1[(s, b)][4] for s in seasons_list for b in dept_buyers_list)
                tem = sum(S1[(s, b)][5] for s in seasons_list for b in dept_buyers_list)
            else:
                tq = sum(S1[(season, b)][0] for b in dept_buyers_list)
                ta = sum(S1[(season, b)][1] for b in dept_buyers_list)
                tn = sum(S1[(season, b)][2] for b in dept_buyers_list)
                tdc = sum(S1[(season, b)][3] for b in dept_buyers_list)
                tp = sum(S1[(season, b)][4] for b in dept_buyers_list)
                tem = sum(S1[(season, b)][5] for b in dept_buyers_list)
        else:
            key = (season, buyer) if season != 'Total' else None
            if season == 'Total':
                tq = sum(S1[(s, buyer)][0] for s in seasons_list)
                ta = sum(S1[(s, buyer)][1] for s in seasons_list)
                tn = sum(S1[(s, buyer)][2] for s in seasons_list)
                tdc = sum(S1[(s, buyer)][3] for s in seasons_list)
                tp = sum(S1[(s, buyer)][4] for s in seasons_list)
                tem = sum(S1[(s, buyer)][5] for s in seasons_list)
            else:
                sd = S1.get((season, buyer), [0,0,0,0,0,0])
                tq, ta, tn, tdc, tp, tem = sd
        
        if tq == 0:
            for mi in range(6):
                cell = ws1.cell(row=r, column=c+mi, value=0)
                cell.font = num_font; cell.border = thin; cell.alignment = Alignment(horizontal='right')
            continue
        
        vals = [
            int(tq),          # 수량
            k(ta),            # 총매출 $K
            k(tn),            # 순매출 $K
            unit_price(tq, ta),  # 평균단가 ($, raw)
            round(tp/ta*100, 2) if ta else 0,  # 마진율%
            round(tem/tn*100, 2) if tn else 0,  # CM평균%
        ]
        for mi, v in enumerate(vals):
            cell = ws1.cell(row=r, column=c+mi, value=v)
            cell.font = num_font; cell.border = thin
            cell.alignment = Alignment(horizontal='right')
            # Bold for Total row or S.TOTAL
            if season == 'Total' or buyer == 'S.TOTAL':
                cell.font = Font(bold=True, size=10, name='Calibri')
    
    r += 1

print('Sheet 1 done')

# ======================================================
# SHEET 2: 시즌별 (Brand x Season)
# ======================================================
ws2 = wb.create_sheet('2.시즌별')
ws2.cell(row=1, column=1, value='2026 Season별 브랜드별 매출현황 (단위: $K)').font = title_font
ws2.merge_cells(start_row=1, start_column=1, end_row=1, end_column=20)

c = 2
col2 = {}
for s in seasons_list:
    ws2.cell(row=3, column=c, value=s).font = hdr_font
    ws2.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+2)
    ws2.cell(row=4, column=c, value='수량').font = hdr_font
    ws2.cell(row=4, column=c+1, value='총매출').font = hdr_font
    ws2.cell(row=4, column=c+2, value='순매출').font = hdr_font
    col2[s] = (c, c+1, c+2)
    c += 3

ws2.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+3)
ws2.cell(row=3, column=c, value='TOTAL').font = hdr_font
ws2.cell(row=4, column=c, value='수량').font = hdr_font
ws2.cell(row=4, column=c+1, value='총매출').font = hdr_font
ws2.cell(row=4, column=c+2, value='순매출').font = hdr_font
ws2.cell(row=4, column=c+3, value='마진율%').font = hdr_font
tc2 = c

ws2.cell(row=3, column=1, value='브랜드/부서').font = hdr_font
ws2.cell(row=4, column=1, value='').font = hdr_font

r = 5
for dn in DEPT_ORDER:
    brands = sorted(dept_brands.get(dn, set()))
    if not brands: continue
    for brand in brands:
        ws2.cell(row=r, column=1, value=brand)
        for s in seasons_list:
            d = S2.get((brand, s), [0,0,0])
            cq, ca, cn_ = col2[s]
            ws2.cell(row=r, column=cq, value=int(d[0]))
            ws2.cell(row=r, column=ca, value=k(d[1]))
            ws2.cell(row=r, column=cn_, value=k(d[2]))
        tq = sum(S2.get((brand, s), [0,0,0])[0] for s in seasons_list)
        ta = sum(S2.get((brand, s), [0,0,0])[1] for s in seasons_list)
        tn = sum(S2.get((brand, s), [0,0,0])[2] for s in seasons_list)
        ws2.cell(row=r, column=tc2, value=int(tq))
        ws2.cell(row=r, column=tc2+1, value=k(ta))
        ws2.cell(row=r, column=tc2+2, value=k(tn))
        ws2.cell(row=r, column=tc2+3, value=round((ta-tn)/ta*100, 2) if ta else 0)
        for ci in range(1, tc2+4):
            ws2.cell(row=r, column=ci).border = thin
            ws2.cell(row=r, column=ci).font = num_font
            ws2.cell(row=r, column=ci).alignment = Alignment(horizontal='right')
        ws2.cell(row=r, column=1).alignment = Alignment(horizontal='left')
        ws2.cell(row=r, column=1).font = num_font
        r += 1
    # Subtotal
    ws2.cell(row=r, column=1, value=f'{dn} 요약').font = Font(bold=True, size=10, name='Calibri')
    for s in seasons_list:
        dq = sum(S2.get((b, s), [0,0,0])[0] for b in brands)
        da = sum(S2.get((b, s), [0,0,0])[1] for b in brands)
        dn_ = sum(S2.get((b, s), [0,0,0])[2] for b in brands)
        cq, ca, cn_ = col2[s]
        ws2.cell(row=r, column=cq, value=int(dq)).font = Font(bold=True)
        ws2.cell(row=r, column=ca, value=k(da)).font = Font(bold=True)
        ws2.cell(row=r, column=cn_, value=k(dn_)).font = Font(bold=True)
    tq = sum(sum(S2.get((b, s), [0,0,0])[0] for s in seasons_list) for b in brands)
    ta = sum(sum(S2.get((b, s), [0,0,0])[1] for s in seasons_list) for b in brands)
    tn = sum(sum(S2.get((b, s), [0,0,0])[2] for s in seasons_list) for b in brands)
    ws2.cell(row=r, column=tc2, value=int(tq)).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+1, value=k(ta)).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+2, value=k(tn)).font = Font(bold=True)
    ws2.cell(row=r, column=tc2+3, value=round((ta-tn)/ta*100, 2) if ta else 0).font = Font(bold=True)
    for ci in range(1, tc2+4):
        ws2.cell(row=r, column=ci).border = thin
    r += 1
print('Sheet 2 done')

# ======================================================
# SHEET 3: 월별 (Month x Org-Brand)
# ======================================================
ws3 = wb.create_sheet('3.월별')
ws3.cell(row=1, column=1, value='2026년 월별 조직-브랜드별 매출현황 (단위: $K)').font = title_font
ws3.merge_cells(start_row=1, start_column=1, end_row=1, end_column=60)

# Column structure same as sheet 1 but only 수량+총매출+순매출 (3 metrics per brand)
col3 = {}  # (dept, buyer) -> col_start
buyer_order3 = []
col = 2
for dept in DEPT_ORDER:
    buyers = dept_buyers.get(dept, [])
    if not buyers: continue
    for buyer in buyers:
        buyer_order3.append((dept, buyer))
        col3[(dept, buyer)] = col
        col += 3
    buyer_order3.append((dept, 'S.TOTAL'))
    col3[(dept, 'S.TOTAL')] = col
    col += 3

# Row 3: Dept headers
for dept in DEPT_ORDER:
    buyers = dept_buyers.get(dept, [])
    if not buyers: continue
    c_start = col3.get((dept, buyers[0]))
    c_end = col3.get((dept, 'S.TOTAL')) + 2
    if c_end > c_start:
        ws3.merge_cells(start_row=3, start_column=c_start, end_row=3, end_column=c_end)
        ws3.cell(row=3, column=c_start, value=dept).font = Font(bold=True, size=10, name='Calibri')
        ws3.cell(row=3, column=c_start).alignment = Alignment(horizontal='center')

# Row 4: Brand names
ws3.cell(row=4, column=1, value='월').font = hdr_font
for (dept, buyer) in buyer_order3:
    c = col3[(dept, buyer)]
    ws3.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c+2)
    ws3.cell(row=4, column=c, value=buyer).font = Font(bold=True, size=9, name='Calibri')
    ws3.cell(row=4, column=c).alignment = Alignment(horizontal='center', wrap_text=True)
    for mi, ml in enumerate(['수량', '매출', '순매출']):
        cell = ws3.cell(row=5, column=c+mi, value=ml)
        cell.font = Font(bold=True, size=8, name='Calibri')
        cell.alignment = Alignment(horizontal='center')
        cell.border = thin

# Data rows: 1월~12월
r = 6
for mn in months_list:
    ws3.cell(row=r, column=1, value=mn).font = hdr_font
    ws3.cell(row=r, column=1).border = thin
    for (dept, buyer) in buyer_order3:
        c = col3[(dept, buyer)]
        if buyer == 'S.TOTAL':
            dq = sum(S3.get((mn, b), [0,0,0])[0] for b in dept_buyers.get(dept, []))
            da = sum(S3.get((mn, b), [0,0,0])[1] for b in dept_buyers.get(dept, []))
            dn = sum(S3.get((mn, b), [0,0,0])[2] for b in dept_buyers.get(dept, []))
        else:
            sd = S3.get((mn, buyer), [0,0,0])
            dq, da, dn = sd
        vals = [int(dq), k(da), k(dn)]
        for mi, v in enumerate(vals):
            cell = ws3.cell(row=r, column=c+mi, value=v)
            cell.font = num_font; cell.border = thin
            cell.alignment = Alignment(horizontal='right')
    r += 1

# Total row
ws3.cell(row=r, column=1, value='총합계').font = Font(bold=True, size=10, name='Calibri')
ws3.cell(row=r, column=1).border = thin
for (dept, buyer) in buyer_order3:
    c = col3[(dept, buyer)]
    if buyer == 'S.TOTAL':
        dq = sum(S3.get((mn, b), [0,0,0])[0] for mn in months_list for b in dept_buyers.get(dept, []))
        da = sum(S3.get((mn, b), [0,0,0])[1] for mn in months_list for b in dept_buyers.get(dept, []))
        dn = sum(S3.get((mn, b), [0,0,0])[2] for mn in months_list for b in dept_buyers.get(dept, []))
    else:
        dq = sum(S3.get((mn, buyer), [0,0,0])[0] for mn in months_list)
        da = sum(S3.get((mn, buyer), [0,0,0])[1] for mn in months_list)
        dn = sum(S3.get((mn, buyer), [0,0,0])[2] for mn in months_list)
    vals = [int(dq), k(da), k(dn)]
    for mi, v in enumerate(vals):
        cell = ws3.cell(row=r, column=c+mi, value=v)
        cell.font = Font(bold=True, size=10, name='Calibri'); cell.border = thin
        cell.alignment = Alignment(horizontal='right')
print('Sheet 3 done')

# ======================================================
# SHEET 4: 시즌 부서공장
# ======================================================
ws4 = wb.create_sheet('4.시즌부서공장')
ws4.cell(row=1, column=1, value='2026 Season 부서별 브랜드별 공장별 배정 수량').font = title_font
ws4.merge_cells(start_row=1, start_column=1, end_row=1, end_column=15)

for ci, h in enumerate(['부서', '브랜드', 'Factory'], 1):
    ws4.cell(row=3, column=ci, value=h).font = hdr_font
col4 = {}; c = 4
for s in seasons_list:
    ws4.cell(row=3, column=c, value=s).font = hdr_font
    ws4.cell(row=3, column=c+1, value='비중').font = hdr_font
    col4[s] = c; c += 2
ws4.cell(row=3, column=c, value='Total').font = hdr_font
ws4.cell(row=3, column=c+1, value='비중').font = hdr_font
tc4 = c

r = 4
for dn in DEPT_ORDER:
    prefix = dn.split('-')[0]
    items = [(k, v) for k, v in S4.items() if k[0].startswith(prefix)]
    if not items: continue
    dg = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    for (dp, br, fac, s), q in items:
        dg[br][fac][s] += q
    for brand in sorted(dg.keys()):
        facs = sorted(dg[brand].keys())
        for fi, fac in enumerate(facs):
            ws4.cell(row=r, column=1, value=dn if fi == 0 else '')
            ws4.cell(row=r, column=2, value=brand if fi == 0 else '')
            ws4.cell(row=r, column=3, value=fac)
            brand_total = sum(dg[brand][fac].get(s, 0) for s in seasons_list)
            for s in seasons_list:
                q = dg[brand][fac].get(s, 0)
                ws4.cell(row=r, column=col4[s], value=int(q))
                ws4.cell(row=r, column=col4[s]+1, value=round(q/brand_total, 4) if brand_total else 0)
            total_q = brand_total
            ws4.cell(row=r, column=tc4, value=int(total_q))
            ws4.cell(row=r, column=tc4+1, value=1.0)
            for ci in range(1, tc4+2):
                ws4.cell(row=r, column=ci).border = thin
                ws4.cell(row=r, column=ci).font = num_font
                ws4.cell(row=r, column=ci).alignment = Alignment(horizontal='right')
            ws4.cell(row=r, column=1).alignment = Alignment(horizontal='left')
            ws4.cell(row=r, column=2).alignment = Alignment(horizontal='left')
            ws4.cell(row=r, column=3).alignment = Alignment(horizontal='left')
            r += 1
print('Sheet 4 done')

# ======================================================
# SHEET 5: 시즌 공장
# ======================================================
ws5 = wb.create_sheet('5.시즌공장')
ws5.cell(row=1, column=1, value='2026 Season 공장별 브랜드별 배정 수량').font = title_font
ws5.merge_cells(start_row=1, start_column=1, end_row=1, end_column=15)

for ci, h in enumerate(['Factory', '브랜드'], 1):
    ws5.cell(row=3, column=ci, value=h).font = hdr_font
col5 = {}; c = 3
for s in seasons_list:
    ws5.cell(row=3, column=c, value=s).font = hdr_font
    ws5.cell(row=3, column=c+1, value='비중').font = hdr_font
    col5[s] = c; c += 2
ws5.cell(row=3, column=c, value='Total').font = hdr_font
ws5.cell(row=3, column=c+1, value='비중').font = hdr_font
tc5 = c

r = 4
fg = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
for (fac, br, s), q in S5.items():
    fg[fac][br][s] += q

for factory in sorted(fg.keys()):
    brands = sorted(fg[factory].keys())
    for bi, brand in enumerate(brands):
        ws5.cell(row=r, column=1, value=factory if bi == 0 else '')
        ws5.cell(row=r, column=2, value=brand)
        fac_brand_total = sum(fg[factory][brand].get(s, 0) for s in seasons_list)
        for s in seasons_list:
            q = fg[factory][brand].get(s, 0)
            ws5.cell(row=r, column=col5[s], value=int(q))
            ws5.cell(row=r, column=col5[s]+1, value=round(q/fac_brand_total, 4) if fac_brand_total else 0)
        ws5.cell(row=r, column=tc5, value=int(fac_brand_total))
        ws5.cell(row=r, column=tc5+1, value=1.0)
        for ci in range(1, tc5+2):
            ws5.cell(row=r, column=ci).border = thin
            ws5.cell(row=r, column=ci).font = num_font
            ws5.cell(row=r, column=ci).alignment = Alignment(horizontal='right')
        ws5.cell(row=r, column=1).alignment = Alignment(horizontal='left')
        ws5.cell(row=r, column=2).alignment = Alignment(horizontal='left')
        r += 1
print('Sheet 5 done')

wb.save(OUT)
print(f'\nSaved: {OUT}')
import os; print(f'Size: {os.path.getsize(OUT)/1024:.0f} KB')
