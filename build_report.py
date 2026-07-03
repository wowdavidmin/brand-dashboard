import pickle
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side
from collections import defaultdict
import os

# Load aggregated data
with open('C:\\업무자료_약진&JS\\업무 인수\\영업데이터파일\\2026년 6월\\4th week\\_report_v2.pkl', 'rb') as f:
    d = pickle.load(f)

# Re-aggregate from file 2 for buyer/season level
f2 = 'C:\\업무자료_약진&JS\\업무 인수\\영업데이터파일\\2026년 6월\\4th week\\2) 영업DATA_Raw_v1.3_20260608_openclaw 작업용.xlsx'
wb2 = openpyxl.load_workbook(f2, read_only=True, data_only=True)
ws = wb2['Raw_20260608']

sheet1_data = defaultdict(lambda: [0, 0, 0])
sheet2_data = defaultdict(lambda: [0, 0, 0])
sheet3_data = defaultdict(lambda: [0, 0, 0])
sheet4_data = defaultdict(lambda: 0)
sheet1_all_buyers = set()
sheet2_brands = set()
sheet1_all_seasons = {'26 SPR', '26 SUM', '26 FAL', '26 HOL'}

row_cnt = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        continue
    v = list(row)
    row_cnt += 1

    qty = v[26] or 0
    amt = v[27] or 0
    disc = v[28] or 0
    net = v[29] or 0

    buyer = str(v[3] or '').strip()
    brand = str(v[4] or '').strip()
    season = str(v[12] or '').strip()
    dept_raw = str(v[5] or '').strip()
    factory = str(v[2] or '').strip()
    month_str = str(v[21] or '').strip()

    if not (qty and amt and season in sheet1_all_seasons):
        continue

    sheet1_data[(season, buyer)][0] += qty
    sheet1_data[(season, buyer)][1] += amt
    sheet1_data[(season, buyer)][2] += net
    sheet1_all_buyers.add(buyer)

    sheet2_data[(brand, season)][0] += qty
    sheet2_data[(brand, season)][1] += amt
    sheet2_data[(brand, season)][2] += net
    sheet2_brands.add(brand)

    if month_str:
        sheet3_data[(brand, month_str)][0] += qty
        sheet3_data[(brand, month_str)][1] += amt
        sheet3_data[(brand, month_str)][2] += net

    sheet4_data[(dept_raw, brand, factory, season)] += qty

    if row_cnt % 20000 == 0:
        print(f'  aggregating: {row_cnt}...')

wb2.close()
print(f'Done: {row_cnt} rows')

# Output workbook
out = 'C:\\업무자료_약진&JS\\업무 인수\\영업데이터파일\\2026년 6월\\4th week\\영업DATA_4Core_20260630.xlsx'
wb = openpyxl.Workbook()
thin = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
hdr_font = Font(bold=True, size=11)
title_font = Font(bold=True, size=13)

seasons_list = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']

def guess_dept(buyer):
    b = buyer.upper()
    if 'GAP' in b and 'IPSS' not in b:
        return 'BU1-영업1부'
    if 'IPSS GAP' in b:
        return 'BU1-영업1부'
    if b in ('BRFS',):
        return 'BU1-영업1부'
    if b in ('MADHAPPY',):
        return 'BU1-영업1부'
    if 'MOTT' in b or 'BOW' in b:
        return 'BU1-영업1부'
    if 'OLD NAVY' in b:
        return 'BU2-영업1부'
    if b.startswith('ON ') or 'ON MEN' in b or 'ON WOM' in b or 'ON KID' in b:
        return 'BU2-영업1부'
    if 'WALMART' in b:
        return 'BU2-영업2부'
    if 'QUINCE' in b:
        return 'BU2-영업2부'
    if b in ('BR',):
        return 'BU3-영업1부'
    if 'IPSS ON' in b:
        return 'BU3-영업1부'
    if 'KOHL' in b:
        return 'BU3-영업2부'
    if 'URBAN' in b:
        return 'BU3-영업2부'
    if 'AERIE' in b:
        return 'BU3-영업3부'
    if 'OFFLINE' in b:
        return 'BU3-영업3부'
    if 'MLB' in b:
        return 'BU3-영업4부'
    return 'BU4부'

# Group brands by dept
dept_brands = defaultdict(set)
for brand in sheet2_brands:
    b = brand.upper()
    if 'GAP' in b and 'IPSS' not in b:
        dept_brands['BU1-영업1부'].add(brand)
    elif 'IPSS GAP' in b:
        dept_brands['BU1-영업1부'].add(brand)
    elif b in ('BRFS',):
        dept_brands['BU1-영업1부'].add(brand)
    elif b in ('MADHAPPY',):
        dept_brands['BU1-영업1부'].add(brand)
    elif 'MOTT' in b or 'BOW' in b:
        dept_brands['BU1-영업1부'].add(brand)
    elif 'OLD NAVY' in b:
        dept_brands['BU2-영업1부'].add(brand)
    elif b.startswith('ON ') or 'ON MEN' in b or 'ON WOM' in b or 'ON KID' in b:
        dept_brands['BU2-영업1부'].add(brand)
    elif 'WALMART' in b:
        dept_brands['BU2-영업2부'].add(brand)
    elif 'QUINCE' in b:
        dept_brands['BU2-영업2부'].add(brand)
    elif b in ('BR',):
        dept_brands['BU3-영업1부'].add(brand)
    elif 'IPSS ON' in b:
        dept_brands['BU3-영업1부'].add(brand)
    elif 'KOHL' in b:
        dept_brands['BU3-영업2부'].add(brand)
    elif 'URBAN' in b:
        dept_brands['BU3-영업2부'].add(brand)
    elif 'AERIE' in b:
        dept_brands['BU3-영업3부'].add(brand)
    elif 'OFFLINE' in b:
        dept_brands['BU3-영업3부'].add(brand)
    elif 'MLB' in b:
        dept_brands['BU3-영업4부'].add(brand)
    else:
        dept_brands['BU4부'].add(brand)

dept_order = ['BU1-영업1부', 'BU2-영업1부', 'BU2-영업2부', 'BU3-영업1부', 'BU3-영업2부', 'BU3-영업3부', 'BU3-영업4부', 'BU4부']

# =========================================
# SHEET 1: 시즌현황
# =========================================
ws1 = wb.active
ws1.title = '1.시즌현황'

ws1.cell(row=1, column=1, value='[약진통상] 2026 Season 부서별 브랜드별 - 수량/매출/마진').font = title_font
ws1.merge_cells('A1:G1')

info_text = '2026 season 기준 (2021.1 ~ 2026.12 오더기준) | 데이터기준: 2026.06.30'
ws1.cell(row=2, column=1, value=info_text).font = Font(size=9, italic=True)
ws1.merge_cells('A2:G2')

# Headers
r = 4
headers = ['시즌', '부서', '바이어', 'ORDER 수량', '총매출액', '순매출액']
for ci, h in enumerate(headers, 1):
    c = ws1.cell(row=r, column=ci, value=h)
    c.font = hdr_font
    c.border = thin

r = 5
for season in seasons_list:
    season_buyers = [(b, sheet1_data.get((season, b), [0, 0, 0])) for b in sheet1_all_buyers]
    season_buyers = [(b, d) for b, d in season_buyers if d[0] > 0]
    season_buyers.sort(key=lambda x: -x[1][0])

    for buyer, sd in season_buyers:
        ws1.cell(row=r, column=1, value=season)
        ws1.cell(row=r, column=2, value=guess_dept(buyer))
        ws1.cell(row=r, column=3, value=buyer)
        ws1.cell(row=r, column=4, value=int(sd[0]))
        ws1.cell(row=r, column=5, value=round(sd[1], 2))
        ws1.cell(row=r, column=6, value=round(sd[2], 2))
        for ci in range(1, 7):
            ws1.cell(row=r, column=ci).border = thin
        r += 1

    # Season subtotal
    sq = sum(sheet1_data.get((season, b), [0, 0, 0])[0] for b in sheet1_all_buyers)
    sa = sum(sheet1_data.get((season, b), [0, 0, 0])[1] for b in sheet1_all_buyers)
    sn = sum(sheet1_data.get((season, b), [0, 0, 0])[2] for b in sheet1_all_buyers)
    ws1.cell(row=r, column=1, value=season).font = Font(bold=True)
    ws1.cell(row=r, column=2, value='소계').font = Font(bold=True)
    ws1.cell(row=r, column=4, value=int(sq)).font = Font(bold=True)
    ws1.cell(row=r, column=5, value=round(sa, 2)).font = Font(bold=True)
    ws1.cell(row=r, column=6, value=round(sn, 2)).font = Font(bold=True)
    for ci in range(1, 7):
        ws1.cell(row=r, column=ci).border = thin
    r += 1

# Grand Total
tq = sum(sheet1_data[(s, b)][0] for s in seasons_list for b in sheet1_all_buyers)
ta = sum(sheet1_data[(s, b)][1] for s in seasons_list for b in sheet1_all_buyers)
tn = sum(sheet1_data[(s, b)][2] for s in seasons_list for b in sheet1_all_buyers)
ws1.cell(row=r, column=1, value='총합계').font = Font(bold=True, size=12)
ws1.cell(row=r, column=4, value=int(tq)).font = Font(bold=True)
ws1.cell(row=r, column=5, value=round(ta, 2)).font = Font(bold=True)
ws1.cell(row=r, column=6, value=round(tn, 2)).font = Font(bold=True)
for ci in range(1, 7):
    ws1.cell(row=r, column=ci).border = thin

# Set column widths
ws1.column_dimensions['A'].width = 12
ws1.column_dimensions['B'].width = 14
ws1.column_dimensions['C'].width = 22
ws1.column_dimensions['D'].width = 14
ws1.column_dimensions['E'].width = 16
ws1.column_dimensions['F'].width = 16

print('Sheet 1 done')

# =========================================
# SHEET 2: 시즌별
# =========================================
ws2 = wb.create_sheet('2.시즌별')
ws2.cell(row=1, column=1, value='[약진통상] 2026 Season별 부서별 브랜드디비전별 매출수량/매출액').font = title_font
ws2.merge_cells('A1:P1')

col_map2 = {}  # season -> (col_qty, col_amt, col_net)
c = 2
for s in seasons_list:
    ws2.cell(row=3, column=c, value=s).font = hdr_font
    ws2.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+2)
    ws2.cell(row=4, column=c, value='수량').font = hdr_font
    ws2.cell(row=4, column=c+1, value='총매출액').font = hdr_font
    ws2.cell(row=4, column=c+2, value='순매출액').font = hdr_font
    col_map2[s] = (c, c+1, c+2)
    c += 3

# Total
ws2.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+3)
ws2.cell(row=3, column=c, value='2026 SEASON TOTAL').font = hdr_font
ws2.cell(row=4, column=c, value='수량').font = hdr_font
ws2.cell(row=4, column=c+1, value='총매출액').font = hdr_font
ws2.cell(row=4, column=c+2, value='순매출액').font = hdr_font
ws2.cell(row=4, column=c+3, value='마진율(%)').font = hdr_font
total_col2 = c

ws2.cell(row=3, column=1, value='브랜드/부서').font = hdr_font
ws2.cell(row=4, column=1, value='').font = hdr_font

r = 5
for dept_name in dept_order:
    brands = sorted(dept_brands.get(dept_name, set()))
    if not brands:
        continue
    for brand in brands:
        ws2.cell(row=r, column=1, value=brand)
        for s in seasons_list:
            sd = sheet2_data.get((brand, s), [0, 0, 0])
            cq, ca, cn_ = col_map2[s]
            ws2.cell(row=r, column=cq, value=int(sd[0]))
            ws2.cell(row=r, column=ca, value=round(sd[1], 2))
            ws2.cell(row=r, column=cn_, value=round(sd[2], 2))
        tq2 = sum(sheet2_data.get((brand, s), [0, 0, 0])[0] for s in seasons_list)
        ta2 = sum(sheet2_data.get((brand, s), [0, 0, 0])[1] for s in seasons_list)
        tn2 = sum(sheet2_data.get((brand, s), [0, 0, 0])[2] for s in seasons_list)
        ws2.cell(row=r, column=total_col2, value=int(tq2))
        ws2.cell(row=r, column=total_col2+1, value=round(ta2, 2))
        ws2.cell(row=r, column=total_col2+2, value=round(tn2, 2))
        margin2 = round((ta2 - tn2) / ta2 * 100, 2) if ta2 else 0
        ws2.cell(row=r, column=total_col2+3, value=margin2)
        for ci in range(1, total_col2 + 4):
            ws2.cell(row=r, column=ci).border = thin
        r += 1

    # Dept subtotal
    ws2.cell(row=r, column=1, value=f'{dept_name} 요약').font = Font(bold=True)
    for s in seasons_list:
        sq2 = sum(sheet2_data.get((b, s), [0, 0, 0])[0] for b in brands)
        sa2 = sum(sheet2_data.get((b, s), [0, 0, 0])[1] for b in brands)
        sn2 = sum(sheet2_data.get((b, s), [0, 0, 0])[2] for b in brands)
        cq, ca, cn_ = col_map2[s]
        ws2.cell(row=r, column=cq, value=int(sq2)).font = Font(bold=True)
        ws2.cell(row=r, column=ca, value=round(sa2, 2)).font = Font(bold=True)
        ws2.cell(row=r, column=cn_, value=round(sn2, 2)).font = Font(bold=True)
    tq2 = sum(sum(sheet2_data.get((b, s), [0, 0, 0])[0] for s in seasons_list) for b in brands)
    ta2 = sum(sum(sheet2_data.get((b, s), [0, 0, 0])[1] for s in seasons_list) for b in brands)
    tn2 = sum(sum(sheet2_data.get((b, s), [0, 0, 0])[2] for s in seasons_list) for b in brands)
    ws2.cell(row=r, column=total_col2, value=int(tq2)).font = Font(bold=True)
    ws2.cell(row=r, column=total_col2+1, value=round(ta2, 2)).font = Font(bold=True)
    ws2.cell(row=r, column=total_col2+2, value=round(tn2, 2)).font = Font(bold=True)
    margin2 = round((ta2 - tn2) / ta2 * 100, 2) if ta2 else 0
    ws2.cell(row=r, column=total_col2+3, value=margin2).font = Font(bold=True)
    for ci in range(1, total_col2 + 4):
        ws2.cell(row=r, column=ci).border = thin
    r += 1

ws2.column_dimensions['A'].width = 22
for c_idx in range(2, total_col2 + 4):
    ws2.column_dimensions[chr(64 + c_idx) if c_idx <= 26 else 'A'].width = 14

print('Sheet 2 done')

# =========================================
# SHEET 3: 월별
# =========================================
ws3 = wb.create_sheet('3.월별')
ws3.cell(row=1, column=1, value='[약진통상] 2026년 월별 부서별 브랜드디비전별 매출수량/매출액').font = title_font
ws3.merge_cells('A1:AH1')

months_list = [f'{m}월' for m in range(1, 13)]
col_map3 = {}
c = 2
for mn in months_list:
    ws3.cell(row=3, column=c, value=mn).font = hdr_font
    ws3.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c+2)
    ws3.cell(row=4, column=c, value='수량').font = hdr_font
    ws3.cell(row=4, column=c+1, value='총매출액').font = hdr_font
    ws3.cell(row=4, column=c+2, value='순매출액').font = hdr_font
    col_map3[mn] = (c, c+1, c+2)
    c += 3

ws3.cell(row=3, column=1, value='브랜드/부서').font = hdr_font
total_col3 = c

r = 5
for dept_name in dept_order:
    brands = sorted(dept_brands.get(dept_name, set()))
    if not brands:
        continue
    for brand in brands:
        ws3.cell(row=r, column=1, value=brand)
        for mn in months_list:
            sd = sheet3_data.get((brand, mn), [0, 0, 0])
            cq, ca, cn_ = col_map3[mn]
            ws3.cell(row=r, column=cq, value=int(sd[0]))
            ws3.cell(row=r, column=ca, value=round(sd[1], 2))
            ws3.cell(row=r, column=cn_, value=round(sd[2], 2))
        for ci in range(1, total_col3):
            ws3.cell(row=r, column=ci).border = thin
        r += 1

    # Dept subtotal
    ws3.cell(row=r, column=1, value=f'{dept_name} 요약').font = Font(bold=True)
    for mn in months_list:
        sq3 = sum(sheet3_data.get((b, mn), [0, 0, 0])[0] for b in brands)
        sa3 = sum(sheet3_data.get((b, mn), [0, 0, 0])[1] for b in brands)
        sn3 = sum(sheet3_data.get((b, mn), [0, 0, 0])[2] for b in brands)
        cq, ca, cn_ = col_map3[mn]
        ws3.cell(row=r, column=cq, value=int(sq3)).font = Font(bold=True)
        ws3.cell(row=r, column=ca, value=round(sa3, 2)).font = Font(bold=True)
        ws3.cell(row=r, column=cn_, value=round(sn3, 2)).font = Font(bold=True)
    for ci in range(1, total_col3):
        ws3.cell(row=r, column=ci).border = thin
    r += 1

print('Sheet 3 done')

# =========================================
# SHEET 4: 시즌 부서.공장
# =========================================
ws4 = wb.create_sheet('4.시즌부서공장')
ws4.cell(row=1, column=1, value='[약진통상] 2026 Season 부서별 브랜드별 공장별 오더 수량').font = title_font
ws4.merge_cells('A1:K1')

c = 1
ws4.cell(row=3, column=c, value='부서').font = hdr_font; c += 1
ws4.cell(row=3, column=c, value='브랜드').font = hdr_font; c += 1
ws4.cell(row=3, column=c, value='Factory').font = hdr_font; c += 1
col_map4 = {}
for s in seasons_list:
    ws4.cell(row=3, column=c, value=s).font = hdr_font
    ws4.cell(row=3, column=c+1, value='비중').font = hdr_font
    col_map4[s] = c
    c += 2
ws4.cell(row=3, column=c, value='Total').font = hdr_font
ws4.cell(row=3, column=c+1, value='비중').font = hdr_font
total_col4 = c

r = 4
for dept_name in dept_order:
    dept_prefix = dept_name.split('-')[0] if '-' in dept_name else dept_name[:3]
    items = [(k, v) for k, v in sheet4_data.items() if k[0].startswith(dept_prefix)]
    if not items:
        continue
    dg = defaultdict(lambda: defaultdict(lambda: defaultdict(int)))
    for (dp, brand, factory, season), qty in items:
        dg[brand][factory][season] += qty

    for brand in sorted(dg.keys()):
        factories = sorted(dg[brand].keys())
        for fi, factory in enumerate(factories):
            ws4.cell(row=r, column=1, value=dept_name if fi == 0 else '')
            ws4.cell(row=r, column=2, value=brand if fi == 0 else '')
            ws4.cell(row=r, column=3, value=factory)
            for s in seasons_list:
                q = dg[brand][factory].get(s, 0)
                ws4.cell(row=r, column=col_map4[s], value=int(q))
            t = sum(dg[brand][factory].get(s, 0) for s in seasons_list)
            ws4.cell(row=r, column=total_col4, value=int(t))
            for ci in range(1, total_col4 + 2):
                ws4.cell(row=r, column=ci).border = thin
            r += 1

print('Sheet 4 done')

wb.save(out)
print(f'\nSaved: {out} ({os.path.getsize(out) / 1024:.0f} KB)')
