"""Fix KOHL'S data in v12 - remove 인터텍스 department quantities"""
import openpyxl
from copy import copy

# Source files
v12_src = r'C:\Users\user\.openclaw\media\inbound\영업데이터파일_v12_20260702---c5b345bd-47c5-47cf-b54c-e337523c2574.xlsx'
raw_path = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---47ea1d42-8317-4a51-a9fa-d5c2e2764f59.xlsx'

# Step 1: Get correct KOHL'S values from raw data (BU3_D2 only, NOT 인터텍스)
import openpyxl as xl
wb_raw = xl.load_workbook(raw_path, read_only=True, data_only=True)
ws_raw = wb_raw.worksheets[0]

# Aggregate by brand + season + dept
from collections import defaultdict
bu3_data = {}

for row in ws_raw.iter_rows(min_row=2, values_only=True):
    if len(row) < 41: continue
    dept = str(row[4] or '').strip() if row[4] else ''
    season = str(row[10] or '').strip() if row[10] else ''
    buyer = str(row[2] or '').strip() if row[2] else ''
    
    if dept not in ('BU3_D1', 'BU3_D2', 'BU3_D3', 'BU3_D4'):
        continue
    if not season.startswith('26 '):
        continue
    
    qty = float(row[23] or 0) if row[23] else 0
    amt = float(row[24] or 0) if row[24] else 0
    disc = float(row[25] or 0) if row[25] else 0
    cmt = float(row[40] or 0) if row[40] else 0
    
    key = (buyer, season)
    if key not in bu3_data:
        bu3_data[key] = {'qty': 0, 'amt': 0, 'disc': 0, 'cmt': 0}
    bu3_data[key]['qty'] += qty
    bu3_data[key]['amt'] += amt
    bu3_data[key]['disc'] += disc
    bu3_data[key]['cmt'] += cmt

wb_raw.close()

print('=== Correct BU3 data from raw (excludingインタ텍스) ===')
for k, v in sorted(bu3_data.items()):
    print(f'  {k[0]:25s} {k[1]:8s} qty={v["qty"]:>10.0f} amt={v["amt"]/1000:>8.1f}K disc={v["disc"]/1000:>8.1f}K cmt_total={v["cmt"]/1000:>8.1f}K')

# Step 2: Fix v12 file
print('\n=== Fixing v12 file ===')
wb = openpyxl.load_workbook(v12_src)
ws = wb.worksheets[0]

# KOHL'S is in column 18 (C18)
# Fix ORDER (rows 6-10), 총매출액 (rows 15-19), 순매출액 (rows 24-28)
# 평균단가 (rows 33-37), 마진율 (rows 42-46), CM (rows 51-55)

sections = {
    'ORDER': {'rows': [6, 7, 8, 9, 10], 'key': 'qty'},
    '총매출액': {'rows': [15, 16, 17, 18, 19], 'key': 'amt'},
    '순매출액': {'rows': [24, 25, 26, 27, 28], 'key': 'net'},  # amt - disc
}

season_map = {6: '26 SPR', 7: '26 SUM', 8: '26 FAL', 9: '26 HOL', 10: 'Total'}
sales_cols = [(r, s) for r, s in season_map.items()]

# Fix KOHL'S (C18) and URBAN (C19) data
for col, brand in [(18, "KOHL`S"), (19, "URBAN OUTFITTERS")]:
    print(f'\n--- Fixing {brand} (C{col}) ---')
    
    for row, season in season_map.items():
        if season == 'Total':
            continue
        
        key = (brand, season)
        if key in bu3_data:
            d = bu3_data[key]
            
            # ORDER 수량
            old_qty = ws.cell(row, col).value
            old_qty = float(old_qty if old_qty not in ('-', None, '') else 0)
            new_qty = d['qty']
            diff_qty = new_qty - old_qty
            ws.cell(row, col).value = new_qty
            print(f'  ORDER {season}: {old_qty:>10.0f} -> {new_qty:>10.0f} (diff={diff_qty:>+10.0f})')
            
            # 총매출액 (amt in $K)
            gross_row = 15 + (row - 6)  # map 6->15, 7->16, etc.
            old_amt = ws.cell(gross_row, col).value
            old_amt = float(old_amt if old_amt not in ('-', None, '') else 0)
            new_amt = round(d['amt'] / 1000, 1)
            diff_amt = new_amt - old_amt
            ws.cell(gross_row, col).value = new_amt
            print(f'  매출 {season}: {old_amt:>8.1f}K -> {new_amt:>8.1f}K (diff={diff_amt:>+8.1f}K)')
            
            # 순매출액 (net in $K)
            net_row = 24 + (row - 6)
            old_net = ws.cell(net_row, col).value
            old_net = float(old_net if old_net not in ('-', None, '') else 0)
            new_net = round((d['amt'] - d['disc']) / 1000, 1)
            diff_net = new_net - old_net
            ws.cell(net_row, col).value = new_net
            print(f'  순매출 {season}: {old_net:>8.1f}K -> {new_net:>8.1f}K (diff={diff_net:>+8.1f}K)')
        else:
            print(f'  {season}: NOT FOUND in BU3 data')

# Also fix S.TOTAL for BU3-2부 (C20 = KOHL'S + URBAN OUTFITTERS)
print('\n--- Fixing S.TOTAL (C20 = BU3-2부) ---')
for season_label, season_code in [('26 SPR', '26 SPR'), ('26 SUM', '26 SUM'), ('26 FAL', '26 FAL'), ('26 HOL', '26 HOL')]:
    kohls = bu3_data.get(("KOHL`S", season_code), {'qty': 0, 'amt': 0, 'disc': 0})
    urban = bu3_data.get(("URBAN OUTFITTERS", season_code), {'qty': 0, 'amt': 0, 'disc': 0})
    
    row_qty = 6 + ['26 SPR', '26 SUM', '26 FAL', '26 HOL'].index(season_code)
    row_amt = 15 + ['26 SPR', '26 SUM', '26 FAL', '26 HOL'].index(season_code)
    row_net = 24 + ['26 SPR', '26 SUM', '26 FAL', '26 HOL'].index(season_code)
    
    # S.TOTAL qty
    total_qty = kohls['qty'] + urban['qty']
    old_total = ws.cell(row_qty, 20).value
    ws.cell(row_qty, 20).value = total_qty
    print(f'  ORDER {season_code} S.TOTAL: {old_total} -> {total_qty}')
    
    # S.TOTAL amt
    total_amt = round((kohls['amt'] + urban['amt']) / 1000, 1)
    ws.cell(row_amt, 20).value = total_amt
    
    # S.TOTAL net
    total_net = round((kohls['amt'] - kohls['disc'] + urban['amt'] - urban['disc']) / 1000, 1)
    ws.cell(row_net, 20).value = total_net

# Save
out_path = r'C:\Users\user\.openclaw\workspace\영업데이터파일_v12_20260702_processed.xlsx'
wb.save(out_path)
print(f'\nSaved to {out_path}')

# Verify
wb2 = openpyxl.load_workbook(out_path)
ws2 = wb2.worksheets[0]
print('\n=== Verification ===')
for col, name in [(18, 'KOHL\'S'), (19, 'URBAN'), (20, 'S.TOTAL')]:
    print(f'\n{name} (C{col}):')
    for r, s in [(6, 'SPR'), (7, 'SUM'), (8, 'FAL'), (9, 'HOL')]:
        v_order = ws2.cell(r, col).value
        v_sales = ws2.cell(r+9, col).value
        v_net = ws2.cell(r+18, col).value
        print(f'  {s:5s} | ORDER={v_order:>10.0f} | 매출={v_sales:>8.1f}K | 순매출={v_net:>8.1f}K')
