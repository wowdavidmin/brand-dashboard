"""Fix KOHL'S and URBAN OUTFITTERS in v12 using pivot data"""
import openpyxl

# Pivot has correct BU3 data (영업3부 only, NO 인터텍스)
pivot_path = r'C:\Users\user\.openclaw\workspace\브랜드별_직접원가_피벗분석_v1.0_20260629.xlsx'
pwb = openpyxl.load_workbook(pivot_path)
pws = pwb['1.시즌별 부서브랜드 원가']

# Extract 영업3부 (BU3) data only for KOHL`S and URBAN OUTFITTERS
pivot_bu3 = {}
for r in range(2, pws.max_row+1):
    dept = str(pws.cell(r, 1).value or '').strip()
    brand = str(pws.cell(r, 2).value or '').strip()
    season = str(pws.cell(r, 3).value or '').strip()
    
    if dept != '영업3부':
        continue
    if brand not in ("KOHL`S", "URBAN OUTFITTERS"):
        continue
    if not season.startswith('26 '):
        continue
    
    amt = float(pws.cell(r, 4).value or 0)  # PO금액
    direct = float(pws.cell(r, 5).value or 0)  # 직접원가
    profit = float(pws.cell(r, 6).value or 0)  # PROFIT
    qty = float(pws.cell(r, 7).value or 0)  # PO수량
    
    pivot_bu3[(brand, season)] = {'amt': amt, 'qty': qty, 'direct': direct, 'profit': profit}

pwb.close()

# Fix the v12 file
v12_path = r'C:\Users\user\.openclaw\media\inbound\영업데이터파일_v12_20260702---c5b345bd-47c5-47cf-b54c-e337523c2574.xlsx'
wb = openpyxl.load_workbook(v12_path)
ws = wb.worksheets[0]

# Season row mapping for various sections
seasons_info = [
    ('26 SPR', 6, 15, 24, 33, 42, 51),
    ('26 SUM', 7, 16, 25, 34, 43, 52),
    ('26 FAL', 8, 17, 26, 35, 44, 53),
    ('26 HOL', 9, 18, 27, 36, 45, 54),
]

# v12 brand -> pivot brand mapping
fix_cols = {
    18: "KOHL`S",   # KOHL'S in v12 -> KOHL`S in pivot
    19: None,       # URBAN in v12 -> but URBAN OUTFITTERS is not in 영업3부 pivot
}

# Also fix S.TOTAL for BU3-2부 (C20)
# Need URBAN data too - check if URBAN OUTFITTERS is in pivot
print('Check URBAN OUTFITTERS in pivot (all depts):')
pwb2 = openpyxl.load_workbook(pivot_path)
pws2 = pwb2['1.시즌별 부서브랜드 원가']
for r in range(2, pws2.max_row+1):
    brand = str(pws2.cell(r, 2).value or '').strip()
    if 'URBAN' in brand.upper():
        dept = str(pws2.cell(r, 1).value or '').strip()
        season = str(pws2.cell(r, 3).value or '').strip()
        print(f'  {dept:15s} {brand:25s} {season:10s} qty={pws2.cell(r,7).value} amt={pws2.cell(r,4).value}')
pwb2.close()

# Build URBAN data by searching all departments
all_pivot = {}
for r in range(2, pws.max_row+1):
    brand = str(pws2.cell(r, 2).value or '').strip()
    season = str(pws2.cell(r, 3).value or '').strip()
    if 'URBAN' in brand.upper() and season.startswith('26 '):
        amt = float(pws2.cell(r, 4).value or 0)
        qty = float(pws2.cell(r, 7).value or 0)
        profit = float(pws2.cell(r, 6).value or 0)
        direct = float(pws2.cell(r, 5).value or 0)
        all_pivot[(brand, season)] = {'amt': amt, 'qty': qty, 'direct': direct, 'profit': profit}
        
        all_pivot[(brand, season)] = {'amt': amt, 'qty': qty, 'direct': direct, 'profit': profit}

# Now fix
col_name = {18: "KOHL`S", 19: "URBAN OUTFITTERS", 20: "S.TOTAL"}
print('\nFixing BU3-영업2부 data...')

for season_name, r_order, r_sales, r_net, r_avg, r_margin, r_cm in seasons_info:
    # Fix KOHL'S (C18)
    k = ("KOHL`S", season_name)
    if k in pivot_bu3:
        d = pivot_bu3[k]
        
        # ORDER qty
        ws.cell(r_order, 18).value = d['qty']
        # 총매출액 ($K)
        ws.cell(r_sales, 18).value = round(d['amt'] / 1000, 1)
        # 순매출액 ($K) - use profit as proxy for net
        net = d['amt'] - d['direct']  # This is profit, which = amt - direct cost
        ws.cell(r_net, 18).value = round(net / 1000, 1)
    
    # Fix URBAN OUTFITTERS (C19)
    k2 = ("URBAN OUTFITTERS", season_name)
    if k2 in all_pivot:
        d = all_pivot[k2]
        ws.cell(r_order, 19).value = d['qty']
        ws.cell(r_sales, 19).value = round(d['amt'] / 1000, 1)
        net = d['amt'] - d['direct']
        ws.cell(r_net, 19).value = round(net / 1000, 1)
    
    # Fix S.TOTAL (C20)
    kohls_d = pivot_bu3.get(k, {'qty': 0, 'amt': 0, 'direct': 0, 'profit': 0})
    urban_d = all_pivot.get(k2, {'qty': 0, 'amt': 0, 'direct': 0, 'profit': 0})
    
    ws.cell(r_order, 20).value = kohls_d['qty'] + urban_d['qty']
    total_amt = kohls_d['amt'] + urban_d['amt']
    total_net = (kohls_d['amt'] - kohls_d['direct']) + (urban_d['amt'] - urban_d['direct'])
    ws.cell(r_sales, 20).value = round(total_amt / 1000, 1)
    ws.cell(r_net, 20).value = round(total_net / 1000, 1)

# Save
out_path = r'C:\Users\user\.openclaw\workspace\영업데이터파일_v12_20260702_processed.xlsx'
wb.save(out_path)
print(f'\nSaved to {out_path}')

# Verify
wb2 = openpyxl.load_workbook(out_path)
ws2 = wb2.worksheets[0]
print('\n=== Verification ===')
for col, name in [(18, "KOHL'S"), (19, 'URBAN'), (20, 'S.TOTAL')]:
    print(f'{name} (C{col}):')
    for r, s in [(6, 'SPR'), (7, 'SUM'), (8, 'FAL'), (9, 'HOL')]:
        v = ws2.cell(r, col).value
        v2 = ws2.cell(r+9, col).value
        v3 = ws2.cell(r+18, col).value
        print(f'  {s:5s} | ORDER={v:>10.0f} | 매출={v2:>8.1f}K | 순매출={v3:>8.1f}K')
