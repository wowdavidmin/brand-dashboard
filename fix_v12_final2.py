"""Fix KOHL'S (C18) and URBAN/URBAN OUTFITTERS (C19) in v12"""
import openpyxl

pivot_path = r'C:\Users\user\.openclaw\workspace\브랜드별_직접원가_피벗분석_v1.0_20260629.xlsx'
pwb = openpyxl.load_workbook(pivot_path)
pws = pwb['1.시즌별 부서브랜드 원가']

# Extract 영업3부 (BU3) data
pivot_data = {}
for r in range(2, pws.max_row+1):
    dept = str(pws.cell(r, 1).value or '').strip()
    brand = str(pws.cell(r, 2).value or '').strip()
    season = str(pws.cell(r, 3).value or '').strip()
    if dept != '영업3부': continue
    if not season.startswith('26 '): continue
    pivot_data[(brand, season)] = {
        'amt': float(pws.cell(r, 4).value or 0),
        'direct': float(pws.cell(r, 5).value or 0),
        'profit': float(pws.cell(r, 6).value or 0),
        'qty': float(pws.cell(r, 7).value or 0),
    }
pwb.close()

# Fix v12
v12_path = r'C:\Users\user\.openclaw\media\inbound\영업데이터파일_v12_20260702---c5b345bd-47c5-47cf-b54c-e337523c2574.xlsx'
wb = openpyxl.load_workbook(v12_path)
ws = wb.worksheets[0]

# Season mapping
seasons = [
    ('26 SPR', 6, 15, 24),
    ('26 SUM', 7, 16, 25),
    ('26 FAL', 8, 17, 26),
    ('26 HOL', 9, 18, 27),
]

# C18 = KOHL'S in v12 -> KOHL`S in pivot
# C19 = URBAN in v12 -> URBAN OUTFITTERS in pivot
fix_map = {18: "KOHL`S", 19: "URBAN OUTFITTERS"}

for season_name, r_qty, r_sales, r_net in seasons:
    for col, pivot_brand in fix_map.items():
        k = (pivot_brand, season_name)
        if k in pivot_data:
            d = pivot_data[k]
            
            # Fix ORDER 수량
            new_qty = int(d['qty'])
            ws.cell(r_qty, col).value = new_qty
            
            # Fix 총매출액 ($K) - preserve original if 0
            new_amt = round(d['amt'] / 1000, 1)
            ws.cell(r_sales, col).value = new_amt
            
            # Fix 순매출액 - use same discount rate as original
            old_amt = float(ws._cells.get((r_sales, col), openpyxl.cell.Cell(ws, None, None)).value or 0)
            old_net = float(ws._cells.get((r_net, col), openpyxl.cell.Cell(ws, None, None)).value or 0)
            
            # Hmm, the cells were already modified above. Let me re-read differently.
    
# Actually, let me redo this properly
wb.close()

# Re-open fresh
wb = openpyxl.load_workbook(v12_path)
ws = wb.worksheets[0]

for season_name, r_qty, r_sales, r_net in seasons:
    for col, pivot_brand in fix_map.items():
        k = (pivot_brand, season_name)
        if k in pivot_data:
            d = pivot_data[k]
            
            # Read original values
            old_sales = ws.cell(r_sales, col).value
            old_sales_f = float(old_sales if old_sales not in ('-', None, '') else 0)
            old_net_f = float(ws.cell(r_net, col).value if ws.cell(r_net, col).value not in ('-', None, '') else 0)
            
            # Calculate discount ratio from original
            if old_sales_f > 0:
                disc_ratio = (old_sales_f - old_net_f) / old_sales_f
            else:
                disc_ratio = 0
            
            # New values
            new_qty = int(d['qty'])
            new_amt_k = round(d['amt'] / 1000, 1)
            new_net_k = round(new_amt_k * (1 - disc_ratio), 1)
            
            ws.cell(r_qty, col).value = new_qty
            ws.cell(r_sales, col).value = new_amt_k
            ws.cell(r_net, col).value = new_net_k
    
    # Fix S.TOTAL (C20) = KOHL`S + URBAN OUTFITTERS
    kohls = pivot_data.get(("KOHL`S", season_name), {'qty': 0, 'amt': 0, 'profit': 0, 'direct': 0})
    urban = pivot_data.get(("URBAN OUTFITTERS", season_name), {'qty': 0, 'amt': 0, 'profit': 0, 'direct': 0})
    
    total_qty = int(kohls['qty'] + urban['qty'])
    total_amt = round((kohls['amt'] + urban['amt']) / 1000, 1)
    total_direct = round((kohls['direct'] + urban['direct']) / 1000, 1)
    # Net = total_amt - total_direct... no, net = amt - disc
    # We don't have disc. Use profit ratio
    k_profit_r = kohls['profit'] / kohls['amt'] if kohls['amt'] > 0 else 0
    u_profit_r = urban['profit'] / urban['amt'] if urban['amt'] > 0 else 0
    
    if kohls['amt'] + urban['amt'] > 0:
        # Actually 순매출액 is not profit. Let me just keep the same ratio as original for S.TOTAL too
        old_s_total_sales = float(ws.cell(r_sales, 20).value if ws.cell(r_sales, 20).value not in ('-', None, '') else 0)
        old_s_total_net = float(ws.cell(r_net, 20).value if ws.cell(r_net, 20).value not in ('-', None, '') else 0)
        if old_s_total_sales > 0:
            st_disc_r = (old_s_total_sales - old_s_total_net) / old_s_total_sales
        else:
            st_disc_r = 0.01  # 1% default
        
        total_net = round(total_amt * (1 - st_disc_r), 1)
    else:
        total_net = 0
    
    ws.cell(r_qty, 20).value = total_qty
    ws.cell(r_sales, 20).value = total_amt
    ws.cell(r_net, 20).value = total_net

# Save
out = r'C:\Users\user\.openclaw\workspace\영업데이터파일_v12_20260702_processed.xlsx'
wb.save(out)
print(f'Saved to {out}')

# Verify
wb2 = openpyxl.load_workbook(out)
ws2 = wb2.worksheets[0]
print('\n=== Fixed BU3-영업2부 ===')
for col, name in [(18, "KOHL'S"), (19, 'URBAN'), (20, 'S.TOTAL')]:
    print(f'\n{name} (C{col}):')
    for r_qty, r_sales, r_net, sl in [(6,15,24,'SPR'), (7,16,25,'SUM'), (8,17,26,'FAL'), (9,18,27,'HOL')]:
        v1 = ws2.cell(r_qty, col).value
        v2 = ws2.cell(r_sales, col).value
        v3 = ws2.cell(r_net, col).value
        print(f'  {sl:5s} | ORDER={v1:>10.0f} | 매출={v2:>8.1f}K | 순매출={v3:>8.1f}K')
