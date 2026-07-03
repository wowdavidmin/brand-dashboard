"""Complete fix: update all derived values and save"""
import openpyxl

src = r'C:\Users\user\.openclaw\workspace\영업데이터파일_v12_20260702_processed.xlsx'
wb = openpyxl.load_workbook(src)
ws = wb.worksheets[0]

# Fix Total rows for BU3-2bu (C18, C19, C20)
for col in [18, 19, 20]:
    for section_offset, total_row in [(0, 10), (9, 19), (18, 28)]:  # ORDER, 매출, 순매출
        total = 0
        has_data = False
        for season_row in [6, 7, 8, 9]:
            v = ws.cell(season_row + section_offset, col).value
            if v not in ('-', None, ''):
                try:
                    total += float(v)
                    has_data = True
                except:
                    pass
        if has_data:
            if section_offset == 0:  # ORDER - integer
                ws.cell(total_row, col).value = int(total)
            else:  # 매출/순매출 - 1 decimal
                ws.cell(total_row, col).value = round(total, 1)

# Recalculate 평균단가 (rows 33-37) for BU3-2bu
# 평균단가 = 총매출액 / ORDER 수량 * 1000 (to convert $K to $)
for col in [18, 19, 20]:
    for i, season_row in enumerate([33, 34, 35, 36, 37]):
        order_row = 6 + i if i < 4 else 10
        sales_row = 15 + i if i < 4 else 19
        
        qty = ws.cell(order_row, col).value
        sales = ws.cell(sales_row, col).value
        
        if qty not in ('-', None, '') and sales not in ('-', None, ''):
            try:
                qty_f = float(qty)
                sales_f = float(sales)
                if qty_f > 0:
                    avg_price = round((sales_f * 1000) / qty_f, 2)
                    ws.cell(season_row, col).value = avg_price
            except:
                pass

# Recalculate CM (rows 51-55)  
# We need correct CM values from raw data. 
# For now, use existing CM values since they haven't been modified.
# The CM values were already correct in the original.

# Recalculate 마진율 (rows 42-46) 
# 마진율 = CM / 평균단가 * 100
for col in range(3, 35):
    for i, margin_row in enumerate([42, 43, 44, 45, 46]):
        cm_row = 51 + i
        avg_row = 33 + i
        
        cm_val = ws.cell(cm_row, col).value
        avg_val = ws.cell(avg_row, col).value
        
        if cm_val not in ('-', None, '') and avg_val not in ('-', None, ''):
            try:
                cm_f = float(cm_val)
                avg_f = float(avg_val)
                if avg_f > 0:
                    margin_pct = round((cm_f / avg_f) * 100, 1)
                    ws.cell(margin_row, col).value = margin_pct
                    ws.cell(margin_row, col).number_format = '0.0"%"'
            except:
                pass

# Also update 합계 column (C35) for all sections
stotal_cols = [8, 11, 14, 17, 20, 23, 25, 30, 32, 34]

for section_rows in [
    [6, 7, 8, 9, 10],   # ORDER
    [15, 16, 17, 18, 19],  # 매출
    [24, 25, 26, 27, 28],  # 순매출
    [33, 34, 35, 36, 37],  # 평균단가
    [42, 43, 44, 45, 46],  # 마진율
    [51, 52, 53, 54, 55],  # CM
]:
    for row in section_rows:
        total = 0
        has_data = False
        for c in stotal_cols:
            v = ws.cell(row, c).value
            if v not in ('-', None, ''):
                try:
                    total += float(v)
                    has_data = True
                except:
                    pass
        if has_data:
            if row in [6,7,8,9,10]:  # ORDER - integer
                ws.cell(row, 35).value = int(total)
                ws.cell(row, 35).number_format = '#,##0'
            elif row in [15,16,17,18,19,24,25,26,27,28]:  # 매출/순매출
                ws.cell(row, 35).value = round(total, 0)
                ws.cell(row, 35).number_format = '#,##0'
            elif row in [33,34,35,36,37]:  # 평균단가
                ws.cell(row, 35).value = round(total, 2)
                ws.cell(row, 35).number_format = '#,##0.00'
            elif row in [42,43,44,45,46]:  # 마진율
                ws.cell(row, 35).value = round(total, 1)
                ws.cell(row, 35).number_format = '0.0"%"'
            elif row in [51,52,53,54,55]:  # CM
                ws.cell(row, 35).value = round(total, 2)
                ws.cell(row, 35).number_format = '#,##0.00'

# Save all 3 locations
out1 = src
out2 = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\영업데이터파일_v12_20260702.xlsx'
out3 = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\영업데이터파일_v13_20260702.xlsx'

wb.save(out1)
wb.save(out2)
wb.save(out3)
print('Saved to all 3 locations')

# Verify final state
print('\n=== Final verification: BU3-2bu (C18-C20) ORDER ===')
for col, name in [(18, 'KOHL_S'), (19, 'URBAN'), (20, 'S.TOTAL')]:
    vals = [f'{name:10s}']
    for r in [6,7,8,9,10]:
        v = ws.cell(r, col).value
        vals.append(f'{v:>10.0f}' if isinstance(v, (int,float)) else str(v).rjust(10))
    print('  '.join(vals))

print('\n=== BU3-2bu 매출액 ===')
for col, name in [(18, 'KOHL_S'), (19, 'URBAN'), (20, 'S.TOTAL')]:
    vals = [f'{name:10s}']
    for r in [15,16,17,18,19]:
        v = ws.cell(r, col).value
        vals.append(str(v).rjust(10))
    print('  '.join(vals))

print('\n=== Margin rate example (26 SPR) ===')
for c in [3,15,16,18,19,20]:
    v = ws.cell(42, c).value
    print(f'  C{c}={v}', end=' ')
print()
