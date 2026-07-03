import openpyxl

src = r'C:\Users\user\.openclaw\workspace\영업데이터파일_v12_20260702_processed.xlsx'
wb = openpyxl.load_workbook(src)
ws = wb.worksheets[0]

# Fix Total rows for BU3-2bu
for col in [18, 19, 20]:
    for section_offset, total_row in [(0, 10), (9, 19), (18, 28)]:
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
            if section_offset == 0:
                ws.cell(total_row, col).value = int(total)
            else:
                ws.cell(total_row, col).value = round(total, 1)

# Fix 평균단가 for BU3-2bu
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
                    ws.cell(season_row, col).value = round((sales_f * 1000) / qty_f, 2)
            except:
                pass

# Recalculate 마진율 with % format
pct_fmt = '0.0"%"'
for col in range(3, 36):
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
                    ws.cell(margin_row, col).number_format = pct_fmt
            except:
                pass

# Update 합계 column
stotal_cols = [8, 11, 14, 17, 20, 23, 25, 30, 32, 34]

fmt_config = [
    ([6,7,8,9,10], '#,##0', 0),
    ([15,16,17,18,19], '#,##0', 0),
    ([24,25,26,27,28], '#,##0', 0),
    ([33,34,35,36,37], '#,##0.00', 2),
    ([42,43,44,45,46], '0.0"%"', 1),
    ([51,52,53,54,55], '#,##0.00', 2),
]

for section_rows, fmt, decimals in fmt_config:
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
            ws.cell(row, 35).value = round(total, decimals)
            ws.cell(row, 35).number_format = fmt

# Save
out1 = src
out2 = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\영업데이터파일_v12_20260702.xlsx'
out3 = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\영업데이터파일_v13_20260702.xlsx'

wb.save(out1)
print(f'Saved: {out1}')

try:
    wb.save(out2)
    print(f'Saved: {out2}')
except Exception as e:
    print(f'v12 save error: {e}')

try:
    wb.save(out3)
    print(f'Saved: {out3}')
except Exception as e:
    print(f'v13 save error: {e}')

# Final verification
print('\n=== 최종 검증 ===')
for col, name in [(18, 'KOHL_S'), (19, 'URBAN'), (20, 'S.TOTAL')]:
    print(f'\n{name} ORDER:')
    for r in [6,7,8,9,10]:
        v = ws.cell(r, col).value
        print(f'  {ws.cell(r,2).value}: {v}')

print('\nMargin rate example (26 SPR C3, C15, C18):')
for c in [3, 15, 16, 18, 19, 20, 35]:
    cell = ws.cell(42, c)
    print(f'  C{c}: {cell.value} ({cell.number_format})')
