"""Update v13 based on July 2 raw data"""
import pandas as pd
import numpy as np

raw_path = r'C:\Users\user\.openclaw\workspace\영업데이터파일_raw_20260702.xlsx'
v13_template = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 7월\영업데이터파일_v13_20260702.xlsx'
out_dir = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 7월'

print('Reading raw_0608...')
df = pd.read_excel(raw_path, sheet_name='raw_0608', engine='openpyxl')
print(f'Read {len(df)} rows, {len(df.columns)} cols')

# Columns: F=부서(col5), M=시즌(col12), AA=PO수량(col26), AB=PO금액(col27), AC=매출할인(col28), AD=순매출(col29)
# Also need CMT계 for CM calculation
# CMT계 is at column 45 (target col 45 = source col 41)
# Let me find columns by name
print('Columns:', list(df.columns[:35]))

# Find relevant columns
col_부서 = '부서'
col_시즌 = '시즌'
col_수량 = 'PO수량'
col_금액 = 'PO금액'
col_할인 = '매출할인'
col_순매출 = '순매출'
col_cmt = 'CMT 계'

# Map dept codes to v13 BU names
dept_to_bu = {
    'BU1_D1': 'BU1', 'BU2_D1': 'BU2_1', 'BU2_D2': 'BU2_2',
    'BU3_D1': 'BU3_1', 'BU3_D2': 'BU3_2', 'BU3_D3': 'BU3_3', 'BU3_D4': 'BU3_4',
    'BU4': 'BU4', 'BU5': 'BU5', 'GTBU': 'GTBU',
    '인터텍스 기획부': '인터텍스'
}

# Aggregate by dept + season
agg = df.groupby([col_부서, col_시즌]).agg({
    col_수량: 'sum',
    col_금액: 'sum',
    col_할인: 'sum',
    col_순매출: 'sum',
}).reset_index()

# Also get CMT for margin calculation
if col_cmt in df.columns:
    cmt_agg = df.groupby([col_부서, col_시즌])[col_cmt].sum().reset_index()
    agg = agg.merge(cmt_agg, on=[col_부서, col_시즌], how='left')
else:
    agg['CMT 계'] = 0

print(f'\nAggregated {len(agg)} department-season combinations')

# Build v13 data structure
seasons_26 = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
bu_order = ['BU1', 'BU2_1', 'BU2_2', 'BU3_1', 'BU3_2', 'BU3_3', 'BU3_4', 'BU4', 'BU5', 'GTBU', '인터텍스']

# For each BU, compute totals with BTS merged into SPR
bu_data = {}
for bu in bu_order:
    bu_data[bu] = {}
    for s in seasons_26:
        key = f'{bu}_{s}'
        bu_data[bu][s] = {'qty': 0, 'amt': 0, 'disc': 0, 'net': 0, 'cmt': 0}

for _, row in agg.iterrows():
    dept = row[col_부서]
    season = row[col_시즌]
    bu = dept_to_bu.get(dept)
    if bu is None:
        print(f'  WARNING: Unknown dept "{dept}"')
        continue
    
    if season in seasons_26:
        s = season
    elif season == '26 BTS':
        s = '26 SPR'  # Merge BTS into SPR
    else:
        continue
    
    bu_data[bu][s]['qty'] += row[col_수량]
    bu_data[bu][s]['amt'] += row[col_금액]
    bu_data[bu][s]['disc'] += row[col_할인]
    bu_data[bu][s]['net'] += row[col_순매출]
    if col_cmt in row and not pd.isna(row.get(col_cmt, 0)):
        bu_data[bu][s]['cmt'] += row[col_cmt]

# Print summary
print('\n=== BU Totals (with BTS→SPR) ===')
grand_qty = 0
for bu in bu_order:
    bu_qty = sum(bu_data[bu][s]['qty'] for s in seasons_26)
    grand_qty += bu_qty
    print(f'{bu:10s}: qty={bu_qty:>10,.0f}')
print(f'{"TOTAL":10s}: qty={grand_qty:>10,.0f}')

# Now load v13 template and update
print('\nUpdating v13 file...')
import openpyxl
from copy import copy

wb = openpyxl.load_workbook(v13_template)
ws = wb.worksheets[0]

# v13 STOTAL to BU mapping
stotal_to_bu = {8: 'BU1', 11: 'BU2_1', 14: 'BU2_2', 17: 'BU3_1',
                20: 'BU3_2', 23: 'BU3_3', 25: 'BU3_4',
                30: 'BU4', 32: 'BU5', 34: 'GTBU'}
inter_col = 36
season_row = {'26 SPR': 6, '26 SUM': 7, '26 FAL': 8, '26 HOL': 9}

# Update ORDER section (base row = 0 offset)
# ORDER = base 0, 매출 = +9, 순매출 = +18, 평균단가 = +27, 마진율 = +36, CM = +45
for offset, field in [(0, 'qty'), (9, 'amt'), (18, 'net')]:
    for season in seasons_26:
        row = season_row[season] + offset
        
        # Update S.TOTAL columns
        for col, bu in stotal_to_bu.items():
            val = bu_data[bu][season][field]
            if field == 'qty':
                ws.cell(row, col).value = int(val)
            else:
                ws.cell(row, col).value = round(val / 1000, 1)
        
        # Update 인터텍스
        inter_val = bu_data['인터텍스'][season][field]
        if field == 'qty':
            ws.cell(row, inter_col).value = int(inter_val)
        else:
            ws.cell(row, inter_col).value = round(inter_val / 1000, 1)
    
    # Total row
    total_row = 10 + offset
    for col, bu in stotal_to_bu.items():
        total = sum(bu_data[bu][s][field] for s in seasons_26)
        if field == 'qty':
            ws.cell(total_row, col).value = int(total)
        else:
            ws.cell(total_row, col).value = round(total / 1000, 1)
    
    # 인터텍스 total
    inter_total = sum(bu_data['인터텍스'][s][field] for s in seasons_26)
    if field == 'qty':
        ws.cell(total_row, inter_col).value = int(inter_total)
    else:
        ws.cell(total_row, inter_col).value = round(inter_total / 1000, 1)

# Recalculate 평균단가, CM, 마진율
print('Recalculating derived sections...')

# 평균단가 = 총매출액 / ORDER 수량 * 1000
for col in list(stotal_to_bu.keys()) + [inter_col]:
    for i, season in enumerate(seasons_26):
        qty_row = season_row[season]
        sales_row = qty_row + 9
        avg_row = 33 + i
        
        qty = ws.cell(qty_row, col).value
        sales = ws.cell(sales_row, col).value
        if qty not in ('-', None, '') and sales not in ('-', None, '') and float(qty) > 0:
            ws.cell(avg_row, col).value = round(float(sales) * 1000 / float(qty), 2)

# CM = CMT계 / PO수량
bu_by_col = {b: bu for bu, b in [(bu, col) for col, bu in stotal_to_bu.items()]}
# Actually simpler: just use the bu name directly
bu_of_col = {col: bu for col, bu in stotal_to_bu.items()}
bu_of_col[inter_col] = '인터텍스'

for col, bu_name in bu_of_col.items():
    for i, season in enumerate(seasons_26):
        cm_row = 51 + i
        qty = bu_data[bu_name][season]['qty']
        cmt = bu_data[bu_name][season]['cmt']
        
        if qty > 0:
            ws.cell(cm_row, col).value = round(cmt / qty, 2)

# 마진율 = CM / 평균단가 * 100  
pct_fmt = '0.0"%"'
for col in range(3, 37):
    for i, margin_row in enumerate([42, 43, 44, 45, 46]):
        cm_row = 51 + i
        avg_row = 33 + i
        cm_val = ws.cell(cm_row, col).value
        avg_val = ws.cell(avg_row, col).value
        if cm_val not in ('-', None, '') and avg_val not in ('-', None, '') and float(avg_val) > 0:
            margin_pct = round(float(cm_val) / float(avg_val) * 100, 1)
            ws.cell(margin_row, col).value = margin_pct
            ws.cell(margin_row, col).number_format = pct_fmt

# Update 합계 column
all_data_cols = list(stotal_to_bu.keys()) + [inter_col]
fmt_config = [
    ([6,7,8,9,10], '#,##0', 0),
    ([15,16,17,18,19], '#,##0', 0),
    ([24,25,26,27,28], '#,##0', 0),
    ([33,34,35,36,37], '#,##0.00', 2),
    ([42,43,44,45,46], '0.0"%"', 1),
    ([51,52,53,54,55], '#,##0.00', 2),
]

for section_rows, fmt_str, decimals in fmt_config:
    for row in section_rows:
        total = 0
        has_data = False
        for c in all_data_cols:
            v = ws.cell(row, c).value
            if v not in ('-', None, ''):
                try:
                    total += float(v)
                    has_data = True
                except:
                    pass
        if has_data:
            ws.cell(row, 35).value = round(total, decimals)
            ws.cell(row, 35).number_format = fmt_str

# Save
out_v13 = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 7월\영업데이터파일_v13_20260702.xlsx'
ws_v13 = r'C:\Users\user\.openclaw\workspace\영업데이터파일_v13_20260702.xlsx'
wb.save(ws_v13)
print(f'Saved to {ws_v13}')

try:
    wb.save(out_v13)
    print(f'Saved to {out_v13}')
except Exception as e:
    print(f'July folder save failed: {e}')
    print('File may be open in Excel. Close and I will retry.')

# Verify
wb2 = openpyxl.load_workbook(out_v13)
ws2 = wb2.worksheets[0]
print(f'\n=== Final Verification ===')
print(f'ORDER Total: {ws2.cell(10, 35).value:>12,.0f}')
for col, name in [(8,'BU1'),(11,'BU2-1'),(14,'BU2-2'),(17,'BU3-1'),
                  (20,'BU3-2'),(23,'BU3-3'),(25,'BU3-4'),
                  (30,'BU4'),(32,'BU5'),(34,'GTBU'),(36,'인터텍스')]:
    v = ws2.cell(10, col).value
    print(f'  {name:10s}: {v:>12,.0f}')

print(f'\n매출 Total: {ws2.cell(19, 35).value}')
print(f'Margin 26 SPR GAP: {ws2.cell(42,3).value}')
wb2.close()
