"""Fix all sections in v13 to match raw data"""
import openpyxl
from collections import defaultdict
from copy import copy

raw_path = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'

# Aggregate raw data
wb_raw = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)
ws_raw = wb_raw.worksheets[0]

raw = {}
for row in ws_raw.iter_rows(min_row=2, values_only=True):
    if len(row) < 41: continue
    season = str(row[10] or '') if row[10] else ''
    dept = str(row[4] or '') if row[4] else ''
    buyer = str(row[2] or '') if row[2] else ''
    if not season.startswith('26 '): continue
    
    qty = float(row[23] or 0) if row[23] else 0
    amt = float(row[24] or 0) if row[24] else 0
    disc = float(row[25] or 0) if row[25] else 0
    cmt = float(row[40] or 0) if row[40] else 0
    
    key = (dept, buyer, season)
    if key not in raw:
        raw[key] = {'qty': 0, 'amt': 0, 'disc': 0, 'cmt': 0}
    raw[key]['qty'] += qty
    raw[key]['amt'] += amt
    raw[key]['disc'] += disc
    raw[key]['cmt'] += cmt

wb_raw.close()

def get_dept_total(dept, season, field='qty'):
    total = 0
    for (d, b, s), v in raw.items():
        if d == dept and s == season:
            total += v[field]
    return total

def season_row_map(base_row, season):
    """Map season to row offset from base"""
    mapping = {6: '26 SPR', 7: '26 SUM', 8: '26 FAL', 9: '26 HOL', 10: 'Total'}
    for r, s in mapping.items():
        if s == season:
            return r
    return None

# Dept to STOTAL column mapping
stotal_to_dept = {8: 'BU1_D1', 11: 'BU2_D1', 14: 'BU2_D2', 17: 'BU3_D1', 
                  20: 'BU3_D2', 23: 'BU3_D3', 25: 'BU3_D4', 
                  30: 'BU4', 32: 'BU5', 34: 'GTBU'}
inter_col = 36

seasons = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
season_rows = {'26 SPR': 6, '26 SUM': 7, '26 FAL': 8, '26 HOL': 9}

# Load v13
v13_path = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\영업데이터파일_v13_20260702.xlsx'
wb = openpyxl.load_workbook(v13_path)
ws = wb.worksheets[0]

# For each section, update S.TOTAL columns and 인터텍스
# Section: ORDER=base 0, 총매출액=base +9, 순매출액=base +18

section_map = [
    (0, 'qty', '#,##0', 0),      # ORDER
    (9, 'amt', '#,##0', 0),       # 총매출액 (in $K)
    (18, 'net', '#,##0', 0),      # 순매출액 (in $K, = amt - disc)
]

for offset, field, fmt, decimals in section_map:
    for season in seasons:
        row = season_rows[season] + offset
        bts_season = '26 BTS' if season == '26 SPR' else None
        
        for col, dept in stotal_to_dept.items():
            # Get raw total for this dept+season
            raw_val = get_dept_total(dept, season, 'qty' if field == 'qty' else 'amt')
            if field == 'net':
                raw_val = get_dept_total(dept, season, 'amt') - get_dept_total(dept, season, 'disc')
            
            # Add BTS to SPR
            if bts_season:
                raw_val += get_dept_total(dept, bts_season, 'qty' if field == 'qty' else 'amt')
                if field == 'net':
                    raw_val += get_dept_total(dept, bts_season, 'amt') - get_dept_total(dept, bts_season, 'disc')
            
            if field == 'qty':
                ws.cell(row, col).value = int(raw_val)
            else:
                ws.cell(row, col).value = round(raw_val / 1000, 1)
        
        # 인터텍스
        inter_qty = get_dept_total('인터텍스 기획부', season, 'qty' if field == 'qty' else 'amt')
        if field == 'net':
            inter_qty = get_dept_total('인터텍스 기획부', season, 'amt') - get_dept_total('인터텍스 기획부', season, 'disc')
        if bts_season:
            inter_qty += get_dept_total('인터텍스 기획부', bts_season, 'qty' if field == 'qty' else 'amt')
            if field == 'net':
                inter_qty += get_dept_total('인터텍스 기획부', bts_season, 'amt') - get_dept_total('인터텍스 기획부', bts_season, 'disc')
        
        if field == 'qty':
            ws.cell(row, inter_col).value = int(inter_qty)
        else:
            ws.cell(row, inter_col).value = round(inter_qty / 1000, 1)
    
    # Total row
    total_row = 10 + offset
    all_stotal = list(stotal_to_cols.keys()) + [inter_col] if False else list(stotal_to_dept.keys()) + [inter_col]
    for col in all_stotal:
        total = 0
        has_data = False
        for season in seasons:
            sr = season_rows[season] + offset
            v = ws.cell(sr, col).value
            if v not in ('-', None, ''):
                try:
                    total += float(v)
                    has_data = True
                except:
                    pass
        if has_data:
            if field == 'qty':
                ws.cell(total_row, col).value = int(total)
            else:
                ws.cell(total_row, col).value = round(total, 1)
    
    # 합계 column
    for section_row in [season_rows[s] + offset for s in seasons] + [total_row]:
        total = 0
        for c in all_stotal:
            v = ws.cell(section_row, c).value
            if v not in ('-', None, ''):
                try:
                    total += float(v)
                    has_data = True
                except:
                    pass
        if has_data:
            if field == 'qty':
                ws.cell(section_row, 35).value = int(total)
                ws.cell(section_row, 35).number_format = '#,##0'
            else:
                ws.cell(section_row, 35).value = round(total, 0)
                ws.cell(section_row, 35).number_format = '#,##0'

# Now fix 평균단가, 마진율, CM sections
# 평균단가 = 총매출액 / ORDER 수량 * 1000 (to convert $K to $)
for col in list(stotal_to_dept.keys()) + [inter_col]:
    for i, season in enumerate(seasons):
        order_row = season_rows[season]  # row 6-9
        sales_row = order_row + 9  # row 15-18
        avg_row = 33 + i  # row 33-36
        total_avg_row = 37
        
        qty = ws.cell(order_row, col).value
        sales = ws.cell(sales_row, col).value
        
        if qty not in ('-', None, '') and sales not in ('-', None, ''):
            try:
                qty_f = float(qty)
                sales_f = float(sales)
                if qty_f > 0:
                    avg_price = round((sales_f * 1000) / qty_f, 2)
                    ws.cell(avg_row, col).value = avg_price
            except:
                pass

# CM section - use raw data CMT계
for col in list(stotal_to_dept.keys()) + [inter_col]:
    for i, season in enumerate(seasons):
        cm_row = 51 + i
        total_cm_row = 55
        
        dept = stotal_to_dept.get(col, '인터텍스 기획부')
        cmt = get_dept_total(dept, season, 'cmt')
        if season == '26 SPR':
            cmt += get_dept_total(dept, '26 BTS', 'cmt')
        
        order_qty = ws.cell(season_rows[season], col).value
        order_qty = float(order_qty if order_qty not in ('-', None, '') else 0)
        
        if order_qty > 0:
            cm_per_unit = round(cmt / order_qty, 2)
            ws.cell(cm_row, col).value = cm_per_unit

# 마진율 = CM / 평균단가 * 100
pct_fmt = '0.0"%"'
for col in range(3, 37):
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

# Update 합계 for all sections
all_data_cols = list(stotal_to_dept.keys()) + [inter_col]
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
wb.save(v13_path)

# Also save to v12 backup
v12_path = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\영업데이터파일_v12_20260702.xlsx'
try:
    wb.save(v12_path)
    print(f'Saved to {v12_path}')
except:
    print('v12 skip (maybe open)')

print(f'Saved to {v13_path}')
print('Done!')

# Quick verify
wb2 = openpyxl.load_workbook(v13_path)
ws2 = wb2.worksheets[0]
print(f'\nORDER Total: {ws2.cell(10, 35).value:>12,.0f}')
print(f'매출 Total: {ws2.cell(19, 35).value:>12,.0f}')
print(f'순매출 Total: {ws2.cell(28, 35).value:>12,.0f}')
print(f'Margin 26 SPR GAP: {ws2.cell(42, 3).value}')
print(f'Margin Total GAP: {ws2.cell(46, 3).value}')
wb2.close()
