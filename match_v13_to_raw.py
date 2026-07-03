"""Update v13 to match raw data PO totals"""
import openpyxl
from collections import defaultdict

# Step 1: Aggregate raw data by dept + season
raw_path = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'
wb_raw = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)
ws_raw = wb_raw.worksheets[0]

raw_dept_season = defaultdict(lambda: {'qty': 0, 'amt': 0, 'disc': 0})

for row in ws_raw.iter_rows(min_row=2, values_only=True):
    if len(row) < 24: continue
    season = str(row[10] or '') if row[10] else ''
    dept = str(row[4] or '') if row[4] else ''
    if not season.startswith('26 '): continue
    
    qty = float(row[23] or 0) if row[23] else 0
    amt = float(row[24] or 0) if row[24] else 0
    disc = float(row[25] or 0) if row[25] else 0
    
    key = (dept, season)
    raw_dept_season[key]['qty'] += qty
    raw_dept_season[key]['amt'] += amt
    raw_dept_season[key]['disc'] += disc

wb_raw.close()

# Dept mapping
dept_map = {
    'BU1_D1': 'BU1', 'BU2_D1': 'BU2_1', 'BU2_D2': 'BU2_2',
    'BU3_D1': 'BU3_1', 'BU3_D2': 'BU3_2', 'BU3_D3': 'BU3_3', 'BU3_D4': 'BU3_4',
    'BU4': 'BU4', 'BU5': 'BU5', 'GTBU': 'GTBU',
    '인터텍스 기획부': '인터텍스'
}

# v13 STOTAL column mapping
stotal_to_dept = {
    8: 'BU1',     11: 'BU2_1',    14: 'BU2_2',
    17: 'BU3_1',   20: 'BU3_2',   23: 'BU3_3',
    25: 'BU3_4',   30: 'BU4',      32: 'BU5',
    34: 'GTBU',
}

# Step 2: Update v13
v13_path = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\영업데이터파일_v13_20260702.xlsx'
wb = openpyxl.load_workbook(v13_path)
ws = wb.worksheets[0]

# Show current vs expected for each BU
seasons = ['26 SPR', '26 SUM', '26 FAL', '26 HOL', '26 BTS']
season_rows = {
    '26 SPR': 6, '26 SUM': 7, '26 FAL': 8, '26 HOL': 9
}
# Note: v13 doesn't have 26 BTS row

print('=== Current vs Expected (ORDER qty) ===')
for col, dept_code in sorted(stotal_to_dept.items()):
    brand = ws.cell(5, col).value
    curr_total = 0
    expected_total = 0
    for s in ['26 SPR', '26 SUM', '26 FAL', '26 HOL']:
        row = season_rows[s]
        curr_val = ws.cell(row, col).value
        curr_val = float(curr_val if curr_val not in ('-', None, '') else 0)
        curr_total += curr_val
        
        rd_key = (list(dept_map.keys())[list(dept_map.values()).index(dept_code)], s) if dept_code in dept_map.values() else None
        expected = 0
        for rd, rd_s in raw_dept_season:
            if dept_map.get(rd) == dept_code and rd_s == s:
                expected += raw_dept_season[(rd, rd_s)]['qty']
        expected_total += expected
    
    # Also include BTS
    for rd, rd_s in raw_dept_season:
        if dept_map.get(rd) == dept_code and rd_s == '26 BTS':
            expected_total += raw_dept_season[(rd, rd_s)]['qty']
    
    print(f'  C{col:2d} ({brand:20s}): v13={curr_total:>10,.0f} | raw={expected_total:>10,.0f} | diff={expected_total-curr_total:>+10,.0f}')

# Also check 인터텍스
inter_qty = sum(v['qty'] for (d, s), v in raw_dept_season.items() if d == '인터텍스 기획부')
print(f'  인터텍스 기획부 (not in v13): raw={inter_qty:>10,.0f}')

total_kohls_c18 = 0
expected_kohls = 0
for s in ['26 SPR', '26 SUM', '26 FAL', '26 HOL']:
    row = season_rows[s]
    curr_val = ws.cell(row, 18).value
    curr_val = float(curr_val if curr_val not in ('-', None, '') else 0)
    total_kohls_c18 += curr_val
    
    for rd, rd_s in raw_dept_season:
        if dept_map.get(rd) == 'BU3_2':
            expected_kohls += raw_dept_season[(rd, rd_s)]['qty']

print(f'\n  BU3_2 (C20): v13={total_kohls_c18:,.0f} | raw={expected_kohls:,.0f}')

wb.close()
