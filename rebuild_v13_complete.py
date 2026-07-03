"""Complete rebuild of v13 to match raw data"""
import openpyxl
from collections import defaultdict
from copy import copy

raw_path = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'

# ========== STEP 1: Aggregate raw data ==========
wb_raw = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)
ws_raw = wb_raw.worksheets[0]

# Aggregate by (dept, buyer, season)
raw = {}  # (dept, buyer, season) -> {qty, amt, disc, cmt}

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
print(f'Aggregated {len(raw)} keys from raw data')

# ========== STEP 2: Map raw data to v13 structure ==========
# v13 column mapping: (col, bu_dept, v13_brand_name, raw_buyers_expected)
# For brand-level detail, keep original v13 proportions but adjust S.TOTAL to match raw

# First get raw dept+season totals
def get_raw_total(dept, season, field='qty'):
    """Get total qty/amt for a department+season combination"""
    total = 0
    for (d, b, s), v in raw.items():
        if d == dept and s == season:
            total += v[field]
    return total

def get_raw_total_buyer(dept, buyer_keywords, season, field='qty'):
    """Get total for buyers matching keywords within a department"""
    total = 0
    for (d, b, s), v in raw.items():
        if d == dept and s == season:
            for kw in buyer_keywords:
                if kw.lower() in b.lower():
                    total += v[field]
                    break
    return total

# Map v13 columns to raw data dept+buyer
# v13 brand columns and their raw department mapping
col_mapping = {
    # BU1-영업1부 (BU1_D1)
    3: ('BU1_D1', ['GAP INC', 'MADHAPPY', 'MOTT']),  # GAP + BRFS
    4: ('BU1_D1', []),  # IPSS GAP - not in raw data as separate buyer
    5: ('BU1_D1', []),  # BRFS
    6: ('BU1_D1', ['MADHAPPY']),  # Madhappy
    7: ('BU1_D1', ['MOTT']),  # Mott & Bow
    8: ('BU1_D1', []),  # S.TOTAL
    
    # BU2-영업1부 (BU2_D1)
    9: ('BU2_D1', ['GAP INC']),  # OLD NAVY
    10: ('BU2_D1', []),  # ON MEN'S
    11: ('BU2_D1', []),  # S.TOTAL
    
    # BU2-영업2부 (BU2_D2)
    12: ('BU2_D2', ['WALMART']),  # WALMART
    13: ('BU2_D2', ['QUINCE']),  # QUINCE
    14: ('BU2_D2', []),  # S.TOTAL
    
    # BU3-영업1부 (BU3_D1)
    15: ('BU3_D1', ['GAP INC']),  # BR/BANANA REPUBLIC
    16: ('BU3_D1', []),  # IPSS ON
    17: ('BU3_D1', []),  # S.TOTAL
    
    # BU3-영업2부 (BU3_D2)
    18: ('BU3_D2', ["KOHL'S", "KOHL"]),  # KOHL'S
    19: ('BU3_D2', ['URBAN']),  # URBAN OUTFITTERS
    20: ('BU3_D2', []),  # S.TOTAL
    
    # BU3-영업3부 (BU3_D3)
    21: ('BU3_D3', ['AEO']),  # AERIE
    22: ('BU3_D3', []),  # Offline
    23: ('BU3_D3', []),  # S.TOTAL
    
    # BU3-영업4부 (BU3_D4)
    24: ('BU3_D4', ['에프앤에프', 'F&F']),  # MLB
    25: ('BU3_D4', []),  # S.TOTAL
    
    # BU4부 (BU4)
    26: ('BU4', ['NUUDS']),  # NUUDS
    27: ('BU4', ['ASMARA']),  # ONER ACTIVE
    28: ('BU4', ['ROSS']),  # ROSS
    29: ('BU4', ['AEO INC', 'AMERICAN']),  # AEO
    30: ('BU4', []),  # S.TOTAL
    
    # BU5-영업5부 (BU5)
    31: ('BU5', ['C CENTS', 'SPORTS', 'SPA']),  # INGRID & ISABEL
    32: ('BU5', []),  # S.TOTAL
    
    # GTBU-과테말라 (GTBU)
    33: ('GTBU', ['HYBRID']),  # OLDNAVY(HYBRID)
    34: ('GTBU', []),  # S.TOTAL
}

# For brand-level columns, calculate proportions from raw data
print('\nCalculating brand-level data from raw...')
col_data = {}
for col, (dept, keywords) in col_mapping.items():
    if not keywords:  # S.TOTAL or brand that needs proportional calc
        continue
    
    col_data[col] = {}
    for season in ['26 SPR', '26 SUM', '26 FAL', '26 HOL', '26 BTS']:
        qty = get_raw_total_buyer(dept, keywords, season)
        col_data[col][season] = qty

# For S.TOTAL columns, sum all buyers in that dept
stotal_cols = [8, 11, 14, 17, 20, 23, 25, 30, 32, 34]
stotal_to_dept = {8: 'BU1_D1', 11: 'BU2_D1', 14: 'BU2_D2', 17: 'BU3_D1', 
                  20: 'BU3_D2', 23: 'BU3_D3', 25: 'BU3_D4', 
                  30: 'BU4', 32: 'BU5', 34: 'GTBU'}

for col, dept in stotal_to_dept.items():
    col_data[col] = {}
    for season in ['26 SPR', '26 SUM', '26 FAL', '26 HOL', '26 BTS']:
        qty = get_raw_total(dept, season)
        col_data[col][season] = qty

# ========== STEP 3: Rebuild v13 ==========
v13_path = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\영업데이터파일_v13_20260702.xlsx'
wb = openpyxl.load_workbook(v13_path)
ws = wb.worksheets[0]

# Update data rows (6=S_PR, 7=SUM, 8=FAL, 9=HOL)
season_to_row = {'26 SPR': 6, '26 SUM': 7, '26 FAL': 8, '26 HOL': 9}

for season in ['26 SPR', '26 SUM', '26 FAL', '26 HOL']:
    row = season_to_row[season]
    
    # BTS qty should be added to SPR
    add_bts = (season == '26 SPR')
    
    for col, dept in stotal_to_dept.items():
        raw_qty = get_raw_total(dept, season)
        if add_bts:
            raw_qty += get_raw_total(dept, '26 BTS')
        ws.cell(row, col).value = int(raw_qty)
    
    # Also update brand-level columns
    for col, (dept, keywords) in col_mapping.items():
        if not keywords:
            continue
        raw_qty = get_raw_total_buyer(dept, keywords, season)
        if add_bts:
            raw_qty += get_raw_total_buyer(dept, keywords, '26 BTS')
        ws.cell(row, col).value = int(raw_qty)

# Update Total row
for col in range(3, 35):
    total = 0
    has_data = False
    for season_row in [6, 7, 8, 9]:
        v = ws.cell(season_row, col).value
        if v not in ('-', None, ''):
            try:
                total += float(v)
                has_data = True
            except:
                pass
    if has_data:
        ws.cell(10, col).value = int(total)

# ========== STEP 4: Add 인터텍스 컬럼 ==========
# Add as C35 (before 합계) or C36 
# 인터텍스 dept total
inter_qty = {}
for season in ['26 SPR', '26 SUM', '26 FAL', '26 HOL']:
    qty = get_raw_total('인터텍스 기획부', season)
    if season == '26 SPR':
        qty += get_raw_total('인터텍스 기획부', '26 BTS')
    inter_qty[season] = qty

# Add header
ws.cell(4, 36).value = '인터텍스 기획부'
ws.cell(5, 36).value = 'TOTAL'

# Copy style from nearby cell
for r in [4, 5]:
    src = ws.cell(r, 34)
    dst = ws.cell(r, 36)
    if src.font: dst.font = copy(src.font)
    if src.alignment: dst.alignment = copy(src.alignment)

# Add data
for season, row in [('26 SPR', 6), ('26 SUM', 7), ('26 FAL', 8), ('26 HOL', 9)]:
    ws.cell(row, 36).value = int(inter_qty[season])

# Total for 인터텍스
total_inter = sum(inter_qty.values())
ws.cell(10, 36).value = int(total_inter)

# ========== STEP 5: Update 합계 (C35) ==========
# Recalculate 합계 with 인터텍스 included
all_stotal = stotal_cols + [36]
for row in [6, 7, 8, 9, 10]:
    total = 0
    has_data = False
    for c in all_stotal:
        v = ws.cell(row, c).value
        if v not in ('-', None, ''):
            try:
                total += float(v)
                has_data = True
            except:
                pass
    if has_data:
        ws.cell(row, 35).value = int(total)

# ========== SAVE ==========
wb.save(v13_path)
print(f'\nSaved to {v13_path}')

# ========== VERIFY ==========
wb2 = openpyxl.load_workbook(v13_path)
ws2 = wb2.worksheets[0]

print('\n=== Final Verification ===')
raw_total = sum(v['qty'] for v in raw.values())
print(f'Raw data total: {raw_total:>12,.0f}')

v13_total = 0
for c in all_stotal:
    v = ws2.cell(10, c).value
    v13_total += float(v if v not in ('-', None, '') else 0)
print(f'v13 Total (C35): {ws2.cell(10, 35).value:>12,.0f} (from S.TOTAL: {v13_total:,.0f})')

# Buy BUs
print(f'\nBU Breakdown (Total row):')
for col, brand in [(8,'BU1'),(11,'BU2-1'),(14,'BU2-2'),(17,'BU3-1'),
                    (20,'BU3-2'),(23,'BU3-3'),(25,'BU3-4'),
                    (30,'BU4'),(32,'BU5'),(34,'GTBU'),(36,'인터텍스')]:
    val = ws2.cell(10, col).value
    # Expected
    dept = stotal_to_dept.get(col, '인터텍스 기획부')
    exp = sum(sum(v['qty'] for v in [raw[(d,b,s)] for (d,b,s) in raw if d == dept and s in ['26 SPR','26 SUM','26 FAL','26 HOL','26 BTS']] if (d,b,s) in raw))
    if dept != '인터텍스 기획부':
        pass
    print(f'  {brand:10s} C{col:2s}: {val:>12,.0f}')

wb2.close()
print('\nDone!')
