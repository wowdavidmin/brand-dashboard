"""Rebuild v12 sheet 1 using pivot analysis data"""
import openpyxl
from copy import copy

# Source: existing v12 file (for structure only)
v12_src = r'C:\Users\user\.openclaw\media\inbound\영업데이터파일_v12_20260702---c5b345bd-47c5-47cf-b54c-e337523c2574.xlsx'
# Pivot analysis (correct data)
pivot = r'C:\Users\user\.openclaw\workspace\브랜드별_직접원가_피벗분석_v1.0_20260629.xlsx'

# Load v12 to preserve structure
wb = openpyxl.load_workbook(v12_src)
ws = wb.worksheets[0]

# Load pivot data
pwb = openpyxl.load_workbook(pivot)
pws_pivot = pwb['1.시즌별 부서브랜드 원가']

# Build lookup: (brand, season) -> {amt, qty, disc, cmt}
pivot_data = {}
for r in range(2, pws_pivot.max_row+1):
    brand = str(pws_pivot.cell(r, 2).value or '').strip()
    season = str(pws_pivot.cell(r, 3).value or '').strip()
    if brand in ('소계', '총계') or not season.startswith('26 '):
        continue
    amt = float(pws_pivot.cell(r, 4).value or 0)
    direct = float(pws_pivot.cell(r, 5).value or 0)
    profit = float(pws_pivot.cell(r, 6).value or 0)
    qty = float(pws_pivot.cell(r, 7).value or 0)
    pivot_data[(brand, season)] = {'amt': amt, 'qty': qty, 'direct': direct, 'profit': profit}

pwb.close()

# v12 column -> brand mapping (with corrections)
col_brand_map = {
    3: 'GAP', 4: None, 5: 'BRFS', 6: 'Madhappy', 7: 'Mott & Bow',
    9: 'OLD NAVY', 10: None,  # ON MEN'S - might not be in pivot
    12: None, 13: 'QUINCE',   # WALMART might be different
    15: 'BANANA REPUBLIC', 16: 'ON IPSS',
    18: "KOHL`S", 19: "URBAN OUTFITTERS",
    21: 'AERIE', 22: 'Offline',
    24: 'MLB',
    26: 'NUUDS', 27: 'ONER ACTIVE', 28: 'ROSS', 29: 'AMERICAN EAGLE',
    31: 'INGRID & ISABEL',
    33: 'OLDNAVY(HYBRID)',
}

# Check what we can find
print('=== Data availability check ===')
seasons = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
missing_brands = []
for col, brand in sorted(col_brand_map.items()):
    if brand is None:
        continue
    found = sum(1 for s in seasons if (brand, s) in pivot_data)
    if found == 0:
        missing_brands.append((col, brand))
        print(f'  C{col}: {brand} -> NOT FOUND')

# For C4 (GAP IPSS) and C10 (ON MEN'S) and C12 (WALMART)
# Check what data exists for these columns in the pivot
for search in ['GAP', 'ON', 'WAL', 'WAL-MART', 'WALMART']:
    matches = [(k, v) for k, v in pivot_data.items() if search.lower() in k[0].lower()]
    if matches:
        print(f'  Searching "{search}": {set(k[0] for k,v in matches)}')

print()
print(f'Missing brands: {missing_brands}')

# These brands might have different names in pivot
# C4 = GAP IPSS → might be part of GAP or not in pivot
# C10 = ON MEN'S → might be part of ON IPSS in pivot  
# C12 = WALMART → might be WAL-MART in pivot
# Let me check
for alt_brand in ['GAP', 'IPSS GAP', 'GAP IPSS', 'ON MEN', 'ON IPSS MEN', 'WAL-MART', 'WALMART']:
    for s in seasons:
        if (alt_brand, s) in pivot_data:
            print(f'  Found: {alt_brand} {s} -> amt={pivot_data[(alt_brand,s)]["amt"]/1000:.1f}K qty={pivot_data[(alt_brand,s)]["qty"]:.0f}')
