import openpyxl

# Load pivot analysis
pivot = r'C:\Users\user\.openclaw\workspace\브랜드별_직접원가_피벗분석_v1.0_20260629.xlsx'
pwb = openpyxl.load_workbook(pivot)
pws = pwb['1.시즌별 부서브랜드 원가']

# Build lookup from pivot: (dept, season) -> qty, amt
pivot_bu3 = {}
for r in range(2, pws.max_row+1):
    dept = str(pws.cell(r, 1).value or '').strip()
    brand = str(pws.cell(r, 2).value or '').strip()
    season = str(pws.cell(r, 3).value or '').strip()
    po_amt = float(pws.cell(r, 4).value or 0)
    po_qty = float(pws.cell(r, 7).value or 0)
    profit = float(pws.cell(r, 6).value or 0)
    direct_cost = float(pws.cell(r, 5).value or 0)
    
    if brand in ('소계', '총계') or not season.startswith('26 '):
        continue
    
    pivot_bu3[(brand, season)] = {'amt': po_amt, 'qty': po_qty, 'profit': profit, 'direct': direct_cost}

pwb.close()

# Map v12 brand columns to pivot brand names
col_to_brand = {
    3: 'GAP', 4: 'GAP IPSS', 5: 'BRFS', 6: 'Madhappy', 7: 'Mott & Bow', 8: 'S.TOTAL',
    9: 'OLD NAVY', 10: 'ON IPSS MEN', 11: 'S.TOTAL',
    12: 'WALMART', 13: 'QUINCE', 14: 'S.TOTAL',
    15: 'BANANA REPUBLIC', 16: 'ON IPSS', 17: 'S.TOTAL',
    18: "KOHL`S", 19: "URBAN", 20: 'S.TOTAL',
    21: 'AERIE', 22: 'Offline', 23: 'S.TOTAL',
    24: 'MLB', 25: 'S.TOTAL',
    26: 'NUUDS', 27: 'ONER ACTIVE', 28: 'ROSS', 29: 'AMERICAN EAGLE', 30: 'S.TOTAL',
    31: 'INGRID & ISABEL', 32: 'S.TOTAL',
    33: 'OLDNAVY(HYBRID)', 34: 'S.TOTAL',
}

# Check which can be found in pivot
print('=== Brand mapping check ===')
seasons = ['26 SPR', '26 SUM', '26 FAL', '26 HOL']
for col, brand_name in sorted(col_to_brand.items()):
    found = 0
    for s in seasons:
        if (brand_name, s) in pivot_bu3:
            found += 1
    if found == 0:
        # Try with different name variations
        for s in seasons:
            for alt_brand in set(k[0] for k in pivot_bu3.keys()):
                if brand_name.lower().replace("'","`").replace("-"," ") in alt_brand.lower().replace("'","`").replace("-"," "):
                    print(f'  C{col}: v12="{brand_name}" -> pivot="{alt_brand}" (fuzzy)')
                    break
            else:
                continue
            break
        else:
            print(f'  C{col}: v12="{brand_name}" -> NOT FOUND in pivot')

print('\nDone.')
