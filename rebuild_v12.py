import openpyxl
from collections import defaultdict

# Raw data (June 29)
raw_path = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---47ea1d42-8317-4a51-a9fa-d5c2e2764f59.xlsx'
wb = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)
ws = wb.worksheets[0]

# Columns: C4=브랜드, C5=부서, C11=시즌, C17=오더, C24=PO수량, C25=PO금액, C26=매출할인, C41=CMT계, C47=PROFIT
# PO금액 = 총매출액(gross sales)
# 매출할인 = sales discount
# 순매출액 = PO금액 - 매출할인
# PO수량 = order quantity
# CMT계 = contribution margin per unit?
# PROFIT = profit amount

# Build mappings
# 부서 -> brand mapping from v12
# BU1-영업1부: GAP, IPSS GAP, BRFS, Madhappy, Mott & Bow
# BU2-영업1부: OLD NAVY, ON MEN'S
# BU2-영업2부: WALMART, QUINCE
# BU3-영업1부: BR, IPSS ON
# BU3-영업2부: KOHL'S, URBAN  <-- URBAN might not be in raw data
# BU3-영업3부: AERIE, Offline
# BU3-영업4부: MLB
# BU4부: NUUDS, ONER ACTIVE, ROSS, AEO
# BU5-영업5부: INGRID & ISABEL
# GTBU-과테말라: OLDNAVY(HYBRID)

# Let me first check what brands exist in the raw data and their departments
brand_depts = defaultdict(set)
for r in range(2, ws.max_row+1):
    brand = str(ws.cell(r, 4).value or '').strip()
    dept = str(ws.cell(r, 5).value or '').strip()
    if brand:
        brand_depts[brand].add(dept)

print('=== All brands in raw data and their departments ===')
for brand, depts in sorted(brand_depts.items()):
    print(f'  {brand}: {sorted(depts)}')
print()

# Also check seasons
seasons = set()
for r in range(2, ws.max_row+1):
    s = str(ws.cell(r, 11).value or '').strip()
    if s.startswith('26 '):
        seasons.add(s)
print(f'Seasons: {sorted(seasons)}')

wb.close()
