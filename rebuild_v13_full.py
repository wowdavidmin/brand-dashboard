"""Full rebuild of v13 first sheet from raw data"""
import openpyxl
from collections import defaultdict

# Step 1: Aggregate raw data
raw_path = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'
wb_raw = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)
ws_raw = wb_raw.worksheets[0]

data = {}  # (dept, buyer, season) -> {'qty':, 'amt':, 'disc':, 'cmt':}
all_seasons = set()
all_buyers = set()
all_depts = set()

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
    if key not in data:
        data[key] = {'qty': 0, 'amt': 0, 'disc': 0, 'cmt': 0}
    data[key]['qty'] += qty
    data[key]['amt'] += amt
    data[key]['disc'] += disc
    data[key]['cmt'] += cmt
    
    all_seasons.add(season)
    all_buyers.add(buyer)
    all_depts.add(dept)

wb_raw.close()

print(f'Seasons: {sorted(all_seasons)}')
print(f'Buyers: {sorted(all_buyers)}')
print(f'Depts: {sorted(all_depts)}')
print(f'Total keys: {len(data)}')

# 인터텍스 buyer breakdown
print('\n=== 인터텍스 기획부 buyer breakdown ===')
inter_data = {}
inter_qty = 0
for (d, b, s), v in data.items():
    if d == '인터텍스 기획부':
        if b not in inter_data:
            inter_data[b] = {'qty': 0, 'amt': 0}
        inter_data[b]['qty'] += v['qty']
        inter_data[b]['amt'] += v['amt']
        inter_qty += v['qty']

for b, v in sorted(inter_data.items(), key=lambda x: -x[1]['qty']):
    print(f'  {b}: qty={v["qty"]:>8,.0f} amt={v["amt"]/1000:>8,.1f}K')
print(f'  TOTAL: {inter_qty:,.0f}')

# Now, I need to figure out which v13 columns correspond to which raw dept+buyer combos
# Let me check what buyers are in each department
print('\n=== Buyers by department (for v13 mapping) ===')
for d in sorted(all_depts):
    buyers = sorted(set(b for (dept, b, s) in data if dept == d))
    print(f'{d}: {buyers}')
