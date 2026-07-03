"""Rebuild v13 first sheet to match raw data PO totals"""
import openpyxl
from collections import defaultdict

# Raw data
raw_path = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'
wb_raw = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)
ws_raw = wb_raw.worksheets[0]

# Aggregate raw data: (dept, buyer, season) -> {qty, amt, disc, cmt}
data = {}

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

wb_raw.close()

print(f'Total keys: {len(data)}')

# Print summary by dept+season
dept_season = defaultdict(lambda: {'qty': 0, 'amt': 0})
for (dept, buyer, season), v in data.items():
    ds = f'{dept}_{season}'
    dept_season[ds]['qty'] += v['qty']
    dept_season[ds]['amt'] += v['amt']

for ds, v in sorted(dept_season.items()):
    print(f'{ds:30s}: qty={v["qty"]:>10,.0f} amt={v["amt"]/1000:>10,.1f}K')

# Total
total_qty = sum(v['qty'] for (d,b,s), v in data.items())
total_amt = sum(v['amt'] for (d,b,s), v in data.items())
print(f'\n{"GRAND TOTAL":30s}: qty={total_qty:>10,.0f} amt={total_amt/1000:>10,.1f}K')
