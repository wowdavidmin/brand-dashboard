"""Calculate raw data PO total accurately for 26 season only"""
import openpyxl

f = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'
wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
ws = wb.worksheets[0]

total_26 = 0
total_all = 0
row_count = 0
row_26 = 0

# Also track by PO Status
from collections import defaultdict
by_status = defaultdict(float)
by_dept = defaultdict(float)

for row in ws.iter_rows(min_row=2, values_only=True):
    if len(row) < 24: continue
    season = str(row[10] or '') if row[10] else ''
    qty = float(row[23] or 0) if row[23] else 0
    dept = str(row[4] or '') if row[4] else ''
    po_status = str(row[17] or '') if len(row) > 17 and row[17] else ''
    
    total_all += qty
    row_count += 1
    
    if season.startswith('26 '):
        total_26 += qty
        row_26 += 1
        by_status[po_status] += qty
        by_dept[dept] += qty

print(f'Total rows: {row_count}')
print(f'26 season rows: {row_26}')
print(f'Total PO (all rows): {total_all:,.0f}')
print(f'Total PO (26 season): {total_26:,.0f}')
print(f'Non-26 season: {total_all - total_26:,.0f}')
print()

print('=== By PO Status (26 season) ===')
for s, q in sorted(by_status.items(), key=lambda x: -x[1]):
    print(f'  {s:20s}: {q:>12,.0f}')
print(f'  {"TOTAL":20s}: {sum(by_status.values()):>12,.0f}')

print()
print('=== By Department (26 season) ===')
for d, q in sorted(by_dept.items(), key=lambda x: -x[1]):
    print(f'  {d:20s}: {q:>12,.0f}')
print(f'  {"TOTAL":20s}: {sum(by_dept.values()):>12,.0f}')

wb.close()
