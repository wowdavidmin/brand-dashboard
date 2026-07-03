"""Compare total order quantities from June 9 vs June 29"""
import openpyxl
from collections import defaultdict

def get_totals_by_dept(filepath):
    """Get PO수량 totals by department for 26 season"""
    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    
    dept_totals = defaultdict(float)
    grand_total = 0
    count = 0
    
    for row in ws.iter_rows(min_row=2, values_only=True):
        # C5=부서, C11=시즌, C24=PO수량
        if len(row) < 24: continue
        season = str(row[10] or '').strip() if row[10] else ''
        if not season.startswith('26 '): continue
        
        dept = str(row[4] or '').strip() if row[4] else ''
        qty = float(row[23] or 0) if row[23] else 0
        
        dept_totals[dept] += qty
        grand_total += qty
        count += 1
    
    wb.close()
    return dept_totals, grand_total, count

print('Reading June 9 file...')
dept9, total9, cnt9 = get_totals_by_dept(r'C:\Users\user\Downloads\브랜드별 직접원가_20260609.xlsx')
print(f'  {cnt9} rows processed')
print(f'  Grand total: {total9:,.0f}')

print('\nReading June 29 file...')
dept29, total29, cnt29 = get_totals_by_dept(r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---47ea1d42-8317-4a51-a9fa-d5c2e2764f59.xlsx')
print(f'  {cnt29} rows processed')
print(f'  Grand total: {total29:,.0f}')

print(f'\n=== Comparison ===')
print(f'June 9 total:  {total9:>12,.0f}')
print(f'June 29 total: {total29:>12,.0f}')
print(f'Difference:    {total29 - total9:>+12,.0f}')

print(f'\n=== By Department ===')
all_depts = sorted(set(list(dept9.keys()) + list(dept29.keys())))
for d in all_depts:
    v9 = dept9.get(d, 0)
    v29 = dept29.get(d, 0)
    diff = v29 - v9
    print(f'  {d:20s}: Jun9={v9:>10,.0f}  Jun29={v29:>10,.0f}  diff={diff:>+10,.0f}')

print(f'\n=== Note ===')
if total29 < total9:
    print(f'June 29 ({total29:,.0f}) < June 9 ({total9:,.0f}) by {total9 - total29:,.0f}')
    print('This is unusual - data should only increase over time.')
elif total29 > total9:
    print(f'June 29 ({total29:,.0f}) > June 9 ({total9:,.0f}) by {total29 - total9:,.0f}')
    print('Normal - data increased.')
else:
    print('Totals are equal.')
