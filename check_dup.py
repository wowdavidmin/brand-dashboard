"""Check if data is duplicated"""
import openpyxl

f = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'
wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
ws = wb.worksheets[0]

# Read first 100 rows and check for duplicates by C8(원가번호) + C11(시즌) + C24(PO수량)
pairs = set()
dup_count = 0
total_qty = 0
row_count = 0

for row in ws.iter_rows(min_row=2, max_row=1000, values_only=True):
    if len(row) < 24: continue
    원가번호 = row[7]  # C8
    시즌 = row[10]     # C11
    po_qty = row[23]   # C24
    po_status = str(row[17] or '') if len(row) > 17 else ''  # C18
    
    if not 시즌 or not str(시즌).startswith('26 '):
        continue
    
    key = (원가번호, 시즌, po_qty, po_status)
    total_qty += float(po_qty or 0)
    row_count += 1
    
    plain_key = (원가번호, 시즌, po_qty)
    if plain_key in pairs:
        dup_count += 1
    else:
        pairs.add(plain_key)

print(f'26 season rows scanned: {row_count}')
print(f'Unique (원가번호+시즌+수량): {len(pairs)}')
print(f'Duplicates found: {dup_count}')
print(f'Total qty: {total_qty:,.0f}')
print(f'Note: C18 PO Status values seen:')

# Check PO Status
wb2 = openpyxl.load_workbook(f, read_only=True, data_only=True)
ws2 = wb2.worksheets[0]
statuses = set()
for row in ws2.iter_rows(min_row=2, max_row=100, values_only=True):
    if len(row) > 17:
        statuses.add(str(row[17] or ''))
print(f'  {statuses}')
wb2.close()
wb.close()
