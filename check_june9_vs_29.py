"""Compare June 9 vs June 29 total order quantities"""
import openpyxl

# June 29 source: existing v12 (already processed)
v12_path = r'C:\Users\user\.openclaw\workspace\영업데이터파일_v12_20260702_processed.xlsx'
wb = openpyxl.load_workbook(v12_path)
ws = wb.worksheets[0]

print('=== June 29 Total (from v12 data) ===')
total_29 = 0
stotal_cols = [8, 11, 14, 17, 20, 23, 25, 30, 32, 34]
for c in stotal_cols:
    v = ws.cell(10, c).value
    print(f'  C{c} ({ws.cell(5, c).value}): {v:>10,.0f}')
    total_29 += float(v if v not in ('-', None, '') else 0)
print(f'  {"합계":>30s}: {total_29:>10,.0f}')
wb.close()

# June 9: check raw data file
print('\n=== June 9 Raw Data (first 5 rows) ===')
june9_path = r'C:\Users\user\Downloads\브랜드별 직접원가_20260609.xlsx'
wb9 = openpyxl.load_workbook(june9_path, read_only=True, data_only=True)
ws9 = wb9.worksheets[0]

# Get column indices for needed data
print(f'Rows: {ws9.max_row}, Cols: {ws9.max_column}')
headers = {}
for c in range(1, ws9.max_column+1):
    h = ws9.cell(1, c).value
    if h:
        headers[str(h).strip()] = c
        print(f'  C{c}: {h}')

wb9.close()
