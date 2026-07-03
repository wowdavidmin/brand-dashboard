"""Simple read - first 10 rows"""
import openpyxl

f = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'
wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
ws = wb.worksheets[0]

for r in range(1, 11):
    vals = {}
    for c in range(1, 55):
        v = ws.cell(r, c).value
        if v is not None:
            vals[c] = v
    
    print(f'Row {r}:')
    for c, v in vals.items():
        print(f'  C{c}: {v}')
    print()

wb.close()
