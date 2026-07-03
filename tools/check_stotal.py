#!/usr/bin/env python3
import openpyxl
BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
f = BASE + r'\리포트_5Sheet_최종_v9_20260630.xlsx'
wb = openpyxl.load_workbook(f, read_only=True)
ws = wb['1.시즌현황']
r = 0
for row in ws.iter_rows(values_only=True):
    r += 1
    if row[0] == 'Total':
        print(f'Found Total at R{r}:')
        for c in range(1, 15):
            v = str(row[c])[:15] if row[c] is not None else '-'
            print(f'  col{c}: {v}')
        break
    if r <= 12:
        print(f'R{r}: {str(row[0])[:12] if row[0] else ""} qty={row[1]} S.TOTAL={row[6] if len(row)>6 else ""}')
wb.close()
