#!/usr/bin/env python3
"""Final formatting: 0->dash, commas for numbers, % for 비중"""
import openpyxl, os, shutil

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
src = os.path.join(BASE, '리포트_5Sheet_최종_v8_20260630.xlsx')
out = os.path.join(BASE, '리포트_5Sheet_최종_FINAL_20260630.xlsx')
shutil.copy2(src, out)

wb = openpyxl.load_workbook(out)

for ws in wb.worksheets:
    name = ws.title
    max_row = ws.max_row
    max_col = ws.max_column
    
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            val = cell.value
            
            if val is None or isinstance(val, str):
                continue
            if not isinstance(val, (int, float)):
                continue
            
            # Sheet 5: 비중 columns (cols 4,6,8,10,12 = even cols from 4)
            if name == '5.시즌공장' and c >= 4 and c % 2 == 0:
                if val == 0:
                    cell.value = '-'
                else:
                    cell.value = round(val * 100, 1)
                    cell.number_format = '0.0'
                continue
            
            # Sheet 4: 비중% columns (cols 5,7,9,11)
            if name == '4.시즌부서공장' and c >= 5 and (c % 2 == 1):
                if val == 0:
                    cell.value = '-'
                else:
                    cell.number_format = '0.0'
                continue
            
            # All other numeric cells: 0->dash, commas for positive
            if val == 0:
                cell.value = '-'
            elif isinstance(val, float) and val != int(val):
                cell.number_format = '#,##0.00'
            else:
                cell.number_format = '#,##0'

wb.save(out)
import os; print(f'Saved: {out} ({os.path.getsize(out)/1024:.0f} KB)')
print('Sheets:', wb.sheetnames)
wb.close()
