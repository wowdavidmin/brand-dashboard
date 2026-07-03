#!/usr/bin/env python3
"""Verify CM calculation and rebuild final report"""
import openpyxl, os
from openpyxl.styles import Font, Border, Side, Alignment
from collections import defaultdict

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
ENH = os.path.join(BASE, '영업DATA_Enhanced_v1.0_20260629.xlsx')

# Quick check: CMT column in enhanced file
print('=== Checking CMT column in enhanced file ===')
wb = openpyxl.load_workbook(ENH, read_only=True, data_only=True)
ws = wb['Raw_20260629']

for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        h = list(row)
        for idx in range(40, 50):
            print(f'  Col {idx}: {h[idx]}')
    if i == 1:
        v = list(row)
        qty = v[26]
        cmt = v[44]  # CMT 계 (column index 44 in enhanced = AS column in Excel)
        print(f'\nSample row: qty={qty}, CMT_total={cmt}')
        if qty and cmt:
            print(f'  CM = {cmt} / {qty} = {cmt/qty:.4f}')
        break
wb.close()

# Read file #1 directly to check CMT column
print('\n=== Checking CMT in original file #1 ===')
f1 = os.path.join(BASE, '1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx')
wb1 = openpyxl.load_workbook(f1, read_only=True, data_only=True)
ws1 = wb1['Raw_20260629']
for i, row in enumerate(ws1.iter_rows(values_only=True)):
    if i == 0:
        h = list(row)
        for idx in [38, 39, 40, 41, 42]:
            print(f'  Col {idx}: {h[idx]}')
    if i == 1:
        v = list(row)
        print(f'Sample: CMT_total(col40)={v[40]}, qty(col23)={v[23]}')
        if v[23] and v[40]:
            print(f'  CM = {v[40]} / {v[23]} = {float(v[40])/float(v[23]):.4f}')
        break
wb1.close()

print('\n=== Current vs Corrected CM calculation ===')
print('Before (wrong): CM% = Exec Margin / Net Sales')
print('After (correct): CM = CMT Total / Qty')
