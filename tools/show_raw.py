#!/usr/bin/env python3
import openpyxl
BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'
f2 = BASE + r'\2) 영업DATA_Raw_v1.3_20260608_openclaw 작업용.xlsx'
wb = openpyxl.load_workbook(f2, read_only=True, data_only=True)
ws = wb['Raw_20260608']

print('=== Raw Data Preview (first 10 rows, key columns) ===')
print()

cnt = 0
for i, row in enumerate(ws.iter_rows(values_only=True)):
    if i == 0:
        # Headers
        h = list(row)
        print(f'Col B(1)=Contract Factory: {h[1]}')
        print(f'Col C(2)=Factory: {h[2]}')
        print(f'Col D(3)=Buyer: {h[3]}')
        print(f'Col E(4)=Brand: {h[4]}')
        print(f'Col F(5)=Dept: {h[5]}')
        print(f'Col K(10)=Division: {h[10]}')
        print(f'Col M(12)=Season: {h[12]}')
        print(f'Col V(21)=Month: {h[21]}')
        print(f'Col Z(26)=PO Qty: {h[26]}')
        print(f'Col AA(27)=PO Amt: {h[27]}')
        print(f'Col AD(29)=Net Sales: {h[29]}')
        print()
        continue
    v = list(row); cnt += 1
    factory = str(v[2] or '')[:20]
    brand = str(v[4] or '')[:18]
    buyer = str(v[3] or '')[:18]
    division = str(v[10] or '')[:18]
    season = str(v[12] or '')[:8]
    month = str(v[21] or '')[:4]
    qty = int(v[26] or 0)
    amt = v[27] or 0
    net = v[29] or 0
    
    print(f'Factory={factory:20s} | Brand={brand:18s} | Div={division:18s} | Season={season:8s} | Mon={month:4s} | Qty={qty:>8,} | Amt={amt:>10,.2f} | Net={net:>10,.2f}')
    if cnt >= 8: break

print(f'\nTotal rows in file: 81,616')
wb.close()
