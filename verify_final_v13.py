"""Verify v13 matches raw data"""
import openpyxl

v13 = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\영업데이터파일_v13_20260702.xlsx'
wb = openpyxl.load_workbook(v13)
ws = wb.worksheets[0]

print('=== v13 ORDER 수량 총합계 ===')
stotal_cols = [8, 11, 14, 17, 20, 23, 25, 30, 32, 34, 36]

for season, row in [('26 SPR',6), ('26 SUM',7), ('26 FAL',8), ('26 HOL',9), ('Total',10)]:
    total = 0
    for c in stotal_cols:
        v = ws.cell(row, c).value
        total += float(v if v not in ('-', None, '') else 0)
    print(f'  {season:8s}: 합계(C35)={ws.cell(row,35).value:>12,.0f} (S.TOTAL={total:>12,.0f})')

print(f'\n=== BU별 합계 (Total row) ===')
for col, name in [(8,'BU1'),(11,'BU2-1'),(14,'BU2-2'),(17,'BU3-1'),
                  (20,'BU3-2'),(23,'BU3-3'),(25,'BU3-4'),
                  (30,'BU4'),(32,'BU5'),(34,'GTBU'),(36,'인터텍스')]:
    v = ws.cell(10, col).value
    print(f'  {name:10s}: {v:>12,.0f}')

print(f'\n=== KOHL_S, URBAN 상세 ===')
for col, name in [(18, "KOHL_S"), (19, 'URBAN'), (20, 'S.TOTAL(BU3-2)')]:
    print(f'{name}:')
    for r in [6,7,8,9,10]:
        print(f'  {ws.cell(r,2).value}: {ws.cell(r,col).value}')
wb.close()
