"""Check the raw data file in the specified folder"""
import zipfile, xml.etree.ElementTree as ET

fpath = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'
z = zipfile.ZipFile(fpath)

print('Parsing shared strings...')
ss = ET.parse(z.open('xl/sharedStrings.xml'))
sm = {i: (t.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t').text or '') for i, t in enumerate(ss.getroot()) if t.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') is not None}

print(f'Shared strings: {len(sm)}')

sheet = ET.parse(z.open('xl/worksheets/sheet1.xml'))
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

total = 0
cnt_26 = 0
cnt_all = 0

for row in sheet.getroot().findall('.//s:row', ns):
    r = int(row.get('r'))
    if r == 1: continue
    season = None
    qty = 0
    for c in row.findall('s:c', ns):
        col = ''.join(filter(str.isalpha, c.get('r')))
        v = c.find('s:v', ns)
        if v is not None and v.text:
            if col == 'K':
                season = sm.get(int(v.text), '') if v.text.isdigit() else ''
            elif col == 'X':
                try: qty = float(v.text)
                except: pass
    cnt_all += 1
    total += qty
    if season and season.startswith('26 '):
        cnt_26 += 1

print(f'Total rows: {cnt_all}')
print(f'Total POsu (all): {total:,.0f}')
print(f'26 season: {cnt_26}')
z.close()
