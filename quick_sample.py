"""Quick sample to understand data structure"""
import zipfile, xml.etree.ElementTree as ET

f = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'
z = zipfile.ZipFile(f)

ss = ET.parse(z.open('xl/sharedStrings.xml'))
sm = {i: (t.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t').text or '') for i, t in enumerate(ss.getroot()) if t.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') is not None}

sheet = ET.parse(z.open('xl/worksheets/sheet1.xml'))
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

from collections import Counter
po_statuses = Counter()
statuses = Counter()
count_26 = 0

for row in list(sheet.getroot().findall('.//s:row', ns))[1:501]:  # rows 2-500
    r = int(row.get('r'))
    vals = {}
    for c in row.findall('s:c', ns):
        col = ''.join(filter(str.isalpha, c.get('r')))
        v = c.find('s:v', ns)
        if v is not None and v.text:
            if v.text.isdigit():
                try: vals[col] = int(v.text)
                except: vals[col] = v.text
            else:
                try: vals[col] = sm.get(int(v.text), v.text)
                except: vals[col] = v.text
    
    season = vals.get('K', '') if isinstance(vals.get('K', ''), str) else ''
    if season.startswith('26 '):
        count_26 += 1
        pops = str(vals.get('R', 'N/A'))
        st = str(vals.get('P', 'N/A'))
        po_statuses[pops] += 1
        statuses[st] += 1

print(f'Sample rows (500): 26 season count={count_26}')
print(f'PO Status (R): {dict(po_statuses)}')
print(f'Status (P): {dict(statuses)}')
z.close()
