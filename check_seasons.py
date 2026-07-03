"""Check all season values in June 29 file"""
import zipfile, xml.etree.ElementTree as ET

fpath = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---47ea1d42-8317-4a51-a9fa-d5c2e2764f59.xlsx'
z = zipfile.ZipFile(fpath)

ss = ET.parse(z.open('xl/sharedStrings.xml'))
sm = {}
for i, item in enumerate(ss.getroot()):
    t = item.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
    if t is not None and t.text:
        sm[i] = t.text

sheet = ET.parse(z.open('xl/worksheets/sheet1.xml'))
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

from collections import Counter
season_counts = Counter()
total_all = 0
total_26 = 0
count_26 = 0

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
                season = sm.get(int(v.text), '')
            elif col == 'X':
                try: qty = float(v.text)
                except: pass
    if season:
        season_counts[season] += 1
    total_all += qty
    if season and season.startswith('26 '):
        total_26 += qty
        count_26 += 1

print(f'All rows total: {total_all:,.0f}')
print(f'26 season total: {total_26:,.0f}')
print(f'26 season count: {count_26}')
print()
print('Season distribution (top 20):')
for s, c in season_counts.most_common(20):
    print(f'  {s}: {c}')
print(f'  ... total seasons: {len(season_counts)}')
z.close()
