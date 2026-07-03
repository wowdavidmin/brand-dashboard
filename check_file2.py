"""Check June 29 file 2 total"""
import zipfile, xml.etree.ElementTree as ET

fpath = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---b99ec911-9409-4c91-ae86-0a63ab160c8c.xlsx'
z = zipfile.ZipFile(fpath)

print('Parsing shared strings...')
ss = ET.parse(z.open('xl/sharedStrings.xml'))
sm = {}
for i, item in enumerate(ss.getroot()):
    t = item.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
    if t is not None and t.text:
        sm[i] = t.text

print(f'Shared strings: {len(sm)}')

print('Parsing sheet...')
sheet = ET.parse(z.open('xl/worksheets/sheet1.xml'))
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

total = 0
cnt = 0
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
    if season and season.startswith('26 '):
        total += qty
        cnt += 1

print(f'File 2: {cnt} rows, total={total:,.0f}')
z.close()
