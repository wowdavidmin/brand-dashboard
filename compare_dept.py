"""Compare department totals between June 9 and June 29"""
import zipfile, xml.etree.ElementTree as ET

def get_dept_totals(filepath):
    z = zipfile.ZipFile(filepath)
    ss = ET.parse(z.open('xl/sharedStrings.xml'))
    ss_map = {}
    for i, item in enumerate(ss.getroot()):
        t = item.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
        if t is not None and t.text:
            ss_map[i] = t.text
    
    ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    sheet = ET.parse(z.open('xl/worksheets/sheet1.xml'))
    
    dept_totals = {}
    for row in sheet.getroot().findall('.//s:row', ns):
        r = int(row.get('r'))
        if r == 1: continue
        season = None; dept = None; qty = 0
        for c in row.findall('s:c', ns):
            col = ''.join(filter(str.isalpha, c.get('r')))
            v = c.find('s:v', ns)
            if v is not None and v.text:
                if col == 'E':
                    dept = ss_map.get(int(v.text), '')
                elif col == 'K':
                    season = ss_map.get(int(v.text), '')
                elif col == 'X':
                    try: qty = float(v.text)
                    except: pass
        if season and season.startswith('26 ') and dept:
            dept_totals[dept] = dept_totals.get(dept, 0) + qty
    z.close()
    return dept_totals

f9 = r'C:\Users\user\Downloads\브랜드별 직접원가_20260609.xlsx'
f29 = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---47ea1d42-8317-4a51-a9fa-d5c2e2764f59.xlsx'

print('Jun 9...')
d9 = get_dept_totals(f9)
print('Jun 29...')
d29 = get_dept_totals(f29)

print(f'\n{"Dept":20s} {"Jun9":>12s} {"Jun29":>12s} {"Diff":>12s}')
print('-' * 60)
for d in sorted(set(list(d9.keys()) + list(d29.keys()))):
    v9 = int(d9.get(d, 0))
    v29 = int(d29.get(d, 0))
    print(f'{d:20s} {v9:>12,} {v29:>12,} {v29-v9:>+12,}')

t9 = int(sum(d9.values()))
t29 = int(sum(d29.values()))
print(f'{"TOTAL":20s} {t9:>12,} {t29:>12,} {t29-t9:>+12,}')
