"""Check if data is double-counted by PO Status"""
import zipfile, xml.etree.ElementTree as ET
from collections import Counter

fpath = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week\1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx'
z = zipfile.ZipFile(fpath)

ss = ET.parse(z.open('xl/sharedStrings.xml'))
sm = {i: (t.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t').text or '') for i, t in enumerate(ss.getroot()) if t.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') is not None}

sheet = ET.parse(z.open('xl/worksheets/sheet1.xml'))
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

# Check 진행구분 (C16=P열) and PO Status (C18=R열) distribution
from collections import defaultdict
status_data = defaultdict(lambda: {'count': 0, 'qty': 0.0})

for row in sheet.getroot().findall('.//s:row', ns):
    r = int(row.get('r'))
    if r == 1: continue
    vals = {}
    for c in row.findall('s:c', ns):
        col = ''.join(filter(str.isalpha, c.get('r')))
        v = c.find('s:v', ns)
        if v is not None and v.text:
            if col in ('K', 'P', 'R'):
                vals[col] = sm.get(int(v.text), '') if v.text.isdigit() else v.text
            elif col == 'X':
                try: vals[col] = float(v.text)
                except: pass
    
    season = vals.get('K', '')
    if not str(season).startswith('26 '): continue
    
    po_status = str(vals.get('R', 'N/A'))
    qty = vals.get('X', 0)
    
    status_data[po_status]['count'] += 1
    status_data[po_status]['qty'] += qty

print('=== By PO Status (C18/R열) ===')
total_qty = 0
for s, d in sorted(status_data.items(), key=lambda x: -x[1]['qty']):
    print(f'  {s:20s} count={d["count"]:>6} qty={d["qty"]:>12,.0f}')
    total_qty += d['qty']
print(f'  {"TOTAL":20s} qty={total_qty:>12,.0f}')

# Maybe use 진행구분 instead?
print()

# Check 진행구분
status_data2 = defaultdict(lambda: {'count': 0, 'qty': 0.0})
z2 = zipfile.ZipFile(fpath)
sheet2 = ET.parse(z2.open('xl/worksheets/sheet1.xml'))
for row in sheet2.getroot().findall('.//s:row', ns):
    r = int(row.get('r'))
    if r == 1: continue
    vals = {}
    for c in row.findall('s:c', ns):
        col = ''.join(filter(str.isalpha, c.get('r')))
        v = c.find('s:v', ns)
        if v is not None and v.text:
            if col in ('K', 'P', 'R'):
                vals[col] = sm.get(int(v.text), '') if v.text.isdigit() else v.text
            elif col == 'X':
                try: vals[col] = float(v.text)
                except: pass
    
    season = vals.get('K', '')
    if not str(season).startswith('26 '): continue
    
    p = str(vals.get('P', 'N/A'))
    qty = vals.get('X', 0)
    status_data2[p]['count'] += 1
    status_data2[p]['qty'] += qty

print('=== By 진행구분 (C16/P열) ===')
for s, d in sorted(status_data2.items(), key=lambda x: -x[1]['qty']):
    print(f'  {s:20s} count={d["count"]:>6} qty={d["qty"]:>12,.0f}')
z2.close()
z.close()
