"""Check PO Status and 진행구분 distribution"""
import zipfile, xml.etree.ElementTree as ET
from collections import Counter

fpath = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---47ea1d42-8317-4a51-a9fa-d5c2e2764f59.xlsx'
z = zipfile.ZipFile(fpath)

ss = ET.parse(z.open('xl/sharedStrings.xml'))
sm = {i: (t.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t').text or '') for i, t in enumerate(ss.getroot()) if t.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t') is not None}

sheet = ET.parse(z.open('xl/worksheets/sheet1.xml'))
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}

status_count = Counter()
po_status_count = Counter()
total_by_status = {}
total_by_po = {}

for row in sheet.getroot().findall('.//s:row', ns):
    r = int(row.get('r'))
    if r == 1: continue
    vals = {}
    for c in row.findall('s:c', ns):
        col = ''.join(filter(str.isalpha, c.get('r')))
        v = c.find('s:v', ns)
        if v is not None and v.text:
            try:
                vals[col] = int(v.text)
            except ValueError:
                try:
                    vals[col] = float(v.text)
                except:
                    vals[col] = v.text
    
    season = sm.get(vals.get('K', -1), '') if isinstance(vals.get('K', ''), int) else ''
    if not season.startswith('26 '): continue
    
    qty = vals.get('X', 0) if isinstance(vals.get('X', 0), (int, float)) else 0
    
    # 진행구분 (C16 = col P)
    status = sm.get(vals.get('P', -1), 'N/A') if isinstance(vals.get('P', ''), int) else 'N/A'
    status_count[status] += 1
    total_by_status[status] = total_by_status.get(status, 0) + qty
    
    # PO Status (C18 = col R)
    po_stat = sm.get(vals.get('R', -1), 'N/A') if isinstance(vals.get('R', ''), int) else 'N/A'
    po_status_count[po_stat] += 1
    total_by_po[po_stat] = total_by_po.get(po_stat, 0) + qty

print('=== 진행구분 (C16) ===')
for s, c in status_count.most_common():
    print(f'  {s:20s} count={c:>6} total_qty={total_by_status[s]:>12,.0f}')

print('\n=== PO Status (C18) ===')
for s, c in po_status_count.most_common():
    print(f'  {s:20s} count={c:>6} total_qty={total_by_po[s]:>12,.0f}')

z.close()
