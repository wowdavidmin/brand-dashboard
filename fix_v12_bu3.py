"""Fix BU3 data in v12 file using raw data pivot"""

import openpyxl
from collections import defaultdict

# ========== STEP 1: Aggregate raw data ==========
raw_path = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---47ea1d42-8317-4a51-a9fa-d5c2e2764f59.xlsx'

wb_raw = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)
ws_raw = wb_raw.worksheets[0]

# Only process BU3 departments
bu3_depts = {'BU3_D1', 'BU3_D2', 'BU3_D3', 'BU3_D4'}
data = defaultdict(lambda: defaultdict(lambda: {'qty': 0, 'amt': 0, 'disc': 0, 'cmt': 0}))

count = 0
for row in ws_raw.iter_rows(min_row=2, values_only=True):
    # C5(부서)=col4, C11(시즌)=col10, C24(PO수량)=col23, C25(PO금액)=col24, C26(매출할인)=col25, C41(CMT계)=col40
    dept = str(row[4] or '').strip() if len(row) > 4 else ''
    season = str(row[10] or '').strip() if len(row) > 10 else ''
    
    if dept not in bu3_depts:
        continue
    if not season.startswith('26 '):
        continue
    
    qty = float(row[23] or 0) if len(row) > 23 and row[23] is not None else 0
    amt = float(row[24] or 0) if len(row) > 24 and row[24] is not None else 0
    disc = float(row[25] or 0) if len(row) > 25 and row[25] is not None else 0
    cmt = float(row[40] or 0) if len(row) > 40 and row[40] is not None else 0
    
    data[dept][season]['qty'] += qty
    data[dept][season]['amt'] += amt
    data[dept][season]['disc'] += disc
    data[dept][season]['cmt'] += cmt
    count += 1

wb_raw.close()
print(f'Processed {count} raw data rows for BU3')

# Print BU3 aggregates
print('\n=== BU3 Aggregated from Raw Data ===')
for dept in ['BU3_D1', 'BU3_D2', 'BU3_D3', 'BU3_D4']:
    print(f'\n--- {dept} ---')
    for s in ['26 SPR', '26 SUM', '26 FAL', '26 HOL']:
        d = data[dept].get(s, {'qty': 0, 'amt': 0, 'disc': 0, 'cmt': 0})
        net = d['amt'] - d['disc']
        avg = d['amt'] / d['qty'] if d['qty'] > 0 else 0
        cm_u = d['cmt'] / d['qty'] if d['qty'] > 0 else 0
        print(f'  {s:8s} qty={d["qty"]:>10.0f} amt={d["amt"]/1000:>8.1f}K disc={d["disc"]/1000:>8.1f}K net={net/1000:>8.1f}K avg_price={avg:.2f} cm={cm_u:.2f}')
    
    total_qty = sum(data[dept][s]['qty'] for s in ['26 SPR', '26 SUM', '26 FAL', '26 HOL'] if s in data[dept])
    total_amt = sum(data[dept][s]['amt'] for s in ['26 SPR', '26 SUM', '26 FAL', '26 HOL'] if s in data[dept])
    print(f'  {"TOTAL":8s} qty={total_qty:>10.0f} amt={total_amt/1000:>8.1f}K')

# ========== STEP 2: Compare with v12 ==========
print('\n\n=== Comparison: v12 vs Raw (BU3 brand columns) ===')
v12_path = r'C:\Users\user\.openclaw\media\inbound\영업데이터파일_v12_20260702---c5b345bd-47c5-47cf-b54c-e337523c2574.xlsx'
wb_v12 = openpyxl.load_workbook(v12_path, data_only=True)
ws = wb_v12.worksheets[0]

# v12 column → what it should represent
# C15=BR → BANANA REPUBLIC (BU3_D1)
# C16=IPSS ON → ON IPSS (BU3_D1)  
# C18=KOHL'S → KOHL`S (BU3_D2)
# C19=URBAN OUTFITTERS → URBAN OUTFITTERS (BU3_D2) - but v12 says "URBAN"
# C21=AERIE → AERIE (BU3_D3)
# C22=Offline → Offline (BU3_D3)
# C24=MLB → MLB (BU3_D4)

# Check brand-level breakdown within BU3
brand_agg = defaultdict(lambda: defaultdict(lambda: {'qty': 0, 'amt': 0}))
wb_raw2 = openpyxl.load_workbook(raw_path, read_only=True, data_only=True)
ws2 = wb_raw2.worksheets[0]
for row in ws2.iter_rows(min_row=2, values_only=True):
    dept = str(row[4] or '').strip() if len(row) > 4 else ''
    season = str(row[10] or '').strip() if len(row) > 10 else ''
    buyer = str(row[2] or '').strip() if len(row) > 2 else ''
    
    if dept not in bu3_depts:
        continue
    if not season.startswith('26 '):
        continue
    
    qty = float(row[23] or 0) if len(row) > 23 and row[23] is not None else 0
    amt = float(row[24] or 0) if len(row) > 24 and row[24] is not None else 0
    
    brand_agg[(dept, buyer)][season]['qty'] += qty
    brand_agg[(dept, buyer)][season]['amt'] += amt

wb_raw2.close()

# Print brand-level for BU3
for dept in ['BU3_D1', 'BU3_D2', 'BU3_D3', 'BU3_D4']:
    print(f'\n--- {dept} Brand Breakdown ---')
    for (d, buyer), seasons in sorted(brand_agg.items()):
        if d != dept: continue
        total_q = sum(s['qty'] for s in seasons.values())
        total_a = sum(s['amt'] for s in seasons.values())
        print(f'{buyer:30s} total_qty={total_q:>10.0f} total_amt={total_a/1000:>10.1f}K')
        for s in ['26 SPR', '26 SUM', '26 FAL', '26 HOL']:
            if s in seasons:
                print(f'  {s:8s} qty={seasons[s]["qty"]:>10.0f} amt={seasons[s]["amt"]/1000:>8.1f}K')
