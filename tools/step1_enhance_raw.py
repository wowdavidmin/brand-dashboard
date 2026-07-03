#!/usr/bin/env python3
"""Step 1: Enhance file #1 raw data with Factory/Division/Month/NetSales columns"""
import openpyxl, os
from openpyxl.styles import Font, Border, Side, Alignment
from collections import defaultdict
from datetime import datetime

BASE = r'C:\업무자료_약진&JS\업무 인수\영업데이터파일\2026년 6월\4th week'

# Read file #2 to extract mappings
print('Phase 1: Extracting Factory/Division mappings from file #2...')
f2 = os.path.join(BASE, '2) 영업DATA_Raw_v1.3_20260608_openclaw 작업용.xlsx')
wb2 = openpyxl.load_workbook(f2, read_only=True, data_only=True)
ws2 = wb2['Raw_20260608']

factory_map = {}  # Contract Factory -> Factory
# Grouped Div -> Division mapping (from file #2's J->K relationship)
gd_div_map = {}

for i, row in enumerate(ws2.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row)
    cf = str(v[1] or '').strip()
    fac = str(v[2] or '').strip()
    gd = str(v[9] or '').strip()   # J열 = Grouped Div
    div = str(v[10] or '').strip()  # K열 = Division
    
    if cf and fac:
        factory_map[cf.upper()] = fac
        # Also store normalized version
        cf_norm = cf.upper().replace(' ','').replace('.','').replace(',','').replace('(','').replace(')','')
        factory_map[cf_norm] = fac
    
    if gd and div:
        if gd not in gd_div_map:
            gd_div_map[gd] = div
        # If multiple divisions map to same Grouped Div, keep the most common one
        # Already handled by first occurrence

wb2.close()
print(f'  Factory mappings: {len(factory_map)}')
print(f'  Grouped Div->Division mappings: {len(gd_div_map)}')

for gd, div in sorted(gd_div_map.items()):
    print(f'    {gd:25s} -> {div}')

# Build reverse lookup: brand -> Division (fallback)
brand_div_fallback = {}
for gd, div in gd_div_map.items():
    # Extract brand name from Grouped Div
    gd_upper = gd.upper()
    if gd_upper.startswith('GAP ') or gd_upper.startswith('GAP('):
        brand_div_fallback['GAP'] = div  # GAP KIDS/MEN'S/WMN'S -> GAP
        brand_div_fallback['GAP OUTLET'] = div
    elif gd_upper == 'GAP':
        brand_div_fallback['GAP'] = div

# Read file #1 and create enhanced version
print(f'\nPhase 2: Reading & enhancing file #1 raw data...')
f1 = os.path.join(BASE, '1) 브랜드별 직접원가_raw_v1.0_20260629.xlsx')
wb1 = openpyxl.load_workbook(f1, read_only=True, data_only=True)
ws1 = wb1['Raw_20260629']

# Output: enhanced raw data file
OUT_ENHANCED = os.path.join(BASE, '영업DATA_Enhanced_v1.0_20260629.xlsx')
# Use write_only for large file output
wb_out = openpyxl.Workbook(write_only=True)
ws_out = wb_out.create_sheet('Raw_20260629')

# Headers (58 columns, matching file #2 format)
headers = [
    'NO', 'Contract Factory', 'Factory',            # A, B, C
    '바이어', '브랜드', '부서',                        # D, E, F
    '바이어스타일', '스타일', '원가번호',               # G, H, I
    'Grouped Div', 'Division', '디비전',              # J, K, L
    '시즌', 'ITEM명', 'ITEM명.1', 'ITEM명.2',         # M, N, O, P
    '담당', '진행구분', '오더', 'PO Status',           # Q, R, S, T
    '선적일', '월', 'ex-Fact', 'DC',                 # U, V, W, X
    'SHIP MODE', 'TERMS',                            # Y, Z
    'PO수량', 'PO금액', '매출할인', '순매출',           # AA, AB, AC, AD
    '원자재비(원사)', '원자재비(원단)', '부자재비(본사)',  # AE, AF, AG
    '염가공', 'I.F염가공', '염가공비 계',              # AH, AI, AJ
    '임가공료(외화)', '임가공료(원화)', '임가공료(합계)', # AK, AL, AM
    'CMT(봉제료)', 'CMT(I.F)', 'CMT(부자재)',          # AN, AO, AP
    'CMT(임가공료해외)', 'CMT(기타수수료)', 'CMT 계',   # AQ, AR, AS
    '해상운임', 'Inspection Fee', 'NCPM',             # AT, AU, AV
    '직접원가(AMOUNT)', '직접원가(%)',                  # AW, AX
    'PROFIT(AMOUNT)', 'PROFIT(%)',                    # AY, AZ
    '손익', 'Buyer Commission', 'Agent Commission',   # BA, BB, BC
    'Duty', 'Execute Margin', 'Execute Margin(%)'      # BD, BE, BF
]
ws_out.append(headers)

months_kr = {1:'1월',2:'2월',3:'3월',4:'4월',5:'5월',6:'6월',
             7:'7월',8:'8월',9:'9월',10:'10월',11:'11월',12:'12월'}

cnt = 0
unmapped_cf = set()
unmapped_gd = {}

for i, row in enumerate(ws1.iter_rows(values_only=True)):
    if i == 0: continue
    v = list(row); cnt += 1
    
    # Original columns from file #1 (0-indexed)
    no = v[0]
    contract_factory = str(v[1] or '').strip()
    buyer = v[2]
    brand = v[3]
    dept = v[4]
    bstyle = v[5]
    style = v[6]
    cost_no = v[7]
    grouped_div = str(v[8] or '').strip()  # I열 in file #1, J열 in file #2
    division_old = v[9]
    season = v[10]
    item = v[11]; item1 = v[12]; item2 = v[13]
    person = v[14]; progress = v[15]
    order = v[16]; po_status = v[17]
    ship_date = v[18]; ex_fact = v[19]; dc_date = v[20]
    ship_mode = v[21]; terms = v[22]
    po_qty = v[23]; po_amt = v[24]; discount = v[25]
    mat_yarn = v[26]; mat_fabric = v[27]; mat_accessory = v[28]
    dyeing = v[29]; if_dyeing = v[30]; dyeing_total = v[31]
    proc_fx = v[32]; proc_krw = v[33]; proc_total = v[34]
    cmt_sewing = v[35]; cmt_if = v[36]; cmt_acc = v[37]
    cmt_proc = v[38]; cmt_fee = v[39]; cmt_total = v[40]
    freight = v[41]; inspection = v[42]; ncpm = v[43]
    dc_cost = v[44]; dc_pct = v[45]
    profit = v[46]; profit_pct = v[47]
    pnl = v[48]; buyer_comm = v[49]; agent_comm = v[50]
    duty = v[51]; exec_margin = v[52]; exec_margin_pct = v[53]
    
    # ---- ENHANCED COLUMNS ----
    
    # 1. Factory (C열): map from Contract Factory
    cf_key = contract_factory.upper()
    factory = factory_map.get(cf_key)
    if not factory:
        cf_norm = cf_key.replace(' ','').replace('.','').replace(',','').replace('(','').replace(')','')
        factory = factory_map.get(cf_norm)
    if not factory:
        # Partial match
        for fk, fv in factory_map.items():
            if cf_norm in fk or fk in cf_norm:
                factory = fv
                break
    if not factory:
        factory = contract_factory
        unmapped_cf.add(contract_factory)
    
    # 2. Division (K열): derive from Grouped Div (J열)
    gd_upper = grouped_div.upper()
    if 'IPSS GAP' in gd_upper:
        division = 'IPSS GAP'
    elif 'IPSS OLD NAVY' in gd_upper:
        division = 'IPSS ON'
    elif 'GAP' in gd_upper:
        division = 'GAP'
    else:
        # Use direct mapping from file #2, or fallback to Grouped Div itself
        division = gd_div_map.get(grouped_div, grouped_div)
    
    # 3. Month (V열): from ship date
    month_str = ''
    if isinstance(ship_date, datetime):
        month_str = months_kr.get(ship_date.month, '')
    elif isinstance(ship_date, str) and len(ship_date) >= 10:
        try:
            d = datetime.strptime(ship_date[:10], '%Y-%m-%d')
            month_str = months_kr.get(d.month, '')
        except: pass
    
    # 4. Net Sales (AD열): PO Amt - Discount
    net_sales = (po_amt or 0) - (discount or 0)
    
    # Also add 디비전 (L열) - use Division for this too
    div_l = division  # same as division
    
    # Ship mode - handle None for TERMS
    ship_mode_val = ship_mode if ship_mode else ''
    
    # Build the enhanced row (58 columns)
    enhanced = [
        no, contract_factory, factory,           # A, B, C
        buyer, brand, dept,                      # D, E, F
        bstyle, style, cost_no,                  # G, H, I
        grouped_div, division, div_l,            # J, K, L
        season, item, item1, item2,              # M, N, O, P
        person, progress, order, po_status,      # Q, R, S, T
        ship_date, month_str, ex_fact, dc_date,  # U, V, W, X
        ship_mode_val, terms,                    # Y, Z
        po_qty, po_amt, discount, net_sales,     # AA, AB, AC, AD
        mat_yarn, mat_fabric, mat_accessory,     # AE, AF, AG
        dyeing, if_dyeing, dyeing_total,         # AH, AI, AJ
        proc_fx, proc_krw, proc_total,           # AK, AL, AM
        cmt_sewing, cmt_if, cmt_acc,             # AN, AO, AP
        cmt_proc, cmt_fee, cmt_total,            # AQ, AR, AS
        freight, inspection, ncpm,               # AT, AU, AV
        dc_cost, dc_pct,                         # AW, AX
        profit, profit_pct,                      # AY, AZ
        pnl, buyer_comm, agent_comm,             # BA, BB, BC
        duty, exec_margin, exec_margin_pct       # BD, BE, BF
    ]
    ws_out.append(enhanced)
    
    if cnt % 20000 == 0: print(f'  {cnt}...')

wb1.close()
wb_out.save(OUT_ENHANCED)
print(f'Enhanced raw data saved: {OUT_ENHANCED}')
print(f'  Total rows: {cnt}')
print(f'  Unmapped factories: {len(unmapped_cf)}')
for cf in sorted(unmapped_cf):
    print(f'    {cf}')
