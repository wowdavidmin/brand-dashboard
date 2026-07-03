"""
Raw data → Template format (1~4 sheets) exactly matching structure
"""
import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')

SRC = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---b99ec911-9409-4c91-ae86-0a63ab160c8c.xlsx'
OUT = r'C:\Users\user\.openclaw\workspace\브랜드별_직접원가_피벗분석_v1.0_20260629.xlsx'

print("Reading raw data...")
df = pd.read_excel(SRC, sheet_name='Raw_20260629')
print(f"Shape: {df.shape}")

# Clean columns
df.columns = [c.replace('\n', ' ').strip() for c in df.columns]
print(f"Columns cleaned: {list(df.columns)}")

# Convert numeric
numerics = ['PO수량', 'PO금액', '매출할인', '원자재비 원사', '원자재비 원단', '부자재비 본사',
            '염가공', 'I.F염가공', '염가공비 계', '임가공료 외화', '임가공료 원화', '임가공료 합계',
            'CMT 봉제료', 'CMT I.F', 'CMT 부자재', 'CMT 임가공료해외', 'CMT 기타수수료', 'CMT 계',
            '해상운임', 'Inspection Fee', 'NCPM', '직접원가 AMOUNT', '직접원가 %',
            'PROFIT AMOUNT', 'PROFIT %', 'Buyer Commission', 'Agent Commission', 'Duty',
            'Execute Margin', 'Execute Margin(%)']
for c in numerics:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

# Clean text cols
for c in ['브랜드', '부서', '시즌', '바이어', 'Contract Factory', 'Grouped Div.', '디비젼']:
    if c in df.columns:
        df[c] = df[c].apply(lambda x: str(x).strip() if pd.notna(x) else '')

# Filter out empty
for c in ['브랜드', '부서', '시즌']:
    df = df[df[c] != ''].copy()

print(f"Unique brands: {df['브랜드'].nunique()}")
print(f"Unique depts: {df['부서'].nunique()}")
print(f"Unique seasons: {df['시즌'].nunique()}")
brands = sorted(df['브랜드'].tolist())
depts = sorted(df['부서'].tolist())
seasons = sorted(df['시즌'].tolist())
print(f"Brands: {brands}")
print(f"Seasons: {seasons}")
print(f"Depts: {depts}")

# ===============================================
# SHEET 1: 시즌별 부서/브랜드 직접원가 비교 (Template: Guide margin)
# Hierarchical: 부서 > 브랜드 with multiple season columns
# ===============================================
print("\n=== Building Sheet 1 ===")
# Dept mapping (simplified from BU codes)
def dept_label(d):
    labels = {
        'BU1_D1': '영업1부(Sales1)', 'BU1_D2': '영업1부(Sales1)',
        'BU2_D1': '영업2부(Sales2)', 'BU2_D2': '영업2부(Sales2)',
        'BU3_D1': '영업3부(Sales3)', 'BU3_D2': '영업3부(Sales3)',
        'BU3_D3': '영업3부(Sales3)', 'BU3_D4': '영업3부(Sales3)',
        'BU4': '영업4부(Sales4)', 'BU5': '영업5부(Sales5)',
        'GTBU': '영업6부(GTBU)',
        '인터텍스 기획부': '인터텍스'
    }
    return labels.get(d, d)

df['부서라벨'] = df['부서'].apply(dept_label)

# Aggregate by season + dept + brand
agg1 = df.groupby(['부서라벨', '브랜드', '시즌']).agg({
    'PO금액': 'sum',
    '직접원가 AMOUNT': 'sum',
    'PROFIT AMOUNT': 'sum',
}).reset_index()

m1 = agg1['PO금액'] > 0
agg1['마진율'] = np.where(m1, (agg1['PROFIT AMOUNT'] / agg1['PO금액'] * 100).round(2), 0)
agg1['직접원가율'] = np.where(m1, (agg1['직접원가 AMOUNT'] / agg1['PO금액'] * 100).round(2), 0)

# ===============================================
# SHEET 2: 브랜드별 직접원가 구성 분석 (Template: 0. Guide margin)
# ===============================================
print("\n=== Building Sheet 2 ===")
# By brand
grp2 = df.groupby(['브랜드']).agg({
    'PO금액': 'sum',
    '원자재비 원사': 'sum',
    '원자재비 원단': 'sum',
    '부자재비 본사': 'sum',
    '염가공비 계': 'sum',
    '임가공료 합계': 'sum',
    'CMT 계': 'sum',
    '해상운임': 'sum',
    '직접원가 AMOUNT': 'sum',
    'PROFIT AMOUNT': 'sum',
    'PO수량': 'sum',
}).reset_index()

m2 = grp2['PO금액'] > 0
grp2['직접원가율(%)'] = np.where(m2, (grp2['직접원가 AMOUNT'] / grp2['PO금액'] * 100).round(2), 0)
grp2['마진율(%)'] = np.where(m2, (grp2['PROFIT AMOUNT'] / grp2['PO금액'] * 100).round(2), 0)
grp2['PO단가'] = np.where(grp2['PO수량'] > 0, (grp2['PO금액'] / grp2['PO수량']).round(2), 0)

# Cost composition
cost_total = grp2['직접원가 AMOUNT'].replace(0, np.nan)
for col in ['원자재비 원사', '원자재비 원단', '부자재비 본사', '염가공비 계', '임가공료 합계', 'CMT 계', '해상운임']:
    grp2[f'{col}_%'] = (grp2[col] / cost_total * 100).fillna(0).round(2)

# ===============================================
# SHEET 3: 시즌별 브랜드 PO금액/직접원가 Pivot (Template: YAKJIN()
# 시즌(row) × 브랜드(col) with PO금액 and 직접원가
# ===============================================
print("\n=== Building Sheet 3 ===")
# Sort seasons
season_order = ['24 SPR','24 SUM','24 FAL','24 HOL',
                '25 SPR','25 SUM','25 FAL','25 HOL',
                '26 BTS','26 SPR','26 SUM','26 FAL','26 HOL',
                '27 SPR']

# Pivot: 시즌 × 브랜드, values = PO금액 + 직접원가 AMOUNT
pivot_po = df.pivot_table(index='시즌', columns='브랜드', values='PO금액', aggfunc='sum', fill_value=0)
pivot_cost = df.pivot_table(index='시즌', columns='브랜드', values='직접원가 AMOUNT', aggfunc='sum', fill_value=0)
pivot_profit = df.pivot_table(index='시즌', columns='브랜드', values='PROFIT AMOUNT', aggfunc='sum', fill_value=0)
pivot_qty = df.pivot_table(index='시즌', columns='브랜드', values='PO수량', aggfunc='sum', fill_value=0)

# Filter seasons - remove empty strings
pivot_po = pivot_po[pivot_po.index != '']
# Reindex with season order
pivot_po = pivot_po.reindex([s for s in season_order if s in pivot_po.index])
pivot_po = pivot_po.fillna(0)
pivot_po['TOTAL'] = pivot_po.sum(axis=1)

# ===============================================
# SHEET 4: Contract Factory별 원가 분석 (Template: YAKJIN)
# Factory 기준 with comparison columns
# ===============================================
print("\n=== Building Sheet 4 ===")
grp4 = df.groupby(['Contract Factory']).agg({
    'PO금액': 'sum',
    '직접원가 AMOUNT': 'sum',
    'PROFIT AMOUNT': 'sum',
    'CMT 계': 'sum',
    '해상운임': 'sum',
    'PO수량': 'sum',
}).reset_index()

m4 = grp4['PO금액'] > 0
grp4['직접원가율(%)'] = np.where(m4, (grp4['직접원가 AMOUNT'] / grp4['PO금액'] * 100).round(2), 0)
grp4['마진율(%)'] = np.where(m4, (grp4['PROFIT AMOUNT'] / grp4['PO금액'] * 100).round(2), 0)
grp4['CMT비중(%)'] = np.where(grp4['직접원가 AMOUNT'] > 0,
    (grp4['CMT 계'] / grp4['직접원가 AMOUNT'] * 100).round(2), 0)
grp4 = grp4.sort_values('PO금액', ascending=False)

# ===============================================
# Write to Excel with better formatting
# ===============================================
print("\n=== Writing Excel ===")
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

with pd.ExcelWriter(OUT, engine='openpyxl') as writer:
    # Sheet 1
    agg1.to_excel(writer, sheet_name='1.시즌별 부서브랜드 원가', index=False)
    
    # Sheet 2  
    grp2.to_excel(writer, sheet_name='2.브랜드별 원가구성 분석', index=False)
    
    # Sheet 3
    pivot_po.to_excel(writer, sheet_name='3.시즌별 브랜드 Pivot')
    
    # Sheet 4
    grp4.to_excel(writer, sheet_name='4.공장별 원가 분석', index=False)

# Now apply formatting
from openpyxl import load_workbook
wb = load_workbook(OUT)

thin = Side(style='thin')
header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
header_font = Font(bold=True, color='FFFFFF', size=10)
title_font = Font(bold=True, size=14)

for ws in wb.worksheets:
    # Auto-width
    for col_idx, col_cells in enumerate(ws.columns, 1):
        max_len = 0
        for cell in col_cells:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 25)
    
    # Header styling (row 1)
    for cell in ws[1]:
        if cell.value:
            cell.font = header_font
            cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
            cell.alignment = Alignment(horizontal='center', wrap_text=True)

wb.save(OUT)
print(f"\n✅ Done! Saved to: {OUT}")
print(f"  Sheet 1: {len(agg1)} rows")
print(f"  Sheet 2: {len(grp2)} rows")
print(f"  Sheet 3: {len(pivot_po)} seasons x {len(pivot_po.columns)-1} brands + Total")
print(f"  Sheet 4: {len(grp4)} factories")
