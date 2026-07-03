"""
브랜드별_직접원가_raw → 4개 피벗 분석 시트 생성
"""
import pandas as pd
import numpy as np
import os, warnings
warnings.filterwarnings('ignore')

SRC = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---b99ec911-9409-4c91-ae86-0a63ab160c8c.xlsx'
OUT = r'C:\Users\user\.openclaw\workspace\브랜드별_직접원가_피벗분석_v1.0_20260629.xlsx'

print("Reading raw data...")
df = pd.read_excel(SRC, sheet_name='Raw_20260629')
print(f"Shape: {df.shape}")

# Clean column names: replace newlines with spaces
clean_cols = {c: c.replace('\n', ' ').strip() for c in df.columns}
df = df.rename(columns=clean_cols)
print("Cleaned columns:", list(df.columns))

# Now reference by cleaned names
c = lambda name: name.replace('\n', ' ') if '\n' not in name else name

# Numeric columns to convert
num_targets = [
    'PO수량', 'PO금액', '매출할인',
    '원자재비 원사', '원자재비 원단', '부자재비 본사',
    '염가공', 'I.F염가공', '염가공비 계',
    '임가공료 외화', '임가공료 원화', '임가공료 합계',
    'CMT 봉제료', 'CMT I.F', 'CMT 부자재',
    'CMT 임가공료해외', 'CMT 기타수수료', 'CMT 계',
    '해상운임', 'Inspection Fee', 'NCPM',
    '직접원가 AMOUNT', '직접원가 %',
    'PROFIT AMOUNT', 'PROFIT %',
    'Buyer Commission', 'Agent Commission', 'Duty',
    'Execute Margin', 'Execute Margin(%)'
]
for cname in num_targets:
    if cname in df.columns:
        df[cname] = pd.to_numeric(df[cname], errors='coerce').fillna(0)
        print(f"  Converted: {cname}")

# Clean text columns
for cname in ['브랜드', '부서', '시즌', '바이어', 'Contract Factory', 'Grouped Div.', '디비젼', '담당', '진행구분', '손익']:
    if cname in df.columns:
        df[cname] = df[cname].astype(str).str.strip()

print(f"\n부서: {df['부서'].unique()}")
print(f"시즌: {df['시즌'].unique()}")
print(f"브랜드 수: {df['브랜드'].nunique()}")
print(f"공장 수: {df['Contract Factory'].nunique() if 'Contract Factory' in df.columns else 0}")

# ==============================================================
# Sheet 1: 시즌별 부서/브랜드 직접원가 비교
# ==============================================================
print("\n=== Sheet 1 ===")
grp1 = df.groupby(['부서', '브랜드', '시즌']).agg({
    'PO금액': 'sum',
    '직접원가 AMOUNT': 'sum',
    'PROFIT AMOUNT': 'sum',
    'PO수량': 'sum',
    'CMT 계': 'sum',
    '해상운임': 'sum',
}).reset_index()

mask = grp1['PO금액'] > 0
grp1['직접원가율'] = np.where(mask, (grp1['직접원가 AMOUNT'] / grp1['PO금액'] * 100).round(2), 0)
grp1['마진율'] = np.where(mask, (grp1['PROFIT AMOUNT'] / grp1['PO금액'] * 100).round(2), 0)

# ==============================================================
# Sheet 2: 브랜드별 직접원가 구성 분석
# ==============================================================
print("\n=== Sheet 2 ===")
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

cost_cols = ['원자재비 원사', '원자재비 원단', '부자재비 본사', '염가공비 계', '임가공료 합계', 'CMT 계', '해상운임']
for col in cost_cols:
    grp2[f'{col} 구성비(%)'] = np.where(grp2['직접원가 AMOUNT'] > 0,
        (grp2[col] / grp2['직접원가 AMOUNT'] * 100).round(2), 0)

m2 = grp2['PO금액'] > 0
grp2['직접원가율(%)'] = np.where(m2, (grp2['직접원가 AMOUNT'] / grp2['PO금액'] * 100).round(2), 0)
grp2['마진율(%)'] = np.where(m2, (grp2['PROFIT AMOUNT'] / grp2['PO금액'] * 100).round(2), 0)

# ==============================================================
# Sheet 3: 시즌별 브랜드 Pivot (PO금액 기준)
# ==============================================================
print("\n=== Sheet 3 ===")
# Pivot: 시즌(rows) x 브랜드(cols), values=PO금액
pivot3 = df.pivot_table(
    index='시즌',
    columns='브랜드',
    values='PO금액',
    aggfunc='sum',
    fill_value=0
)
pivot3 = pivot3.reset_index()
pivot3['Total'] = pivot3.select_dtypes(include=[np.number]).sum(axis=1)

# Sort seasons
season_sorter = {
    '24 SPR':1,'24 SUM':2,'24 FAL':3,'24 HOL':4,
    '25 SPR':5,'25 SUM':6,'25 FAL':7,'25 HOL':8,
    '26 SPR':9,'26 SUM':10,'26 FAL':11,'26 HOL':12
}
pivot3['_sort'] = pivot3['시즌'].map(season_sorter).fillna(99)
pivot3 = pivot3.sort_values('_sort').drop(columns='_sort')

# ==============================================================
# Sheet 4: Contract Factory별 원가 분석
# ==============================================================
print("\n=== Sheet 4 ===")
grp4 = df.groupby(['Contract Factory']).agg({
    'PO금액': 'sum',
    'CMT 계': 'sum',
    '해상운임': 'sum',
    '직접원가 AMOUNT': 'sum',
    'PROFIT AMOUNT': 'sum',
    'PO수량': 'sum',
}).reset_index()

m4 = grp4['PO금액'] > 0
grp4['직접원가율(%)'] = np.where(m4, (grp4['직접원가 AMOUNT'] / grp4['PO금액'] * 100).round(2), 0)
grp4['마진율(%)'] = np.where(m4, (grp4['PROFIT AMOUNT'] / grp4['PO금액'] * 100).round(2), 0)
grp4['CMT비율(%)'] = np.where(grp4['직접원가 AMOUNT'] > 0,
    (grp4['CMT 계'] / grp4['직접원가 AMOUNT'] * 100).round(2), 0)

# ==============================================================
# Write Excel
# ==============================================================
print("\n=== Writing Excel ===")
with pd.ExcelWriter(OUT, engine='openpyxl') as writer:
    grp1.to_excel(writer, sheet_name='1.시즌별 부서브랜드 원가', index=False)
    grp2.to_excel(writer, sheet_name='2.브랜드별 원가구성 분석', index=False)
    pivot3.to_excel(writer, sheet_name='3.시즌별 브랜드 Pivot', index=False)
    grp4.to_excel(writer, sheet_name='4.공장별 원가 분석', index=False)

print(f"Done! -> {OUT}")
print(f"  Sheet1: {len(grp1)} rows")
print(f"  Sheet2: {len(grp2)} rows")
print(f"  Sheet3: {len(pivot3)} rows x {len(pivot3.columns)} cols")
print(f"  Sheet4: {len(grp4)} rows")
