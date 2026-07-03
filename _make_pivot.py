"""
브랜드별_직접원가_raw_v1.0_20260629.xlsx → 4개 피벗 분석 시트 생성
Template format: 20212026_영업_DATA_FILE_working (1~4번 시트 구조 참조)
"""
import pandas as pd
import numpy as np
import os, warnings
warnings.filterwarnings('ignore')

SRC = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---b99ec911-9409-4c91-ae86-0a63ab160c8c.xlsx'
OUT = r'C:\Users\user\.openclaw\workspace\브랜드별_직접원가_피벗분석_v1.0_20260629.xlsx'

print("=== Reading raw data ===")
df = pd.read_excel(SRC, sheet_name='Raw_20260629')
print(f"Shape: {df.shape}")
print(f"Columns: {list(df.columns)}")

# Clean column names
col_map = {}
for c in df.columns:
    clean = str(c).replace('\n','').strip()
    col_map[c] = clean
df = df.rename(columns=col_map)

# Key numeric columns
num_cols = ['PO수량', 'PO금액', '매출할인', '원자재비\n원사', '원자재비\n원단', '부자재비\n본사',
            '염가공', 'I.F염가공', '염가공비 계', '임가공료\n외화', '임가공료\n원화', '임가공료\n합계',
            'CMT\n봉제료', 'CMT\nI.F', 'CMT\n부자재', 'CMT\n임가공료해외', 'CMT\n기타수수료', 'CMT 계',
            '해상운임', 'Inspection Fee', 'NCPM', '직접원가\nAMOUNT', '직접원가 %',
            'PROFIT\nAMOUNT', 'PROFIT %', 'Buyer\nCommission', 'Agent\nCommission', 'Duty',
            'Execute\nMargin', 'Execute\nMargin(%)']
for c in num_cols:
    df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

# Clean text cols
for c in ['브랜드', '부서', '시즌', '바이어', 'Contract\nFactory', 'Grouped Div.', '디비젼', '담당', '진행구분', '손익']:
    if c in df.columns:
        df[c] = df[c].astype(str).str.strip()

# Print unique values for key columns
print(f"\n부서 unique: {df['부서'].unique()}")
print(f"시즌 unique: {df['시즌'].unique()}")
print(f"브랜드 unique (count): {df['브랜드'].nunique()}")
print(f"Contract Factory unique (count): {df['Contract\nFactory'].nunique()}")

# ==============================================================
# Sheet 1: 시즌별 부서/브랜드 직접원가 비교 (Template Sheet 1 구조 참조)
# ==============================================================
print("\n=== Sheet 1: 시즌별 부서/브랜드 직접원가 비교 ===")
sheet1 = df.groupby(['부서', '브랜드', '시즌']).agg({
    'PO금액': 'sum',
    '직접원가\nAMOUNT': 'sum',
    'PROFIT\nAMOUNT': 'sum',
    'PO수량': 'sum',
    'CMT 계': 'sum',
    '해상운임': 'sum',
}).reset_index()

sheet1['직접원가율'] = np.where(sheet1['PO금액'] > 0, 
    (sheet1['직접원가\nAMOUNT'] / sheet1['PO금액'] * 100).round(2), 0)
sheet1['마진율'] = np.where(sheet1['PO금액'] > 0,
    (sheet1['PROFIT\nAMOUNT'] / sheet1['PO금액'] * 100).round(2), 0)

# ==============================================================
# Sheet 2: 브랜드별 직접원가 구성 분석 (Template Sheet 2 구조 참조)
# ==============================================================
print("\n=== Sheet 2: 브랜드별 직접원가 구성 분석 ===")
sheet2 = df.groupby(['브랜드']).agg({
    'PO금액': 'sum',
    '원자재비\n원사': 'sum',
    '원자재비\n원단': 'sum',
    '부자재비\n본사': 'sum',
    '염가공비 계': 'sum',
    '임가공료\n합계': 'sum',
    'CMT 계': 'sum',
    '해상운임': 'sum',
    '직접원가\nAMOUNT': 'sum',
    'PROFIT\nAMOUNT': 'sum',
    'PO수량': 'sum',
}).reset_index()

# Calculate percentages
for col in ['원자재비\n원사', '원자재비\n원단', '부자재비\n본사', '염가공비 계', '임가공료\n합계', 'CMT 계', '해상운임']:
    sheet2[f'{col}_%'] = np.where(sheet2['직접원가\nAMOUNT'] > 0,
        (sheet2[col] / sheet2['직접원가\nAMOUNT'] * 100).round(2), 0)

sheet2['직접원가율'] = np.where(sheet2['PO금액'] > 0,
    (sheet2['직접원가\nAMOUNT'] / sheet2['PO금액'] * 100).round(2), 0)
sheet2['마진율'] = np.where(sheet2['PO금액'] > 0,
    (sheet2['PROFIT\nAMOUNT'] / sheet2['PO금액'] * 100).round(2), 0)

# ==============================================================
# Sheet 3: 시즌별 브랜드 PO금액 Pivot (Template Sheet 3 구조 참조)
# ==============================================================
print("\n=== Sheet 3: 시즌별 브랜드 PO금액 Pivot ===")
sheet3_pivot = df.pivot_table(
    index='시즌',
    columns='브랜드',
    values=['PO금액', '직접원가\nAMOUNT', 'PROFIT\nAMOUNT', 'PO수량'],
    aggfunc='sum',
    fill_value=0
)

# Flatten multi-level columns
sheet3_pivot.columns = [f'{col[0]}_{col[1]}' for col in sheet3_pivot.columns]
sheet3_pivot = sheet3_pivot.reset_index()

# Add season sort order
season_order = {'24 SPR':1,'24 SUM':2,'24 FAL':3,'24 HOL':4,'25 SPR':5,'25 SUM':6,'25 FAL':7,'25 HOL':8,
                '26 SPR':9,'26 SUM':10,'26 FAL':11,'26 HOL':12}

# ==============================================================
# Sheet 4: Contract Factory별 원가 분석 (Template Sheet 4 구조 참조)
# ==============================================================
print("\n=== Sheet 4: Contract Factory별 원가 분석 ===")
sheet4 = df.groupby(['Contract\nFactory', '브랜드']).agg({
    'PO금액': 'sum',
    'CMT 계': 'sum',
    '해상운임': 'sum',
    '직접원가\nAMOUNT': 'sum',
    'PROFIT\nAMOUNT': 'sum',
    'PO수량': 'sum',
}).reset_index()

sheet4_pivot_factory = df.groupby(['Contract\nFactory']).agg({
    'PO금액': 'sum',
    'CMT 계': 'sum',
    '해상운임': 'sum',
    '직접원가\nAMOUNT': 'sum',
    'PROFIT\nAMOUNT': 'sum',
    'PO수량': 'sum',
}).reset_index()
sheet4_pivot_factory['직접원가율'] = np.where(sheet4_pivot_factory['PO금액'] > 0,
    (sheet4_pivot_factory['직접원가\nAMOUNT'] / sheet4_pivot_factory['PO금액'] * 100).round(2), 0)
sheet4_pivot_factory['마진율'] = np.where(sheet4_pivot_factory['PO금액'] > 0,
    (sheet4_pivot_factory['PROFIT\nAMOUNT'] / sheet4_pivot_factory['PO금액'] * 100).round(2), 0)

# ==============================================================
# Write to Excel with formatting
# ==============================================================
print("\n=== Writing Excel ===")
with pd.ExcelWriter(OUT, engine='openpyxl') as writer:
    # Sheet 1
    sheet1.to_excel(writer, sheet_name='1. 시즌별 부서브랜드 원가', index=False)
    
    # Sheet 2
    sheet2.to_excel(writer, sheet_name='2. 브랜드별 원가구성 분석', index=False)
    
    # Sheet 3
    sheet3_pivot.to_excel(writer, sheet_name='3. 시즌별 브랜드 Pivot', index=False)
    
    # Sheet 4
    sheet4_pivot_factory.to_excel(writer, sheet_name='4. 공장별 원가 분석', index=False)

print(f"\n✅ 완료! 출력 파일: {OUT}")
print(f"  Sheet 1: {len(sheet1)} rows")
print(f"  Sheet 2: {len(sheet2)} rows")
print(f"  Sheet 3: {len(sheet3_pivot)} rows")
print(f"  Sheet 4: {len(sheet4_pivot_factory)} rows")
