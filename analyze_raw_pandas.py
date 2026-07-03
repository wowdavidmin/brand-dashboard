import pandas as pd
import numpy as np

raw_path = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---47ea1d42-8317-4a51-a9fa-d5c2e2764f59.xlsx'

# Read specific columns only, filter for 26 season
df = pd.read_excel(raw_path, 
                   usecols=[3, 4, 10, 23, 24, 25, 40, 46],  # 0-indexed
                   header=0)

# Rename columns
df.columns = ['바이어', '브랜드', '시즌', 'PO수량', 'PO금액', '매출할인', 'CMT계', 'PROFIT']

# Filter to only 26 season
df_26 = df[df['시즌'].str.startswith('26 ', na=False)].copy()

print(f'Total rows (26 season): {len(df_26)}')
print()

# Group by 브랜드 and 시즌
grouped = df_26.groupby(['브랜드', '시즌']).agg({
    'PO수량': 'sum',
    'PO금액': 'sum',
    '매출할인': 'sum',
    'CMT계': 'sum',
    'PROFIT': 'sum'
}).reset_index()

# Get unique brands
brands = sorted(df_26['브랜드'].unique())
print('=== Brands in 26 season data ===')
for b in brands:
    print(f'  {b}')
print()

# Get unique departments
depts = sorted(df_26['바이어'].unique()) if '바이어' in df_26.columns else ['N/A']
print(f'Depts (buver): {depts}')

# Check URBAN
urban = df_26[df_26['브랜드'].str.contains('URBAN|urban', na=False)]
print(f'\nURBAN related rows: {len(urban)}')
if len(urban) > 0:
    print(urban[['브랜드', '시즌', 'PO수량', 'PO금액']].head(20))

# Check AERIE FAL
aerie = df_26[(df_26['브랜드'] == 'AERIE') & (df_26['시즌'] == '26 FAL')]
print(f'\nAERIE 26 FAL rows: {len(aerie)}, PO수량={aerie["PO수량"].sum()}, PO금액={aerie["PO금액"].sum()}')

# Print summary for key BU3 brands
bu3_brands = ['BANANA REPUBLIC', 'ON IPSS', "KOHL'S", 'AERIE', 'Offline', 'MLB']
print(f'\n=== BU3 brand summary (PO수량) ===')
for b in bu3_brands:
    sub = grouped[grouped['브랜드'] == b]
    if len(sub) > 0:
        total_qty = sub['PO수량'].sum()
        print(f'{b}: total_qty={total_qty:>10.0f}')
        for _, row in sub.iterrows():
            print(f'  {row["시즌"]}: qty={row["PO수량"]:>10.0f}, amt={row["PO금액"]:>12.2f}')
    else:
        print(f'{b}: NOT FOUND')

# Also print all brand+season combos
print(f'\n=== All 브랜드 x 시즌 (PO수량) ===')
for _, row in grouped.iterrows():
    qty = row['PO수량']
    amt = row['PO금액']
    print(f'{row["브랜드"]:25s} | {row["시즌"]:6s} | qty={qty:>10.0f} | amt={amt:>12.2f}')
