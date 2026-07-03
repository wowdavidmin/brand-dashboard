"""Vectorized processing - no Python loops"""
import pandas as pd
import numpy as np

src_path = r'C:\Users\user\Downloads\브랜드별 직접원가 (7).xlsx'
out_path = r'C:\Users\user\.openclaw\workspace\영업데이터파일_raw_20260702.xlsx'

print('Reading...')
df = pd.read_excel(src_path, header=0, engine='openpyxl')
n = len(df)
print(f'Read {n} rows, {len(df.columns)} cols')

# Factory mapping function
def map_factory(series):
    cf = series.fillna('').astype(str).str.upper()
    result = pd.Series(index=cf.index, dtype=str)
    result[:] = ''
    
    mappings = [
        ('YAKJIN CAMBO', '약진캄보'), ('YAKJIN VIETNAM', '약진베트남'),
        ('YAKJIN INTERTEX', '약진인터텍스'), ('YAKJIN GUATEMALA', '약진과테'),
        ('ARON YBL', '외창(VTN)'), ('VINA KOREA', '비나코리아'),
        ('JS VINA', 'JS VINA'), ('MICHIGAN', '미시간'),
        ('HANA KOVI', '하나코비'), ('TDT', 'TDT'), ('YTC', 'YTC'),
        ('HA PHONG', '하풍'), ('DASAN', 'DASAN'), ('WILLBES', 'WILLBES'),
        ('YEON HEUNG', '연흥'), ('INJAE', '인제'),
        ('GSE GUATEMALA', 'GSE과테'), ('PT YAKJIN JAYA 2', '자야2'),
        ('PT YAKJIN JAYA', '자야'), ('YAKJIN', '약진'),
    ]
    
    for kw, val in mappings:
        mask = cf.str.contains(kw, na=False)
        result[mask] = val
    
    # Fill remaining with first word
    remaining = result == ''
    if remaining.any():
        result[remaining] = cf[remaining].str.split().str[0]
    
    return result

print('Building target DataFrame...')
col_src = list(df.columns)

# Build target: 58 columns
target_cols = [
    'NO', 'Contract\nFactory', 'Factory', '바이어', '브랜드', '부서',
    '바이어스타일', '스타일', '원가번호', 'Grouped Div.', 'Division', '디비전',
    '시즌', 'ITEM명', 'ITEM명', 'ITEM명', '담당', '진행구분', '오더',
    'PO Status', '선적일', '월', 'ex-Fact', 'DC', 'SHIP MODE', 'TERMS',
    'PO수량', 'PO금액', '매출할인', '순매출'
]
# Add remaining columns (source indices 26-53 -> target 31-58)
for i in range(26, 54):
    target_cols.append(str(col_src[i]))

# Source to target column mapping
src_to_tgt = {
    'NO': 0, 'Contract\nFactory': 1, '바이어': 3, '브랜드': 4, '부서': 5,
    '바이어스타일': 6, '스타일': 7, '원가번호': 8, 'Grouped Div.': 9,
    'Grouped Div.': 10,  # Division
    '디비젼': 11,  # 디비전
    '시즌': 12,
}

# Proper ITEM mapping
item_cols = [c for c in col_src if 'ITEM' in str(c).upper()]
for i, item_col in enumerate(item_cols[:3]):
    src_to_tgt[item_col] = 13 + i

other_maps = {
    '담당': 16, '진행구분': 17, '오더': 18, 'PO Status': 19, '선적일': 20,
    'ex-Fact': 22, 'DC': 23, 'SHIP MODE': 24, 'TERMS': 25,
    'PO수량': 26, 'PO금액': 27, '매출할인': 28,
}

# Initialize target DataFrame with zeros/empty
T = pd.DataFrame(index=range(n), columns=target_cols, dtype=object)

# Copy source columns to target
for src_name, tgt_idx in src_to_tgt.items():
    if src_name in col_src:
        col_data = df[src_name].values
        if isinstance(tgt_idx, list):
            for ti in tgt_idx:
                T.iloc[:, ti] = col_data
        else:
            T.iloc[:, tgt_idx] = col_data

for src_name, tgt_idx in other_maps.items():
    if src_name in col_src:
        T.iloc[:, tgt_idx] = df[src_name].values

# Also copy ITEM명 columns
for i, item_col in enumerate(item_cols[:3]):
    T.iloc[:, 13 + i] = df[item_col].values

print('Adding Factory...')
# Factory (col 3, index 2)
T.iloc[:, 2] = map_factory(df.iloc[:, 1])  # Contract Factory

print('Adding Month...')
# Month (col 22, index 21)
ship = df.iloc[:, 18]  # 선적일
T.iloc[:, 21] = ship.apply(lambda x: f'{x.month}월' if pd.notna(x) and hasattr(x, 'month') else '')

print('Adding Net Sales...')
# Net sales (col 30, index 29) = PO금액 - 매출할인
po = pd.to_numeric(df.iloc[:, 24], errors='coerce').fillna(0)
disc = pd.to_numeric(df.iloc[:, 25], errors='coerce').fillna(0)
T.iloc[:, 29] = po - disc

# Copy remaining cost columns
print('Copying cost columns...')
for src_i in range(26, 54):
    tgt_i = src_i + 4  # source 26 -> target 30 (index 29, but 0-based -> 30)
    T.iloc[:, tgt_i] = df.iloc[:, src_i].values

print('Saving...')
with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    T.to_excel(writer, sheet_name='raw_0608', index=False)

print(f'Saved to {out_path}')
print(f'Sheet1: {df.shape}, raw_0608: {T.shape}')

# Verify
print('\n=== Verification (raw_0608 row 0) ===')
for c_idx, name in [(2, 'Factory'), (10, 'Division'), (12, '시즌'), (21, '월'), (29, '순매출')]:
    print(f'  {name}: {T.iloc[0, c_idx]}')
print(f'  Total rows: {len(T)}')
