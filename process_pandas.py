"""Fast processing using pandas for both read and write"""
import pandas as pd
import numpy as np
from datetime import datetime

src_path = r'C:\Users\user\Downloads\브랜드별 직접원가 (7).xlsx'
out_path = r'C:\Users\user\.openclaw\workspace\영업데이터파일_raw_20260702.xlsx'

print('Reading data...')
# Read all data
df = pd.read_excel(src_path, header=0, engine='openpyxl')
print(f'Read {len(df)} rows, {len(df.columns)} cols')

# Print columns
for i, c in enumerate(df.columns):
    print(f'  [{i}] {repr(c)}')

# Build target columns based on Raw_20260608 format
# Map source columns by position
col_idx = {str(c).strip().lower(): i for i, c in enumerate(df.columns)}

# Find source column indices
def find(patterns):
    for p in patterns:
        pl = p.lower()
        for name, idx in col_idx.items():
            if pl in name:
                return idx
    return None

# Read source data as numpy arrays for speed
data = df.values

# Define factory mapping function
def get_factory(contract):
    if pd.isna(contract):
        return ''
    cf = str(contract).upper()
    mapping = {
        'YAKJIN CAMBO': '약진캄보',
        'YAKJIN VIETNAM': '약진베트남',
        'YAKJIN INTERTEX': '약진인터텍스',
        'YAKJIN GUATEMALA': '약진과테',
        'ARON YBL': '외창(VTN)',
        'VINA KOREA': '비나코리아',
        'JS VINA': 'JS VINA',
        'MICHIGAN HAIDUONG': '미시간하이증',
        'MICHIGAN HEE': '미시간히이정',
        'HANA KOVI': '하나코비',
        'TDT': 'TDT',
        'YTC': 'YTC',
        'HA PHONG': '하풍',
        'DASAN': 'DASAN',
        'WILLBES': 'WILLBES',
        'YEON HEUNG': '연흥',
        'INJAE': '인제',
        'GSE GUATEMALA': 'GSE과테',
        'PT YAKJIN JAYA 2': '자야2',
        'PT YAKJIN JAYA': '자야',
        'YAKJIN': '약진',
    }
    for key, val in mapping.items():
        if key in cf:
            return val
    return cf.split()[0] if cf.split() else cf[:20]

# Derive month function
def get_month(val):
    if pd.isna(val):
        return ''
    if hasattr(val, 'month'):
        return f'{val.month}월'
    return ''

# Build target data
# Target columns (58):
# 1=NO, 2=Contract Factory, 3=Factory, 4=바이어, 5=브랜드, 6=부서, 7=바이어스타일, 8=스타일, 
# 9=원가번호, 10=Grouped Div., 11=Division, 12=디비전, 13=시즌, 14-16=ITEM명, 17=담당, 
# 18=진행구분, 19=오더, 20=PO Status, 21=선적일, 22=월, 23=ex-Fact, 24=DC, 25=SHIP MODE,
# 26=TERMS, 27=PO수량, 28=PO금액, 29=매출할인, 30=순매출, 31-58 = cost columns

src_cols_map = [
    (1, 'NO'),          # 1
    (2, 'contract'),    # 2
    (3, None),          # 3 Factory - derive
    (4, '바이어'),      # 4
    (5, '브랜드'),      # 5
    (6, '부서'),        # 6
    (7, '바이어스타일'), # 7
    (8, '스타일'),      # 8
    (9, '원가번호'),    # 9
    (10, 'grouped'),    # 10
    (11, '디비젼'),     # 11 - Division (from 디비전)
    (12, '디비젼'),     # 12 - 디비전 (same data)
    (13, '시즌'),       # 13 - season
    (14, 'ITEM명'),     # 14 - ITEM명 (first)
    (15, 'ITEM명.1'),   # 15 - ITEM명.1
    (16, 'ITEM명.2'),   # 16 - ITEM명.2
    (17, '담당'),       # 17
    (18, '진행구분'),   # 18
    (19, '오더'),       # 19
    (20, 'po status'),  # 20
    (21, '선적일'),     # 21
    (22, None),         # 22 월 - derive
    (23, 'ex-fact'),    # 23
    (24, 'DC'),         # 24
    (25, 'ship mode'),  # 25
    (26, 'TERMS'),      # 26
    (27, 'PO수량'),     # 27
    (28, 'PO금액'),     # 28
    (29, '매출할인'),   # 29
    (30, None),         # 30 순매출 - derive
]

# Find source column indices  
src_idx = {}
for target_col, col_name in src_cols_map:
    if col_name:
        # Find the column
        cl = col_name.lower()
        found = False
        for name, idx in col_idx.items():
            if cl in name:
                src_idx[target_col] = idx
                found = True
                break
        if not found:
            print(f'  WARNING: Column "{col_name}" not found!')

print('\nSource column mapping:')
for t, c in sorted(src_idx.items()):
    print(f'  Target col {t:2d} <- Source idx {c:2d} = {df.columns[c]}')

# Build target array
n = len(df)
target = np.empty((n, 58), dtype=object)

print(f'\nProcessing {n} rows...')

for i in range(n):
    # Simple columns (most data)
    for t_col, name in src_cols_map:
        if name and t_col in src_idx:
            val = data[i, src_idx[t_col]]
            target[i, t_col - 1] = val if not pd.isna(val) else None
    
    # Factory (col 3)
    contract_val = data[i, src_idx[2]] if 2 in src_idx else None
    target[i, 2] = get_factory(contract_val)
    
    # Month (col 22)
    ship_date = data[i, src_idx[21]] if 21 in src_idx else None
    target[i, 21] = get_month(ship_date)
    
    # Net sales (col 30) = PO금액 - 매출할인
    po_amt = float(data[i, src_idx[28]] if 28 in src_idx else 0) if not pd.isna(data[i, src_idx[28]]) else 0
    disc = float(data[i, src_idx[29]] if 29 in src_idx else 0) if not pd.isna(data[i, src_idx[29]]) else 0
    target[i, 29] = po_amt - disc
    
    # Remaining cost columns (31-58) need to be figured out
    # For now, leave them at 0 - will be mapped properly based on actual column positions

# Build 58 column headers
headers_58 = [
    'NO', 'Contract\nFactory', 'Factory', '바이어', '브랜드', '부서', '바이어스타일', '스타일',
    '원가번호', 'Grouped Div.', 'Division', '디비전', '시즌', 'ITEM명', 'ITEM명', 'ITEM명',
    '담당', '진행구분', '오더', 'PO Status', '선적일', '월', 'ex-Fact', 'DC', 'SHIP MODE',
    'TERMS', 'PO수량', 'PO금액', '매출할인', '순매출',
]
# Add remaining cost columns (columns 31-58 in target = source columns 26-53)
for c in range(26, 54):  # source columns 26..53
    if c < len(df.columns):
        headers_58.append(str(df.columns[c]))
    else:
        headers_58.append(f'Col{c+1}')

assert len(headers_58) == 58, f'Expected 58 headers, got {len(headers_58)}'

# Create DataFrame
df_target = pd.DataFrame(target, columns=headers_58)

print('Saving to Excel...')
# Save to new file
with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
    # Write Sheet1 (original data)
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    # Write raw_0608 (processed)
    df_target.to_excel(writer, sheet_name='raw_0608', index=False)

print(f'Saved to {out_path}')
print(f'Sheet1: {len(df)} rows, {len(df.columns)} cols (original)')
print(f'raw_0608: {len(df_target)} rows, {len(df_target.columns)} cols (processed)')

# Quick verify
print(f'\nSample data (raw_0608, row 1):')
for c in [3, 11, 13, 22, 30]:
    col_name = headers_58[c-1]
    print(f'  Col {c} ({col_name}): {df_target.iloc[0, c-1]}')
