"""Final processing - hardcode known column positions"""
import pandas as pd
import numpy as np

src_path = r'C:\Users\user\Downloads\브랜드별 직접원가 (7).xlsx'
out_path = r'C:\Users\user\.openclaw\workspace\영업데이터파일_raw_20260702.xlsx'

print('Reading data...')
df = pd.read_excel(src_path, header=0, engine='openpyxl')
n = len(df)
print(f'Read {n} rows, {len(df.columns)} cols')

# Source column indices (0-based) from the file
S = {
    'NO': 0, 'CONTRACT': 1, '바이어': 2, '브랜드': 3, '부서': 4,
    '바이어스타일': 5, '스타일': 6, '원가번호': 7, 'GROUPED': 8, '디비젼': 9,
    '시즌': 10, 'ITEM1': 11, 'ITEM2': 12, 'ITEM3': 13,
    '담당': 14, '진행구분': 15, '오더': 16, 'PO_STATUS': 17, '선적일': 18,
    'EX_FACT': 19, 'DC': 20, 'SHIP_MODE': 21, 'TERMS': 22,
    'PO수량': 23, 'PO금액': 24, '매출할인': 25
}

# Factory mapping
def get_factory(cf):
    if pd.isna(cf): return ''
    cf = str(cf).upper()
    m = {
        'YAKJIN CAMBO': '약진캄보', 'YAKJIN VIETNAM': '약진베트남',
        'YAKJIN INTERTEX': '약진인터텍스', 'YAKJIN GUATEMALA': '약진과테',
        'ARON YBL': '외창(VTN)', 'VINA KOREA': '비나코리아',
        'JS VINA': 'JS VINA', 'MICHIGAN H': '미시간',
        'HANA KOVI': '하나코비', 'TDT': 'TDT', 'YTC': 'YTC',
        'HA PHONG': '하풍', 'DASAN': 'DASAN', 'WILLBES': 'WILLBES',
        'YEON HEUNG': '연흥', 'INJAE': '인제',
        'GSE GUATEMALA': 'GSE과테', 'PT YAKJIN JAYA 2': '자야2',
        'PT YAKJIN JAYA': '자야', 'YAKJIN': '약진',
    }
    for k, v in m.items():
        if k in cf: return v
    return cf.split()[0] if cf.split() else cf[:20]

def get_month(val):
    if pd.isna(val) or not hasattr(val, 'month'): return ''
    return f'{val.month}월'

# Get data as numpy array for fast access
data = df.values

# Build target: 58 columns
T = np.empty((n, 58), dtype=object)

# Map simple columns
simple_map = [
    (0, 'NO'), (1, 'CONTRACT'), (3, '바이어'), (4, '브랜드'),
    (5, '부서'), (6, '바이어스타일'), (7, '스타일'), (8, '원가번호'),
    (9, 'GROUPED'), (10, '디비젼'), (11, '디비젼'),  # Division + 디비전
    (12, '시즌'),                                    # 시즌
    (13, 'ITEM1'), (14, 'ITEM2'), (15, 'ITEM3'),    # ITEM명
    (16, '담당'), (17, '진행구분'), (18, '오더'),
    (19, 'PO_STATUS'), (20, '선적일'),
    (22, 'EX_FACT'), (23, 'DC'), (24, 'SHIP_MODE'), (25, 'TERMS'),
    (26, 'PO수량'), (27, 'PO금액'), (28, '매출할인'),
]

for t_idx, s_key in simple_map:
    s_idx = S[s_key]
    for i in range(n):
        val = data[i, s_idx]
        T[i, t_idx] = val if not pd.isna(val) else None

print('Processing Factory, Month, Net Sales...')
for i in range(n):
    # Factory (col 3, index 2)
    cf = data[i, S['CONTRACT']]
    T[i, 2] = get_factory(cf)
    
    # Month (col 22, index 21)
    ship = data[i, S['선적일']]
    T[i, 21] = get_month(ship)
    
    # Net sales (col 30, index 29)
    po = float(data[i, S['PO금액']] or 0) if not pd.isna(data[i, S['PO금액']]) else 0
    disc = float(data[i, S['매출할인']] or 0) if not pd.isna(data[i, S['매출할인']]) else 0
    T[i, 29] = po - disc

# Copy remaining cost columns (source indices 26-53 -> target 30-57)
print('Copying cost columns...')
for s_col in range(26, 54):
    t_col = s_col + 4  # target index: source 26 -> target 30 (index 29)
    for i in range(n):
        val = data[i, s_col]
        T[i, t_col] = val if not pd.isna(val) else None

# Build headers
headers = [
    'NO', 'Contract\nFactory', 'Factory', '바이어', '브랜드', '부서',
    '바이어스타일', '스타일', '원가번호', 'Grouped Div.', 'Division', '디비전',
    '시즌', 'ITEM명', 'ITEM명', 'ITEM명', '담당', '진행구분', '오더',
    'PO Status', '선적일', '월', 'ex-Fact', 'DC', 'SHIP MODE', 'TERMS',
    'PO수량', 'PO금액', '매출할인', '순매출'
]
for c in range(26, 54):
    headers.append(str(df.columns[c]))

print(f'Target dimension: {T.shape}, headers: {len(headers)}')

# Create DataFrame
df_target = pd.DataFrame(T, columns=headers)

print('Saving...')
with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    df_target.to_excel(writer, sheet_name='raw_0608', index=False)

print(f'Saved to {out_path}')
print(f'Sheet1: {len(df)} cols, raw_0608: {len(df_target)} cols')

# Verify
print('\n=== Verification (raw_0608 row 0) ===')
for c in [3, 11, 13, 22, 30]:
    print(f'  Col {c} ({headers[c-1]}): {df_target.iloc[0, c-1]}')
