"""Fast Excel processing with xlsxwriter for output"""
import pandas as pd

src_path = r'C:\Users\user\Downloads\브랜드별 직접원가 (7).xlsx'
out_path = r'C:\Users\user\.openclaw\workspace\영업데이터파일_raw_20260702.xlsx'

print('Reading...')
df = pd.read_excel(src_path, header=0, engine='openpyxl')
n = len(df)
print(f'Read {n} rows, {len(df.columns)} cols')

col_src = list(df.columns)

# Factory mapping
def map_factory(series):
    cf = series.fillna('').astype(str).str.upper()
    result = pd.Series(index=cf.index, dtype=str)
    result[:] = ''
    maps = [
        ('YAKJIN CAMBO','약진캄보'),('YAKJIN VIETNAM','약진베트남'),
        ('YAKJIN INTERTEX','약진인터텍스'),('YAKJIN GUATEMALA','약진과테'),
        ('ARON YBL','외창(VTN)'),('VINA KOREA','비나코리아'),
        ('JS VINA','JS VINA'),('MICHIGAN','미시간'),
        ('HANA KOVI','하나코비'),('TDT','TDT'),('YTC','YTC'),
        ('HA PHONG','하풍'),('DASAN','DASAN'),('WILLBES','WILLBES'),
        ('YEON HEUNG','연흥'),('INJAE','인제'),
        ('GSE GUATEMALA','GSE과테'),('PT YAKJIN JAYA 2','자야2'),
        ('PT YAKJIN JAYA','자야'),('YAKJIN','약진'),
    ]
    for kw, val in maps:
        result[cf.str.contains(kw, na=False)] = val
    result[result==''] = cf[result==''].str.split().str[0]
    return result

# Build target DataFrame
target_cols = [
    'NO', 'Contract\nFactory', 'Factory', '바이어', '브랜드', '부서',
    '바이어스타일', '스타일', '원가번호', 'Grouped Div.', 'Division', '디비전',
    '시즌', 'ITEM명', 'ITEM명.1', 'ITEM명.2', '담당', '진행구분', '오더',
    'PO Status', '선적일', '월', 'ex-Fact', 'DC', 'SHIP MODE', 'TERMS',
    'PO수량', 'PO금액', '매출할인', '순매출'
]
for c in range(26, 54):
    target_cols.append(str(col_src[c]))

target = pd.DataFrame(index=range(n), columns=target_cols, dtype=object)

# Copy columns
simple_map = {
    'NO': 0, 'Contract\nFactory': 1,
    '바이어': 3, '브랜드': 4, '부서': 5,
    '바이어스타일': 6, '스타일': 7, '원가번호': 8, 'Grouped Div.': 9,
    '담당': 16, '진행구분': 17, '오더': 18, 'PO Status': 19, '선적일': 20,
    'ex-Fact': 22, 'DC': 23, 'SHIP MODE': 24, 'TERMS': 25,
    'PO수량': 27, 'PO금액': 28, '매출할인': 29,
}
# Fix: PO수량=26, PO금액=27, 매출할인=28 in my mapping (0-based)
# Actually let me recalculate: 
# 0=NO,1=Contract,2=Factory,3=바이어,... Target col 27 = PO수량 (0-based index 26)
simple_map = {
    'NO': 0, 'Contract\nFactory': 1,
    '바이어': 3, '브랜드': 4, '부서': 5,
    '바이어스타일': 6, '스타일': 7, '원가번호': 8, 'Grouped Div.': 9,
    '담당': 16, '진행구분': 17, '오더': 18, 'PO Status': 19, '선적일': 20,
    'ex-Fact': 22, 'DC': 23, 'SHIP MODE': 24, 'TERMS': 25,
    'PO수량': 26, 'PO금액': 27, '매출할인': 28,
}

# Map using the column index in the SOURCE
col_index_src = {str(c): i for i, c in enumerate(col_src)}

for src_name, tgt_idx in simple_map.items():
    if src_name in col_index_src:
        si = col_index_src[src_name]
        target.iloc[:, tgt_idx] = df.iloc[:, si].values

# Division (C11 = index 10) = Grouped Div. (C10 = index 9)
target.iloc[:, 10] = df.iloc[:, 8].values  # Grouped Div.

# 디비전 (C12 = index 11) = 디비젼
if '디비젼' in col_index_src:
    target.iloc[:, 11] = df.iloc[:, col_index_src['디비젼']].values

# ITEM명 (C14-C16 = indices 13-15)
item_cols = [c for c in col_src if 'ITEM' in str(c).upper()]
for i, ic in enumerate(item_cols[:3]):
    target.iloc[:, 13 + i] = df[ic].values

# 시즌 (C13 = index 12)
if '시즌' in col_index_src:
    target.iloc[:, 12] = df.iloc[:, col_index_src['시즌']].values

# Factory
print('Factory...')
target.iloc[:, 2] = map_factory(df.iloc[:, 1])

# Month
print('Month...')
ship = df.iloc[:, 18]
target.iloc[:, 21] = ship.apply(lambda x: f'{x.month}월' if pd.notna(x) and hasattr(x, 'month') else '')

# Net sales
print('Net sales...')
po = pd.to_numeric(df.iloc[:, 24], errors='coerce').fillna(0)
disc = pd.to_numeric(df.iloc[:, 25], errors='coerce').fillna(0)
target.iloc[:, 29] = po - disc

# Cost columns
print('Cost columns...')
for src_i in range(26, 54):
    tgt_i = src_i + 4  # source 26 -> target 30 (index 29)
    target.iloc[:, tgt_i] = df.iloc[:, src_i].values

# Write with xlsxwriter (fast)
print('Writing Excel...')
with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    target.to_excel(writer, sheet_name='raw_0608', index=False)

print(f'Saved to {out_path}')
print(f'Sheet1: {df.shape}, raw_0608: {target.shape}')

# Quick verify
print('\n=== Verification ===')
print(f'Factory: {target.iloc[0, 2]}')
print(f'Division: {target.iloc[0, 10]}')
print(f'디비전: {target.iloc[0, 11]}')
print(f'시즌: {target.iloc[0, 12]}')
print(f'월: {target.iloc[0, 21]}')
print(f'순매출: {target.iloc[0, 29]}')
print(f'Total rows: {len(target)}')
