"""Process new raw data to match Raw_20260608 format"""
import openpyxl
from copy import copy

src_path = r'C:\Users\user\Downloads\브랜드별 직접원가 (7).xlsx'
out_path = r'C:\Users\user\.openclaw\workspace\영업데이터파일_raw_20260702.xlsx'

# Read new raw data
wb = openpyxl.load_workbook(src_path)
ws_src = wb.worksheets[0]

print(f'Source: {ws_src.max_row} rows, {ws_src.max_column} cols')

# Create new sheet for processed data
ws_new = wb.create_sheet('raw_0608')

# Define Raw_20260608 column headers and positions (1-indexed)
# Template: [target_col, header_name, source_col_or_formula]
template = [
    (1, 'NO', 1),
    (2, 'Contract\nFactory', 2),
    (3, 'Factory', None),         # Need to derive
    (4, '바이어', 3),
    (5, '브랜드', 4),
    (6, '부서', 5),
    (7, '바이어스타일', 6),
    (8, '스타일', 7),
    (9, '원가번호', 8),
    (10, 'Grouped Div.', 9),
    (11, 'Division', 10),          # 디비젼 from source
    (12, '디비전', 10),            # Same data
    (13, '시즌', 11),              # 시즌 from source
    (14, 'ITEM명', 12),
    (15, 'ITEM명', 13),
    (16, 'ITEM명', 14),
    (17, '담당', 15),
    (18, '진행구분', 16),
    (19, '오더', 17),
    (20, 'PO Status', 18),
    (21, '선적일', 19),
    (22, '월', None),              # Need to derive
    (23, 'ex-Fact', 20),
    (24, 'DC', 21),
    (25, 'SHIP MODE', 22),
    (26, 'TERMS', 23),
    (27, 'PO수량', 24),
    (28, 'PO금액', 25),
    (29, '매출할인', 26),
    (30, '순매출', None),          # PO금액 - 매출할인
    (31, '원자재비\n원사', 27),
    (32, '원자재비\n원단', 28),
    (33, '부자재비\n본사', 29),
    (34, '염가공', 30),
    (35, 'I.F염가공', 31),
    (36, '염가공비 계', 32),
    (37, '임가공료\n외화', 33),
    (38, '임가공료\n원화', 34),
    (39, '임가공료\n합계', 35),
    (40, 'CMT\n봉제료', 36),
    (41, 'CMT\nI.F', 37),
    (42, 'CMT\n부자재', 38),
    (43, 'CMT\n임가공료해외', 39),
    (44, 'CMT\n기타수수료', 40),
    (45, 'CMT 계', 41),
    (46, '해상운임', 42),
    (47, 'Inspection Fee', 43),
    (48, 'NCPM', 44),
    (49, '직접원가\nAMOUNT', 45),
    (50, '직접원가 %', 46),
    (51, 'PROFIT\nAMOUNT', 47),
    (52, 'PROFIT %', 48),
    (53, '손익', 49),
    (54, 'Buyer\nCommission', 50),
    (55, 'Agent Commission', 51),
    (56, 'Duty', 52),
    (57, 'Execute\nMargin', 53),
    (58, 'Execute\nMargin(%)', 54),
]

# Write header row
for t_col, header, _ in template:
    ws_new.cell(1, t_col, header)

# Process data rows
# Build factory mapping from Contract Factory
def get_factory(contract_factory):
    """Derive short factory name from Contract Factory"""
    if not contract_factory:
        return ''
    cf = str(contract_factory).upper()
    
    # Common factory mapping
    mapping = {
        'YAKJIN': '약진',
        'VINA KOREA': '비나코리아',
        'JS VINA': 'JS VINA',
        'MICHIGAN': '미시간',
        'HANA KOVI': '하나코비',
        'TDT': 'TDT',
        'YTC': 'YTC',
        'HA PHONG': '하풍',
        'DASAN': 'DASAN',
        'WILLBES': 'WILLBES',
        'YEON HEUNG': '연흥',
        'INJAE': '인제',
        'GSE GUATEMALA': 'GSE과테',
        'YAKJIN CAMBO': '약진캄보',
        'YAKJIN VIETNAM': '약진베트남',
        'YAKJIN INTERTEX': '약진인터텍스',
        'PT YAKJIN JAYA': '자야',
        'PT. YAKJIN JAYA 2': '자야2',
        'ARON YBL': '외창(VTN)',
        'GARMENT': 'GARMENT',
    }
    
    for key, val in mapping.items():
        if key in cf:
            return val
    
    # Fallback: try to extract first meaningful part
    parts = cf.replace('CO.,LTD', '').replace('COMPANY', '').replace('LIMITED', '').replace(',', '').strip()
    parts = parts.split()
    if parts:
        return parts[0]
    return cf[:20]

def derive_month(row_src):
    """Derive month from 선적일 (C19 = col 18)"""
    ship_date = row_src[18] if len(row_src) > 18 else None
    if ship_date:
        try:
            if hasattr(ship_date, 'month'):
                m = ship_date.month
                return f'{m}월'
            return ''
        except:
            return ''
    return ''

# Process all data rows
MAX_ROWS = ws_src.max_row
print(f'Processing {MAX_ROWS - 1} data rows...')

for src_row in range(2, MAX_ROWS + 1):
    row_src = [ws_src.cell(src_row, c).value for c in range(1, 55)]
    
    for t_col, header, src_col in template:
        if src_col is not None and src_col > 0 and src_col <= len(row_src):
            val = row_src[src_col - 1]  # 1-indexed to 0-indexed
            ws_new.cell(src_row, t_col, val)
        elif t_col == 3:  # Factory
            contract_factory = row_src[1] if len(row_src) > 1 else ''  # C2
            ws_new.cell(src_row, t_col, get_factory(contract_factory))
        elif t_col == 22:  # 월 (month)
            ws_new.cell(src_row, t_col, derive_month(row_src))
        elif t_col == 30:  # 순매출 = PO금액 - 매출할인
            po_amt = float(row_src[24] or 0) if len(row_src) > 24 and row_src[24] else 0  # C25
            disc = float(row_src[25] or 0) if len(row_src) > 25 and row_src[25] else 0  # C26
            ws_new.cell(src_row, t_col, po_amt - disc)
    
    if (src_row - 1) % 10000 == 0:
        print(f'  ... {src_row - 1} rows done')

# Rename Sheet1
ws_src.title = 'Sheet1'

# Save
wb.save(out_path)
print(f'\nSaved to {out_path}')
print(f'Sheet1: raw data (54 cols)')
print(f'raw_0608: processed (58 cols, with Factory/Division/시즌/월/순매출 added)')

# Verify
wb2 = openpyxl.load_workbook(out_path)
ws_v = wb2['raw_0608']
print(f'\n=== Verification ===')
print(f'Rows: {ws_v.max_row}, Cols: {ws_v.max_column}')
# Print headers
headers_v = {c: ws_v.cell(1, c).value for c in range(1, ws_v.max_column+1) if ws_v.cell(1, c).value}
print(f'Headers (cols with values): {len(headers_v)}')
for c in [3, 11, 13, 22, 30]:
    print(f'  Col {c} ({ws_v.cell(1,c).value}): {ws_v.cell(2,c).value}')
# Check row count
data_rows = ws_v.max_row - 1
print(f'Data rows: {data_rows}')
wb2.close()
