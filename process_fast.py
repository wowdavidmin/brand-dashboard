"""Fast processing of new raw data using pandas"""
import pandas as pd
import numpy as np
from openpyxl import load_workbook
from copy import copy

src_path = r'C:\Users\user\Downloads\브랜드별 직접원가 (7).xlsx'
out_path = r'C:\Users\user\.openclaw\workspace\영업데이터파일_raw_20260702.xlsx'

print('Reading raw data...')
df = pd.read_excel(src_path, header=0, engine='openpyxl')
print(f'Read {len(df)} rows, {len(df.columns)} cols')

# Print column names
for i, c in enumerate(df.columns, 1):
    print(f'  C{i}: {repr(c)}')

# Load workbook to add sheet
wb = load_workbook(src_path)

# Create new sheet
if 'raw_0608' in wb.sheetnames:
    del wb['raw_0608']
ws_new = wb.create_sheet('raw_0608')

# Rename Sheet1
ws_src = wb.worksheets[0]

# Map source column index (0-based) to target column index (1-based)
# Target = 58 columns matching Raw_20260608 format
# Source columns:
col_map = {}

# Let me build the mapping from actual column names
src_cols = list(df.columns)

# Helper to find column index by name pattern
def find_col(patterns, src_cols):
    for i, col in enumerate(src_cols):
        col_str = str(col).strip().lower()
        for p in patterns:
            if p.lower() in col_str:
                return i  # 0-based
    return None

# Map source to target
idx_NO = find_col(['NO'], src_cols)
idx_Contract = find_col(['Contract'], src_cols)
idx_Factory_data = None  # Not in source - will derive
idx_바이어 = find_col(['바이어'], src_cols)
idx_브랜드 = find_col(['브랜드'], src_cols)  # C4 in source
idx_부서 = find_col(['부서'], src_cols)  # C5
idx_바이어스타일 = find_col(['바이어스타일'], src_cols)  # C6
idx_스타일 = find_col(['스타일'], src_cols)  # C7
idx_원가번호 = find_col(['원가번호'], src_cols)  # C8
idx_Grouped = find_col(['Grouped'], src_cols)  # C9
idx_디비젼 = find_col(['디비젼', 'Division', '디비전'], src_cols)  # C10
idx_시즌 = find_col(['시즌'], src_cols)  # C11
idx_ITEM1 = find_col(['ITEM명'], src_cols)  # C12
# C13, C14 in source are also ITEM명 (duplicate column names)

# For ITEM명 cols - pandas may have renamed duplicates
item_cols = [i for i, c in enumerate(src_cols) if 'ITEM' in str(c).upper() or '아이템' in str(c)]
print(f'ITEM명 columns (0-based): {item_cols}')
idx_ITEM1 = item_cols[0] if len(item_cols) > 0 else None
idx_ITEM2 = item_cols[1] if len(item_cols) > 1 else None
idx_ITEM3 = item_cols[2] if len(item_cols) > 2 else None

idx_담당 = find_col(['담당'], src_cols)  # C15
idx_진행구분 = find_col(['진행구분'], src_cols)  # C16
idx_오더 = find_col(['오더'], src_cols)  # C17
idx_PO_Status = find_col(['PO Status', 'Status'], src_cols)  # C18
idx_선적일 = find_col(['선적일'], src_cols)  # C19
idx_exFact = find_col(['ex-Fact', 'ex-Factory'], src_cols)  # C20
idx_DC_date = find_col(['DC'], src_cols)  # C21
idx_SHIP_MODE = find_col(['SHIP MODE', 'SHIP'], src_cols)  # C22
idx_TERMS = find_col(['TERMS'], src_cols)  # C23
idx_PO수량 = find_col(['PO수량'], src_cols)  # C24
idx_PO금액 = find_col(['PO금액'], src_cols)  # C25
idx_매출할인 = find_col(['매출할인'], src_cols)  # C26

# Price/cost columns
idx_원자재원사 = find_col(['원자재비'], src_cols)  # C27 - may match first
idx_원자재원단 = find_col(['원자재비'], src_cols)  # C28
idx_부자재 = find_col(['부자재비'], src_cols)  # C29
idx_염가공 = find_col(['염가공'], src_cols)  # C30
idx_IF염가공 = find_col(['I.F염가공', 'IF염가공'], src_cols)  # C31
idx_염가공계 = find_col(['염가공비 계', '염가공비'], src_cols)  # C32
idx_임가공외화 = find_col(['임가공료'], src_cols)  # C33
idx_임가공원화 = find_col(['임가공료'], src_cols)  # C34
idx_임가공합계 = find_col(['임가공료'], src_cols)  # C35
idx_CMT봉제료 = find_col(['CMT'], src_cols)  # C36
idx_CMTIF = find_col(['CMT'], src_cols)  # C37
idx_CMT부자재 = find_col(['CMT'], src_cols)  # C38
idx_CMT임가공 = find_col(['CMT'], src_cols)  # C39
idx_CMT기타 = find_col(['CMT'], src_cols)  # C40
idx_CMT계 = find_col(['CMT 계'], src_cols)  # C41
idx_해상운임 = find_col(['해상운임'], src_cols)  # C42
idx_Inspection = find_col(['Inspection'], src_cols)  # C43
idx_NCPM = find_col(['NCPM'], src_cols)  # C44
idx_직접원가AMOUNT = find_col(['직접원가'], src_cols)  # C45
idx_직접원가율 = find_col(['직접원가 %'], src_cols)  # C46
idx_PROFIT = find_col(['PROFIT'], src_cols)  # C47
idx_PROFIT율 = find_col(['PROFIT %'], src_cols)  # C48
idx_손익 = find_col(['손익'], src_cols)  # C49
idx_BuyerComm = find_col(['Buyer'], src_cols)  # C50
idx_AgentComm = find_col(['Agent'], src_cols)  # C51
idx_Duty = find_col(['Duty'], src_cols)  # C52
idx_ExecMargin = find_col(['Execute'], src_cols)  # C53
idx_ExecMarginPct = find_col(['Execute'], src_cols)  # C54 - might be same column

print(f'\nKey column mapping:')
print(f'  NO: {idx_NO}')
print(f'  바이어: {idx_바이어}')
print(f'  부서: {idx_부서}')
print(f'  디비젼: {idx_디비젼}')
print(f'  시즌: {idx_시즌}')
print(f'  PO수량: {idx_PO수량}')
print(f'  PO금액: {idx_PO금액}')
print(f'  매출할인: {idx_매출할인}')
print(f'  ITEM명 cols: {item_cols}')
print(f'  CMT 계: {idx_CMT계}')
