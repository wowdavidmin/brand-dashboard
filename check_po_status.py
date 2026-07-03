"""Quick check of 진행구분 and PO Status values"""
import pandas as pd

f = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---47ea1d42-8317-4a51-a9fa-d5c2e2764f59.xlsx'

# Read only first 100 rows to get unique values
df = pd.read_excel(f, usecols=[15, 16, 17, 23], header=0, nrows=100, engine='openpyxl')
print(df.columns.tolist())
# Actual headers from row 1: C16=진행구분, C17=오더, C18=PO Status, C24=PO수량
print(df.head())
print()
print('Unique values in C16 (진행구분):', df.iloc[:,0].unique() if len(df.columns)>0 else 'N/A')
