import openpyxl

# Check the raw data files
path1 = r'C:\Users\user\.openclaw\media\inbound\20212026_영업_DATA_FILE_working_v1_2_20260609_openclaw_작업용---5aca5cd0-82de-4f6d-a748-02126a5b502a.xlsx'
path2 = r'C:\Users\user\.openclaw\media\inbound\20212026_영업_DATA_FILE_working_v1_2_20260609_openclaw_작업용---aa7ac6b1-5682-42c2-b6ac-40b3c069f1da.xlsx'

for i, p in enumerate([path1, path2], 1):
    try:
        wb = openpyxl.load_workbook(p)
        ws = wb.worksheets[0]
        print(f'File {i} - Sheets: {wb.sheetnames}')
        print(f'  Rows: {ws.max_row}, Cols: {ws.max_column}')
        # Print first 20 rows
        for r in range(1, min(25, ws.max_row+1)):
            row_vals = []
            for c in range(1, min(15, ws.max_column+1)):
                v = ws.cell(r, c).value
                if v is not None:
                    row_vals.append(f'C{c}={v}')
            if row_vals:
                print(f'  Row {r}: {", ".join(row_vals)}')
        print()
    except Exception as e:
        print(f'File {i} Error: {e}')
