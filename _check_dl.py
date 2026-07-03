import os, datetime, glob
dl = r'C:\Users\user\Downloads'
results = []
for f in glob.glob(dl + r'\*.xlsx') + glob.glob(dl + r'\*.xls'):
    name = os.path.basename(f)
    if not ('직접원가' in name or 'raw' in name.lower() or '영업' in name):
        continue
    ct = datetime.datetime.fromtimestamp(os.path.getctime(f))
    sz = os.path.getsize(f)
    results.append((ct, sz, name))
results.sort(reverse=True)
for ct, sz, name in results[:20]:
    print(ct.strftime('%m-%d %H:%M:%S'), '|', f'{sz:>10,}', '|', name)
