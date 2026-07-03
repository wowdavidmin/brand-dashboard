import os, datetime, glob

dl = r'C:\Users\user\Downloads'
for f in sorted(glob.glob(dl + r'\*'), key=os.path.getctime, reverse=True):
    name = os.path.basename(f)
    if not (name.endswith('.xlsx') or name.endswith('.xls') or name.endswith('.csv')):
        continue
    ct = datetime.datetime.fromtimestamp(os.path.getctime(f))
    if (datetime.datetime.now() - ct).total_seconds() > 3600:
        continue
    sz = os.path.getsize(f)
    print(f"{ct.strftime('%H:%M:%S')} | {sz:>8,} | {name}")
