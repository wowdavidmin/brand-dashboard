import os, datetime, glob

locations = [
    r'C:\Users\user\.openclaw\media\inbound',
    r'C:\Users\user\Downloads',
]

seen = set()
for loc in locations:
    if not os.path.isdir(loc):
        continue
    for f in os.listdir(loc):
        fp = os.path.join(loc, f)
        if not os.path.isfile(fp):
            continue
        if fp in seen:
            continue
        seen.add(fp)
        ext = os.path.splitext(f)[1].lower()
        if ext not in ('.xlsx', '.xls', '.csv'):
            continue
        sz = os.path.getsize(fp)
        ct = datetime.datetime.fromtimestamp(os.path.getctime(fp))
        mt = datetime.datetime.fromtimestamp(os.path.getmtime(fp))
        print(f"{ct.strftime('%m-%d %H:%M:%S')} | {mt.strftime('%m-%d %H:%M:%S')} | {sz:>10,} | {f}")
