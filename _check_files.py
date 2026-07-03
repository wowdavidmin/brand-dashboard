import os, datetime
path = r'C:\Users\user\.openclaw\media\inbound'
files = []
for f in os.listdir(path):
    fp = os.path.join(path, f)
    if os.path.getsize(fp) > 10000:
        ct = datetime.datetime.fromtimestamp(os.path.getctime(fp))
        mt = datetime.datetime.fromtimestamp(os.path.getmtime(fp))
        files.append((ct, mt, os.path.getsize(fp), f))
files.sort(reverse=True)
for ct, mt, sz, f in files[:10]:
    print(f"{ct.strftime('%m-%d %H:%M:%S')} | {mt.strftime('%m-%d %H:%M:%S')} | {sz:>10,} | {f}")
