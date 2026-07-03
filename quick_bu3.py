import pandas as pd, time

raw_path = r'C:\Users\user\.openclaw\media\inbound\브랜드별_직접원가_raw_v1.0_20260629---47ea1d42-8317-4a51-a9fa-d5c2e2764f59.xlsx'

t0 = time.time()
df = pd.read_excel(raw_path, usecols=[2, 4, 10, 23, 24, 25, 40], header=0, engine='openpyxl')
print(f'Read {len(df)} rows in {time.time()-t0:.1f}s')

df.columns = ['buyer', 'dept', 'season', 'qty', 'amt', 'disc', 'cmt']
df_26 = df[df['season'].astype(str).str.startswith('26 ', na=False)]

bu3 = df_26[df_26['dept'].isin(['BU3_D1','BU3_D2','BU3_D3','BU3_D4'])]
print(f'BU3 rows: {len(bu3)}\n')

# By department + brand
agg2 = bu3.groupby(['dept','buyer','season']).agg({'qty':'sum','amt':'sum','disc':'sum','cmt':'sum'}).reset_index()

for dept in ['BU3_D1','BU3_D2','BU3_D3','BU3_D4']:
    print(f'=== {dept} ===')
    sub = agg2[agg2['dept']==dept]
    for _, row in sub.iterrows():
        net = row['amt'] - row['disc']
        avg = row['amt']/row['qty'] if row['qty']>0 else 0
        cm_u = row['cmt']/row['qty'] if row['qty']>0 else 0
        print(f'  {row["buyer"]:30s} {row["season"]:8s} qty={row["qty"]:>10.0f} amt={row["amt"]/1000:>8.1f}K net={net/1000:>8.1f}K avg={avg:.2f} cm={cm_u:.3f}')
    print()
