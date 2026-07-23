import pandas as pd, json, urllib.request, urllib.error
from datetime import datetime, timedelta

EXCEL = 'feishu_revenue_auto.xlsx'
SUPABASE_URL = 'https://tbxxhmtqufzzhshivija.supabase.co'
KEY = 'sb_publishable_vWXg8v_JCLanAlM4nHtMtA_-5qmiJIW'

def supabase_req(method, path, body=None):
    url = f'{SUPABASE_URL}/rest/v1{path}'
    data = json.dumps(body).encode() if body else None
    req = urllib.request.Request(url, data=data, method=method)
    req.add_header('apikey', KEY)
    req.add_header('Authorization', f'Bearer {KEY}')
    req.add_header('Content-Type', 'application/json')
    if method == 'POST':
        req.add_header('Prefer', 'return=representation')
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.status, resp.read().decode()
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode()

# 1. 导出飞书表格
import subprocess, sys, os
result = subprocess.run(
    ['bash', '/c/Users/DZ/.workbuddy/binaries/node/cli-connector-packages/lark-cli', 'sheets', '+workbook-export',
     '--url', 'https://my.feishu.cn/sheets/YqR1sDuMdhq62utI0dvcM5IynJg',
     '--output-path', EXCEL],
    capture_output=True, text=True, cwd='C:/Users/DZ/WorkBuddy/2026-06-29-19-43-45')
print('导出结果:', result.returncode, result.stdout[-100:] if result.stdout else '')

# 2. 读取并转换
df = pd.read_excel(EXCEL, sheet_name=0, header=None)
header = [str(h).strip() for h in df.iloc[0]]
print(f'表头: {header[:4]}...')

data_rows = df.iloc[1:].dropna(subset=[0])
print(f'数据行数: {len(data_rows)}')

import_data = []
skipped = 0

for _, row in data_rows.iterrows():
    date_val = row[0]
    if pd.isna(date_val): continue
    if isinstance(date_val, (int, float)):
        d = datetime(1899, 12, 30) + timedelta(days=int(date_val))
        date_str = d.strftime('%Y-%m-%d')
    else:
        date_str = str(date_val)[:10]

    coin0 = float(row[1]) if pd.notna(row[1]) else 0
    first0 = float(row[2]) if pd.notna(row[2]) else 0
    renew0 = float(row[3]) if pd.notna(row[3]) else 0

    # 当日
    total0 = coin0 + first0 + renew0
    if total0 > 0:
        import_data.append({'date': date_str, 'installDays': 0,
                           'coin': coin0, 'firstSub': first0, 'renewSub': renew0,
                           'coinRenew': 0, 'total': total0})
    else:
        skipped += 1

    # 8/14/30/45日
    for days, rCol, crCol in [(8,4,8),(14,5,9),(30,6,10),(45,7,11)]:
        r = float(row[rCol]) if pd.notna(row[rCol]) else 0
        cr = float(row[crCol]) if pd.notna(row[crCol]) else 0
        ttl = r + cr
        if ttl > 0:
            import_data.append({'date': date_str, 'installDays': days,
                               'coin': 0, 'firstSub': 0,
                               'renewSub': r, 'coinRenew': cr,
                               'total': ttl})
        else:
            skipped += 1

print(f'转换完成: {len(import_data)} 条 (跳过全0行: {skipped})')

# 3. 分批删除旧数据 + 插入新数据
dates = sorted(set(r['date'] for r in import_data))
print(f'涉及日期: {len(dates)} 个')

BATCH = 50

# 先一次性清空所有涉及日期的旧数据
# 避免按 50 条切批时，同一日期的 D0 与 N日被切到相邻批次，
# 后批次的 DELETE 把前批次刚 INSERT 的数据抹掉
all_dates = sorted(set(r['date'] for r in import_data))
print(f'清空旧数据，涉及 {len(all_dates)} 个日期...')
for d in all_dates:
    status, _ = supabase_req('DELETE', f'/revenues?date=eq.{d}')
    if status != 204:
        print(f'  删除 {d} 失败: {status}')

for i in range(0, len(import_data), BATCH):
    batch = import_data[i:i+BATCH]

    # 插入新数据
    payload = [{'date': r['date'], 'install_days': r['installDays'],
                'coin': r['coin'], 'first_sub': r['firstSub'],
                'renew_sub': r['renewSub'], 'coin_renew': r['coinRenew'],
                'total': r['total']} for r in batch]
    status, resp = supabase_req('POST', '/revenues', payload)
    if status == 201:
        print(f'  批次 {i//BATCH+1}: 插入 {len(payload)} 条 OK')
    else:
        print(f'  批次 {i//BATCH+1} 失败: {status} {resp[:200]}')
        break
else:
    print(f'\n同步完成！共 {len(dates)} 天，{len(import_data)} 条记录')
