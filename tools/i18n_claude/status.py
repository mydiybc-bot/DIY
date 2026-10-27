"""平板外語食譜：看各語言 live／review／fail 數與待檢查清單（2026-10-27）
用法：python3 <本資料夾>/status.py          → 只看狀態
      python3 <本資料夾>/status.py --clean  → 另外清掉暫存資料夾裡、該組已 live 的 src／tr／review／brief 檔（opus_backup 不動）
取代原本的 python3 -c "from common import *; ..."，讓排程巡檢可以免按允許。"""
import os, sys, glob
from collections import Counter
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from common import call, tag, WORK

j = call('i18nStatus')
if not j.get('ok', True) and 'data' not in j:
    raise SystemExit('讀不到 i18nStatus：%s' % j.get('msg'))
s = j['data']
langs = s['langs']
cnt = {l: Counter() for l in langs}
st = {}
for r in s['matrix']:
    for l in langs:
        v = r['st'].get(l, 'none')
        cnt[l][v] += 1
        st[tag(r['id'], l)] = v
print('食譜 %d 支' % len(s['matrix']))
for l in langs:
    print('  %s：%s' % (l, '、'.join('%s %d' % kv for kv in sorted(cnt[l].items()))))
rv = s.get('review') or []
print('待檢查 %d 組' % len(rv))
for x in rv:
    print('  ', x)

if '--clean' in sys.argv:
    gone = 0
    for p in glob.glob(os.path.join(WORK, '*')):
        b = os.path.basename(p)
        if not os.path.isfile(p) or not b.endswith(('.src.json', '.tr.json', '.review.json', '.brief.txt')):
            continue
        if st.get(b.split('.')[0]) == 'live':
            os.remove(p); gone += 1
    print('清掉已上線的暫存檔 %d 個' % gone)
