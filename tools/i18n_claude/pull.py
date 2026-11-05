#!/usr/bin/env python3
"""拿原稿：python3 pull.py <食譜id或前8碼> <語言> → 存 WORK/<id8>_<lang>.src.json 與 .brief.txt
不給參數 → 列出待翻清單；加 --fresh → 今天 06:00 後還沒更新過原文就先更新（早上巡檢用）。
2026-11-05 只翻改到的頁：上一版放行過、字沒變的頁不放進 src.json（伺服器投遞時自動沿用），
src.json 的 slides 只剩要翻的頁；同頁碼的舊譯文放在 prev（參考用）；名稱／保存方式沒變就給空字串。"""
import json, sys
from common import *

if len(sys.argv) < 3:
    kw = dict(wkey=key('I18N_PUT_KEY'))
    if '--fresh' in sys.argv: kw['fresh'] = 1
    r = call('i18nPull', **kw)
    d = r.get('data') or {}
    if d.get('refresh'):
        f = d['refresh']
        print('已更新原文：檢查 %s、重抓 %s、要翻的字有變 %s%s' % (f.get('checked'), f.get('fetched'), f.get('changed'), '；' + '；'.join(f['errors']) if f.get('errors') else ''))
    print('待翻', d.get('pending'), '（原文更新於 %s）' % d.get('refresh_at'))
    for x in d.get('items', []):
        print(x['id'], x['lang'], x['why'], '平板顯示上一版' if x.get('showing') else '平板沒有', x['name'])
    sys.exit(0 if r.get('ok') else 1)

idp, lang = sys.argv[1], sys.argv[2]
if len(idp) < 36:   # 前 8 碼 → 從待翻清單或狀態找完整 id
    st = call('i18nStatus')
    ids = [m['id'] for m in st['data']['matrix'] if m['id'].startswith(idp)]
    if len(ids) != 1:
        sys.exit('前綴 %s 對到 %d 支' % (idp, len(ids)))
    idp = ids[0]
r = call('i18nPull', wkey=key('I18N_PUT_KEY'), id=idp, lang=lang)
if not r.get('ok'):
    sys.exit('拿稿失敗：%s' % r.get('msg'))
d = r['data']
t = tag(d['id'], lang)
keep = set(d.get('keep') or [])
total = len(d['slides'])
if keep:
    d['partial'] = True
    d['total'] = total
    d['slides'] = [s for s in d['slides'] if s['k'] not in keep]
    d['prev'] = {str(k): v for k, v in (d.get('prev') or {}).items() if int(k) not in keep}
    if d.get('name_keep'): d['name'] = ''
    if d.get('pres_keep'): d['pres'] = ''
else:
    d['prev'] = {}
json.dump(d, open('%s/%s.src.json' % (WORK, t), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
open('%s/%s.brief.txt' % (WORK, t), 'w', encoding='utf-8').write(d['brief'])
print(t, '全部', total, '頁，要翻', len(d['slides']), '頁（沿用', len(keep), '頁）', '名稱沿用' if d.get('partial') and d.get('name_keep') else '', '已有詞', len(d['dict']), '缺詞', len(d['missing']), '現況', d.get('now'))
