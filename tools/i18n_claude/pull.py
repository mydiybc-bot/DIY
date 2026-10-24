#!/usr/bin/env python3
"""拿原稿：python3 pull.py <食譜id或前8碼> <語言> → 存 WORK/<id8>_<lang>.src.json 與 .brief.txt
不給參數 → 列出待翻清單。"""
import json, sys
from common import *

if len(sys.argv) < 3:
    r = call('i18nPull', wkey=key('I18N_PUT_KEY'))
    d = r.get('data') or {}
    print('待翻', d.get('pending'))
    for x in d.get('items', []):
        print(x['id'], x['lang'], x['why'], x['name'])
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
json.dump(d, open('%s/%s.src.json' % (WORK, t), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
open('%s/%s.brief.txt' % (WORK, t), 'w', encoding='utf-8').write(d['brief'])
print(t, d['name'], '頁數', len(d['slides']), '已有詞', len(d['dict']), '缺詞', len(d['missing']), '現況', d.get('now'))
