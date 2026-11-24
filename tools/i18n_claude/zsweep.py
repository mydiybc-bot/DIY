#!/usr/bin/env python3
"""總檢（2026-11-25）：python3 zsweep.py [id8 …] → 讀回中文原文食譜各語言「放行中的譯文」與原文，存到 WORK/sweep/<id8>_<lang>.json
不給 id 就讀全部中文原文（on=Y）。只用唯讀預覽通行碼，不改任何資料、不花 API。"""
import sys, json, os, re
from concurrent.futures import ThreadPoolExecutor
from common import *
OUT = os.path.join(WORK, 'sweep'); os.makedirs(OUT, exist_ok=True)
lst = call('i18nZSet', wkey=key('I18N_PUT_KEY'), op='list')['data']['items']
ids = [x['id'] for x in lst if x.get('on')]
if sys.argv[1:]: ids = [i for i in ids if any(i.startswith(p) for p in sys.argv[1:])]
pk = key('I18N_PREVIEW_KEY')
def one(a):
    i, l = a
    j = call('i18nPreview', pkey=pk, id=i, lang=l, timeout=120)
    if not j.get('ok'): return '%s %s 失敗 %s' % (i[:8], l, j.get('msg'))
    d = j['data']
    o = {'id': i, 'lang': l, 'state': d.get('state'), 'name': d.get('name'), 'src_name': d.get('src_name'),
         'pres': d.get('pres'), 'src_pres': d.get('src_pres'),
         'slides': [{'t': s.get('t'), 'h': s.get('h')} for s in (d.get('slides') or [])],
         'src': [{'t': s.get('t'), 'h': s.get('h')} for s in (d.get('src') or [])]}
    json.dump(o, open(os.path.join(OUT, '%s_%s.json' % (i[:8], l)), 'w', encoding='utf-8'), ensure_ascii=False)
    return '%s %s %s %d/%d' % (i[:8], l, o['state'], len(o['slides']), len(o['src']))
with ThreadPoolExecutor(8) as ex:
    res = list(ex.map(one, [(i, l) for i in ids for l in ['en', 'ja', 'ko', 'vi', 'id', 'th']]))
bad = [r for r in res if '失敗' in r or ' live ' not in r]
print('共', len(res), '組；非 live 或失敗', len(bad)); print('\n'.join(bad[:40]))
