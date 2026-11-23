#!/usr/bin/env python3
"""中文詞典（第 2 段，2026-11-24）：
python3 zdict.py terms            → 掃全部中文候選食譜用到的詞（品項／取放區／單位）＋提示，存 WORK/zdict/terms.json（分好幾次叫，一次最多 4 分鐘）
python3 zdict.py put <rows.json> [--overwrite] → 整批寫進中文詞典 i18n_zitems（rows＝[{kind, zh, en, ja, ko, vi, id, th}]）"""
import json, os, sys
from common import *
D = os.path.join(WORK, 'zdict'); os.makedirs(D, exist_ok=True)
op = sys.argv[1] if len(sys.argv) > 1 else 'terms'
if op == 'terms':
    frm, items, seen, errs = 0, [], {}, []
    while True:
        r = call('i18nZDict', wkey=key('I18N_PUT_KEY'), op='terms', **{'from': frm}, timeout=330)
        if not r.get('ok'): sys.exit('失敗：%s' % r.get('msg'))
        d = r['data']; errs += d['errors']
        for x in d['items']:
            k = x['kind'] + '|' + x['zh'].lower()
            if k in seen: seen[k]['n'] += x['n']
            else: seen[k] = x; items.append(x)
        print('掃到 %d/%d 支，詞 %d 個（%d 秒）' % (d['done'], d['cands'], len(items), d['sec']))
        if d['left'] <= 0 or d['done'] <= frm: break
        frm = d['done']
    json.dump({'items': items, 'errors': errs}, open(os.path.join(D, 'terms.json'), 'w'), ensure_ascii=False, indent=1)
    print('存好', os.path.join(D, 'terms.json'), '錯誤', errs)
elif op == 'put':
    rows = json.load(open(sys.argv[2]))
    r = call('i18nZDict', wkey=key('I18N_PUT_KEY'), op='put', rows=rows, overwrite=1 if '--overwrite' in sys.argv else 0, timeout=330)
    print(json.dumps(r.get('data') or r, ensure_ascii=False))
