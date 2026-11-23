#!/usr/bin/env python3
"""中文原文開關（2026-11-24）：哪幾支中文食譜收進外語食譜（打開才會抓原文、進待翻清單）
python3 zset.py                 → 列候選（台灣區、繁中、啟用中、客人用、非裝飾蛋糕）與開關、對應的舊英文原文
python3 zset.py on <id…>        → 打開（id 可給前 8 碼）
python3 zset.py off <id…>       → 關掉
python3 zset.py old <id> <舊英文原文id|none> → 手動指定對應的舊英文原文
有對應舊英文原文的：中文版 6 種語言都放行之前，平板照舊列舊的那支；都好了才換。"""
import sys
from common import *

args = sys.argv[1:]
op = args[0] if args else 'list'
kw = dict(wkey=key('I18N_PUT_KEY'), op='list')
if op in ('on', 'off', 'old'):
    lst = call('i18nZSet', **kw)['data']['items']
    full = lambda p: [x['id'] for x in lst if x['id'].startswith(p.lower())]
    ids = []
    for p in (args[1:2] if op == 'old' else args[1:]):
        f = full(p)
        if len(f) != 1: sys.exit('前綴 %s 對到 %d 支' % (p, len(f)))
        ids += f
    kw.update(op=op, ids=ids)
    if op == 'old':
        o = args[2] if len(args) > 2 else 'none'
        kw['old'] = '' if o == 'none' else o
r = call('i18nZSet', **kw)
if not r.get('ok'): sys.exit('失敗：%s' % r.get('msg'))
d = r['data']
print('候選 %d 支、打開 %d 支、有舊英文原文 %d 支' % (d['cands'], d['on'], d['paired']))
if op == 'list':
    for x in d['items']:
        if not x['cand']: continue
        print('%s %s %-3s 步%-3s %s%s' % (x['id'][:8], 'ON ' if x['on'] else '-- ', x['pair'] or '', x['steps'], x['name'], ('  ⇐ 舊：' + x['old_name']) if x['old'] else ''))
else:
    for x in d['items']:
        if x['id'] in kw.get('ids', []): print(x['id'], 'ON' if x['on'] else 'off', x['name'], x['old'] or '')
