#!/usr/bin/env python3
"""列出下一批要翻的食譜（還有待翻語言的）：python3 nextwave.py [N]  → 每行：前8碼 待翻語言 名稱"""
import sys
from common import *
r = call('i18nPull', wkey=key('I18N_PUT_KEY'))
items = r['data']['items']; order = []; by = {}
for x in items:
    if x['id'] not in by: by[x['id']] = []; order.append(x['id'])
    by[x['id']].append(x['lang'])
n = int(sys.argv[1]) if len(sys.argv) > 1 else 3
print('待翻', r['data']['pending'], '組；涉及食譜', len(order), '支')
for i in order[:n]:
    nm = next(x['name'] for x in items if x['id'] == i)
    print(i[:8], ','.join(by[i]), nm)
