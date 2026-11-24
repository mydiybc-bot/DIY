#!/usr/bin/env python3
"""第 4 段一輪（2026-11-24）：python3 zround.py <id8> <id8> …
打開這幾支中文原文 → 強制更新原文 → 並行拿 6 語言稿（已放行、沒改到的組 pull 會給空的 slides）"""
import sys, json, subprocess, os
from concurrent.futures import ThreadPoolExecutor
from common import *
ids = sys.argv[1:]
lst = call('i18nZSet', wkey=key('I18N_PUT_KEY'), op='list')['data']['items']
full = []
for p in ids:
    f = [x['id'] for x in lst if x['id'].startswith(p)]
    if len(f) != 1: sys.exit('前綴 %s 對到 %d 支' % (p, len(f)))
    full += f
r = call('i18nZSet', wkey=key('I18N_PUT_KEY'), op='on', ids=full); print('打開', r['data']['on'], '支')
r = call('i18nPull', wkey=key('I18N_PUT_KEY'), fresh='force', timeout=330); print('原文', json.dumps((r.get('data') or {}).get('refresh'), ensure_ascii=False))
here = os.path.dirname(os.path.abspath(__file__))
def one(a):
    o = subprocess.run(['python3', os.path.join(here, 'pull.py'), a[0], a[1]], capture_output=True, text=True)
    return (o.stdout.strip().splitlines() or [o.stderr.strip()])[-1]
with ThreadPoolExecutor(8) as ex:
    for line in ex.map(one, [(i, l) for i in full for l in ['en', 'ja', 'ko', 'vi', 'id', 'th']]):
        w = line.split()
        print(' '.join(w[:1] + w[2:6]))
