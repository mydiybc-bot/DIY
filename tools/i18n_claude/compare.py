#!/usr/bin/env python3
"""Sonnet 譯文 vs Opus 昨晚譯文並排：python3 compare.py <tag> → 印出逐頁差異摘要（只列文字不同的頁，截短）"""
import json, re, sys, difflib
from common import *
t = sys.argv[1]
son = json.load(open('%s/%s.tr.json' % (WORK, t), encoding='utf-8'))
opu = json.load(open('%s/opus_backup/%s.preview.json' % (WORK, t), encoding='utf-8'))
src = json.load(open('%s/%s.src.json' % (WORK, t), encoding='utf-8'))
strip = lambda h: re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', h or '')).strip()
same = diff = 0
rows = []
for i, s in enumerate(son['slides']):
    a, b = strip(s['h']), strip(opu['slides'][i]['h'])
    if a == b: same += 1
    else:
        diff += 1
        r = difflib.SequenceMatcher(None, a, b).ratio()
        rows.append((i + 1, round(r, 2), a[:90], b[:90]))
print(t, '頁數', len(son['slides']), '完全相同', same, '不同', diff)
print('名稱 Sonnet:', son['name'], '| Opus:', opu['name'])
for p, r, a, b in sorted(rows, key=lambda x: x[1])[:6]:
    print(' 第%d頁 相似度%.2f\n   S: %s\n   O: %s' % (p, r, a, b))
