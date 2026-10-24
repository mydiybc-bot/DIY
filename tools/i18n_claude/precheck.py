#!/usr/bin/env python3
"""投遞前的本機預檢（跟伺服器同一套規則的簡化版）：python3 precheck.py <id8_lang>
檢查：頁數、k 連續、[[IMGn]] 位置、每頁數字、HTML 標籤只留允許的、有沒有留下整句英文。"""
import json, re, sys
from common import *
t = sys.argv[1]
src = json.load(open('%s/%s.src.json' % (WORK, t), encoding='utf-8'))
tr = json.load(open('%s/%s.tr.json' % (WORK, t), encoding='utf-8'))
bad = []
if len(tr['slides']) != len(src['slides']): bad.append('頁數 %d≠%d' % (len(tr['slides']), len(src['slides'])))
ph = lambda h: sorted(re.findall(r'\[\[IMG\d+\]\]', h or ''))
nums = lambda h: sorted(re.findall(r'\d+(?:[.,]\d+)?', re.sub(r'<[^>]+>', ' ', re.sub(r'\[\[IMG\d+\]\]', ' ', h or ''))))
ALLOWED = {'p', 'br', 'b', 'u', 'i', 'font'}
for s, x in zip(src['slides'], tr['slides']):
    k = s['k']
    if x.get('k') != k: bad.append('第%d頁 k 不對' % (k + 1))
    if ph(s['h']) != ph(x['h']): bad.append('第%d頁 圖片位置' % (k + 1))
    a, b = nums(s['h']), nums(x['h'])
    miss = [n for n in a if n not in b or a.count(n) > b.count(n)]
    if miss: bad.append('第%d頁 數字少了 %s' % (k + 1, sorted(set(miss))))
    for tg in re.findall(r'</?\s*([a-zA-Z0-9]+)', x['h']):
        if tg.lower() not in ALLOWED: bad.append('第%d頁 多了標籤 <%s>' % (k + 1, tg)); break
    if s['t'] and not x.get('t', '').strip(): bad.append('第%d頁 標題空' % (k + 1))
    for tg in ('p', 'b', 'u', 'font'):   # 開關標籤數量要跟原稿一致（缺 </p> 等會讓畫面壞掉）
        co = lambda h, pat: len(re.findall(pat, h or '', re.I))
        so, sc = co(s['h'], r'<%s[\s>]' % tg), co(s['h'], r'</%s>' % tg)
        to, tc = co(x['h'], r'<%s[\s>]' % tg), co(x['h'], r'</%s>' % tg)
        if to != tc: bad.append('第%d頁 <%s> 開關不成對（開%d關%d）' % (k + 1, tg, to, tc))
if src['pres'] and not tr.get('pres', '').strip(): bad.append('保存方式空')
if not tr.get('name', '').strip(): bad.append('名稱空')
miss_terms = [m for m in src['missing'] if not any(d['kind'] == m['kind'] and d['en'].lower() == m['en'].lower() for d in tr.get('dict', []))]
if miss_terms: bad.append('詞典還缺 %d 個：%s' % (len(miss_terms), ', '.join(m['en'] for m in miss_terms[:8])))
print(t, '預檢', '全過' if not bad else '有 %d 個問題' % len(bad))
for b in bad: print(' -', b)
sys.exit(1 if bad else 0)
