#!/usr/bin/env python3
"""總檢：把 _check.json 依語言切成給核對員看的小檔 → sweep/_review_<lang>.json（只留有問題那一條步驟的中文與譯文）"""
import json, os, re, collections
from common import WORK
D = os.path.join(WORK, 'sweep'); C = json.load(open(os.path.join(D, '_check.json'), encoding='utf-8'))
PAT = {'A尺寸': r'.', 'B剪花嘴': r'剪.{0,6}花嘴|花嘴的?\s*1/4', 'B剪掉1/4': r'1/4', 'B保持乾燥': r'乾燥', 'B朝上朝下': r'朝上|朝下', 'B無借冰': r'借冰|保冰',
       'B小心燙': r'燙', 'B冷藏冷凍': r'冷藏|冷凍', 'B表面抹平': r'抹平', 'B擀': r'擀', 'B重置': r'重置', 'B套入': r'套入|套進'}
def items(s): return [x.strip() for x in re.split(r'(?=(?:^|\s)\d+\.\s)|\n|●', s) if x.strip()]
def num(x): m = re.match(r'(\d+)\.', x); return m.group(1) if m else None
out = collections.defaultdict(list)
for kind, L in (('auto', C['auto']), ('manual', C['manual'])):
    for f in L:
        z, t = f['zh'], f['tr']; pat = PAT.get(f['rule'], r'.')
        zi = [x for x in items(z) if re.search(pat, x)] if f['rule'] != 'A尺寸' else [z]
        ns = {num(x) for x in zi if num(x)}
        ti = [x for x in items(t) if num(x) in ns] if ns and f['k'] != 'pres' else [t]
        out[f['lang']].append({'id8': f['id8'], 'k': f['k'], 'rule': f.get('why') or f['rule'], 'extra': f.get('extra'), 'zh': ' / '.join(zi)[:400], 'tr': ' / '.join(ti)[:600]})
for l, L in out.items():
    json.dump(L, open(os.path.join(D, '_review_%s.json' % l), 'w', encoding='utf-8'), ensure_ascii=False, indent=0)
    print(l, len(L), os.path.getsize(os.path.join(D, '_review_%s.json' % l)) // 1024, 'KB')
