#!/usr/bin/env python3
"""總檢標記（2026-11-25）：讀 WORK/sweep/*.json，列出回讀員常漏看的地方 → WORK/sweep/_flags.json＋摘要
A 自動判：①「撕下再貼」沒寫撕下 ②譯文多出原文沒有的尺寸（5''、6吋、No.7…）
B 列給人工看：剪花嘴／剪掉1/4、保持乾燥、無借冰、冷藏冷凍、小心燙、擀、朝上朝下、套入、表面抹平、重置"""
import json, os, re, glob, collections
from common import WORK
D = os.path.join(WORK, 'sweep')
txt = lambda h: re.sub(r'\s+', ' ', re.sub(r'<[^>]+>', ' ', h or '')).strip()
PEEL = {'en': r'peel|remove|take (it )?off|detach|unstick|pull off', 'ja': r'はが|剥が|はず|外し|取り', 'ko': r'떼|뗀|벗기', 'vi': r'gỡ|bóc|lột|tháo', 'id': r'lepas|copot|cabut|kelupas|ambil', 'th': r'ลอก|แกะ|ดึง'}
UNIT = r"(\d+)\s*(?:''|\"|”|″|吋|インチ|인치|นิ้ว|inch(?:es)?\b|号|호|เบอร์|No\.?\s*)" 
UNIT2 = r"(?:No\.?\s*|số\s*|เบอร์\s*|ขนาด\s*)(\d+)"
B = [('剪花嘴', r'剪.{0,6}花嘴|花嘴的?\s*1/4'), ('剪掉1/4', r'剪掉\s*1/4|剪.{0,4}1/4'), ('保持乾燥', r'乾燥'), ('無借冰', r'借冰'),
     ('冷藏冷凍', r'冷藏|冷凍'), ('小心燙', r'燙'), ('擀', r'擀'), ('朝上朝下', r'朝上|朝下'), ('套入', r'套入|套進'), ('表面抹平', r'表面.{0,3}抹平|抹平.{0,3}表面'), ('重置', r'重置')]
flags, cnt = [], collections.Counter()
for f in sorted(glob.glob(os.path.join(D, '*_*.json'))):
    if os.path.basename(f).startswith('_'): continue
    d = json.load(open(f, encoding='utf-8')); l = d['lang']; i8 = d['id'][:8]
    for k, (s, z) in enumerate(zip(d['slides'], d['src'])):
        zt, tt = txt(z['t']) + '｜' + txt(z['h']), txt(s['t']) + '｜' + txt(s['h'])
        if re.search(r'貼至|貼在|貼到', zt) and re.search(r'便條紙|烤盤紙|紙', zt) and not re.search(PEEL[l], tt, re.I):
            flags.append({'id8': i8, 'lang': l, 'k': k, 'rule': 'A撕下', 'zh': zt[:300], 'tr': tt[:400]}); cnt['A撕下'] += 1
        zn = set(re.findall(r'\d+', zt))
        extra = sorted({m for m in re.findall(UNIT, tt) + re.findall(UNIT2, tt) if m not in zn})
        if extra:
            flags.append({'id8': i8, 'lang': l, 'k': k, 'rule': 'A尺寸', 'extra': extra, 'zh': zt[:300], 'tr': tt[:400]}); cnt['A尺寸'] += 1
        for nm, pat in B:
            if re.search(pat, zt):
                flags.append({'id8': i8, 'lang': l, 'k': k, 'rule': 'B' + nm, 'zh': zt[:300], 'tr': tt[:400]}); cnt['B' + nm] += 1
    for nm, pat in B:   # 到店須知／保存說明
        if re.search(pat, txt(d.get('src_pres'))):
            flags.append({'id8': i8, 'lang': l, 'k': 'pres', 'rule': 'B' + nm, 'zh': txt(d.get('src_pres'))[:500], 'tr': txt(d.get('pres'))[:600]}); cnt['B' + nm + '(pres)'] += 1
json.dump(flags, open(os.path.join(D, '_flags.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('檔數', len(glob.glob(os.path.join(D, '*_*.json'))) - 1, '標記', len(flags)); print(dict(cnt))
