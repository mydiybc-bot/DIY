#!/usr/bin/env python3
"""總檢修正（2026-11-25）：python3 zfix.py <lang> [--dry] [--skip id8,id8]
讀 sweep/_found_<lang>.json → 改 WORK/<id8>_<lang>.tr.json（先備份到 sweep/backup/）→ 預檢 → put --approve → 讀回確認新字已上線。"""
import json, os, sys, shutil, subprocess, collections
from common import *
lang = sys.argv[1]; dry = '--dry' in sys.argv
skip = set(sys.argv[sys.argv.index('--skip') + 1].split(',')) if '--skip' in sys.argv else set()
D = os.path.join(WORK, 'sweep'); B = os.path.join(D, 'backup'); os.makedirs(B, exist_ok=True)
here = os.path.dirname(os.path.abspath(__file__))
found = json.load(open(os.path.join(D, '_found_%s.json' % lang), encoding='utf-8'))
by = collections.defaultdict(list)
for f in found: by[f['id8']].append(f)
res = []
for i8, L in by.items():
    if i8 in skip: res.append((i8, '跳過（另有人在做）')); continue
    tag = '%s_%s' % (i8, lang); p = os.path.join(WORK, tag + '.tr.json')
    tr = json.load(open(p, encoding='utf-8')); sl = {s['k']: s for s in tr['slides']}
    ok, miss = [], []
    for f in L:
        fld = f.get('field') or ('pres' if f['k'] == 'pres' else 'h')
        if f['k'] == 'pres' or fld == 'pres': obj, key_ = tr, 'pres'
        else:
            obj = sl.get(int(f['k'])); key_ = fld
            if obj is None: miss.append((f['k'], '沒有這頁')); continue
        cur = obj.get(key_) or ''
        if cur.count(f['old']) != 1: miss.append((f['k'], 'old 出現 %d 次' % cur.count(f['old']))); continue
        obj[key_] = cur.replace(f['old'], f['new']); ok.append(f)
    if not ok: res.append((i8, '沒改到：%s' % miss)); continue
    if dry: res.append((i8, '試算：可改 %d、不行 %s' % (len(ok), miss))); continue
    shutil.copy(p, os.path.join(B, tag + '.tr.json'))
    json.dump(tr, open(p, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    src = json.load(open(os.path.join(WORK, tag + '.src.json'), encoding='utf-8'))
    j = call('i18nPut', wkey=key('I18N_PUT_KEY'), id=src['id'], lang=lang, fp=src['fp'], name=tr.get('name', ''), pres=tr.get('pres', ''),
             slides=tr['slides'], partial=1, dict=[], approve=1, overwrite=0, model='claude-code-sweep', by='Claude Code 總檢')
    st = j.get('data') or {}
    res.append((i8, '改 %d 處｜ok=%s %s｜%s｜沒改 %s' % (len(ok), j.get('ok'), json.dumps({k: st.get(k) for k in ('status', 'approved', 'state')}, ensure_ascii=False), (st.get('check') or j.get('msg') or '')[:150], miss)))
for r in res: print(*r)
