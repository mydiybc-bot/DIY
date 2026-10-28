#!/usr/bin/env python3
"""核對：convert.py 輸出的 rows_<版本>.jsonl（獨立重算）vs 線上 stats2 彙總，逐店逐題逐選項比對。
  python3 -I tools/survey-import/verify_agg.py <convert 輸出資料夾> <版本> [版本 ...]"""
import sys, os, json, subprocess, collections
URL = 'https://script.google.com/macros/s/AKfycbxrsYaYLdXqjUeGvG8_CZ--5znw5izXr6tcgmfC0MbJ32T1KY_TS8UkVqimxS2Jajx0/exec?action=stats2&from=2015-01-01&to=2099-12-31'
out, labs = os.path.expanduser(sys.argv[1]), sys.argv[2:]
exp = {}
for lab in labs:
    for l in open(os.path.join(out, 'rows_%s.jsonl' % lab), encoding='utf-8'):
        r = json.loads(l); g = exp.setdefault(r['sn'], {'n': 0, 'secSum': 0, 'secN': 0, 'c': collections.Counter()})
        g['n'] += 1
        if 0 < r['sec'] < 3600: g['secSum'] += r['sec']; g['secN'] += 1
        for q, v in r['a'].items():
            g['c']['#' + q] += 1
            if q == 'q17': continue
            for o in v.split('|'): g['c'][q + '|' + o] += 1
d = json.loads(subprocess.run(['curl', '-sL', '--max-time', '120', URL], capture_output=True, text=True).stdout)
live = {g['sn']: g for g in d['stores']}
bad = 0
for sn in sorted(set(exp) | set(live)):
    e, l = exp.get(sn), live.get(sn)
    if not e or not l: print('店', sn, '只有一邊有', bool(e), bool(l)); bad += 1; continue
    for f in ('n', 'secSum', 'secN'):
        if e[f] != l[f]: print('店', sn, f, '預期', e[f], '線上', l[f]); bad += 1
    keys = set(e['c']) | set(l['c'])
    for k in keys:
        if e['c'].get(k, 0) != l['c'].get(k, 0): print('店', sn, k, '預期', e['c'].get(k, 0), '線上', l['c'].get(k, 0)); bad += 1
tot_e = sum(g['n'] for g in exp.values()); tot_l = sum(g['n'] for g in live.values())
print('預期 %d 筆，線上彙總 %d 筆；店 %d 家；逐項比對差異 %d 處；月表合計 %d' % (tot_e, tot_l, len(live), bad, sum(m['n'] for m in d['months'])))
sys.exit(1 if bad or tot_e != tot_l else 0)
