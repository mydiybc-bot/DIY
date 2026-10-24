#!/usr/bin/env python3
"""投遞：python3 put.py <id8_lang> [--approve] [--model 名稱]
讀 WORK/<id8_lang>.tr.json（{name, pres, slides:[{k,t,h}], dict:[{kind,en,tr}]}），連同拿稿時的指紋一起送。
--approve：硬檢查全過才放行（沒過只存成待檢查）。"""
import json, sys
from common import *

t = sys.argv[1]; approve = '--approve' in sys.argv
model = sys.argv[sys.argv.index('--model') + 1] if '--model' in sys.argv else 'claude-code-sonnet-5-5'
src = json.load(open('%s/%s.src.json' % (WORK, t), encoding='utf-8'))
tr = json.load(open('%s/%s.tr.json' % (WORK, t), encoding='utf-8'))
r = call('i18nPut', wkey=key('I18N_PUT_KEY'), id=src['id'], lang=src['lang'], fp=src['fp'], name=tr.get('name', ''), pres=tr.get('pres', ''),
         slides=tr['slides'], dict=tr.get('dict', []), approve=1 if approve else 0, overwrite=1 if '--overwrite' in sys.argv else 0, model=model, by='Claude Code')
print(json.dumps(r, ensure_ascii=False, indent=1)[:1500])
sys.exit(0 if r.get('ok') else 1)
