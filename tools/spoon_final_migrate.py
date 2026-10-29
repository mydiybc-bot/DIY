#!/usr/bin/env python3
"""把舊的執行進度（static/spoon-progress.json，桌面 Excel 轉出）搬進專案推動 spoon_final 定案明細（2026-10-29，一次性）。

為什麼要搬：執行進度改讀試算表 spoon_final（平匙換算頁按「✅ 確認轉換到執行進度」寫入），
舊的 6 種食材 204 處不能丟；搬進去時每一處要對到「後台編號＋第幾步」，之後才能照它改正式食譜。

用法：
  python3 tools/spoon_final_migrate.py <spoon.json> <平匙行.tsv> <輸出.gs> [--check]
    spoon.json  ：專案推動 GET ?action=spoon 的回傳（快照 lines 要已經有 recipe_id／step_no／seq，也就是新版 spoon_feed 跑過一次）
    平匙行.tsv  ：後台目前已是「平匙」的行（內建瀏覽器讀 /steps/recipeprint 產生）：recipe_id, 甜點, 第幾步, 步驟標題, 食材, 平匙數
    輸出.gs     ：暫時檔 zz_final_migrate.gs（內含資料＋zzFinalMigrate()；clasp push → 編輯器執行一次 → 刪檔再 push）
對法：
  ① 快照裡同甜點（改過名的照 ALIAS）同食材同克數、照順序 → 後台還寫 g
  ② 對不到 → 後台已是平匙的行，同甜點同食材且平匙數＝定案的 → 已改完後台（泡打粉記成 10/27 Ver.2 替換已套用）
  ③ 還是對不到 → 照樣搬，後台編號留白，回報清單
"""
import json, sys
from collections import defaultdict

ALIAS = {   # Excel 定案時的名稱 → 後台現在的名稱（聖誕新品 10/28 改名）
    '115 聖誕節_雪人紅茶戚風蛋糕': '耶誕花圈紅茶戚風',
    '115 聖誕節_堆雪人慕斯(葷)': '堆雪人莓果慕斯 (葷)',
    '115 聖誕節_雪人樹幹蛋糕': '暖暖雪人樹幹蛋糕',
    '115 聖誕節_聖誕老人生乳酪(葷)': '老公公蘋果生乳酪 (葷)',
    '115 聖誕節_聖誕餅乾': '亮晶晶耶誕餅乾',
}
APPLIED = {'泡打粉': ('2026-10-27', '泡打粉 Ver.2 替換（10/22～10/27）')}


def ns(x):
    x = float(x)
    return str(int(x)) if x == int(x) else str(round(x, 3))


def main():
    if len(sys.argv) < 4:
        sys.exit(__doc__)
    prog = json.load(open('static/spoon-progress.json', encoding='utf-8'))
    sp = json.load(open(sys.argv[1], encoding='utf-8'))
    sp = sp.get('data', sp)
    snap = defaultdict(list)
    for l in sp['lines']:
        if not l.get('recipe_id'):
            sys.exit('快照還沒有後台編號（新版 spoon_feed 要先跑一次）')
        snap[(l['recipe'], l['ingredient'], float(l['grams']))].append(l)
    ml = {s['ingredient']: s.get('spoon_ml', '') for s in sp['set']}
    pf = defaultdict(list)
    for ln in open(sys.argv[2], encoding='utf-8'):
        f = ln.rstrip('\n').split('\t')
        if len(f) == 6:
            pf[(f[1], f[4], float(f[5]))].append({'recipe_id': f[0], 'step_no': f[2], 'step_title': f[3]})
    rows, miss, stat = [], [], defaultdict(int)
    for f in prog['ings']:
        ing, batch = f['ing'], 'M' + f['date'].replace('-', '')[2:] + '-' + f['ing']
        for i, r in enumerate(f['rows']):
            name = ALIAS.get(r['recipe'], r['recipe'])
            k = (name, ing, float(r['grams']))
            o = {'fin_id': batch + '-' + ('000' + str(i))[-4:], 'batch': batch, 'ingredient': ing, 'recipe': name, 'grams': ns(r['grams']),
                 'g_per_spoon': ns(f['g']), 'spoon_ml': str(f.get('ml') or ml.get(ing, '') or ''), 'max_spoons': str(f['max']), 'tol_pct': str(f['tol']),
                 'calc_spoons': ns(r['spoons']) if r['spoons'] != '' else '', 'calc_ok': 'Y' if r['ok'] else 'N', 'calc_why': '' if r['ok'] else (r.get('why') or '維持秤重'),
                 'manual_spoons': '', 'manual_reason': '', 'manual_at': '', 'final_spoons': ns(r['spoons']) if r['ok'] else '',
                 'uses': '' if r.get('uses') in (None, '') else str(r['uses']), 'purpose': r.get('purpose', ''), 'use_step': r.get('step', ''),
                 'source': '搬家：' + f['file'], 'confirmed_at': f.get('mtime', f['date']), 'applied_at': '', 'applied_batch': '',
                 'updated_at': '', 'deleted': '', 'replaced_by': ''}
            if snap[k]:
                l = snap[k].pop(0)
                o.update(recipe_id=l['recipe_id'], step_no=l['step_no'], step_title=l['step_title'], seq=l.get('seq') or '1')
                stat['後台還寫 g'] += 1
            elif r['ok'] and pf[(name, ing, float(r['spoons']))]:
                l = pf[(name, ing, float(r['spoons']))].pop(0)
                o.update(recipe_id=l['recipe_id'], step_no=l['step_no'], step_title=l['step_title'], seq='1')
                if ing in APPLIED:
                    o['applied_at'], o['applied_batch'] = APPLIED[ing]
                stat['後台已是平匙'] += 1
            else:
                o.update(recipe_id='', step_no='', step_title='', seq='1')
                miss.append(name + '｜' + ing + ' ' + ns(r['grams']) + ' g')
                stat['對不到後台'] += 1
            o['line_key'] = '|'.join([o['recipe_id'], o['step_no'], ing, ns(o['grams']), o['seq']]) if o['recipe_id'] else ''
            rows.append(o)
    out = '''/* 暫時檔（2026-10-29）：把舊執行進度 %d 處搬進 spoon_final；編輯器執行 zzFinalMigrate 一次後刪掉這個檔。
   已經有搬家資料（batch 開頭 M）就不重搬。工具：tools/spoon_final_migrate.py */
var ZZ_FINAL_ROWS = %s;
var ZZ_FINAL_SET = %s;
function zzFinalMigrate() {
  var live = rows_('spoon_final').filter(function (r) { return r.deleted !== 'Y'; });
  if (live.some(function (r) { return /^M/.test(r.batch); })) throw new Error('已經搬過了（有 M 開頭的批號），不重搬');
  var t = now_();
  append_('spoon_final', ZZ_FINAL_ROWS.map(function (o) { o.updated_at = t; return o; }));
  /* 平匙規格表同步成 Excel 定案的規格（10/28 只改了一平匙克數，最多幾匙／誤差還是舊的） */
  ZZ_FINAL_SET.forEach(function (x) {
    var h = rows_('spoon_set').filter(function (r) { return r.ingredient === x.ingredient; })[0]; if (!h) return;
    var o = clean_(h); o.g_per_spoon = x.g; o.max_spoons = x.max; o.tol_pct = x.tol; if (x.ml) o.spoon_ml = x.ml; o.note = '執行進度定案（搬家）'; o.saved_at = t;
    writeRow_('spoon_set', h._row, o); specLog_(o, '執行進度定案（搬家）');
  });
  log_('finalMigrate', 'spoon_final', '搬家', ZZ_FINAL_ROWS.length + ' 處（%s）');
  return { added: ZZ_FINAL_ROWS.length, total: rows_('spoon_final').length };
}
''' % (len(rows), json.dumps(rows, ensure_ascii=False, indent=0),
       json.dumps([{'ingredient': f['ing'], 'g': ns(f['g']), 'max': str(f['max']), 'tol': str(f['tol']), 'ml': str(f.get('ml') or '')} for f in prog['ings']], ensure_ascii=False),
       '、'.join('%s %d' % kv for kv in stat.items()))
    if '--check' not in sys.argv:
        open(sys.argv[3], 'w', encoding='utf-8').write(out)
    print('共 %d 處：%s' % (len(rows), '、'.join('%s %d' % kv for kv in stat.items())))
    if miss:
        print('對不到後台：\n  ' + '\n  '.join(miss))


if __name__ == '__main__':
    main()
