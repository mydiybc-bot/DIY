#!/usr/bin/env python3
"""外國客人到店須知（QR Code 頁 store-guide.html）的甜點資料：跟平板外語食譜同一份（2026-11-26）
經營者：「QR code 掃描後的資料還是停留在 26 支食譜，這不對…未來要能和每天的外語食譜更新一起進行更新檢查」；
「88 支都要上，不要自行合併」；甜點介紹「一律用後台中文食譜」。翻譯一律照台灣區中文逐字逐句翻（不花 API，Claude Code 翻）。

python3 guide.py pull   → 讀平板中文原文上線中的每一支（申請單 guideSrc：後台售價／尺寸／製作時間／產品特色＋平板放行中的外語品名）
                          → 跟到店須知 API 已有的逐行譯文比 → WORK/guide/todo_<語言>.json（還沒翻的行），印出摘要
python3 guide.py put [--dry] → 用 WORK/guide/tr_<語言>.json（新譯文）＋已有譯文組好 88 支 → 送到店須知 API 的 put（伺服器檢查過才發佈）
python3 guide.py check [語言…] → 硬檢查譯文（代號、換行、殘留中文、數字）
python3 guide.py fix <語言> → 套用回讀員 review_<語言>.json 的 fix
python3 guide.py status → 到店須知 API 目前發佈的款數、最後更新、缺譯
通行碼：I18N_PREVIEW_KEY（讀申請單）、I18N_PUT_KEY（到店須知寫入；伺服器只存雜湊）。不印出。"""
import sys, json, os, re, hashlib, subprocess
from concurrent.futures import ThreadPoolExecutor
from common import *

GUIDE_API = 'https://script.google.com/macros/s/AKfycbzKmiZv95tW2b_1oAr-DiJf90cmpadfkb0MI2WBikqD_uVE6uUSUf_vTNa6KcbjMYc4BQ/exec'
LANGS = ['en', 'ja', 'ko', 'vi', 'id', 'th']
GW = os.path.join(WORK, 'guide'); os.makedirs(GW, exist_ok=True)
CJK = re.compile(r'[㐀-鿿]')


def gpost(body, timeout=300):
    data = json.dumps(body, ensure_ascii=False).encode('utf-8')
    r = subprocess.run(['curl', '-sL', '--max-time', str(timeout), '-H', 'Content-Type:', '--data-binary', '@-', GUIDE_API], input=data, capture_output=True)
    try: return json.loads(r.stdout.decode('utf-8'))
    except Exception: return {'ok': False, 'error': '回應看不懂：' + r.stdout[:200].decode('utf-8', 'ignore')}


def gget(action):
    r = subprocess.run(['curl', '-sL', '--max-time', '120', GUIDE_API + '?action=' + action], capture_output=True)
    return json.loads(r.stdout.decode('utf-8'))


def lkey(s):
    return hashlib.sha1(s.encode('utf-8')).hexdigest()[:12]


def needs_tr(line):   # 有中文字才要翻；分隔線、純符號照抄
    return bool(CJK.search(line))


def src_lines(it):
    """要翻的中文行：產品特色每一行＋尺寸（照原字）"""
    e = it.get('edit') or {}
    out = [l for l in (e.get('content') or '').split('\n') if needs_tr(l)]
    if needs_tr(e.get('size') or ''): out.append(e['size'])
    return out


def pull():
    pk = key('I18N_PREVIEW_KEY')
    lst = call('guideSrc', pkey=pk)
    if not lst.get('ok'): sys.exit('讀清單失敗：%s' % lst.get('msg'))
    ids = [x['id'] for x in lst['data']['items']]
    chunks = [ids[i:i + 8] for i in range(0, len(ids), 8)]
    def one(ch):
        r = call('guideSrc', pkey=pk, ids=ch, timeout=300)
        if not r.get('ok'): raise SystemExit('讀內容失敗：%s' % r.get('msg'))
        return r['data']['items']
    with ThreadPoolExecutor(4) as ex:
        items = [x for part in ex.map(one, chunks) for x in part]
    order = {i: n for n, i in enumerate(ids)}
    items.sort(key=lambda x: order[x['id']])
    json.dump({'at': lst['data']['at'], 'items': items}, open(os.path.join(GW, 'src.json'), 'w'), ensure_ascii=False, indent=1)
    json.dump({x['name']: x['names'] for x in items}, open(os.path.join(GW, 'names.json'), 'w'), ensure_ascii=False, separators=(',', ':'))   # 翻譯員用（src.json 太大）
    json.dump([{'zh': x['name'], 'content': (x.get('edit') or {}).get('content') or ''} for x in items], open(os.path.join(GW, 'ctx.json'), 'w'), ensure_ascii=False, separators=(',', ':'))
    have = gget('lines').get('lines') or {}
    errs = [x['name'] + '：' + x['err'] for x in items if x.get('err')]
    miss_name = [x['name'] + '（' + ','.join(l for l in LANGS if not x['names'].get(l)) + '）' for x in items if any(not x['names'].get(l) for l in LANGS)]
    lines = {}
    for it in items:
        for l in src_lines(it): lines[lkey(l)] = l
    print('平板中文原文上線 %d 支；後台讀取失敗 %d；外語品名缺 %d；要翻的中文行（不重複）%d' % (len(items), len(errs), len(miss_name), len(lines)))
    for e in errs: print('  ⚠', e)
    for m in miss_name: print('  ⚠ 品名缺', m)
    for lg in LANGS:
        todo = {k: v for k, v in lines.items() if not (have.get(k) or {}).get(lg)}
        p = os.path.join(GW, 'todo_%s.json' % lg)
        if todo: json.dump(todo, open(p, 'w'), ensure_ascii=False, indent=1)
        elif os.path.exists(p): os.remove(p)
        print('  %s 待翻 %d 行' % (lg, len(todo)))


# ---------- 組資料 ----------
def size_of(s):
    m = re.search(r'(\d+)\s*吋\s*[×xX]\s*(\d+)\s*個', s)
    if m: return ['tart', int(m.group(1)), int(m.group(2))]
    m = re.fullmatch(r'\s*(\d+)\s*吋\s*', s)
    if m: return ['inch', int(m.group(1))]
    m = re.fullmatch(r'\s*(\d+)\s*(?:個|片|顆|入|條|塊)\s*', s)
    if m: return ['pcs', int(m.group(1))]
    return None


def tags_of(it):
    nm, c = it['name'], (it.get('edit') or {}).get('content') or ''
    t = []
    if re.search(r'限當日壽星', nm): t.append('bday')
    if re.search(r'信義、高雄無法製作', nm + c): t.append('noXK')
    if re.search(r'[(（]\s*葷\s*[)）]', nm): t.append('meat')
    if re.search(r'限信義、高雄', nm + ' '.join((it.get('edit') or {}).get('cats') or [])): t.append('bread')
    if re.search(r'1\s*人或\s*2\s*人製作|2\s*人一起製作', c): t.append('pair')
    return t


BDAY_CUT = re.compile(r'當日壽星優惠|^生日當天')


def shown_content(it):
    """客人頁顯示的產品特色：壽星款（限當日壽星）的壽星優惠說明那段起不顯示，改用頁面固定說明（2026-10-29 經營者：出示證件即可、不用官網會員；
    後台舊食譜還寫要官網會員）。前面的分隔線、空行一起拿掉。"""
    lines = ((it.get('edit') or {}).get('content') or '').split('\n')
    if 'bday' not in tags_of(it): return '\n'.join(lines)
    for i, l in enumerate(lines):
        if BDAY_CUT.search(l):
            lines = lines[:i]
            while lines and (not lines[-1].strip() or re.fullmatch(r'[-—─\s]+', lines[-1])): lines.pop()
            break
    return '\n'.join(lines)


def build():
    S = json.load(open(os.path.join(GW, 'src.json')))
    have = gget('lines').get('lines') or {}
    for lg in LANGS:
        p = os.path.join(GW, 'tr_%s.json' % lg)
        if not os.path.exists(p): continue
        for k, v in json.load(open(p)).items():
            have.setdefault(k, {})[lg] = v
    items, miss = [], []
    for it in S['items']:
        e = it.get('edit') or {}
        if it.get('err') or not e: miss.append(it['name'] + '：後台讀不到'); continue
        price = float(e.get('price') or 'nan') if re.fullmatch(r'\s*\d+(\.\d+)?\s*', e.get('price') or '') else None
        content = shown_content(it)
        t = {'zh': ['', content, e.get('size') or '']}
        for lg in LANGS:
            def tr(line):
                if not needs_tr(line): return line
                v = (have.get(lkey(line)) or {}).get(lg)
                if not v: miss.append('%s %s：%s' % (lg, it['name'], line[:30]))
                return v or ''
            desc = '\n'.join(tr(l) for l in content.split('\n'))
            t[lg] = [it['names'].get(lg) or '', desc, tr(e.get('size') or '')]
        items.append({'v': 2, 'id': it['id'], 'en': it['names'].get('en') or '', 'zh': it['name'], 'p': int(round(price)) if price is not None else None,
                      'img': it['cover'], 's': size_of(e.get('size') or ''), 'size_raw': e.get('size') or '',
                      'prep': float(e['prep']) if re.fullmatch(r'\s*\d+(\.\d+)?\s*', e.get('prep') or '') else None,
                      'tags': tags_of(it), 'cats': e.get('cats') or [], 't': t})
    used = {}
    for it in S['items']:
        for l in src_lines(it):
            k = lkey(l)
            used[k] = dict(have.get(k) or {}, zh=l)
    return items, used, miss


def put(dry):
    items, used, miss = build()
    print('組好 %d 款；缺譯 %d 處' % (len(items), len(miss)))
    for m in miss[:40]: print('  ⚠', m)
    if miss: sys.exit('還有缺譯，先翻完（todo_<語言>.json）再 put')
    if dry:
        json.dump(items, open(os.path.join(GW, 'items_dry.json'), 'w'), ensure_ascii=False, indent=1)
        print('試算：', os.path.join(GW, 'items_dry.json')); return
    r = gpost({'action': 'put', 'key': key('I18N_PUT_KEY'), 'items': items, 'lines': used})
    print(json.dumps(r.get('status') or r, ensure_ascii=False, indent=1)[:3000])
    if r.get('ok'):
        for lg in LANGS:
            p = os.path.join(GW, 'tr_%s.json' % lg)
            if os.path.exists(p): os.replace(p, p + '.done')
    else: sys.exit(1)


def check(lgs):
    """硬檢查 tr_<語言>.json：代號跟 todo 一樣、一行、沒有殘留中文（ja 除外）、數字一樣、價格與日期一樣"""
    bad = 0
    for lg in lgs:
        tp, rp = os.path.join(GW, 'todo_%s.json' % lg), os.path.join(GW, 'tr_%s.json' % lg)
        if not os.path.exists(tp): continue
        todo = json.load(open(tp)); tr = json.load(open(rp)) if os.path.exists(rp) else {}
        errs = ['少了 %s' % k for k in todo if k not in tr] + ['多了 %s' % k for k in tr if k not in todo]
        for k, zh in todo.items():
            v = tr.get(k)
            if v is None: continue
            if '\n' in v: errs.append('%s 有換行' % k)
            if lg != 'ja' and CJK.search(v): errs.append('%s 還有中文：%s' % (k, v[:40]))
            nz = sorted(re.findall(r'\d+(?:\.\d+)?', zh)); nv = sorted(re.findall(r'\d+(?:\.\d+)?', v))
            miss = [x for x in nz if x not in nv and not (x == '69' and '31' in nv)]   # 只抓「中文有、譯文沒有」的數字（譯文多出「1 個」之類量詞不算）；69 折＝31% off
            if miss: errs.append('%s 少了數字 %s：%s → %s' % (k, miss, zh[:30], v[:50]))
        print('%s：%d 行，問題 %d' % (lg, len(todo), len(errs)))
        for e in errs[:30]: print('   ', e)
        bad += len(errs)
    return bad


def fix(lg):
    """把回讀員 review_<語言>.json 的 fix 套進 tr_<語言>.json（只修一輪）"""
    rp, tp = os.path.join(GW, 'review_%s.json' % lg), os.path.join(GW, 'tr_%s.json' % lg)
    rv = json.load(open(rp)); tr = json.load(open(tp)); n = 0
    for it in rv.get('issues') or []:
        k, f = it.get('k'), (it.get('fix') or '').strip()
        if k in tr and f and f != tr[k]: tr[k] = f; n += 1
    json.dump(tr, open(tp, 'w'), ensure_ascii=False, indent=1)
    print('%s：套用 %d 處（回讀列 %d 處）' % (lg, n, len(rv.get('issues') or [])))


def status():
    d = gget('data')
    s = d.get('status') or {}
    print('到店須知：發佈 %d 款｜最後更新 %s（%s）｜結果 %s' % (len(d.get('items') or []), s.get('at'), s.get('by'), s.get('result')))
    for k in ('added', 'removed', 'price', 'text', 'errors', 'warnings'):
        if s.get(k): print('  %s：%s' % (k, '、'.join(s[k])))


if __name__ == '__main__':
    a = sys.argv[1:] or ['status']
    if a[0] == 'pull': pull()
    elif a[0] == 'put': put('--dry' in a)
    elif a[0] == 'check': sys.exit(1 if check(a[1:] or LANGS) else 0)
    elif a[0] == 'fix': fix(a[1])
    else: status()
