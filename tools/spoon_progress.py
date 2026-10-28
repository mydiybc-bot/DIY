"""平匙換裝「執行進度」資料：把經營者每處理完一種食材匯出的《平匙換算_<食材>_<日期>.xlsx》
（專案推動 → 🥄 平匙換算 →「📥 匯出 Excel」的「目前畫面」那張）整理成 static/spoon-progress.json，
給 project-hub.html 的「✅ 執行進度」分頁讀。

用法：python3 tools/spoon_progress.py ~/Desktop/專案-平匙換裝/
同一種食材有好幾個檔 → 取檔名日期最新的（同日取修改時間最新的），其餘列在 superseded。
"""
import glob, json, os, re, sys, datetime
import openpyxl

# 欄位照表頭名稱對位（2026-10-28 起匯出多一欄「量匙規格」；「N 月使用次數」月份會變）
COLS = {'食材': 'ing', '甜點食譜名稱': 'recipe', '食材重量(g)': 'grams', '量匙規格': 'mlt', '一平匙(g)': 'g', '平匙數量': 'spoons',
        '可否更換': 'ok', '不可更換原因': 'why', '說明': 'note', '使用目的': 'purpose', '用在哪一步': 'step'}
NEED = ['ing', 'recipe', 'grams', 'g', 'spoons', 'ok', 'why', 'note', 'purpose', 'step']


def numv(v):
    if v in (None, ''): return ''
    try: return float(v) if float(v) != int(float(v)) else int(float(v))
    except (TypeError, ValueError): return str(v)


def read(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    title = str(rows[0][0] or '')
    m = re.search(r'食材：(.+?)｜(?:量匙\s*(.+?)｜)?一平匙\s*([\d.]+)\s*g｜最多\s*([\d.]+\s*平匙|不限)｜容許誤差\s*([\d.]+)%', title)
    if not m: raise SystemExit('看不懂第一行：%s（%s）' % (title, path))
    ing, mlt, g, mx, tol = m.group(1), (m.group(2) or '').strip(), m.group(3), m.group(4), m.group(5)
    mln = re.match(r'([\d.]+)\s*ml', mlt)
    hi = next(i for i, r in enumerate(rows) if r and r[0] == '食材')
    idx = {}
    for j, c in enumerate(rows[hi]):
        c = str(c or '').strip()
        if c in COLS: idx[COLS[c]] = j
        elif re.match(r'(\d+ 月|上月)使用次數', c): idx['uses'] = j
    miss = [k for k in NEED if k not in idx]
    if miss: raise SystemExit('缺欄位 %s（%s）' % (miss, path))
    out = []
    for r in rows[hi + 1:]:
        if not r or all(c in (None, '') for c in r): continue
        o = {k: (r[j] if j < len(r) else None) for k, j in idx.items()}
        if o['ing'] != ing: raise SystemExit('檔內食材不一致：%s（%s）' % (o['ing'], path))
        for k in ('grams', 'g', 'spoons', 'uses'): o[k] = numv(o.get(k))
        o['ok'] = o['ok'] == '可更換'
        for k in ('recipe', 'why', 'note', 'purpose', 'step'): o[k] = '' if o[k] is None else str(o[k]).strip()
        o.pop('ing'); o.pop('mlt', None)
        out.append(o)
    d = re.search(r'_(\d{8})\.xlsx$', os.path.basename(path))
    day = d.group(1) if d else ''
    mt = datetime.datetime.fromtimestamp(os.path.getmtime(path))
    return {
        'ing': ing, 'ml': numv(mln.group(1)) if mln else '', 'g': numv(g), 'max': numv(re.sub(r'\s*平匙', '', mx)), 'tol': numv(tol),
        'trial': '不一定已存檔' in title,
        'date': '%s-%s-%s' % (day[:4], day[4:6], day[6:]) if day else mt.strftime('%Y-%m-%d'),
        'mtime': mt.strftime('%Y-%m-%d %H:%M'), 'file': os.path.basename(path), 'rows': out,
    }


def main():
    src = os.path.expanduser(sys.argv[1] if len(sys.argv) > 1 else '~/Desktop/專案-平匙換裝/')
    here = os.path.dirname(os.path.abspath(__file__))
    dst = os.path.join(here, '..', 'static', 'spoon-progress.json')
    files = sorted(glob.glob(os.path.join(src, '平匙換算_*.xlsx')))
    if not files: raise SystemExit('資料夾裡沒有《平匙換算_*.xlsx》：' + src)
    by, old = {}, []
    for f in files:
        x = read(f)
        k = (x['date'], x['mtime'])
        if x['ing'] in by and (by[x['ing']]['date'], by[x['ing']]['mtime']) >= k: old.append(x['file']); continue
        if x['ing'] in by: old.append(by[x['ing']]['file'])
        by[x['ing']] = x
    ings = sorted(by.values(), key=lambda x: (x['date'], x['mtime']))
    doc = {'built': datetime.datetime.now().strftime('%Y-%m-%d %H:%M'), 'folder': os.path.basename(os.path.normpath(src)),
           'ings': ings, 'superseded': old}
    with open(dst, 'w', encoding='utf-8') as fh:
        json.dump(doc, fh, ensure_ascii=False, indent=1)
        fh.write('\n')
    n = sum(len(x['rows']) for x in ings)
    print('寫入 %s：%d 種食材、%d 處%s' % (os.path.normpath(dst), len(ings), n, '；被較新檔取代：' + '、'.join(old) if old else ''))
    for x in ings:
        print('  %s %s：量匙 %s、1 平匙 %s g、最多 %s、誤差 %s%%、%d 處（可換 %d）' % (x['date'], x['ing'], (str(x['ml']) + ' ml') if x['ml'] != '' else '未填（Excel 沒寫）', x['g'], x['max'], x['tol'], len(x['rows']), sum(r['ok'] for r in x['rows'])))


if __name__ == '__main__':
    main()
