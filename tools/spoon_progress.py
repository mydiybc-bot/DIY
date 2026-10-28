"""平匙換裝「執行進度」資料：把經營者每處理完一種食材匯出的《平匙換算_<食材>_<日期>.xlsx》
（專案推動 → 🥄 平匙換算 →「📥 匯出 Excel」的「目前畫面」那張）整理成 static/spoon-progress.json，
給 project-hub.html 的「✅ 執行進度」分頁讀。

用法：python3 tools/spoon_progress.py ~/Desktop/專案-平匙換裝/
同一種食材有好幾個檔 → 取檔名日期最新的（同日取修改時間最新的），其餘列在 superseded。
"""
import glob, json, os, re, sys, datetime
import openpyxl

COLS = ['食材', '甜點食譜名稱', '食材重量(g)', '一平匙(g)', '平匙數量', '可否更換', '不可更換原因', '說明', '9 月使用次數（12 店）', '使用目的', '用在哪一步']
KEYS = ['ing', 'recipe', 'grams', 'g', 'spoons', 'ok', 'why', 'note', 'uses', 'purpose', 'step']


def numv(v):
    if v in (None, ''): return ''
    try: return float(v) if float(v) != int(float(v)) else int(float(v))
    except (TypeError, ValueError): return str(v)


def read(path):
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    title = str(rows[0][0] or '')
    m = re.search(r'食材：(.+?)｜一平匙\s*([\d.]+)\s*g｜最多\s*([\d.]+)\s*平匙｜容許誤差\s*([\d.]+)%', title)
    if not m: raise SystemExit('看不懂第一行：%s（%s）' % (title, path))
    hi = next(i for i, r in enumerate(rows) if r and r[0] == '食材')
    head = [str(c or '').strip() for c in rows[hi]][:len(COLS)]
    if head != COLS: raise SystemExit('欄位跟預期不同：%s（%s）' % (head, path))
    out = []
    for r in rows[hi + 1:]:
        if not r or all(c in (None, '') for c in r): continue
        o = dict(zip(KEYS, r[:len(KEYS)]))
        if o['ing'] != m.group(1): raise SystemExit('檔內食材不一致：%s（%s）' % (o['ing'], path))
        for k in ('grams', 'g', 'spoons', 'uses'): o[k] = numv(o[k])
        o['ok'] = o['ok'] == '可更換'
        for k in ('recipe', 'why', 'note', 'purpose', 'step'): o[k] = '' if o[k] is None else str(o[k]).strip()
        del o['ing']
        out.append(o)
    d = re.search(r'_(\d{8})\.xlsx$', os.path.basename(path))
    day = d.group(1) if d else ''
    mt = datetime.datetime.fromtimestamp(os.path.getmtime(path))
    return {
        'ing': m.group(1), 'g': numv(m.group(2)), 'max': numv(m.group(3)), 'tol': numv(m.group(4)),
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
        print('  %s %s：1 平匙 %s g、最多 %s、誤差 %s%%、%d 處（可換 %d）' % (x['date'], x['ing'], x['g'], x['max'], x['tol'], len(x['rows']), sum(r['ok'] for r in x['rows'])))


if __name__ == '__main__':
    main()
