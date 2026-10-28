#!/usr/bin/env python3
"""
把 convert.py 的輸出分批送進「顧客問卷」GAS（import_ 動作）。
  python3 -I tools/survey-import/send_import.py <convert 輸出資料夾> count            只查各表筆數（試密碼、不寫入）
  python3 -I tools/survey-import/send_import.py <convert 輸出資料夾> backup           備份相關表
  python3 -I tools/survey-import/send_import.py <convert 輸出資料夾> rows 6.0 5.0 ...  匯入 rows_<版本>.jsonl → fact_response（同 id 會跳過，可重跑）
  python3 -I tools/survey-import/send_import.py <convert 輸出資料夾> raw  1.0_第一版:raw_v1 2.0:raw_v2 ...  原樣封存成 raw_ 表
一次性匯入碼從環境變數 IMPORT_CODE 讀（GAS 編輯器執行 importOpen 產生，60 分鐘有效），不印、不存檔。
"""
import sys, os, json, subprocess, time, tempfile

URL = 'https://script.google.com/macros/s/AKfycbxrsYaYLdXqjUeGvG8_CZ--5znw5izXr6tcgmfC0MbJ32T1KY_TS8UkVqimxS2Jajx0/exec'

def pwd():
    c = os.environ.get('IMPORT_CODE', '').strip()
    if not c: raise SystemExit('請先在 GAS 編輯器執行 importOpen，把執行記錄裡的碼放進環境變數 IMPORT_CODE')
    return c

def post(body, tries=3):
    last = None
    for k in range(tries):
        with tempfile.NamedTemporaryFile('w', suffix='.json', delete=False, encoding='utf-8') as f:
            json.dump(body, f, ensure_ascii=False); name = f.name
        try:
            r = subprocess.run(['curl', '-sL', '--max-time', '420', '-H', 'Content-Type: text/plain', '--data-binary', '@' + name, URL], capture_output=True, text=True)
            try: d = json.loads(r.stdout)
            except Exception: d = {'ok': False, 'msg': 'not json: ' + r.stdout[:120]}
            if d.get('ok'): return d
            last = d
            if d.get('code') in ('auth', 'input'): break          # 重試也沒用
        finally:
            os.unlink(name)
        time.sleep(5 * (k + 1))
    raise SystemExit('失敗：%s' % json.dumps(last, ensure_ascii=False)[:300])

def log(msg, fh=None):
    print(msg, flush=True)
    if fh: fh.write(msg + '\n'); fh.flush()

def main():
    if len(sys.argv) < 3: print(__doc__); sys.exit(1)
    outdir, cmd, args = os.path.expanduser(sys.argv[1]), sys.argv[2], sys.argv[3:]
    P = pwd(); fh = open(os.path.join(outdir, 'import_log.txt'), 'a', encoding='utf-8')
    log('=== %s %s %s' % (time.strftime('%Y-%m-%d %H:%M:%S'), cmd, ' '.join(args)), fh)
    if cmd == 'count':
        log(json.dumps(post({'action': 'import', 'code': P, 'op': 'count'})['counts'], ensure_ascii=False), fh)
    elif cmd == 'backup':
        d = post({'action': 'import', 'code': P, 'op': 'backup'}); log('備份：' + json.dumps(d['backup'], ensure_ascii=False), fh); log('現況：' + json.dumps(d['counts'], ensure_ascii=False), fh)
    elif cmd == 'rows':
        for lab in args:
            recs = [json.loads(l) for l in open(os.path.join(outdir, 'rows_%s.jsonl' % lab), encoding='utf-8')]
            recs.sort(key=lambda r: r['t'])
            added = skipped = 0; t0 = time.time(); B = 2500
            for i in range(0, len(recs), B):
                batch = recs[i:i + B]
                d = post({'action': 'import', 'code': P, 'rows': batch})
                if d['added'] + d['skipped'] != len(batch): raise SystemExit('筆數對不上：送 %d、回 %s' % (len(batch), d))
                added += d['added']; skipped += d['skipped']
                log('  %s  %d/%d  新增 %d 跳過 %d  表內合計 %d  (%ds)' % (lab, min(i + B, len(recs)), len(recs), added, skipped, d['total'], time.time() - t0), fh)
            log('版本 %s 完成：檔案 %d 筆 → 新增 %d、跳過 %d' % (lab, len(recs), added, skipped), fh)
    elif cmd == 'raw':
        for spec in args:
            lab, sheet = spec.split(':')
            d0 = json.load(open(os.path.join(outdir, 'raw_%s.json' % lab), encoding='utf-8'))
            rows = d0['rows']; added = skipped = 0; t0 = time.time(); B = 3000 if len(d0['head']) < 30 else 1200
            for i in range(0, len(rows), B):
                batch = rows[i:i + B]
                d = post({'action': 'import', 'code': P, 'op': 'raw', 'sheet': sheet, 'head': d0['head'], 'key': d0['key'], 'rows': batch})
                if d['added'] + d['skipped'] != len(batch): raise SystemExit('筆數對不上：%s' % d)
                added += d['added']; skipped += d['skipped']
                log('  %s→%s  %d/%d  新增 %d 跳過 %d  (%ds)' % (lab, sheet, min(i + B, len(rows)), len(rows), added, skipped, time.time() - t0), fh)
            log('封存 %s → %s 完成：檔案 %d 列 → 新增 %d、跳過 %d' % (lab, sheet, len(rows), added, skipped), fh)
    else:
        print(__doc__)

if __name__ == '__main__':
    main()
