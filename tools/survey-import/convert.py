#!/usr/bin/env python3
"""
SurveyCake 歷史回覆 → 新問卷 fact_response 格式（2026-10-28）

用法（請用隔離模式，因為輸入檔是下載來的）：
  python3 -I tools/survey-import/convert.py <輸入資料夾> <輸出資料夾> [版本 ...]
  例：python3 -I tools/survey-import/convert.py ~/Desktop/Claude/Code/資料交換/輸入/顧客問卷 /tmp/out 6.0 5.0 4.0

做的事：
  1. 每個版本資料夾（6.0／5.0／4.0／3.0／2.0）讀 .xlsx，用「題目文字」對到新問卷的題號 q1…q16（不靠欄位順序）。
  2. 選項文字正規化成新問卷的選項（全形括號、～、空格、店名等）；舊選項改名的用 SYN 表；對不上的不猜，記在「未對應」。
  3. 輸出：
       <輸出>/rows_<版本>.jsonl   一行一筆，給 GAS import_ 分批送（欄位 t／id／sn／s／sec／a／ua）
       <輸出>/preview.json        摘要（給 make_preview 做報表）
本程式只讀檔、不連網路寫資料；新問卷的標準選項從線上 ?action=form 讀（讀不到就用內建備份）。
"""
import sys, os, re, json, glob, collections, subprocess, datetime
import openpyxl

FORM_URL = 'https://script.google.com/macros/s/AKfycbxrsYaYLdXqjUeGvG8_CZ--5znw5izXr6tcgmfC0MbJ32T1KY_TS8UkVqimxS2Jajx0/exec?action=form'

# 標題文字（關鍵字）→ 新題號。順序重要：先比對先贏。
HEADER_Q = [
    (r'今日體驗的分店', 'q1'),
    (r'第幾次來做甜點', 'q2'),
    (r'入座後，服務人員的解說', 'q3'),
    (r'第一個考慮的是', 'q4'),
    (r'您這次來的原因', 'q5'),
    (r'平板教學影片', 'q6'),
    (r'店內的環境', 'q7'),
    (r'(服務人員|人員)的服務', 'q8'),
    (r'是否有提供您任何的協助', 'q9'),            # 6.0、5.0。4.0 問的是「是否遇到需要協助」，意思不同，不對應
    (r'(服務人員有協助您的地方|服務人員協助您的地方)', 'q10'),
    (r'您今天做的甜點是', 'q11'),
    (r'甜點是否有製作成功', 'q12'),
    (r'製作甜點的時間', 'q13'),
    (r'會推薦其他人來嗎', 'q14'),
    (r'請問您的年齡', 'q15'),
    (r'透過哪個途徑得知', 'q16'),
]
MULTI = {'q10', 'q11'}

# 店名（正規化後）→ (新店名, 店號)。台南新光西門店已結束 → 店號 0
def _k(s): return re.sub(r'[\s　]', '', s).lower().replace('(', '（').replace(')', '）')
STORE_MAP = {
    '台北南京店': ('台北南京', 3), '台北士林店': ('台北士林', 4), '吳寶春自己做（台北信義a13）': ('吳寶春自己做（台北信義A13）', 11),
    '新北板橋店': ('新北板橋', 7), '新北新店店': ('新北新店', 8), '桃園中壢店': ('桃園中壢', 9), '桃園藝文店': ('桃園藝文', 10),
    '桃園中正藝文店': ('桃園藝文', 10), '新竹文化店': ('新竹文化', 6), '台中精明店': ('台中精明', 1), '台中草悟道店': ('台中草悟道', 2),
    '台南focus店': ('台南Focus', 5), '吳寶春自己做（高雄skmpark）': ('吳寶春自己做（高雄SKM Park）', 12),
    '台南新光西門店': ('台南新光西門（已結束）', 0),
}

# 舊選項改過名／合併的（正規化後的舊值 → 新選項）。每一條都會列在預覽報告給經營者確認。
SYN = {
    'q4': {'難易度': '不容易失敗', '流行性': '現在流行'},
    # 5.0 的「協助的地方」12 個細項，6.0 合併成「製作過程時的協助，如補救、教學、設備使用、蛋糕抹面、找食材／器模具…」
    'q10': {
        '找工具': 'MAKE', '找食材': 'MAKE', '設備使用，如烤箱、攪拌機、微波爐': 'MAKE', '甜點補救': 'MAKE', '製作技巧示範教學': 'MAKE',
        '烤模拿取': 'MAKE', '確認烤熟了沒': 'MAKE', '蛋糕抹面': 'MAKE', '蛋糕裝飾': 'MAKE', '讚美鼓勵我': '給予讚美鼓勵',
    },
}
# 4.0 沒有 q9／q13；3.0、2.0 的結構會在讀到時補（見 profile）
IGNORE_HEAD = {'填答時間', '填答秒數', 'IP紀錄', '額滿結束註記', '使用者紀錄', '會員時間', 'Hash', '會員編號', '自訂ID', '備註'}


def nz(s):
    s = str(s).strip()
    s = s.replace('~', '～').replace('/', '／').replace('...', '…').replace('..', '…').replace(',', '，').replace('(', '（').replace(')', '）')
    return re.sub(r'[\s　]', '', s).lower()


def load_canon():
    try:
        out = subprocess.run(['curl', '-sL', '--max-time', '40', FORM_URL], capture_output=True, text=True).stdout
        d = json.loads(out)
        return {q['qid']: q['options'] for q in d['questions']}
    except Exception as e:  # 讀不到就用備份
        sys.stderr.write('讀不到線上題目，改用內建備份：%s\n' % e)
        return json.load(open(os.path.join(os.path.dirname(__file__), 'canon_backup.json'), encoding='utf-8'))


def header_map(head):
    m = {}
    for i, h in enumerate(head):
        if h in IGNORE_HEAD or h is None: continue
        for pat, q in HEADER_Q:
            if re.search(pat, str(h)):
                if q not in m.values(): m[i] = q
                break
    return m


def parse_time(v):
    if isinstance(v, datetime.datetime): return v
    s = str(v).strip()
    for f in ('%Y-%m-%d %H:%M:%S', '%Y/%m/%d %H:%M:%S', '%Y-%m-%d %H:%M'):
        try: return datetime.datetime.strptime(s, f)
        except ValueError: pass
    return None


def convert_version(label, folder, canon, outdir):
    files = sorted(glob.glob(os.path.join(folder, '*.xlsx')))
    if not files: return None
    wb = openpyxl.load_workbook(files[0], read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    it = ws.iter_rows(values_only=True)
    head = list(next(it))
    hm = header_map(head)
    iT, iSec, iHash = head.index('填答時間'), head.index('填答秒數'), head.index('Hash')
    canon_nz = {q: {nz(o): o for o in opts} for q, opts in canon.items() if opts}
    unmapped = collections.defaultdict(collections.Counter)      # qid -> 原值 -> 次數
    syn_used = collections.defaultdict(collections.Counter)      # qid -> "原值→新值" -> 次數
    answered = collections.Counter()
    stores = collections.Counter(); months = collections.Counter()
    n = bad_time = no_store = 0; tmin = tmax = None; seen = set(); dup = 0
    out = open(os.path.join(outdir, 'rows_%s.jsonl' % label), 'w', encoding='utf-8')
    sample = []
    for r in it:
        n += 1
        t = parse_time(r[iT])
        if t is None: bad_time += 1; continue
        h = str(r[iHash] or '').strip()
        if not h: bad_time += 1; continue
        rid = 'sc%s-%s' % (label.split('_')[0].replace('.', ''), h)
        if rid in seen: dup += 1; continue
        seen.add(rid)
        a = {}; sn = ''; sname = ''
        for i, q in hm.items():
            raw = r[i]
            if raw is None or str(raw).strip() == '' or str(raw).strip() == '-': continue
            if q == 'q1':
                k = _k(str(raw))
                if k in STORE_MAP: sname, sn = STORE_MAP[k]; a['q1'] = sname
                else: unmapped['q1'][str(raw)] += 1
                continue
            parts = [p.strip() for p in str(raw).split('\n') if p.strip()] if q in MULTI else [str(raw).strip()]
            outv = []
            for p in parts:
                key = nz(p)
                val = canon_nz.get(q, {}).get(key)
                if val is None and q in SYN:
                    for old, new in SYN[q].items():
                        if nz(old) == key:
                            if new == 'MAKE':
                                new = [o for o in canon[q] if o.startswith('製作過程時的協助')][0]
                            val = new; syn_used[q]['%s → %s' % (p, val)] += 1; break
                if val is None:
                    # 「其他-xxx」自由填寫 → 歸「其他」（若該題有「其他」選項）
                    if key.startswith('其他') and nz('其他') in canon_nz.get(q, {}): val = '其他'
                if val is None: unmapped[q][p] += 1; continue
                if val not in outv: outv.append(val)
            if outv:
                a[q] = '|'.join(outv); answered[q] += 1
        if sn == '' : no_store += 1; continue
        sec = 0
        try: sec = int(float(str(r[iSec]).strip()))
        except Exception: sec = 0
        iso = t.strftime('%Y-%m-%dT%H:%M:%S+08:00')
        stores[sname] += 1; months[t.strftime('%Y-%m')] += 1
        tmin = t if tmin is None or t < tmin else tmin; tmax = t if tmax is None or t > tmax else tmax
        rec = {'t': iso, 'id': rid, 'sn': sn, 's': sname, 'sec': sec, 'a': a, 'ua': 'surveycake-import:' + label}
        out.write(json.dumps(rec, ensure_ascii=False) + '\n')
        if len(sample) < 8 and n % 997 == 3: sample.append({'raw': [str(x)[:30] if x is not None else None for x in r[:16]], 'new': rec})
    out.close()
    written = sum(stores.values())
    return {
        'label': label, 'file': os.path.basename(files[0]), 'excel_rows': n, 'written': written, 'bad_time': bad_time, 'no_store': no_store, 'dup_hash': dup,
        'time_from': tmin.strftime('%Y-%m-%d %H:%M:%S') if tmin else '', 'time_to': tmax.strftime('%Y-%m-%d %H:%M:%S') if tmax else '',
        'stores': dict(stores), 'months': dict(sorted(months.items())),
        'headers_mapped': {str(head[i]): q for i, q in hm.items()},
        'headers_unmapped': [str(h) for i, h in enumerate(head) if i not in hm and h not in IGNORE_HEAD and h is not None],
        'answered': dict(answered), 'unmapped': {q: dict(c) for q, c in unmapped.items()},
        'syn_used': {q: dict(c) for q, c in syn_used.items()}, 'sample': sample,
    }


def main():
    if len(sys.argv) < 3: print(__doc__); sys.exit(1)
    inp, outdir = os.path.expanduser(sys.argv[1]), os.path.expanduser(sys.argv[2])
    labels = sys.argv[3:] or ['6.0', '5.0', '4.0']
    os.makedirs(outdir, exist_ok=True)
    canon = load_canon()
    json.dump(canon, open(os.path.join(outdir, 'canon_used.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    res = []
    for lab in labels:
        r = convert_version(lab, os.path.join(inp, lab), canon, outdir)
        if r: res.append(r); print(lab, 'Excel', r['excel_rows'], '→ 轉出', r['written'], '｜未對應', sum(sum(c.values()) for c in r['unmapped'].values()), '｜', r['time_from'][:10], '~', r['time_to'][:10])
    json.dump(res, open(os.path.join(outdir, 'preview.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)


if __name__ == '__main__':
    main()
