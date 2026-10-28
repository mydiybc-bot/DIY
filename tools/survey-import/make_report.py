#!/usr/bin/env python3
"""產生「SurveyCake 歷史回覆匯入結果」Excel 報表（給經營者看）。
  python3 -I tools/survey-import/make_report.py <convert 輸出資料夾> <報表.xlsx> [--live <線上筆數 json>]"""
import sys, os, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

out = os.path.expanduser(sys.argv[1]); dest = os.path.expanduser(sys.argv[2])
live = {}
if '--live' in sys.argv: live = json.load(open(sys.argv[sys.argv.index('--live') + 1], encoding='utf-8'))
pv = json.load(open(os.path.join(out, 'preview.json'), encoding='utf-8'))
rw = json.load(open(os.path.join(out, 'raw_preview.json'), encoding='utf-8'))
BACKEND = {'6.0': 33527, '5.0': 21272, '4.0': 5358, '3.0': 7003, '2.0': 15221, '1.0_第一版': 44594, '年度問卷_2020': 2676}
RAWNAME = {'1.0_第一版': 'raw_v1_第一版', '2.0': 'raw_v2', '3.0': 'raw_v3', '4.0': 'raw_v4', '年度問卷_2020': 'raw_annual2020'}
wb = Workbook(); H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='B4532A'); W = Alignment(wrap_text=True, vertical='top')

def sheet(title, head, rows, widths=None):
    ws = wb.create_sheet(title); ws.append(head)
    for c in ws[1]: c.font = H; c.fill = HF; c.alignment = W
    for r in rows: ws.append(r)
    for i, w in enumerate(widths or [], 1): ws.column_dimensions[get_column_letter(i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row: c.alignment = W
    ws.freeze_panes = 'A2'; return ws

wb.remove(wb.active)
rows = []
for r in pv:
    rows.append([r['label'], BACKEND.get(r['label']), r['excel_rows'], r['written'], '✔ 一致' if BACKEND.get(r['label']) == r['excel_rows'] == r['written'] else '✘ 不一致',
                 '匯進 fact_response（儀表板看得到）', r['time_from'][:10] + ' ～ ' + r['time_to'][:10]])
for r in rw:
    rows.append([r['label'], BACKEND.get(r['label']), r['rows'], r['rows'], '✔ 一致' if BACKEND.get(r['label']) == r['rows'] else '✘ 不一致',
                 '原樣封存：' + RAWNAME.get(r['label'], '?') + '（' + str(r['cols']) + ' 欄，拿掉客人 IP）', ''])
sheet('① 筆數對帳', ['問卷', 'SurveyCake 後台筆數', '匯出檔筆數', '轉出／封存筆數', '對帳', '去向', '期間'], rows, [18, 16, 14, 16, 10, 52, 26])
if live:
    ws = wb['① 筆數對帳']; ws.append([]); ws.append(['線上資料庫現況（匯入後）']); ws.cell(ws.max_row, 1).font = Font(bold=True)
    for k, v in live.items(): ws.append([k, v])

rows = []
for r in pv:
    for h, q in r['headers_mapped'].items(): rows.append([r['label'], q, h])
sheet('② 欄位對應', ['問卷', '新問卷題號', 'SurveyCake 題目文字'], rows, [12, 14, 80])

rows = []
for r in pv:
    for h in r['headers_unmapped']: rows.append([r['label'], h, '沒有對應的新題目；整張原樣收在封存表 ' + RAWNAME.get(r['label'], '')])
sheet('③ 沒對應的題目', ['問卷', 'SurveyCake 題目文字', '處理'], rows, [12, 60, 60])

rows = []
for r in pv:
    for q, d in r['syn_used'].items():
        for k, n in sorted(d.items(), key=lambda kv: -kv[1]): rows.append([r['label'], q, k, n, '舊選項改名，視為同一個選項'])
    for q, d in r['legacy_kept'].items():
        for k, n in sorted(d.items(), key=lambda kv: -kv[1]): rows.append([r['label'], q, k, n, '舊選項和新問卷不同，保留原文（儀表板列在新選項後面）'])
sheet('④ 選項改名與保留原文', ['問卷', '題號', '舊選項 → 新選項／保留', '人次', '說明'], rows, [12, 8, 70, 10, 50])

stores = sorted({s for r in pv for s in r['stores']})
rows = [[s] + [r['stores'].get(s, 0) for r in pv] + [sum(r['stores'].get(s, 0) for r in pv)] for s in stores]
rows.append(['合計'] + [r['written'] for r in pv] + [sum(r['written'] for r in pv)])
sheet('⑤ 店別筆數', ['店別'] + [r['label'] for r in pv] + ['合計'], rows, [34, 10, 10, 10, 10, 10])

months = sorted({m for r in pv for m in r['months']})
rows = [[m] + [r['months'].get(m, 0) for r in pv] + [sum(r['months'].get(m, 0) for r in pv)] for m in months]
sheet('⑥ 月別筆數', ['月份'] + [r['label'] for r in pv] + ['合計'], rows, [12, 10, 10, 10, 10, 10])

rows = []
for r in pv:
    for sm in r['sample']:
        rows.append([r['label'], ' ｜ '.join(str(x) for x in sm['raw'] if x), json.dumps(sm['new'], ensure_ascii=False)])
sheet('⑦ 抽樣對照', ['問卷', 'SurveyCake 原始（前 16 欄）', '轉成新格式後'], rows, [10, 70, 90])

notes = [
 ['時間', '填答時間照 SurveyCake 的時間（台北時區）寫入；編號＝sc＋版本＋SurveyCake 的 Hash，重跑不會重複。'],
 ['填答秒數', '舊問卷的「填答秒數」是從食譜頁開啟到送出（約 2～3 小時），不是填問卷的時間；原值留在 fact_response 的 sec 欄，儀表板「平均填答時間」只算 1 小時內的，所以舊資料顯示「－」。'],
 ['店別', '桃園中正藝文店＝桃園藝文；台南focus店＝台南Focus；台南新光西門店已結束，店號記 0、店名「台南新光西門（已結束）」，儀表板篩選有獨立選項。'],
 ['選項改名', '4.0／3.0／5.0 的「難易度」→「不容易失敗」、「流行性」→「現在流行」（6.0 改名）；2.0 的「其它」→「其他」。「其他-自由填」歸「其他」。'],
 ['協助的地方（q10）', '5.0 是 12 個細項、6.0 合併成 5 個選項，兩邊不硬對，5.0 保留原文。'],
 ['4.0 的 q9', '4.0 問「是否遇到需要協助」，6.0／5.0 問「服務人員是否有提供協助」，意思相反方向，所以 4.0 這題不放進 q9（原文在 raw_v4）。'],
 ['5 分制量表', '1.0、2.0 的滿意度是「非常滿意～非常不滿意」5 分制，跟新問卷「是／否」不能比，所以整張原樣封存，不進儀表板。'],
 ['隱私', '封存表拿掉「IP 紀錄」欄（客人 IP，用不到）。'],
 ['還原', '備份分頁 backup_*_20261028-1236；編輯器執行 rollbackImport 會拿掉所有匯入的回覆並重建彙總；封存表 raw_* 手動刪分頁即可。'],
]
sheet('⑧ 注意事項', ['項目', '說明'], notes, [20, 120])
os.makedirs(os.path.dirname(dest), exist_ok=True); wb.save(dest); print('報表已存：', dest)
