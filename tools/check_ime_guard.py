#!/usr/bin/env python3
"""檢查全站儀表板的中文輸入保護（IME-GUARD）是否都在、而且一字不差。

由來：2026-10-17 經營者：注音／倉頡選字時輸入框被程式重畫打斷、打不出中文，
「已經發生在不同儀表板很多次，一定要解決」。修法＝每頁 <head> 最前面放同一段 IME-GUARD
（選字中攔下 input／keydown／keyup／keypress，選完字補送一次 input；window.__ime.on＝正在選字）。

正本：static/dashboard-recipe.html 的那一段（第一個做的頁面）。
檢查：
  1. 下面 REQUIRED 每一頁都有 IME-GUARD，剛好一段，與正本一字不差（md5 相同）
  2. IME-GUARD 是該頁第一個 <script>（要比 AUTH-CORE 與頁面自己的程式早載入，才攔得到）
  3. static/ 裡新冒出來、不在 REQUIRED 也不在 SKIP 的頁面，只要有可打字的輸入框就列為 FAIL（提醒新頁也要加）
  4. SKIP 裡（訓練站、歷史頁）若哪天出現可打字的輸入框，列 WARN 提醒

用法：python3 tools/check_ime_guard.py [static 資料夾路徑]
結果：最後一行 RESULT: PASS / FAIL；FAIL 時 exit code 1。
"""
import hashlib
import os
import re
import sys

ROOT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'static')
ROOT = os.path.abspath(ROOT)
SOURCE = 'dashboard-recipe.html'

# 要有 IME-GUARD 的頁面（正本除外）
REQUIRED = [
    'hub.html', 'decision-center.html', 'dashboard-monthly-review.html', 'store-today.html',
    'dashboard-announce.html', 'dashboard-pnl.html', 'pos-dashboard.html', 'zijiren.html',
    'google-reviews.html', 'dashboard-schedule.html', 'dashboard-purchase.html',
    'dashboard-reservation.html', 'recipe-upload.html', 'project-hub.html', 'admin-auth.html',
    'recipe-studio.html', 'index.html',   # 訓練站首頁（核心檔；2026-10-18 經營者同意加）
    'dashboard-supplier.html',   # 新品供應商查詢（2026-10-19 新建）
    'tablet-recipe.html', 'tablet-recipe-review.html',   # 平板外語食譜＋對照核准（2026-10-24 新建）
    'petty-cash.html',   # 零用金報銷（2026-10-24 新建）
    'schedule-workbench.html',   # 門市排班工作台（2026-10-26 新建；對話框要能打中文）
    'survey.html', 'dashboard-survey.html',   # 顧客問卷（2026-10-28 新建；留言題／留言搜尋要能打中文）
    'poll.html', 'dashboard-poll.html',   # 門市內部調查＋門市問卷管理（2026-10-30 新建；填寫、編輯題目要能打中文）
    'store-notes.html',   # 門市雜記（2026-10-30 新建；決策中心 📝 分頁，記事、搜尋要能打中文）
]

# 不加的頁面與原因
SKIP = {
    # 廢棄檔（不改）
    'pnl-dashboard.html': '廢棄', 'budget-2026.html': '廢棄', 'dashboard-vendor-price.html': '廢棄',
    'dashboard-recipe-cost.html': '廢棄', 'dashboard_m4.1.html': '廢棄',
    # 訓練站（核心檔或訓練站頁面，要改先問經營者）
    'self-dashboard.html': '訓練站（只有日期框）',
    'training-guide.html': '訓練站教學頁（輸入框都是唯讀示意）',
    # 歷史整合入口（沒有輸入框）
    'dashboard.html': '舊整合入口（沒有輸入框）',
}
SKIP_PATTERNS = [r'^pos-dashboard-backup-.*\.html$', r'.*-v2\.html$']   # 廢棄檔

GUARD_RE = re.compile(r'<script>\s*/\* ===== IME-GUARD.*?</script>', re.S)
# 可打字的輸入框：textarea，或 type 不是勾選／按鈕／檔案／日期／數字／密碼等、也不是唯讀的 input
TEXT_INPUT_RE = re.compile(
    r'<textarea\b|<input\b(?![^>]*\btype\s*=\s*["\']?(?:checkbox|radio|button|submit|file|range|color|date|time|'
    r'datetime-local|month|week|hidden|image|reset|number|password)\b)(?![^>]*\breadonly\b)', re.I)


def md5(s):
    return hashlib.md5(s.encode('utf-8')).hexdigest()


def read(name):
    with open(os.path.join(ROOT, name), encoding='utf-8') as fh:
        return fh.read()


def skipped(name):
    return name in SKIP or any(re.match(p, name) for p in SKIP_PATTERNS)


ok = True
src = read(SOURCE)
ref_list = GUARD_RE.findall(src)
if len(ref_list) != 1:
    print(f'FAIL 正本 {SOURCE}：IME-GUARD 有 {len(ref_list)} 段（應為 1）')
    sys.exit(1)
ref = ref_list[0]
print(f'正本 {SOURCE}：{len(ref)} 字  md5={md5(ref)}')
print()
print(f'{"頁面":34s} {"段數":>4s}  {"md5":32s}  結果')

for name in [SOURCE] + REQUIRED:
    path = os.path.join(ROOT, name)
    if not os.path.exists(path):
        print(f'{name:34s} {"-":>4s}  {"":32s}  FAIL 找不到檔案')
        ok = False
        continue
    s = read(name)
    found = GUARD_RE.findall(s)
    no_comment = re.sub(r'<!--.*?-->', lambda m: ' ' * len(m.group(0)), s, flags=re.S)   # 註解裡寫的 <script> 不算
    first_script = re.search(r'<script\b[^>]*>', no_comment)
    is_first = bool(found) and first_script is not None and s.find(found[0]) == first_script.start()
    if len(found) != 1:
        res, good = f'FAIL 段數 {len(found)}（應為 1）', False
    elif found[0] != ref:
        res, good = 'FAIL 內容與正本不同', False
    elif not is_first:
        res, good = 'FAIL 不是第一個 <script>（要放在 <head> 最前面）', False
    else:
        res, good = 'PASS', True
    ok &= good
    print(f'{name:34s} {len(found):4d}  {md5(found[0]) if found else "":32s}  {res}')

print()
others = sorted(f for f in os.listdir(ROOT) if f.endswith('.html') and f != SOURCE and f not in REQUIRED)
for name in others:
    s = read(name)
    n_text = len(TEXT_INPUT_RE.findall(s))
    if skipped(name):
        reason = SKIP.get(name, '廢棄')
        if n_text and reason != '廢棄' and not GUARD_RE.search(s):
            print(f'WARN {name}：不在檢查清單（{reason}），但有 {n_text} 個可打字的輸入框 → 要不要加 IME-GUARD 請先問經營者')
        else:
            print(f'略過 {name}（{reason}）')
        continue
    if n_text and not GUARD_RE.search(s):
        print(f'FAIL {name}：新頁面有 {n_text} 個可打字的輸入框但沒有 IME-GUARD → 照 dashboard-recipe.html 那段加在 <head> 最前面，並把檔名加進本腳本 REQUIRED')
        ok = False
    else:
        print(f'注意 {name}：不在 REQUIRED 清單（{"已有 IME-GUARD" if GUARD_RE.search(s) else "沒有可打字的輸入框"}）→ 建議加進 REQUIRED')

print()
print('RESULT:', 'PASS' if ok else 'FAIL')
sys.exit(0 if ok else 1)
