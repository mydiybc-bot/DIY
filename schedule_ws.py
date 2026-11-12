"""🗓 門市排班工作台（2026-10-26）

店長上傳「每月排班設定表」＋「104 匯入範本（時間排班）」→ Claude（Sonnet 5.5）排初稿 → 檢查程式（schedule_check104）把關，
沒過自動退回 Claude 修正（2026-11-04 起修正一律用 Opus 5.5）→ 網頁日曆可點格子改、也可在對話框用講的改（2026-11-04 起對話用 Opus 5.5）→ 下載填好的 104 匯入檔。

2026-10-28 經營者裁定：設定表改成直接在工作台填（Excel 上傳留著當備用）——
  ・店別基本設定（班別時段、每天需求人數、連上上限、晚接早、跨店夥伴、店規、挪假未補清單）：每店一份 store-<店號>.json，填一次每月自動帶入。
  ・本月設定（正職基準、營收預估、每位夥伴正職／兼職與目標時數、指休／特休／兼職可上日點格子、特殊日）：存在該月 JSON 的 mset。
  ・兩者組成跟 Excel 一樣的文字給 AI（form_text），AI 那段不用改；表單同時先算出 ext（班別、需求、員工），104 同步進來的班表也能直接檢查、顯示班別名。
  ・班表加「營業額」（排班分析 GAS getStoreDailyData，含甜點券$100 加回）與「日獎金」列（口徑同排班儀表板：級距、跨級 ×2、每人 200×工時÷8），
    另可下載「各店日獎金表」給會計（bonus_export；總部一次全店、店長只有自己店）。

server.py 只把 /api/schedule/* 整個交給 handle()；身分用權限中心 store_N（店）／task_hq（總部）驗證，結果快取 10 分鐘（只存雜湊）。
每店每月一份 JSON，存在 Render Disk：<TRAINING_DATA_DIR>/schedule_ws/<店號>-<YYYYMM>.json。
"""
from __future__ import annotations

import asyncio
import base64
import copy
import datetime as dt
import hashlib
import io
import json
import os
import re
import ssl
import threading
import time
import unicodedata
import urllib.parse
import urllib.request
import warnings

import openpyxl

import schedule_check104 as c104
import schedule_solver
from storage_paths import atomic_write_json, data_dir

warnings.filterwarnings("ignore")

AUTH_API = "https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec"
AUTH_TTL = 10 * 60
MODEL = "claude-sonnet-5-5"
MODEL_BACKUP = "claude-opus-5-5"
# 2026-11-04 經營者「節省店長的時間」：初稿留 Sonnet（10/26 實測 3 分鐘、0 警告；Opus 要 7 分鐘、2 警告），
# 初稿沒過檢查第 1 次就換 Opus 修；對話改班用 Opus（實測 10/10 直接改好，Sonnet 約 8/10 會只給選項、店長要再講一次）
CHAT_MODEL = MODEL_BACKUP
PRICE = {MODEL: (2.00, 10.00, 2.50, 0.20), MODEL_BACKUP: (4.00, 20.00, 5.00, 0.20)}   # 美元／百萬 tokens：輸入、輸出、快取寫、快取讀
TWD = 32
STORES = {1: "精明", 2: "草悟道", 3: "南京", 4: "士林", 5: "Focus", 6: "新竹", 7: "板橋", 8: "新店",
          9: "中壢", 10: "藝文", 11: "A13", 12: "SKM"}
# 2026-11-05 經營者：開 13 號測試店試對話改班，測完會刪。只有總部登入看得到；第一次打開某月時從 12 號店同月複製一份；
# 104 同步、全店日獎金表、跨店支援都不算它；營業額照 12 號店讀。
TEST_STORES = {13: "測試店", 14: "測試店二"}
TEST_SRC = 12
# 2026-11-06 經營者「做測試區」：14 號測試店二（假名字夥伴、自己的設定），跟 13 號組成「測試區」試區域排班；只有 13 號會從 12 號店複製
TEST_COPY = {13}


def _sname(store):
    return STORES.get(store) or TEST_STORES.get(store, "")
WEEKDAY = "一二三四五六日"
# 104 崗位（13 個「X店」）→ 店號；台南店＝5 Focus、高雄店＝12 SKM（2026-10-28 跨店支援日獎金歸被支援店用）
POS_STORE = {"精明店": 1, "草悟道店": 2, "南京店": 3, "士林店": 4, "台南店": 5, "新竹店": 6, "板橋店": 7, "新店店": 8,
             "中壢店": 9, "藝文店": 10, "A13店": 11, "高雄店": 12}
RULES = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "schedule_rules.md"), encoding="utf-8").read()
MAX_FILE_B64 = 4_000_000
MAX_SETTINGS_CHARS = 80_000
MAX_API_MESSAGES = 60

# 營業額／日獎金（2026-10-28）：營收用排班分析 GAS（fact_daily_sales＋甜點券$100），口徑見 dashboard-schedule skill「🎁 日獎金口徑」
REV_API = "https://script.google.com/macros/s/AKfycbwTE1H-uh6mtsoZP_HKscF8XGgpv8Mxu6lH1M8Xrbp0kjd8M2OR1sUtPZrnSVBAJkF6Og/exec"
REV_TTL = 2 * 3600

# 行政院人事行政總處「政府行政機關辦公日曆表」114～116 年（2025～2027，2026-10-31 從 dgpa.gov.tw 附件逐格讀放假底色）
# TW_OFF＝落在週一～五的放假日（國定假日、補假、春節）；TW_WORK＝週末補行上班日。新年度公告後照同樣方式補一行。
TW_OFF = {
    "2025": "01-01 01-27 01-28 01-29 01-30 01-31 02-28 04-03 04-04 05-30 09-29 10-06 10-10 10-24 12-25",
    "2026": "01-01 02-16 02-17 02-18 02-19 02-20 02-27 04-03 04-06 05-01 06-19 09-25 09-28 10-09 10-26 12-25",
    "2027": "01-01 02-04 02-05 02-08 02-09 02-10 03-01 04-05 04-06 04-30 06-09 09-15 09-28 10-11 10-25 12-24 12-31",
}
TW_WORK = {
    "2025": "02-08",
}
REV_EST_RETRY = 3600
BONUS_TIERS = [(10, 10000), (20, 14000), (30, 29000), (40, 42000), (50, 53000), (60, 65000), (70, 77000), (80, 88000), (90, 100000), (100, 110000)]
BONUS_BASE, BONUS_STD_H = 200, 8
BONUS_LEAVE_RE = re.compile(r"特休|病假|事假|喪假|婚假|產假|陪產|公假|補休|休假|月休|例假|休息日|指休|生理假|家庭照顧假|請假|天災")
BONUS_MEET_RE = re.compile(r"會議|區會|上課|內訓|培訓|教育訓練|研習|經管會")
BONUS_NOFLOOR_RE = re.compile(r"不計現場人力|不算現場人力")
DEFAULT_SHIFTS = [  # 新店第一次打開表單的預設列（店長自己改）
    {"name": "平日早班", "segs": [["10:20", "12:00"], ["12:30", "18:50"]], "wd": 1, "we": 0},
    {"name": "平日晚班", "segs": [["13:00", "15:00"], ["15:30", "21:30"]], "wd": 1, "we": 0},
    {"name": "假日早班", "segs": [["09:50", "12:00"], ["12:30", "18:20"]], "wd": 0, "we": 1},
    {"name": "假日中班", "segs": [["12:00", "15:00"], ["15:30", "20:30"]], "wd": 0, "we": 1},
    {"name": "假日晚班", "segs": [["13:30", "15:30"], ["16:00", "22:00"]], "wd": 0, "we": 1},
]
_rev_jobs: set = set()

_auth_ok: dict[str, float] = {}
_locks: dict[str, threading.Lock] = {}
_jobs: dict[str, dict] = {}
_glock = threading.Lock()
_client = None


# ───────────────────────── 共用 ─────────────────────────

def _client_():
    global _client
    if _client is None:
        import anthropic
        _client = anthropic.Anthropic(max_retries=3)   # 金鑰讀環境變數 ANTHROPIC_API_KEY
    return _client


def _dir():
    p = data_dir() / "schedule_ws"
    p.mkdir(parents=True, exist_ok=True)
    return p


def _key(store, ym):
    return f"{int(store)}-{ym}"


def _lock(k):
    with _glock:
        return _locks.setdefault(k, threading.Lock())


def _load(store, ym):
    p = _dir() / f"{_key(store, ym)}.json"
    if p.exists():
        return json.loads(p.read_text(encoding="utf-8"))
    src = _dir() / f"{_key(TEST_SRC, ym)}.json"
    if int(store) in TEST_COPY and src.exists():
        st = json.loads(src.read_text(encoding="utf-8"))
        st.update(store=int(store), history=[], api_messages=[], usage=[], s104=None, s104_adopted=0,
                  chat=[{"role": "ai", "ts": time.time(), "text": f"🧪 這是測試店：班表從 {TEST_SRC} 號店 {int(ym[4:])} 月複製過來，怎麼改都不會影響正式店和 104。"}])
        return st
    return {"store": int(store), "ym": ym, "settings_files": [], "note": "", "template": None,
            "ext": None, "schedule": None, "report": "", "history": [], "chat": [], "api_messages": [],
            "check": None, "usage": [], "updated": ""}


def _auto_report(rep):
    """這份說明是不是程式寫的（空白、舊版「方案 A／B」、新版「排班結果／目前班表」）→ 是的話每次存檔照目前班表重寫；AI 寫的四段報告不動"""
    rep = (rep or "").lstrip()
    return not rep or rep.startswith(("### 方案", "### 排班結果", "### 目前班表"))


def _save(st):
    st["updated"] = dt.datetime.now().isoformat(timespec="seconds")
    if st.get("schedule") and st.get("check") and (st.get("draft") or _auto_report(st.get("report"))):   # 2026-11-10：說明跟著目前班表重寫（改班、對話、復原後都一致）
        try:
            if "cover" not in st["check"]:
                st["check"] = run_check(st)
            st["report"] = live_report(st)
        except Exception:   # noqa: BLE001 — 說明寫不出來不可以擋住存檔
            pass
    atomic_write_json(_dir() / f"{_key(st['store'], st['ym'])}.json", st)


def _cost(model, u):
    p = PRICE.get(model, PRICE[MODEL])
    cw = getattr(u, "cache_creation_input_tokens", 0) or 0
    cr = getattr(u, "cache_read_input_tokens", 0) or 0
    usd = (u.input_tokens * p[0] + u.output_tokens * p[1] + cw * p[2] + cr * p[3]) / 1e6
    return {"model": model, "in": u.input_tokens, "out": u.output_tokens, "cw": cw, "cr": cr,
            "twd": round(usd * TWD, 2)}


def _log_usage(st, kind, c):
    c = dict(c, kind=kind, ts=dt.datetime.now().isoformat(timespec="seconds"))
    st.setdefault("usage", []).append(c)
    with open(_dir() / "usage.jsonl", "a", encoding="utf-8") as f:
        f.write(json.dumps(dict(c, store=st["store"], ym=st["ym"]), ensure_ascii=False) + "\n")


def _blocks(content):
    """回應的 content 轉成可存檔、可原樣送回的 dict（thinking 區塊要原封不動送回）。"""
    return [b.model_dump(mode="json", exclude_none=True) for b in content]


# ───────────────────────── 身分 ─────────────────────────

def verify(store, pwd):
    """store＝1～12 或 'hq'；密碼對 → (True, '')"""
    pwd = str(pwd or "")
    if not pwd:
        return False, "請輸入密碼"
    if str(store) == "hq":
        dash = "task_hq"
    else:
        try:
            n = int(store)
        except (TypeError, ValueError):
            return False, "請選擇門市"
        if n not in STORES:
            return False, "請選擇門市"
        dash = f"store_{n}"
    h = hashlib.sha256((dash + "\x00" + pwd).encode("utf-8")).hexdigest()
    now = time.time()
    with _glock:
        if _auth_ok.get(h, 0) > now:
            return True, ""
    url = AUTH_API + "?" + urllib.parse.urlencode({"action": "verify", "dashboard": dash, "pwd": pwd, "callback": "cb"})
    last = "連不到權限中心，請再試一次"
    for _ in range(3):
        try:
            with urllib.request.urlopen(url, timeout=25) as r:
                t = r.read().decode("utf-8")
        except Exception:
            continue
        m = re.search(r"\((\{.*\})\)", t, re.S)
        if not m:
            continue
        d = json.loads(m.group(1))
        if d.get("ok"):
            with _glock:
                _auth_ok[h] = now + AUTH_TTL
            return True, ""
        return False, str(d.get("error") or "密碼錯誤")
    return False, last


def _can(store_req, store_auth):
    """總部可看任何一店；店只能看自己"""
    return str(store_auth) == "hq" or str(store_auth) == str(store_req)


# ───────────────────────── 檔案讀取 ─────────────────────────

def _cell_text(v):
    if v is None:
        return ""
    if isinstance(v, dt.datetime):
        return v.strftime("%Y-%m-%d %H:%M") if (v.hour or v.minute) else v.strftime("%Y-%m-%d")
    if isinstance(v, (dt.date, dt.time)):
        return v.isoformat()
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def settings_to_text(name, raw):
    """設定表（xlsx／xls／csv／txt）→ 給 Claude 看的純文字，每個分頁一段"""
    low = name.lower()
    out = []
    if low.endswith((".xlsx", ".xlsm")):
        wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
        for ws in wb.worksheets:
            if ws.sheet_state != "visible":
                continue
            out.append(f"### 分頁「{ws.title}」")
            for row in ws.iter_rows(values_only=True):
                cells = [_cell_text(v) for v in row]
                while cells and not cells[-1]:
                    cells.pop()
                if any(cells):
                    out.append("\t".join(cells))
    elif low.endswith(".xls"):
        import xlrd
        bk = xlrd.open_workbook(file_contents=raw)
        for sh in bk.sheets():
            out.append(f"### 分頁「{sh.name}」")
            for r in range(sh.nrows):
                cells = []
                for c in range(sh.ncols):
                    cell = sh.cell(r, c)
                    if cell.ctype == xlrd.XL_CELL_DATE:
                        cells.append(_cell_text(xlrd.xldate.xldate_as_datetime(cell.value, bk.datemode)))
                    else:
                        cells.append(_cell_text(cell.value))
                while cells and not cells[-1]:
                    cells.pop()
                if any(cells):
                    out.append("\t".join(cells))
    else:
        out.append(raw.decode("utf-8", errors="replace"))
    return "\n".join(out)[:MAX_SETTINGS_CHARS]


def template_info(raw):
    """104 範本 → 班表名稱、日期、員工、合法清單"""
    wb = openpyxl.load_workbook(io.BytesIO(raw))
    ws = wb.worksheets[0]
    hdr, days, emps = c104.layout(ws)
    if not days or not emps:
        raise ValueError("看不懂這個檔：找不到「員工編號」標題列或日期欄。請上傳 104 下載的「時間排班」匯入範本。")
    lists = c104.read_lists(wb)
    if not lists["dayTypes"] or not lists["timeStrs"]:
        raise ValueError("範本缺少 104 的下拉選單清單（hidden 分頁），請重新從 104 下載當月範本。")
    return {
        "title": ws.title,
        "days": [{"d": int(d["label"][3:5]), "m": int(d["label"][:2]), "label": d["label"], "wd": d["wd"]} for d in days],
        "emps": [{"id": e["id"], "name": e["name"]} for e in emps],
        "tags": sorted(lists["tagStrs"]), "positions": sorted(lists["positionStrs"]),
        "dayTypes": sorted(lists["dayTypes"]),
    }


# ───────────────────────── 格子文法 ─────────────────────────

SEG_RE = re.compile(r"^(\d{1,2}:\d{2})-(次日)?(\d{1,2}:\d{2})((?:@[^@#]+)?)((?:#[^@#]+)?)$")


def parse_cell(v, shifts, day_types):
    """格子 → {type, segs:[[起,迄,標籤,崗位]], dtag}；寫錯丟 ValueError"""
    v = unicodedata.normalize("NFKC", v or "").strip()   # 店長在日曆打全形「１１：００－１５：００」也收
    out = {"type": "", "segs": [], "dtag": ""}
    m = re.match(r"^\[([^\]]+)\](.*)$", v)
    if m:
        out["dtag"], v = m.group(1).strip(), m.group(2).strip()
    if not v:
        return out
    if v in day_types:
        out["type"] = v
        return out
    m = re.match(r"^([^:：0-9]+)[:：](.+)$", v)
    if m and m.group(1).strip() in day_types:
        out["type"], v = m.group(1).strip(), m.group(2).strip()
    if v not in shifts and v and not v[0].isdigit():   # 班別名稱有全形字（「早班（短）」）：格子已轉半形，名稱也轉了再比
        v = next((k for k in shifts if unicodedata.normalize("NFKC", k) == v), v)
    if v in shifts:
        out["segs"] = [[a, b, "", ""] for a, b in shifts[v]]
        return out
    for part in v.split(","):
        part = part.strip().replace("：", ":")
        mm = SEG_RE.match(part)
        if not mm:
            raise ValueError(f"看不懂「{v}」：班別名稱要跟設定一樣，自訂時段要寫成 10:20-12:00")
        s, e = mm.group(1).zfill(5), ("次日" if mm.group(2) else "") + mm.group(3).zfill(5)
        out["segs"].append([s, e, mm.group(4)[1:].strip(), mm.group(5)[1:].strip()])
    if len(out["segs"]) > 4:
        raise ValueError(f"「{v}」超過 4 段，104 一天最多 4 段")
    return out


def _mins(t):
    return c104.mins(t)


def _elsewhere(pos, title):
    """崗位填了別家店 → 那段人在別店（104 規則：在本店上班崗位留空；只有實際所在店≠本班表的店才填崗位）。
    「支援」標籤本身不代表在別店：外店夥伴來本店支援時，在本店的班表上就是「支援」＋崗位空白（2026-10-28 用高雄SKM 10/3 張于歡實例確認）"""
    pos = (pos or "").strip()
    if not pos:
        return False
    return re.sub(r"店$", "", pos) not in (title or "")   # 只去結尾的「店」：「新店店」原本變成「新」，新竹店的班表會誤判成本店


def _pos_store(pos):
    """崗位字串 → 店號（對不上回 None）"""
    pos = (pos or "").strip()
    if pos in POS_STORE:
        return POS_STORE[pos]
    for k, v in POS_STORE.items():
        if re.sub(r"店$", "", k) and re.sub(r"店$", "", k) in pos:
            return v
    return None


def cell_hours(c):
    return sum(_mins(e) - _mins(s) for s, e, _, _ in c["segs"]) / 60


# ───────────────────────── 設定表單（2026-10-28：取代 Excel 設定表） ─────────────────────────

TIME_RE = re.compile(r"^\d{1,2}:\d{2}$")


def _hhmm(v):
    # 全形「１０：００」先轉半形（2026-11-05 模擬：原本存成「10:００」，排班、匯入 104 都會壞）
    v = unicodedata.normalize("NFKC", str(v or "")).strip()
    if not v:
        return ""
    # 沒有冒號一定要 4 位數（1030）；3 位數「100」分不出是 1:00 還是少打一個 0 的 10:00（原本讀成 01:00）
    m = re.match(r"^([0-9]{1,2}):([0-9]{2})$", v) or re.match(r"^([0-9]{2})([0-9]{2})$", v)
    if not m or int(m.group(1)) > 23 or int(m.group(2)) > 59:
        raise ValueError(f"時間「{v}」要寫成 10:20")
    if int(m.group(2)) % 5:
        raise ValueError(f"時間「{v}」104 只收 5 分鐘一格（例：10:20、10:25）")
    return f"{int(m.group(1)):02d}:{m.group(2)}"


def _num(v, default=None, lo=None, hi=None, what="數字"):
    if v is None or str(v).strip() == "":
        return default
    try:
        x = float(v)
    except (TypeError, ValueError):
        raise ValueError(f"{what}「{v}」不是數字")
    if (lo is not None and x < lo) or (hi is not None and x > hi):
        raise ValueError(f"{what}「{v}」超出範圍")
    return int(x) if x.is_integer() else x


def _cfg_path(store):
    return _dir() / f"store-{int(store)}.json"


# 2026-11-11 經營者：「連續上班最多幾天」「晚班接隔天早班最多幾次」不再讓店長填，一律照勞基法：
# 七休一＝連續上班最多 6 天；晚班接早班不限次數，只要下班到隔天上班有 11 小時（不到 11 小時照舊提醒）
# 2026-11-12 經營者：公司規定連續上班最多 5 天（正職、兼職都適用；跨店合併算）。排初稿照 5 天排；
# 店長手動改成連上 6 天只提醒不擋（6 天仍合法），7 天以上照舊算錯誤（七休一）
LAW_MAX_CONSEC = 5


def _law(cfg):
    cfg["max_consec"] = LAW_MAX_CONSEC
    cfg["late_early_limit"] = None
    return cfg


def cfg_load(store):
    p = _cfg_path(store)
    if int(store) in TEST_COPY and not p.exists():
        p = _cfg_path(TEST_SRC)
    if p.exists():
        try:
            return _law(json.loads(p.read_text(encoding="utf-8")))
        except (OSError, json.JSONDecodeError):
            pass
    return _law({"shifts": [], "support": "", "rules": "", "carry": [], "updated": ""})


def _t(m):
    return f"{m // 60:02d}:{m % 60:02d}"


def split_span(start, end, rest=0, rest_at="", what="時段"):
    """2026-11-10 經營者：店長只填「上班、下班、休息幾分鐘」，系統自己拆成 104 收的分段（104 時間排班不扣休息，
    只能靠中間留空檔；空檔不算工時、不會變加班）。休息從幾點開始可指定；沒指定＝上班滿 4 小時後，
    剩下不到 4 小時又放不下就放正中間（5 分鐘一格）。回 [[起,迄], …]；寫錯丟 ValueError"""
    a, b = _hhmm(start), _hhmm(end)
    if not a or not b:
        raise ValueError(f"{what}要填上班和下班時間")
    ma, mb = _mins(a), _mins(b)
    if mb <= ma:
        raise ValueError(f"{what} {a}-{b} 下班要晚於上班")
    try:
        rest = int(float(unicodedata.normalize("NFKC", str(rest if rest not in (None, "") else 0)).strip()))
    except (ValueError, OverflowError):
        raise ValueError(f"{what}的休息要填分鐘數（例：30；不休息填 0）")
    if rest < 0 or rest % 5:
        raise ValueError(f"{what}的休息分鐘要是 5 的倍數（104 只收 5 分鐘一格）")
    if not rest:
        return [[a, b]]
    if mb - ma - rest < 10:
        raise ValueError(f"{what} {a}-{b} 扣掉休息 {rest} 分鐘就沒有上班時間了")
    at = _hhmm(rest_at)
    if at:
        r0 = _mins(at)
        if r0 <= ma or r0 + rest >= mb:
            raise ValueError(f"{what}的休息 {at} 起 {rest} 分鐘要落在 {a}-{b} 中間")
    else:
        r0 = ma + 240
        if r0 + rest > mb - 60:   # 休息後剩不到 1 小時 → 改放正中間
            r0 = ma + ((mb - ma - rest) // 2 + 4) // 5 * 5
    return [[a, _t(r0)], [_t(r0 + rest), b]]


def span_of(segs):
    """分段 → (上班, 下班, 休息分鐘, 休息從幾點)；一段班休息 0"""
    segs = [x for x in segs or [] if x and x[0] and x[1]]
    if not segs:
        return "", "", 0, ""
    rest = sum(_mins(segs[i][0]) - _mins(segs[i - 1][1]) for i in range(1, len(segs)))
    return segs[0][0], segs[-1][1], rest, (segs[0][1] if len(segs) > 1 and rest else "")


def span_txt(segs):
    """給人看的寫法：10:30–19:00（休 30 分）"""
    a, b, rest, _ = span_of(segs)
    return f"{a}–{b}" + (f"（休 {rest} 分）" if rest else "") if a else ""


def cell_from_span(sp):
    """改班視窗填的「上班／下班／休息」→ 格子文字（10:30-14:30,15:00-19:00）"""
    segs = split_span(sp.get("start"), sp.get("end"), sp.get("rest"), sp.get("rest_at") or "", "自訂時段")
    return ",".join(f"{a}-{b}" for a, b in segs)


def _carry_date(v, name=""):
    """挪假未補的日期一律存成西元 YYYY-MM-DD（2026-11-11 經營者）。沒寫年份的（11/3）程式分不出是哪一年，
    會被當成工作台以前的舊帳、跟班表上的挪假重複算一次 → 擋下請店長補年份；民國年自動換成西元"""
    v = unicodedata.normalize("NFKC", str(v or "")).strip()
    m = re.fullmatch(r"(\d{2,4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?", v)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if 100 <= y <= 200:
            y += 1911
        try:
            if y >= 2000:
                return dt.date(y, mo, d).isoformat()
        except ValueError:
            pass
    raise ValueError(f"挪假未補「{name}」的日期要寫西元年月日，例：2026-10-05（現在填的是「{v or '空白'}」）")


def cfg_clean(body):
    """店別基本設定：檢查＋整理（寫錯丟 ValueError）"""
    out = {"shifts": [], "carry": []}
    names = set()
    for r in (body.get("shifts") or [])[:20]:
        name = str(r.get("name") or "").strip()[:20]
        segs = []
        if r.get("start") or r.get("end"):   # 2026-11-10：新表單填「上班／下班／休息」→ 自動拆段
            r = dict(r, segs=split_span(r.get("start"), r.get("end"), r.get("rest"), r.get("rest_at") or "", f"班別「{name}」"))
        for a, b in (r.get("segs") or [])[:3]:
            a, b = _hhmm(a), _hhmm(b)
            if a and b:
                if _mins(b) <= _mins(a):
                    raise ValueError(f"班別「{name}」的時段 {a}-{b} 結束要晚於開始")
                if segs and _mins(a) < _mins(segs[-1][1]):
                    raise ValueError(f"班別「{name}」的時段 {a}-{b} 跟上一段 {segs[-1][0]}-{segs[-1][1]} 重疊或順序顛倒（104 不收）")
                segs.append([a, b])
            elif a or b:
                raise ValueError(f"班別「{name}」有一段只填了一半")
        if not name and not segs:
            continue
        if not name:
            raise ValueError("有一列班別沒有名稱")
        if not segs:
            raise ValueError(f"班別「{name}」至少要填一段時間")
        if name in names:
            raise ValueError(f"班別名稱「{name}」重複")
        names.add(name)
        out["shifts"].append({"name": name, "segs": segs, "wd": _num(r.get("wd"), 0, 0, 20, "平日人數"), "we": _num(r.get("we"), 0, 0, 20, "假日人數")})
    _law(out)   # 連上天數、晚接早照勞基法（不收店長填的值）
    out["support"] = str(body.get("support") or "").strip()[:1000]
    out["rules"] = str(body.get("rules") or "").strip()[:4000]
    for r in (body.get("carry") or [])[:60]:
        name = str(r.get("name") or "").strip()[:20]
        if not name:
            continue
        out["carry"].append({"name": name, "type": ("B" if str(r.get("type") or "").strip().upper() == "B" else "A"),
                             "date": _carry_date(r.get("date"), name), "note": str(r.get("note") or "").strip()[:100]})
    return out


def cfg_save(store, cfg):
    cfg["updated"] = dt.datetime.now().isoformat(timespec="seconds")
    atomic_write_json(_cfg_path(store), cfg)
    return cfg


def _prev_ym(ym):
    y, m = int(ym[:4]), int(ym[4:])
    return f"{y - 1}12" if m == 1 else f"{y}{m - 1:02d}"


def _tw_off(day):
    """政府辦公日曆的放假日（週末＋平日放假；扣掉週末補班）"""
    k = day.strftime("%m-%d")
    y = str(day.year)
    if k in TW_WORK.get(y, "").split():
        return False
    return day.weekday() >= 5 or k in TW_OFF.get(y, "").split()


def base_est(ym):
    """正職基準估算（2026-10-31 經營者：照行政院辦公日曆估，店長可改）：休假＝週末、國定假日＝平日放假、標準時數＝上班日×8"""
    y, m = int(ym[:4]), int(ym[4:])
    if str(y) not in TW_OFF:
        return None
    n = (dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)).day
    days = [dt.date(y, m, d) for d in range(1, n + 1)]
    wk = [d for d in days if d.weekday() >= 5 and d.strftime("%m-%d") not in TW_WORK.get(str(y), "").split()]
    hol = [d for d in days if d.weekday() < 5 and d.strftime("%m-%d") in TW_OFF[str(y)].split()]
    work = n - len(wk) - len(hol)
    src = f"行政院 {y} 年辦公日曆：{m} 月 {n} 天－週末 {len(wk)} 天－國定假日 {len(hol)} 天" + \
          (f"（{'、'.join(f'{d.month}/{d.day}' for d in hol)}）" if hol else "") + f"＝上班 {work} 天 × 8h"
    return {"std_hours": work * 8, "off_days": len(wk), "holidays": len(hol), "src": src}


def rev_est_fetch(store, ym):
    """營收預估（2026-10-31 經營者：照去年同月實際）：去年同月每天營業額，平日／假日（週末＋國定假日）各取平均，取到百位"""
    ly = f"{int(ym[:4]) - 1}{ym[4:]}"
    days = fetch_rev(store, ly)["days"]
    wd, we = [], []
    for d, v in days.items():
        (we if _tw_off(dt.date(int(ly[:4]), int(ly[4:]), int(d))) else wd).append(v["rev"])
    if len(wd) < 3 or len(we) < 2:
        return {"weekday": None, "weekend": None, "src": f"去年 {int(ly[4:])} 月營業資料不足（平日 {len(wd)} 天／假日 {len(we)} 天），請自己填", "fetched": time.time()}
    avg = lambda a: int(round(sum(a) / len(a), -2))
    return {"weekday": avg(wd), "weekend": avg(we), "fetched": time.time(),
            "src": f"去年 {ly[:4]} 年 {int(ly[4:])} 月 POS 實際：平日 {len(wd)} 天平均、假日 {len(we)} 天平均（含國定假日）"}


def rev_est_ensure(store, ym):
    """還沒算過（或上次失敗超過 1 小時）就抓一次去年同月；去年的數字不會變，算過就留著"""
    k = _key(store, ym)
    with _lock(k):
        st = _load(store, ym)
        re_ = st.get("rev_est") or {}
        if not st.get("template") or re_.get("fetched") or time.time() - (re_.get("tried") or 0) < REV_EST_RETRY:
            return
    with _glock:
        if ("est", k) in _rev_jobs:   # 別的請求正在抓，這次先不等
            return
        _rev_jobs.add(("est", k))
    try:
        try:
            data = rev_est_fetch(store, ym)
        except Exception as ex:   # noqa: BLE001
            data = {"error": f"{type(ex).__name__}: {str(ex)[:120]}", "tried": time.time()}
        with _lock(k):
            st = _load(store, ym)
            st["rev_est"] = data
            _save(st)
    finally:
        with _glock:
            _rev_jobs.discard(("est", k))


def mset_view(st):
    """本月設定（含預設值）：夥伴清單跟著 104 範本；正職／兼職、目標時數沒填的先抄上個月；
    正職基準、營收預估沒填的先放估算值（est 記哪些是估的，畫面標「估算」，店長改了存檔就是店長的）"""
    ms = copy.deepcopy(st.get("mset") or {})
    ms.setdefault("base", {})
    ms.setdefault("revenue", {})
    est = {"fields": []}
    be = base_est(st["ym"])
    if be:
        est["base_src"] = be["src"]
        for f in ("std_hours", "off_days", "holidays"):
            if ms["base"].get(f) is None:
                ms["base"][f] = be[f]
                est["fields"].append(f)
    rv = st.get("rev_est") or {}
    est["rev_src"] = rv.get("src") or ("去年同月營收讀取失敗：" + rv["error"] if rv.get("error") else "")
    for f in ("weekday", "weekend"):
        if ms["revenue"].get(f) is None and rv.get(f):
            ms["revenue"][f] = rv[f]
            est["fields"].append("rev_" + f)
    ms["est"] = est
    ms.setdefault("emps", {})
    ms.setdefault("marks", {})
    ms.setdefault("special", [])
    prev = None
    info = (st.get("template") or {}).get("info") or {}
    s104 = {e.get("name"): e for e in ((st.get("s104") or {}).get("emps") or [])}
    for e in info.get("emps") or []:
        cur = ms["emps"].get(e["id"])
        if cur is None:
            if prev is None:
                pp = _dir() / f"{_key(st['store'], _prev_ym(st['ym']))}.json"
                prev, prev_std = {}, None
                if pp.exists():
                    try:
                        pm = json.loads(pp.read_text(encoding="utf-8")).get("mset") or {}
                        prev = pm.get("emps") or {}
                        prev_std = (pm.get("base") or {}).get("std_hours")
                    except (OSError, json.JSONDecodeError):
                        prev = {}
                if prev_std is None:
                    pb = base_est(_prev_ym(st["ym"]))
                    prev_std = pb and pb["std_hours"]
            cur = dict(prev.get(e["id"]) or {"type": "正職", "target": None, "min_hours": None, "note": ""})
            for f in MARK_TXT.values():   # 日期是這個月的，不跟著抄
                cur.pop(f, None)
            # 2026-11-04 經營者：排 12 月卻看到 11 月的數字。正職目標時數＝上個月標準時數的，是跟著月份走的，不抄；
            # 只有店長特別改過的（例：區主管 120h）才帶過來
            if cur.get("type") == "正職" and prev_std is not None and cur.get("target") not in (None, "") and float(cur["target"]) == float(prev_std):
                cur["target"] = None
            cur["_default"] = True
            ms["emps"][e["id"]] = cur
        cur["move"] = "可"   # 2026-11-11 經營者：調店不再讓店長選，後台固定全部夥伴都可調店（舊存檔的「否」也不算）
        if cur.get("type") == "正職" and not cur.get("target"):
            must = (s104.get(e["name"]) or {}).get("mustHour")
            cur["target_hint"] = ms["base"].get("std_hours") or must
            if ms["base"].get("std_hours"):   # 2026-10-31 經營者：正職目標時數直接套本月標準時數（店長可改）
                cur["target"] = ms["base"]["std_hours"]
                cur["target_est"] = True
    ms["emps"] = {k: v for k, v in ms["emps"].items() if any(e["id"] == k for e in info.get("emps") or [])}
    ms["marks"] = {k: v for k, v in ms["marks"].items() if any(e["id"] == k for e in info.get("emps") or [])}
    av = ms.get("avail_shift") or {}
    for eid, cur in ms["emps"].items():   # 還沒用文字存過的（以前點格子的）→ 換成文字顯示
        mk = ms["marks"].get(eid) or {}
        for k, f in MARK_TXT.items():
            if f not in cur:
                ds = [int(d) for d, x in mk.items() if x == k]
                if k == "可上" and av.get(eid):
                    by = {}
                    for d in ds:
                        by.setdefault((av[eid].get(str(d)) or ""), []).append(d)
                    cur[f] = "；".join((_days_txt(v) + (" " + sh if sh else "")) for sh, v in by.items())
                else:
                    cur[f] = _days_txt(ds)
    return ms


def can_move(v):
    """區域排班能不能調去同城別家店：2026-11-11 經營者裁定後台固定「全部可」（原本由店長在本月設定選，正職預設可、兼職預設否）"""
    return True


def mset_clean(st, body):
    """本月設定（不含指休格子，那個走 mark）"""
    ms = copy.deepcopy(st.get("mset") or {})
    b = body.get("base") or {}
    ms["base"] = {"std_hours": _num(b.get("std_hours"), None, 0, 400, "標準時數"), "off_days": _num(b.get("off_days"), None, 0, 31, "休假天數"),
                  "holidays": _num(b.get("holidays"), None, 0, 31, "國定假日天數")}
    r = body.get("revenue") or {}
    ms["revenue"] = {"weekday": _num(r.get("weekday"), None, 0, 10_000_000, "平日營收"), "weekend": _num(r.get("weekend"), None, 0, 10_000_000, "假日營收")}
    ids = {e["id"] for e in ((st.get("template") or {}).get("info") or {}).get("emps") or []}
    emps = {}
    for eid, v in (body.get("emps") or {}).items():
        if eid not in ids:
            continue
        v = v or {}
        emps[eid] = {"type": ("兼職" if str(v.get("type") or "") == "兼職" else "正職"), "target": _num(v.get("target"), None, 0, 400, "目標時數"),
                     "min_hours": _num(v.get("min_hours"), None, 0, 400, "最低時數"), "note": str(v.get("note") or "").strip()[:300]}
        mv = str(v.get("move") or "")
        emps[eid]["move"] = mv if mv in ("可", "否") else ("否" if emps[eid]["type"] == "兼職" else "可")
        for f in MARK_TXT.values():
            emps[eid][f] = str(v.get(f) or "").strip()[:300]
    ms["emps"] = emps
    days = {d["d"] for d in ((st.get("template") or {}).get("info") or {}).get("days") or []}
    sp = []
    for x in (body.get("special") or [])[:40]:
        d = _num(x.get("day"), None, 1, 31, "日期")
        if d is None:
            continue
        if days and d not in days:
            raise ValueError(f"這個月沒有 {d} 號")
        row = {"day": int(d), "revenue": _num(x.get("revenue"), None, 0, 10_000_000, "營收預估"),
               "staff": str(x.get("staff") or "").strip()[:200], "note": str(x.get("note") or "").strip()[:200],
               "as_holiday": bool(x.get("as_holiday"))}
        if row["revenue"] is None and not row["staff"] and not row["note"] and not row["as_holiday"]:
            continue   # 整列空白（按了「加一天」沒填）不存
        if any(y["day"] == row["day"] for y in sp):
            raise ValueError(f"特殊日 {row['day']} 號填了兩列，請合成一列")
        sp.append(row)
    ms["special"] = sorted(sp, key=lambda x: x["day"])
    ms["marks"], ms["avail_shift"] = marks_from_text(st, emps, cfg_load(st["store"]))
    ms["updated"] = dt.datetime.now().isoformat(timespec="seconds")
    return ms


MARKS = ("指休", "特休", "可上")
# 2026-11-05 經營者：日曆一格一格點太慢 → 改成每人打字（指休／特休填日期，兼職填可上日＋時段），存檔時程式換成格子
MARK_TXT = {"指休": "off_txt", "特休": "pto_txt", "可上": "avail_txt"}
_WD = {"一": "一", "二": "二", "三": "三", "四": "四", "五": "五", "六": "六", "日": "日", "天": "日"}
_DAY_RE = re.compile(
    r"(?P<rng>(?:\d{1,2}/)?(?P<a>\d{1,2})\s*[號日]?\s*[-~～到至]\s*(?:\d{1,2}/)?(?P<b>\d{1,2})\s*[號日]?)"
    r"|(?P<wk>(?:每週|每周|週|周|星期|禮拜)(?P<wds>[一二三四五六日天]+))"
    r"|(?P<wday>平日)|(?P<wend>假日|週末|周末)|(?P<all>每天|全月|整個月|天天)"
    r"|(?P<one>(?:\d{1,2}/)?(?P<n>\d{1,2})\s*[號日]?)")
_FILLER = re.compile(r"[\s、，,／/。.和及跟與]|可上班|可以上|可上|都可以|都行|都OK|都ok|上班|其他|其餘|其它|剩下的?|都")
# 2026-11-05 經營者「提高店長自然語言的辨識率、不要給太多限制」：存檔與即時回顯前先把常見口語換成固定寫法
_WK = r"(?:每週|每周|週|周|星期|禮拜)"
_WD_ORDER = "一二三四五六日"
_NEG_WORDS = r"(?:以外|之外|除外|不行|不可以|不能|不要|沒空|沒辦法|不方便|休息|不上)"
# 班別口語：店長寫「早班／早上」→ 對到店別設定裡名稱有「早」的班（平日、假日各自對）
_SHIFT_ALIAS = (("早", r"早班|早上|上午|白天|早場"), ("中", r"中班|中午|午班|中場"), ("晚", r"晚班|晚上|夜班|傍晚|晚場"),
                ("下午", r"下午"))


def _wd_span(a, b):
    a, b = _WD_ORDER.index(_WD[a]), _WD_ORDER.index(_WD[b])
    return "".join(_WD_ORDER[(a + i) % 7] for i in range((b - a) % 7 + 1))


def _norm_days(seg, last, month=None):
    """口語 → 固定寫法（只換日期部分；班別另外處理）"""
    s = unicodedata.normalize("NFKC", seg)
    s = re.sub(r"\d{1,2}\s*月\s*(?=\d)", "", s)                                       # 12月4日 → 4日
    if month:   # 12/4 → 4；4/9/17（不是本月月份）→ 4、9、17（2026-11-05 模擬：原本把 4/9 當 4 月 9 日，4 號默默不見）
        s = re.sub(r"(\d{1,2})\s*/\s*(?=\d)", lambda m: "" if int(m.group(1)) == month else m.group(1) + "、", s)
    s = re.sub(_WK + r"([一二三四五六日天])\s*[-~到至]\s*" + _WK + r"?([一二三四五六日天])",
               lambda m: "每週" + _wd_span(m.group(1), m.group(2)), s)                 # 週一到週三 → 每週一二三
    while True:                                                                         # 週一、週三、五 → 每週一三五
        t = re.sub("(" + _WK + r"[一二三四五六日天]+)\s*[、,/和跟及與]?\s*" + _WK + r"?([一二三四五六日天])(?![一二三四五六日天]*[號日])",
                   lambda m: m.group(1) + m.group(2) if not m.group(0).endswith(m.group(1)) else m.group(0), s, count=1)
        if t == s:
            break
        s = t
    s = s.replace("上半月", "1-15").replace("下半月", f"16-{last}").replace("上旬", "1-10").replace("中旬", "11-20").replace("下旬", f"21-{last}")
    s = re.sub(r"(\d{1,2})\s*[號日]?\s*(?:以後|之後|後|起|開始)", lambda m: f"{m.group(1)}-{last}", s)
    s = re.sub(r"(\d{1,2})\s*[號日]?\s*(?:以前|之前|前)", lambda m: f"1-{int(m.group(1)) - 1}" if int(m.group(1)) > 1 else "", s)
    return s


def _scan_days(seg, wd_of, label):
    """固定寫法 → (日期 list, 剩下看不懂的字)"""
    days, rng, flt, rest = [], [], [], seg
    for m in _DAY_RE.finditer(seg):
        if m.group("rng"):
            a, b = int(m.group("a")), int(m.group("b"))
            if a > b:
                raise ValueError(f"{label}「{m.group(0)}」前面的日期要比後面小")
            rng += list(range(a, b + 1))
        elif m.group("wk"):
            ws = {_WD[c] for c in m.group("wds")}
            flt += [d for d, w in wd_of.items() if w in ws]
        elif m.group("wday"):
            flt += [d for d, w in wd_of.items() if w not in ("六", "日")]
        elif m.group("wend"):
            flt += [d for d, w in wd_of.items() if w in ("六", "日")]
        elif m.group("all"):
            days += list(wd_of)
        else:
            days.append(int(m.group("n")))
        rest = rest.replace(m.group(0), " ", 1)
    # 「下半月 假日」「10-20 平日」＝區間裡的那幾天（取交集）；其他寫法全部加起來
    if rng and flt:
        days += [d for d in rng if d in set(flt)]
    else:
        days += rng + flt
    return days, _FILLER.sub("", rest)


def _shift_of(seg, names):
    """這段寫了哪個班 → (把班別字拿掉的段落, 依日子挑班別的函式)；沒寫班別＝哪個班都可以"""
    for nm in names:
        if nm and nm in seg:
            return seg.replace(nm, " "), (lambda d, w, nm=nm: nm)
    for key, pat in _SHIFT_ALIAS:
        m = re.search(pat, seg)
        if not m:
            continue
        cand = [n for n in names if key in n]
        if key == "下午" and not cand:          # 店裡沒有「下午」班 → 有中班就算中班，沒有就算晚班
            cand = [n for n in names if "中" in n] or [n for n in names if "晚" in n]
        if not cand:
            continue

        def pick(d, w, cand=cand, key=key):
            # 2026-11-05 模擬：「每週一三五 下午」平日沒有中班，原本拿了假日中班塞到平日
            want, other = ("假日", "平日") if w in ("六", "日") else ("平日", "假日")
            hit = next((n for n in cand if want in n or other not in n), "")
            if not hit and key in ("中", "下午"):   # 這天類型沒有中班 → 下午／中班算晚班
                hit = next((n for n in names if "晚" in n and (want in n or other not in n)), "")
            return hit   # 都沒有＝這天不限班別
        return seg.replace(m.group(0), " "), pick
    return seg, (lambda d, w: "")


def _parse_seg(seg, wd_of, names, shift_mode, label, month=None):
    """一段（；之間）→ {日: 班別}"""
    last = max(wd_of)
    sh_fn = lambda d, w: ""
    seg0 = seg
    if shift_mode:
        seg, sh_fn = _shift_of(seg, names)
    had_shift = seg != seg0 or bool(re.search(r"全天|整天|任何班|都可以|都行", seg))
    seg = re.sub(r"全天|整天|任何班|都可以", " ", seg)
    seg = _norm_days(seg, last, month)
    neg = []
    # 「除了 5、6 號」「5 號以外」「20 號不行」→ 先拿出要排除的日子
    for pat in (r"(?:除了|除去|除)\s*(?P<x>[^，,。；]*?)\s*" + _NEG_WORDS + r"?(?=\s*$|\s*[，,。]|\s*都|\s*其他|\s*其餘|\s*其它)",
                r"(?P<x>[\d、,\s號日/\-~到至一二三四五六日天週周星期禮拜每平假末]+?)\s*" + _NEG_WORDS):
        for m in list(re.finditer(pat, seg)):
            d, left = _scan_days(m.group("x"), wd_of, label)
            if not d or left:
                continue
            neg += d
            seg = seg.replace(m.group(0), " ", 1)
    days, left = _scan_days(seg, wd_of, label)
    if left:
        hint = "；可上時段要寫店別基本設定裡的班別名稱（" + "、".join(names) + "）" if shift_mode and names else ""
        raise ValueError(f"{label}看不懂「{left}」：日期寫成 4、9、17 或 10-12，也可寫 每週一三五／平日／假日{hint}")
    if not days and (neg or (shift_mode and had_shift)):   # 「20 號不行」「早班」「全天都可以」＝整個月（「平日早班」只算平日）
        def fits(d):
            n, we = sh_fn(d, wd_of[d]), wd_of[d] in ("六", "日")
            return not (("平日" in n and we) or ("假日" in n and not we))
        days = [d for d in wd_of if fits(d)]
    out = {}
    for d in days:
        if d not in wd_of:
            raise ValueError(f"{label}這個月沒有 {d} 號")
        if d not in neg:
            out[d] = sh_fn(d, wd_of[d])
    return out


def parse_days_text(txt, info_days, shift_names=None, label=""):
    """「4、9、17」「10-12」「每週一三五 晚班；20、21 早班」「平日」→ [(日, 班別或空字串)]；看不懂丟 ValueError"""
    out = {}
    wd_of = {d["d"]: d["wd"] for d in info_days}
    names = sorted(shift_names or [], key=len, reverse=True)
    for seg in re.split(r"[；;\n]+", str(txt or "")):
        seg = seg.strip()
        if seg:
            out.update(_parse_seg(seg, wd_of, names, shift_names is not None, label, (info_days[0].get("m") if info_days else None)))
    return sorted(out.items())


class NeedAI(ValueError):
    """規則看不懂、還沒叫 AI（店長還在打字）"""


_ai_lock = threading.Lock()
KIND_TXT = {"指休": "指休（夥伴指定要休假的日子）", "特休": "特休（夥伴請特休的日子）", "可上": "兼職可上日與時段"}


def _ai_rewrite(st, seg, info_days, shift_names, kind):
    """規則看不懂的那一段 → 請 AI 改寫成固定寫法（同月同欄同一句只問一次，記在快取）→ (固定寫法, 這次有沒有真的問 AI)；改不了回空字串"""
    ym = st["ym"]
    ck = "|".join([ym, kind, ",".join(shift_names or []), seg])
    cp = _dir() / "parse_ai_cache.json"
    with _ai_lock:
        cache = json.loads(cp.read_text(encoding="utf-8")) if cp.exists() else {}
    if ck in cache:
        return cache[ck], False
    cal = "、".join(f"{d['d']}({d['wd']})" for d in info_days)
    sys_p = ("你把門市店長寫的排班日期，改寫成程式看得懂的固定寫法。只回 JSON：{\"std\":\"…\"}，看不出意思就回 {\"std\":\"\"}。\n"
             "固定寫法只能用這些：日期數字用「、」分開（4、9、17）；連續日期寫 10-12；每週一三五；平日；假日"
             + ("；這段可上的班別名稱寫在日期後面、空一格（只能用下面列的班別名稱，沒指定班別就不寫）；不同班別的段落用「；」分開" if shift_names is not None else "")
             + "。不要加其他字、不要解釋。意思是「排除某幾天」時，直接列出剩下的日期。")
    user = (f"月份：{ym[:4]} 年 {int(ym[4:])} 月，日期（星期）：{cal}\n欄位：{KIND_TXT.get(kind, kind)}\n"
            + (f"班別名稱：{'、'.join(shift_names) or '（沒有設定）'}\n" if shift_names is not None else "")
            + f"店長寫的：{seg}")
    try:
        msg = _client_().messages.create(model=MODEL, max_tokens=800, system=sys_p, messages=[{"role": "user", "content": user}])
        _log_usage(st, "parse", _cost(MODEL, msg.usage))
        txt = "".join(getattr(b, "text", "") for b in msg.content)
        m = re.search(r"\{.*\}", txt, re.S)
        std = str(json.loads(m.group(0)).get("std") or "").strip()[:300] if m else ""
    except Exception:   # noqa: BLE001  AI 連不上就當看不懂，照原本的錯誤訊息
        return "", True
    with _ai_lock:
        cache = json.loads(cp.read_text(encoding="utf-8")) if cp.exists() else {}
        cache[ck] = std
        atomic_write_json(cp, cache)
    return std, True


def _parse_log(st, kind, text, std, ok):
    """規則看不懂的說法記下來，每月看一次、常見的加進規則"""
    with _ai_lock, open(_dir() / "parse_log.jsonl", "a", encoding="utf-8") as f:
        f.write(json.dumps({"ts": dt.datetime.now().isoformat(timespec="seconds"), "store": st["store"], "ym": st["ym"],
                            "kind": kind, "text": text, "ai": std, "ok": ok}, ensure_ascii=False) + "\n")


def parse_days_smart(st, txt, info_days, shift_names, label, kind, ai=True):
    """先用規則；看不懂的那段請 AI 改寫再用規則讀一次 → (結果, 有沒有靠 AI, 固定寫法)"""
    out, used, std_parts = {}, False, []
    wd_of = {d["d"]: d["wd"] for d in info_days}
    names = sorted(shift_names or [], key=len, reverse=True)
    mon = info_days[0].get("m") if info_days else None
    for seg in re.split(r"[；;\n]+", str(txt or "")):
        seg = seg.strip()
        if not seg:
            continue
        try:
            r = _parse_seg(seg, wd_of, names, shift_names is not None, label, mon)
            std_parts.append(seg)
        except ValueError as ex:
            if not ai:
                raise NeedAI(str(ex)) from None
            std, fresh = _ai_rewrite(st, seg, info_days, shift_names, kind)
            try:
                if not std:
                    raise ValueError
                r = {}
                for s2 in re.split(r"[；;\n]+", std):
                    if s2.strip():
                        r.update(_parse_seg(s2.strip(), wd_of, names, shift_names is not None, label, mon))
            except ValueError:
                if fresh:
                    _parse_log(st, kind, seg, std, False)
                raise ValueError(f"{label}看不懂「{seg}」，AI 也看不出來。請改寫成 4、9、17／10-12／每週一三五／平日／假日") from None
            if fresh:
                _parse_log(st, kind, seg, std, True)
            used = True
            std_parts.append(std)
        out.update(r)
    return sorted(out.items()), used, "；".join(std_parts)


def _days_txt(ds):
    """[3,4,5,9] → 「3-5、9」"""
    ds, out, i = sorted(ds), [], 0
    while i < len(ds):
        j = i
        while j + 1 < len(ds) and ds[j + 1] == ds[j] + 1:
            j += 1
        out.append(f"{ds[i]}-{ds[j]}" if j - i >= 2 else "、".join(map(str, ds[i:j + 1])))
        i = j + 1
    return "、".join(out)


def marks_from_text(st, emps, cfg):
    """本月設定存檔時：每人三格文字 → marks（取代舊格子）＋兼職可上時段 avail_shift"""
    info = st["template"]["info"]
    names = {e["id"]: e["name"] for e in info["emps"]}
    shift_names = [r["name"] for r in (cfg.get("shifts") or [])]
    marks, avail, clash = {}, {}, []
    for eid, v in emps.items():
        mk = {}
        for k, f in MARK_TXT.items():
            days, _, _ = parse_days_smart(st, v.get(f), info["days"], shift_names if k == "可上" else None, f"{names.get(eid, eid)} {k}：", k)
            for d, sh in days:
                if str(d) in mk and mk[str(d)] != k:   # 衝突一次列完（2026-11-05 模擬：原本一次只說一天）
                    clash.append(f"{names.get(eid, eid)} {d} 號同時填了{mk[str(d)]}和{k}")
                    continue
                mk[str(d)] = k
                if k == "可上" and sh:
                    avail.setdefault(eid, {})[str(d)] = sh
        if mk:
            marks[eid] = mk
    if clash:
        raise ValueError("；".join(clash) + "。請改掉其中一個再存")
    return marks, avail



def mark_set(st, eid, day, value):
    ids = {e["id"] for e in st["template"]["info"]["emps"]}
    days = {d["d"] for d in st["template"]["info"]["days"]}
    if eid not in ids:
        raise ValueError(f"沒有工號 {eid}")
    if int(day) not in days:
        raise ValueError(f"這個月沒有 {day} 號")
    if value and value not in MARKS:
        raise ValueError("標記只能是 指休／特休／可上")
    ms = st.setdefault("mset", {})
    marks = ms.setdefault("marks", {}).setdefault(eid, {})
    if value:
        marks[str(int(day))] = value
    else:
        marks.pop(str(int(day)), None)
    ms["updated"] = dt.datetime.now().isoformat(timespec="seconds")


def form_ready(st, cfg):
    """表單填到可以排班了嗎（班別有填、營收預估有填）"""
    if not st.get("template"):
        return False
    rv = mset_view(st)["revenue"]   # 估算值也算（店長沒改就照估算排）
    return bool(cfg.get("shifts")) and bool(rv.get("weekday")) and bool(rv.get("weekend"))


# ───────── 挪假累積（2026-11-05 經營者：每月要看到每人累積欠幾天、補了幾天） ─────────
# 「彈性挪假」那天有上班＝夥伴休假日來上班 → 公司欠夥伴休假（要補休）；沒上班＝夥伴多放一天 → 夥伴欠公司上班（要補班）
# 「挪假補班」那天沒上班＝補休給夥伴；有上班＝補上班給公司。標籤在整天或時段上都算。
# 期初＝店別基本設定「挪假未補」（工作台以前的舊帳；日期落在工作台已有班表的月份就不重算）
NUO_REST, NUO_WORK = "要補休給夥伴", "要補上班給公司"
NUO_DAY_H = 8   # 沒上班的挪假日／補休日記 8 小時（2026-11-05 經營者：挪假、補班要自動記時數）


def _nuo_h(c, tag):
    """這格挪假／補班算幾小時：有上班＝標籤所在時段的實際時數（整天標籤＝全天），沒上班＝NUO_DAY_H"""
    if not c["segs"]:
        return NUO_DAY_H
    segs = c["segs"] if tag in c["dtag"] else [x for x in c["segs"] if tag in x[2]]
    return round(cell_hours(dict(c, segs=segs)), 2)


def _nuo_txt(d, h):
    return f"{d:g} 天（{h:g}h）" if d or h else "0 天"


NUO_F = ["rest_add", "work_add", "rest_paid", "work_paid", "rest_add_h", "work_add_h", "rest_paid_h", "work_paid_h"]


def _nuo_month(st):
    """這個月班表 → {工號: {"rest_add","work_add","rest_paid","work_paid", "days":[(日, 說明)]}}"""
    info = (st.get("template") or {}).get("info") or {}
    if not st.get("schedule") or not info:
        return {}
    shifts = _ext(st).get("shifts") or {}
    dtypes = set(info.get("dayTypes") or [])
    out = {}
    for e in info.get("emps") or []:
        row = st["schedule"].get(e["id"]) or {}
        for d in info.get("days") or []:
            try:
                c = parse_cell(row.get(str(d["d"]), ""), shifts, dtypes)
            except ValueError:
                continue
            tags = c["dtag"] + "、" + "、".join(x[2] for x in c["segs"])
            work = bool(c["segs"])
            r = out.setdefault(e["id"], dict({f: 0 for f in NUO_F}, days=[]))
            if "彈性挪假" in tags:
                f, h = "rest_add" if work else "work_add", _nuo_h(c, "彈性挪假")
                r[f] += 1
                r[f + "_h"] += h
                r["days"].append((d["d"], f"挪假（來上班 {h:g}h，之後補休）" if work else f"挪假（多放假 {h:g}h，之後補上班）"))
            if "挪假補班" in tags:
                f, h = "work_paid" if work else "rest_paid", _nuo_h(c, "挪假補班")
                r[f] += 1
                r[f + "_h"] += h
                r["days"].append((d["d"], f"補上班 {h:g}h" if work else f"補休 {h:g}h"))
    return {k: v for k, v in out.items() if v["days"]}


def nuo_ledger(st, cfg=None):
    """這家店到這個月為止的挪假累積：[{id, name, open_rest, open_work, ...本月..., close_rest, close_work, days}]"""
    info = (st.get("template") or {}).get("info") or {}
    if not info:
        return []
    cfg = cfg if cfg is not None else cfg_load(st["store"])
    ym = st["ym"]
    months = {}
    for p in _dir().glob(f"{int(st['store'])}-20*.json"):
        m = p.stem.split("-", 1)[1]
        if m < ym:
            try:
                ost = json.loads(p.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            if ost.get("schedule"):
                months[m] = ost
    months[ym] = st
    first = min(m for m, x in months.items() if x.get("schedule")) if any(x.get("schedule") for x in months.values()) else ym
    names = {e["id"]: e["name"] for e in info.get("emps") or []}
    by_name = {}
    for m in sorted(months):
        for e in ((months[m].get("template") or {}).get("info") or {}).get("emps") or []:
            names.setdefault(e["id"], e["name"])
            by_name[e["name"]] = e["id"]
    acc = {}
    for x in cfg.get("carry") or []:   # 期初舊帳
        m = re.match(r"(\d{4})\D?(\d{1,2})", x.get("date") or "")
        if m and f"{m.group(1)}{int(m.group(2)):02d}" >= first:
            continue
        eid = by_name.get(x.get("name")) or ("name:" + x.get("name", ""))
        a = acc.setdefault(eid, [0, 0, 0, 0])
        i = 0 if x.get("type") != "B" else 1
        a[i] += 1
        a[i + 2] += NUO_DAY_H
        names.setdefault(eid, x.get("name", ""))
    out = {}
    for m in sorted(months):
        mm = _nuo_month(months[m]) if months[m].get("schedule") else {}
        for eid, r in mm.items():
            a = acc.setdefault(eid, [0, 0, 0, 0])
            if m == ym:
                out[eid] = dict(r, open_rest=a[0], open_work=a[1], open_rest_h=a[2], open_work_h=a[3])
            a[0] += r["rest_add"] - r["rest_paid"]
            a[1] += r["work_add"] - r["work_paid"]
            a[2] += r["rest_add_h"] - r["rest_paid_h"]
            a[3] += r["work_add_h"] - r["work_paid_h"]
    rows = []
    for eid, a in acc.items():
        cur = out.get(eid) or dict({f: 0 for f in NUO_F}, days=[], open_rest=a[0], open_work=a[1], open_rest_h=a[2], open_work_h=a[3])
        if not (cur["open_rest"] or cur["open_work"] or cur["days"] or any(a)):
            continue
        rows.append(dict({f: round(cur[f], 2) for f in NUO_F}, id=eid, name=names.get(eid, eid),
                         open_rest=cur["open_rest"], open_work=cur["open_work"],
                         open_rest_h=round(cur["open_rest_h"], 2), open_work_h=round(cur["open_work_h"], 2),
                         close_rest=a[0], close_work=a[1], close_rest_h=round(a[2], 2), close_work_h=round(a[3], 2),
                         days=[f"{int(ym[4:])}/{d} {t}" for d, t in sorted(cur["days"])]))
    order = {e["id"]: i for i, e in enumerate(info.get("emps") or [])}
    return sorted(rows, key=lambda r: order.get(r["id"], 999))


def form_text(st, cfg):
    """把表單組成跟 Excel 設定表同樣結構的文字給 AI"""
    info = st["template"]["info"]
    ms = mset_view(st)
    L = ["### 分頁「各店基本設定」", f"店名\t{info['title']}", "班別\t第1段\t第2段\t第3段\t平日每天需要幾人\t假日每天需要幾人"]
    for r in cfg.get("shifts") or []:
        segs = [f"{a}-{b}" for a, b in r["segs"]] + [""] * (3 - len(r["segs"]))
        L.append(f"{r['name']}\t" + "\t".join(segs) + f"\t{r.get('wd', 0)}\t{r.get('we', 0)}")
    L.append(f"連續上班上限\t{LAW_MAX_CONSEC} 天（公司規定；勞基法七休一是 6 天）")
    L.append("休假分散\t正職休假要平均散開：連休最多 2 天（店長指休、特休自己連成的不算），避免「休、上 1 天、又休」")
    L.append("晚班接隔天早班\t可以，不限次數；下班到隔天上班要隔 11 小時以上（勞基法）")
    L.append("指休原則\t正職每月可自訂 4 天：平日 3 天、假日（週末、國定假日）1 天；超過的照店長填的排，但要在報告提醒店長")
    L.append("同一天指休\t同一家店兩人以上同一天指休，一定要提醒店長協調夥伴錯開休假日")
    if cfg.get("support"):
        L.append("跨店支援夥伴\t" + cfg["support"].replace("\n", "；"))
    if cfg.get("rules"):
        L.append("店規（一定要守）\t" + cfg["rules"].replace("\n", "；"))
    L.append("")
    L.append("### 分頁「每月排班」")
    L.append(f"月份\t{st['ym'][:4]}-{st['ym'][4:]}")
    b = ms["base"]
    if any(b.get(k) is not None for k in ("std_hours", "off_days", "holidays")):
        L.append(f"正職基準\t標準時數 {b.get('std_hours') or '—'}h／休假 {b.get('off_days') if b.get('off_days') is not None else '—'} 天／國定假日 {b.get('holidays') if b.get('holidays') is not None else '—'} 天")
    r = ms["revenue"]
    L.append(f"平日營收預估\t{r.get('weekday') or '—'}")
    L.append(f"假日營收預估\t{r.get('weekend') or '—'}")
    L.append("【夥伴】工號\t姓名\t正職/兼職\t目標時數\t兼職最低時數\t個人特殊班別／備註")
    for e in info["emps"]:
        v = ms["emps"].get(e["id"]) or {}
        tgt = v.get("target") or v.get("target_hint") or ""
        L.append(f"{e['id']}\t{e['name']}\t{v.get('type', '正職')}\t{tgt}\t{v.get('min_hours') or ''}\t{v.get('note') or ''}")
    L.append("【指休／特休／兼職可上日】（指休＝那天一定要休；可上＝兼職只有這些天能排，括號是只能上的班別）")
    av = ms.get("avail_shift") or {}
    for e in info["emps"]:
        mk = ms["marks"].get(e["id"]) or {}
        parts = []
        for k in MARKS:
            ds = sorted(int(d) for d, v in mk.items() if v == k)
            if ds:
                sh = av.get(e["id"]) or {}
                parts.append(f"{k} " + "、".join(f"{d}（{sh[str(d)]}）" if sh.get(str(d)) else str(d) for d in ds))
        if parts:
            L.append(f"{e['id']} {e['name']}：" + "；".join(parts))
    if ms["special"]:
        L.append("【特殊日】")
        for x in ms["special"]:
            bits = []
            if x.get("revenue"):
                bits.append(f"營收預估 {x['revenue']}")
            if x.get("staff"):
                bits.append(f"人力／班別調整：{x['staff']}")
            if x.get("note"):
                bits.append(f"備註：{x['note']}")
            if x.get("as_holiday"):
                bits.append("當假日（人力需求照假日班）")
            L.append(f"{x['day']} 日：" + "；".join(bits))
    led = nuo_ledger(st, cfg)
    if any(r["open_rest"] or r["open_work"] for r in led):
        L.append("【挪假累積（上個月底還沒補的，程式從各月班表算的）】")
        for r in led:
            if r["open_rest"] or r["open_work"]:
                L.append(f"{r['name']}：{NUO_REST} {_nuo_txt(r['open_rest'], r['open_rest_h'])}；{NUO_WORK} {_nuo_txt(r['open_work'], r['open_work_h'])}")
    if cfg.get("carry"):
        L.append("【挪假未補清單】姓名\t類型\t挪假日期\t備註")
        for x in cfg["carry"]:
            L.append(f"{x['name']}\t{x['type']}\t{x['date']}\t{x.get('note', '')}")
    return "\n".join(L)


def _ext_from_form(st, cfg):
    """表單直接算得出來的設定（AI 還沒排也能檢查、顯示班別名）"""
    if not cfg.get("shifts"):
        return {}
    ms = mset_view(st)
    info = (st.get("template") or {}).get("info") or {}
    shifts = {r["name"]: [list(x) for x in r["segs"]] for r in cfg["shifts"]}
    special = {str(x["day"]): x.get("revenue") for x in ms["special"] if x.get("revenue")}
    holi = _hol_days(ms)
    demand = {}
    for d in info.get("days") or []:
        wk = d["wd"] in ("六", "日") or d["d"] in holi
        need = []
        for r in cfg["shifts"]:
            need += [r["name"]] * int(r.get("we" if wk else "wd") or 0)
        demand[str(d["d"])] = need
    emps = []
    for e in info.get("emps") or []:
        v = ms["emps"].get(e["id"]) or {}
        emps.append({"id": e["id"], "name": e["name"], "type": v.get("type", "正職"),
                     "target_hours": v.get("target") or (v.get("target_hint") if v.get("type", "正職") == "正職" else None),
                     "min_hours": v.get("min_hours") if v.get("type") == "兼職" else None})
    return {"shifts": shifts, "demand": demand, "employees": emps, "max_consec": cfg.get("max_consec", 5),
            "late_early_limit": cfg.get("late_early_limit", 0),
            "late_shift_ends": sorted({r["segs"][-1][1] for r in cfg["shifts"] if "晚" in r["name"]}),
            "early_shift_starts": sorted({r["segs"][0][0] for r in cfg["shifts"] if "早" in r["name"]}),
            "revenue": {"weekday": ms["revenue"].get("weekday"), "weekend": ms["revenue"].get("weekend"), "special": special},
            "hol_days": sorted(holi), "notes": [], "from_form": True}


# 2026-11-11 經營者：員工排休其中 4 天可自訂（平日 3 天、假日 1 天）；假日（週末、國定假日、當假日）以 1 天為原則，
# 實際仍看各店營業狀況與店長安排 → 超過只提醒、不擋
PICK_OFF = {"total": 4, "weekday": 3, "holiday": 1}


def _gov_hol(st):
    """這個月落在平日的政府放假日（國定假日）"""
    y, m = int(st["ym"][:4]), int(st["ym"][4:])
    return sorted(d["d"] for d in st["template"]["info"]["days"] if d["wd"] not in ("六", "日") and _tw_off(dt.date(y, m, d["d"])))


def _pick_off_warns(st, names, etype):
    """正職指休超過「平日 3、假日 1、共 4 天」→ 提醒文字"""
    ms = st.get("mset") or {}
    info = st["template"]["info"]
    wk = {d["d"] for d in info["days"] if d["wd"] in ("六", "日")} | set(_gov_hol(st)) | _hol_days(ms)
    out = []
    for eid, mk in (ms.get("marks") or {}).items():
        if eid not in names or etype.get(eid) == "兼職":
            continue
        ds = [int(d) for d, v in mk.items() if v == "指休"]
        h = sum(1 for d in ds if d in wk)
        w = len(ds) - h
        if len(ds) > PICK_OFF["total"] or h > PICK_OFF["holiday"] or w > PICK_OFF["weekday"]:
            out.append(f"{names[eid]}：指休 {len(ds)} 天（平日 {w}、假日 {h}），原則上每月可自訂 4 天：平日 3 天、假日 1 天；超過的要店長視營業狀況決定")
    # 2026-11-11 經營者：同一家店兩人以上同一天指休，一定要提醒店長協調夥伴錯開
    by = {}
    for eid, mk in (ms.get("marks") or {}).items():
        if eid in names and etype.get(eid) != "兼職":
            for d, v in mk.items():
                if v == "指休":
                    by.setdefault(int(d), []).append(names[eid])
    lab = {d["d"]: d["label"] for d in info["days"]}
    for d in sorted(by):
        if len(by[d]) > 1:
            out.append(f"{lab.get(d, d)}：{'、'.join(by[d])} 同一天指休，請店長協調夥伴錯開休假日")
    return out


def _hol_days(ms):
    """特殊日勾了「當假日」的平日（2026-11-05 經營者：店長說 25、31 號視為假日，檢查卻還照平日班在比）"""
    return {int(x["day"]) for x in (ms or {}).get("special") or [] if x.get("as_holiday")}


def _ext(st, cfg=None):
    """AI 排過用 AI 的（班別名以表單為準）；沒排過用表單算的"""
    cfg = cfg if cfg is not None else cfg_load(st["store"])
    form = _ext_from_form(st, cfg) if st.get("template") else {}
    ai = st.get("ext") or {}
    if not ai:
        return form
    if not form:
        return ai
    if "demand_delta" in ai:
        # 2026-11-11：程式排的初稿只記「備註把哪幾天的人數改成多少」，其餘（每天幾人、目標時數、營收）一律讀現在的設定。
        # 原本整份需求在排初稿當下就定住，店長事後改「每天幾人」或目標時數，檢查和摘要都不會跟著變。
        out = dict(form, notes=ai.get("notes") or [], from_form=False)
        out["demand"] = dict(form.get("demand") or {}, **{str(d): v for d, v in (ai.get("demand_delta") or {}).items()})
        return out
    if "from_form" in ai:   # 2026-11-10 以前程式排的初稿：整份設定定在排班當下 → 只保留每天需求與備註，其餘讀現在的設定
        out = dict(form, demand=ai.get("demand") or form.get("demand") or {}, notes=ai.get("notes") or [], from_form=False)
    else:
        out = dict(form, **ai)
    out["shifts"] = dict(ai.get("shifts") or {}, **form["shifts"])
    # 排完之後才改「當假日」的日子：那幾天的需求改照表單（平日班↔假日班），其他天保留排班時算好的需求
    flip = set(form.get("hol_days") or []) ^ set(ai.get("hol_days") or [])
    if flip and ai.get("demand"):
        out["demand"] = dict(ai["demand"], **{str(d): form["demand"].get(str(d), []) for d in flip})
    out["hol_days"] = form.get("hol_days") or []
    return out


# ───────────────────────── 填 104 範本 ─────────────────────────

def fill(st, schedule=None):
    """把班表填進店長上傳的原始範本，回傳 xlsx bytes（只改格子的值，下拉選單、合併儲存格不動）"""
    schedule = schedule if schedule is not None else st["schedule"]
    raw = base64.b64decode(st["template"]["b64"])
    wb = openpyxl.load_workbook(io.BytesIO(raw))
    ws = wb.worksheets[0]
    hdr, days, emps = c104.layout(ws)
    shifts = _ext(st).get("shifts") or {}
    dtypes = set(st["template"]["info"]["dayTypes"])
    for e in emps:
        row = (schedule or {}).get(e["id"]) or {}
        for d in days:
            dn = int(d["label"][3:5])
            col = d["col"]
            for k in range(4):
                for i in range(6):
                    cell = ws.cell(e["row"] + k, col + i)
                    if type(cell).__name__ != "MergedCell":
                        cell.value = None
            c = parse_cell(row.get(str(dn), ""), shifts, dtypes)
            if c["type"]:
                ws.cell(e["row"], col).value = c["type"]
            if c["dtag"]:
                ws.cell(e["row"], col + 5).value = c["dtag"]
            for k, (s, en, tag, pos) in enumerate(c["segs"][:4]):
                ws.cell(e["row"] + k, col + 1).value = s
                ws.cell(e["row"] + k, col + 2).value = en
                if pos:
                    ws.cell(e["row"] + k, col + 3).value = pos
                if tag:
                    ws.cell(e["row"] + k, col + 4).value = tag
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ───────────────────────── 檢查 ─────────────────────────

# 不算現場人力的時段標籤／整天標籤（schedule_rules.md：開會、內訓、事務不佔現場班；特休、補休照常排班但人不在）
# 2026-11-05 經營者：謝雨庭 12/4、12/11 排事務班，早班沒人，檢查卻沒列出缺口
NOFLOOR_SEG_TAGS = {"事務班", "交流會議", "區會", "台北內訓課", "經管會", "特休", "補休"}   # 特休／補休標在時段上（104 吳予茜 11/24）也不在店
NOFLOOR_DAY_TAGS = {"特休", "補休"}


GAP_MIN = 15   # 少於 15 分鐘的空窗不算（交接）


def _hm2m(t):
    """'HH:MM'／'次日HH:MM' → 分鐘"""
    nxt = str(t).startswith("次日")
    h, m = str(t).replace("次日", "").split(":")
    return int(h) * 60 + int(m) + (1440 if nxt else 0)


def _m2hm(t):
    return ("次日" if t >= 1440 else "") + f"{t % 1440 // 60:02d}:{t % 60:02d}"


def _coverage(info, parsed, ext, names, stat=None):
    """每天需要幾人／哪些班有沒有人顧（只算在本店現場的時段）。回傳 (errors, warns)"""
    errors, warns = [], []
    demand = ext.get("demand") or {}
    shifts = ext.get("shifts") or {}
    for d in info["days"]:
        need = demand.get(str(d["d"])) or []
        if not need:
            continue
        working, away = [], []
        for eid in parsed:
            c = parsed[eid][d["d"]]
            here = [s for s in c["segs"] if not _elsewhere(s[3], info["title"])]
            if not here:
                continue
            why = next((t for t in NOFLOOR_DAY_TAGS if t in c["dtag"]), "")
            floor = [] if why else [s for s in here if not (set(re.split(r"[、,，]", s[2] or "")) & NOFLOOR_SEG_TAGS)]   # 104 多標籤「支援、事務班」也要認得
            if floor:
                working.append((names.get(eid, eid), floor))
            else:
                away.append(f"{names.get(eid, eid)}{why or next((s[2] for s in here if s[2]), '')}")
        note = f"（{'、'.join(away)}不算現場人力）" if away else ""
        if stat is not None:
            stat.append({"d": d["d"], "label": d["label"], "wd": d["wd"], "need": len(need), "got": len(working)})
        if len(working) < len(need):
            errors.append(f"{d['label']} 需要 {len(need)} 人（{'、'.join(need)}），只排了 {len(working)} 人{note}")
            continue
        # 2026-11-05：原本只比「開始時間一模一樣」，葉思柔 09:50–18:20 上平日也被說「早班沒人在那個時間上班」。
        # 改成：需要的班的時段（含午休空檔）疊起來，跟實際在店現場的時段比，真的有空窗（店裡沒人）才提醒
        req, cov = set(), set()
        for nm in need:
            for a, b in shifts.get(nm) or []:
                req.update(range(_hm2m(a), _hm2m(b)))
        for _, segs in working:
            for sg in segs:
                cov.update(range(_hm2m(sg[0]), _hm2m(sg[1])))
        holes, start = [], None
        for t in range(min(req, default=0), max(req, default=-1) + 2):
            empty = t in req and t not in cov
            if empty and start is None:
                start = t
            elif not empty and start is not None:
                if t - start >= GAP_MIN:
                    holes.append(f"{_m2hm(start)}–{_m2hm(t)}")
                start = None
        if holes:
            warns.append(f"{d['label']} {'、'.join(holes)} 店裡沒有人（需要：{'、'.join(need)}）{note}")
    return errors, warns


_POS_WARN_RE = re.compile(r"^(.+?)\(\d+\) \d+/\d+ 第\d+段：崗位填了本店「(.+?)」")


def _merge_pos_warn(warns):
    """「崗位填了本店」每個時段一條，士林／板橋／A13 一個月 150～190 條，把其他提醒都淹沒了
    → 合併成一條（2026-11-11 經營者）：幾個時段、哪幾位夥伴。其他提醒順序不變，合併的那條放在第一次出現的位置"""
    hit = [(i, m) for i, m in ((i, _POS_WARN_RE.match(x)) for i, x in enumerate(warns)) if m]
    if len(hit) < 2:
        return list(warns)
    names = list(dict.fromkeys(m.group(1) for _, m in hit))
    pos = hit[0][1].group(2)
    one = (f"有 {len(hit)} 個時段的崗位填了本店「{pos}」（{'、'.join(names[:6])}{' 等 ' + str(len(names)) + ' 位' if len(names) > 6 else ''}）："
           f"在本店上班的時段崗位要留空，只有去別家店才填。不影響時數與匯入，之後排班留空即可")
    skip = {i for i, _ in hit}
    out = []
    for i, x in enumerate(warns):
        if i == hit[0][0]:
            out.append(one)
        if i not in skip:
            out.append(x)
    return out


def _tail_merged(store, ym, ids):
    """排初稿用的「上個月底已連上幾天」：本店算的與各店名單合起來算的取大的"""
    out = dict(_prev_tail(store, ym))
    try:
        for eid, n in _prev_tail_all(ym, ids).items():
            out[eid] = max(int(out.get(eid) or 0), n)
    except Exception:   # noqa: BLE001
        pass
    return out


def _prev_tail_all(ym, ids):
    """這些工號上個月底「連續上班到月底」幾天（各店名單合起來算）→ {工號: 天數}；上個月沒班表就沒有"""
    pym = _prev_ym(ym)
    worked, last = {}, 0
    for s in STORES:
        p = _dir() / f"{_key(s, pym)}.json"
        if not p.exists():
            continue
        try:
            st2 = json.loads(p.read_text(encoding="utf-8"))
            tpl, sch = st2.get("template"), st2.get("schedule") or {}
            if not tpl or not sch:
                continue
            shifts, dtypes = (_ext(st2) or {}).get("shifts") or {}, set(tpl["info"]["dayTypes"])
            last = max(last, max(d["d"] for d in tpl["info"]["days"]))
        except (OSError, json.JSONDecodeError, KeyError, TypeError, ValueError):
            continue
        for eid in ids:
            for dd, v in (sch.get(eid) or {}).items():
                try:
                    if v and parse_cell(v, shifts, dtypes)["segs"]:
                        worked.setdefault(eid, set()).add(int(dd))
                except (ValueError, TypeError):
                    continue
    out = {}
    for eid, ds in worked.items():
        n, d = 0, last
        while d in ds:
            n, d = n + 1, d - 1
        if n:
            out[eid] = n
    return out


def _other_store_days(store, ym, ids, override=None):
    """這些工號在「別家店的名單」上有排班的日子 → {工號: {日: 店名}}（不含 @支援 調店：那是寫在自己店的班表上）"""
    out = {}
    for s in STORES:
        if int(s) == int(store):
            continue
        p = _dir() / f"{_key(s, ym)}.json"
        if not p.exists():
            continue
        try:
            st2 = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        sch, tpl = st2.get("schedule") or {}, st2.get("template")
        if override and s in override:   # 還沒採用的區域方案：用方案裡的班表，不讀存檔
            sch = override[s] or {}
        hit = [e for e in ids if e in sch]
        if not hit or not tpl:
            continue
        shifts, dtypes = (_ext(st2) or {}).get("shifts") or {}, set(tpl["info"]["dayTypes"])
        for eid in hit:
            for dd, v in (sch.get(eid) or {}).items():
                try:
                    if v and parse_cell(v, shifts, dtypes)["segs"]:
                        out.setdefault(eid, {})[int(dd)] = _sname(s) + "店"
                except (ValueError, TypeError):
                    continue
    return out


def run_check(st, schedule=None, guests=None, sibs=None):
    """檢查程式：104 格式＋勞基法（check_104）＋本店需求人數／正職時數／晚接早。回傳 dict
    guests：同區外店調來本店的夥伴（_guest_cells）；None＝讀同區各店存檔（2026-11-06 區域一起排）"""
    schedule = schedule if schedule is not None else st["schedule"]
    ext = _ext(st)
    shifts = ext.get("shifts") or {}
    info = st["template"]["info"]
    dtypes = set(info["dayTypes"])
    errors, warns = [], []
    parsed = {}
    badcells = set()
    names = {e["id"]: e["name"] for e in info["emps"]}
    for e in info["emps"]:
        row = (schedule or {}).get(e["id"]) or {}
        parsed[e["id"]] = {}
        for d in info["days"]:
            try:
                parsed[e["id"]][d["d"]] = parse_cell(row.get(str(d["d"]), ""), shifts, dtypes)
            except ValueError as ex:
                errors.append(f"{names[e['id']]} {d['label']}：{ex}")
                parsed[e["id"]][d["d"]] = {"type": "", "segs": [], "dtag": ""}
                badcells.add((e["id"], str(d["d"])))
    hours = []
    if True:   # 2026-11-11 經營者：有格子讀不懂時，那格當空白，其他人的時數與檢查照算（原本整份時數都不顯示）
        raw_t = base64.b64decode(st["template"]["b64"])
        sched_ok = schedule if not badcells else {eid: {dd: ("" if (eid, str(dd)) in badcells else v) for dd, v in (row or {}).items()} for eid, row in (schedule or {}).items()}
        out = fill(st, sched_ok)
        tpl = io.BytesIO(raw_t)
        tmp_dir = _dir() / "_tmp"
        tmp_dir.mkdir(exist_ok=True)
        p = tmp_dir / f"{_key(st['store'], st['ym'])}_{threading.get_ident()}.xlsx"
        tp = tmp_dir / f"{_key(st['store'], st['ym'])}_{threading.get_ident()}_t.xlsx"
        p.write_bytes(out)
        tp.write_bytes(raw_t)
        try:
            rep, rows, _ = c104.check(str(p), str(tp), int(ext.get("max_consec") or 5), {})
        finally:
            for f in (p, tp):
                try:
                    f.unlink()
                except OSError:
                    pass
        errors += rep.err
        warns += _merge_pos_warn(rep.warn)
        for emp, _, inf in rows:
            tot = sum(b - a for _, segs, _ in inf for a, b in segs) / 60
            # 104 把休息日、國定假日出勤列為加班，不算進總工時 → 跟正職標準時數比的是「正常工時」
            ot = sum(b - a for t, segs, _ in inf if t in ("休息日", "國定假日") for a, b in segs) / 60
            cnt = lambda x: sum(1 for t, _, _ in inf if t == x)
            hours.append({"id": emp["id"], "name": emp["name"], "days": sum(1 for _, s, _ in inf if s),
                          "hours": round(tot, 2), "regular": round(tot - ot, 2), "ot": round(ot, 2),
                          "例假": cnt("例假日"), "休息": cnt("休息日"), "國定": cnt("國定假日"),
                          # 特休照排班＋當日標籤「特休」（104 寫法），類型欄是空的 → 也要算（2026-11-05 模擬：特休欄一直是 0）
                          "特休": sum(1 for t, _, tg in inf if t == "特休" or "特休" in str(tg or ""))})
    # 本店規則（現場人力含同區外店調來支援的人）
    if guests is None:
        guests = _guest_cells(st["store"], st["ym"])
    p2, n2 = dict(parsed), dict(names)
    for gk, g in (guests or {}).items():
        p2[gk] = {d["d"]: g["days"].get(d["d"], {"type": "", "segs": [], "dtag": ""}) for d in info["days"]}
        n2[gk] = g["name"]
    cover = []
    e2, w2 = _coverage(info, p2, ext, n2, cover)
    errors += e2
    warns += w2
    late_starts = set(ext.get("late_shift_ends") or [])
    early_starts = set(ext.get("early_shift_starts") or [])
    for emp in ext.get("employees") or []:
        tgt = emp.get("target_hours")
        h = next((x for x in hours if x["id"] == emp.get("id")), None)
        if h:   # 目標、最低時數直接寫在每人時數裡，AI 不用自己找設定再比大小（2026-11-05 模擬：64h 被說成沒到最低 60h）
            h["target"], h["min"] = tgt, emp.get("min_hours")
        if tgt and h and abs(h["regular"] - float(tgt)) > 0.01:
            extra = f"（另有休息日／國定假日出勤 {h['ot']:g}h，104 列加班）" if h["ot"] else ""
            warns.append(f"{h['name']}：正常工時 {h['regular']:g}h，目標 {float(tgt):g}h{extra}")
    marks = ((st.get("mset") or {}).get("marks") or {})
    avail = ((st.get("mset") or {}).get("avail_shift") or {})
    etype = {e.get("id"): e.get("type") for e in ext.get("employees") or []}
    try:
        warns += _pick_off_warns(st, {k: names[k] for k in parsed if k in names}, etype)
    except Exception:   # noqa: BLE001 — 提醒算不出來不影響其他檢查
        pass
    for eid, mk in marks.items():
        if eid not in parsed:
            continue
        allow = {int(d) for d, v in mk.items() if v == "可上"}
        for d in info["days"]:
            c = parsed[eid][d["d"]]
            v = mk.get(str(d["d"]), "")
            if v == "指休" and c["segs"]:
                errors.append(f"{names[eid]} {d['label']}：本月設定是指休，卻排了班")
            elif v == "特休" and "特休" not in (c["type"] + c["dtag"] + "".join(x[2] or "" for x in c["segs"])):
                warns.append(f"{names[eid]} {d['label']}：本月設定要特休，班表沒有標特休")
            elif allow and etype.get(eid) == "兼職" and c["segs"] and d["d"] not in allow:
                warns.append(f"{names[eid]} {d['label']}：兼職只有填「可上」的日子能排，這天沒填")
            elif v == "可上" and c["segs"] and avail.get(eid, {}).get(str(d["d"])) in shifts:
                want = avail[eid][str(d["d"])]
                if [[x[0], x[1]] for x in c["segs"]] != [list(x) for x in shifts[want]]:
                    warns.append(f"{names[eid]} {d['label']}：兼職這天只能上{want}，班表排的不一樣")
    limit = ext.get("late_early_limit")
    if late_starts and early_starts and limit is not None:
        for eid in parsed:
            n = 0
            ds = [d["d"] for d in info["days"]]
            for a, b in zip(ds, ds[1:]):
                ca, cb = parsed[eid][a], parsed[eid][b]
                if ca["segs"] and cb["segs"] and ca["segs"][-1][1] in late_starts and cb["segs"][0][0] in early_starts:
                    n += 1
            if n > int(limit):
                warns.append(f"{names[eid]}：晚班接隔天早班 {n} 次，本店上限 {limit} 次")
    # 2026-11-11 經營者：同一人掛在別家店名單也有排班 → 連上天數要兩邊合起來算、同一天不可以兩家店都有班；
    # 上個月底連續上班的天數（本店＋別店）也要接上來算
    try:
        mc = int(ext.get("max_consec") or 5)
        others = _other_store_days(st["store"], st["ym"], set(parsed), sibs)
        tails = _prev_tail_all(st["ym"], set(parsed))
        dlist = [x["d"] for x in info["days"]]

        def _run(ds, tail=0):
            best, cur, st0 = (0, 0, 0), tail, (dlist[0] if tail else 0)
            for dd in dlist:
                if dd in ds:
                    cur += 1
                    st0 = st0 or dd
                    best = max(best, (cur, st0, dd))
                else:
                    cur, st0 = 0, 0
            return best
        for eid in parsed:
            od, tail = others.get(eid) or {}, int(tails.get(eid) or 0)
            if not od and not tail:
                continue
            own = {d["d"] for d in info["days"] if parsed[eid][d["d"]]["segs"]}
            both = sorted(own & set(od))
            if both:
                errors.append(f"{names[eid]}：{'、'.join(str(x) for x in both)} 號同一天在{'、'.join(sorted({od[x] for x in both}))}也有排班")
            n, a, b = _run(own | set(od), tail)
            if n > _run(own)[0] and n > min(mc, 6):
                used_tail = bool(tail) and a == dlist[0]
                used_other = any(a <= x <= b for x in od)
                why = "、".join(([f"上個月底已連上 {tail} 天"] if used_tail else []) + ([f"在{'、'.join(sorted({od[x] for x in od if a <= x <= b}))}的班"] if used_other else []))
                msg = f"{names[eid]}：加上{why}，到 {b} 號連續上班 {n} 天" if used_tail else f"{names[eid]}：加上{why}，{a}～{b} 號連續上班 {n} 天"
                (errors if n > 6 else warns).append(msg + ("，超過 6 天（七休一）" if n > 6 else f"，超過公司上限 {mc} 天"))
    except Exception:   # noqa: BLE001 — 讀別店存檔失敗不影響本店檢查
        pass
    # 2026-11-10：每天需要／排到幾人、正職「可支援其他店」的日子（標彈性挪假又沒排班）也放進檢查結果，
    # 畫面的摘要、報告、檢查都讀這一份（原本方案比較用排班當時的數字、檢查用現在的班表，兩邊對不起來）
    support = {eid: [d["d"] for d in info["days"] if "彈性挪假" in parsed[eid][d["d"]]["dtag"] and not parsed[eid][d["d"]]["segs"]]
               for eid in parsed if etype.get(eid) != "兼職"}
    return {"ok": not errors, "errors": errors, "warns": warns, "hours": hours, "cover": cover,
            "support": {k: v for k, v in support.items() if v}}


def check_text(chk):
    lines = ["【每人時數】"]
    for h in chk["hours"]:
        cmp = ""
        if h.get("target"):
            d = h["regular"] - float(h["target"])
            cmp += f"；目標 {float(h['target']):g}h（{'剛好' if abs(d) < 0.01 else ('多 ' if d > 0 else '少 ') + f'{abs(d):g}h'}）"
        if h.get("min"):
            d = h["regular"] - float(h["min"])
            cmp += f"；兼職最低 {float(h['min']):g}h（{'已達到' if d >= -0.01 else f'還差 {-d:g}h'}）"
        lines.append(f"{h['id']} {h['name']}：上班 {h['days']} 天、共 {h['hours']:g}h（正常工時 {h['regular']:g}h＋休息日／國定出勤 {h['ot']:g}h）、例假 {h['例假']}、休息 {h['休息']}、國定 {h['國定']}、特休 {h['特休']}{cmp}")
    lines.append(f"【❌ 一定要改】{len(chk['errors'])} 項")
    lines += ["❌ " + x for x in chk["errors"][:40]]
    lines.append(f"【⚠ 請店長確認】{len(chk['warns'])} 項")
    lines += ["⚠ " + x for x in chk["warns"][:40]]
    return "\n".join(lines)


def table_text(st, schedule=None):
    schedule = schedule if schedule is not None else st["schedule"]
    info = st["template"]["info"]
    out = []
    for e in info["emps"]:
        row = (schedule or {}).get(e["id"]) or {}
        out.append(f"{e['id']} {e['name']}：" + "；".join(f"{d['d']}={row.get(str(d['d']), '')}" for d in info["days"]))
    return "\n".join(out)


# ───────────────────────── AI 排初稿 ─────────────────────────

DRAFT_CONTRACT = """
## 這次的工作：排整月初稿
1. 先輸出「排班分析報告」四段（用表格、精簡）：①每日時數與人力配置 ②人效燈號（表格下方附一行圖例：🟢優秀（人力精實、人效高）／🟡正常／🔴偏低（人可能排太多））③規則檢核與人力盤點（含七休一、國定假日位置、同日跨店）④提醒與建議（缺口、無法達標的最低時數、零備援日、資料矛盾、不在 104 群組的人；🔴 紅燈日指名建議誰改特休或轉店務）。
2. 報告最後，用一個 ```json 區塊輸出下面格式（一定要完整，程式會直接讀）：
{
 "settings": {
  "shifts": {"平日早班": [["10:20","12:00"],["12:30","18:50"]], ...},   // 設定表裡每個班別的時段，名稱自己取但要清楚（平日／假日分開）
  "demand": {"1": ["平日早班","平日晚班"], "2": [...], ...},            // 每一天必須有人的班（含特殊日加班），一天一個陣列
  "employees": [{"id":"104 工號","name":"姓名","type":"正職或兼職","target_hours":176}],
  "max_consec": 5,
  "late_early_limit": 0,                                                 // 晚接早最多幾次（不允許＝0）
  "late_shift_ends": ["21:30"], "early_shift_starts": ["10:20","09:50"],  // 晚班下班時間、早班上班時間
  "revenue": {"weekday": 14000, "weekend": 32000, "special": {"25": 30000}},
  "notes": ["店規或讀設定時的重要判斷，一句一條"]
 },
 "schedule": {"工號": {"1": "格子", "2": "格子", ...}, ...}              // 每位員工、每一天都要有值（兼職沒班填 ""）
}
3. 程式會用檢查程式驗你的班表，有 ❌ 會退回給你改。
"""


def _draft_user(st):
    info = st["template"]["info"]
    parts = ["我要排班，月份 " + st["ym"][:4] + "-" + st["ym"][4:] + "。"]
    parts.append("## 104 匯入範本\n- 班表名稱：" + info["title"])
    parts.append("- 日期：" + "、".join(d["label"] for d in info["days"]))
    parts.append("- 員工（工號 姓名）：" + "；".join(f"{e['id']} {e['name']}" for e in info["emps"]))
    parts.append("- 可用的出勤類型：" + "、".join(info["dayTypes"]))
    parts.append("- 可用的標籤：" + "、".join(info["tags"]))
    parts.append("- 可用的崗位：" + "、".join(info["positions"]))
    cfg = cfg_load(st["store"])
    if form_ready(st, cfg):
        parts.append("## 每月排班設定表（店長在工作台填的；班別名稱請照這裡的用、不要自己取）\n" + form_text(st, cfg))
    for f in st["settings_files"]:
        parts.append(f"## 每月排班設定表{'（Excel 補充）' if form_ready(st, cfg) else ''}（{f['name']}）\n{f['text']}")
    if st.get("note"):
        parts.append("## 店長補充說明\n" + st["note"])
    return "\n".join(parts)


def _extract_json(text):
    m = re.findall(r"```json\s*(\{.*\})\s*```", text, re.S)
    if not m:
        raise ValueError("回覆裡沒有 JSON 班表")
    s = re.sub(r"//[^\n]*", "", m[-1])
    return json.loads(s)


def _validate_draft(st, d):
    if not isinstance(d, dict) or not isinstance(d.get("settings"), dict) or not isinstance(d.get("schedule"), dict):
        raise ValueError("JSON 要有 settings 和 schedule 兩塊")
    ids = [e["id"] for e in st["template"]["info"]["emps"]]
    miss = [i for i in ids if i not in d["schedule"]]
    if miss:
        raise ValueError("schedule 少了這些工號：" + "、".join(miss))
    sched = {i: {str(k): str(v or "") for k, v in (d["schedule"].get(i) or {}).items()} for i in ids}
    return d["settings"], sched


def _stream(model, system, messages, effort):
    with _client_().messages.stream(model=model, max_tokens=64000, system=system, messages=messages,
                                    output_config={"effort": effort}) as s:
        return s.get_final_message()


def _job(k, **kw):
    with _glock:
        old = _jobs.get(k) or {}
        now = time.time()
        # t0＝這次工作開始的時間（running 期間換步驟不重算），畫面用來顯示「已等幾秒」
        t0 = old.get("t0", now) if kw.get("status") == "running" and old.get("status") == "running" else now
        _jobs[k] = dict(kw, ts=now, t0=t0)


def _draft_worker_ai(store, ym):
    k = _key(store, ym)
    lk = _lock(k)
    try:
        with lk:
            st = _load(store, ym)
        system = RULES + DRAFT_CONTRACT
        messages = [{"role": "user", "content": _draft_user(st)}]
        report, ext, sched, chk = "", None, None, None
        plan = [MODEL, MODEL_BACKUP, MODEL_BACKUP, MODEL_BACKUP]
        for i, model in enumerate(plan):
            _job(k, status="running", step=("AI 正在排班（約 3～5 分鐘）" if i == 0 else f"檢查沒過，AI 正在修正（第 {i} 次{'，用 Opus' if model == MODEL_BACKUP else ''}）"))
            msg = _stream(model, system, messages, "high")
            text = "".join(b.text for b in msg.content if b.type == "text")
            with lk:
                st = _load(store, ym)
                _log_usage(st, "draft", _cost(model, msg.usage))
                _save(st)
            messages.append({"role": "assistant", "content": _blocks(msg.content)})
            try:
                ext_i, sched_i = _validate_draft(st, _extract_json(text))
            except (ValueError, json.JSONDecodeError) as ex:
                messages.append({"role": "user", "content": f"程式讀不到你的 JSON：{ex}。請重新輸出完整的 ```json 區塊（settings＋schedule），報告不用重寫。"})
                continue
            ext, sched = ext_i, sched_i
            form = _ext_from_form(st, cfg_load(store))
            if form:
                ext["shifts"] = dict(ext.get("shifts") or {}, **form["shifts"])   # 班別名以表單為準
                for kk in ("max_consec", "late_early_limit"):
                    ext[kk] = form[kk]
            body = text[:text.rfind("```json")].strip()
            report = body if i == 0 or not report else report + f"\n\n---\n### 修正說明（第 {i} 次）\n" + body
            tmp = dict(st, ext=ext, schedule=sched)
            chk = run_check(tmp)
            if chk["ok"]:
                break
            messages.append({"role": "user", "content": "檢查程式結果如下，請把 ❌ 全部改掉，再輸出完整的 ```json 區塊（settings＋schedule）。報告不用重寫，只要用幾句話說改了什麼。\n\n" + check_text(chk)})
        if sched is None:
            _job(k, status="error", step="AI 沒有產出可讀的班表，請再按一次「AI 排初稿」。")
            return
        with lk:
            st = _load(store, ym)
            if st.get("schedule"):
                st["history"] = (st.get("history") or [])[-19:] + [{"schedule": st["schedule"], "why": "AI 重排前"}]
            st.pop("draft", None)   # 這次是 AI 排的：說明用 AI 寫的，不要被「照目前班表重寫」蓋掉
            st.update(ext=ext, schedule=sched, report=report, check=chk, api_messages=[],
                      chat=(st.get("chat") or []) + [{"role": "ai", "text": "✅ 初稿排好了。" + ("檢查全部通過。" if chk["ok"] else f"還有 {len(chk['errors'])} 項 ❌ 沒改掉，請看下方檢查結果或跟我說要怎麼改。"), "ts": time.time()}])
            _save(st)
        _job(k, status="done", step="初稿完成")
    except Exception as ex:   # noqa: BLE001 — 背景工作要把錯誤留給畫面看
        _job(k, status="error", step=f"排班失敗：{type(ex).__name__}: {str(ex)[:300]}")


# ───────────────────────── 對話框 ─────────────────────────

# ───────────────────────── 程式排初稿（2026-11-05） ─────────────────────────
# 經營者「程式先排、AI 只做它擅長的事」：AI 先把店長寫的備註／店規翻成規則（約 10～30 秒，沒備註就不叫 AI），
# 程式（schedule_solver.py，OR-Tools）幾秒排完，排出來就過硬規則檢查，只給店長一份結果。
# 舊的「AI 排整月、檢查沒過再叫 Opus 修」留著當備援（_draft_worker_ai：沒裝 OR-Tools 或只有 Excel 設定表時用）。

RULES_SYS = """你把門市店長在排班設定裡寫的文字，翻成排班程式看得懂的規則。只回 JSON：{"rules":[…],"unapplied":["原文：為什麼沒辦法翻"]}
規則種類（days 一律是這個月的日期數字陣列；emp 是工號；shift 只能用給你的班別名稱）：
- {"type":"off","emp":"工號","days":[…]}：那幾天不能上班
- {"type":"work","emp":"工號","days":[…],"shift":"班別或null"}：那幾天一定要上班
- {"type":"only_shift","emp":"工號","shifts":[…],"days":[…]或null}：只能上這些班（null＝整個月）
- {"type":"prefer_shift","emp":"工號","shifts":[…],"days":[…]或null}：最好上這些班（「白天班最好」「盡量排早班」）
- {"type":"apart","emps":["工號","工號"],"days":[…]或null}：兩人不要同一天上班
- {"type":"demand","shift":"班別","days":[…],"count":人數}：這幾天這班要幾人（取代原本人數，例：週三晚班只排 1 人）
- {"type":"extra","shift":"班別","days":[…],"count":人數}：這幾天這班多加幾人（特殊日加人）
- {"type":"fixed_cell","emp":"工號","day":日期,"cell":"格子"}：那天固定用這個格子（上課、會議、跨店支援、包館等自訂時段）
格子寫法：時段用逗號分段，段後可加「@標籤」「#崗位」，當日標籤放最前面中括號，例「10:20-12:00,12:30-15:00@交流會議」「[支援]11:00-15:00#高雄店,15:30-19:30#高雄店」「休息日:13:00-15:00@交流會議」。標籤、崗位只能用給你的清單；崗位只在人實際在別家店時才填。自訂上班時間超過 4 小時的，中間要留 30 分鐘休息、拆成兩段（上班滿 4 小時後開始休，例：店長說「10:30 到 19:00」→「10:30-14:30,15:00-19:00」；104 不會自動扣休息，寫成一段會把休息也算進工時）；店長明講不休息或指定休息時間才照他說的。
店長用語對照：上課／南京上課／內訓→時段標籤「@台北內訓課」＋那段的崗位（例 10:20-12:00,12:30-15:00@台北內訓課#南京店）；店長會議／區會→@區會；交流會→@交流會議；事務→@事務班；經管會→@經管會；包館→@包館；到別家店支援→那幾段 #那家店＋@支援。會議、上課是時段標籤（@），不是當日標籤（[ ]）；當日標籤只用在整天性質（指休、特休、補休、彈性挪假、挪假補班、整天支援）。
朝會、巡店、消毒這類日常筆記不用翻，也不用列進 unapplied。看不懂或沒辦法用上面規則表達的，原文放進 unapplied 並說明原因。不要自己發明規則。"""


def _note_texts(st, cfg, ms):
    info = st["template"]["info"]
    L = []
    for e in info["emps"]:
        v = ms["emps"].get(e["id"]) or {}
        if (v.get("note") or "").strip():
            L.append(f"{e['id']} {e['name']}（{v.get('type', '正職')}）的個人備註：{v['note'].strip()}")
    if (cfg.get("rules") or "").strip():
        L.append("店規：" + cfg["rules"].strip().replace("\n", "；"))
    if (cfg.get("support") or "").strip():
        L.append("跨店支援夥伴：" + cfg["support"].strip().replace("\n", "；"))
    for x in ms.get("special") or []:
        bits = [b for b in (x.get("staff"), x.get("note")) if (b or "").strip()]
        if bits:
            L.append(f"{x['day']} 日特殊日：" + "；".join(bits))
    return L


def _daytype_shift(name, holiday_type, shifts):
    """「平日晚班」用在假日型的日子 → 「假日晚班」（有這個班才換）；反過來也是"""
    a, b = ("平日", "假日") if holiday_type else ("假日", "平日")
    alt = name.replace(a, b)
    return alt if a in name and alt in shifts else name


def _note_rules(st, cfg, ms):
    """備註／店規 → (rules, 沒套用的說明)；沒有備註就不叫 AI"""
    texts = _note_texts(st, cfg, ms)
    if not texts:
        return [], []
    info = st["template"]["info"]
    shifts = [r["name"] for r in cfg.get("shifts") or []]
    holi = _hol_days(ms)
    user = "\n".join([
        f"月份：{st['ym'][:4]} 年 {int(st['ym'][4:])} 月；日期（星期）：" + "、".join(f"{d['d']}({d['wd']}{'・當假日' if d['d'] in holi else ''})" for d in info["days"]),
        "（標「當假日」的平日，人力照假日班排，要加人就加假日的班）" if holi else "",
        "員工：" + "；".join(f"{e['id']} {e['name']}" for e in info["emps"]),
        "班別：" + "、".join(f"{r['name']}（" + "、".join(f"{a}-{b}" for a, b in r["segs"]) + "）" for r in cfg.get("shifts") or []),
        "可用標籤：" + "、".join(info["tags"]), "可用崗位：" + "、".join(info["positions"]),
        "本店是：" + info["title"], "", "店長寫的："] + texts)
    msg = _client_().messages.create(model=MODEL, max_tokens=6000, system=RULES_SYS,
                                     messages=[{"role": "user", "content": user}], output_config={"effort": "low"})
    with _lock(_key(st["store"], st["ym"])):
        s2 = _load(st["store"], st["ym"])
        _log_usage(s2, "rules", _cost(MODEL, msg.usage))
        _save(s2)
    txt = "".join(getattr(b, "text", "") for b in msg.content)
    m = re.search(r"\{.*\}", txt, re.S)
    try:
        d = json.loads(m.group(0)) if m else {}
    except json.JSONDecodeError:
        d = {}
    ids = {e["id"] for e in info["emps"]}
    days = {x["d"] for x in info["days"]}
    dtypes = set(info["dayTypes"])
    shmap = {r["name"]: [list(x) for x in r["segs"]] for r in cfg.get("shifts") or []}
    rules, bad = [], [str(x) for x in d.get("unapplied") or []]
    for r in d.get("rules") or []:
        t = r.get("type")
        try:
            if t not in schedule_solver.RULE_TYPES:
                raise ValueError("不認得的規則")
            if t == "apart":
                if len(r.get("emps") or []) < 2 or not set(r["emps"][:2]) <= ids:
                    raise ValueError("工號不對")
            elif t in ("demand", "extra"):
                if r.get("shift") not in shifts:
                    raise ValueError("班別不對")
                r["count"] = int(r["count"])
            elif r.get("emp") not in ids:
                raise ValueError("工號不對")
            if t == "fixed_cell":
                if int(r.get("day") or 0) not in days:
                    raise ValueError("日期不對")
                r["day"] = int(r["day"])
                parse_cell(r["cell"], shmap, dtypes)
            else:
                ds = r.get("days")
                if ds is None and t in ("only_shift", "prefer_shift", "apart"):
                    r["days"] = None
                else:
                    r["days"] = sorted({int(x) for x in ds or []} & days)
                    if not r["days"]:
                        raise ValueError("沒有日期")
            if t in ("only_shift", "prefer_shift"):
                r["shifts"] = [x for x in r.get("shifts") or [] if x in shifts]
                if not r["shifts"]:
                    raise ValueError("班別不對")
            if t == "work" and r.get("shift") not in shifts:
                r["shift"] = None
            if t in ("demand", "extra"):   # 平日班套到假日型的日子（或反過來）→ 換成同名另一種（2026-11-05 模擬：當假日的 31 號被加「平日晚班」）
                wkd = {x["d"] for x in info["days"] if x["wd"] in ("六", "日")} | holi
                groups = {}
                for dd in r["days"]:
                    groups.setdefault(_daytype_shift(r["shift"], dd in wkd, shifts), []).append(dd)
                for sh, dds in groups.items():
                    rules.append(dict(r, shift=sh, days=dds))
                continue
            rules.append(r)
        except (ValueError, TypeError, KeyError) as ex:
            bad.append(f"{json.dumps(r, ensure_ascii=False)[:120]}：{ex}")
    return rules, bad


def _rule_text(r, names):
    ds = lambda x: "整個月" if not x else _days_txt(x) + " 號"
    t = r["type"]
    if t == "off":
        return f"{names.get(r['emp'])} {ds(r['days'])}不排班"
    if t == "work":
        return f"{names.get(r['emp'])} {ds(r['days'])}一定上班" + (f"（{r['shift']}）" if r.get("shift") else "")
    if t == "only_shift":
        return f"{names.get(r['emp'])} {ds(r.get('days'))}只上{'／'.join(r['shifts'])}"
    if t == "prefer_shift":
        return f"{names.get(r['emp'])} {ds(r.get('days'))}盡量上{'／'.join(r['shifts'])}"
    if t == "apart":
        return f"{names.get(r['emps'][0])}、{names.get(r['emps'][1])} {ds(r.get('days'))}不同天上班"
    if t == "demand":
        return f"{ds(r['days'])}{r['shift']}排 {r['count']} 人"
    if t == "extra":
        return f"{ds(r['days'])}{r['shift']}多加 {r['count']} 人"
    if t == "fixed_cell":
        return f"{names.get(r['emp'])} {r['day']} 號固定：{r['cell']}"
    return str(r)


def _prev_tail(store, ym):
    """上個月最後連續上班幾天（七休一跨月用）；上個月沒班表就當 0"""
    p = _dir() / f"{_key(store, _prev_ym(ym))}.json"
    if not p.exists():
        return {}
    try:
        st = json.loads(p.read_text(encoding="utf-8"))
        info, sch = st["template"]["info"], st.get("schedule") or {}
        shifts = _ext(st).get("shifts") or {}
        dtypes = set(info["dayTypes"])
    except (OSError, json.JSONDecodeError, KeyError, TypeError):
        return {}
    out = {}
    for eid, row in sch.items():
        n = 0
        for d in reversed(info["days"]):
            try:
                c = parse_cell(row.get(str(d["d"]), ""), shifts, dtypes)
            except ValueError:
                break
            if not c["segs"]:
                break
            n += 1
        out[eid] = n
    return out


# 2026-11-10 經營者：只保留兼職優先（兼職排到希望的時數；正職本店排不滿的去別店支援）。ft 留著只為了讀舊存檔。
PLAN_NAME = {"ft": "正職優先（已停用）", "pt": "兼職優先"}
PLAN_DESC = {"ft": "正職先排滿標準時數，兼職只補剩下的班（兼職最低時數不保證）。",
             "pt": "兼職先排到希望的時數（沒填就用最低時數）；正職在本店排不滿的天數列為「可支援其他店」，可以去別店上班補時數。"}


def _solver_report(st, res, rules, bad, chk, secs_ai):
    info = st["template"]["info"]
    names = {e["id"]: e["name"] for e in info["emps"]}
    mode = res.get("mode")
    L = [f"### {PLAN_NAME.get(mode, '排班結果')}（程式排，{res['stats']['seconds'] + secs_ai:.0f} 秒）"]
    if mode in PLAN_DESC:
        L.append(PLAN_DESC[mode])
    L += [
         "✅ 硬規則檢查全部通過" if chk["ok"] else f"❌ 還有 {len(chk['errors'])} 項沒過（多半是人不夠，見下方缺口）"]
    L += ["", "| 夥伴 | 正職／兼職 | 排了 | 目標 | 最低 |", "|---|---|---|---|---|"]
    for eid, h in res["hours"].items():
        tg = f"{h['target']:g}h" if h["target"] else "—"
        mn = f"{h['min']:g}h" if h["min"] else "—"
        L.append(f"| {h['name']} | {'正職' if h['ft'] else '兼職'} | {h['hours']:g}h | {tg} | {mn} |")
    if res["gaps"]:
        L += ["", "**缺口（人不夠，沒有硬塞）**"]
        for g in res["gaps"]:
            L.append(f"- {g['day']} 號 {g['shift']}：要 {g['need']} 人、只排得到 {g['got']} 人")
        L.append("可以：請兼職多填可上日、調整指休、或在下方對話框跟 AI 說要找誰支援。")
    if res.get("extras"):
        n = sum(g["got"] - g["need"] for g in res["extras"])
        L += ["", f"**比設定多排的（共 {n} 人次）**：正職要排滿標準時數、兼職要排到最低時數，班不夠分才多出來。可以把這幾天改成特休、補休或事務班，或下修目標時數。"]
        L += [f"- {g['day']} 號：設定 {g['need']} 人、排了 {g['got']} 人（多在{'、'.join(g['shifts'])}）" for g in res["extras"]]
    if res.get("nuo"):
        L += ["", f"**正職{SUPPORT_TXT}的日子（本店排不滿；沒去別店的話這天標「彈性挪假」）**"]
        for eid, ds in res["nuo"].items():
            h = res["hours"][eid]
            L.append(f"- {names.get(eid)}：{'、'.join(str(d) for d in sorted(ds))} 號（共 {len(ds)} 天，本月 {h['hours']:g}h／目標 {h['target']:g}h）")
    if res.get("unmet"):
        L += ["", "**⚠ 備註／店規沒做到的（人手不夠才破例，請看要不要調整）**"]
        L += [f"- {_rule_text(u['rule'], names)}：{'、'.join(str(d) for d in u['days'])} 號沒做到" for u in res["unmet"]]
    if rules:
        L += ["", "**照備註／店規套用的**"] + ["- " + _rule_text(r, names) for r in rules]
    if bad:
        L += ["", "**備註裡沒辦法自動套用的（請看班表確認，要改就在下方對話框跟 AI 說）**"] + ["- " + b for b in bad]
    led = nuo_ledger(st)
    if any(r["open_rest"] or r["open_work"] for r in led):
        L += ["", "**挪假還沒補的（程式沒有自動補，要補哪天請跟 AI 說）**"]
        L += [f"- {r['name']}：{NUO_REST} {_nuo_txt(r['open_rest'], r['open_rest_h'])}、{NUO_WORK} {_nuo_txt(r['open_work'], r['open_work_h'])}" for r in led if r["open_rest"] or r["open_work"]]
    return "\n".join(L)


SUPPORT_TXT = "可支援其他店"   # 2026-11-10 經營者：「挪假 N 天」改叫「可支援其他店 N 天」（本店排不滿，可以去別店上班補時數）


def live_summary(st, chk=None):
    """目前班表的摘要（每人時數／目標、可支援其他店、缺人、多排、檢查）：一律從檢查結果算，改班後跟著變"""
    chk = chk or st.get("check")
    if not chk or not st.get("schedule") or not st.get("template"):
        return None
    ext = _ext(st)
    typ = {e.get("id"): e.get("type") for e in ext.get("employees") or []}
    sup = chk.get("support") or {}
    rows = [{"id": h["id"], "name": h["name"], "type": typ.get(h["id"], ""), "hours": h["hours"], "days": h["days"],
             "target": h.get("target"), "min": h.get("min"), "support": sup.get(h["id"]) or []} for h in chk.get("hours") or []]
    cover = chk.get("cover") or []
    return {"rows": rows, "has_demand": bool(cover),
            "gaps": [c for c in cover if c["got"] < c["need"]], "extras": [c for c in cover if c["got"] > c["need"]],
            "gap_n": sum(c["need"] - c["got"] for c in cover if c["got"] < c["need"]),
            "extra_n": sum(c["got"] - c["need"] for c in cover if c["got"] > c["need"]),
            "ok": chk["ok"], "errors": len(chk["errors"]), "warns": len(chk["warns"])}


def live_report(st):
    """排班結果說明：數字全部照「目前班表」重算（每次存檔都重寫），排初稿當時 AI 讀備註的結果附在後面"""
    sm = live_summary(st)
    if not sm:
        return st.get("report", "")
    dm = st.get("draft") or {}
    if dm:
        L = ["### 排班結果（兼職優先；照目前班表算，改班後會跟著更新）", PLAN_DESC["pt"]]
    else:   # 班表不是這一版程式排的（104 同步來的、或舊初稿）
        L = ["### 目前班表摘要（照目前班表算，改班後會跟著更新）"]
    L += ["✅ 硬規則檢查全部通過" if sm["ok"] else f"❌ 還有 {sm['errors']} 項沒過（見「檢查結果」）"]
    L += ["", "| 夥伴 | 正職／兼職 | 排了 | 目標 | 最低 | " + SUPPORT_TXT + " |", "|---|---|---|---|---|---|"]
    for r in sm["rows"]:
        tg = f"{float(r['target']):g}h" if r["target"] else "—"
        mn = f"{float(r['min']):g}h" if r["min"] else "—"
        L.append(f"| {r['name']} | {r['type'] or '—'} | {r['hours']:g}h | {tg} | {mn} | {str(len(r['support'])) + ' 天' if r['support'] else '—'} |")
    if sm["gaps"]:
        L += ["", f"**缺人的日子（共 {sm['gap_n']} 人次）**"]
        L += [f"- {c['label']}（{c['wd']}）：要 {c['need']} 人、排了 {c['got']} 人" for c in sm["gaps"]]
        L.append("可以：請兼職多填可上日、調整指休、請同區正職來支援，或在下方對話框跟 AI 說。")
    if sm["extras"]:
        L += ["", f"**比設定多排的日子（共 {sm['extra_n']} 人次）**"]
        L += [f"- {c['label']}（{c['wd']}）：設定 {c['need']} 人、排了 {c['got']} 人" for c in sm["extras"]]
    sup = [r for r in sm["rows"] if r["support"]]
    if sup:
        L += ["", f"**正職{SUPPORT_TXT}的日子**（本店這幾天不缺人，可以排去別店上班補時數；沒去的話，匯入 104 時這天標「彈性挪假」）"]
        for r in sup:
            tg = f"／目標 {float(r['target']):g}h" if r["target"] else ""
            L.append(f"- {r['name']}：{'、'.join(str(d) for d in r['support'])} 號（共 {len(r['support'])} 天；本月 {r['hours']:g}h{tg}）")
    try:   # 調店：去別店的看本店班表（崗位填別店的格子），來本店的看同區各店班表
        info = st["template"]["info"]
        names = {e["id"]: e["name"] for e in info["emps"]}
        shifts, dtypes = (_ext(st).get("shifts") or {}), set(info["dayTypes"])
        outs = []
        for e in info["emps"]:
            for d in info["days"]:
                c = parse_cell((st["schedule"].get(e["id"]) or {}).get(str(d["d"]), ""), shifts, dtypes)
                away = [x for x in c["segs"] if _elsewhere(x[3], info["title"]) and not BONUS_MEET_RE.search(x[2] or "")]
                if away:
                    outs.append(f"- {d['label']}（{d['wd']}）{names[e['id']]} 去{away[0][3]}（{away[0][0]}–{away[-1][1]}）")
        ins = []
        for g in (_guest_cells(st["store"], st["ym"]) or {}).values():
            for dd, c in sorted(g["days"].items()):
                if c.get("segs"):
                    ins.append(f"- {dd} 號 {g['name']} 來本店（{c['segs'][0][0]}–{c['segs'][-1][1]}）")
        if outs:
            L += ["", f"**去別店上班的（{len(outs)} 人次）**"] + outs
        if ins:
            L += ["", f"**別店來本店支援的（{len(ins)} 人次）**"] + ins
    except Exception:   # noqa: BLE001
        pass
    if dm.get("unmet"):
        L += ["", "**⚠ 排初稿時沒做到的備註／店規（人手不夠才破例）**"] + ["- " + x for x in dm["unmet"]]
    if dm.get("rules"):
        L += ["", "**排初稿時照備註／店規套用的**"] + ["- " + x for x in dm["rules"]]
    if dm.get("bad"):
        L += ["", "**備註裡沒辦法自動套用的（請看班表確認，要改就在下方對話框跟 AI 說）**"] + ["- " + x for x in dm["bad"]]
    led = nuo_ledger(st)
    if any(r["open_rest"] or r["open_work"] for r in led):
        L += ["", "**挪假還沒補的（程式沒有自動補，要補哪天請跟 AI 說）**"]
        L += [f"- {r['name']}：{NUO_REST} {_nuo_txt(r['open_rest'], r['open_rest_h'])}、{NUO_WORK} {_nuo_txt(r['open_work'], r['open_work_h'])}" for r in led if r["open_rest"] or r["open_work"]]
    return "\n".join(L)


def _plan_summary(res, chk):
    """方案比較用的數字：每人時數、缺口、多排、挪假、檢查"""
    return {"hours": {eid: round(h["hours"], 2) for eid, h in res["hours"].items()},
            "gaps": sum(g["need"] - g["got"] for g in res["gaps"]),
            "extras": sum(g["got"] - g["need"] for g in res.get("extras") or []),
            "nuo": {eid: sorted(ds) for eid, ds in (res.get("nuo") or {}).items()},
            "ok": chk["ok"], "errors": len(chk["errors"]), "warns": len(chk["warns"])}


def _draft_worker(store, ym):
    k = _key(store, ym)
    lk = _lock(k)
    with lk:
        st = _load(store, ym)
    cfg = cfg_load(store)
    try:
        import schedule_solver as _sv  # noqa: F401
        from ortools.sat.python import cp_model  # noqa: F401
    except ImportError:
        return _draft_worker_ai(store, ym)
    if not form_ready(st, cfg):          # 只有 Excel 設定表 → 走舊的 AI 排法
        return _draft_worker_ai(store, ym)
    try:
        ms = mset_view(st)
        info = st["template"]["info"]
        t0 = time.time()
        rules, bad = [], []
        if _note_texts(st, cfg, ms):
            _job(k, status="running", step="AI 正在讀備註和店規（約 10～30 秒）")
            try:
                rules, bad = _note_rules(st, cfg, ms)
            except Exception as ex:   # noqa: BLE001 — AI 連不上就不套備註，照樣排
                bad = [f"AI 讀備註失敗（{type(ex).__name__}），這次備註沒有套用"]
        secs_ai = time.time() - t0
        _job(k, status="running", step="程式正在排班（約 10 秒）")
        shmap = {r["name"]: [list(x) for x in r["segs"]] for r in cfg["shifts"]}
        dtypes = set(info["dayTypes"])
        fixed = {}
        for r in rules:
            if r["type"] == "fixed_cell":
                c = parse_cell(r["cell"], shmap, dtypes)
                fixed[(r["emp"], r["day"])] = (r["cell"], int(round(cell_hours(c) * 60)))
        y, mth = int(ym[:4]), int(ym[4:])
        hol = [d["d"] for d in info["days"] if d["wd"] not in ("六", "日") and _tw_off(dt.date(y, mth, d["d"]))]
        # 2026-11-06 經營者：同一份規則排兩種（A 正職優先、B 兼職優先），同時給店長看、選一個採用
        tail = _tail_merged(store, ym, {e["id"] for e in info["emps"]})
        plans = {}
        for mode in ("pt",):
            r_ = schedule_solver.solve(info, cfg, ms, rules=rules, hol_days=hol, prev_tail=tail,
                                       fixed_minutes={kk: v[1] for kk, v in fixed.items()}, mode=mode)
            for (eid, d), (cell, _) in fixed.items():
                r_["schedule"][eid][str(d)] = cell
            plans[mode] = r_
        res = plans["pt"]
        ext = _ext_from_form(st, cfg)
        _dm = {str(d): [n for n, c in res["need"][d].items() for _ in range(c)] for d in res["need"]}
        ext = {"demand_delta": {d: v for d, v in _dm.items() if sorted(v) != sorted((ext.get("demand") or {}).get(d) or [])}}
        ext["notes"] = [_rule_text(r, {e["id"]: e["name"] for e in info["emps"]}) for r in rules] + bad
        ext["from_form"] = False
        with lk:
            st = _load(store, ym)
            saved = {}
            for mode, r_ in plans.items():
                tmp = dict(st, ext=ext, schedule=r_["schedule"])
                chk_ = run_check(tmp)
                saved[mode] = {"schedule": r_["schedule"], "check": chk_, "report": _solver_report(tmp, r_, rules, bad, chk_, secs_ai),
                               "summary": _plan_summary(r_, chk_)}
            sched, chk, report = saved["pt"]["schedule"], saved["pt"]["check"], saved["pt"]["report"]
            if st.get("schedule"):
                st["history"] = (st.get("history") or [])[-19:] + [{"schedule": st["schedule"], "why": "重排前"}]
            names_ = {e["id"]: e["name"] for e in info["emps"]}
            st["plans"], st["plan"] = {}, "pt"   # 只排一個方案，不再有「兩個方案比較」
            st["draft"] = {"mode": "pt", "at": time.time(), "rules": [_rule_text(r, names_) for r in rules], "bad": list(bad),
                           "unmet": [f"{_rule_text(u['rule'], names_)}：{'、'.join(str(d) for d in u['days'])} 號沒做到" for u in res.get("unmet") or []]}
            nsup = sum(len(v) for v in (chk.get("support") or {}).values())
            msg = "✅ 排好了（兼職優先）。" + ("硬規則全部通過。" if chk["ok"] else f"有 {len(chk['errors'])} 項沒過（多半是人不夠），看「檢查結果」。") + (f"正職共有 {nsup} 人天{SUPPORT_TXT}，列在排班結果裡。" if nsup else "")
            if res.get("extras"):
                msg += f"有 {sum(g['got'] - g['need'] for g in res['extras'])} 人次比設定多（正職時數要排滿、班不夠分），列在報告裡。"
            if bad:
                msg += f"有 {len(bad)} 條備註沒辦法自動套用，列在報告裡。"
            if res.get("unmet"):
                msg += f"⚠ 有 {len(res['unmet'])} 條備註／店規因為人手不夠沒完全做到，列在報告裡。"
            st.update(ext=ext, schedule=sched, report=report, check=chk, api_messages=[],
                      chat=(st.get("chat") or []) + [{"role": "ai", "text": msg + "要改哪裡直接跟我說。", "ts": time.time()}])
            _save(st)
        _job(k, status="done", step="初稿完成")
    except Exception as ex:   # noqa: BLE001 — 背景工作要把錯誤留給畫面看
        _job(k, status="error", step=f"排班失敗：{type(ex).__name__}: {str(ex)[:300]}")


# ───────────────────────── 區域一起排（2026-11-06） ─────────────────────────
# 經營者：魔導師把同城的店一起排，自己店優先，缺人／多人才調店。流程：20 日前各店填好「本月設定」→
# 魔導師在「🗺 區域排班」按排初稿（A 正職優先、B 兼職優先兩個方案，整區一起比）→ 採用一個方案（一次寫進整區每家店的班表）→
# 店長再看自己店、微調、下載 104 匯入檔。調去別店的那天寫在夥伴自己店的班表（每段 @支援#那家店，104 規則），
# 被支援店的檢查會把外店調來的人算進現場人力（_guest_cells）。
REGIONS = {"台中": [1, 2], "台北": [3, 4, 7, 8, 11], "桃園": [9, 10], "測試": [13, 14]}
REGION_FILE = {"台中": "tc", "台北": "tp", "桃園": "ty", "測試": "ts"}
STORE_POS = {v: k for k, v in POS_STORE.items()}
# 測試區：13 號範本名稱是「高雄SKM店」、14 號是「台南測試店」，調店時崗位借用 104 清單裡的「高雄店」「台南店」（只在測試區這樣對應）
TEST_POS = {13: "高雄店", 14: "台南店"}


def _region_pos(rg, store):
    return TEST_POS.get(store, "") if rg == "測試" else STORE_POS.get(store, "")


def _pos_in_region(rg, pos):
    """崗位 → 這一區的店號（測試區用 TEST_POS，其他照 POS_STORE）"""
    if rg == "測試":
        return next((s for s, p in TEST_POS.items() if p == (pos or "").strip()), None)
    return _pos_store(pos)


def _region_of(store):
    return next((r for r, ids in REGIONS.items() if int(store) in ids), None)


def _rg_path(rg, ym):
    return _dir() / f"rg_{REGION_FILE[rg]}_{ym}.json"


def _guest_cells(store, ym, override=None):
    """同區別家店調來本店的夥伴 → {"外店:店號:工號": {"name", "days": {日: 格子（只留在本店的段）}}}
    override＝{店號: 班表}：還沒採用的區域方案用（不讀存檔）"""
    rg = _region_of(store)
    if not rg:
        return {}
    out = {}
    for s in REGIONS[rg]:
        if s == int(store):
            continue
        p = _dir() / f"{_key(s, ym)}.json"
        if not p.exists():
            continue
        try:
            st2 = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            continue
        tpl = st2.get("template")
        sched = (override or {}).get(s) if override and s in override else st2.get("schedule")
        if not tpl or not sched:
            continue
        shifts = (_ext(st2) or {}).get("shifts") or {}
        dtypes = set(tpl["info"]["dayTypes"])
        for e in tpl["info"]["emps"]:
            for d, v in (sched.get(e["id"]) or {}).items():
                if not v or "#" not in str(v):
                    continue
                try:
                    c = parse_cell(v, shifts, dtypes)
                except ValueError:
                    continue
                here = [[a, b, t, ""] for a, b, t, pos in c["segs"] if _pos_in_region(rg, pos) == int(store)]
                if here:
                    g = out.setdefault(f"外店:{s}:{e['id']}", {"name": f"{e['name']}（{_sname(s)}支援）", "days": {}})
                    g["days"][int(d)] = {"type": "", "segs": here, "dtag": ""}
    return out


def _region_ready(rg, ym):
    """整區每家店能不能排：要有 104 範本、店別基本設定的班別、本月設定的營收預估"""
    out = []
    for s in REGIONS[rg]:
        st = _load(s, ym)
        cfg = cfg_load(s)
        why = "" if st.get("template") else "還沒有 104 範本"
        if not why and not form_ready(st, cfg):
            why = "店別基本設定或本月設定還沒填好" if cfg.get("shifts") else "還沒填店別基本設定（班別、每天幾人）"
        out.append({"store": s, "name": _sname(s), "ready": not why, "why": why,
                    "emps": len(((st.get("template") or {}).get("info") or {}).get("emps") or []),
                    "region": (st.get("region") or {}).get("at")})
    return out


def _region_ctx(rg, ym):
    stores = REGIONS[rg]
    sts = {s: _load(s, ym) for s in stores}
    cfgs = {s: cfg_load(s) for s in stores}
    mss = {s: mset_view(sts[s]) for s in stores}
    return stores, sts, cfgs, mss


def _region_got(sts, s, sched, moves):
    """這家店每天在店人數（自己的人在本店上班＋外店調來的）"""
    days = sts[s]["template"]["info"]["days"]
    got = [0] * len(days)
    for row in sched.values():
        for i, d in enumerate(days):
            v = row.get(str(d["d"])) or ""
            if v and "@支援#" not in v and v.split("]")[-1] not in ("例假日", "休息日", "國定假日") and not v.startswith("[特休]"):
                got[i] += 1
    idx = {d["d"]: i for i, d in enumerate(days)}
    for mv in moves:
        if mv["to"] == s and mv["day"] in idx:
            got[idx[mv["day"]]] += 1
    return got


def _region_moves(rg, sts, scheds):
    """從各店班表找出調店（格子裡 @支援# 的崗位是同區別家店）"""
    out = []
    for s, sched in scheds.items():
        info = sts[s]["template"]["info"]
        names = {e["id"]: e["name"] for e in info["emps"]}
        for eid, row in sched.items():
            for d, v in row.items():
                m = re.search(r"@支援#([^,@#]+)", v or "")
                if not m:
                    continue
                to = _pos_in_region(rg, m.group(1))
                if to and to != s and to in scheds:
                    segs = [x.split("@")[0] for x in v.split(",")]
                    sh = next((r["name"] for r in cfg_load(to).get("shifts") or [] if [f"{a}-{b}" for a, b in r["segs"]] == segs), "")
                    out.append({"store": s, "emp": eid, "name": names.get(eid, eid), "day": int(d), "to": to,
                                "shift": sh or "、".join(segs), "from_name": _sname(s), "to_name": _sname(to)})
    return sorted(out, key=lambda x: (x["day"], x["store"]))


def _region_eval(rg, ym, sts, P):
    """方案改過格子後：重算各店檢查、每天人數、調店清單（報告保留排班當時的，前面加註）"""
    scheds = {int(k): v["schedule"] for k, v in P["stores"].items()}
    P["moves"] = _region_moves(rg, sts, scheds)
    for k, v in P["stores"].items():
        s = int(k)
        tmp = dict(sts[s], ext=v["ext"], schedule=v["schedule"])
        chk = run_check(tmp, guests=_guest_cells(s, ym, scheds), sibs=scheds)
        v["check"] = chk
        v["got"] = _region_got(sts, s, v["schedule"], P["moves"])
        sm = v.get("summary") or {}
        sm.update(ok=chk["ok"], errors=len(chk["errors"]), warns=len(chk["warns"]))
        v["summary"] = sm


def _region_solve(rg, ym, ctx, rules, bad, secs_ai, modes, jk):
    """用同一份規則整區排（modes 裡每個方案各排一次）→ {mode: 方案}"""
    stores, sts, cfgs, mss = ctx
    y, mth = int(ym[:4]), int(ym[4:])
    units, fixed = [], {}
    for s in stores:
        info = sts[s]["template"]["info"]
        shmap = {r["name"]: [list(x) for x in r["segs"]] for r in cfgs[s]["shifts"]}
        dtypes = set(info["dayTypes"])
        fx = {}
        for r in rules.get(s, []):
            if r["type"] == "fixed_cell":
                c = parse_cell(r["cell"], shmap, dtypes)
                fx[(r["emp"], r["day"])] = (r["cell"], int(round(cell_hours(c) * 60)))
        fixed[s] = fx
        units.append({"store": s, "pos": _region_pos(rg, s), "info": info, "cfg": cfgs[s], "ms": mss[s], "rules": rules.get(s, []),
                      "hol_days": [d["d"] for d in info["days"] if d["wd"] not in ("六", "日") and _tw_off(dt.date(y, mth, d["d"]))],
                      "prev_tail": _tail_merged(s, ym, {e["id"] for e in info["emps"]}), "fixed_minutes": {kk: v[1] for kk, v in fx.items()}})
    out = {}
    for i, mode in enumerate(modes):
        _job(jk, status="running", step=f"程式正在排整區（{PLAN_NAME[mode]}，{i + 1}／{len(modes)}，約 1～3 分鐘）")
        r = schedule_solver.solve_multi(units, mode=mode, lend=True, time_limit=90)
        for u in r["stores"]:
            for (eid, d), (cell, _) in fixed[u["store"]].items():
                u["schedule"][eid][str(d)] = cell
        scheds = {u["store"]: u["schedule"] for u in r["stores"]}
        P = {"stats": r["stats"], "moves": [dict(mv, from_name=_sname(mv["store"]), to_name=_sname(mv["to"])) for mv in r["moves"]],
             "stores": {}, "history": [],
             "pos_block": [{"from": a, "to": b, "from_name": _sname(a), "to_name": _sname(b), "pos": _region_pos(rg, b)} for a, b in r.get("pos_block") or []]}
        for u in r["stores"]:
            s = u["store"]
            ext = _ext_from_form(sts[s], cfgs[s])
            _dm = {str(d): [n for n, c in u["need"][d].items() for _ in range(c)] for d in u["need"]}
            ext = {"demand_delta": {d: v for d, v in _dm.items() if sorted(v) != sorted((ext.get("demand") or {}).get(d) or [])}}
            names = {e["id"]: e["name"] for e in sts[s]["template"]["info"]["emps"]}
            ext["notes"] = [_rule_text(x, names) for x in rules.get(s, [])] + bad.get(s, [])
            ext["from_form"] = False
            tmp = dict(sts[s], ext=ext, schedule=u["schedule"])
            chk = run_check(tmp, guests=_guest_cells(s, ym, scheds), sibs=scheds)
            rep = _solver_report(tmp, dict(u, stats=r["stats"]), rules.get(s, []), bad.get(s, []), chk, secs_ai)
            outs = [mv for mv in P["moves"] if mv["store"] == s]
            ins = [mv for mv in P["moves"] if mv["to"] == s]
            if outs or ins:
                rep += "\n\n**區域調店（魔導師整區一起排）**"
                rep += "".join(f"\n- {mv['day']} 號 {mv['name']} 去{mv['to_name']}支援（{mv['shift']}）" for mv in outs)
                rep += "".join(f"\n- {mv['day']} 號 {mv['from_name']}的 {mv['name']} 來本店支援（{mv['shift']}）" for mv in ins)
            need = [sum(u["need"][d["d"]].values()) for d in sts[s]["template"]["info"]["days"]]
            P["stores"][str(s)] = {"schedule": u["schedule"], "ext": ext, "check": chk, "report": rep, "need": need,
                                   "draft": {"rules": [_rule_text(x, names) for x in rules.get(s, [])], "bad": list(bad.get(s, [])),
                                             "unmet": [f"{_rule_text(x['rule'], names)}：{'、'.join(str(d) for d in x['days'])} 號沒做到" for x in u.get("unmet") or []]},
                                   "got": _region_got(sts, s, u["schedule"], P["moves"]),
                                   "summary": _plan_summary(u, chk), "names": names}
        out[mode] = P
    return out


def _region_rules(rg, ym, ctx, jk):
    """各店備註／店規請 AI 翻成規則（並行）→ (rules, bad, 秒)"""
    stores, sts, cfgs, mss = ctx
    t0 = time.time()
    _job(jk, status="running", step="AI 正在讀各店的備註和店規（約 10～30 秒）")
    res = {}

    def run(s):
        if not _note_texts(sts[s], cfgs[s], mss[s]):
            res[s] = ([], [])
            return
        try:
            res[s] = _note_rules(sts[s], cfgs[s], mss[s])
        except Exception as ex:   # noqa: BLE001 — AI 連不上就不套備註，照樣排
            res[s] = ([], [f"AI 讀備註失敗（{type(ex).__name__}），這次備註沒有套用"])
    ths = [threading.Thread(target=run, args=(s,), daemon=True) for s in stores]
    for th in ths:
        th.start()
    for th in ths:
        th.join()
    return {s: res[s][0] for s in stores}, {s: res[s][1] for s in stores}, time.time() - t0


def _merge_rules(rules, extra):
    out = {s: list(v) for s, v in rules.items()}
    for r in extra or []:
        out.setdefault(int(r["store"]), []).append({k: v for k, v in r.items() if k != "store"})
    return out


def _region_worker(rg, ym):
    jk = f"rg:{rg}:{ym}"
    try:
        ctx = _region_ctx(rg, ym)
        stores = ctx[0]
        p = _rg_path(rg, ym)
        old = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
        extra = old.get("extra_rules") or []      # 魔導師在對話框加的條件，重排時照樣套
        rules, bad, secs_ai = _region_rules(rg, ym, ctx, jk)
        plans = _region_solve(rg, ym, ctx, _merge_rules(rules, extra), bad, secs_ai, ("pt",), jk)
        out = {"region": rg, "ym": ym, "at": dt.datetime.now().isoformat(timespec="seconds"), "secs_ai": round(secs_ai, 1),
               "stores": stores, "plans": plans, "picked": None, "rules": {str(s): v for s, v in rules.items()},
               "bad": {str(s): v for s, v in bad.items()}, "extra_rules": extra, "chat": old.get("chat") or []}
        atomic_write_json(p, out)
        _job(jk, status="done", step="整區初稿完成")
    except Exception as ex:   # noqa: BLE001 — 背景工作要把錯誤留給畫面看
        _job(jk, status="error", step=f"區域排班失敗：{type(ex).__name__}: {str(ex)[:300]}")


# ── 魔導師對話框（2026-11-06 經營者：魔導師也要有 AI 對話框提出調整）：改的是「目前這個方案」，按「採用」才寫進各店 ──
RG_CHAT_RULES = """
## 這次的工作：魔導師在「區域排班」對話框要你調整整區的方案
- 你看到的是整區排好的班表（兼職優先：兼職排到希望的時數，正職本店排不滿的去同區別店支援）（同區每家店），還沒寫進各店；魔導師按「採用」才會寫進去。
- 工具：
 1. rg_apply：直接改某店某人某天的格子（一次可改多格、可跨店）。調去別家店上班＝在夥伴「自己店」的格子寫那家店的班別時段，每段加「@支援#崗位」（崗位見下方各店的調店崗位），例「13:00-15:00@支援#台南店,15:30-21:30@支援#台南店」；原本被調走的那天改回自己店上班就寫自己店的班別名稱。
 2. rg_rerun：加條件後讓程式整區重排這個方案（約 1～3 分鐘），適合「某人整月不要調店」「某人某幾天一定要調去某店」「某人某幾天不能上班」「某店某天某班要幾人」這類條件。條件會記住，之後按「整區排初稿」也會照樣套。
 3. rg_undo：復原上一次改動。
- 所有夥伴（正職、兼職）都可以調去同城別家店（2026-11-11 經營者固定規定）；魔導師說某人不要調店，就用 rg_rerun 加條件。
- 回覆規矩：每句話都要給答案；能做就直接做好再說明；做不到說清楚卡在哪一條規則並給 2～3 個做法；改完用一兩句白話說改了什麼、檢查結果（各店 ❌／⚠ 有幾項）；不要貼整張表、不要提程式內部用語；用繁體中文。
- 檢查結果有新的 ❌ 一定要再改到沒有才回覆（人真的不夠時說清楚缺在哪天哪店）。
"""

RG_TOOLS = [
    {"name": "rg_apply", "description": "改整區方案的格子（可跨店，一次可改多格）。回傳改後各店檢查結果。",
     "input_schema": {"type": "object", "additionalProperties": False, "required": ["changes"], "properties": {
         "changes": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                     "required": ["store", "emp", "day", "value"], "properties": {
                         "store": {"type": "integer", "description": "夥伴自己店的店號"},
                         "emp": {"type": "string", "description": "104 工號"},
                         "day": {"type": "integer", "description": "日期 1～31"},
                         "value": {"type": "string", "description": "格子新內容（照格子寫法）"}}}}}},
     "strict": True},
    {"name": "rg_rerun", "description": "加條件後整區重排這個方案（約 1～3 分鐘）。條件種類：off（某人某幾天不能上班）、work（某人某幾天一定上班，可指定班別）、no_lend（某人〔某幾天，null＝整月〕不調去別店）、lend（某人某幾天一定調去某店，可指定那家店的班別）、demand（某店某幾天某班要幾人）、extra（某店某幾天某班多加幾人）。",
     "input_schema": {"type": "object", "additionalProperties": False, "required": ["rules"], "properties": {
         "rules": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                   "required": ["type", "store", "emp", "days", "to", "shift", "count"], "properties": {
                       "type": {"type": "string", "enum": ["off", "work", "no_lend", "lend", "demand", "extra"]},
                       "store": {"type": "integer", "description": "夥伴自己店（demand／extra＝那家店）"},
                       "emp": {"type": ["string", "null"], "description": "工號（demand／extra 填 null）"},
                       "days": {"type": ["array", "null"], "items": {"type": "integer"}, "description": "日期；null＝整個月（只有 no_lend 可以）"},
                       "to": {"type": ["integer", "null"], "description": "lend：調去哪家店"},
                       "shift": {"type": ["string", "null"], "description": "班別名稱（work／lend／demand／extra 可填）"},
                       "count": {"type": ["integer", "null"], "description": "demand／extra 的人數"}}}}}},
     "strict": True},
    {"name": "rg_undo", "description": "復原上一次改動（rg_apply 或 rg_rerun）。",
     "input_schema": {"type": "object", "additionalProperties": False, "properties": {}}, "strict": True},
]


def _rg_check_text(rg, P):
    lines = []
    for k, v in P["stores"].items():
        c = v["check"]
        lines.append(f"【{k} {_sname(int(k))}】❌ {len(c['errors'])} 項、⚠ {len(c['warns'])} 項")
        lines += ["❌ " + x for x in c["errors"][:12]] + ["⚠ " + x for x in c["warns"][:8]]
    lines.append(f"【調店】{len(P['moves'])} 人次：" + "；".join(f"{m['day']} 號 {m['name']}（{m['from_name']}→{m['to_name']}）" for m in P["moves"][:40]))
    return "\n".join(lines)


def _rg_table_text(sts, P):
    out = []
    for k, v in P["stores"].items():
        s = int(k)
        info = sts[s]["template"]["info"]
        out.append(f"【{s} {_sname(s)}】每天需要：" + "、".join(f"{d['d']}={n}" for d, n in zip(info["days"], v["need"])))
        out.append("　每天在店（含外店調來）：" + "、".join(f"{d['d']}={n}" for d, n in zip(info["days"], v["got"])))
        for e in info["emps"]:
            row = v["schedule"].get(e["id"]) or {}
            out.append(f"{s}／{e['id']} {e['name']}：" + "；".join(f"{d['d']}={row.get(str(d['d']), '')}" for d in info["days"]))
    return "\n".join(out)


def _rg_system(rg, ym, ctx):
    stores, sts, cfgs, mss = ctx
    parts = [RULES, RG_CHAT_RULES, f"\n## 區域：{rg}（{ym[:4]} 年 {int(ym[4:])} 月）"]
    for s in stores:
        info = sts[s]["template"]["info"]
        parts.append(f"\n### {s} {_sname(s)}（104 班表：{info['title']}；調來這家店時崗位填「{_region_pos(rg, s)}」）")
        parts.append("班別：" + "、".join(f"{r['name']}（" + "、".join(f"{a}-{b}" for a, b in r["segs"]) + f"；平日 {r.get('wd', 0)} 人、假日 {r.get('we', 0)} 人）" for r in cfgs[s]["shifts"]))
        parts.append("夥伴：" + "；".join(f"{e['id']} {e['name']}（{(mss[s]['emps'].get(e['id']) or {}).get('type', '正職')}" + ("" if can_move(mss[s]['emps'].get(e['id'])) else "，不調店") + "）" for e in info["emps"]))
        parts.append("本月設定：\n" + marks_text(sts[s]))
    parts.append("\n日期：" + "、".join(d["label"] for d in sts[stores[0]]["template"]["info"]["days"]))
    return "\n".join(parts)


def _rg_rule_check(rg, r, sts):
    s = int(r["store"])
    if s not in REGIONS[rg]:
        raise ValueError(f"{s} 號店不在{rg}區")
    info = sts[s]["template"]["info"]
    days = {d["d"] for d in info["days"]}
    t = r["type"]
    out = {"type": t, "store": s}
    if t in ("demand", "extra"):
        shifts = {x["name"] for x in cfg_load(s).get("shifts") or []}
        if r.get("shift") not in shifts:
            raise ValueError(f"{_sname(s)}沒有「{r.get('shift')}」這個班別")
        out.update(shift=r["shift"], count=int(r.get("count") or 0), days=sorted(set(r.get("days") or []) & days))
    else:
        nm = {e["id"]: e["name"] for e in info["emps"]}
        if r.get("emp") not in nm:
            raise ValueError(f"{_sname(s)}沒有工號 {r.get('emp')}")
        out["emp"], out["name"] = r["emp"], nm[r["emp"]]
        out["days"] = None if (t == "no_lend" and r.get("days") is None) else sorted(set(r.get("days") or []) & days)
        if t in ("work", "lend") and r.get("shift"):
            out["shift"] = r["shift"]
        if t == "lend":
            if int(r.get("to") or 0) not in REGIONS[rg] or int(r["to"]) == s:
                raise ValueError("lend 要指定同區另一家店")
            if not can_move(((sts[s].get("mset") or {}).get("emps") or {}).get(r["emp"])):
                raise ValueError(f"{_sname(s)} {nm[r['emp']]} 本月設定「調店」是「否」，不能調去別店；要調請店長先改成「可」再存檔")
            out["to"] = int(r["to"])
    if out.get("days") == [] and t != "no_lend":
        raise ValueError("沒有有效的日期")
    return out


def _region_chat_worker(rg, ym, mode, text):
    jk = f"rg:{rg}:{ym}"
    p = _rg_path(rg, ym)
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        P = data["plans"][mode]
        ctx = _region_ctx(rg, ym)
        stores, sts = ctx[0], ctx[1]
        msgs = (data.get("api_messages") or {}).get(mode) or []
        if len(msgs) > MAX_API_MESSAGES:
            msgs = []
        msgs.append({"role": "user", "content": f"（目前方案：{PLAN_NAME[mode]}）\n（整區班表）\n{_rg_table_text(sts, P)}\n\n（目前檢查）\n{_rg_check_text(rg, P)}\n\n魔導師：{text}"})
        changed, reran = 0, False
        system = _rg_system(rg, ym, ctx)
        for i in range(8):
            _job(jk, status="running", step=f"AI 正在想（第 {i + 1} 輪）…" if i else "AI 正在想…")
            msg = _client_().messages.create(model=CHAT_MODEL, max_tokens=16000, system=system, tools=RG_TOOLS,
                                             output_config={"effort": "medium"}, cache_control={"type": "ephemeral"}, messages=msgs)
            c = _cost(CHAT_MODEL, msg.usage)
            with open(_dir() / "usage.jsonl", "a", encoding="utf-8") as f:
                f.write(json.dumps(dict(c, kind="rg_chat", region=rg, ym=ym, ts=dt.datetime.now().isoformat(timespec="seconds")), ensure_ascii=False) + "\n")
            data["cost"] = round((data.get("cost") or 0) + c["twd"], 2)
            msgs.append({"role": "assistant", "content": _blocks(msg.content)})
            if msg.stop_reason != "tool_use":
                break
            results = []
            for b in msg.content:
                if b.type != "tool_use":
                    continue
                try:
                    if b.name == "rg_undo":
                        h = P.get("history") or []
                        if not h:
                            out = "沒有可以復原的改動"
                        else:
                            prev = h.pop()
                            for k, sch in prev["schedules"].items():
                                P["stores"][k]["schedule"] = sch
                            if "extra_rules" in prev:   # 復原「加條件重排」：條件也要退回去，不然下次重排又套用
                                data["extra_rules"] = prev["extra_rules"]
                            _region_eval(rg, ym, sts, P)
                            out = "已復原。\n" + _rg_check_text(rg, P)
                    elif b.name == "rg_apply":
                        chg = b.input.get("changes") or []
                        for c_ in chg:
                            s = int(c_["store"])
                            if str(s) not in P["stores"]:
                                raise ValueError(f"{s} 號店不在這一區")
                            info = sts[s]["template"]["info"]
                            if c_["emp"] not in {e["id"] for e in info["emps"]}:
                                raise ValueError(f"{_sname(s)}沒有工號 {c_['emp']}")
                            if int(c_["day"]) not in {d["d"] for d in info["days"]}:
                                raise ValueError(f"這個月沒有 {c_['day']} 號")
                            parse_cell(c_["value"], P["stores"][str(s)]["ext"].get("shifts") or {}, set(info["dayTypes"]))
                            if "@支援#" in str(c_["value"]) and not can_move(((sts[s].get("mset") or {}).get("emps") or {}).get(c_["emp"])):
                                nm_ = next(e["name"] for e in info["emps"] if e["id"] == c_["emp"])
                                raise ValueError(f"{_sname(s)} {nm_} 本月設定「調店」是「否」，不能調去別店；要調請店長先改成「可」再存檔")
                        P.setdefault("history", []).append({"schedules": {k: copy.deepcopy(v["schedule"]) for k, v in P["stores"].items()}, "why": "對話：" + text[:40]})
                        P["history"] = P["history"][-20:]
                        for c_ in chg:
                            P["stores"][str(int(c_["store"]))]["schedule"].setdefault(c_["emp"], {})[str(int(c_["day"]))] = c_["value"]
                        _region_eval(rg, ym, sts, P)
                        changed += len(chg)
                        out = "已改。\n" + _rg_check_text(rg, P)
                    else:
                        new = [_rg_rule_check(rg, r, sts) for r in b.input.get("rules") or []]
                        hist = (P.get("history") or []) + [{"schedules": {k: copy.deepcopy(v["schedule"]) for k, v in P["stores"].items()}, "why": "重排前",
                                                            "extra_rules": list(data.get("extra_rules") or [])}]
                        data["extra_rules"] = (data.get("extra_rules") or []) + new
                        rules = {int(k): v for k, v in (data.get("rules") or {}).items()}
                        bad = {int(k): v for k, v in (data.get("bad") or {}).items()}
                        P2 = _region_solve(rg, ym, ctx, _merge_rules(rules, data["extra_rules"]), bad, 0, (mode,), jk)[mode]
                        P2["history"] = hist[-20:]
                        data["plans"][mode] = P = P2
                        reran = True
                        out = "已加條件並整區重排。\n" + _rg_check_text(rg, P)
                except ValueError as ex:
                    out = f"這次沒有改：{ex}"
                results.append({"type": "tool_result", "tool_use_id": b.id, "content": out})
            _job(jk, status="running", step=f"已改 {changed} 格{'、已重排' if reran else ''}，AI 正在看檢查結果…")
            msgs.append({"role": "user", "content": results})
        reply = "".join(b.text for b in msg.content if b.type == "text").strip() or "（我剛剛沒有產生回覆，請再說一次。）"
        if (changed or reran) and (data.get("picked") or {}).get("mode") == mode:
            data["picked"]["stale"] = True
            reply += "\n\n（這個方案已採用過，改的內容要再按一次「採用」才會寫進各店。）"
        data.setdefault("api_messages", {})[mode] = msgs
        data["chat"] = (data.get("chat") or []) + [{"role": "me", "text": text, "mode": mode, "ts": time.time()},
                                                   {"role": "ai", "text": reply, "mode": mode, "changed": changed, "reran": reran, "ts": time.time()}]
        atomic_write_json(p, data)
        _job(jk, status="done", step="")
    except Exception as ex:   # noqa: BLE001
        _job(jk, status="error", step=f"對話失敗：{type(ex).__name__}: {str(ex)[:300]}")


def _region_public(rg, ym):
    p = _rg_path(rg, ym)
    data = json.loads(p.read_text(encoding="utf-8")) if p.exists() else None
    with _glock:
        job = dict(_jobs.get(f"rg:{rg}:{ym}") or {})
    if job.get("status") == "running":
        job["secs"] = int(time.time() - job.get("t0", time.time()))
    days = None
    meta = {}   # 2026-11-09 經營者：調店清單右邊列各店每人總時數 → 帶正職／兼職、目標或最低時數、調店
    for s in REGIONS[rg]:
        st = _load(s, ym)
        if st.get("template"):
            days = days or st["template"]["info"]["days"]
            try:
                meta[str(s)] = {eid: {"type": v.get("type", "正職"), "target": v.get("target"), "min": v.get("min_hours"),
                                      "move": "可" if can_move(v) else "否"} for eid, v in mset_view(st)["emps"].items()}
            except Exception:   # noqa: BLE001 — 只是補充資訊，讀不到就不顯示目標
                meta[str(s)] = {}
    plans = {}
    for mode, P in ((data or {}).get("plans") or {}).items():
        plans[mode] = {"name": PLAN_NAME[mode], "desc": PLAN_DESC[mode], "stats": P["stats"], "moves": P["moves"], "pos_block": P.get("pos_block") or [],
                       "stores": {k: {"schedule": v["schedule"], "need": v["need"], "got": v["got"], "summary": v["summary"],
                                      "names": v["names"], "ok": v["check"]["ok"], "errors": v["check"]["errors"][:20],
                                      "warns": len(v["check"]["warns"]),
                                      "hours": [dict(h, **((meta.get(k) or {}).get(h.get("id")) or {})) for h in v["check"].get("hours") or []]}
                                  for k, v in P["stores"].items()}}
    return {"ok": True, "region": rg, "ym": ym, "regions": REGIONS, "stores": _region_ready(rg, ym), "days": days, "job": job,
            "plans": plans, "at": (data or {}).get("at"), "picked": (data or {}).get("picked"),
            "chat": ((data or {}).get("chat") or [])[-80:], "cost": (data or {}).get("cost") or 0,
            "extraRules": [_rule_text_rg(r) for r in (data or {}).get("extra_rules") or []],
            "storeNames": {s: _sname(s) for s in REGIONS[rg]}}


def _rule_text_rg(r):
    t, s = r["type"], int(r["store"])
    days = "整個月" if not r.get("days") else _days_txt(r["days"]) + " 號"
    nm = f"{_sname(s)} {r.get('name') or r.get('emp') or ''}".strip()
    if t == "no_lend":
        return f"{nm} {days}不調去別店"
    if t == "lend":
        return f"{nm} {days}調去{_sname(r['to'])}" + (f"（{r['shift']}）" if r.get("shift") else "")
    if t == "off":
        return f"{nm} {days}不上班"
    if t == "work":
        return f"{nm} {days}一定上班" + (f"（{r['shift']}）" if r.get("shift") else "")
    if t == "demand":
        return f"{_sname(s)} {days}{r['shift']}排 {r['count']} 人"
    if t == "extra":
        return f"{_sname(s)} {days}{r['shift']}多加 {r['count']} 人"
    return str(r)


def _region_pick(rg, ym, mode):
    p = _rg_path(rg, ym)
    if not p.exists():
        return 400, {"ok": False, "error": "還沒有整區初稿，請先按「整區排初稿」"}
    data = json.loads(p.read_text(encoding="utf-8"))
    P = (data.get("plans") or {}).get(mode)
    if not P:
        return 400, {"ok": False, "error": "沒有這個方案"}
    with _glock:   # 2026-11-11：某家店的 AI 還在跑（對話、排初稿）時採用，會被它結束時整份存回的舊班表蓋掉
        busy = [_sname(s) for s in REGIONS[rg] if (_jobs.get(_key(s, ym)) or {}).get("status") == "running"]
    if busy:
        return 409, {"ok": False, "error": "這幾家店的 AI 還在處理（對話或排初稿），等它完成再採用：" + "、".join(busy)}
    now = dt.datetime.now().isoformat(timespec="seconds")
    for s in REGIONS[rg]:          # 先把整區的班表都寫進去，再逐店重算檢查（外店支援要讀到別家新的班表）
        x = P["stores"].get(str(s))
        if not x:
            continue
        with _lock(_key(s, ym)):
            st = _load(s, ym)
            if st.get("schedule"):
                st["history"] = (st.get("history") or [])[-19:] + [{"schedule": st["schedule"], "why": "區域排班前"}]
            st.update(schedule=copy.deepcopy(x["schedule"]), ext=x["ext"], report=x["report"], plans={}, plan=None,
                      region={"name": rg, "mode": mode, "at": now},
                      draft=dict(x.get("draft") or {"rules": list((x.get("ext") or {}).get("notes") or []), "bad": [], "unmet": []}, mode=mode, at=time.time(), region=rg))
            st["chat"] = (st.get("chat") or []) + [{"role": "ai", "ts": time.time(),
                                                    "text": f"魔導師用「{PLAN_NAME[mode]}」排好整區（{rg}）的班表，本店班表已換成這一版（調店的格子標「支援＋店名」）。要改哪裡直接跟我說；換錯可按「復原上一步」。"}]
            _save(st)
    for s in REGIONS[rg]:
        with _lock(_key(s, ym)):
            st = _load(s, ym)
            if st.get("template") and st.get("schedule"):
                st["check"] = run_check(st)
                _save(st)
    data["picked"] = {"mode": mode, "at": now, "stale": False}
    atomic_write_json(p, data)
    return 200, _region_public(rg, ym)


CHAT_RULES = """
## 這次的工作：店長在對話框要你改班或問問題
- 用 apply_changes 改格子（程式會回檢查結果），用 undo 復原上一次改動。班表、檢查結果每句話都會附在店長訊息前面。
- 一定要守的回覆規矩：
 1. 每句話都要給答案，不准只回「我不懂」「做不到」。
 2. 能做就直接做好（例：改了某人的班造成缺人，就順手找最合適的人補上），再說明。做不到或會違規：說清楚卡在哪一條規則，給 2～3 個可以選的做法。
 3. 意思不明確（沒說哪一天、哪個人）：先說「我理解成：…」，照最合理的理解做，最後問一句要不要改。
 4. 改完用一兩句白話說「改了什麼、誰補上、檢查結果」，不要貼整張表，不要提程式內部用語（格子文法、工具、JSON）。
 5. 只處理本店排班；其他話題回「我只負責本店排班，其他問題請洽總部」。
 6. 檢查結果有新的 ❌ 一定要再改到沒有才回覆；⚠ 要告訴店長。
 7. 回覆用繁體中文，不要夾雜英文字。
- 「本月設定」是店長填的設定表；格子開頭的 [指休]、[包館] 等是 104 班表上的標籤。說「設定表上是指休」只能根據本月設定的【指休】日期，不可以用格子標籤推論。
- 店長問「別家店能不能派人來補班」「缺的人能不能請 X 號店支援」：一定先用 support_eval 評估（約 1～3 分鐘）。它會把本店和那幾家店的人放在一起重新排一次，看本店補不補得滿、來支援的店會不會反而缺人——**只是評估，不會改任何一家店的班表**。你也**絕對不可以**改別家店的班表（apply_changes 只能改本店的人；別家店的人不在本店名單上，也寫不進本店班表）。回覆時用白話說：本店缺的日子能不能補滿、建議誰哪天來支援、對方店會不會因此缺人、他們的工時會不會超過或不足、有沒有 104 崗位要先加；最後提醒「真的要調，請魔導師在『區域排班』採用，或請兩家店長自己協調後各改各的班表」。店長沒說哪家店：先評估同區的店；本店不在任何一區（或同區沒有別家店），就問店長想找哪幾家店。
- 店長說某個平日「視為假日／當假日／照假日排」（或反過來改回平日）：先用 set_day_type 改那天的需求（檢查會改照假日班），再用 apply_changes 排人。不要叫店長去改設定表。
"""

TOOLS = [
    {"name": "apply_changes", "description": "改班表格子，一次可改多格。回傳改後的檢查結果與班表。",
     "input_schema": {"type": "object", "additionalProperties": False, "required": ["changes"], "properties": {
         "changes": {"type": "array", "items": {"type": "object", "additionalProperties": False,
                     "required": ["emp", "day", "value"], "properties": {
                         "emp": {"type": "string", "description": "104 工號"},
                         "day": {"type": "integer", "description": "日期 1～31"},
                         "value": {"type": "string", "description": "格子新內容（照格子寫法）"}}}}}},
     "strict": True},
    {"name": "set_day_type", "description": "把某幾天設成「當假日」（需求、人效照假日算）或改回照星期幾。存在本月設定的特殊日。回傳改後的檢查結果。",
     "input_schema": {"type": "object", "additionalProperties": False, "required": ["days", "holiday"], "properties": {
         "days": {"type": "array", "items": {"type": "integer"}, "description": "日期 1～31"},
         "holiday": {"type": "boolean", "description": "true＝當假日；false＝改回照星期幾"}}},
     "strict": True},
    {"name": "undo", "description": "復原上一次改動（不論是你改的還是店長在日曆上改的）。",
     "input_schema": {"type": "object", "additionalProperties": False, "properties": {}}, "strict": True},
    {"name": "support_eval", "description": "【只評估、不改任何班表】把本店和別家店的人放在一起重新排一次，看別家店能不能派人來補本店缺的班、對方店會不會因此缺人。約 1～3 分鐘。回傳評估結果（不會存檔）。",
     "input_schema": {"type": "object", "additionalProperties": False, "required": ["stores"], "properties": {
         "stores": {"type": "array", "items": {"type": "integer"}, "description": "要評估的支援店店號（1～14）；空陣列＝本店同區的其他店"}}},
     "strict": True},
]


def _apply(st, changes, why):
    shifts = _ext(st).get("shifts") or {}
    dtypes = set(st["template"]["info"]["dayTypes"])
    ids = {e["id"] for e in st["template"]["info"]["emps"]}
    days = {d["d"] for d in st["template"]["info"]["days"]}
    for c in changes:
        if c.get("emp") not in ids:
            raise ValueError(f"沒有工號 {c.get('emp')}")
        if int(c.get("day", 0)) not in days:
            raise ValueError(f"這個月沒有 {c.get('day')} 號")
        parse_cell(c.get("value", ""), shifts, dtypes)
    st["history"] = (st.get("history") or [])[-19:] + [{"schedule": copy.deepcopy(st["schedule"]), "why": why}]
    for c in changes:   # 全形「１０：００－１９：００」存成半形（畫面與之後的比對都用半形）
        st["schedule"].setdefault(c["emp"], {})[str(int(c["day"]))] = unicodedata.normalize("NFKC", c.get("value", "") or "").strip()
    st["check"] = run_check(st)
    return st["check"]


def _set_day_type(st, days, holiday):
    """特殊日「當假日」開關（對話用；畫面上是本月設定特殊日的勾選）"""
    valid = {d["d"] for d in st["template"]["info"]["days"]}
    ms = st.setdefault("mset", {})
    sp = ms.setdefault("special", [])
    for d in days:
        d = int(d)
        if d not in valid:
            raise ValueError(f"這個月沒有 {d} 號")
        x = next((x for x in sp if int(x["day"]) == d), None)
        if x is None:
            if not holiday:
                continue
            x = {"day": d, "revenue": None, "staff": "", "note": ""}
            sp.append(x)
        x["as_holiday"] = bool(holiday)
    ms["special"] = sorted(sp, key=lambda x: x["day"])
    st["check"] = run_check(st)
    return st["check"]


def _undo(st):
    if not st.get("history"):
        return None
    st["schedule"] = st["history"].pop()["schedule"]
    st["check"] = run_check(st)
    return st["check"]


def _support_eval(store, ym, helpers, jk):
    """2026-11-11 經營者：店長想知道別家店能不能來補班 → 把本店＋支援店的人放一起重排一次，只評估、不存檔、不改任何店。
    回傳給 AI 看的文字（現在各店缺幾人、重排後缺幾人、誰哪天去哪家店、工時、104 崗位缺口）"""
    store = int(store)
    helpers = [int(x) for x in helpers or [] if int(x) != store]
    if not helpers:
        rg = _region_of(store)
        helpers = [x for x in REGIONS.get(rg, []) if x != store] if rg else []
        if not helpers:
            return "本店不在任何一區，或同區沒有別家店。請問店長想找哪幾家店來支援（給店號）？"
    bad = [x for x in helpers if x not in STORES and x not in TEST_STORES]
    if bad:
        return f"沒有這些店號：{'、'.join(map(str, bad))}"
    stores = [store] + helpers
    sts, cfgs, not_ready = {}, {}, []
    for x in stores:
        st_ = _load(x, ym)
        cf = cfg_load(x)
        why = "" if st_.get("template") else "還沒有這個月的 104 範本"
        if not why and not form_ready(st_, cf):
            why = "店別基本設定（班別、每天幾人）或本月設定還沒填好"
        if why:
            not_ready.append(f"{_sname(x)}（{x} 號）：{why}")
        else:
            sts[x], cfgs[x] = st_, cf
    if store not in sts:
        return "本店自己的設定還沒填好，沒辦法評估：" + next(x for x in not_ready if x.startswith(f"{_sname(store)}（"))
    stores = [x for x in stores if x in sts]
    if len(stores) < 2:
        return "要評估的店都還沒準備好，沒辦法評估：" + "；".join(not_ready)
    mss = {x: mset_view(sts[x]) for x in stores}
    ctx = (stores, sts, cfgs, mss)
    rg_name = "測試" if all(x in TEST_STORES for x in stores) else ""
    rules, badn, _ = _region_rules(rg_name, ym, ctx, jk)
    y, mth = int(ym[:4]), int(ym[4:])
    pos = {x: (TEST_POS.get(x, "") if x in TEST_STORES else STORE_POS.get(x, "")) for x in stores}
    units, block = [], []
    for x in stores:
        info = sts[x]["template"]["info"]
        for o in stores:   # 104 崗位缺口：評估時先當作可以調，另外列出要先在 104 加崗位
            if o != x and pos[o] and info.get("positions") and pos[o] not in info["positions"]:
                block.append(f"{_sname(x)}的 104 範本沒有「{pos[o]}」崗位（{_sname(x)}的人要去{_sname(o)}支援，104 要先加這個崗位）")
        units.append({"store": x, "pos": pos[x], "info": dict(info, positions=[]), "cfg": cfgs[x], "ms": mss[x], "rules": rules.get(x, []),
                      "hol_days": [d["d"] for d in info["days"] if d["wd"] not in ("六", "日") and _tw_off(dt.date(y, mth, d["d"]))],
                      "prev_tail": _tail_merged(x, ym, {e["id"] for e in info["emps"]}), "fixed_minutes": {}})
    _job(jk, status="running", step=f"正在評估 {len(stores)} 家店一起排（只評估、不改班表，約 1～3 分鐘）")
    r = schedule_solver.solve_multi(units, mode="pt", lend=True, time_limit=90)

    def cur_gap(x):
        chk = sts[x].get("check") or {}
        if not sts[x].get("schedule"):
            return "還沒排班"
        g = [c for c in chk.get("cover") or [] if c["got"] < c["need"]]
        return f"缺 {sum(c['need'] - c['got'] for c in g)} 人次" + (f"（{'、'.join(str(c.get('label') or c.get('day')) for c in g)}）" if g else "")
    L = ["【只評估，沒有改任何一家店的班表】", "評估的店：" + "、".join(f"{_sname(x)}（{x} 號）" for x in stores)]
    if not_ready:
        L.append("沒辦法一起評估（設定沒填好）：" + "；".join(not_ready))
    L.append("\n## 現在各店（照目前存的班表）")
    for x in stores:
        L.append(f"- {_sname(x)}：{cur_gap(x)}")
    L.append("\n## 把這幾家店的人一起重排後（別店人力可互相支援）")
    for u in r["stores"]:
        x = u["store"]
        g = {}
        for c in u["gaps"]:
            g.setdefault(c["day"], []).append(f"{c['shift']} 缺 {c['need'] - c['got']}")
        L.append(f"- {_sname(x)}：" + ("不缺人" if not g else "還缺 " + str(sum(c["need"] - c["got"] for c in u["gaps"])) + " 人次（" + "；".join(f"{d} 號 {'、'.join(v)}" for d, v in sorted(g.items())) + "）"))
        hrs = [f"{h['name']} {h['hours']:g}h" + (f"／目標 {h['target']:g}h" if h.get("target") else (f"／最低 {h['min']:g}h" if h.get("min") else ""))
               for h in u["hours"].values()]
        L.append("  每人工時：" + "、".join(hrs))
        if u.get("unmet"):
            nm = {e["id"]: e["name"] for e in sts[x]["template"]["info"]["emps"]}
            L.append("  備註規則沒做到：" + "；".join(f"{_rule_text(q['rule'], nm)}（{'、'.join(map(str, q['days']))} 號）" for q in u["unmet"]))
    mv = r["moves"]
    L.append(f"\n## 建議的支援（{len(mv)} 人次）")
    L += [f"- {m['day']} 號 {_sname(m['store'])}的 {m['name']} 去{_sname(m['to'])}上{m['shift']}" for m in sorted(mv, key=lambda m: (m["day"], m["store"]))] or ["- 不用調人"]
    if block:
        L.append("\n## 104 要先處理")
        L += [f"- {b}" for b in dict.fromkeys(block)]
    for x, v in badn.items():
        if v:
            L.append(f"（{_sname(x)}有備註 AI 讀不懂，這次沒套：{'；'.join(v)}）")
    L.append(f"\n（程式排了 {r['stats']['seconds']} 秒，結果：{r['stats']['status']}。這是重排的建議，跟各店現在的班表可能不一樣；真的要調，請魔導師在「區域排班」採用，或兩家店長協調後各改各的班表。）")
    return "\n".join(L)


def _chat_system(st):
    cfg = cfg_load(st["store"])
    return (RULES + CHAT_RULES + "\n## 本店這個月的設定（AI 從設定表讀到的）\n" +
            json.dumps(_ext(st, cfg), ensure_ascii=False) +
            "\n\n## 本月設定（店長填的表單原文）\n" + form_text(st, cfg) +
            "\n\n## 104 範本\n- 班表：" + st["template"]["info"]["title"] +
            "\n- 日期：" + "、".join(d["label"] for d in st["template"]["info"]["days"]) +
            "\n- 可用標籤：" + "、".join(st["template"]["info"]["tags"]) +
            "\n- 可用崗位：" + "、".join(st["template"]["info"]["positions"]))


def marks_text(st):
    """每句話都附上設定表的指休／特休／可上日（2026-11-05 模擬：AI 把格子上的 [指休] 標籤當成設定表指休，之前答錯後還會照舊說法）"""
    ms = mset_view(st)
    out = []
    for e in st["template"]["info"]["emps"]:
        mk = ms["marks"].get(e["id"]) or {}
        parts = [f"{k} " + "、".join(str(d) for d in sorted(int(d) for d, v in mk.items() if v == k))
                 for k in MARKS if any(v == k for v in mk.values())]
        out.append(f"{e['name']}：" + ("；".join(parts) if parts else "沒有"))
    hol = sorted(_hol_days(ms))
    if hol:
        out.append("當假日的日子：" + "、".join(map(str, hol)))
    out.append("（班表格子開頭的 [指休] 等是 104 標籤，不是設定表；之前的回覆若說法不同，以這裡為準）")
    return "\n".join(out)


def _chat_turn(st, text, model, k, tag=" "):
    """跑一句對話（可能多次呼叫工具）；回傳 (回覆, 改了幾格)"""
    st["api_messages"].append({"role": "user", "content": f"（本月設定的指休／特休／可上，以這裡為準）\n{marks_text(st)}\n\n（目前班表）\n{table_text(st)}\n\n（目前檢查）\n{check_text(st['check'])}\n\n店長：{text}"})
    changed = 0
    for i in range(8):
        _job(k, status="running", step=f"AI{tag}正在想（第 {i + 1} 輪）…" if i else f"AI{tag}正在想…")
        msg = _client_().messages.create(model=model, max_tokens=16000, system=_chat_system(st), tools=TOOLS,
                                         output_config={"effort": "medium"}, cache_control={"type": "ephemeral"},
                                         messages=st["api_messages"])
        _log_usage(st, "chat", _cost(model, msg.usage))
        st["api_messages"].append({"role": "assistant", "content": _blocks(msg.content)})
        if msg.stop_reason != "tool_use":
            break
        results = []
        for b in msg.content:
            if b.type != "tool_use":
                continue
            try:
                if b.name == "undo":
                    chk = _undo(st)
                    out = "沒有可以復原的改動" if chk is None else "已復原。\n" + check_text(chk)
                elif b.name == "support_eval":   # 只評估、不改班表、不存檔
                    out = _support_eval(st["store"], st["ym"], b.input.get("stores") or [], k)
                elif b.name == "set_day_type":
                    chk = _set_day_type(st, b.input.get("days") or [], b.input.get("holiday"))
                    out = "已改日子類型。\n" + check_text(chk)
                else:
                    chk = _apply(st, b.input.get("changes") or [], "對話：" + text[:40])
                    changed += len(b.input.get("changes") or [])
                    out = "已改。\n" + check_text(chk) + "\n\n改後班表：\n" + table_text(st)
            except ValueError as ex:
                out = f"這次沒有改任何格子：{ex}"
            results.append({"type": "tool_result", "tool_use_id": b.id, "content": out})
        _job(k, status="running", step=f"已改 {changed} 格，AI{tag}正在看檢查結果…")
        st["api_messages"].append({"role": "user", "content": results})
    reply = "".join(b.text for b in msg.content if b.type == "text").strip()
    return reply or "（我剛剛沒有產生回覆，請再說一次。）", changed


def _chat_worker(store, ym, text):
    k = _key(store, ym)
    lk = _lock(k)
    try:
        # AI 想的時候不鎖住檔案（畫面要能照常讀狀態）；工作中其他寫入動作都會被 409 擋下，所以最後整份存回是安全的
        with lk:
            st = _load(store, ym)
        if len(st.get("api_messages") or []) > MAX_API_MESSAGES:
            st["api_messages"] = []   # 對話太長就開新的一段（班表、檢查每句都會重新附上）
        before = copy.deepcopy(st)
        n_err = len(st["check"]["errors"])
        reply, changed = _chat_turn(st, text, CHAT_MODEL, k)
        if len(st["check"]["errors"]) > n_err:
            # 改完反而多了 ❌ → 退回這句之前，再用 Opus 重做一次
            usage = st["usage"]
            st = before
            st["usage"] = usage
            reply, changed = _chat_turn(st, text, MODEL_BACKUP, k, tag="（改完問題變多，重想一次）")
        st["chat"] = (st.get("chat") or []) + [{"role": "me", "text": text, "ts": time.time()},
                                               {"role": "ai", "text": reply, "changed": changed, "ts": time.time()}]
        with lk:
            _save(st)
        _job(k, status="done", step="")
    except Exception as ex:   # noqa: BLE001
        _job(k, status="error", step=f"對話失敗：{type(ex).__name__}: {str(ex)[:300]}")


# ───────────────────────── 每天比對 104（2026-10-26 經營者：店長可直接在 104 改，儀表板每天自動跟上） ─────────────────────────
# 資料由本機 Claude 每天用內建瀏覽器（經營者登入的 104）讀出（tools/sync104/extract.js），再由 tools/sync104/push.py 送到這裡。
# 規則：104 班表有排班內容、而且 104 存檔時間比上次採用的新 → 工作台改成 104 的版本（舊的放進復原紀錄）；
#       104 沒有內容（還沒排）→ 不動工作台，只記錄；工作台改了還沒匯入 104 → 不動，畫面顯示差幾格。
STATUS_104 = {1: "還沒排", 3: "已發佈", 5: "排班中"}


def _hm(ms, base_day):
    """104 的毫秒時間 → 'HH:MM'（台灣時間；跨到隔天加「次日」）"""
    t = dt.datetime.utcfromtimestamp(ms / 1000) + dt.timedelta(hours=8)
    s = t.strftime("%H:%M")
    return ("次日" + s) if t.date() > base_day else s


def _cells_from_104(item, info, shifts_named):
    """104 一份班表 → {工號: {日: 格子}}；回傳 (班表, 對不上的員工名, 看不懂的種類數)"""
    by_name = {}
    for e in info["emps"]:
        by_name.setdefault(e["name"], e["id"])
    emp_no = {}
    missing = []
    for e in item.get("emps") or []:
        if e["name"] in by_name:
            emp_no[str(e["empId"])] = by_name[e["name"]]
        else:
            missing.append(e["name"])
    tags, pos, cls = item.get("tags") or {}, item.get("pos") or {}, item.get("cls") or {}
    month_days = {d["d"] for d in info["days"]}
    m0 = int(item["ym"][4:])
    raw = {}
    unknown = 0
    for empId, base, typ, frm, to, tag_ids, pos_id, cls_id in item.get("shifts") or []:
        no = emp_no.get(str(empId))
        if not no:
            continue
        bd = (dt.datetime.utcfromtimestamp(base / 1000) + dt.timedelta(hours=8)).date()
        if bd.month != m0 or bd.day not in month_days:
            continue
        c = raw.setdefault(no, {}).setdefault(bd.day, {"type": "", "segs": [], "dtag": []})
        tag_names = [tags.get(str(t), "") for t in tag_ids if tags.get(str(t))]
        if typ in (0, 8, 10) and cls.get(str(cls_id)):   # 休假類（例假日／休息日／國定假日…），種類看 classId
            c["type"] = cls[str(cls_id)]
        elif typ == 3 and frm and to:
            c["segs"].append((frm, _hm(frm, bd), _hm(to, bd), "、".join(tag_names), pos.get(str(pos_id), "")))
        elif typ == 2:
            c["dtag"] += tag_names
        else:
            unknown += 1
    out = {}
    named = {tuple(map(tuple, v)): k for k, v in (shifts_named or {}).items()}
    for no, days in raw.items():
        row = {}
        for d, c in days.items():
            segs = sorted(c["segs"])
            plain = tuple((s, e) for _, s, e, t, p in segs if not t and not p)
            if segs and len(plain) == len(segs) and plain in named:
                body = named[plain]
            else:
                body = ",".join(f"{s}-{e}" + (f"@{t}" if t else "") + (f"#{p}" if p else "") for _, s, e, t, p in segs)
            if c["type"] and body:
                body = c["type"] + ":" + body
            elif c["type"]:
                body = c["type"]
            if c["dtag"]:
                body = "[" + "、".join(dict.fromkeys(c["dtag"])) + "]" + body
            row[str(d)] = body
        out[no] = row
    for e in info["emps"]:
        out.setdefault(e["id"], {})
        for d in info["days"]:
            out[e["id"]].setdefault(str(d["d"]), "")
    return out, missing, unknown


def _canon(v, shifts, dtypes):
    try:
        c = parse_cell(v, shifts, dtypes)
    except ValueError:
        return "壞:" + (v or "")
    return json.dumps([c["type"], c["dtag"], c["segs"]], ensure_ascii=False)


def _diff(st, a, b):
    """兩份班表不同的格子：[(工號, 日, a 格子, b 格子)]"""
    shifts = _ext(st).get("shifts") or {}
    dtypes = set(st["template"]["info"]["dayTypes"])
    out = []
    for e in st["template"]["info"]["emps"]:
        for d in st["template"]["info"]["days"]:
            k = str(d["d"])
            va, vb = ((a or {}).get(e["id"]) or {}).get(k, ""), ((b or {}).get(e["id"]) or {}).get(k, "")
            if _canon(va, shifts, dtypes) != _canon(vb, shifts, dtypes):
                out.append((e["id"], d["d"], va, vb))
    return out


def sync104(payload, force_dash=False):
    """收 tools/sync104/push.py 送來的 104 班表；回傳 (每份的處理結果, 排班分析儀表板佇列摘要)"""
    results = []
    dash_items, dash_skipped = [], 0
    for item in (payload or {}).get("items") or []:
        try:
            store, ym = int(item["store"]), str(item["ym"])
        except (KeyError, TypeError, ValueError):
            continue
        if store not in STORES or not re.fullmatch(r"20\d{2}(0[1-9]|1[0-2])", ym):
            continue
        if item.get("error"):
            results.append({"store": store, "ym": ym, "error": item["error"]})
            continue
        k = _key(store, ym)
        with _glock:
            if (_jobs.get(k) or {}).get("status") == "running":
                results.append({"store": store, "ym": ym, "skipped": "AI 正在處理，下次再同步"})
                continue
        with _lock(k):
            st = _load(store, ym)
            res = {"store": store, "ym": ym, "name": item.get("name"), "status": STATUS_104.get(item.get("status"), str(item.get("status")))}
            if item.get("template_b64"):
                raw = base64.b64decode(item["template_b64"])
                try:
                    info = template_info(raw)
                    st["template"] = {"name": f"{item.get('name', '')}（104 自動取得）.xlsx", "b64": item["template_b64"],
                                      "info": info, "source": "104", "fetched": time.time()}
                    res["template"] = "已更新"
                except Exception as ex:   # noqa: BLE001
                    res["template"] = f"範本讀不了：{ex}"
            if not st.get("template"):
                res["error"] = "沒有範本，無法對照"
                results.append(res)
                continue
            cells, missing, unknown = _cells_from_104(item, st["template"]["info"], _ext(st).get("shifts"))
            has_content = any(v for row in cells.values() for v in row.values())
            emps = {e["empId"]: e for e in item.get("emps") or []}
            st["s104"] = {"id": item.get("id"), "name": item.get("name"), "status": res["status"], "modify": item.get("modify"),
                          "fetched": time.time(), "cells": cells, "missing": missing, "unknown": unknown,
                          "emps": [{"name": e["name"], "mustHour": e.get("mustHour"), "annualHour": e.get("annualHour"),
                                    "compHour": e.get("compHour")} for e in emps.values()]}
            adopted = False
            if has_content and (not st.get("schedule") or (item.get("modify") or 0) > (st.get("s104_adopted") or 0)):
                d = _diff(st, st.get("schedule"), cells) if st.get("schedule") else [1]
                if d:
                    if st.get("schedule"):
                        st["history"] = (st.get("history") or [])[-19:] + [{"schedule": st["schedule"], "why": "104 同步前"}]
                    st["schedule"] = cells
                    adopted = True
                    st["chat"] = (st.get("chat") or []) + [{"role": "ai", "ts": time.time(), "text":
                        f"📥 已照 104 的班表（{res['status']}，104 最後存檔 {dt.datetime.utcfromtimestamp((item.get('modify') or 0) / 1000) + dt.timedelta(hours=8):%m/%d %H:%M}）更新工作台"
                        + (f"，跟原本差 {len(d)} 格。" if st.get("history") and isinstance(d[0], tuple) else "。")}]
                st["s104_adopted"] = item.get("modify") or 0
            if st.get("schedule"):
                st["check"] = run_check(st)
                st["s104"]["diff"] = len(_diff(st, st["schedule"], cells)) if has_content else None
            res.update(adopted=adopted, has_content=has_content, missing=missing, unknown=unknown,
                       diff=st["s104"].get("diff"))
            # 2026-11-09 第一段：104 有內容 → 整月轉成排班分析儀表板的列，內容沒變就不重送
            if has_content:
                pk, ph = _dash104_item(item)
                if force_dash or ph != (st.get("dash104") or {}).get("hash"):
                    dash_items.append((k, pk, ph))
                else:
                    dash_skipped += 1
            _save(st)
            results.append(res)
    if dash_items:
        threading.Thread(target=_dash104_worker, args=(dash_items,), daemon=True).start()
    return results, {"queued": len(dash_items), "skipped": dash_skipped}


# ───────────────────────── 104 → 排班分析儀表板（2026-11-09 經營者裁定第一段：儀表板改跟 104 同步） ─────────────────────────
# 工作台同步到 104 班表後，把整月轉成「排班分析資料倉」fact_schedule 的列，POST 給同一支 GAS（REV_API）的 upsert104，
# 再呼叫 processUploads 重算營收人效。口徑（2026-11-09 經營者同意 Claude 建議）：
#   ①特休／補休（時段或當日標籤）：人不在店 → 那段不算工時；整天都不算就不寫 fact_schedule 列，改寫 fact_note「特休 1000-1800」讓日曆看得到。
#   ②崗位＝別家店（支援）：那段算在被支援店（店號照 POS_STORE），備註「跨店支援（自○○店）」；跟 NUEIP 店班表、日獎金一致。
#   ③經管會／交流會議／區會／台北內訓課：工時照寫、備註「區會 1000-1200」→ 儀表板 bonusNonWork 自己扣；事務班、備品班照算現場工時。
#   ④崗位＝出貨中心（不是門市）：留在自己店、備註「出貨中心 1000-1800 不計現場人力」→ 儀表板整天不算。
#   ⑤職稱 104 沒有 → GAS 用 fact_schedule 既有資料同姓名最後一次的職稱；對不到的回報 unknownRole。
#   ⑥例假日／休息日／國定假日／空班、只有指休等當日標籤而沒時段的日子：不寫列（NUEIP 月休也是跳過）。
# token：push.py 從 secrets.env 帶 dash_token 來，存在 _dash104_cfg.json；GAS 端放指令碼屬性 DASH104_TOKEN。
DASH104_LEAVE = {"特休", "補休"}
DASH104_STATUS = "_dash104_status.json"
DASH104_CFG = "_dash104_cfg.json"


def _dash104_cfg(update=None):
    p = _dir() / DASH104_CFG
    cfg = {}
    if p.exists():
        try:
            cfg = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            cfg = {}
    if update:
        cfg.update(update)
        atomic_write_json(p, cfg)
    return cfg


def _dash104_status(update=None):
    p = _dir() / DASH104_STATUS
    st = {}
    if p.exists():
        try:
            st = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            st = {}
    if update is not None:
        st.update(update)
        atomic_write_json(p, st)
    return st


def _tw_dt(ms):
    return dt.datetime.fromtimestamp(ms / 1000, dt.timezone.utc) + dt.timedelta(hours=8)


def _hhmm4(ms, base_day):
    """104 毫秒 → 'HHMM'（fact_schedule 的班段寫法）；跨到隔天一律寫 2400"""
    t = _tw_dt(ms)
    if t.date() > base_day:
        return "2400"
    return t.strftime("%H%M")


def _dash104_item(item):
    """一份 104 班表 → upsert104 的 item（rows／notes）與內容雜湊"""
    tags, pos = item.get("tags") or {}, item.get("pos") or {}
    emps = {str(e["empId"]): e for e in item.get("emps") or []}
    home, ym = int(item["store"]), str(item["ym"])
    days = {}
    for empId, base, typ, frm, to, tag_ids, pos_id, _cls in item.get("shifts") or []:
        if str(empId) not in emps:
            continue
        bd = _tw_dt(base).date()
        if f"{bd:%Y%m}" != ym:
            continue
        c = days.setdefault((bd, str(empId)), {"segs": [], "dtag": []})
        names = [tags[str(t)] for t in (tag_ids or []) if tags.get(str(t))]
        if typ == 3 and frm and to:
            c["segs"].append((frm, to, names, pos.get(str(pos_id), "") if pos_id else ""))
        elif typ == 2:
            c["dtag"] += names
    rows, notes = [], []
    for (bd, empId), c in sorted(days.items()):
        name, date = emps[empId]["name"], bd.isoformat()
        day_leave = [t for t in c["dtag"] if t in DASH104_LEAVE]
        other_dtag = [t for t in dict.fromkeys(c["dtag"]) if t not in DASH104_LEAVE]
        by_store = {}
        leave_notes = []
        for frm, to, names, pname in sorted(c["segs"]):
            label = f"{_hhmm4(frm, bd)}-{_hhmm4(to, bd)}"
            target = POS_STORE.get(pname) or home
            cell = label + (("「" + "、".join(names) + "」") if names else "") + (f"＠{pname}" if pname and target != home else "")
            if pname and target == home and pname not in POS_STORE:
                cell += f"＠{pname}"
            leave = day_leave or [n for n in names if n in DASH104_LEAVE]
            if leave:
                leave_notes.append("、".join(dict.fromkeys(leave)) + " " + label)
                continue
            t = by_store.setdefault(target, {"segs": [], "cell": [], "note": []})
            t["segs"].append(label)
            t["cell"].append(cell)
            for n in names:
                if n != "支援":
                    t["note"].append(f"{n} {label}")
            if pname and target == home and pname not in POS_STORE:
                t["note"].append(f"{pname} {label} 不計現場人力")
            if target != home and f"跨店支援（自{STORES.get(home, home)}店）" not in t["note"]:
                t["note"].append(f"跨店支援（自{STORES.get(home, home)}店）")
        for target, t in sorted(by_store.items()):
            note = "；".join(dict.fromkeys(other_dtag + t["note"] + leave_notes))
            rows.append([date, target, empId, name, "", "|".join(t["segs"]), "\n".join(t["cell"]), note])
        if leave_notes and not by_store:
            notes.append([date, home, empId, name, "；".join(leave_notes), 0])
    payload = {"store": home, "ym": ym, "source": f"104:{home}-{ym}:{item.get('id')}@{item.get('modify') or 0}",
               "rows": rows, "notes": notes}
    h = hashlib.sha1(json.dumps({"rows": rows, "notes": notes}, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()
    return payload, h


def _gas_post(url, body, timeout=300):
    req = urllib.request.Request(url, data=json.dumps(body, ensure_ascii=False).encode("utf-8"), method="POST",
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout, context=_ssl_ctx()) as r:
        return json.loads(r.read().decode("utf-8"))


def _gas_get(url, params, timeout=300):
    with urllib.request.urlopen(url + "?" + urllib.parse.urlencode(params), timeout=timeout, context=_ssl_ctx()) as r:
        t = r.read().decode("utf-8")
    m = re.match(r"^\s*\w+\((.*)\)\s*$", t, re.S)
    return json.loads(m.group(1) if m else t)


def _dash104_worker(dash_items, backup=False):
    """背景：把這批 104 班表送到排班分析儀表板，再重算；結果寫 _dash104_status.json 與各月 state 的 dash104"""
    started = time.time()
    keys = [k for k, _, _ in dash_items]
    _dash104_status({"running": True, "started": started, "finished": None, "items": keys, "results": None, "process": None, "error": None})
    token = _dash104_cfg().get("token")
    out = {"running": False, "finished": None}
    try:
        if not token:
            raise RuntimeError("工作台還沒有排班分析儀表板的 dash_token（push.py 要從 secrets.env 帶 SCHED_DASH_TOKEN 來）")
        res = _gas_post(REV_API, {"action": "upsert104", "token": token, "backup": bool(backup),
                                  "items": [pk for _, pk, _ in dash_items]})
        if not res.get("ok"):
            raise RuntimeError(res.get("error") or "GAS 回傳不 ok")
        out["results"] = res
        for (k, _pk, ph), r in zip(dash_items, res.get("results") or []):
            store, ym = k.split("-")
            with _lock(k):
                st = _load(int(store), ym)
                st["dash104"] = {"hash": ph, "at": time.time(), "written": r.get("written"), "removedOld": r.get("removedOld"),
                                 "unknownRole": r.get("unknownRole") or []}
                _save(st)
        proc = _gas_get(REV_API, {"action": "processUploads", "callback": "cb"}, timeout=420)
        out["process"] = proc
    except Exception as ex:   # noqa: BLE001
        out["error"] = str(ex)
    out["finished"] = time.time()
    _dash104_status(out)


# ───────────────────────── 雲端自動同步 104（2026-11-09 經營者：「104 不開 API，你處理自動化的流程，不要麻煩我」） ─────────────────────────
# 本伺服器每天 06:05（台灣）自己用無頭瀏覽器（tools/sync104/fetch104.py，Playwright）登入 104 讀 12 店班表 → sync104()（工作台＋排班分析儀表板）。
# 104 帳號：程式專用的排班唯讀帳號，由 Mac 的排程從 secrets.env（HR104_EMAIL／HR104_PWD）用 cred104_set 送來、存硬碟 _104_cred.json；不放環境變數、不回顯。
# 只「讀」104。帳密錯一次就停（code 2），不重試，避免被 104 鎖。狀態 _104_cloud_status.json；Mac 10:13 排程會看它決定要不要補跑。
TW_TZ = dt.timezone(dt.timedelta(hours=8))
CRED104 = "_104_cred.json"
CLOUD104_STATUS = "_104_cloud_status.json"
CLOUD104_AT = (6, 5)   # 每天台灣 06:05


def _json_file(name, update=None):
    p = _dir() / name
    cur = {}
    if p.exists():
        try:
            cur = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            cur = {}
    if update is not None:
        cur.update(update)
        atomic_write_json(p, cur)
    return cur


def _cred104(update=None):
    return _json_file(CRED104, update)


def _cloud104_status(update=None):
    return _json_file(CLOUD104_STATUS, update)


def _fetch104_module():
    import importlib.util
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tools", "sync104", "fetch104.py")
    spec = importlib.util.spec_from_file_location("fetch104", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


_cloud104_lock = threading.Lock()


def cloud104_run(trigger="schedule", selftest=False):
    """登入 104 → 讀班表 → sync104；selftest＝只確認無頭瀏覽器開得到 104 登入頁。結果寫 _104_cloud_status.json"""
    if not _cloud104_lock.acquire(blocking=False):
        return False
    st = {"running": True, "started": time.time(), "trigger": trigger, "selftest": selftest, "finished": None,
          "code": None, "msg": None, "results": None, "dash": None}
    _cloud104_status(st)
    try:
        mod = _fetch104_module()
        if selftest:
            ok, msg = asyncio.run(mod.probe())
            st.update(code=0 if ok else 6, msg=msg)
        else:
            cred = _cred104()
            if not cred.get("email") or not cred.get("pwd"):
                st.update(code=7, msg="還沒設定 104 排班帳號（Mac 排程會從 secrets.env 的 HR104_EMAIL／HR104_PWD 送來）")
            else:
                os.environ["HR104_EMAIL"], os.environ["HR104_PWD"] = cred["email"], cred["pwd"]
                code, msg, payload = asyncio.run(mod.fetch())
                st.update(code=code, msg=msg)
                if code == 0:
                    results, dash = sync104(payload)
                    st["results"] = [{k: r.get(k) for k in ("store", "ym", "status", "adopted", "diff", "error", "template")} for r in results]
                    st["dash"] = dash
    except Exception as ex:  # noqa: BLE001
        st.update(code=6, msg=f"{type(ex).__name__}: {str(ex)[:500]}")
    finally:
        st.update(running=False, finished=time.time())
        _cloud104_status(st)
        _cloud104_lock.release()
    return True


def _cloud104_loop():
    while True:
        now = dt.datetime.now(TW_TZ)
        nxt = now.replace(hour=CLOUD104_AT[0], minute=CLOUD104_AT[1], second=0, microsecond=0)
        if nxt <= now:
            nxt += dt.timedelta(days=1)
        time.sleep(max(1, (nxt - now).total_seconds()))
        if _cred104().get("email"):
            cloud104_run("schedule")


# 2026-11-09 深夜：實測 104 登入 API（bsignin /api/is-valid-account）被 Cloudflare 機器人防護擋下回 403（本機與 Render 都一樣），
# 偽裝瀏覽器身分繞過 bot 偵測不在允許範圍 → 每天 06:05 的雲端自動登入預設關閉（SCHED_CLOUD104=1 才開），等 104 開放 API 再啟用。
if os.environ.get("SCHED_CLOUD104", "0") == "1":
    threading.Thread(target=_cloud104_loop, daemon=True, name="cloud104").start()


# ───────────────────────── 營業額與日獎金（2026-10-28 經營者：班表要有營業額及日獎金、產各店日獎金表給會計） ─────────────────────────

def _ssl_ctx():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        return ssl.create_default_context()


# ── 每人每月早班／晚班次數（2026-11-13 經營者：排班不考慮早晚班公平，但班表要列出每人每月早、晚班次數，一年份）──
# 早班＝開店那一批、晚班＝做到打烊那一批；中班、全天班、特休／補休不算。
# 開店／打烊時間＝那個月「平日」「假日」各自最常見的當天最早上班、最晚下班時間（單看當天會被偶爾的事務班、開會拉歪）。
# 資料：這個月用工作台目前的班表；其他月份先看工作台存檔（104 每天同步），沒有再讀排班分析資料倉（fact_schedule，2026-03 起才有）。
_EL_OFF_RE = re.compile(r"特休|補休")
_el_lock = threading.Lock()


def _el_classify(days):
    """days＝[{"wk": 假日?, "cells": [(誰, 上班分鐘, 下班分鐘)]}] → {誰: [早, 晚]}"""
    import collections
    op, cl = {}, {}
    for wk in (False, True):
        a = collections.Counter(min(c[1] for c in d["cells"]) for d in days if d["wk"] == wk and d["cells"])
        b = collections.Counter(max(c[2] for c in d["cells"]) for d in days if d["wk"] == wk and d["cells"])
        if a:
            top = max(a.values())
            op[wk] = min(t for t, n in a.items() if n == top)
            top = max(b.values())
            cl[wk] = max(t for t, n in b.items() if n == top)
    out = {}
    for d in days:
        if d["wk"] not in op:
            continue
        o, c = op[d["wk"]], cl[d["wk"]]
        for who, a, b in d["cells"]:
            n = out.setdefault(who, [0, 0])
            if a <= o + 45 and b <= c - 60:
                n[0] += 1
            elif b >= c - 45 and a >= o + 60:
                n[1] += 1
    return out


def _el_from_state(st):
    """工作台的一份班表 → {姓名: [早, 晚]}；沒有班表回 None"""
    tpl, sch = st.get("template"), st.get("schedule") or {}
    if not tpl or not sch:
        return None
    info = tpl["info"]
    shifts, dtypes = (_ext(st) or {}).get("shifts") or {}, set(info["dayTypes"])
    y, mo = int(st["ym"][:4]), int(st["ym"][4:])
    days = []
    for d in info["days"]:
        cells = []
        for e in info["emps"]:
            try:
                c = parse_cell((sch.get(e["id"]) or {}).get(str(d["d"]), ""), shifts, dtypes)
            except (ValueError, TypeError):
                continue
            segs = [x for x in c["segs"] if x and x[0] and x[1] and not _EL_OFF_RE.search((x[2] if len(x) > 2 else "") or "")]
            if not segs or _EL_OFF_RE.search(c.get("dtag") or ""):
                continue
            try:
                a, b = schedule_solver._mins(segs[0][0]), schedule_solver._mins(segs[-1][1])
            except (ValueError, TypeError):
                continue
            cells.append((e["name"].strip(), a, b + (1440 if b < a else 0)))
        days.append({"wk": d["wd"] in ("六", "日") or _tw_off(dt.date(y, mo, d["d"])), "cells": cells})
    return _el_classify(days) if any(d["cells"] for d in days) else None


def _el_from_gas(store, ym):
    """排班分析資料倉那個月的班表 → {姓名: [早, 晚]}；那個月沒資料回 {}；讀不到丟 Exception"""
    src = TEST_SRC if int(store) in TEST_STORES else store
    url = REV_API + "?" + urllib.parse.urlencode({"action": "getStoreDailyData", "store_id": int(src), "month": f"{ym[:4]}-{ym[4:]}", "callback": "cb"})
    with urllib.request.urlopen(url, timeout=90, context=_ssl_ctx()) as r:
        t = r.read().decode("utf-8")
    m = re.search(r"^\s*cb\((.*)\)\s*;?\s*$", t, re.S)
    d = json.loads(m.group(1) if m else t)
    if not d.get("ok"):
        raise ValueError(str(d.get("error") or "班表資料讀不到"))
    days = []
    for row in d.get("daily") or []:
        cells = []
        for e in row.get("employees") or []:
            if _EL_OFF_RE.search(str(e.get("cell") or "") + str(e.get("note") or "")):
                continue
            segs = re.findall(r"(\d{2})(\d{2})-(\d{2})(\d{2})", str(e.get("shift") or ""))
            if not segs or not str(e.get("name") or "").strip():
                continue
            a, b = int(segs[0][0]) * 60 + int(segs[0][1]), int(segs[-1][2]) * 60 + int(segs[-1][3])
            cells.append((str(e["name"]).strip(), a, b + (1440 if b < a else 0)))
        days.append({"wk": bool(row.get("is_weekend") or row.get("is_holiday")), "cells": cells})
    return _el_classify(days)


def el_year(st):
    """班表下方「每人每月早班／晚班次數」：這份班表所在那一年 1～12 月 → {year, months, rows, note}"""
    import concurrent.futures
    store, ym = st["store"], st["ym"]
    year = ym[:4]
    cp = _dir() / "el_cache.json"
    with _el_lock:
        try:
            cache = json.loads(cp.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            cache = {}
    now = time.time()
    today = dt.date.today()
    recent = {f"{today.year}{today.month:02d}", _prev_ym(f"{today.year}{today.month:02d}")}
    data, need, failed = {}, [], []
    for mo in range(1, 13):
        mm = f"{year}{mo:02d}"
        if mm == ym:
            data[mm] = _el_from_state(st)
            continue
        p = _dir() / f"{_key(store, mm)}.json"
        got = None
        if p.exists():
            try:
                got = _el_from_state(json.loads(p.read_text(encoding="utf-8")))
            except Exception:   # noqa: BLE001
                got = None
        if got is not None:
            data[mm] = got
            continue
        if mm < "202603":            # 排班分析資料倉 2026-03 才開始有班表
            data[mm] = None
            continue
        c = cache.get(f"{int(store)}-{mm}")
        ttl = 6 * 3600 if (mm in recent or mm > max(recent)) else 30 * 86400
        if c and now - c.get("t", 0) < ttl:
            data[mm] = c["cnt"] or None
        else:
            need.append(mm)
    if need:
        with concurrent.futures.ThreadPoolExecutor(max_workers=4) as ex:
            futs = {mm: ex.submit(_el_from_gas, store, mm) for mm in need}
        for mm, f in futs.items():
            try:
                cnt = f.result()
                cache[f"{int(store)}-{mm}"] = {"t": now, "cnt": cnt}
                data[mm] = cnt or None
            except Exception:   # noqa: BLE001
                old = cache.get(f"{int(store)}-{mm}")
                data[mm] = (old or {}).get("cnt") or None
                failed.append(int(mm[4:]))
        with _el_lock:
            try:
                cur = json.loads(cp.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                cur = {}
            cur.update({k: v for k, v in cache.items() if k.split("-")[1] in need and k.startswith(f"{int(store)}-")})
            atomic_write_json(cp, cur)
    ext = _ext(st) if st.get("template") else {}
    typ = {e.get("id"): e.get("type") for e in ext.get("employees") or []}
    rows = []
    for e in (st.get("template") or {}).get("info", {}).get("emps") or []:
        nm = e["name"].strip()
        ms = {str(mo): ((data[f"{year}{mo:02d}"] or {}).get(nm) if data[f"{year}{mo:02d}"] is not None else None) for mo in range(1, 13)}
        rows.append({"id": e["id"], "name": e["name"], "type": typ.get(e["id"], ""), "m": ms,
                     "total": [sum((v or [0, 0])[i] for v in ms.values()) for i in (0, 1)]})
    return {"ok": True, "year": int(year), "cur": int(ym[4:]), "has": [mo for mo in range(1, 13) if data[f"{year}{mo:02d}"] is not None],
            "rows": rows, "failed": failed}


def fetch_rev(store, ym):
    """排班分析 GAS getStoreDailyData（bonus=1）→ {日: {net, voucher, qty, rev}}；失敗丟 Exception"""
    store = TEST_SRC if int(store) in TEST_STORES else store
    url = REV_API + "?" + urllib.parse.urlencode({"action": "getStoreDailyData", "store_id": int(store), "month": f"{ym[:4]}-{ym[4:]}", "bonus": "1", "callback": "cb"})
    with urllib.request.urlopen(url, timeout=90, context=_ssl_ctx()) as r:
        t = r.read().decode("utf-8")
    m = re.search(r"^\s*cb\((.*)\)\s*;?\s*$", t, re.S)
    d = json.loads(m.group(1) if m else t)
    if not d.get("ok"):
        raise ValueError(str(d.get("error") or "營收資料讀不到"))
    out = {}
    for row in d.get("daily") or []:
        try:
            day = int(str(row.get("date"))[8:10])
        except ValueError:
            continue
        net = float(row.get("revenue") or 0)
        v = float(row.get("voucher100") or 0)
        if net <= 0 and v <= 0:
            continue
        out[str(day)] = {"net": round(net), "voucher": round(v), "qty": int(row.get("voucher100_qty") or 0),
                         "rev": round(float(row.get("bonus_rev") if row.get("bonus_rev") is not None else net + v))}
    bonus = d.get("bonus") or {}
    return {"days": out, "voucher_ok": bool(bonus.get("ok")), "fetched": time.time(), "error": ""}


def _rev_refresh_bg(store, ym):
    k = _key(store, ym)
    with _glock:
        if k in _rev_jobs:
            return
        _rev_jobs.add(k)

    def run():
        try:
            data = fetch_rev(store, ym)
        except Exception as ex:   # noqa: BLE001
            data = None
            err = f"{type(ex).__name__}: {str(ex)[:120]}"
        with _lock(k):
            st = _load(store, ym)
            if data:
                st["rev"] = data
            else:
                st["rev"] = dict(st.get("rev") or {"days": {}}, error=err, tried=time.time())
            _save(st)
        with _glock:
            _rev_jobs.discard(k)
    threading.Thread(target=run, daemon=True).start()


def _rev_stale(st):
    rv = st.get("rev") or {}
    last = max(rv.get("fetched") or 0, rv.get("tried") or 0)
    return time.time() - last > REV_TTL


def _emp_day_work(c, title=""):
    """一位夥伴一天「真正在本店工作」的時數、不算的原因、跨店支援段（規則同排班儀表板 bonusNonWork，改成讀 104 格子）。
    回 (work, why, away)；away=[{store, hours, span}]：崗位填別店且不是會議／假別的時段 → 日獎金算在被支援店（2026-10-28 經營者裁定）"""
    if not c["segs"]:
        return 0.0, "", []
    if c["type"] and c["type"] not in ("休息日", "國定假日"):   # 2026-11-11 經營者：休息日／國定假日有來上班的要算日獎金（會議等標籤的時段照樣不算）
        return 0.0, f"班表寫「{c['type']}」", []
    dtag = c.get("dtag") or ""
    if BONUS_NOFLOOR_RE.search(dtag):
        return 0.0, f"「{dtag}」", []
    any_tag = any(t for _, _, t, _ in c["segs"])
    if not any_tag and (BONUS_MEET_RE.search(dtag) or BONUS_LEAVE_RE.search(dtag)) and "指休" != dtag:
        return 0.0, f"整天「{dtag}」", []
    work, why, away = 0.0, [], []
    for s_, e_, t, pz in c["segs"]:
        h = (_mins(e_) - _mins(s_)) / 60
        if BONUS_MEET_RE.search(t) or BONUS_NOFLOOR_RE.search(t):
            why.append(f"{s_}-{e_}「{t}」")
        elif BONUS_LEAVE_RE.search(t) and "指休" not in t and "挪假" not in t:
            why.append(f"{s_}-{e_}「{t}」")
        elif _elsewhere(pz, title):
            sto = _pos_store(pz)
            if sto:
                away.append({"store": sto, "hours": round(h, 2), "span": f"{s_}-{e_}"})
                why.append(f"{s_}-{e_} 支援{pz}（獎金算{pz}）")
            else:
                why.append(f"{s_}-{e_} 在{pz}")
        else:
            work += h
    return round(work, 2), "、".join(why), away


_sup_cache = {}


def _support_index(ym):
    """全部店這個月班表裡「崗位填別店」的時段 → {被支援店: {日: {工號: {id, name, from, hours, span, trainee}}}}。
    用各店月檔的修改時間當快取鍵，檔沒變就不重算。"""
    files = []
    for store in sorted(STORES):
        p = _dir() / f"{_key(store, ym)}.json"
        try:
            files.append((store, p.stat().st_mtime_ns))
        except FileNotFoundError:
            continue
    sig = tuple(files)
    with _glock:
        hit = _sup_cache.get(ym)
    if hit and hit[0] == sig:
        return hit[1]
    idx = {}
    for store, _ in files:
        try:
            st = _load(store, ym)
        except (OSError, ValueError):
            continue
        tpl = st.get("template")
        if not tpl or not st.get("schedule"):
            continue
        shifts = (_ext(st) or {}).get("shifts") or {}
        dtypes = set(tpl["info"]["dayTypes"])
        for e in tpl["info"]["emps"]:
            for k, v in (st["schedule"].get(e["id"]) or {}).items():
                try:
                    c = parse_cell(v, shifts, dtypes)
                except ValueError:
                    continue
                _, _, away = _emp_day_work(c, tpl["info"]["title"])
                for a in away:
                    if a["store"] == store:
                        continue
                    g = idx.setdefault(a["store"], {}).setdefault(str(k), {}).setdefault(e["id"], {
                        "id": e["id"], "name": e["name"], "from": store, "hours": 0.0, "span": [], "trainee": "實作" in e["name"]})
                    g["hours"] = round(g["hours"] + a["hours"], 2)
                    g["span"].append(a["span"])
    with _glock:
        _sup_cache[ym] = (sig, idx)
    return idx


def bonus_calc(st, ext=None):
    """每天：獎金營收、店日工時、級距、達標、倍數、每人獎金。回傳 {days:{日:{…}}, month:{…}, rev_ok}"""
    tpl = st.get("template")
    if not tpl or not st.get("schedule"):
        return None
    ext = ext if ext is not None else _ext(st)
    shifts = ext.get("shifts") or {}
    dtypes = set(tpl["info"]["dayTypes"])
    rev = (st.get("rev") or {}).get("days") or {}
    emps = tpl["info"]["emps"]
    own_ids = {e["id"] for e in emps}
    sup = (_support_index(st["ym"]).get(int(st["store"])) or {})
    guests = {}   # 外店來支援的人：工號 → {id, name, from, days, work, bonus}
    days = {}
    by_emp = {e["id"]: 0 for e in emps}
    tot = hit_n = cross_n = biz = 0
    for d in tpl["info"]["days"]:
        k = str(d["d"])
        staff = {}
        hours = 0.0
        for e in emps:
            try:
                c = parse_cell((st["schedule"].get(e["id"]) or {}).get(k, ""), shifts, dtypes)
            except ValueError:
                continue
            w, why, away = _emp_day_work(c, tpl["info"]["title"])
            if w <= 0 and not why:
                continue
            trainee = "實作" in e["name"]
            staff[e["id"]] = {"work": w, "why": why, "trainee": trainee, "bonus": 0, "away": away}
            hours += w
        for g in (sup.get(k) or {}).values():
            if g["id"] in own_ids and (staff.get(g["id"]) or {}).get("work", 0) > 0:
                continue   # 本店班表已經排了這個人，以本店班表為準，不重複算
            staff[g["id"]] = {"work": g["hours"], "why": "", "trainee": g["trainee"], "bonus": 0, "away": [],
                              "guest": True, "from": g["from"], "name": g["name"], "span": "、".join(g["span"])}
            hours += g["hours"]
            gg = guests.setdefault(g["id"], {"id": g["id"], "name": g["name"], "from": g["from"], "days": 0, "work": 0.0, "bonus": 0})
            gg["days"] += 1
            gg["work"] = round(gg["work"] + g["hours"], 2)
            by_emp.setdefault(g["id"], 0)
        hours = round(hours, 2)
        r = rev.get(k) or {}
        rv = r.get("rev") or 0
        tier = next((i for i, (cap, _) in enumerate(BONUS_TIERS) if hours <= cap), -1)
        row = {"hours": hours, "rev": rv, "net": r.get("net") or 0, "voucher": r.get("voucher") or 0, "qty": r.get("qty") or 0,
               "cap": BONUS_TIERS[tier][0] if tier >= 0 else None, "threshold": BONUS_TIERS[tier][1] if tier >= 0 else None,
               "next": BONUS_TIERS[tier + 1][1] if 0 <= tier < len(BONUS_TIERS) - 1 else None,
               "hit": False, "cross": False, "mult": 0, "total": 0, "no_tier": hours > 0 and tier < 0, "staff": staff}
        if hours > 0 and rv > 0:
            biz += 1
            if tier >= 0:
                row["hit"] = rv >= row["threshold"]
                row["cross"] = row["hit"] and row["next"] is not None and rv >= row["next"]
                row["mult"] = 2 if row["cross"] else (1 if row["hit"] else 0)
        if row["mult"]:
            hit_n += 1
            cross_n += 1 if row["cross"] else 0
            for eid, x in staff.items():
                if x["work"] > 0 and not x["trainee"]:
                    x["bonus"] = int(BONUS_BASE * x["work"] / BONUS_STD_H * row["mult"] + 0.5)
                    by_emp[eid] += x["bonus"]
                    row["total"] += x["bonus"]
                    if x.get("guest"):
                        guests[eid]["bonus"] += x["bonus"]
        tot += row["total"]
        days[k] = row
    rvx = st.get("rev") or {}
    return {"days": days, "month": {"total": tot, "hit_days": hit_n, "cross_days": cross_n, "biz_days": biz, "by_emp": by_emp,
                                    "guests": sorted(guests.values(), key=lambda g: (g["from"], g["id"]))},
            "rev_fetched": rvx.get("fetched"), "rev_error": rvx.get("error") or "", "rev_days": len(rev), "voucher_ok": rvx.get("voucher_ok", False)}


def _rev_refresh_all(stores, ym, wait=50):
    """匯出前把過期的營收一次抓齊（並行，最多等 wait 秒）"""
    todo = []
    for store in stores:
        with _lock(_key(store, ym)):
            st = _load(store, ym)
        if st.get("schedule") and _rev_stale(st):
            todo.append(store)
    if not todo:
        return
    res = {}

    def run(store):
        try:
            res[store] = fetch_rev(store, ym)
        except Exception as ex:   # noqa: BLE001
            res[store] = {"error": f"{type(ex).__name__}: {str(ex)[:120]}"}
    ths = [threading.Thread(target=run, args=(x,), daemon=True) for x in todo]
    for t in ths:
        t.start()
    for t in ths:
        t.join(wait)
    for store, data in res.items():
        with _lock(_key(store, ym)):
            st = _load(store, ym)
            st["rev"] = data if "days" in data else dict(st.get("rev") or {"days": {}}, error=data["error"], tried=time.time())
            _save(st)


def bonus_xlsx(stores, ym):
    """各店日獎金表（給會計）：各店各人合計／每日明細／每人每日。回傳 (bytes, 有資料的店數)"""
    from openpyxl.styles import Alignment, Font, PatternFill
    _rev_refresh_all(stores, ym)
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "各店各人合計"
    ws2 = wb.create_sheet("每日明細")
    ws3 = wb.create_sheet("每人每日")
    ws4 = wb.create_sheet("說明")
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="F3E9D8")
    ws1.append(["店號", "店名", "工號", "姓名", "正職/兼職", "上班天數", "計獎工時", "日獎金合計", "備註"])
    ws2.append(["店號", "店名", "日期", "星期", "POS 淨營收", "甜點券$100 加回", "獎金營收", "店日工時", "級距上限(h)", "門檻", "達標", "倍數", "當日獎金合計", "當天上班（工時）"])
    hdr3 = None
    n_store = 0
    for store in stores:
        st = _load(store, ym)
        b = bonus_calc(st)
        if not b:
            continue
        n_store += 1
        info = st["template"]["info"]
        ext = _ext(st)
        etype = {e.get("id"): e.get("type") for e in ext.get("employees") or []}
        name = f"{store} {_sname(store)}"
        if hdr3 is None:
            hdr3 = ["店號", "店名", "工號", "姓名"] + [str(d["d"]) for d in info["days"]] + ["合計"]
            ws3.append(hdr3)
        for e in info["emps"]:
            days_on = sum(1 for k, r in b["days"].items() if (r["staff"].get(e["id"]) or {}).get("work", 0) > 0)
            work = round(sum((r["staff"].get(e["id"]) or {}).get("work", 0) for r in b["days"].values()), 2)
            away = {}
            for k, r in sorted(b["days"].items(), key=lambda kv: int(kv[0])):
                for a in (r["staff"].get(e["id"]) or {}).get("away") or []:
                    kk = (int(k), a["store"])
                    away[kk] = round(away.get(kk, 0) + a["hours"], 2)
            note = ("；".join(f"{int(ym[4:])}/{d} 支援{STORES.get(sto, sto)}店 {h:g}h" for (d, sto), h in away.items()) + "（獎金算在被支援店）") if away else ""
            ws1.append([store, STORES.get(store, ""), e["id"], e["name"], etype.get(e["id"], ""), days_on, work, b["month"]["by_emp"].get(e["id"], 0), note])
            row = [store, STORES.get(store, ""), e["id"], e["name"]]
            for d in info["days"]:
                x = (b["days"].get(str(d["d"])) or {}).get("staff", {}).get(e["id"])
                row.append(x["bonus"] if x and x["bonus"] else None)
            row.append(b["month"]["by_emp"].get(e["id"], 0))
            ws3.append(row)
        for g in b["month"].get("guests") or []:
            frm = f"{g['from']}{STORES.get(g['from'], '')}店"
            ws1.append([store, STORES.get(store, ""), g["id"], g["name"], f"外店支援（{frm}）", g["days"], g["work"], g["bonus"], f"從{frm}來支援，獎金算本店"])
            row = [store, STORES.get(store, ""), g["id"], f"{g['name']}（{frm}支援）"]
            for d in info["days"]:
                x = (b["days"].get(str(d["d"])) or {}).get("staff", {}).get(g["id"])
                row.append(x["bonus"] if x and x["bonus"] else None)
            row.append(g["bonus"])
            ws3.append(row)
        names = {e["id"]: e["name"] for e in info["emps"]}
        for d in info["days"]:
            r = b["days"].get(str(d["d"])) or {}
            if not r.get("hours") and not r.get("rev"):
                continue
            who = "、".join(f"{x.get('name') or names.get(eid, eid)} {x['work']:g}h" + ("（實作）" if x["trainee"] else "")
                            + (f"（{STORES.get(x['from'], x['from'])}店支援）" if x.get("guest") else "")
                            for eid, x in r["staff"].items() if x["work"] > 0)
            ws2.append([store, STORES.get(store, ""), f"{ym[:4]}-{ym[4:]}-{d['d']:02d}", d["wd"], r.get("net"), r.get("voucher"), r.get("rev"), r.get("hours"),
                        r.get("cap"), r.get("threshold"), ("跨級 ×2" if r.get("cross") else "達標" if r.get("hit") else ("超過100h" if r.get("no_tier") else "未達標")) if r.get("rev") else "沒有營收資料",
                        r.get("mult"), r.get("total"), who])
    for ws in (ws1, ws2, ws3):
        for c in ws[1]:
            c.font = bold
            c.fill = fill
            c.alignment = Alignment(horizontal="center")
        ws.freeze_panes = "E2" if ws is ws3 else "A2"
    ws1.column_dimensions["D"].width = 12
    ws1.column_dimensions["E"].width = 18
    ws1.column_dimensions["I"].width = 50
    ws2.column_dimensions["N"].width = 60
    ws2.column_dimensions["C"].width = 12
    ws4["A1"] = "各店日獎金表（依門市排班工作台的班表估算）"
    ws4["A1"].font = bold
    for i, t in enumerate([
        f"月份：{ym[:4]}-{ym[4:]}；產出時間：{dt.datetime.now():%Y-%m-%d %H:%M}",
        "班表來源：門市排班工作台（每天早上自動跟 104 同步；店長在 104 改班，隔天會跟上）。實際發放請以出勤紀錄核對。",
        "營收來源：排班分析資料倉（POS 淨營收，已扣全部折扣）＋「甜點券$100」折抵額加回（「自己人$100甜點券」不算）。",
        "店日工時＝當天在本店真正工作的時數合計：假別（特休、病假…）、交流會議／區會／內訓／經管會的時段不算；事務班、備品班、包館照算。",
        "跨店支援（崗位填別家店的時段）：那幾個小時算在「被支援的店」，用被支援店的營收與級距算獎金，列在被支援店底下（正職/兼職欄寫「外店支援」）；原店那幾個小時不算、備註欄會註明。被支援店的班表若本來就排了這位夥伴，以被支援店的班表為準，不重複算。",
        "級距（店日工時→營收門檻）：≤10h 10,000／≤20h 14,000／≤30h 29,000／≤40h 42,000／≤50h 53,000／≤60h 65,000／≤70h 77,000／≤80h 88,000／≤90h 100,000／≤100h 110,000；超過 100h 不發。",
        "達標＝獎金營收 ≥ 門檻（剛好等於也算）；營收 ≥ 下一級門檻＝跨級，獎金 ×2。",
        "每人日獎金＝200 元 × 個人計獎工時 ÷ 8 × 倍數（四捨五入）；名字含「實作」的實作夥伴工時計入全店、但不發。",
        "「沒有營收資料」：那天 POS 還沒進資料倉（當天與前一天常見），隔天再下載一次。",
    ], start=2):
        ws4.cell(i, 1, t)
    ws4.column_dimensions["A"].width = 120
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), n_store


# ───────────────────────── 給前端的狀態 ─────────────────────────

def _check_stale(st):
    """本店存檔之後，同區別店的班表或本店的店別基本設定有沒有改過（有＝存著的檢查結果可能過期）"""
    try:
        own = (_dir() / f"{_key(st['store'], st['ym'])}.json").stat().st_mtime
        others = [_cfg_path(st["store"])]
        rg = _region_of(st["store"])
        if rg:
            others += [_dir() / f"{_key(s, st['ym'])}.json" for s in REGIONS[rg] if s != int(st["store"])]
        return any(p.exists() and p.stat().st_mtime > own for p in others)
    except OSError:
        return False


def _public(st):
    k = _key(st["store"], st["ym"])
    with _glock:
        job = dict(_jobs.get(k) or {})
    if job.get("status") == "running":
        job["secs"] = int(time.time() - job.get("t0", time.time()))
    tpl = st.get("template")
    cfg = cfg_load(st["store"])
    ext = _ext(st, cfg)
    shifts = ext.get("shifts") or {}
    cells = {}
    if st.get("schedule") and tpl:
        dtypes = set(tpl["info"]["dayTypes"])
        for eid, row in st["schedule"].items():
            cells[eid] = {}
            for d, v in row.items():
                try:
                    c = parse_cell(v, shifts, dtypes)
                    c["hours"] = cell_hours(c)
                except ValueError:
                    c = {"type": "", "segs": [], "dtag": "", "hours": 0, "bad": True}
                cells[eid][d] = c
    month_cost = round(sum(u.get("twd", 0) for u in st.get("usage") or []), 2)
    s104 = None
    if st.get("s104"):
        x = st["s104"]
        diffs = []
        if tpl and st.get("schedule") and x.get("diff"):
            names = {e["id"]: e["name"] for e in tpl["info"]["emps"]}
            for eid, d, a, b in _diff(st, st["schedule"], x["cells"])[:30]:
                diffs.append({"name": names.get(eid, eid), "day": d, "ws": a, "s104": b})
        s104 = {"name": x.get("name"), "status": x.get("status"), "modify": x.get("modify"), "fetched": x.get("fetched"),
                "diff": x.get("diff"), "diffs": diffs, "missing": x.get("missing") or [], "emps": x.get("emps") or [],
                "unknown": x.get("unknown") or 0}
    chk = st.get("check")
    report = st.get("report", "")
    if chk and tpl and st.get("schedule") and ("cover" not in chk or _check_stale(st)):
        # 舊存檔沒有每天人數；或同區別店的班表（調店格子）、店別基本設定在本店存檔之後改過 → 現算一次（不存檔）
        try:
            chk = run_check(st)
            if st.get("draft") or _auto_report(report):
                report = live_report(dict(st, check=chk))
        except Exception:   # noqa: BLE001
            pass
    return {
        "summary": live_summary(st, chk),
        "ok": True, "store": st["store"], "storeName": _sname(st["store"]), "ym": st["ym"],
        "settingsFiles": [{"name": f["name"], "chars": len(f["text"])} for f in st.get("settings_files") or []],
        "note": st.get("note", ""), "s104": s104,
        "template": ({"name": tpl["name"], "info": tpl["info"], "source": tpl.get("source", "upload"),
                      "fetched": tpl.get("fetched")} if tpl else None),
        "ext": ext or None, "schedule": st.get("schedule"), "cells": cells, "report": report,
        "check": chk, "chat": (st.get("chat") or [])[-80:], "canUndo": bool(st.get("history")),
        "plans": {}, "plan": st.get("plan"),   # 2026-11-10：只排兼職優先，不再回兩個方案
        "job": job, "cost": month_cost, "updated": st.get("updated", ""),
        "cfg": cfg, "mset": mset_view(st) if tpl else None, "formReady": form_ready(st, cfg),
        "bonus": bonus_calc(st, ext), "stores": {**STORES, **TEST_STORES}, "nuo": nuo_ledger(st, cfg) if tpl else [], "gov_hol": _gov_hol(st) if tpl else [],
    }


# ───────────────────────── 入口 ─────────────────────────

def handle(path, body):
    """server.py 把 /api/schedule/* 交給這裡；回傳 (HTTP 狀態碼, JSON)"""
    body = body if isinstance(body, dict) else {}
    act = path.rsplit("/", 1)[-1]
    who = body.get("as")
    ok, err = verify(who, body.get("pwd"))
    if not ok:
        return 401, {"ok": False, "error": err}
    if act == "login":
        return 200, {"ok": True, "as": str(who), "stores": STORES}
    if act == "sync104":
        if str(who) != "hq":
            return 403, {"ok": False, "error": "只有總部可以送 104 班表"}
        if body.get("dash_token"):   # 2026-11-09：排班分析儀表板寫入 token，push.py 從 secrets.env 帶來，存硬碟
            _dash104_cfg({"token": str(body["dash_token"])})
        results, dash = sync104(body.get("payload"), force_dash=bool(body.get("force_dash")))
        dash["has_token"] = bool(_dash104_cfg().get("token"))
        return 200, {"ok": True, "results": results, "dash104": dash}
    if act == "cred104_set":   # 2026-11-09：Mac 排程把 104 排班唯讀帳號送來存硬碟（不回顯）
        if str(who) != "hq":
            return 403, {"ok": False, "error": "只有總部可以設定"}
        email, pwd = str(body.get("email") or "").strip(), str(body.get("pwd104") or "")
        if not email or not pwd:
            return 400, {"ok": False, "error": "email／pwd104 必填"}
        _cred104({"email": email, "pwd": pwd, "set_at": time.time()})
        return 200, {"ok": True, "email": email}
    if act == "cred104_clear":   # 2026-11-09 深夜：無頭瀏覽器登入被 104 的 Cloudflare 機器人防護擋（403），不可繞過 → 把存的帳密刪掉
        if str(who) != "hq":
            return 403, {"ok": False, "error": "只有總部可以清除"}
        p = _dir() / CRED104
        try:
            if p.exists():
                p.unlink()
        except OSError as ex:
            return 500, {"ok": False, "error": str(ex)}
        return 200, {"ok": True, "has_cred": bool(_cred104().get("email"))}
    if act == "cloud104_status":   # 2026-11-09：雲端 06:05 自動同步的狀態
        if str(who) != "hq":
            return 403, {"ok": False, "error": "只有總部可以看"}
        c = _cred104()
        return 200, {"ok": True, "has_cred": bool(c.get("email")), "email": c.get("email"), "status": _cloud104_status()}
    if act == "cloud104_run":   # 2026-11-09：立刻跑一次（selftest=true 只測無頭瀏覽器）
        if str(who) != "hq":
            return 403, {"ok": False, "error": "只有總部可以觸發"}
        if (_cloud104_status() or {}).get("running"):
            return 409, {"ok": False, "error": "正在跑"}
        threading.Thread(target=cloud104_run, args=("manual", bool(body.get("selftest"))), daemon=True).start()
        return 200, {"ok": True, "started": True}
    if act == "dash104_status":   # 2026-11-09：排班分析儀表板寫入進度（push.py 等它跑完）
        if str(who) != "hq":
            return 403, {"ok": False, "error": "只有總部可以看"}
        return 200, {"ok": True, "status": _dash104_status()}
    if act == "sync104_status":
        if str(who) != "hq":
            return 403, {"ok": False, "error": "只有總部可以看"}
        out = []
        for p in sorted(_dir().glob("*-*.json")):
            try:
                s = json.loads(p.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            x = s.get("s104") or {}
            out.append({"store": s.get("store"), "ym": s.get("ym"), "name": x.get("name"), "status": x.get("status"),
                        "fetched": x.get("fetched"), "diff": x.get("diff"),
                        "template": (s.get("template") or {}).get("source")})
        return 200, {"ok": True, "items": out}
    if act in ("region_state", "region_draft", "region_pick", "region_chat"):   # 🗺 區域排班（2026-11-06；目前總部登入可用，魔導師登入待權限中心加）
        if str(who) != "hq":
            return 403, {"ok": False, "error": "只有總部／魔導師可以用區域排班"}
        rg, ym = str(body.get("region") or ""), str(body.get("ym") or "")
        if rg not in REGIONS or not re.fullmatch(r"20\d{2}(0[1-9]|1[0-2])", ym):
            return 400, {"ok": False, "error": "區域或月份不對"}
        jk = f"rg:{rg}:{ym}"
        with _glock:
            running = (_jobs.get(jk) or {}).get("status") == "running"
        if act == "region_state":
            return 200, _region_public(rg, ym)
        if running:
            return 409, {"ok": False, "error": "整區還在排，請等它跑完"}
        if act == "region_draft":
            nr = [x for x in _region_ready(rg, ym) if not x["ready"]]
            if nr:
                return 400, {"ok": False, "error": "這幾家店還不能排：" + "；".join(f"{x['name']}（{x['why']}）" for x in nr)}
            _job(jk, status="running", step="準備整區排班…")
            threading.Thread(target=_region_worker, args=(rg, ym), daemon=True).start()
            return 200, _region_public(rg, ym)
        m = str(body.get("plan") or "")
        if act == "region_chat":
            text = str(body.get("text") or "").strip()[:2000]
            if not text:
                return 400, {"ok": False, "error": "請輸入要調整的內容"}
            p = _rg_path(rg, ym)
            if not p.exists() or m not in (json.loads(p.read_text(encoding="utf-8")).get("plans") or {}):
                return 400, {"ok": False, "error": "還沒有整區初稿，請先按「整區排初稿」"}
            _job(jk, status="running", step="AI 正在想…")
            threading.Thread(target=_region_chat_worker, args=(rg, ym, m, text), daemon=True).start()
            return 200, _region_public(rg, ym)
        return _region_pick(rg, ym, m)
    try:
        store = int(body.get("store") if str(who) == "hq" else who)
    except (TypeError, ValueError):
        return 400, {"ok": False, "error": "請選擇門市"}
    ym = str(body.get("ym") or "")
    if not (store in STORES or (store in TEST_STORES and str(who) == "hq")) or not re.fullmatch(r"20\d{2}(0[1-9]|1[0-2])", ym):
        return 400, {"ok": False, "error": "門市或月份不對"}
    if not _can(store, who):
        return 403, {"ok": False, "error": "只能看自己店的班表"}
    k = _key(store, ym)
    lk = _lock(k)
    with _glock:
        running = (_jobs.get(k) or {}).get("status") == "running"

    if act == "el_year":   # 2026-11-13：班表下方每人每月早／晚班次數（畫面載入後另外讀，不拖慢主畫面）
        with lk:
            st = _load(store, ym)
        return 200, el_year(st)
    if act == "state":
        rev_est_ensure(store, ym)
        with lk:
            st = _load(store, ym)
        if st.get("schedule") and _rev_stale(st):
            _rev_refresh_bg(store, ym)
        with lk:
            return 200, _public(_load(store, ym))

    # 2026-11-11：這道擋原本放在 cfg_save／mset_save／mark 後面，AI 處理中存的設定會被 AI 存檔整份蓋掉
    if act in ("upload", "draft", "edit", "undo", "chat", "note", "cfg_save", "mset_save", "mark", "plan_pick", "remove_settings", "export") and running:
        return 409, {"ok": False, "error": "AI 還在處理上一個動作，請等它完成"}

    if act == "cfg_save":
        try:
            cfg = cfg_clean(body.get("cfg") or {})
        except ValueError as ex:
            return 400, {"ok": False, "error": str(ex)}
        with lk:
            cfg_save(store, cfg)
            st = _load(store, ym)
            if st.get("schedule"):
                st["check"] = run_check(st)
                _save(st)
            return 200, _public(st)

    if act == "mset_save":
        with lk:
            st = _load(store, ym)
            if not st.get("template"):
                return 400, {"ok": False, "error": "還沒有 104 範本（每天早上自動取得；或自己上傳）"}
            try:
                st["mset"] = mset_clean(st, body.get("mset") or {})
            except ValueError as ex:
                return 400, {"ok": False, "error": str(ex)}
            if st.get("schedule"):
                st["check"] = run_check(st)
            _save(st)
            return 200, _public(st)

    if act == "parse_preview":   # 2026-11-05：指休／特休／可上 打字時即時回顯「我讀成…」；ai＝離開格子才請 AI 看規則看不懂的
        with lk:
            st = _load(store, ym)
            if not st.get("template"):
                return 400, {"ok": False, "error": "還沒有 104 範本"}
            info, n_use = st["template"]["info"], len(st.get("usage") or [])
            shift_names = [r["name"] for r in (cfg_load(store).get("shifts") or [])]
            out = []
            for it in (body.get("items") or [])[:60]:
                k, txt = it.get("k"), str(it.get("text") or "")[:300]
                if k not in MARK_TXT:
                    out.append({"ok": False, "error": "欄位不對"})
                    continue
                try:
                    days, used, std = parse_days_smart(st, txt, info["days"], shift_names if k == "可上" else None, "", k, ai=bool(body.get("ai")))
                    out.append({"ok": True, "days": days, "ai": used, "std": std})
                except NeedAI as ex:
                    out.append({"ok": False, "needAi": True, "error": str(ex)})
                except ValueError as ex:
                    out.append({"ok": False, "error": str(ex)})
            if len(st.get("usage") or []) != n_use:
                _save(st)
            return 200, {"ok": True, "items": out}

    if act == "parse_log":   # 看不懂的說法（總部每月看一次，常見的加進規則）
        if str(who) != "hq":
            return 403, {"ok": False, "error": "只有總部可以看"}
        p = _dir() / "parse_log.jsonl"
        rows = [json.loads(x) for x in p.read_text(encoding="utf-8").splitlines() if x.strip()] if p.exists() else []
        return 200, {"ok": True, "items": rows[-500:]}

    if act == "mark":
        with lk:
            st = _load(store, ym)
            if not st.get("template"):
                return 400, {"ok": False, "error": "還沒有 104 範本"}
            try:
                for c in (body.get("marks") or [])[:200]:
                    mark_set(st, str(c.get("emp")), int(c.get("day", 0)), str(c.get("value") or ""))
            except (ValueError, TypeError) as ex:
                return 400, {"ok": False, "error": str(ex)}
            if st.get("schedule"):
                st["check"] = run_check(st)
            _save(st)
            return 200, _public(st)

    if act == "rev_refresh":
        with lk:
            st = _load(store, ym)
        if st.get("schedule"):
            _rev_refresh_bg(store, ym)
        return 200, {"ok": True}

    if act == "bonus_export":
        if str(who) == "hq" and body.get("all"):
            stores = sorted(STORES)
            label = "全店"
        else:
            stores = [store]
            label = f"{store}{_sname(store)}" + ("" if _sname(store).endswith("店") else "店")
        out, n = bonus_xlsx(stores, ym)
        if not n:
            return 400, {"ok": False, "error": "這個月還沒有班表，沒辦法算日獎金"}
        return 200, {"ok": True, "b64": base64.b64encode(out).decode("ascii"), "stores": n,
                     "filename": f"日獎金表_{ym}_{label}_{dt.datetime.now():%Y%m%d}.xlsx"}

    if act == "upload":
        b64 = str(body.get("b64") or "")
        name = str(body.get("name") or "")[:120]
        if not b64 or len(b64) > MAX_FILE_B64:
            return 400, {"ok": False, "error": "檔案是空的或太大（上限約 3MB）"}
        raw = base64.b64decode(b64)
        with lk:
            st = _load(store, ym)
            try:
                if body.get("kind") == "template":
                    st["template"] = {"name": name, "b64": b64, "info": template_info(raw)}
                    if st.get("schedule"):
                        st["check"] = run_check(st)
                else:
                    text = settings_to_text(name, raw)
                    st["settings_files"] = [f for f in st.get("settings_files") or [] if f["name"] != name][-2:] + [{"name": name, "text": text}]
            except Exception as ex:   # noqa: BLE001
                return 400, {"ok": False, "error": f"讀不了這個檔：{ex}"}
            _save(st)
            return 200, _public(st)

    if act == "note":
        with lk:
            st = _load(store, ym)
            st["note"] = str(body.get("note") or "")[:4000]
            _save(st)
            return 200, _public(st)

    if act == "remove_settings":
        with lk:
            st = _load(store, ym)
            st["settings_files"] = [f for f in st.get("settings_files") or [] if f["name"] != body.get("name")]
            _save(st)
            return 200, _public(st)

    if act == "draft":
        with lk:
            st = _load(store, ym)
        if not st.get("template"):
            return 400, {"ok": False, "error": "還沒有 104 範本（每天早上自動取得；或自己上傳）"}
        if not form_ready(st, cfg_load(store)) and not st.get("settings_files"):
            return 400, {"ok": False, "error": "請先填好「店別基本設定」的班別，和「本月設定」的平日／假日營收預估（或上傳 Excel 設定表）"}
        _job(k, status="running", step="準備排班…")
        threading.Thread(target=_draft_worker, args=(store, ym), daemon=True).start()
        with lk:
            return 200, _public(_load(store, ym))

    if act == "chat":
        text = str(body.get("text") or "").strip()[:2000]
        if not text:
            return 400, {"ok": False, "error": "請輸入要說的話"}
        with lk:
            st = _load(store, ym)
        if not st.get("schedule"):
            return 400, {"ok": False, "error": "要先排出初稿，才能用對話改班"}
        _job(k, status="running", step="AI 正在想…")
        threading.Thread(target=_chat_worker, args=(store, ym, text), daemon=True).start()
        with lk:
            out = _public(_load(store, ym))
        if (out.get("job") or {}).get("status") == "running":
            out["chat"] = out["chat"] + [{"role": "me", "text": text, "ts": time.time(), "pending": True}]
        return 200, out

    if act == "edit":
        with lk:
            st = _load(store, ym)
            if not st.get("schedule"):
                return 400, {"ok": False, "error": "還沒有班表"}
            try:
                chs = [dict(c, value=cell_from_span(c["span"])) if isinstance(c.get("span"), dict) else c for c in body.get("changes") or []]
                _apply(st, chs, "店長在日曆上改")
            except ValueError as ex:
                return 400, {"ok": False, "error": str(ex)}
            _save(st)
            return 200, _public(st)

    if act == "plan_pick":   # 2026-11-06：採用方案 A／B（換掉目前班表，舊的可「復原上一步」）
        m = str(body.get("plan") or "")
        with lk:
            st = _load(store, ym)
            p = (st.get("plans") or {}).get(m)
            if not p:
                return 400, {"ok": False, "error": "沒有這個方案，請重新排初稿"}
            if st.get("schedule"):
                st["history"] = (st.get("history") or [])[-19:] + [{"schedule": st["schedule"], "why": "換方案前"}]
            st["schedule"] = copy.deepcopy(p["schedule"])
            st["check"] = run_check(st)
            st["report"], st["plan"] = p["report"], m
            st["chat"] = (st.get("chat") or []) + [{"role": "ai", "text": f"已改用{PLAN_NAME.get(m, m)}。要改哪裡直接跟我說。", "ts": time.time()}]
            _save(st)
            return 200, _public(st)

    if act == "undo":
        with lk:
            st = _load(store, ym)
            if _undo(st) is None:
                return 400, {"ok": False, "error": "沒有可以復原的改動"}
            _save(st)
            return 200, _public(st)

    if act == "export":
        with lk:
            st = _load(store, ym)
            if not st.get("schedule") or not st.get("template"):
                return 400, {"ok": False, "error": "還沒有班表"}
            st["check"] = run_check(st)
            out = fill(st)
            _save(st)
        base = re.sub(r"\.xlsx?$", "", st["template"]["name"]) or f"{_sname(store)}店_{ym}"
        return 200, {"ok": True, "b64": base64.b64encode(out).decode("ascii"),
                     "filename": f"{base}_已排班.xlsx", "check": st["check"]}

    return 404, {"ok": False, "error": "沒有這個動作"}
