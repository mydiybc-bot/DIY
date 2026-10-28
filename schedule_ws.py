"""🗓 門市排班工作台（2026-10-26）

店長上傳「每月排班設定表」＋「104 匯入範本（時間排班）」→ Claude（Sonnet 5.5）排初稿 → 檢查程式（schedule_check104）把關，
沒過自動退回 Claude 修正（兩次不過換 Opus 5.5）→ 網頁日曆可點格子改、也可在對話框用講的改 → 下載填好的 104 匯入檔。

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
import urllib.parse
import urllib.request
import warnings

import openpyxl

import schedule_check104 as c104
from storage_paths import atomic_write_json, data_dir

warnings.filterwarnings("ignore")

AUTH_API = "https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec"
AUTH_TTL = 10 * 60
MODEL = "claude-sonnet-5-5"
MODEL_BACKUP = "claude-opus-5-5"
PRICE = {MODEL: (2.00, 10.00, 2.50, 0.20), MODEL_BACKUP: (4.00, 20.00, 5.00, 0.20)}   # 美元／百萬 tokens：輸入、輸出、快取寫、快取讀
TWD = 32
STORES = {1: "精明", 2: "草悟道", 3: "南京", 4: "士林", 5: "Focus", 6: "新竹", 7: "板橋", 8: "新店",
          9: "中壢", 10: "藝文", 11: "A13", 12: "SKM"}
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
    return {"store": int(store), "ym": ym, "settings_files": [], "note": "", "template": None,
            "ext": None, "schedule": None, "report": "", "history": [], "chat": [], "api_messages": [],
            "check": None, "usage": [], "updated": ""}


def _save(st):
    st["updated"] = dt.datetime.now().isoformat(timespec="seconds")
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
    v = (v or "").strip()
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
    return pos.replace("店", "") not in (title or "")


def _pos_store(pos):
    """崗位字串 → 店號（對不上回 None）"""
    pos = (pos or "").strip()
    if pos in POS_STORE:
        return POS_STORE[pos]
    for k, v in POS_STORE.items():
        if k.replace("店", "") and k.replace("店", "") in pos:
            return v
    return None


def cell_hours(c):
    return sum(_mins(e) - _mins(s) for s, e, _, _ in c["segs"]) / 60


# ───────────────────────── 設定表單（2026-10-28：取代 Excel 設定表） ─────────────────────────

TIME_RE = re.compile(r"^\d{1,2}:\d{2}$")


def _hhmm(v):
    v = str(v or "").strip().replace("：", ":")
    if not v:
        return ""
    m = re.match(r"^(\d{1,2}):?(\d{2})$", v)
    if not m or int(m.group(1)) > 23 or int(m.group(2)) > 59:
        raise ValueError(f"時間「{v}」要寫成 10:20")
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


def cfg_load(store):
    p = _cfg_path(store)
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            pass
    return {"shifts": [], "max_consec": 5, "late_early_limit": 0, "support": "", "rules": "", "carry": [], "updated": ""}


def cfg_clean(body):
    """店別基本設定：檢查＋整理（寫錯丟 ValueError）"""
    out = {"shifts": [], "carry": []}
    names = set()
    for r in (body.get("shifts") or [])[:20]:
        name = str(r.get("name") or "").strip()[:20]
        segs = []
        for a, b in (r.get("segs") or [])[:3]:
            a, b = _hhmm(a), _hhmm(b)
            if a and b:
                if _mins(b) <= _mins(a):
                    raise ValueError(f"班別「{name}」的時段 {a}-{b} 結束要晚於開始")
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
    out["max_consec"] = _num(body.get("max_consec"), 5, 1, 12, "連續上班上限")
    out["late_early_limit"] = _num(body.get("late_early_limit"), 0, 0, 31, "晚接早上限")
    out["support"] = str(body.get("support") or "").strip()[:1000]
    out["rules"] = str(body.get("rules") or "").strip()[:4000]
    for r in (body.get("carry") or [])[:60]:
        name = str(r.get("name") or "").strip()[:20]
        if not name:
            continue
        out["carry"].append({"name": name, "type": ("B" if str(r.get("type") or "").strip().upper() == "B" else "A"),
                             "date": str(r.get("date") or "").strip()[:12], "note": str(r.get("note") or "").strip()[:100]})
    return out


def cfg_save(store, cfg):
    cfg["updated"] = dt.datetime.now().isoformat(timespec="seconds")
    atomic_write_json(_cfg_path(store), cfg)
    return cfg


def _prev_ym(ym):
    y, m = int(ym[:4]), int(ym[4:])
    return f"{y - 1}12" if m == 1 else f"{y}{m - 1:02d}"


def mset_view(st):
    """本月設定（含預設值）：夥伴清單跟著 104 範本；正職／兼職、目標時數沒填的先抄上個月"""
    ms = copy.deepcopy(st.get("mset") or {})
    ms.setdefault("base", {})
    ms.setdefault("revenue", {})
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
                prev = {}
                if pp.exists():
                    try:
                        prev = (json.loads(pp.read_text(encoding="utf-8")).get("mset") or {}).get("emps") or {}
                    except (OSError, json.JSONDecodeError):
                        prev = {}
            cur = dict(prev.get(e["id"]) or {"type": "正職", "target": None, "min_hours": None, "note": ""})
            cur["_default"] = True
            ms["emps"][e["id"]] = cur
        if cur.get("type") == "正職" and not cur.get("target"):
            must = (s104.get(e["name"]) or {}).get("mustHour")
            cur["target_hint"] = must or ms["base"].get("std_hours")
    ms["emps"] = {k: v for k, v in ms["emps"].items() if any(e["id"] == k for e in info.get("emps") or [])}
    ms["marks"] = {k: v for k, v in ms["marks"].items() if any(e["id"] == k for e in info.get("emps") or [])}
    return ms


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
    ms["emps"] = emps
    days = {d["d"] for d in ((st.get("template") or {}).get("info") or {}).get("days") or []}
    sp = []
    for x in (body.get("special") or [])[:40]:
        d = _num(x.get("day"), None, 1, 31, "日期")
        if d is None:
            continue
        if days and d not in days:
            raise ValueError(f"這個月沒有 {d} 號")
        sp.append({"day": int(d), "revenue": _num(x.get("revenue"), None, 0, 10_000_000, "營收預估"),
                   "staff": str(x.get("staff") or "").strip()[:200], "note": str(x.get("note") or "").strip()[:200]})
    ms["special"] = sorted(sp, key=lambda x: x["day"])
    ms.setdefault("marks", st.get("mset", {}).get("marks") or {})
    ms["updated"] = dt.datetime.now().isoformat(timespec="seconds")
    return ms


MARKS = ("指休", "特休", "可上")


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
    ms = st.get("mset") or {}
    return bool(cfg.get("shifts")) and bool((ms.get("revenue") or {}).get("weekday")) and bool((ms.get("revenue") or {}).get("weekend"))


def form_text(st, cfg):
    """把表單組成跟 Excel 設定表同樣結構的文字給 AI"""
    info = st["template"]["info"]
    ms = mset_view(st)
    L = ["### 分頁「各店基本設定」", f"店名\t{info['title']}", "班別\t第1段\t第2段\t第3段\t平日每天需要幾人\t假日每天需要幾人"]
    for r in cfg.get("shifts") or []:
        segs = [f"{a}-{b}" for a, b in r["segs"]] + [""] * (3 - len(r["segs"]))
        L.append(f"{r['name']}\t" + "\t".join(segs) + f"\t{r.get('wd', 0)}\t{r.get('we', 0)}")
    L.append(f"連續上班上限\t{cfg.get('max_consec', 5)} 天")
    L.append(f"晚班接隔天早班\t{'不允許' if not cfg.get('late_early_limit') else '最多 ' + str(cfg['late_early_limit']) + ' 次'}")
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
    L.append("【指休／特休／兼職可上日】（指休＝那天一定要休；可上＝兼職只有這些天能排）")
    for e in info["emps"]:
        mk = ms["marks"].get(e["id"]) or {}
        parts = []
        for k in MARKS:
            ds = sorted(int(d) for d, v in mk.items() if v == k)
            if ds:
                parts.append(f"{k} " + "、".join(map(str, ds)))
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
            L.append(f"{x['day']} 日：" + "；".join(bits))
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
    demand = {}
    for d in info.get("days") or []:
        wk = d["wd"] in ("六", "日")
        need = []
        for r in cfg["shifts"]:
            need += [r["name"]] * int(r.get("we" if wk else "wd") or 0)
        demand[str(d["d"])] = need
    emps = []
    for e in info.get("emps") or []:
        v = ms["emps"].get(e["id"]) or {}
        emps.append({"id": e["id"], "name": e["name"], "type": v.get("type", "正職"),
                     "target_hours": v.get("target") or (v.get("target_hint") if v.get("type", "正職") == "正職" else None)})
    return {"shifts": shifts, "demand": demand, "employees": emps, "max_consec": cfg.get("max_consec", 5),
            "late_early_limit": cfg.get("late_early_limit", 0),
            "late_shift_ends": sorted({r["segs"][-1][1] for r in cfg["shifts"] if "晚" in r["name"]}),
            "early_shift_starts": sorted({r["segs"][0][0] for r in cfg["shifts"] if "早" in r["name"]}),
            "revenue": {"weekday": ms["revenue"].get("weekday"), "weekend": ms["revenue"].get("weekend"), "special": special},
            "notes": [], "from_form": True}


def _ext(st, cfg=None):
    """AI 排過用 AI 的（班別名以表單為準）；沒排過用表單算的"""
    cfg = cfg if cfg is not None else cfg_load(st["store"])
    form = _ext_from_form(st, cfg) if st.get("template") else {}
    ai = st.get("ext") or {}
    if not ai:
        return form
    if not form:
        return ai
    out = dict(form, **ai)
    out["shifts"] = dict(ai.get("shifts") or {}, **form["shifts"])
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

def run_check(st, schedule=None):
    """檢查程式：104 格式＋勞基法（check_104）＋本店需求人數／正職時數／晚接早。回傳 dict"""
    schedule = schedule if schedule is not None else st["schedule"]
    ext = _ext(st)
    shifts = ext.get("shifts") or {}
    info = st["template"]["info"]
    dtypes = set(info["dayTypes"])
    errors, warns = [], []
    parsed = {}
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
    hours = []
    if not errors:
        raw_t = base64.b64decode(st["template"]["b64"])
        out = fill(st, schedule)
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
        warns += rep.warn
        for emp, _, inf in rows:
            tot = sum(b - a for _, segs, _ in inf for a, b in segs) / 60
            # 104 把休息日、國定假日出勤列為加班，不算進總工時 → 跟正職標準時數比的是「正常工時」
            ot = sum(b - a for t, segs, _ in inf if t in ("休息日", "國定假日") for a, b in segs) / 60
            cnt = lambda x: sum(1 for t, _, _ in inf if t == x)
            hours.append({"id": emp["id"], "name": emp["name"], "days": sum(1 for _, s, _ in inf if s),
                          "hours": round(tot, 2), "regular": round(tot - ot, 2), "ot": round(ot, 2),
                          "例假": cnt("例假日"), "休息": cnt("休息日"), "國定": cnt("國定假日"), "特休": cnt("特休")})
    # 本店規則
    demand = ext.get("demand") or {}
    late_starts = set(ext.get("late_shift_ends") or [])
    early_starts = set(ext.get("early_shift_starts") or [])
    for d in info["days"]:
        need = demand.get(str(d["d"])) or []
        if not need:
            continue
        working = []
        for eid in parsed:
            c = parsed[eid][d["d"]]
            here = [s for s in c["segs"] if not _elsewhere(s[3], info["title"])]
            if here:
                working.append(here)
        if len(working) < len(need):
            errors.append(f"{d['label']} 需要 {len(need)} 人（{'、'.join(need)}），只排了 {len(working)} 人")
            continue
        free = list(working)
        miss = []
        for nm in need:
            sh = shifts.get(nm)
            if not sh:
                continue
            hit = next((w for w in free if w[0][0] == sh[0][0]), None)
            if hit:
                free.remove(hit)
            else:
                miss.append(nm)
        if miss:
            warns.append(f"{d['label']} 可能缺{'、'.join(miss)}（沒有人在那個時間上班）")
    for emp in ext.get("employees") or []:
        tgt = emp.get("target_hours")
        h = next((x for x in hours if x["id"] == emp.get("id")), None)
        if tgt and h and abs(h["regular"] - float(tgt)) > 0.01:
            extra = f"（另有休息日／國定假日出勤 {h['ot']:g}h，104 列加班）" if h["ot"] else ""
            warns.append(f"{h['name']}：正常工時 {h['regular']:g}h，目標 {float(tgt):g}h{extra}")
    marks = ((st.get("mset") or {}).get("marks") or {})
    etype = {e.get("id"): e.get("type") for e in ext.get("employees") or []}
    for eid, mk in marks.items():
        if eid not in parsed:
            continue
        allow = {int(d) for d, v in mk.items() if v == "可上"}
        for d in info["days"]:
            c = parsed[eid][d["d"]]
            v = mk.get(str(d["d"]), "")
            if v == "指休" and c["segs"]:
                errors.append(f"{names[eid]} {d['label']}：本月設定是指休，卻排了班")
            elif v == "特休" and "特休" not in (c["type"] + c["dtag"]):
                warns.append(f"{names[eid]} {d['label']}：本月設定要特休，班表沒有標特休")
            elif allow and etype.get(eid) == "兼職" and c["segs"] and d["d"] not in allow:
                warns.append(f"{names[eid]} {d['label']}：兼職只有勾「可上」的日子能排，這天沒勾")
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
    return {"ok": not errors, "errors": errors, "warns": warns, "hours": hours}


def check_text(chk):
    lines = ["【每人時數】"]
    for h in chk["hours"]:
        lines.append(f"{h['id']} {h['name']}：上班 {h['days']} 天、共 {h['hours']:g}h（正常工時 {h['regular']:g}h＋休息日／國定出勤 {h['ot']:g}h）、例假 {h['例假']}、休息 {h['休息']}、國定 {h['國定']}、特休 {h['特休']}")
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
        _jobs[k] = dict(kw, ts=time.time())


def _draft_worker(store, ym):
    k = _key(store, ym)
    lk = _lock(k)
    try:
        with lk:
            st = _load(store, ym)
        system = RULES + DRAFT_CONTRACT
        messages = [{"role": "user", "content": _draft_user(st)}]
        report, ext, sched, chk = "", None, None, None
        plan = [MODEL, MODEL, MODEL, MODEL_BACKUP]
        for i, model in enumerate(plan):
            _job(k, status="running", step=("AI 正在排班（約 3～5 分鐘）" if i == 0 else f"檢查沒過，AI 正在修正（第 {i} 次{'，換 Opus' if model == MODEL_BACKUP else ''}）"))
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
            st.update(ext=ext, schedule=sched, report=report, check=chk, api_messages=[],
                      chat=(st.get("chat") or []) + [{"role": "ai", "text": "✅ 初稿排好了。" + ("檢查全部通過。" if chk["ok"] else f"還有 {len(chk['errors'])} 項 ❌ 沒改掉，請看下方檢查結果或跟我說要怎麼改。"), "ts": time.time()}])
            _save(st)
        _job(k, status="done", step="初稿完成")
    except Exception as ex:   # noqa: BLE001 — 背景工作要把錯誤留給畫面看
        _job(k, status="error", step=f"排班失敗：{type(ex).__name__}: {str(ex)[:300]}")


# ───────────────────────── 對話框 ─────────────────────────

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
    {"name": "undo", "description": "復原上一次改動（不論是你改的還是店長在日曆上改的）。",
     "input_schema": {"type": "object", "additionalProperties": False, "properties": {}}, "strict": True},
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
    for c in changes:
        st["schedule"][c["emp"]][str(int(c["day"]))] = c.get("value", "")
    st["check"] = run_check(st)
    return st["check"]


def _undo(st):
    if not st.get("history"):
        return None
    st["schedule"] = st["history"].pop()["schedule"]
    st["check"] = run_check(st)
    return st["check"]


def _chat_system(st):
    return (RULES + CHAT_RULES + "\n## 本店這個月的設定（AI 從設定表讀到的）\n" +
            json.dumps(_ext(st), ensure_ascii=False) +
            "\n\n## 104 範本\n- 班表：" + st["template"]["info"]["title"] +
            "\n- 日期：" + "、".join(d["label"] for d in st["template"]["info"]["days"]) +
            "\n- 可用標籤：" + "、".join(st["template"]["info"]["tags"]) +
            "\n- 可用崗位：" + "、".join(st["template"]["info"]["positions"]))


def _chat_turn(st, text, model, lk):
    """跑一句對話（可能多次呼叫工具）；回傳 (回覆, 改了幾格)"""
    st["api_messages"].append({"role": "user", "content": f"（目前班表）\n{table_text(st)}\n\n（目前檢查）\n{check_text(st['check'])}\n\n店長：{text}"})
    changed = 0
    for _ in range(8):
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
                else:
                    chk = _apply(st, b.input.get("changes") or [], "對話：" + text[:40])
                    changed += len(b.input.get("changes") or [])
                    out = "已改。\n" + check_text(chk) + "\n\n改後班表：\n" + table_text(st)
            except ValueError as ex:
                out = f"這次沒有改任何格子：{ex}"
            results.append({"type": "tool_result", "tool_use_id": b.id, "content": out})
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
        reply, changed = _chat_turn(st, text, MODEL, lk)
        if len(st["check"]["errors"]) > n_err:
            # Sonnet 改完反而多了 ❌ → 退回這句之前，換 Opus 重做
            _job(k, status="running", step="換更強的模型再想一次…")
            usage = st["usage"]
            st = before
            st["usage"] = usage
            reply, changed = _chat_turn(st, text, MODEL_BACKUP, lk)
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


def sync104(payload):
    """收 tools/sync104/push.py 送來的 104 班表；回傳每份的處理結果"""
    results = []
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
            _save(st)
            results.append(res)
    return results


# ───────────────────────── 營業額與日獎金（2026-10-28 經營者：班表要有營業額及日獎金、產各店日獎金表給會計） ─────────────────────────

def _ssl_ctx():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        return ssl.create_default_context()


def fetch_rev(store, ym):
    """排班分析 GAS getStoreDailyData（bonus=1）→ {日: {net, voucher, qty, rev}}；失敗丟 Exception"""
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
    if c["type"]:
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
        name = f"{store} {STORES.get(store, '')}"
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

def _public(st):
    k = _key(st["store"], st["ym"])
    with _glock:
        job = dict(_jobs.get(k) or {})
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
    return {
        "ok": True, "store": st["store"], "storeName": STORES.get(st["store"], ""), "ym": st["ym"],
        "settingsFiles": [{"name": f["name"], "chars": len(f["text"])} for f in st.get("settings_files") or []],
        "note": st.get("note", ""), "s104": s104,
        "template": ({"name": tpl["name"], "info": tpl["info"], "source": tpl.get("source", "upload"),
                      "fetched": tpl.get("fetched")} if tpl else None),
        "ext": ext or None, "schedule": st.get("schedule"), "cells": cells, "report": st.get("report", ""),
        "check": st.get("check"), "chat": (st.get("chat") or [])[-80:], "canUndo": bool(st.get("history")),
        "job": job, "cost": month_cost, "updated": st.get("updated", ""),
        "cfg": cfg, "mset": mset_view(st) if tpl else None, "formReady": form_ready(st, cfg),
        "bonus": bonus_calc(st, ext), "stores": STORES,
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
        return 200, {"ok": True, "results": sync104(body.get("payload"))}
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
    try:
        store = int(body.get("store") if str(who) == "hq" else who)
    except (TypeError, ValueError):
        return 400, {"ok": False, "error": "請選擇門市"}
    ym = str(body.get("ym") or "")
    if store not in STORES or not re.fullmatch(r"20\d{2}(0[1-9]|1[0-2])", ym):
        return 400, {"ok": False, "error": "門市或月份不對"}
    if not _can(store, who):
        return 403, {"ok": False, "error": "只能看自己店的班表"}
    k = _key(store, ym)
    lk = _lock(k)
    with _glock:
        running = (_jobs.get(k) or {}).get("status") == "running"

    if act == "state":
        with lk:
            st = _load(store, ym)
        if st.get("schedule") and _rev_stale(st):
            _rev_refresh_bg(store, ym)
        with lk:
            return 200, _public(_load(store, ym))

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
            label = f"{store}{STORES[store]}店"
        out, n = bonus_xlsx(stores, ym)
        if not n:
            return 400, {"ok": False, "error": "這個月還沒有班表，沒辦法算日獎金"}
        return 200, {"ok": True, "b64": base64.b64encode(out).decode("ascii"), "stores": n,
                     "filename": f"日獎金表_{ym}_{label}_{dt.datetime.now():%Y%m%d}.xlsx"}

    if act in ("upload", "draft", "edit", "undo", "chat", "note", "cfg_save", "mset_save", "mark") and running:
        return 409, {"ok": False, "error": "AI 還在處理上一個動作，請等它完成"}

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
        _job(k, status="running", step="AI 正在排班（約 3～5 分鐘）")
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
                _apply(st, body.get("changes") or [], "店長在日曆上改")
            except ValueError as ex:
                return 400, {"ok": False, "error": str(ex)}
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
        base = re.sub(r"\.xlsx?$", "", st["template"]["name"]) or f"{STORES[store]}店_{ym}"
        return 200, {"ok": True, "b64": base64.b64encode(out).decode("ascii"),
                     "filename": f"{base}_已排班.xlsx", "check": st["check"]}

    return 404, {"ok": False, "error": "沒有這個動作"}
