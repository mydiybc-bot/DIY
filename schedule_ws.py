"""🗓 門市排班工作台（2026-10-26）

店長上傳「每月排班設定表」＋「104 匯入範本（時間排班）」→ Claude（Sonnet 5.5）排初稿 → 檢查程式（schedule_check104）把關，
沒過自動退回 Claude 修正（兩次不過換 Opus 5.5）→ 網頁日曆可點格子改、也可在對話框用講的改 → 下載填好的 104 匯入檔。

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
RULES = open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "schedule_rules.md"), encoding="utf-8").read()
MAX_FILE_B64 = 4_000_000
MAX_SETTINGS_CHARS = 80_000
MAX_API_MESSAGES = 60

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


def cell_hours(c):
    return sum(_mins(e) - _mins(s) for s, e, _, _ in c["segs"]) / 60


# ───────────────────────── 填 104 範本 ─────────────────────────

def fill(st, schedule=None):
    """把班表填進店長上傳的原始範本，回傳 xlsx bytes（只改格子的值，下拉選單、合併儲存格不動）"""
    schedule = schedule if schedule is not None else st["schedule"]
    raw = base64.b64decode(st["template"]["b64"])
    wb = openpyxl.load_workbook(io.BytesIO(raw))
    ws = wb.worksheets[0]
    hdr, days, emps = c104.layout(ws)
    shifts = (st.get("ext") or {}).get("shifts") or {}
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
    ext = st.get("ext") or {}
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
            here = [s for s in c["segs"] if "支援" not in s[2] and not s[3]]
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
    for f in st["settings_files"]:
        parts.append(f"## 每月排班設定表（{f['name']}）\n{f['text']}")
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
    shifts = (st.get("ext") or {}).get("shifts") or {}
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
            json.dumps(st.get("ext") or {}, ensure_ascii=False) +
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


# ───────────────────────── 給前端的狀態 ─────────────────────────

def _public(st):
    k = _key(st["store"], st["ym"])
    with _glock:
        job = dict(_jobs.get(k) or {})
    tpl = st.get("template")
    shifts = (st.get("ext") or {}).get("shifts") or {}
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
    return {
        "ok": True, "store": st["store"], "storeName": STORES.get(st["store"], ""), "ym": st["ym"],
        "settingsFiles": [{"name": f["name"], "chars": len(f["text"])} for f in st.get("settings_files") or []],
        "note": st.get("note", ""), "template": ({"name": tpl["name"], "info": tpl["info"]} if tpl else None),
        "ext": st.get("ext"), "schedule": st.get("schedule"), "cells": cells, "report": st.get("report", ""),
        "check": st.get("check"), "chat": (st.get("chat") or [])[-80:], "canUndo": bool(st.get("history")),
        "job": job, "cost": month_cost, "updated": st.get("updated", ""),
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
            return 200, _public(_load(store, ym))

    if act in ("upload", "draft", "edit", "undo", "chat", "note") and running:
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
        if not st.get("template") or not st.get("settings_files"):
            return 400, {"ok": False, "error": "要先上傳「每月排班設定表」和「104 匯入範本」"}
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
