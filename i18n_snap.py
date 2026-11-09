"""🌏 平板外語食譜快照（2026-11-09 經營者：「未來如果我希望可以讓 100 人同時使用」→ 方案 A）。

為什麼要有這個：平板原本每開一支食譜都直接問 Google Apps Script（申請單 API 的 i18nList／i18nGet），
GAS 每次要當場讀試算表、組內容（約 5 秒），而 Google 帳號同時最多只跑 30 個 → 約 25～30 人同時點開就會被擋。
改成：GAS 把「平板清單＋每支×語言組好的內容」推到本伺服器存在硬碟（TRAINING_DATA_DIR/i18n_snap/），
平板向本伺服器讀，一支約 1 秒內、幾百人同時也撐得住；本伺服器讀不到（還沒推、重開中）→ 平板自動退回問 GAS（照舊保留）。

網址（都是 POST，body＝JSON；回覆格式跟 GAS 一樣 {ok, act, data}，平板不用分兩套）：
  /api/i18n/list  {tpw}             → {ok, act:'i18nList', data:{at, items}}
  /api/i18n/get   {tpw, id, lang}   → {ok, act:'i18nGet', data:{…}}；沒有這一版 → 404 code 'nosnap'（平板改問 GAS）
  /api/i18n/snap  {key, op, …}      → 只給 GAS 推資料（i18n.gs 的 i18nSnapPush_）：
        op 'manifest' → {vers:{'id|lang': ver}}        伺服器現有哪些版本（GAS 只推變了的）
        op 'put'      {items:[{id, lang, ver, data}]}   存組好的內容
        op 'list'     {at, items}                      存清單；清單上沒有（下架、取消勾、沒放行）的內容檔一併刪掉
平板密碼（tpw）：跟 GAS 一樣問權限中心 verify（dashboard tablet_i18n），對了記 2 小時（只存雜湊）；
  同一組密碼同時很多台來問，只有第一台真的去問權限中心，其他台等它的結果（不會一次把權限中心打爆）。
推送通行碼：GAS 指令碼屬性 I18N_SNAP_KEY；這裡只放 SHA-256。
"""
from __future__ import annotations

import hashlib
import json
import re
import ssl
import threading
import time
import urllib.parse
import urllib.request
from pathlib import Path

from storage_paths import atomic_write_json, data_dir

SNAP_KEY_SHA256 = "9b2d3b3fd95098c58c7799708862708a30937325f34b9cb3140435ffaab4827b"
AUTH_API = "https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec"
AUTH_DASH = "tablet_i18n"
LANGS = ("ja", "ko", "vi", "id", "th")
_ID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$")
_FINAL_RE = re.compile(r"密碼錯誤|已停用|查無|嘗試次數過多")   # 權限中心的明確答案（不是連線問題）

_LOCK = threading.RLock()
_LIST: dict | None = None          # {at, items}
_RCP: dict[str, tuple[str, dict]] = {}   # 'id|lang' → (ver, data)，記憶體快取（全部約 10MB 內）
_LOADED = False


def _ssl_ctx() -> ssl.SSLContext:
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


_CTX = _ssl_ctx()


def _dir() -> Path:
    d = data_dir() / "i18n_snap"
    (d / "rcp").mkdir(parents=True, exist_ok=True)
    return d


def _rcp_path(key: str) -> Path:
    i, l = key.split("|", 1)
    return _dir() / "rcp" / f"{i}_{l}.json"


def _load() -> None:
    """開機後第一次用到時把硬碟上的快照讀進記憶體（伺服器重開不必等 GAS 重推）。"""
    global _LIST, _LOADED
    with _LOCK:
        if _LOADED:
            return
        d = _dir()
        try:
            _LIST = json.loads((d / "list.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            _LIST = None
        for p in (d / "rcp").glob("*.json"):
            try:
                o = json.loads(p.read_text(encoding="utf-8"))
                _RCP[o["id"] + "|" + o["lang"]] = (o["ver"], o["data"])
            except (OSError, ValueError, KeyError):
                continue
        _LOADED = True


# ---------------- 平板密碼（問權限中心，同一組密碼只問一次） ----------------
_AUTH_OK: dict[str, float] = {}              # 雜湊 → 到期時間
_AUTH_BAD: dict[str, tuple[float, str]] = {}  # 雜湊 → (到期, 訊息)：錯的密碼 20 秒內不再問
_AUTH_WAIT: dict[str, threading.Event] = {}
_AUTH_LOCK = threading.Lock()


def _ask_center(pw: str) -> tuple[bool | None, str]:
    """回 (True, '') 對；(False, 訊息) 明確錯；(None, 訊息) 連不上。"""
    qs = urllib.parse.urlencode({"action": "verify", "dashboard": AUTH_DASH, "pwd": pw, "callback": "cb"})
    try:
        with urllib.request.urlopen(urllib.request.Request(AUTH_API + "?" + qs), timeout=25, context=_CTX) as r:
            t = r.read().decode("utf-8", "replace")
    except Exception as e:  # noqa: BLE001
        return None, "連不上權限中心（" + str(e)[:60] + "）"
    m = re.match(r"^\s*[A-Za-z_$][\w$]*\(([\s\S]*)\);?\s*$", t)
    try:
        j = json.loads(m.group(1) if m else t)
    except ValueError:
        return None, "權限中心回應看不懂"
    if j.get("ok"):
        return True, ""
    err = str(j.get("error") or "")
    return (False, err) if _FINAL_RE.search(err) or not err else (None, err)


def _check_pw(pw: str) -> tuple[bool | None, str]:
    pw = str(pw or "").strip()
    if not pw:
        return False, "請輸入平板密碼"
    h = hashlib.sha256((AUTH_DASH + "|" + pw).encode("utf-8")).hexdigest()
    for _ in range(3):
        now = time.time()
        with _AUTH_LOCK:
            if _AUTH_OK.get(h, 0) > now:
                return True, ""
            bad = _AUTH_BAD.get(h)
            if bad and bad[0] > now:
                return False, bad[1]
            ev = _AUTH_WAIT.get(h)
            mine = ev is None
            if mine:
                ev = _AUTH_WAIT[h] = threading.Event()
        if not mine:   # 別台正在問同一組密碼 → 等它的結果
            ev.wait(30)
            continue
        try:
            ok, msg = _ask_center(pw)
            with _AUTH_LOCK:
                if ok:
                    _AUTH_OK[h] = time.time() + 7200
                elif ok is False:
                    _AUTH_BAD[h] = (time.time() + 20, msg)
            return ok, msg
        finally:
            with _AUTH_LOCK:
                _AUTH_WAIT.pop(h, None)
            ev.set()
    return None, "權限中心忙碌，請再試一次"


def _auth_fail(act: str, ok: bool | None, msg: str) -> tuple[int, dict]:
    if ok is False:   # 跟 GAS 同一個代號 auth：平板會請店員重新輸入密碼
        return 401, {"ok": False, "act": act, "code": "auth", "msg": "平板密碼不對（" + msg + "）"}
    return 503, {"ok": False, "act": act, "code": "authnet", "msg": msg}   # 連不上 → 平板改問 GAS


# ---------------- 平板讀 ----------------
def _list(body: dict) -> tuple[int, dict]:
    ok, msg = _check_pw(body.get("tpw"))
    if not ok:
        return _auth_fail("i18nList", ok, msg)
    _load()
    with _LOCK:
        lst = _LIST
    if not lst:
        return 404, {"ok": False, "act": "i18nList", "code": "nosnap", "msg": "伺服器還沒有快照"}
    return 200, {"ok": True, "act": "i18nList", "data": lst}


def _get(body: dict) -> tuple[int, dict]:
    ok, msg = _check_pw(body.get("tpw"))
    if not ok:
        return _auth_fail("i18nGet", ok, msg)
    rid, lang = str(body.get("id") or "").strip().lower(), str(body.get("lang") or "").strip()
    _load()
    with _LOCK:
        it = next((x for x in (_LIST or {}).get("items", []) if x.get("id") == rid), None)
        hit = _RCP.get(rid + "|" + lang)
    ver = ((it or {}).get("ver") or {}).get(lang)
    if not it or (it.get("langs") or {}).get(lang) != "live" or not hit or hit[0] != ver:
        return 404, {"ok": False, "act": "i18nGet", "code": "nosnap", "msg": "伺服器沒有這一版"}
    return 200, {"ok": True, "act": "i18nGet", "data": hit[1]}


# ---------------- GAS 推資料 ----------------
def _snap(body: dict) -> tuple[int, dict]:
    global _LIST
    key = str(body.get("key") or "")
    if not key or hashlib.sha256(key.encode("utf-8")).hexdigest() != SNAP_KEY_SHA256:
        time.sleep(1)
        return 403, {"ok": False, "msg": "通行碼不對"}
    _load()
    op = body.get("op")
    if op == "manifest":
        with _LOCK:
            return 200, {"ok": True, "vers": {k: v[0] for k, v in _RCP.items()}, "list_at": (_LIST or {}).get("at", "")}
    if op == "put":
        n = 0
        for x in body.get("items") or []:
            rid, lang, ver, data = str(x.get("id") or ""), str(x.get("lang") or ""), str(x.get("ver") or ""), x.get("data")
            if not _ID_RE.match(rid) or lang not in LANGS or not ver or not isinstance(data, dict) or not isinstance(data.get("slides"), list):
                return 400, {"ok": False, "msg": "內容格式不對：" + rid + "|" + lang}
            k = rid + "|" + lang
            with _LOCK:
                atomic_write_json(_rcp_path(k), {"id": rid, "lang": lang, "ver": ver, "data": data})
                _RCP[k] = (ver, data)
            n += 1
        return 200, {"ok": True, "saved": n}
    if op == "list":
        items = body.get("items")
        if not isinstance(items, list):
            return 400, {"ok": False, "msg": "清單格式不對"}
        lst = {"at": str(body.get("at") or ""), "items": items}
        live = {str(it.get("id")) + "|" + l for it in items for l, st in (it.get("langs") or {}).items() if st == "live"}
        removed = 0
        with _LOCK:
            atomic_write_json(_dir() / "list.json", lst)
            _LIST = lst
            for k in [k for k in _RCP if k not in live]:
                try:
                    _rcp_path(k).unlink()
                except OSError:
                    pass
                _RCP.pop(k, None)
                removed += 1
        return 200, {"ok": True, "items": len(items), "removed": removed}
    return 400, {"ok": False, "msg": "不認得的動作"}


def handle(path: str, raw: bytes) -> tuple[int, dict]:
    try:
        body = json.loads((raw or b"{}").decode("utf-8") or "{}")
        if not isinstance(body, dict):
            raise ValueError
    except ValueError:
        return 400, {"ok": False, "msg": "送來的資料不是 JSON"}
    try:
        if path == "/api/i18n/list":
            return _list(body)
        if path == "/api/i18n/get":
            return _get(body)
        if path == "/api/i18n/snap":
            return _snap(body)
    except Exception as e:  # noqa: BLE001  任何意外都回 500，平板會改問 GAS
        return 500, {"ok": False, "code": "err", "msg": "伺服器錯誤：" + str(e)[:120]}
    return 404, {"ok": False, "msg": "找不到這個網址"}
