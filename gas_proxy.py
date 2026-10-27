"""儀表板 → Google Apps Script 代轉（2026-10-27，經營者同意）。

為什麼要有這個：儀表板原本由手機瀏覽器直接向 script.google.com 要資料，Google 會順便檢查
手機上登入的 Google 帳號，某些帳號會被擋（畫面「很抱歉，目前無法開啟這個檔案」；無痕正常）。
改由本伺服器代為向 Google 取資料，Google 只看到匿名的伺服器，手機登什麼帳號都能用。

網址：GET  /api/gas/<部署ID>?action=…        → 回 JSON（Google 回的 JSONP 會拆成純 JSON）
      POST /api/gas/<部署ID>  body＝JSON 字串 → 原樣轉給 Google、回 JSON
只放行 static/*.html 裡出現過的部署 ID（開機時掃一次），別人不能拿本伺服器當跳板。
"""
from __future__ import annotations

import json
import re
import ssl
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

_ID_RE = re.compile(r"script\.google\.com/macros/s/(AKfycb[\w-]+)/exec")
_ROOT = Path(__file__).parent
_STATIC = _ROOT / "static"
_TIMEOUT = 60          # Google 端偶爾慢（冷啟動），給足
_MAX_BODY = 512 * 1024  # 寫入內容上限（evoSnap 一次最多 600 列，遠在這之下）


def _ssl_ctx() -> ssl.SSLContext:
    """Render（Linux）系統憑證可用；本機 Mac 的 python.org Python 沒裝憑證會 CERTIFICATE_VERIFY_FAILED，
    有 certifi（requirements.txt 已列）就用它的憑證包。"""
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


_CTX = _ssl_ctx()


def _allowed_ids() -> set[str]:
    ids: set[str] = set()
    try:
        for p in _STATIC.glob("*.html"):
            try:
                ids.update(_ID_RE.findall(p.read_text(encoding="utf-8", errors="ignore")))
            except OSError:
                continue
    except OSError:
        pass
    return ids


_ALLOWED = _allowed_ids()


def _parse(path: str):
    """'/api/gas/<id>' → id；不合格回 None。"""
    m = re.fullmatch(r"/api/gas/(AKfycb[\w-]+)/?", path)
    return m.group(1) if m else None


def _strip_jsonp(text: str) -> str:
    """Google 回的可能是 callback({...})，拆成 {...}；本來就是 JSON 就原樣回。"""
    t = text.strip()
    if t.startswith("{") or t.startswith("["):
        return t
    m = re.match(r"^[\w$]+\((.*)\)\s*;?\s*$", t, re.S)
    return m.group(1) if m else t


def _fetch(url: str, data: bytes | None = None) -> tuple[int, str]:
    req = urllib.request.Request(url, data=data, method="POST" if data is not None else "GET")
    req.add_header("User-Agent", "diybc-dashboard-proxy/1.0")
    if data is not None:
        req.add_header("Content-Type", "text/plain; charset=utf-8")
    try:
        with urllib.request.urlopen(req, timeout=_TIMEOUT, context=_CTX) as r:
            raw = r.read()
            ctype = (r.headers.get("Content-Type") or "").lower()
    except urllib.error.HTTPError as e:
        return 502, json.dumps({"ok": False, "error": f"Google 回應 {e.code}"}, ensure_ascii=False)
    except Exception as e:  # 逾時、連不上
        return 504, json.dumps({"ok": False, "error": f"Google 連線失敗：{e.__class__.__name__}"}, ensure_ascii=False)
    text = raw.decode("utf-8", errors="replace")
    if "html" in ctype and not text.lstrip().startswith(("{", "[")):
        # Google 回的是 HTML（錯誤頁、登入頁），不是資料
        return 502, json.dumps({"ok": False, "error": "Google 回的不是資料（可能忙碌中），請稍後再試"}, ensure_ascii=False)
    body = _strip_jsonp(text)
    try:
        json.loads(body)
    except ValueError:
        return 502, json.dumps({"ok": False, "error": "Google 回的資料格式不對"}, ensure_ascii=False)
    return 200, body


def handle_get(path: str, query: str) -> tuple[int, str]:
    """回 (HTTP 狀態, JSON 字串)。"""
    sid = _parse(path)
    if not sid or sid not in _ALLOWED:
        return 404, json.dumps({"ok": False, "error": "不在放行清單"}, ensure_ascii=False)
    q = urllib.parse.parse_qs(query, keep_blank_values=True)
    q.pop("callback", None)  # 代轉一律回純 JSON，不需要 JSONP
    qs = urllib.parse.urlencode(q, doseq=True)
    url = f"https://script.google.com/macros/s/{sid}/exec" + (f"?{qs}" if qs else "")
    return _fetch(url)


def handle_post(path: str, body: bytes) -> tuple[int, str]:
    sid = _parse(path)
    if not sid or sid not in _ALLOWED:
        return 404, json.dumps({"ok": False, "error": "不在放行清單"}, ensure_ascii=False)
    if len(body) > _MAX_BODY:
        return 413, json.dumps({"ok": False, "error": "內容太大"}, ensure_ascii=False)
    url = f"https://script.google.com/macros/s/{sid}/exec"
    return _fetch(url, data=body)
