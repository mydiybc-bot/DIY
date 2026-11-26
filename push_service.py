"""🔔 自己做 App 推播（2026-11-26，經營者：「工作清單發出後，被指定處理的人員如何知道有工作要處理」→「好，做推播」）。

為什麼在這台伺服器：手機推播（Web Push）要用 ES256 簽章＋加密內容，Google Apps Script 做不到；
「誰該收到、收到什麼」仍由工作清單 GAS（公告及工作清單 API）決定，這裡只負責「把訊息送到手機」。

網址（POST，body＝JSON）：
  /api/push/key    → {ok, key}            App 訂閱時要的公開金鑰（任何人可讀，本來就是公開的）
  /api/push/setup  {hq_pw}                → {ok, token}  體驗中心密碼（dim_auth task_hq）驗過才發送件通行碼；
                                            每次 setup 都換新通行碼（舊的失效），通行碼只存雜湊
  /api/push/send   {token, items:[{sub:{endpoint,keys:{p256dh,auth}}, payload:{title,body,url,n,tag}}]}
                   → {ok, sent, gone:[endpoint], errors:[...]}   gone＝手機已取消訂閱（404／410），GAS 會刪掉
金鑰與通行碼雜湊存在資料夾 TRAINING_DATA_DIR（Render 掛的硬碟，重開不會掉）：push_vapid.pem、push_vapid.json。
"""
from __future__ import annotations

import base64
import hashlib
import hmac
import json
import secrets
import ssl
import threading
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor

from storage_paths import FILE_LOCK, atomic_write_json, data_file

AUTH_API = "https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec"
SUB_CLAIM = "https://diybc-dash.onrender.com"   # VAPID 聯絡方式（不放個人信箱）
MAX_ITEMS = 300
_KEY_LOCK = threading.Lock()


def _b64url(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode().rstrip("=")


def _ssl_ctx() -> ssl.SSLContext:
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


def _state() -> dict:
    p = data_file("push_vapid.json")
    if p.exists():
        try:
            return json.loads(p.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _ensure_keys() -> dict:
    """第一次用到時產生 VAPID 金鑰（P-256），之後一直用同一把（換了手機的舊訂閱就全部失效）。"""
    with _KEY_LOCK:
        st = _state()
        pem = data_file("push_vapid.pem")
        if st.get("public") and pem.exists():
            return st
        from cryptography.hazmat.primitives import serialization
        from cryptography.hazmat.primitives.asymmetric import ec
        k = ec.generate_private_key(ec.SECP256R1())
        pem.write_bytes(k.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8, serialization.NoEncryption()))
        pub = k.public_key().public_bytes(serialization.Encoding.X962, serialization.PublicFormat.UncompressedPoint)
        st = {"public": _b64url(pub), "token_sha256": st.get("token_sha256", "")}
        with FILE_LOCK:
            atomic_write_json(data_file("push_vapid.json"), st)
        return st


def _verify_hq(pw: str) -> bool:
    pw = str(pw or "").strip()
    if not pw:
        return False
    url = AUTH_API + "?" + urllib.parse.urlencode({"action": "verify", "dashboard": "task_hq", "pwd": pw, "callback": "cb"})
    try:
        with urllib.request.urlopen(url, timeout=30, context=_ssl_ctx()) as r:
            t = r.read().decode("utf-8", "replace")
    except Exception:
        return False
    t = t.strip()
    if t.startswith("cb(") and t.endswith(");"):
        t = t[3:-2]
    elif t.startswith("cb(") and t.endswith(")"):
        t = t[3:-1]
    try:
        return bool(json.loads(t).get("ok"))
    except Exception:
        return False


def _token_ok(token: str) -> bool:
    want = _state().get("token_sha256", "")
    got = hashlib.sha256(str(token or "").encode()).hexdigest()
    return bool(want) and hmac.compare_digest(want, got)


def _send_one(item: dict) -> tuple[str, str, str]:
    """回 (結果 ok/gone/err, endpoint, 說明)"""
    from pywebpush import WebPushException, webpush
    sub = item.get("sub") or {}
    ep = str(sub.get("endpoint") or "")
    keys = sub.get("keys") or {}
    if not ep.startswith("https://") or not keys.get("p256dh") or not keys.get("auth"):
        return "err", ep, "訂閱資料不完整"
    payload = json.dumps(item.get("payload") or {}, ensure_ascii=False)
    if len(payload.encode()) > 3000:
        return "err", ep, "內容太長"
    try:
        webpush(subscription_info={"endpoint": ep, "keys": {"p256dh": keys["p256dh"], "auth": keys["auth"]}},
                data=payload, vapid_private_key=str(data_file("push_vapid.pem")),
                vapid_claims={"sub": SUB_CLAIM}, ttl=86400, headers={"Urgency": "high"}, timeout=20)
        return "ok", ep, ""
    except WebPushException as e:
        code = getattr(getattr(e, "response", None), "status_code", 0) or 0
        if code in (404, 410):
            return "gone", ep, str(code)
        return "err", ep, f"{code} {str(e)[:160]}"
    except Exception as e:  # 網路等
        return "err", ep, str(e)[:160]


def handle(path: str, data: dict) -> tuple[int, dict]:
    data = data or {}
    if path == "/api/push/key":
        return 200, {"ok": True, "key": _ensure_keys()["public"]}
    if path == "/api/push/setup":
        if not _verify_hq(data.get("hq_pw")):
            return 403, {"ok": False, "msg": "體驗中心密碼不對"}
        st = _ensure_keys()
        token = secrets.token_urlsafe(32)
        st["token_sha256"] = hashlib.sha256(token.encode()).hexdigest()
        with FILE_LOCK:
            atomic_write_json(data_file("push_vapid.json"), st)
        return 200, {"ok": True, "token": token, "key": st["public"]}
    if path == "/api/push/send":
        if not _token_ok(data.get("token")):
            return 403, {"ok": False, "msg": "通行碼不對"}
        _ensure_keys()
        items = data.get("items") or []
        if not isinstance(items, list) or len(items) > MAX_ITEMS:
            return 400, {"ok": False, "msg": f"items 要是陣列、最多 {MAX_ITEMS} 筆"}
        with ThreadPoolExecutor(max_workers=8) as ex:
            res = list(ex.map(_send_one, items))
        return 200, {"ok": True, "sent": sum(1 for r in res if r[0] == "ok"),
                     "gone": [r[1] for r in res if r[0] == "gone"],
                     "errors": [{"endpoint": r[1][:80], "msg": r[2]} for r in res if r[0] == "err"]}
    return 404, {"ok": False, "msg": "沒有這個網址"}
