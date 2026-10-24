"""
食譜小幫手獨立認證 + Anthropic API 代理
- 與 auth_store.py 完全隔離
- 密碼從環境變數 RECIPE_STUDIO_PASSWORD 讀取
- Anthropic key 從環境變數 ANTHROPIC_API_KEY 讀取
- Token 4 小時固定過期(非滑動)
- 多裝置可同時登入
"""
import json
import os
import secrets
import ssl
import time
import urllib.request
import urllib.error
from threading import Lock

RECIPE_SESSION_TTL = 4 * 60 * 60  # 4 小時(秒)
ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_VERSION = "2023-06-01"


class RecipeAuthStore:
    def __init__(self):
        self._tokens = {}  # token -> expires_at(epoch)
        self._lock = Lock()

    def login(self, password: str):
        expected = os.environ.get("RECIPE_STUDIO_PASSWORD", "")
        if not expected:
            return None  # 環境變數沒設,直接拒絕
        if password != expected:
            return None
        token = secrets.token_urlsafe(32)
        expires_at = time.time() + RECIPE_SESSION_TTL
        with self._lock:
            self._tokens[token] = expires_at
            self._gc()
        return {"token": token, "expires_at": int(expires_at)}

    def verify(self, token: str) -> bool:
        if not token:
            return False
        with self._lock:
            exp = self._tokens.get(token)
            if not exp:
                return False
            if time.time() > exp:
                self._tokens.pop(token, None)
                return False
            return True

    def logout(self, token: str):
        with self._lock:
            self._tokens.pop(token, None)

    def _gc(self):
        now = time.time()
        for t in [k for k, v in self._tokens.items() if v < now]:
            self._tokens.pop(t, None)


recipe_auth = RecipeAuthStore()


def call_anthropic(payload: dict) -> tuple[int, dict]:
    """
    轉發到 Anthropic Messages API
    回傳 (status_code, response_dict)
    """
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        return 500, {"error": "ANTHROPIC_API_KEY 未設定"}

    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        ANTHROPIC_API_URL,
        data=body,
        method="POST",
        headers={
            "Content-Type": "application/json",
            "x-api-key": api_key,
            "anthropic-version": ANTHROPIC_VERSION,
        },
    )
    ctx = ssl.create_default_context()
    try:
        with urllib.request.urlopen(req, timeout=120, context=ctx) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return resp.status, data
    except urllib.error.HTTPError as e:
        try:
            err_body = json.loads(e.read().decode("utf-8"))
        except Exception:
            err_body = {"error": str(e)}
        return e.code, err_body
    except Exception as e:
        return 500, {"error": f"Anthropic 連線失敗:{type(e).__name__}: {e}"}


# ============ 🌐 食譜翻譯（2026-10-19）：用「新品申請表的身分密碼」驗證，不用另外登入食譜小幫手 ============
# 經營者 10/19：「這個食譜翻譯不需要食譜小幫手的密碼輸入，簡單點」。
# 流程：前端送 role＋password＋payload → 先拿身分密碼到「食譜系統申請單 API」whoami 驗證（通過記 30 分鐘，存雜湊不存密碼）
#       → 通過才轉給 Anthropic；payload 只留翻譯用得到的欄位，模型限定、輸出上限 16000（不讓這個入口被拿去做別的事）。
import hashlib

REQ_API_URL = "https://script.google.com/macros/s/AKfycbz7JhcgUxC0fLMTc2sk6EcN_XwNyLSdPiWE-Z0pP7IeumMZ_IaPVp7PM8sEekS-jtwm/exec"
REQ_ID_TTL = 30 * 60
TRANSLATE_MODELS = ("claude-opus-5-5", "claude-sonnet-5-5", "claude-sonnet-4-5", "claude-haiku-5-5")   # haiku：Google 評論 AI 判讀（2026-10-24）
TRANSLATE_MAX_TOKENS = 16000
TRANSLATE_MAX_CHARS = 200000
_id_ok = {}
_id_lock = Lock()


def _whoami(role: str, password: str) -> dict:
    body = json.dumps({"action": "whoami", "role": role, "password": password}).encode("utf-8")
    req = urllib.request.Request(REQ_API_URL, data=body, method="POST",
                                 headers={"Content-Type": "text/plain;charset=utf-8"})
    ctx = ssl.create_default_context()
    with urllib.request.urlopen(req, timeout=30, context=ctx) as resp:   # Google 會 302 轉到結果頁，urllib 自動用 GET 跟過去
        return json.loads(resp.read().decode("utf-8"))


def verify_req_identity(role, password):
    """身分密碼對 → (True, '')；不對或連不到 → (False, 原因)"""
    role = str(role or "").strip()
    password = str(password or "")
    if not role or not password:
        return False, "請輸入身分密碼"
    key = hashlib.sha256((role + "\x00" + password).encode("utf-8")).hexdigest()
    now = time.time()
    with _id_lock:
        exp = _id_ok.get(key)
        if exp and exp > now:
            return True, ""
    last = "驗證身分時連不到申請單系統，請再試一次"
    for _ in range(3):   # Google 偶爾回錯動作或逾時 → 再試
        try:
            data = _whoami(role, password)
        except Exception:
            continue
        if data.get("act") != "whoami":
            continue
        if data.get("ok"):
            with _id_lock:
                _id_ok[key] = now + REQ_ID_TTL
                for k in [k for k, v in _id_ok.items() if v < now]:
                    _id_ok.pop(k, None)
            return True, ""
        if data.get("code") == "auth":
            return False, "身分密碼不對"
        last = str(data.get("msg") or "驗證身分失敗")
        break
    return False, last


# 🧾 零用金報銷（2026-10-24）：收據照片辨識也走這個入口 → content 可以是「文字＋圖片」區塊陣列（只收 base64 圖片，每則最多 4 張）
IMAGE_TYPES = ("image/jpeg", "image/png", "image/webp")
IMAGE_MAX_CHARS = 3_000_000   # 單張 base64 上限（約 2.2MB 圖檔）；前端已先縮到 1600px


def _clean_content(c):
    """回傳 (乾淨的 content, 文字字數)；不合格回 (None, 0)"""
    if isinstance(c, str):
        return c, len(c)
    if not isinstance(c, list):
        return None, 0
    out, n, imgs = [], 0, 0
    for b in c[:8]:
        if not isinstance(b, dict):
            continue
        if b.get("type") == "text" and isinstance(b.get("text"), str):
            out.append({"type": "text", "text": b["text"]})
            n += len(b["text"])
        elif b.get("type") == "image" and imgs < 4:
            s = b.get("source") if isinstance(b.get("source"), dict) else {}
            if s.get("type") == "base64" and s.get("media_type") in IMAGE_TYPES \
                    and isinstance(s.get("data"), str) and 0 < len(s["data"]) <= IMAGE_MAX_CHARS:
                out.append({"type": "image", "source": {"type": "base64", "media_type": s["media_type"], "data": s["data"]}})
                imgs += 1
    return (out or None), n


def clean_translate_payload(p: dict) -> dict:
    """只留翻譯用得到的欄位；模型、長度設上限"""
    p = p if isinstance(p, dict) else {}
    model = p.get("model") if p.get("model") in TRANSLATE_MODELS else TRANSLATE_MODELS[0]
    try:
        mt = int(p.get("max_tokens") or TRANSLATE_MAX_TOKENS)
    except (TypeError, ValueError):
        mt = TRANSLATE_MAX_TOKENS
    msgs, total = [], 0
    for m in (p.get("messages") or []):
        if not isinstance(m, dict) or m.get("role") not in ("user", "assistant"):
            continue
        content, n = _clean_content(m.get("content"))
        if content is None:
            continue
        msgs.append({"role": m["role"], "content": content})
        total += n
        if len(msgs) >= 4:
            break
    system = p.get("system") if isinstance(p.get("system"), str) else ""
    total += len(system)
    if total > TRANSLATE_MAX_CHARS:
        msgs = []   # 太長 → 送空的 messages，Anthropic 直接回 400（不花錢），前端顯示錯誤
    out = {"model": model, "max_tokens": max(256, min(mt, TRANSLATE_MAX_TOKENS)), "system": system,
           "messages": msgs}
    oc = p.get("output_config")
    if isinstance(oc, dict):
        keep = {}
        if oc.get("effort") in ("low", "medium", "high"):
            keep["effort"] = oc["effort"]
        if isinstance(oc.get("format"), dict) and oc["format"].get("type") == "json_schema":
            keep["format"] = {"type": "json_schema", "schema": oc["format"].get("schema") or {}}
        if keep:
            out["output_config"] = keep
    return out
