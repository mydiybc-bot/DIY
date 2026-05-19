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
