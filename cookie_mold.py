"""🍪 AI 專屬餅乾模・第二階段後端（2026-11-18；經營者決定照交辦書放訓練站伺服器）。

做什麼：LINE 登入換一次免費 AI 線稿（寫實版／Q 版）、已付訂金可無限次、客人確認設計送總部審核、
        總部審核後台（通過／退回／下載 STL／每週待印清單）、每次 AI 呼叫記成本、每日預算上限。
不做什麼：線上付款（另案；「已付訂金」由總部在後台手動開）。不用 Claude API（成本考量）。

網址（全部 POST、body＝JSON、回 {ok, ...}；不 ok 時 {ok:false, code, msg}，msg 是給客人看的白話）：
  客人：/api/cookie-mold/status   {token?}                 → AI 是否開放、登入狀態、額度
        /api/cookie-mold/login    {id_token} 或 {test_code,name} → {token, user}
        /api/cookie-mold/generate {token, photo, style}    → {design_id, line(PNG base64), left}
        /api/cookie-mold/confirm  {token, design_id, line, stl, customer} → 送總部審核
        /api/cookie-mold/my       {token}                  → 我的設計與審核結果（退回原因在這裡看）
  總部（每個都要帶 pwd，問權限中心 dashboard=cookiemold）：
        /api/cookie-mold/admin/list   {pwd, status?}       /admin/file {pwd, id, kind}
        /api/cookie-mold/admin/review {pwd, id, action, reason}   /admin/delete {pwd, id}（永久刪除）
        /api/cookie-mold/admin/paid   {pwd, uid, on}       /admin/users {pwd}
        /api/cookie-mold/admin/stats  {pwd, days?}         /admin/weekly {pwd}

金鑰一律讀環境變數（程式庫是公開的，絕不寫進程式）：
  GEMINI_API_KEY            Google Gemini
  LINE_LOGIN_CHANNEL_ID     LINE Login 頻道編號（還沒設 → LINE 登入關閉）
  COOKIE_MOLD_TEST_CODE     測試登入通行碼（有設才開放測試登入；正式開放前拿掉）
設定檔 data/cookie_mold_config.json：每日預算、模型名稱、單價、匯率（改檔即生效，不必改程式）。
資料放 TRAINING_DATA_DIR/cookie_mold/：users.json、sessions.json、designs/*.json、files/*、cost_log.jsonl。
照片保留：沒送審的 7 天後刪；已送審的保留到到店日後 30 天。
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
import re
import secrets
import ssl
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from storage_paths import atomic_write_json, data_dir

TZ = timezone(timedelta(hours=8))
AUTH_API = "https://script.google.com/macros/s/AKfycbyQ9LWY74Ix8VBe1gDoMdSF5eH74ratL7_f0EUUHi37IH3bXhkIJMx4WB2I-c-MLsWtLQ/exec"
AUTH_DASH = "cookiemold"
STYLES = {"q": "Q 版"}   # 2026-11-23 起只有 Q 版（經營者決定）
STORES = ["1店", "2店", "3店", "4店", "5店", "6店", "7店", "8店", "9店", "10店", "11店", "12店"]
MAX_PHOTO = 3_000_000      # 照片 base64 上限（頁面會先縮到最長邊 1024）
MAX_STL = 12_000_000   # base64 長度上限（約 9MB 的列印檔；大塊壓紋的檔比較大）
DEFAULT_CFG = {
    "daily_budget_twd": 50,
    "usd_twd": 32.5,
    "check_model": "gemini-3.1-flash-lite",
    "image_model": "gemini-3.1-flash-lite-image",   # 2026-11-18 實測：線條粗、乾淨、約 5 秒；2.5-flash-image 寫實版線太細不能用
    # 單價（美元／每百萬 token，2026-11-18 查官方定價頁）；image_out＝圖片輸出、image_text_out＝圖像模型順帶輸出的文字
    "price": {"check_in": 0.25, "check_out": 1.50, "image_in": 0.25, "image_out": 30.0, "image_text_out": 1.50},
}
_LOCK = threading.RLock()
_LAST_CLEAN = 0.0
# 多人同時使用（2026-11-18 經營者：一天可能一千人、同時多人在線）
MAX_GEN = 3                 # 同時最多幾個人在生成（其餘請他稍等，避免被 Gemini 以超過頻率擋下、也不拖慢訓練站）
COOLDOWN = 15               # 同一個人兩次生成至少隔幾秒（防狂按）
_GEN_SEM = threading.BoundedSemaphore(MAX_GEN)
_LAST_GEN: dict[str, float] = {}
_SPENT: dict[str, float] = {}   # 當日花費記在記憶體（{日期: 元}），不必每次重讀整份紀錄檔
_AI_DOWN = {"until": 0.0, "note": ""}   # Gemini 帳戶沒額度／金鑰失效時，先暫停 AI 生成 10 分鐘並在後台顯示原因


class Fail(Exception):
    def __init__(self, code: str, msg: str, status: int = 400):
        super().__init__(msg)
        self.code, self.msg, self.status = code, msg, status


def _ctx() -> ssl.SSLContext:
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:  # noqa: BLE001
        return ssl.create_default_context()


_CTX = _ctx()


# ---------------- 檔案 ----------------
def _dir() -> Path:
    d = data_dir() / "cookie_mold"
    (d / "designs").mkdir(parents=True, exist_ok=True)
    (d / "files").mkdir(parents=True, exist_ok=True)
    return d


def _load(name: str, default):
    try:
        return json.loads((_dir() / name).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return default


def _save(name: str, obj) -> None:
    atomic_write_json(_dir() / name, obj)


def cfg() -> dict:
    c = json.loads(json.dumps(DEFAULT_CFG))
    try:
        j = json.loads((Path(__file__).parent / "data" / "cookie_mold_config.json").read_text(encoding="utf-8"))
        for k, v in j.items():
            if k == "price" and isinstance(v, dict):
                c["price"].update(v)
            elif not k.startswith("_"):
                c[k] = v
    except (OSError, ValueError):
        pass
    return c


def _now() -> datetime:
    return datetime.now(TZ)


def _b64(s: str, limit: int) -> bytes:
    s = re.sub(r"^data:[^,]*,", "", str(s or ""))
    if not s or len(s) > limit:
        raise Fail("size", "檔案太大或是空的，請重新上傳一次。")
    try:
        return base64.b64decode(s, validate=False)
    except Exception:  # noqa: BLE001
        raise Fail("size", "檔案讀不出來，請重新上傳一次。")


# ---------------- 成本與每日預算 ----------------
def _log_cost(kind: str, model: str, usage: dict, uid: str) -> float:
    c, p = cfg(), cfg()["price"]
    tin, tout = int(usage.get("promptTokenCount") or 0), int(usage.get("candidatesTokenCount") or 0)
    img = sum(int(d.get("tokenCount") or 0) for d in usage.get("candidatesTokensDetails") or [] if d.get("modality") == "IMAGE")
    if kind == "image" and img:      # 圖片 token 與文字 token 單價不同，分開算
        usd = (tin * p["image_in"] + img * p["image_out"] + max(0, tout - img) * p.get("image_text_out", 0)) / 1e6
    else:
        usd = (tin * p[kind + "_in"] + tout * p[kind + "_out"]) / 1e6
    twd = round(usd * c["usd_twd"], 4)
    row = {"at": _now().isoformat(timespec="seconds"), "day": _now().strftime("%Y-%m-%d"), "kind": kind, "model": model,
           "in": tin, "out": tout, "twd": twd, "uid": uid}
    with _LOCK:
        spent_today()   # 先確定今天的累計已載入
        with open(_dir() / "cost_log.jsonl", "a", encoding="utf-8") as f:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")
        _SPENT[row["day"]] = _SPENT.get(row["day"], 0.0) + twd
    return twd


def _cost_rows() -> list[dict]:
    try:
        return [json.loads(x) for x in (_dir() / "cost_log.jsonl").read_text(encoding="utf-8").splitlines() if x.strip()]
    except (OSError, ValueError):
        return []


def spent_today() -> float:
    d = _now().strftime("%Y-%m-%d")
    with _LOCK:
        if d not in _SPENT:      # 每天第一次（或伺服器重開後）才讀檔
            _SPENT.clear()
            _SPENT[d] = sum(r["twd"] for r in _cost_rows() if r.get("day") == d)
        return round(_SPENT[d], 2)


def ai_open() -> bool:
    if _AI_DOWN["until"] > time.time():
        return False
    return bool(os.environ.get("GEMINI_API_KEY") or os.environ.get("COOKIE_MOLD_MOCK")) and spent_today() < float(cfg()["daily_budget_twd"])


# ---------------- Gemini ----------------
_MOCK_PNG = ("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=")


def _gemini(model: str, parts: list, want_image: bool) -> tuple[str, str, dict]:
    """回 (文字, 圖片 base64, 用量)。"""
    if os.environ.get("COOKIE_MOLD_MOCK"):   # 本機測試用：不連外、不花錢
        if want_image:
            mp = os.environ.get("COOKIE_MOLD_MOCK_IMG")
            return "", (base64.b64encode(Path(mp).read_bytes()).decode() if mp else _MOCK_PNG), {"promptTokenCount": 300, "candidatesTokenCount": 1290}
        return os.environ.get("COOKIE_MOLD_MOCK_CHECK") or '{"single_subject":true,"subject":"pet","cartoon_or_brand":false,"celebrity":false,"clear":true}', "", {"promptTokenCount": 300, "candidatesTokenCount": 40}
    key = os.environ.get("GEMINI_API_KEY", "")
    body: dict = {"contents": [{"role": "user", "parts": parts}]}
    body["generationConfig"] = {"responseModalities": ["TEXT", "IMAGE"]} if want_image else {"responseMimeType": "application/json", "temperature": 0}
    req = urllib.request.Request(
        f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent",
        data=json.dumps(body).encode("utf-8"), headers={"Content-Type": "application/json", "x-goog-api-key": key})
    try:
        with urllib.request.urlopen(req, timeout=110, context=_CTX) as r:
            j = json.loads(r.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        detail = e.read()[:300]
        if e.code in (401, 402, 403):     # 帳戶沒額度、金鑰失效：不是客人的問題，暫停 AI、讓總部在後台看到
            _AI_DOWN.update(until=time.time() + 600, note=f"Gemini 回覆 {e.code}：帳戶額度用完或金鑰失效，請到 Google AI Studio 檢查（{_now().strftime('%m/%d %H:%M')}）")
            raise Fail("closed", "AI 生成暫時休息中，請晚一點再來；免費試玩照常可以玩。", 503) from RuntimeError(detail)
        raise Fail("ai", "AI 現在有點忙，請過一下再試一次。", 502) from RuntimeError(detail)
    except Exception as e:  # noqa: BLE001
        raise Fail("ai", "AI 現在連不上，請過一下再試一次。", 502) from e
    text, img = "", ""
    for part in ((j.get("candidates") or [{}])[0].get("content") or {}).get("parts") or []:
        if part.get("text"):
            text += part["text"]
        d = part.get("inlineData") or part.get("inline_data")
        if d and d.get("data"):
            img = d["data"]
    return text, img, j.get("usageMetadata") or {}


CHECK_PROMPT = (
    "You screen customer photos for a custom cookie-stamp service. Look at the photo and answer in JSON only with keys: "
    "single_subject (true if exactly one clear main subject: one pet or one person), "
    "subject ('pet', 'person' or 'other'), "
    "cartoon_or_brand (true if the main subject is a cartoon/anime/game character, mascot, logo, branded artwork or product packaging), "
    "celebrity (true if the image looks like a press photo, poster, screenshot or publicity image of a public figure), "
    "clear (true if the subject is in focus and large enough to draw), "
    "traits (one sentence of at most 45 words listing the 4 or 5 most distinctive visible traits needed to recognise this individual in a very simple cartoon: "
    "expression, ear type and how they are carried or hairstyle and hair colour (dark or light), face or muzzle shape, glasses / hat / beard, "
    "and exactly which parts of the face or head are dark and which are light, and for a pet also the body build if the body is visible: "
    "body length, leg length, tail, and how it stands or sits), "
    "accessory (a short phrase naming the one most eye-catching thing a PET is wearing that its owner would recognise it by - for example 'heart-shaped tag "
    "with a small bell on the collar', 'bow tie', 'bandana', 'bow on the ear', 'sweater'; an empty string if it wears nothing special or only a plain collar, "
    "and always an empty string for a person), "
    "framing ('full_body' if the main subject is a PET whose whole body including legs or paws is visible in the photo; otherwise 'head'), "
    "facing ('left', 'right' or 'front': the side of the picture the subject's nose points to)."
)
# 2026-11-23 經營者：「全部用 Q 版、簡單線條、要掌握表情神韻、一眼看得出是誰、要考慮 3D 列印做得出來」。
# 反覆試了 7 版提示詞、26 張人物與動物照片後定案（原本一律只畫頭，全身版見下方 DRAW_BODY；先由照片檢查列出這個人／這隻寵物的特徵 traits，再帶進畫圖提示）：
#   DRAW_HEAD＝構圖（只畫頭、不畫脖子身體；帶身體的版本線條會變細、印不出來的比例從 5% 升到 25%）
#   DRAW_LIKE＝相似度（表情、耳朵或髮型、臉型、眼鏡鬍子帽子、深淺分界）
#   DRAW_PRINT＝3D 列印限制（純黑白、粗線、最多約 12 筆、深色毛髮才塗黑、灰白金橘不塗）
DRAW_HEAD = (
    'Draw a cute chibi (Q-version) icon of the main subject of this photo for a cookie stamp: THE HEAD ONLY, as one closed rounded shape seen from the '
    'front or three-quarter view. Nothing below the chin: no neck, no collar, no shoulders, no body, no paws, no scarf. Hair, ears, a hat or glasses belong'
    ' to the head and are included. The head fills about 85% of a square image, centred. Draw only the one main subject; ignore hands, other people or '
    'animals, objects and background. The outer outline must be one continuous unbroken line all the way round. Do not draw any circle, ring, frame, badge '
    'or sticker edge around the head, and no necklace or jewellery below the chin. '
)
# 2026-11-23 下午 經營者用臘腸狗全身照測試：「還是不太像」。只畫頭會丟掉最好認的體型（長身體、短腿）和照片的姿勢、方向，
# 所以寵物照片有拍到全身（照片檢查回 framing=full_body）就改畫全身、姿勢與朝向照照片；人物和只拍到頭的寵物照舊畫頭。
# 3 版提示詞、臘腸狗 11 次＋其他寵物 30 多張實測後定案（BODY3）。
DRAW_BODY = (
    'Draw a cute chibi (Q-version) icon of the pet in this photo for a cookie stamp: THE WHOLE ANIMAL, in exactly the pose of the photo. Keep the same '
    'facing direction as the photo (the nose points to the {FACING} of the picture - do not mirror it), the same stance (standing, sitting or lying) and the same '
    'turn and tilt of the head. Keep the real body build that makes this animal recognisable - body length, leg length, chest, tail, ear size - and gently '
    'exaggerate it. The head is drawn larger than life, about 45% of the drawing, with big eyes placed and spaced as in the photo; the body is simplified '
    'but must stay clearly recognisable. The outer outline is one single continuous unbroken closed line all the way round the whole animal - only one '
    'outline, never a second line, halo or sticker edge around it. Legs are simple chunky rounded shapes with no toes, no claws and no paw lines; show at '
    'most four legs, each clearly separated by a wide white gap or else merged into one shape - never thin slits between them. A collar, if there is one, '
    'is one simple curved band; no leash. {ACCESSORY}The animal fills about 90% of a square image, centred. Draw only the one main subject: no '
    'ground, grass, leaves, shadow, objects or background, and no circle, frame or badge around it. '
)
DRAW_BODY_PRINT = (
    'Striped, spotted, tabby, tortoiseshell or mottled coats get NO fill and no pattern marks at all - only a coat with large clearly separate solid '
    'black areas (such as a tuxedo cat or a black-and-white dog) gets those areas filled. Absolutely no colour anywhere: an orange, ginger or brown patch '
    'stays white. FINAL CHECK before you answer: the picture is black ink on white paper only - if any part is coloured, tinted or grey, redraw it in '
    'pure black and white. '
)
DRAW_LIKE = (
    'LIKENESS comes first. Before drawing, notice what makes THIS individual recognisable and keep it: the expression (smile, open mouth, tongue out, '
    'sleepy or closed eyes, wide eyes, head tilt), eye shape and spacing, face shape, and for animals the real ear type exactly as it is (upright, pointed,'
    ' bat-like, folded tip or hanging - never change it), muzzle shape and where dark and light fur meet on the face; for people the real hairstyle '
    '(length, parting, fringe, curls, bun, ponytail, bald), glasses, hat, beard or moustache. Gently exaggerate the three most distinctive traits like a '
    'friendly caricature. Someone who knows the subject must recognise them at a glance. Keep the subject looking their real age and gender. '
)
DRAW_PRINT = (
    'PRINTING LIMITS (this will be 3D-printed as a 6 cm cookie stamp, so obey strictly): pure white background; ONLY pure black and pure white - no gray, '
    'no colours, no shading, no dots, no hatching, no texture; draw with ONE very thick round marker: every line the same width, about 40 pixels on a '
    '1024-pixel image (never thinner than 30); leave white gaps at least as wide as a line between any two lines; at most about 12 strokes in total; no '
    'thin strokes, no individual hairs, whiskers, eyelashes, teeth, wrinkles, fur texture or clothing patterns; no text, no border, no background, no '
    'shadow. Eyes are solid black dots or ovals at least 50 pixels across (or one thick arc each if closed or smiling). For an animal the nose is one solid'
    ' black shape; for a PERSON the nose is only one small curved line, never a black blob, and the mouth is one simple line or a simple open smile. Solid '
    'black fill is allowed only for areas that are really BLACK or very dark brown in the photo: black fur patches, black or dark-brown hair, a dark beard '
    'or moustache, dark sunglasses, a dark hat. Never fill light, grey, white, blond, ginger, tan or red areas: grey or white hair, blond hair, tan or '
    'ginger fur and rosy cheeks all stay white with just an outline. Fill those completely with flat black following their real shape, and keep a clear '
    'white gap (at least a line width) between a black area and any eye, nose or other line inside or beside it. Light fur, light hair, skin and everything'
    ' else stays white with an outline. Fluffy or curly outlines become a few big rounded bumps. If the animal is ginger, cream, tan, grey or white, use no'
    ' black fill at all except eyes and nose. '
)


# 2026-11-23 下午 經營者：「狗胸前的鈴鐺，這類型的裝飾物，具有強烈的辨識度，應該要保留」。照片檢查回 accessory（寵物身上最顯眼的飾品），畫圖時保留成一個印得出來的大圖形。
DRAW_ACC_BODY = (
    'This pet wears: {ACC}. KEEP IT - the owner recognises the pet by it. Draw it in its real place as ONE bold simple solid black shape in its own real '
    'form (do not turn it into something else), drawn BIG like a cartoon medal so it survives printing: roughly one fifth of the height of the '
    'drawing (about 180 pixels). It hangs directly from the middle of the collar on the chest, its top touching the collar - no ring, chain, string or loop '
    'between them - and it sits on open white chest with a wide white gap from the body outline, the legs and every other line. No small details, '
    'holes or highlights inside it, no lettering. It and the collar are pure black and white like everything else - never red or any colour. '
)
DRAW_ACC_HEAD = (
    'Exception to "nothing below the chin": this pet wears: {ACC}. KEEP IT - the owner recognises the pet by it. Draw it touching the bottom of the head '
    'outline (or on the ear or head if that is where it is worn) as ONE bold simple solid black shape in its own real form (do not turn it into something '
    'else), drawn BIG like a cartoon medal so it survives printing: about 180 pixels across, joined directly to the head - no ring, chain or string - '
    'so the whole drawing is still one closed shape. No small details, holes or highlights inside it, no lettering. Pure black and white only - never '
    'red or any colour. '
)
DRAW_ACC_NONE = 'No tag, heart, bell or other ornament. '


def accessory_of(chk: dict) -> str:
    if chk.get("subject") != "pet":
        return ""
    a = re.sub(r"[^A-Za-z0-9 ,'()-]+", " ", str(chk.get("accessory") or "")).strip()[:80]
    low = a.lower()
    if low in ("none", "no", "n a", "nothing") or re.fullmatch(r"(?:an? |the )?(?:[a-z-]+ ){0,3}(?:collar|harness|leash)", low):
        return ""          # 只有素面項圈、胸背帶不算飾品（實測：米格魯只有紫色項圈，卻被畫出一顆愛心）
    return a


def layout_of(chk: dict) -> str:
    return "body" if chk.get("subject") == "pet" and chk.get("framing") == "full_body" else "head"


def draw_prompt(traits: str, chk: dict = None) -> str:
    chk = chk if isinstance(chk, dict) else {}
    t = re.sub(r"[\r\n]+", " ", str(traits or ""))[:500]
    acc = accessory_of(chk)
    tr = f"The traits of this individual that MUST be visible in the drawing: {t} " if t else ""
    if layout_of(chk) == "body":
        facing = chk.get("facing") if chk.get("facing") in ("left", "right", "front") else "same side"
        return (DRAW_BODY.replace("{FACING}", facing).replace("{ACCESSORY}", DRAW_ACC_BODY.replace("{ACC}", acc) if acc else DRAW_ACC_NONE)
                + DRAW_LIKE + tr + DRAW_PRINT + DRAW_BODY_PRINT)
    return DRAW_HEAD + (DRAW_ACC_HEAD.replace("{ACC}", acc) if acc else "") + DRAW_LIKE + tr + DRAW_PRINT


# ---------------- 登入 ----------------
def _users() -> dict:
    return _load("users.json", {})


def _login(b: dict) -> dict:
    uid = name = ""
    test = os.environ.get("COOKIE_MOLD_TEST_CODE", "")
    if b.get("test_code"):
        if not test or not secrets.compare_digest(str(b["test_code"]), test):
            raise Fail("auth", "測試通行碼不正確。", 401)
        name = re.sub(r"[^\w一-鿿-]", "", str(b.get("name") or "tester"))[:20] or "tester"
        uid = "test:" + name
    else:
        cid = os.environ.get("LINE_LOGIN_CHANNEL_ID", "")
        if not cid or not b.get("id_token"):
            raise Fail("auth", "LINE 登入還沒開放，敬請期待。", 401)
        data = urllib.parse.urlencode({"id_token": str(b["id_token"]), "client_id": cid}).encode()
        try:
            with urllib.request.urlopen(urllib.request.Request("https://api.line.me/oauth2/v2.1/verify", data=data), timeout=20, context=_CTX) as r:
                j = json.loads(r.read().decode("utf-8"))
        except Exception as e:  # noqa: BLE001
            raise Fail("auth", "LINE 登入沒有成功，請再試一次。", 401) from e
        if not j.get("sub"):
            raise Fail("auth", "LINE 登入沒有成功，請再試一次。", 401)
        uid, name = "line:" + j["sub"], str(j.get("name") or "")[:40]
    token = secrets.token_urlsafe(32)
    with _LOCK:
        users = _users()
        u = users.setdefault(uid, {"name": name, "created": _now().isoformat(timespec="seconds"), "free_used": False, "paid": False, "gen": 0})
        if name:
            u["name"] = name
        _save("users.json", users)
        ses = {k: v for k, v in _load("sessions.json", {}).items() if v.get("exp", 0) > time.time()}
        ses[hashlib.sha256(token.encode()).hexdigest()] = {"uid": uid, "exp": time.time() + 30 * 86400}
        _save("sessions.json", ses)
    return {"token": token, "user": _pub_user(uid, u)}


def _pub_user(uid: str, u: dict) -> dict:
    return {"name": u.get("name", ""), "paid": bool(u.get("paid")), "free_used": bool(u.get("free_used")),
            "can_generate": bool(u.get("paid")) or not u.get("free_used"), "test": uid.startswith("test:")}


def _who(b: dict, required: bool = True) -> tuple[str, dict]:
    t = str(b.get("token") or "")
    s = _load("sessions.json", {}).get(hashlib.sha256(t.encode()).hexdigest()) if t else None
    if not s or s.get("exp", 0) < time.time():
        if required:
            raise Fail("login", "請先用 LINE 登入。", 401)
        return "", {}
    return s["uid"], _users().get(s["uid"], {})


# ---------------- 客人：狀態、生成、確認 ----------------
def _status(b: dict) -> dict:
    uid, u = _who(b, required=False)
    return {"ai_open": ai_open(), "line_login": bool(os.environ.get("LINE_LOGIN_CHANNEL_ID")),
            "test_login": bool(os.environ.get("COOKIE_MOLD_TEST_CODE")), "user": _pub_user(uid, u) if uid else None,
            "styles": STYLES, "stores": STORES}


def _design_path(did: str) -> Path:
    if not re.fullmatch(r"[0-9a-f]{16}", did or ""):
        raise Fail("id", "找不到這個設計。", 404)
    return _dir() / "designs" / f"{did}.json"


def _generate(b: dict) -> dict:
    uid, u = _who(b)
    style = "q"
    if not ai_open():
        raise Fail("closed", "今天的 AI 生成名額已經用完了，明天再來；免費試玩照常可以玩。", 429)
    if u.get("free_used") and not u.get("paid"):
        raise Fail("quota", "免費的 AI 生成已經用過了。付訂金預約後，就可以無限次重新生成、慢慢調整到滿意。", 402)
    photo = _b64(b.get("photo"), MAX_PHOTO)
    with _LOCK:
        if time.time() - _LAST_GEN.get(uid, 0) < COOLDOWN:
            raise Fail("slow", "剛剛才生成過，請等幾秒再按一次。", 429)
        _LAST_GEN[uid] = time.time()
    if not _GEN_SEM.acquire(blocking=False):
        _LAST_GEN.pop(uid, None)
        raise Fail("busy", "現在同時很多人在生成，請等 30 秒左右再按一次（不會扣掉你的次數）。", 429)
    try:
        return _generate_locked(uid, u, style, photo)
    finally:
        _GEN_SEM.release()


def _generate_locked(uid: str, u: dict, style: str, photo: bytes) -> dict:
    c = cfg()
    img_part = {"inline_data": {"mime_type": "image/jpeg", "data": base64.b64encode(photo).decode()}}
    # 1. AI 照片檢查（輕量模型）
    text, _, usage = _gemini(c["check_model"], [{"text": CHECK_PROMPT}, img_part], False)
    _log_cost("check", c["check_model"], usage, uid)
    try:
        chk = json.loads(re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.M))
    except ValueError:
        chk = {}
    if isinstance(chk, list):      # 偶爾會回一個陣列（照片裡不只一個主角時）
        chk = chk[0] if len(chk) == 1 and isinstance(chk[0], dict) else {"single_subject": False}
    if not isinstance(chk, dict):
        chk = {}
    if chk.get("cartoon_or_brand") or chk.get("celebrity"):
        raise Fail("rights", "這張看起來是卡通角色、品牌圖案或名人照片，有版權或肖像權的疑慮，沒辦法做成餅乾模。請換一張自己拍的寵物或親友照片。")
    if chk.get("single_subject") is False:
        raise Fail("subject", "這張照片裡的主角不只一個，或是看不出主角。請換一張只有一隻寵物或一個人、主角清楚的照片。")
    if chk.get("clear") is False:
        raise Fail("blurry", "這張照片的主角不夠清楚，畫出來會不像。請換一張對焦清楚、主角大一點的照片。")
    # 2. 線稿生成（圖像模型）
    _, line, usage = _gemini(c["image_model"], [{"text": draw_prompt(chk.get("traits"), chk)}, img_part], True)
    twd = _log_cost("image", c["image_model"], usage, uid)
    if not line:
        raise Fail("ai", "這次沒有畫成功，請再試一次（不會扣掉你的免費次數）。", 502)
    did = secrets.token_hex(8)
    (_dir() / "files" / f"{did}_photo.jpg").write_bytes(photo)
    (_dir() / "files" / f"{did}_raw.png").write_bytes(base64.b64decode(line))
    with _LOCK:
        users = _users()
        uu = users.setdefault(uid, u)
        uu["free_used"], uu["gen"] = True, int(uu.get("gen") or 0) + 1
        _save("users.json", users)
        atomic_write_json(_design_path(did), {"id": did, "uid": uid, "name": uu.get("name", ""), "style": style, "layout": layout_of(chk), "status": "draft",
                                               "created": _now().isoformat(timespec="seconds"), "check": chk, "twd": twd})
    return {"design_id": did, "line": line, "style": style, "layout": layout_of(chk), "user": _pub_user(uid, uu)}


def _confirm(b: dict) -> dict:
    uid, _u = _who(b)
    p = _design_path(str(b.get("design_id") or ""))
    cu = b.get("customer") or {}
    name, phone = str(cu.get("name") or "").strip()[:30], re.sub(r"\D", "", str(cu.get("phone") or ""))[:12]
    store, visit = str(cu.get("store") or ""), str(cu.get("visit") or "")
    if not name or len(phone) < 8:
        raise Fail("form", "請填寫姓名和手機號碼，總部審核有問題時才聯絡得到你。")
    if store not in STORES:
        raise Fail("form", "請選擇要去的門市。")
    try:
        vd = date.fromisoformat(visit)
    except ValueError:
        raise Fail("form", "請選擇預計到店的日期。")
    if vd < print_cutoff_min():
        raise Fail("form", f"模具要先列印再送到門市，最快可以選 {print_cutoff_min().isoformat()} 之後的日期。")
    with _LOCK:
        try:
            d = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            raise Fail("id", "找不到這個設計，請重新生成一次。", 404)
        if d.get("uid") != uid:
            raise Fail("id", "找不到這個設計。", 404)
        if d.get("status") in ("submitted", "approved"):
            raise Fail("dup", "這個設計已經送出了，請等總部審核。")
        (_dir() / "files" / f"{d['id']}_line.png").write_bytes(_b64(b.get("line"), MAX_PHOTO))
        (_dir() / "files" / f"{d['id']}.stl").write_bytes(_b64(b.get("stl"), MAX_STL))
        d.update({"status": "submitted", "submitted": _now().isoformat(timespec="seconds"), "reason": "",
                  "customer": {"name": name, "phone": phone, "store": store, "visit": visit}})
        atomic_write_json(p, d)
    return {"status": "submitted"}


def _designs() -> list[dict]:
    out = []
    for p in (_dir() / "designs").glob("*.json"):
        try:
            out.append(json.loads(p.read_text(encoding="utf-8")))
        except (OSError, ValueError):
            continue
    return sorted(out, key=lambda d: d.get("submitted") or d.get("created") or "", reverse=True)


def _my(b: dict) -> dict:
    uid, u = _who(b)
    keep = ("id", "style", "status", "created", "submitted", "reason", "customer")
    return {"user": _pub_user(uid, u), "designs": [{k: d.get(k) for k in keep} for d in _designs() if d.get("uid") == uid and d.get("status") != "draft"]}


# ---------------- 列印排程：到店日前一個週一中午截止、週二列印、週三隨貨出貨 ----------------
def cutoff_for(visit: date) -> date:
    """這個到店日要趕上哪一個週一截止。週三出貨要留運送時間 → 到店日至少在截止週一的 4 天後（週五）。"""
    mon = visit - timedelta(days=visit.weekday())      # 到店那一週的週一
    return mon if (visit - mon).days >= 4 else mon - timedelta(days=7)


def print_cutoff_min() -> date:
    """現在送出，最快可以選哪一天到店。"""
    now = _now()
    mon = now.date() - timedelta(days=now.weekday())
    if now.weekday() > 0 or now.hour >= 12:            # 已過本週一中午 → 下週一
        mon += timedelta(days=7)
    return mon + timedelta(days=4)


# ---------------- 總部 ----------------
_PW_OK: dict[str, float] = {}


def _admin(b: dict) -> None:
    pw = str(b.get("pwd") or "").strip()
    h = hashlib.sha256(pw.encode()).hexdigest()
    if pw and _PW_OK.get(h, 0) > time.time():
        return
    local = os.environ.get("COOKIE_MOLD_ADMIN_TEST", "")      # 只給本機測試用
    if local and pw and secrets.compare_digest(pw, local):
        return
    if not pw:
        raise Fail("auth", "請輸入密碼。", 401)
    qs = urllib.parse.urlencode({"action": "verify", "dashboard": AUTH_DASH, "pwd": pw, "callback": "cb"})
    try:
        with urllib.request.urlopen(AUTH_API + "?" + qs, timeout=25, context=_CTX) as r:
            t = r.read().decode("utf-8", "replace")
        m = re.match(r"^\s*[A-Za-z_$][\w$]*\(([\s\S]*)\);?\s*$", t)
        j = json.loads(m.group(1) if m else t)
    except Exception as e:  # noqa: BLE001
        raise Fail("auth", "連不上權限中心，請再試一次。", 503) from e
    if not j.get("ok"):
        raise Fail("auth", str(j.get("error") or "密碼錯誤"), 401)
    _PW_OK[h] = time.time() + 4 * 3600


def _admin_list(b: dict) -> dict:
    want = str(b.get("status") or "submitted")
    rows = [d for d in _designs() if d.get("status") != "draft" and (want == "all" or d.get("status") == want)]
    for d in rows:
        try:
            d["cutoff"] = cutoff_for(date.fromisoformat(d["customer"]["visit"])).isoformat()
        except (KeyError, ValueError):
            d["cutoff"] = ""
    return {"designs": rows}


def _admin_file(b: dict) -> dict:
    did, kind = str(b.get("id") or ""), str(b.get("kind") or "")
    _design_path(did)
    name = {"photo": f"{did}_photo.jpg", "raw": f"{did}_raw.png", "line": f"{did}_line.png", "stl": f"{did}.stl"}.get(kind)
    if not name:
        raise Fail("id", "沒有這種檔案。", 404)
    try:
        return {"name": name, "data": base64.b64encode((_dir() / "files" / name).read_bytes()).decode()}
    except OSError:
        raise Fail("id", "檔案已經不在了（可能超過保留期限）。", 404)


def _admin_review(b: dict) -> dict:
    p, act, reason = _design_path(str(b.get("id") or "")), b.get("action"), str(b.get("reason") or "").strip()[:300]
    if act not in ("approve", "reject"):
        raise Fail("form", "請選通過或退回。")
    if act == "reject" and not reason:
        raise Fail("form", "退回要寫原因，客人才知道怎麼改。")
    with _LOCK:
        d = json.loads(p.read_text(encoding="utf-8"))
        d.update({"status": "approved" if act == "approve" else "rejected", "reason": reason if act == "reject" else "",
                  "reviewed": _now().isoformat(timespec="seconds")})
        atomic_write_json(p, d)
    return {"status": d["status"]}


def _admin_delete(b: dict) -> dict:
    """永久刪除一個設計（資料與照片、線稿、列印檔一起刪，不能復原）。測試資料、客人取消時用。"""
    p = _design_path(str(b.get("id") or ""))
    with _LOCK:
        try:
            d = json.loads(p.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            raise Fail("id", "找不到這個設計（可能已經刪掉了）。", 404)
        for f in (_dir() / "files").glob(d["id"] + "*"):
            f.unlink(missing_ok=True)
        p.unlink(missing_ok=True)
    print("[cookie_mold] deleted", d["id"], d.get("status"), flush=True)
    return {"deleted": d["id"]}


def _admin_paid(b: dict) -> dict:
    with _LOCK:
        users = _users()
        u = users.get(str(b.get("uid") or ""))
        if not u:
            raise Fail("id", "找不到這位客人。", 404)
        u["paid"] = bool(b.get("on"))
        _save("users.json", users)
    return {"paid": u["paid"]}


def _admin_users(_b: dict) -> dict:
    return {"users": [dict(uid=k, **v) for k, v in sorted(_users().items(), key=lambda kv: kv[1].get("created", ""), reverse=True)]}


def _admin_stats(b: dict) -> dict:
    days = max(1, min(90, int(b.get("days") or 14)))
    since = (_now().date() - timedelta(days=days - 1)).isoformat()
    by: dict[str, dict] = {}
    def slot(day: str) -> dict:
        return by.setdefault(day, {"day": day, "twd": 0.0, "gen": 0, "check": 0, "logins": 0, "confirmed": 0})
    for r in _cost_rows():
        if r.get("day", "") >= since:
            s = slot(r["day"]); s["twd"] = round(s["twd"] + r["twd"], 2); s["gen" if r["kind"] == "image" else "check"] += 1
    for u in _users().values():
        if u.get("created", "")[:10] >= since:
            slot(u["created"][:10])["logins"] += 1
    for d in _designs():
        if (d.get("submitted") or "")[:10] >= since:
            slot(d["submitted"][:10])["confirmed"] += 1
    c = cfg()
    return {"days": sorted(by.values(), key=lambda s: s["day"], reverse=True), "today_twd": spent_today(),
            "budget_twd": c["daily_budget_twd"], "ai_open": ai_open(),
            "ai_note": _AI_DOWN["note"] if _AI_DOWN["until"] > time.time() else ("沒有設定 GEMINI_API_KEY" if not (os.environ.get("GEMINI_API_KEY") or os.environ.get("COOKIE_MOLD_MOCK")) else ""), "models": {"check": c["check_model"], "image": c["image_model"]}}


def _admin_weekly(_b: dict) -> dict:
    weeks: dict[str, list] = {}
    for d in _designs():
        if d.get("status") != "approved":
            continue
        try:
            cut = cutoff_for(date.fromisoformat(d["customer"]["visit"]))
        except (KeyError, ValueError):
            continue
        if cut >= _now().date() - timedelta(days=7):
            weeks.setdefault(cut.isoformat(), []).append({"id": d["id"], "style": d["style"], "customer": d["customer"]})
    return {"weeks": [{"cutoff": k, "print": (date.fromisoformat(k) + timedelta(days=1)).isoformat(),
                       "ship": (date.fromisoformat(k) + timedelta(days=2)).isoformat(), "items": v} for k, v in sorted(weeks.items())]}


# ---------------- 照片保留期限 ----------------
def _cleanup() -> None:
    global _LAST_CLEAN
    if time.time() - _LAST_CLEAN < 3600:
        return
    _LAST_CLEAN = time.time()
    today = _now().date()
    for d in _designs():
        try:
            if d.get("status") in ("draft", "rejected"):
                gone = (today - date.fromisoformat((d.get("reviewed") or d["created"])[:10])).days > 7
            else:
                gone = (today - date.fromisoformat(d["customer"]["visit"])).days > 30
        except (KeyError, ValueError):
            gone = False
        if gone:
            for f in (_dir() / "files").glob(d["id"] + "*"):
                f.unlink(missing_ok=True)
            (_dir() / "designs" / f"{d['id']}.json").unlink(missing_ok=True)


ROUTES = {"status": _status, "login": _login, "generate": _generate, "confirm": _confirm, "my": _my}
ADMIN = {"list": _admin_list, "file": _admin_file, "review": _admin_review, "delete": _admin_delete, "paid": _admin_paid,
         "users": _admin_users, "stats": _admin_stats, "weekly": _admin_weekly}


def handle(path: str, raw: bytes) -> tuple[int, dict]:
    try:
        body = json.loads(raw.decode("utf-8") or "{}") if raw else {}
        if not isinstance(body, dict):
            raise ValueError
    except (ValueError, UnicodeDecodeError):
        return 400, {"ok": False, "code": "json", "msg": "資料格式不正確。"}
    name = path.split("/api/cookie-mold/", 1)[-1].strip("/")
    try:
        _cleanup()
        if name.startswith("admin/"):
            fn = ADMIN.get(name[6:])
            if not fn:
                return 404, {"ok": False, "code": "act", "msg": "沒有這個功能。"}
            _admin(body)
            return 200, dict(ok=True, **fn(body))
        fn = ROUTES.get(name)
        if not fn:
            return 404, {"ok": False, "code": "act", "msg": "沒有這個功能。"}
        return 200, dict(ok=True, **fn(body))
    except Fail as e:
        return e.status, {"ok": False, "code": e.code, "msg": e.msg}
    except Exception as e:  # noqa: BLE001
        print("[cookie_mold]", name, repr(e)[:300], flush=True)
        return 500, {"ok": False, "code": "err", "msg": "系統出了一點問題，請再試一次。"}
