"""🍪 cookie_mold.py 的自動測試（模擬 AI，不連外、不花錢）。執行：python3 -m pytest test_cookie_mold.py -q"""
import base64, importlib, json, os, tempfile
from datetime import timedelta

import pytest


@pytest.fixture()
def cm(monkeypatch):
    d = tempfile.mkdtemp()
    monkeypatch.setenv("TRAINING_DATA_DIR", d)
    monkeypatch.setenv("COOKIE_MOLD_MOCK", "1")
    monkeypatch.setenv("COOKIE_MOLD_TEST_CODE", "t-code")
    monkeypatch.setenv("COOKIE_MOLD_ADMIN_TEST", "adm")
    monkeypatch.delenv("COOKIE_MOLD_MOCK_CHECK", raising=False)
    import cookie_mold
    m = importlib.reload(cookie_mold)
    m.COOLDOWN = 0          # 測試連續生成時不等冷卻
    return m


def call(cm, _act, **body):
    return cm.handle("/api/cookie-mold/" + _act, json.dumps(body).encode())


PHOTO = base64.b64encode(b"\xff\xd8fakejpeg").decode()


def login(cm, name="amy"):
    s, r = call(cm, "login", test_code="t-code", name=name)
    assert s == 200 and r["ok"]
    return r["token"]


def test_login_required_and_bad_code(cm):
    assert call(cm, "generate", photo=PHOTO)[0] == 401
    assert call(cm, "login", test_code="wrong")[0] == 401
    assert call(cm, "login", id_token="x")[1]["code"] == "auth"      # LINE 還沒設


def test_free_once_then_blocked_then_paid_unlimited(cm):
    t = login(cm)
    s, r = call(cm, "generate", token=t, photo=PHOTO, style="q")
    assert s == 200 and r["line"] and r["style"] == "q" and r["user"]["free_used"]
    s, r = call(cm, "generate", token=t, photo=PHOTO)
    assert s == 402 and r["code"] == "quota" and "訂金" in r["msg"]            # 驗收 3：第二次被擋並引導付訂金
    assert call(cm, "admin/paid", pwd="adm", uid="test:amy", on=True)[1]["paid"] is True
    for _ in range(3):                                                           # 驗收 4：已付訂金可重複生成
        assert call(cm, "generate", token=t, photo=PHOTO)[0] == 200
    assert call(cm, "admin/paid", pwd="adm", uid="test:amy", on=False)[1]["paid"] is False
    assert call(cm, "generate", token=t, photo=PHOTO)[0] == 402


def test_rights_block_does_not_use_quota(cm, monkeypatch):
    t = login(cm)
    monkeypatch.setenv("COOKIE_MOLD_MOCK_CHECK", '{"single_subject":true,"cartoon_or_brand":true,"celebrity":false,"clear":true}')
    s, r = call(cm, "generate", token=t, photo=PHOTO)
    assert s == 400 and r["code"] == "rights"
    monkeypatch.setenv("COOKIE_MOLD_MOCK_CHECK", '{"single_subject":false,"cartoon_or_brand":false,"celebrity":false,"clear":true}')
    assert call(cm, "generate", token=t, photo=PHOTO)[1]["code"] == "subject"
    monkeypatch.delenv("COOKIE_MOLD_MOCK_CHECK")
    assert call(cm, "generate", token=t, photo=PHOTO)[0] == 200                  # 被擋不扣免費次數


def test_confirm_review_and_download(cm):
    t = login(cm)
    did = call(cm, "generate", token=t, photo=PHOTO)[1]["design_id"]
    visit = (cm.print_cutoff_min() + timedelta(days=1)).isoformat()
    cu = {"name": "王小明", "phone": "0912-345-678", "store": "3店", "visit": visit}
    assert call(cm, "confirm", token=t, design_id=did, line=PHOTO, stl=PHOTO, customer=dict(cu, store="99店"))[1]["code"] == "form"
    assert call(cm, "confirm", token=t, design_id=did, line=PHOTO, stl=PHOTO, customer=dict(cu, visit="2020-01-01"))[1]["code"] == "form"
    assert call(cm, "confirm", token=t, design_id=did, line=PHOTO, stl=PHOTO, customer=cu)[0] == 200
    assert call(cm, "admin/list", pwd="nope")[0] in (401, 503)
    rows = call(cm, "admin/list", pwd="adm")[1]["designs"]
    assert len(rows) == 1 and rows[0]["customer"]["phone"] == "0912345678" and rows[0]["cutoff"]
    assert call(cm, "admin/review", pwd="adm", id=did, action="reject")[1]["code"] == "form"      # 退回要寫原因
    assert call(cm, "admin/review", pwd="adm", id=did, action="reject", reason="線條太細，請改選 Q 版")[1]["status"] == "rejected"
    mine = call(cm, "my", token=t)[1]["designs"]
    assert mine[0]["status"] == "rejected" and "Q 版" in mine[0]["reason"]                         # 客人看得到退回原因
    assert call(cm, "confirm", token=t, design_id=did, line=PHOTO, stl=PHOTO, customer=cu)[0] == 200  # 改完可再送
    assert call(cm, "admin/review", pwd="adm", id=did, action="approve")[1]["status"] == "approved"
    f = call(cm, "admin/file", pwd="adm", id=did, kind="stl")[1]
    assert base64.b64decode(f["data"]) == base64.b64decode(PHOTO) and f["name"].endswith(".stl")   # 驗收 5：下載 STL
    w = call(cm, "admin/weekly", pwd="adm")[1]["weeks"]
    assert w and w[0]["items"][0]["id"] == did


def test_budget_cap_pauses_ai(cm):
    t = login(cm)
    assert call(cm, "status", token=t)[1]["ai_open"] is True
    cm._log_cost("image", "x", {"promptTokenCount": 0, "candidatesTokenCount": 60000}, "test:amy")   # 模擬花掉 58.5 元
    assert cm.spent_today() >= 50
    st = call(cm, "status", token=t)[1]
    assert st["ai_open"] is False                                                                    # 驗收 6：超過上限自動暫停
    s, r = call(cm, "generate", token=t, photo=PHOTO)
    assert s == 429 and r["code"] == "closed" and "免費試玩" in r["msg"]
    stats = call(cm, "admin/stats", pwd="adm")[1]
    assert stats["budget_twd"] == 50 and stats["today_twd"] >= 50 and stats["days"][0]["logins"] == 1


def test_cutoff_rule(cm):
    from datetime import date
    assert cm.cutoff_for(date(2026, 12, 4)) == date(2026, 11, 30)     # 週五到店 → 同週週一截止
    assert cm.cutoff_for(date(2026, 12, 6)) == date(2026, 11, 30)     # 週日到店 → 同週週一
    assert cm.cutoff_for(date(2026, 12, 2)) == date(2026, 11, 23)     # 週三到店（還沒送到）→ 前一週週一


def test_busy_and_cooldown(cm):
    t = login(cm)
    for _ in range(cm.MAX_GEN):
        assert cm._GEN_SEM.acquire(blocking=False)                    # 模擬已有 3 個人正在生成
    s, r = call(cm, "generate", token=t, photo=PHOTO)
    assert s == 429 and r["code"] == "busy"
    for _ in range(cm.MAX_GEN):
        cm._GEN_SEM.release()
    assert call(cm, "generate", token=t, photo=PHOTO)[0] == 200       # 忙線被擋不扣次數、也不進冷卻
    assert call(cm, "admin/paid", pwd="adm", uid="test:amy", on=True)[0] == 200
    cm.COOLDOWN = 15
    assert call(cm, "generate", token=t, photo=PHOTO)[1]["code"] == "slow"   # 剛生成完馬上再按 → 請稍等


def test_spent_cache_survives_reload(cm):
    cm._log_cost("image", "x", {"promptTokenCount": 0, "candidatesTokenCount": 10000}, "u")
    a = cm.spent_today()
    cm._SPENT.clear()                                                  # 模擬伺服器重開
    assert a > 0 and cm.spent_today() == a


def test_admin_delete(cm):
    t = login(cm)
    did = call(cm, "generate", token=t, photo=PHOTO)[1]["design_id"]
    visit = (cm.print_cutoff_min() + timedelta(days=1)).isoformat()
    assert call(cm, "confirm", token=t, design_id=did, line=PHOTO, stl=PHOTO, customer={"name": "測試", "phone": "0900000000", "store": "1店", "visit": visit})[0] == 200
    assert call(cm, "admin/review", pwd="adm", id=did, action="approve")[0] == 200
    assert call(cm, "admin/delete", pwd="wrong", id=did)[0] in (401, 503)          # 沒密碼不能刪
    assert call(cm, "admin/delete", pwd="adm", id=did)[1]["deleted"] == did
    assert call(cm, "admin/list", pwd="adm", status="all")[1]["designs"] == []
    assert call(cm, "admin/weekly", pwd="adm")[1]["weeks"] == []                    # 待印清單也跟著消失
    assert call(cm, "admin/file", pwd="adm", id=did, kind="stl")[0] == 404          # 檔案也刪了
    assert call(cm, "my", token=t)[1]["designs"] == []
    assert call(cm, "admin/delete", pwd="adm", id=did)[0] == 404


def test_only_q_style_and_traits_in_prompt(cm):
    t = login(cm)
    assert call(cm, "generate", token=t, photo=PHOTO, style="real")[1]["style"] == "q"     # 2026-11-23 起只有 Q 版
    p = cm.draw_prompt("floppy ears, tongue out")
    assert "HEAD ONLY" in p and "floppy ears, tongue out" in p and "3D-printed" in p
    assert "traits" in cm.CHECK_PROMPT
