#!/usr/bin/env python3
"""用無頭瀏覽器登入 104 企業大師、讀 12 店當月＋下月班表，送到門市排班工作台（再由工作台寫進排班分析儀表板）。
2026-11-09 經營者：「104 不開 API，你處理自動化的流程，不要麻煩我」→ 不再靠經營者在 Code 內建瀏覽器登入。

環境變數（GitHub Actions secrets 或本機 ~/.config/diybc/secrets.env）：
  HR104_EMAIL／HR104_PWD   104 會員信箱＋密碼（給程式專用的排班唯讀帳號；只讀取使用、不印出）
  DIYBC_HQ_PWD             工作台總部密碼（push.py 用）
  SCHED_DASH_TOKEN         排班分析儀表板寫入通行碼（可省，省了儀表板不會更新）
  SYNC104_URL              工作台網址（預設 https://diybc-training.onrender.com）
用法：python3 tools/sync104/fetch104.py [--out 結果檔] [--no-push] [--headed]
結束碼：0 成功；2 帳密不對（不重試，避免被鎖）；3 104 要額外驗證（簡訊／信箱／裝置）；4 讀班表失敗；5 送工作台失敗；6 瀏覽器或網路問題；7 還沒設定帳號（略過，結束碼 0）
只「讀」104：不按 104 上任何存檔／發佈鈕。
"""
import argparse
import asyncio
import json
import os
import re
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
SIGNIN = "https://pro.104.com.tw/user/signin?redirectURL=%2Fhrm2%2Fm9%2FF0901"
GRID = "https://pro.104.com.tw/hrm2/m9/F0901"
# 2026-11-09 實測：用 Playwright 預設的 HeadlessChrome 身分，bsignin.104.com.tw 不會出現登入表單；要用一般 Chrome 的 User-Agent
UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"


def load_secrets_env():
    """本機跑時從 secrets.env 補環境變數（已有的不覆蓋）"""
    p = os.path.expanduser("~/.config/diybc/secrets.env")
    if not os.path.exists(p):
        return
    for line in open(p, encoding="utf-8"):
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip())


def log(msg):
    print(time.strftime("%H:%M:%S"), msg, flush=True)


async def fetch(headed=False):
    from playwright.async_api import async_playwright
    email, pwd = os.environ.get("HR104_EMAIL", ""), os.environ.get("HR104_PWD", "")
    if not email or not pwd:
        return 7, "尚未設定 HR104_EMAIL／HR104_PWD（104 排班唯讀帳號），這次略過不登入", None
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=not headed)
        ctx = await browser.new_context(locale="zh-TW", timezone_id="Asia/Taipei", user_agent=UA)
        page = await ctx.new_page()
        try:
            await page.goto(SIGNIN, wait_until="domcontentloaded", timeout=60000)
            await page.wait_for_selector("input[name=email]", timeout=60000)
            await page.fill("input[name=email]", email)
            await page.fill("input[name=password]", pwd)
            await page.click("button:has-text('立即登入')")
            # 等跳回 pro.104.com.tw；還留在登入站就是失敗或要額外驗證
            deadline = time.time() + 60
            while time.time() < deadline:
                await page.wait_for_timeout(1000)
                url = page.url
                if url.startswith("https://pro.104.com.tw/") and "/user/" not in url:
                    break
            url = page.url
            if not url.startswith("https://pro.104.com.tw/") or "/user/" in url:
                txt = (await page.evaluate("() => document.body.innerText.replace(/\\s+/g,' ').slice(0,600)")) or ""
                if re.search(r"驗證碼|簡訊|OTP|裝置|二階段|驗證信|verify", txt + url, re.I):
                    return 3, "104 要求額外驗證（" + url + "）：" + txt[:200], None
                return 2, "104 登入失敗（" + url + "）：" + txt[:200], None
            await page.goto(GRID, wait_until="networkidle", timeout=60000)
            await page.wait_for_timeout(2000)
            if "/hrm2/" not in page.url:
                return 2, "登入後進不了排班頁（" + page.url + "）", None
            js = open(os.path.join(HERE, "extract.js"), encoding="utf-8").read()
            result = await page.evaluate(js)
            m = re.search(r"SYNC104_BEGIN(.*)SYNC104_END", str(result), re.S)
            if not m:
                return 4, "extract.js 沒有回傳班表資料：" + str(result)[:200], None
            payload = json.loads(m.group(1))
            if not payload.get("ok"):
                return 4, "104 讀班表失敗：" + str(payload.get("error")), None
            return 0, f"讀到 {len(payload.get('items') or [])} 份班表", payload
        except Exception as ex:  # noqa: BLE001
            return 6, f"瀏覽器／網路問題：{type(ex).__name__}: {str(ex)[:300]}", None
        finally:
            await browser.close()


async def probe():
    """自我檢查：無頭瀏覽器能不能開到 104 登入表單（不輸入任何東西）。回 (ok, 訊息)"""
    try:
        from playwright.async_api import async_playwright
    except ImportError as ex:
        return False, f"沒有 playwright：{ex}"
    try:
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True)
            page = await browser.new_page(locale="zh-TW", timezone_id="Asia/Taipei", user_agent=UA)   # 預設 HeadlessChrome 身分 104 不給表單
            await page.goto(SIGNIN, wait_until="domcontentloaded", timeout=60000)
            try:
                await page.wait_for_selector("input[name=email]", timeout=45000)
                has_form = True
            except Exception:  # noqa: BLE001
                has_form = False
            title = await page.title()
            url = page.url
            try:
                text = await page.evaluate("() => (document.body && document.body.innerText || '').replace(/\\s+/g,' ').trim().slice(0, 300)")
            except Exception:  # noqa: BLE001
                text = ""
            await browser.close()
            if has_form and "錯誤" not in title:
                return True, f"登入頁正常（{title}）｜{url}"
            return has_form, f"{'有表單但' if has_form else '沒有表單，'}頁面標題「{title}」｜{url}｜{text}"
    except Exception as ex:  # noqa: BLE001
        return False, f"{type(ex).__name__}: {str(ex)[:400]}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--probe", action="store_true", help="只檢查無頭瀏覽器能不能開 104 登入頁")
    ap.add_argument("--out", default=os.environ.get("SYNC104_OUT") or os.path.join(os.environ.get("RUNNER_TEMP") or "/tmp", "sync104_payload.txt"))
    ap.add_argument("--no-push", action="store_true")
    ap.add_argument("--headed", action="store_true")
    a = ap.parse_args()
    load_secrets_env()
    if a.probe:
        ok, msg = asyncio.run(probe())
        log(("✅ " if ok else "❌ ") + msg)
        sys.exit(0 if ok else 6)
    code, msg, payload = asyncio.run(fetch(a.headed))
    log(("✅ " if code == 0 else "⚠️ " if code == 7 else "❌ ") + msg)
    if code == 7:
        sys.exit(0)
    if code:
        sys.exit(code)
    with open(a.out, "w", encoding="utf-8") as f:
        f.write("SYNC104_BEGIN" + json.dumps(payload, ensure_ascii=False) + "SYNC104_END")
    for it in payload.get("items") or []:
        log(f"  {it['store']:>2} 號店 {it['ym']} {it.get('name', '')} {'錯誤：' + it['error'] if it.get('error') else ''}")
    if a.no_push:
        log("（--no-push）結果檔：" + a.out)
        return
    url = os.environ.get("SYNC104_URL", "https://diybc-training.onrender.com")
    r = subprocess.run([sys.executable, os.path.join(HERE, "push.py"), a.out, "--url", url])
    sys.exit(0 if r.returncode == 0 else 5)


if __name__ == "__main__":
    main()
