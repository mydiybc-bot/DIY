#!/usr/bin/env python3
"""把 extract.js 讀到的 104 班表送到門市排班工作台（2026-10-26）。

用法：python3 tools/sync104/push.py <javascript_tool 存下的結果檔> [--url https://diybc-training.onrender.com]
- 結果檔裡要有 'SYNC104_BEGIN' … 'SYNC104_END'（extract.js 的回傳值；內容大時工具會自動存成檔案）。
- 身分用總部密碼：讀 ~/.config/diybc/secrets.env 的 DIYBC_HQ_PWD（權限中心 task_hq 那組），不顯示在畫面上。
- 2026-11-09 起：工作台收到後會把 104 班表轉成「排班分析儀表板」的資料（fact_schedule）並重算；本程式帶 secrets.env 的
  SCHED_DASH_TOKEN（寫入通行碼）過去，然後等它跑完、印出每店寫入幾列。--no-dash 不等、--force-dash 內容沒變也重寫。
- 結束碼：0＝全部送達；1＝有錯（104 沒登入、密碼不對、連不上、儀表板寫入失敗）。
"""
import time
import argparse
import json
import os
import re
import ssl
import sys
import urllib.request


def ssl_ctx():
    """Mac 上 python.org 版 Python 沒帶憑證，連 https 會失敗 → 用 certifi，沒有就用系統的 /etc/ssl/cert.pem"""
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except ImportError:
        if os.path.exists("/etc/ssl/cert.pem"):
            return ssl.create_default_context(cafile="/etc/ssl/cert.pem")
        return ssl.create_default_context()


def read_payload(path):
    raw = open(path, encoding="utf-8").read()
    try:   # 工具存檔格式：[{type, text}]
        arr = json.loads(raw)
        if isinstance(arr, list):
            raw = "".join(x.get("text", "") for x in arr if isinstance(x, dict))
        elif isinstance(arr, str):
            raw = arr
    except json.JSONDecodeError:
        pass
    m = re.search(r"SYNC104_BEGIN(.*)SYNC104_END", raw, re.S)
    if not m:
        sys.exit("結果檔裡找不到 SYNC104_BEGIN／SYNC104_END")
    body = m.group(1)
    try:
        return json.loads(body)
    except json.JSONDecodeError:
        return json.loads(json.loads('"' + body + '"'))   # 被多包一層字串跳脫時


def hq_pwd():
    if os.environ.get("DIYBC_HQ_PWD"):   # GitHub Actions／伺服器用環境變數（2026-11-09）
        return os.environ["DIYBC_HQ_PWD"]
    p = os.path.expanduser("~/.config/diybc/secrets.env")
    if os.path.exists(p):
        for line in open(p, encoding="utf-8"):
            if line.startswith("DIYBC_HQ_PWD="):
                return line.split("=", 1)[1].strip()
    sys.exit("secrets.env 沒有 DIYBC_HQ_PWD（總部密碼），請經營者加一行")


def dash_token():
    """排班分析儀表板寫入通行碼（沒有就回 None，工作台會提醒）"""
    if os.environ.get("SCHED_DASH_TOKEN"):
        return os.environ["SCHED_DASH_TOKEN"]
    p = os.path.expanduser("~/.config/diybc/secrets.env")
    if not os.path.exists(p):
        return None
    for line in open(p, encoding="utf-8"):
        if line.startswith("SCHED_DASH_TOKEN="):
            return line.split("=", 1)[1].strip() or None
    return None


def call(url, act, body, timeout=600):
    req = urllib.request.Request(url.rstrip("/") + "/api/schedule/" + act, data=json.dumps(body).encode("utf-8"),
                                 method="POST", headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout, context=ssl_ctx() if url.startswith("https") else None) as r:
        return json.loads(r.read().decode("utf-8"))


def wait_dash(url, auth, t0, limit=540):
    """等工作台把這批送進排班分析儀表板；回傳 status dict（逾時回 None）"""
    end = time.time() + limit
    while time.time() < end:
        time.sleep(10)
        try:
            st = (call(url, "dash104_status", auth, timeout=60).get("status") or {})
        except (urllib.error.URLError, json.JSONDecodeError, OSError):
            continue
        if (st.get("started") or 0) >= t0 - 5 and st.get("finished"):
            return st
    return None


def print_dash(st):
    """印排班分析儀表板寫入結果；回傳 True＝有錯"""
    if st is None:
        print("⚠️ 排班分析儀表板：等了 9 分鐘還沒跑完，請稍後看工作台 dash104_status")
        return True
    if st.get("error"):
        print("❌ 排班分析儀表板寫入失敗：" + str(st["error"]))
        return True
    res = st.get("results") or {}
    for r in res.get("results") or []:
        line = f"📊 {r['store']} 號店 {r['ym']}：寫入 {r.get('written')} 列"
        if r.get("removedOld"):
            line += f"、清掉舊 NUEIP {r['removedOld']} 列"
        if r.get("unknownRole"):
            line += "、職稱對不到（當一般夥伴算）：" + "、".join(r["unknownRole"])
        if r.get("badSeg"):
            line += f"、看不懂的時段 {r['badSeg']} 段"
        print(line)
    if res.get("backups"):
        print("🗂 備份分頁：" + "、".join(res["backups"]))
    print(f"📊 fact_schedule 筆數 {res.get('before')} → {res.get('after')}；備註表寫入 {(res.get('notes') or {}).get('written')} 列")
    pr = st.get("process") or {}
    if pr.get("ok"):
        print(f"📊 重算完成（{pr.get('totalSeconds')} 秒）")
    else:
        print("⚠️ 重算沒完成：" + str(pr.get("error") or [x for x in pr.get("steps") or [] if not x.get("ok")]))
        return True
    return False


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--url", default="https://diybc-training.onrender.com")
    ap.add_argument("--pwd", help="只給本機測試用；正式一律讀 secrets.env")
    ap.add_argument("--no-dash", action="store_true", help="不等排班分析儀表板寫入結果")
    ap.add_argument("--force-dash", action="store_true", help="內容沒變也重寫排班分析儀表板")
    a = ap.parse_args()
    payload = read_payload(a.file)
    if not payload.get("ok"):
        print("❌ 104 讀取失敗：" + str(payload.get("error")))
        sys.exit(1)
    auth = {"as": "hq", "pwd": a.pwd or hq_pwd()}
    body = dict(auth, payload=payload, force_dash=a.force_dash)
    tok = dash_token()
    if tok:
        body["dash_token"] = tok
    t0 = time.time()
    try:
        res = call(a.url, "sync104", body)
    except urllib.error.HTTPError as e:
        print(f"❌ 工作台回 {e.code}：{e.read().decode('utf-8', 'replace')[:300]}")
        sys.exit(1)
    bad = 0
    for x in res.get("results") or []:
        if x.get("error"):
            bad += 1
            print(f"❌ {x['store']} 號店 {x['ym']}：{x['error']}")
            continue
        flag = "📥 已採用 104 版本" if x.get("adopted") else ("—" if not x.get("has_content") else "一致" if not x.get("diff") else f"工作台和 104 差 {x['diff']} 格")
        extra = (f"｜範本{x['template']}" if x.get("template") else "") + (f"｜名單對不上：{'、'.join(x['missing'])}" if x.get("missing") else "")
        print(f"✅ {x['store']} 號店 {x['ym']}（{x.get('status')}）：{flag}{extra}")
    print(f"共 {len(res.get('results') or [])} 份，{bad} 份有錯")
    dash = res.get("dash104") or {}
    if not a.no_dash:
        if not dash.get("has_token"):
            print("⚠️ 排班分析儀表板：工作台沒有寫入通行碼，這次沒寫（本機 secrets.env 要有 SCHED_DASH_TOKEN）")
            bad += 1
        elif not dash.get("queued"):
            print(f"📊 排班分析儀表板：{dash.get('skipped', 0)} 份內容跟上次一樣，不用重寫")
        else:
            print(f"📊 排班分析儀表板：{dash['queued']} 份有更動，等工作台寫入與重算…")
            if print_dash(wait_dash(a.url, auth, t0)):
                bad += 1
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
