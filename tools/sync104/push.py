#!/usr/bin/env python3
"""把 extract.js 讀到的 104 班表送到門市排班工作台（2026-10-26）。

用法：python3 tools/sync104/push.py <javascript_tool 存下的結果檔> [--url https://diybc-training.onrender.com]
- 結果檔裡要有 'SYNC104_BEGIN' … 'SYNC104_END'（extract.js 的回傳值；內容大時工具會自動存成檔案）。
- 身分用總部密碼：讀 ~/.config/diybc/secrets.env 的 DIYBC_HQ_PWD（權限中心 task_hq 那組），不顯示在畫面上。
- 結束碼：0＝全部送達；1＝有錯（104 沒登入、密碼不對、連不上）。
"""
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
    p = os.path.expanduser("~/.config/diybc/secrets.env")
    for line in open(p, encoding="utf-8"):
        if line.startswith("DIYBC_HQ_PWD="):
            return line.split("=", 1)[1].strip()
    sys.exit("secrets.env 沒有 DIYBC_HQ_PWD（總部密碼），請經營者加一行")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file")
    ap.add_argument("--url", default="https://diybc-training.onrender.com")
    ap.add_argument("--pwd", help="只給本機測試用；正式一律讀 secrets.env")
    a = ap.parse_args()
    payload = read_payload(a.file)
    if not payload.get("ok"):
        print("❌ 104 讀取失敗：" + str(payload.get("error")))
        sys.exit(1)
    body = json.dumps({"as": "hq", "pwd": a.pwd or hq_pwd(), "payload": payload}).encode("utf-8")
    req = urllib.request.Request(a.url.rstrip("/") + "/api/schedule/sync104", data=body, method="POST",
                                 headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=600, context=ssl_ctx() if a.url.startswith("https") else None) as r:
            res = json.loads(r.read().decode("utf-8"))
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
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
