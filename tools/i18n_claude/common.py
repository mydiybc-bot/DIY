"""平板外語食譜：Claude Code 投遞口的共用小工具（2026-10-24）
通行碼只從本機 ~/.config/diybc/secrets.env 讀（I18N_PUT_KEY 投遞、I18N_PREVIEW_KEY 唯讀預覽），不印出、不寫進任何檔。"""
import json, os, re, subprocess

API = 'https://script.google.com/macros/s/AKfycbz7JhcgUxC0fLMTc2sk6EcN_XwNyLSdPiWE-Z0pP7IeumMZ_IaPVp7PM8sEekS-jtwm/exec'
WORK = os.path.expanduser('~/Desktop/Claude/Code/資料交換/輸出/暫存/i18n_claude')
os.makedirs(WORK, exist_ok=True)

def key(name):
    t = open(os.path.expanduser('~/.config/diybc/secrets.env')).read()
    m = re.search(r'^%s=(.*)$' % name, t, re.M)
    if not m or not m.group(1).strip():
        raise SystemExit('secrets.env 沒有 %s' % name)
    return m.group(1).strip()

def call(action, tries=4, timeout=170, **kw):
    """POST 到申請單 API；Google 偶爾回錯動作或逾時，自動重試"""
    body = dict(action=action); body.update(kw)
    data = json.dumps(body, ensure_ascii=False).encode('utf-8')
    last = None
    for _ in range(tries):
        r = subprocess.run(['curl', '-sL', '--max-time', str(timeout), '-H', 'Content-Type:', '--data-binary', '@-', API], input=data, capture_output=True)
        try:
            j = json.loads(r.stdout.decode('utf-8'))
            if j.get('act') == action:
                return j
            last = j
        except Exception:
            last = {'raw': r.stdout[:200].decode('utf-8', 'ignore')}
    return {'ok': False, 'msg': '連續 %d 次沒有拿到正確回覆' % tries, 'last': last}

def tag(id_, lang):
    return '%s_%s' % (id_[:8], lang)
