#!/usr/bin/env python3
"""熊主廚開場動畫 → 各語言版：在原本英文的位置換成該語言（逐字打出效果）
用法：python3 build.py <語言 ja|ko|vi|id|th> [只輸出幾秒到幾秒的預覽 s e]"""
import json, subprocess, sys, os, math
import numpy as np
from PIL import Image, ImageDraw, ImageFilter
HERE = os.path.dirname(os.path.abspath(__file__))
FF = os.path.expanduser('~/.local/bin/ffmpeg')
SRC = HERE + '/bear.mp4'
W, H = 854, 480
FPS = 30000 / 1001
lang = sys.argv[1]
prev = (float(sys.argv[2]), float(sys.argv[3])) if len(sys.argv) > 3 else None
FONT = {'ja': ('HiraMaruProN-W4', 18.5), 'ko': ('AppleSDGothicNeo-Medium', 19), 'th': ('SukhumvitSet-Medium', 20), 'vi': ('rounded', 19), 'id': ('rounded', 19)}
FONTB = {'ja': 'HiraMaruProN-W4', 'ko': 'AppleSDGothicNeo-Bold', 'th': 'SukhumvitSet-Bold', 'vi': 'rounded', 'id': 'rounded'}
BROWN, RED, BLACK, PILL = [92, 70, 44], [205, 8, 30], [20, 12, 10], [64, 33, 15]
PEG_BREAK = {'ja': '有孔ボード\n壁エリア', 'ko': '타공판\n구역', 'vi': 'Khu tường\ntreo đồ', 'id': 'Area dinding\npegboard', 'th': 'โซนผนัง\nเพกบอร์ด'}
segs = json.load(open(HERE + '/lines.json'))['seg']
ds = sorted(x['s'] for x in segs if x['k'] == 'd')
for x in segs:   # 對話框淡出那幾格也要蓋：結束時間往後延 0.4 秒，但不壓到下一段
    if x['k'] == 'd': x['e'] = min(x['e'] + 0.4, min([v for v in ds if v > x['s']] or [999]))
    else: x['e0'] = x['e']; x['e'] += 0.5
fn, fs = FONT[lang]

# ---- 1. 文字圖（Mac 原生排版）----
jobs = []
for i, s in enumerate(segs):
    tx = s[lang]
    if s['k'] == 'd': jobs.append(dict(id='s%d' % i, lines=tx.split('\n'), font=fn, size=fs, minSize=fs - 4, maxW=615, color=BROWN, red=RED, bold=False))
    elif s['k'] == 'q': jobs.append(dict(id='s%d' % i, lines=[tx], font=fn, size=fs, minSize=13, maxW=268, color=BROWN, red=RED, bold=False))
    elif s['k'] == 'b': jobs.append(dict(id='s%d' % i, lines=[tx], font=fn, size=14.5, minSize=11, maxW=196, color=BROWN, red=RED, bold=False))
    elif s['k'] == 'peg':
        jobs.append(dict(id='s%d' % i, lines=PEG_BREAK[lang].split('\n'), font=FONTB[lang], size=20, minSize=13, maxW=126, color=BLACK, red=RED, bold=True))
        jobs.append(dict(id='p%d' % i, lines=[s['btn'][lang]], font=FONTB[lang], size=13, minSize=10, maxW=122, color=PILL, red=RED, bold=True))
    elif s['k'] == 'bowl':
        for n, t in enumerate(tx.split('|')): jobs.append(dict(id='w%d_%d' % (i, n), lines=[t], font=fn, size=15, minSize=11, maxW=132, color=RED, red=RED, bold=False))
od = HERE + '/txt_' + lang
os.makedirs(od, exist_ok=True)
json.dump(jobs, open(od + '/jobs.json', 'w'), ensure_ascii=False)
subprocess.run([HERE + '/render', od + '/jobs.json', od], check=True, capture_output=True)
meta = {}
for m in json.load(open(od + '/out.json')):
    im = Image.open('%s/%s_%d.png' % (od, m['id'], m['n'])).convert('RGBA')
    meta.setdefault(m['id'], []).append((im, m))

# ---- 2. 乾淨底圖（對話框沒字的那一格；碗標籤出現前那一格）----
def grab(t):
    r = subprocess.run([FF, '-v', 'error', '-ss', str(t), '-i', SRC, '-frames:v', '1', '-f', 'rawvideo', '-pix_fmt', 'rgb24', '-'], capture_output=True, check=True)
    return np.frombuffer(r.stdout, np.uint8).reshape(H, W, 3).copy()
def clean_dialog():   # 4.5～4.8 秒逐格找對話框文字區最乾淨的那一格（只有 1 格是空的）
    r = subprocess.run([FF, '-v', 'error', '-ss', '4.5', '-t', '0.3', '-i', SRC, '-f', 'rawvideo', '-pix_fmt', 'rgb24', '-'], capture_output=True, check=True)
    d = np.frombuffer(r.stdout, np.uint8).reshape(-1, H, W, 3)
    dark = [int((f[372:436, 160:795].astype(int).sum(2) < 450).sum()) for f in d]
    assert min(dark) == 0, dark
    c = d[dark.index(0)].copy()
    reg = c[372:436, 160:790]   # 補掉打字第一筆的淡殘點（框邊不動）
    bad = reg.astype(int).sum(2) < 640
    reg[bad] = np.median(reg[~bad], 0).astype(np.uint8)
    return c
CLEAN_D = clean_dialog()
CLEAN_W = grab(30.75)
DR = (160, 372, 795, 436)   # 對話框文字區 x0,y0,x1,y1

def paste(fr, im, x, y, alpha=1.0, cut=None, off=False):
    """把文字圖貼到畫面（im 左上角放在 x,y）；cut＝只顯示左邊幾像素（打字效果）"""
    a = np.asarray(im).astype(np.float32)
    if cut is not None:
        c = max(0, min(a.shape[1], int(cut)))
        a = a.copy(); a[:, c:, 3] = 0
    x, y = int(round(x)), int(round(y))
    h, w = a.shape[:2]
    FH, FW = fr.shape[:2]
    x0, y0, x1, y1 = max(0, x), max(0, y), min(FW, x + w), min(FH, y + h)
    if x1 <= x0 or y1 <= y0: return
    sub = a[y0 - y:y1 - y, x0 - x:x1 - x]
    al = sub[..., 3:4] / 255.0 * alpha
    reg = fr[y0:y1, x0:x1].astype(np.float32)
    fr[y0:y1, x0:x1] = (reg * (1 - al) + sub[..., :3] * al).astype(np.uint8)

def ringfill(fr, x0, y0, x1, y1):
    ring = np.concatenate([fr[y0:y1, x0 - 2:x0].reshape(-1, 3), fr[y0:y1, x1:x1 + 2].reshape(-1, 3), fr[y0 - 2:y0, x0:x1].reshape(-1, 3), fr[y1:y1 + 2, x0:x1].reshape(-1, 3)])
    fr[y0:y1, x0:x1] = np.median(ring, 0).astype(np.uint8)

LAST = {}
def fill_mask(F, mk, r=6):
    """把 mk 標記的像素用周圍（非 mk）像素的模糊平均補平"""
    wgt = (~mk).astype(np.float32)
    num = np.stack([np.asarray(Image.fromarray(np.clip(F[..., c] * wgt, 0, 255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(r))).astype(np.float32) for c in range(3)], 2)
    den = np.asarray(Image.fromarray((wgt * 255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(r))).astype(np.float32)[..., None] / 255
    return np.where(mk[..., None], num / np.maximum(den, 1e-3), F).astype(np.uint8)
def grow(mask, k=5):
    return np.asarray(Image.fromarray((mask * 255).astype(np.uint8)).filter(ImageFilter.MaxFilter(k))) > 0
def slope(F, B, mask):
    fm, lm = F[mask].mean(1), B[mask].mean(1); fm = fm - fm.mean(); lm = lm - lm.mean()
    return float(np.clip((fm * lm).sum() / max(1e-6, (lm * lm).sum()), 0, 1))
NM = (slice(334, 361), slice(91, 272))   # 「Chef Bear」名牌
def opacity(fr):   # 對話框目前的不透明度：名牌字樣跟乾淨框的回歸斜率（背景跟名牌字形無關，不影響）
    a = fr[NM].astype(np.float32).mean(2); c = CLEAN_D[NM].astype(np.float32).mean(2)
    a -= a.mean(); c -= c.mean()
    return float(np.clip((a * c).sum() / (c * c).sum(), 0, 1)) if boxness(fr) > 0.02 or (a * c).sum() / (c * c).sum() > 0.2 else 0.0
def boxness(fr):
    r = fr[330:368, 320:780].astype(int)
    return ((r[..., 0] > 235) & (r[..., 1] > 222) & (r[..., 2] > 185) & (r[..., 2] < 235)).mean()

def draw(fr, t):
    if 17.95 <= t < 18.5:   # 「看位置」標籤彈出動畫：改成「彈出前畫面」淡入「已換好譯文的標籤」
        x0, y0, x1, y1 = 300, 130, 700, 300
        pre, post = PEG_IN()
        a = min(1.0, max(0.0, (t - 18.05) / 0.4))
        fr[y0:y1, x0:x1] = (pre[y0:y1, x0:x1] * (1 - a) + post[y0:y1, x0:x1] * a).astype(np.uint8)
    for i, s in enumerate(segs):
        if not (s['s'] <= t < s['e']): continue
        k = s['k']
        if k == 'd':
            x0, y0, x1, y1 = DR
            o = opacity(fr)
            if o < 0.04: continue
            O = CLEAN_D[y0:y1, x0:x1].copy()
            L = meta['s%d' % i]
            tot = sum(m['w'] for _, m in L)
            td = min(1.3, max(0.45, tot / 520.0))
            budget = (t - s['s']) / td * tot
            for n, (im, m) in enumerate(L):
                base = 395 + 24 * n       # 每行基線（照原片）
                paste(O, im, 169 - m['pad'] - x0, base - m['asc'] - m['pad'] - y0, 1, cut=m['pad'] + budget, off=True)
                budget -= m['w']
                if budget <= 0: budget = -1e9
            F = fr[y0:y1, x0:x1].astype(np.float32)
            if o >= 0.97:
                LAST[i] = F.copy(); fr[y0:y1, x0:x1] = O
            elif i in LAST:   # 淡出：框比字淡得快 → 先把英文字跡用周圍顏色補平，再依英文當下濃淡疊譯文
                Lf = LAST[i]
                mask = np.abs(Lf - CLEAN_D[y0:y1, x0:x1].astype(np.float32)).sum(2) > 40
                mk = np.asarray(Image.fromarray((mask * 255).astype(np.uint8)).filter(ImageFilter.MaxFilter(5))) > 0
                if mask.sum() > 50:
                    fm, lm = F[mask].mean(1), Lf[mask].mean(1); fm = fm - fm.mean(); lm = lm - lm.mean()
                    ot = float(np.clip((fm * lm).sum() / max(1e-6, (lm * lm).sum()), 0, 1))
                else: ot = o
                wgt = (~mk).astype(np.float32)
                num = np.stack([np.asarray(Image.fromarray(np.clip(F[..., c] * wgt, 0, 255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(6))).astype(np.float32) for c in range(3)], 2)
                den = np.asarray(Image.fromarray((wgt * 255).astype(np.uint8)).filter(ImageFilter.GaussianBlur(6))).astype(np.float32)[..., None] / 255
                fill = num / np.maximum(den, 1e-3)
                G = np.where(mk[..., None], fill, F).astype(np.uint8)
                for n, (im, m) in enumerate(L):
                    paste(G, im, 169 - m['pad'] - x0, 395 + 24 * n - m['asc'] - m['pad'] - y0, ot, off=True)
                fr[y0:y1, x0:x1] = G
            else:
                fr[y0:y1, x0:x1] = (F * (1 - o) + O * o).astype(np.uint8)
        else:
            static(fr, i, s, k)

REG = {'q': (280, 262, 572, 295), 'b': (323, 307, 532, 328), 'peg': (326, 154, 462, 278), 'bowl': (150, 320, 428, 348)}
REF = {}   # 原片定格畫面（判斷彈出框／標籤是否已經定位）
for x in segs:
    if x['k'] in REG: REF[id(x)] = None
def ref_of(s):
    if REF.get(id(s)) is None:
        mid = {'q': 0.55, 'b': 0.55, 'peg': 0.6, 'bowl': 0.25}[s['k']]
        REF[id(s)] = grab(s['s'] + (s['e0'] - s['s']) * mid)
    return REF[id(s)]
LASTS = {}
def overlay(fr, i, s, k):
    """在 fr（整張畫面）畫上譯文"""
    if k == 'q':
        ringfill(fr, 280, 262, 572, 295)
        im, m = meta['s%d' % i][0]
        paste(fr, im, 425 - im.width / 2, 284 - m['asc'] / 2 - m['pad'] - 2 + m['desc'] / 2, 1)
    elif k == 'b':
        ringfill(fr, 323, 307, 532, 328)
        im, m = meta['s%d' % i][0]
        paste(fr, im, 427 - im.width / 2, 318 - (m['asc'] - m['desc']) / 2 - m['pad'] - m['desc'], 1)
    elif k == 'peg':
        orig = fr.copy()
        ringfill(fr, 326, 154, 462, 278)
        red = (orig[..., 0].astype(int) > 150) & (orig[..., 1] < 110) & (orig[..., 2] < 110)
        red[:, :424] = False
        fr[red] = orig[red]
        L = meta['s%d' % i]
        for n, (im, m) in enumerate(L):
            base = 183 + 28 * n if len(L) > 1 else 197
            paste(fr, im, 393 - im.width / 2, base - m['asc'] - m['pad'], 1)
        im, m = meta['p%d' % i][0]
        pw, ph, cy = int(m['w'] + 16), 26, 247
        pill = Image.new('RGBA', (pw + 2, ph + 2), (0, 0, 0, 0))
        ImageDraw.Draw(pill).rounded_rectangle([0, 0, pw, ph], radius=13, fill=(255, 255, 255, 255), outline=(216, 200, 190, 255), width=1)
        paste(fr, pill, int(393 - pw / 2), cy - ph / 2, 1)
        paste(fr, im, 393 - im.width / 2, cy - (m['asc'] - m['desc']) / 2 - m['pad'] - m['desc'] + 0.5, 1)
    elif k == 'bowl':
        fr[320:348, 150:428] = CLEAN_W[320:348, 150:428]
        for n, cx in enumerate((222, 357)):
            im, m = meta['w%d_%d' % (i, n)][0]
            paste(fr, im, cx - im.width / 2, 339 - m['asc'] - m['pad'], 1)
def static(fr, i, s, k):
    """彈出框／標籤：跟原片定格畫面一致才換字；之後淡出時依濃淡同步淡出"""
    gk = 'q' if k == 'b' else k          # 按鈕跟著同一個彈出框判斷（按下時按鈕會變色）
    x0, y0, x1, y1 = REG[gk]
    ref = ref_of(s)[y0:y1, x0:x1].astype(np.float32)
    F = ORIG[0][y0:y1, x0:x1].astype(np.float32)
    d = np.abs(F - ref).mean()
    X0, Y0, X1, Y1 = REG[k]
    if d < (30 if gk == 'q' else 15):
        before = fr[Y0:Y1, X0:X1].astype(np.float32).copy()
        overlay(fr, i, s, k)
        LASTS[i] = (before, fr[Y0:Y1, X0:X1].astype(np.float32).copy())
    elif i in LASTS and t_now[0] > s['s'] + 0.3:   # 淡出
        bo, af = LASTS[i]
        G = fr[Y0:Y1, X0:X1].astype(np.float32)
        if k in ('peg', 'bowl'):
            if k == 'bowl': en = (bo[..., 0] > 150) & (bo[..., 1] < 110) & (bo[..., 2] < 110)   # 紅字
            else:
                en = (bo.sum(2) < 300) | ((bo[..., 0] > 150) & (bo[..., 1] < 110) & (bo[..., 2] < 110))
                en[:, 424 - X0:] = False      # 箭頭不動
            if en.sum() < 30: return
            ot = slope(G, bo, en)
            if ot < 0.04: return
            ours = np.abs(af - bo).sum(2) > 30   # 我們畫上去的範圍（譯文、按鈕）
            H2 = fill_mask(G, grow(en, 5))
            Hf = H2.astype(np.float32)
            Hf[ours] = Hf[ours] * (1 - ot) + af[ours] * ot
            fr[Y0:Y1, X0:X1] = np.clip(Hf, 0, 255).astype(np.uint8)
        else:
            a = G.mean(2) - G.mean(); c = bo.mean(2) - bo.mean()
            o = float(np.clip((a * c).sum() / max(1e-6, (c * c).sum()), 0, 1))
            if o > 0.05: fr[Y0:Y1, X0:X1] = np.clip(G + o * (af - bo), 0, 255).astype(np.uint8)
t_now = [0]
_PEG = []
def PEG_IN():
    if not _PEG:
        pi = [j for j, x in enumerate(segs) if x['k'] == 'peg'][0]
        post = grab(18.55); overlay(post, pi, segs[pi], 'peg')
        _PEG.extend([grab(17.95).astype(np.float32), post.astype(np.float32)])
    return _PEG
ORIG = [None]

out = HERE + ('/prev_%s.mp4' % lang if prev else '/chef-bear-%s.mp4' % lang)
ss = ['-ss', str(prev[0]), '-to', str(prev[1])] if prev else []
dec = subprocess.Popen([FF, '-v', 'error'] + ss + ['-i', SRC, '-f', 'rawvideo', '-pix_fmt', 'rgb24', '-'], stdout=subprocess.PIPE)
enc = subprocess.Popen([FF, '-v', 'error', '-y', '-f', 'rawvideo', '-pix_fmt', 'rgb24', '-s', '%dx%d' % (W, H), '-r', '30000/1001', '-i', '-'] + ss + ['-i', SRC,
    '-map', '0:v', '-map', '1:a', '-c:v', 'libx264', '-preset', 'slow', '-crf', '23', '-pix_fmt', 'yuv420p', '-profile:v', 'main', '-c:a', 'aac', '-b:a', '160k', '-shortest', '-movflags', '+faststart', out], stdin=subprocess.PIPE)
i = 0; t0 = prev[0] if prev else 0
while True:
    b = dec.stdout.read(W * H * 3)
    if len(b) < W * H * 3: break
    fr = np.frombuffer(b, np.uint8).reshape(H, W, 3).copy()
    t_now[0] = t0 + i / FPS; ORIG[0] = fr.copy(); draw(fr, t_now[0])
    enc.stdin.write(fr.tobytes()); i += 1
enc.stdin.close(); enc.wait(); dec.wait()
print(out, i, 'frames', os.path.getsize(out) // 1024, 'KB')
