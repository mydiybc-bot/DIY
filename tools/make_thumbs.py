#!/usr/bin/env python3
"""食譜封面縮圖（2026-10-17 經營者：「有些圖片太大無法直接看，應該縮小讓使用者一眼看清楚」）

門市食譜後台的封面放在公開的雲端空間 recipeimages/{食譜id}/{食譜id}.jpg，有的到 19MB。
本程式把「超過 200KB」的封面縮成長邊 480px 的 JPG，放進 static/thumbs/r/{食譜id}.jpg，
新品申請表「🧬 食譜演化」優先顯示縮圖（200KB 以下的直接用原圖）。

- 已經有縮圖、而且原圖沒換過（檔案大小一樣）的不重做；後台刪掉的食譜，縮圖也刪掉。
- 每天由 GitHub Actions（.github/workflows/recipe-thumbs.yml）在台灣時間凌晨 2:40 跑一次；本機也可以直接跑：
    python3 tools/make_thumbs.py
- 縮圖用 Pillow；本機沒有 Pillow 時改用 macOS 內建的 sips。
"""
import json, os, re, shutil, subprocess, sys, tempfile, urllib.error, urllib.request

LIST = 'https://diybcstorage.blob.core.windows.net/recipeimages?restype=container&comp=list&maxresults=5000'
BASE = 'https://diybcstorage.blob.core.windows.net/recipeimages/'
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(ROOT, 'static', 'thumbs', 'r')
INDEX = os.path.join(OUT, 'index.json')
MIN_KB = 200      # 超過這個才縮
EDGE = 480        # 長邊幾 px
QUALITY = 72


def fetch(url, timeout=60):
    req = urllib.request.Request(url, headers={'User-Agent': 'DIYBC-thumbs'})
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return r.read()
    except urllib.error.URLError as e:
        if 'CERTIFICATE' not in str(e):
            raise
        # 經營者的 Mac 上 Python 憑證不完整（2026-10 起已知）→ 改用 curl
        r = subprocess.run(['curl', '-sfL', '--max-time', str(timeout), url], capture_output=True)
        if r.returncode != 0:
            raise
        return r.stdout


def list_covers():
    out, marker = {}, ''
    for _ in range(12):
        x = fetch(LIST + ('&marker=' + urllib.request.quote(marker) if marker else '')).decode('utf-8')
        for b in re.findall(r'<Blob>(.*?)</Blob>', x, re.S):
            name = re.search(r'<Name>([^<]*)</Name>', b)
            size = re.search(r'<Content-Length>(\d+)</Content-Length>', b)
            if not name or not size:
                continue
            m = re.match(r'^([0-9a-f-]{36})/\1\.(jpe?g|png|gif|webp)$', name.group(1), re.I)
            if m:
                out[m.group(1).lower()] = {'path': name.group(1), 'bytes': int(size.group(1))}
        nm = re.search(r'<NextMarker>([^<]*)</NextMarker>', x)
        marker = nm.group(1) if nm else ''
        if not marker:
            break
    return out


def shrink(src, dst):
    try:
        from PIL import Image, ImageOps
        with Image.open(src) as im:
            im = ImageOps.exif_transpose(im)
            im.thumbnail((EDGE, EDGE))
            if im.mode not in ('RGB', 'L'):
                bg = Image.new('RGB', im.size, (255, 255, 255))
                bg.paste(im, mask=im.convert('RGBA').split()[-1])
                im = bg
            im.convert('RGB').save(dst, 'JPEG', quality=QUALITY, optimize=True, progressive=True)
        return True
    except ImportError:
        tmp = dst + '.tmp.jpg'
        r = subprocess.run(['sips', '-Z', str(EDGE), '-s', 'format', 'jpeg', '-s', 'formatOptions', str(QUALITY), src, '--out', tmp], capture_output=True)
        if r.returncode == 0 and os.path.exists(tmp):
            os.replace(tmp, dst)
            return True
        return False


def main():
    os.makedirs(OUT, exist_ok=True)
    try:
        index = json.load(open(INDEX, encoding='utf-8'))
    except Exception:
        index = {}
    covers = list_covers()
    if len(covers) < 100:
        print('雲端封面清單只讀到 %d 張，先不動' % len(covers))
        return 1
    want = {k: v for k, v in covers.items() if v['bytes'] > MIN_KB * 1024}
    made = kept = failed = removed = 0
    tmpdir = tempfile.mkdtemp()
    try:
        for rid, c in sorted(want.items()):
            dst = os.path.join(OUT, rid + '.jpg')
            if os.path.exists(dst) and index.get(rid) == c['bytes']:
                kept += 1
                continue
            src = os.path.join(tmpdir, 'src')
            try:
                open(src, 'wb').write(fetch(BASE + c['path'], timeout=120))
                if shrink(src, dst):
                    index[rid] = c['bytes']
                    made += 1
                else:
                    failed += 1
            except Exception as e:
                print('失敗', rid, e)
                failed += 1
        for f in os.listdir(OUT):
            rid = f[:-4]
            if f.endswith('.jpg') and rid not in want:
                os.remove(os.path.join(OUT, f))
                index.pop(rid, None)
                removed += 1
        index = {k: v for k, v in index.items() if k in want}
        json.dump(index, open(INDEX, 'w', encoding='utf-8'), ensure_ascii=False, sort_keys=True, separators=(',', ':'))
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)
    total = sum(os.path.getsize(os.path.join(OUT, f)) for f in os.listdir(OUT) if f.endswith('.jpg'))
    print('封面 %d 張，要縮 %d 張：新做 %d、沿用 %d、失敗 %d、刪除 %d；縮圖合計 %.1f MB' % (len(covers), len(want), made, kept, failed, removed, total / 1048576))
    return 0


if __name__ == '__main__':
    sys.exit(main())
