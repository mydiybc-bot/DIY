# 兩張台灣限定頁海報 → 6 種語言版（保留原圖照片與插畫，只蓋掉中文字換成譯文）
import json, os, subprocess, sys, html
S = os.path.dirname(os.path.abspath(__file__))
IMG = '/private/tmp/claude-501/-Users-diybc-Desktop-SBIR/4450ec64-e24b-49f8-82c3-bcd5bba393e4/scratchpad/imgs/'
OUT = sys.argv[1] if len(sys.argv) > 1 else S + '/out'
os.makedirs(OUT, exist_ok=True)
CHROME = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
FONT = {
 'en': "'Arial Rounded MT Bold','Helvetica Neue',Arial,sans-serif", 'vi': "'Helvetica Neue',Arial,sans-serif", 'id': "'Arial Rounded MT Bold','Helvetica Neue',Arial,sans-serif",
 'ja': "'Hiragino Maru Gothic ProN','Hiragino Sans',sans-serif", 'ko': "'Apple SD Gothic Neo',sans-serif", 'th': "'Sukhumvit Set','Thonburi',sans-serif"}
T = json.load(open(S + '/text.json'))

def box(x1, y1, x2, y2, txt, size, color, bg, extra='', cls=''):
    return ('<div class="b %s" style="left:%dpx;top:%dpx;width:%dpx;height:%dpx;font-size:%dpx;color:%s;background:%s;%s" data-max="%d">%s</div>'
            % (cls, x1, y1, x2 - x1, y2 - y1, size, color, bg, extra, size, txt))

def esc(s): return html.escape(s).replace('\n', '<br>')

def addon(L, t):
    M, TL, P = '#b1d5c7', '#38766b', '#f39b97'
    h = [box(300, 16, 900, 102, '<b>' + esc(t['title']) + '</b>', 58, '#fff', P),
         box(48, 188, 134, 272, '<b>' + esc(t['candles']) + '</b>', 24, '#c9a66f', '#fff', 'border-radius:30px;overflow-wrap:normal;word-break:normal;padding:0 3px', 'lb'),
         box(616, 192, 700, 278, '<b>' + esc(t['deco']) + '</b>', 24, '#c9a66f', '#fff', 'border-radius:30px;overflow-wrap:normal;word-break:normal;padding:0 3px', 'lb'),
         box(294, 536, 376, 616, '<b>' + esc(t['sprinkles']) + '</b>', 22, '#c9a66f', '#fff', 'border-radius:30px;overflow-wrap:normal;word-break:normal;padding:0 3px', 'lb')]
    items = [(150, 446, 338, 516, 'candle'), (372, 446, 584, 516, 'sparkler'), (700, 446, 916, 516, 'balloon'), (936, 446, 1150, 516, 'hat'),
             (1, 1, 1, 1, ''), (40, 1043, 350, 1118, 'choco'), (358, 1043, 578, 1118, 'bag'), (622, 1043, 846, 1118, 'tableware'), (870, 1043, 1128, 1118, 'ice')]
    for x1, y1, x2, y2, k in items:
        if not k: continue
        n, p = t[k]
        h.append(box(x1, y1, x2, y2, '<div class="nm">• ' + esc(n) + '</div><div class="pr">' + esc(p) + '</div>', 27, TL, M, 'flex-direction:column;gap:2px', 'it'))
    h.append(box(60, 640, 304, 712, '<div class="nm">• ' + esc(t['pearls'][0]) + '</div><div class="pr">' + esc(t['pearls'][1]) + '</div>', 27, TL, M, 'flex-direction:column;align-items:flex-end;text-align:right;gap:2px', 'it'))
    h.append(box(898, 640, 1170, 712, '<div class="nm">• ' + esc(t['rainbow'][0]) + '</div><div class="pr">' + esc(t['rainbow'][1]) + '</div>', 27, TL, M, 'flex-direction:column;align-items:flex-start;text-align:left;gap:2px', 'it'))
    h.append(box(0, 1132, 1200, 1200, '<b>★ ' + esc(t['more']) + '</b>', 36, '#fff', P))
    return 'p27_0', 1200, h

def member(L, t):
    K = '#f6d27c'
    feather = 'box-shadow:0 0 14px 12px %s;border-radius:30px' % K
    h = [box(236, 42, 846, 116, '<b>' + esc(t['join']) + '</b>', 54, '#bf5211', K, feather)]
    for (x1, y1, x2, y2), k in zip([(62, 166, 506, 252), (578, 166, 1028, 252), (62, 306, 506, 394), (578, 306, 1028, 394)], ['v1', 'v2', 'v3', 'v4']):
        h.append(box(x1, y1, x2, y2, '<b>' + esc(t[k]) + '</b>', 46, '#7b8e22', '#fff', 'border-radius:40px', 'bb'))
    h.append(box(268, 452, 806, 500, esc(t['note']), 28, '#3a1406', '#f7d88b', 'box-shadow:0 0 10px 8px #f7d88b;border-radius:20px'))
    h.append('<div class="b ttl" style="left:294px;top:512px;width:504px;height:318px" data-max="0"><div class="t1" data-max="74">' + esc(t['t1']) + '</div><div class="t2" data-max="104">' + esc(t['t2']) + '</div></div>')
    h.append(box(166, 982, 914, 1040, '<b>' + esc(t['foot']) + '</b>', 30, '#410908', '#fbe0a6', 'border-radius:29px;box-shadow:0 2px 0 #d9a85a'))
    return 'p28_0', 1080, h

CSS = '''*{margin:0;padding:0;box-sizing:border-box}body{width:%dpx;height:%dpx;overflow:hidden;font-family:%s}
img.bg{position:absolute;left:0;top:0;width:%dpx;height:%dpx}
.b{position:absolute;display:flex;align-items:center;justify-content:center;text-align:center;line-height:1.15;overflow:hidden;padding:0 6px;word-break:keep-all;overflow-wrap:anywhere}
.nm{font-weight:700}.pr{font-size:.78em;font-weight:600}
.ttl{flex-direction:column;gap:6px;background:#d15b19;border:7px solid #5c1d0c;border-bottom:none;border-radius:34px 34px 0 0;padding:10px 14px}
.t1,.t2{font-weight:900;-webkit-text-stroke:6px #5c1d0c;paint-order:stroke fill;line-height:1.05;white-space:nowrap}
.t1{color:#fff}.t2{color:#f6cf2a}'''
FIT = '''<script>
function fit(el,max){var s=max;el.style.fontSize=s+'px';while(s>9&&(el.scrollWidth>el.clientWidth+1||el.scrollHeight>el.clientHeight+1)){s-=1;el.style.fontSize=s+'px';}}
document.querySelectorAll('.t1,.t2').forEach(function(e){var p=e.parentNode;var s=+e.dataset.max;e.style.fontSize=s+'px';while(s>12&&e.scrollWidth>p.clientWidth-28){s--;e.style.fontSize=s+'px';}});
document.querySelectorAll('.b').forEach(function(e){if(+e.dataset.max)fit(e,+e.dataset.max);});
['it','lb','bb'].forEach(function(c){var es=[].slice.call(document.querySelectorAll('.'+c));if(!es.length)return;var m=Math.min.apply(null,es.map(function(e){return parseFloat(e.style.fontSize);}));es.forEach(function(e){e.style.fontSize=m+'px';});});
</script>'''

for L in ['en', 'ja', 'ko', 'vi', 'id', 'th']:
    for name, fn in (('addon', addon), ('member', member)):
        src, W, parts = fn(L, T[name][L])
        doc = '<!doctype html><html lang="%s"><head><meta charset="utf-8"><style>%s</style></head><body><img class="bg" src="file://%s">%s%s</body></html>' % (
            L, CSS % (W, W, FONT[L], W, W), IMG + src, ''.join(parts), FIT)
        hp = '%s/%s_%s.html' % (S, name, L); open(hp, 'w').write(doc)
        png = '%s/%s_%s.png' % (S, name, L)
        subprocess.run([CHROME, '--headless=new', '--disable-gpu', '--hide-scrollbars', '--force-device-scale-factor=1', '--allow-file-access-from-files',
                        '--window-size=%d,%d' % (W, W), '--screenshot=' + png, 'file://' + hp], capture_output=True, timeout=60)
        subprocess.run(['sips', '-s', 'format', 'jpeg', '-s', 'formatOptions', '88', png, '--out', '%s/%s_%s.jpg' % (OUT, name, L)], capture_output=True)
        print(name, L, os.path.getsize('%s/%s_%s.jpg' % (OUT, name, L)))
