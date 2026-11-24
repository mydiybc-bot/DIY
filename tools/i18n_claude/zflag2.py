#!/usr/bin/env python3
"""總檢第二層（2026-11-25）：從 _flags.json 用各語言關鍵字自動篩出「很可能真的錯」與「要人看」的 → _check.json"""
import json, os, re, collections
from common import WORK
D = os.path.join(WORK, 'sweep'); F = json.load(open(os.path.join(D, '_flags.json'), encoding='utf-8'))
HOTWATER = {'en': r'hot water|boiling', 'ja': r'熱湯|お湯', 'ko': r'뜨거운 물|끓는 물', 'vi': r'nước nóng|nước sôi', 'id': r'air panas|air mendidih', 'th': r'น้ำร้อน'}
KNEAD = {'en': (r'knead', r'roll'), 'ja': (r'こね', r'伸ば|のば'), 'ko': (r'반죽하|치대', r'밀대|밀어|펴'), 'vi': (r'nhào', r'cán'), 'id': (r'uleni|menguleni', r'gilas|giling|rata'), 'th': (r'นวด', r'รีด|คลึง')}
FREEZE = {'en': r'freez|frozen', 'ja': r'冷凍', 'ko': r'냉동', 'vi': r'đông|ngăn đá', 'id': r'beku|freezer|bekukan', 'th': r'แช่แข็ง|ช่องแข็ง|ช่องฟรีซ'}
FRIDGE = {'en': r'refrigerat|fridge|chill|cold', 'ja': r'冷蔵', 'ko': r'냉장', 'vi': r'lạnh|ngăn mát', 'id': r'kulkas|lemari es|dingin', 'th': r'ตู้เย็น|แช่เย็น'}
ICE = {'en': r'\bice\b', 'ja': r'氷', 'ko': r'얼음', 'vi': r'\bđá\b', 'id': r'\bes\b', 'th': r'น้ำแข็ง'}
ICEPACK = r'ice ?pack|保冷剤|아이스팩|túi chườm|túi đá|ice gel|cool ?pack|เจลเย็น'
SIDE = {'en': (r'\bsides?\b', r'surface|top'), 'ko': (r'옆면|측면', r'표면|윗면'), 'ja': (r'側面', r'表面|上面'), 'vi': (r'thành bên|mặt bên|xung quanh', r'bề mặt|mặt trên'), 'id': (r'sisi samping', r'permukaan'), 'th': (r'ด้านข้าง', r'ผิวหน้า|ด้านบน|พื้นผิว')}
REPLAY = {'en': r'replay', 'vi': r'phát lại', 'id': r'putar ulang', 'th': r'เล่นซ้ำ', 'ja': r'再生し直', 'ko': r'다시 재생'}
OVER = {'en': r'\bover\b|cover', 'vi': r'\búp\b', 'ko': r'씌우|덮', 'th': r'สวม|ครอบ', 'ja': r'かぶせ|被せ', 'id': r'tutup|telungkup'}
seg = lambda s, pat: [x for x in re.split(r'(?=\b\d+\.\s)|(?<=[。.!?])\s', s) if re.search(pat, x)]
auto, manual = [], []
for f in F:
    r, l, z, t = f['rule'], f['lang'], f['zh'], f['tr']
    why = ''
    if r == 'B小心燙' and not re.search(r'熱水|溫水|水', z) and re.search(HOTWATER[l], t, re.I): why = '小心燙傷→譯成熱水燙'
    elif r == 'B擀' and re.search(KNEAD[l][0], t, re.I) and not re.search(KNEAD[l][1], t, re.I): why = '擀→譯成揉'
    elif r == 'B冷藏冷凍':
        zf, zr = bool(re.search(r'冷凍', z)), bool(re.search(r'冷藏', z))
        tf, tr_ = bool(re.search(FREEZE[l], t, re.I)), bool(re.search(FRIDGE[l], t, re.I))
        if zr and not zf and tf: why = '冷藏→譯成冷凍'
        elif zf and not zr and not tf: why = '冷凍→沒譯出冷凍'
        elif zr and not tr_ and not tf: why = '冷藏→沒譯出冷藏'
    elif r == 'B無借冰' and re.search(r'냉동 ?보관|保冷剤を?(お)?貸|冷凍保管|penitipan|freezer storage|ice ?packs? (are not|is not) (lent|provided)|no ice ?packs?', t, re.I): why = '店內無借冰→譯成不借保冷劑／不提供冷凍保管'
    elif r == 'B表面抹平' and re.search(SIDE[l][0], t, re.I) and not re.search(SIDE[l][1], t, re.I): why = '表面→譯成側面'
    elif r == 'B重置' and re.search(REPLAY[l], t, re.I): why = '重置→譯成重播'
    elif r == 'B套入' and re.search(OVER[l], t, re.I): why = '套入→可能譯成罩上／倒扣（要人看）'
    if why: auto.append(dict(f, why=why))
    elif r in ('A尺寸', 'B剪花嘴', 'B剪掉1/4', 'B保持乾燥', 'B朝上朝下') or (r == 'B無借冰' and f['k'] == 'pres'): manual.append(f)
uniq = lambda L: list({(x['lang'], x['zh'], x['tr'], x.get('why', x['rule'])): x for x in L}.values())
auto, manual = uniq(auto), uniq(manual)
json.dump({'auto': auto, 'manual': manual}, open(os.path.join(D, '_check.json'), 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
print('自動抓到', len(auto), collections.Counter(x['why'] for x in auto))
print('要人看', len(manual), collections.Counter((x['rule'], x['lang']) for x in manual))
