# 平板外語食譜：熊主廚開場動畫各語言版（2026-11-03）

外語平板食譜第一頁的「熊主廚」動畫原本只有英文（所有外語食譜共用同一支，檔案指紋 Content-MD5 `zxY7HTTs9Oub93tFKzn3aw==`）。
這裡把對話框裡的英文換成日／韓／越／印／泰文，成品在 `media/tablet-intro/chef-bear-<語言>.mp4`。
平板怎麼換：後端 `gas/recipe-req/i18n.gs` 的 `I18N_VSWAP` 認指紋 → 換成該語言影片網址（jsDelivr，固定 commit）。

## 重做（改譯文或加語言）
1. 改 `lines.json`（每段的出現時間與各語言譯文；`[ ]` 內是紅字）。
2. 把原片放在本資料夾 `bear.mp4`（從任一外語食譜第一頁影片網址下載，指紋要對）。
3. `swiftc -O render.swift -o render`（Mac 原生排版，泰文聲調、越南文、表情符號才正確）。
4. `python3 build.py ja`（ko／vi／id／th 同理）→ 產出 `chef-bear-ja.mp4`；需要 `~/.local/bin/ffmpeg`、numpy、Pillow。
5. 抽格檢查（每段出現、淡出瞬間不可露英文）→ 放進 `media/tablet-intro/` → push → 把新 commit 編號換進 `I18N_VSWAP` 再部署後端。

`build.py` 做法：對話框用「乾淨空白框」蓋掉英文後逐字打出譯文；淡出時依英文濃淡同步淡出；問答框、取料區標籤、碗標籤在定位後換字。
