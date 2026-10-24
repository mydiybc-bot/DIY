# 平板外語食譜：每日巡檢與翻譯作業手冊（Claude Code，2026-10-24）

經營者裁定（2026-10-24）：用 Claude Code（訂閱）翻，不花 API；主廚不核准，**Claude 自己檢查通過就放行**；週一到週五早上 10 點檢查一次。

## 每次執行的流程
工作資料夾與工具：`~/Desktop/Claude/Code/DIY/tools/i18n_claude/`（先 `git -C ~/Desktop/Claude/Code/DIY pull --ff-only`）。通行碼從 `~/.config/diybc/secrets.env` 讀（I18N_PUT_KEY），**不印出、不回顯、不寫進任何檔**。

1. **看有沒有事**：`python3 pull.py`（不帶參數）→ 待翻清單（why＝還沒翻／原文改版／失敗）。`python3 -c "from common import *; print(call('i18nStatus')['data']['review'])"` 看有沒有卡在「待檢查」的。待翻 0 且無待檢查 → 寫一行「今天沒事」就結束。
2. **挑這次要做的**：最多 **15 組**（一次不要太多，訂閱額度有限）。**同一語言一次只做一組**（詞典是每語言共用的，同語言兩組同時翻，新詞的譯法會不一致）；優先挑同一支食譜的不同語言、其次挑原文改版的。
3. **每一組**（依序）：
   1. `python3 pull.py <id前8碼> <語言>` → 產生 `<tag>.src.json` 與 `.brief.txt`。
   2. 派 **翻譯員**（Agent，model=sonnet）：照 `TRANSLATE.md` 翻，寫 `<tag>.tr.json`，跑 `precheck.py` 到全過。不同語言的翻譯員可並行。
   3. 派 **獨立回讀員**（另一個 Agent，model=sonnet，沒參與翻譯）：照 `REVIEW.md` 逐頁對照，寫 `<tag>.review.json`。
   4. 回讀 `verdict=pass`（沒有 block）→ `python3 put.py <tag> --approve`（伺服器再驗一次硬檢查，過了就自動放行，放行人「Claude」）。
   5. 回讀有 block → 把問題交回原翻譯員修（SendMessage 或新派一個帶著 issues 的翻譯員），改完重跑 precheck 與回讀，**最多修一輪**。還是不過 → `python3 put.py <tag>`（不帶 --approve，只存成待檢查），並在報告列出。
   6. 伺服器硬檢查回 `review`（例：多了原文沒有的數字）→ 看 check 訊息，通常是把譯文裡多的量詞（1장、1本）拿掉就過；改完重新 pull（更新缺詞清單）→ precheck → put --approve。
   7. 詞典明顯錯的（譯名跟英文意思對不上）→ 該詞放進 `tr.json` 的 `dict`，投遞時加 `--overwrite`，並在報告列出。
4. **收尾**：`i18nStatus` 看各語言 live／review 數；清掉自己產生、已投遞成功的 `*.src.json／.tr.json／.review.json／.brief.txt`（留 `opus_backup/`）。
5. **報告**（口語、給老闆看）：這次翻了哪幾組、放行幾組、卡在待檢查的有幾組與原因、剩幾組、有沒有發現詞典錯誤。
