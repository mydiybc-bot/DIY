# 平板外語食譜：每日巡檢與翻譯作業手冊（Claude Code，2026-10-24；2026-11-05 改版）

> **2026-11-05 晚起，平常不用這份手冊**：經營者改用 **Google 排程自動翻**（「改 Google，花 API」，Mac 不用開）。申請單 GAS v38（部署 @94）：每天 06:30 `i18nDaily` 更新原文 → `i18nAutoStart_` 讓 5 個語言各自接力（`i18nAuto_ja`…）直到全部翻完；Sonnet 只翻改到的頁 → AI 回讀（`i18nReview_`）→ 有問題修一輪 → 過了自動放行（放行人「AI 自動檢查」），沒過存待檢查、平板照舊顯示上一版。Claude 排程 `weekday-recipe-i18n-check` 已停用。
> 這份手冊留著當**備用**：Google 那邊卡住（status.py 有待檢查／失敗）或經營者指定用 Claude Code 翻時才照下面做。開關：`i18nAutoSet`（投遞通行碼，on:0／1）。

經營者裁定：
- 2026-10-24：用 Claude Code（訂閱）翻，不花 API；主廚不核准，**Claude 自己檢查通過就放行**。
- 2026-11-05：**每天早上 06:40**（原文 06:30 更新後）抓到有改的就動手翻，**不設組數上限，全部改完才算完成**；改的過程**平板不下架**（照舊顯示上一版，頁面不提示），新版通過才換上；**全程不能跳「允許」要人按**。

後台怎麼運作（2026-11-05 起，申請單 GAS v37）：
- 只改食材份量／名稱／單位、圖片、影片 → **不算改版**，翻譯照常上線（份量與圖影平板直接拿最新原文）。
- 改到文字（標題、內文、甜點名、保存方式）才算改版；**只要重翻字有改的那幾頁**，其他頁自動沿用。
- 改版到新譯文放行之前，平板顯示放行過的上一版。

## 每次執行的流程
工作資料夾與工具：`~/Desktop/Claude/Code/DIY/tools/i18n_claude/`。通行碼從 `~/.config/diybc/secrets.env` 讀（I18N_PUT_KEY），**不印出、不回顯、不寫進任何檔**。

0. **免按允許的鐵則**（排程是一般權限，只有下列寫法免按；2026-11-05 查到早上要按允許的原因＝主控端與翻譯員臨時跑了 `ls`、`grep`、`cat`、`sed`）：
   - Bash 只准這幾種，一律絕對路徑、一個指令一行，不要 `cd`、`&&`、`;`、`|`、`python3 -c`、`rm`：
     - `git -C /Users/diybc/Desktop/Claude/Code/DIY pull --ff-only`
     - `python3 /Users/diybc/Desktop/Claude/Code/DIY/tools/i18n_claude/pull.py …`、`precheck.py …`、`put.py …`、`status.py …`
   - 看檔案內容一律用 **Read 工具**（不要 `cat`／`grep`／`ls`）；改 `tr.json` 用 **Edit／Write 工具**（不要 `sed`）。
   - 派出去的翻譯員／回讀員也照這個寫法（TRANSLATE.md、REVIEW.md 已寫）。
1. **看有沒有事**：`python3 /Users/diybc/Desktop/Claude/Code/DIY/tools/i18n_claude/pull.py --fresh`（今天還沒更新原文就先更新）→ 待翻清單（why＝還沒翻／原文改版／失敗；後面寫「平板顯示上一版」或「平板沒有」）。再跑 `status.py` 看待檢查。待翻 0 且無待檢查 → 寫一行「今天沒事」就結束。
2. **全部做完，不設上限**：5 個語言同時開工，**同一語言一次只做一組**（詞典每語言共用，同語言同時翻新詞會不一致）；一組放行後，該語言接著做下一組，直到 `pull.py` 回待翻 0。優先順序：「平板沒有」的（新食譜）→ 原文改版的。
3. **每一組**（依序）：
   1. `python3 /Users/diybc/Desktop/Claude/Code/DIY/tools/i18n_claude/pull.py <id前8碼> <語言>` → 產生 `<tag>.src.json` 與 `.brief.txt`；印出「全部 N 頁，要翻 M 頁（沿用 K 頁）」。
   2. 派 **翻譯員**（Agent，model=sonnet）：照 `TRANSLATE.md` 翻（只翻 src.json 裡的頁），寫 `<tag>.tr.json`，跑 `precheck.py` 到全過。
   3. 派 **獨立回讀員**（另一個 Agent，model=sonnet，沒參與翻譯）：照 `REVIEW.md` 逐頁對照，寫 `<tag>.review.json`。
   4. 回讀 `verdict=pass`（沒有 block）→ `python3 …/put.py <tag> --approve`（伺服器再驗一次硬檢查，過了就放行，放行人「Claude」，平板換新版）。
   5. 回讀有 block → 交回原翻譯員修（SendMessage），改完重跑 precheck 與回讀，**最多修一輪**。還是不過 → `python3 …/put.py <tag>`（不帶 --approve，存成待檢查；**平板照舊顯示上一版**），報告列出。
   6. 伺服器硬檢查回 `review`（例：多了原文沒有的數字）→ 看 check 訊息，用 Edit 工具改 tr.json（通常是拿掉多的量詞 1장、1本）→ 重新 pull → precheck → put --approve。
   7. 詞典明顯錯的 → 該詞放進 `tr.json` 的 `dict`，投遞加 `--overwrite`，報告列出。
   8. 主控端自己要修譯文時，也用 Edit 工具改 `tr.json`，不要用 `sed`。
4. **收尾**：`python3 …/status.py --clean` → 各語言 live／review 數，清掉已上線那幾組的暫存檔（留 `opus_backup/`）。
5. **報告**（口語、給老闆看）：翻了幾組（各幾頁）、放行幾組、卡在待檢查的幾組與原因（平板仍顯示上一版）、剩幾組（應為 0）、有沒有發現詞典錯誤。
