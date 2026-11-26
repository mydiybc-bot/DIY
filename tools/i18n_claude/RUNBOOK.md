# 平板外語食譜：每日巡檢與翻譯作業手冊（Claude Code，2026-10-24；2026-11-05 改版）

> **2026-11-23 起改回這份手冊（經營者：「未來這個外語食譜的更新，不要使用 API，全部都由 Claude code 對話處理就好」）**：Google 自動翻譯已關（`i18nAutoSet` on:0，`I18N_AUTO`＝0），Google 只剩每天 06:30 `i18nDaily` 重抓原文、`i18nSnapTick` 推快照，**都不呼叫 API**。原文改版後新譯文放行前，平板照舊顯示上一版（不會下架）；新食譜要等 Claude Code 翻完才會出現在平板。翻譯一律照下面流程由 Claude Code（訂閱）做：翻譯員、回讀員都是 Claude Code 的 Agent，不要呼叫 `i18nKick`、`i18nAutoSet test`、`i18nZTry`／`i18nZReview`／`i18nZGloss`（這些都會花 API）。
> （舊註記）2026-11-05 晚起，平常不用這份手冊：經營者改用 **Google 排程自動翻**（「改 Google，花 API」，Mac 不用開）。申請單 GAS v38（部署 @94）：每天 06:30 `i18nDaily` 更新原文 → `i18nAutoStart_` 讓 5 個語言各自接力（`i18nAuto_ja`…）直到全部翻完；Sonnet 只翻改到的頁 → AI 回讀（`i18nReview_`）→ 有問題修一輪 → 過了自動放行（放行人「AI 自動檢查」），沒過存待檢查、平板照舊顯示上一版。Claude 排程 `weekday-recipe-i18n-check` 已停用。
> 這份手冊留著當**備用**：Google 那邊卡住（status.py 有待檢查／失敗）或經營者指定用 Claude Code 翻時才照下面做。開關：`i18nAutoSet`（投遞通行碼，on:0／1）。

> **2026-11-24 原文改用中文食譜**（經營者：「平板英文和其他語言食譜全部改用中文直接翻譯」、92 支）：`zset.py` 看候選／打開哪幾支；打開的中文食譜才進待翻清單（pull.py 標【中文原文】，語言多一個 en）。翻譯、回讀照 TRANSLATE.md／REVIEW.md 的「中文原文」段。有舊英文原文的那支，中文版 6 種語言都放行後平板自動換成中文版、舊的不再列也不再翻；之前照舊顯示舊的。`zset.py` 只由主控端在對話裡用，每天排程不用。

經營者裁定：
- 2026-10-24：用 Claude Code（訂閱）翻，不花 API；主廚不核准，**Claude 自己檢查通過就放行**。
- 2026-11-05：**每天早上 06:40**（原文 06:30 更新後）抓到有改的就動手翻（**2026-11-26 經營者改成週一到週五 10:00、排程與子代理都用 Opus 5.5 中等思考**），**不設組數上限，全部改完才算完成**；改的過程**平板不下架**（照舊顯示上一版，頁面不提示），新版通過才換上；**全程不能跳「允許」要人按**。

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
   2. 派 **翻譯員**（Agent，model=opus、effort=medium；2026-11-26 起）：照 `TRANSLATE.md` 翻（只翻 src.json 裡的頁），寫 `<tag>.tr.json`，跑 `precheck.py` 到全過。
   3. 派 **獨立回讀員**（另一個 Agent，model=opus、effort=medium，沒參與翻譯）：照 `REVIEW.md` 逐頁對照，寫 `<tag>.review.json`。
   4. 回讀 `verdict=pass`（沒有 block）→ `python3 …/put.py <tag> --approve`（伺服器再驗一次硬檢查，過了就放行，放行人「Claude」，平板換新版）。
   5. 回讀有 block → 交回原翻譯員修（SendMessage），改完重跑 precheck 與回讀，**最多修一輪**。還是不過 → `python3 …/put.py <tag>`（不帶 --approve，存成待檢查；**平板照舊顯示上一版**），報告列出。
   6. 伺服器硬檢查回 `review`（例：多了原文沒有的數字）→ 看 check 訊息，用 Edit 工具改 tr.json（通常是拿掉多的量詞 1장、1本）→ 重新 pull → precheck → put --approve。
   7. 詞典明顯錯的 → 該詞放進 `tr.json` 的 `dict`，投遞加 `--overwrite`，報告列出。
   8. 主控端自己要修譯文時，也用 Edit 工具改 `tr.json`，不要用 `sed`。
4. **收尾**：`python3 …/status.py --clean` → 各語言 live／review 數，清掉已上線那幾組的暫存檔（留 `opus_backup/`）。
5. **報告**（口語、給老闆看）：翻了幾組（各幾頁）、放行幾組、卡在待檢查的幾組與原因（平板仍顯示上一版）、剩幾組（應為 0）、有沒有發現詞典錯誤。

## 到店須知（QR Code 頁）甜點資料：每次巡檢最後一起做（2026-11-26 起）
經營者：「QR code 掃描後的資料還是停留在 26 支食譜，這不對…未來要能和每天的外語食譜更新一起進行更新檢查」；「88 支都要上，不要自行合併」；介紹「一律用後台中文食譜」。
外國客人掃店內 QR Code → `static/store-guide.html` → 甜點清單＝平板外語食譜「中文原文」上線中的每一支；售價、尺寸、製作時間、產品特色讀門市食譜後台（申請單 `guideSrc`，唯讀），外語品名＝平板已放行的譯名，產品特色逐行照中文翻。
平板那段做完（待翻 0）之後：
1. `python3 /Users/diybc/Desktop/Claude/Code/DIY/tools/i18n_claude/guide.py pull` → 印出各語言「待翻 N 行」（只有後台改了產品特色／尺寸、或平板新增食譜才會有）。
2. 有待翻的語言：派翻譯員（Agent，opus、medium）照 `GUIDE_TRANSLATE.md`「翻譯員」寫 `guide/tr_<語言>.json` → 另派回讀員照「回讀員」寫 `guide/review_<語言>.json` → 有 issues 就用 Edit 工具把 fix 改進 tr_<語言>.json（只修一輪）。
3. `python3 …/guide.py put`（缺譯會擋下不送）→ 到店須知伺服器檢查（每款 6 語品名、介紹行數＝中文、沒有殘留中文、照片打得開；款數驟減擋下）過了才發佈，沒過保留上一版。回報「到店須知：N 款，新上架／下架／改價／說明有改」。
4. 待翻 0 也要跑一次 `guide.py put`（售價、尺寸、上下架不用翻也要更新）。
- 壽星款（名稱有「限當日壽星」）：產品特色從「當日壽星優惠」那段起不顯示，改用頁面固定說明（2026-10-29 經營者：出示證件即可、不用官網會員）。
- 原本「每天 23:00 讀官網 English Menu」已停用（到店須知 GAS `RETIRED=true`，回滾改 false 重新部署）。

## 總檢（2026-11-25，大批上線後或懷疑某類誤譯時）
回讀員常漏看固定幾類錯（剪花嘴、撕下再貼、多出尺寸、冷藏冷凍、表面／側面、擀／揉、店內無借冰、保持乾燥、熱水燙、重置）。做法（都不花 API、只用唯讀預覽通行碼讀）：
1. `python3 zsweep.py [id8 …]` 讀回上線譯文與原文 → `WORK/sweep/<id8>_<lang>.json`
2. `python3 zflag.py` → `python3 zflag2.py` → `python3 zpack.py`：標記、各語言關鍵字自動判、切成 `sweep/_review_<lang>.json`
3. 每語言派一位核對員（sonnet），照 `WORK/sweep/TASK_CHECK.md` 寫 `sweep/_found_<lang>.json`（old／new 片段）
4. `python3 zfix.py <lang> --dry` 試算 → `python3 zfix.py <lang>`：改 tr.json（備份到 sweep/backup/）、partial 投遞＋放行
5. 再跑一次 zsweep 確認新字已上線。
