# Render 上線步驟

這個專案已經整理成可直接部署到 Render 的 `Web Service`。

## 最快上線方式

1. 把這個專案推到 GitHub。
2. 登入 Render。
3. 點 `New +` -> `Blueprint`。
4. 選你的 GitHub repo。
5. Render 會自動讀取 `render.yaml`。
6. 按 `Apply` 建立服務。
7. 等部署完成後，Render 會給你一個 `https://xxxx.onrender.com` 網址。

## 目前設定

- Build Command: `pip install -r requirements.txt`
- Start Command: `python3 server.py`
- Runtime: `Python`
- Plan: `free`

## 很重要

免費方案適合先拿外網網址做快速測試，但資料不保證持久保存。

如果你要給 50 位員工穩定測試，建議至少改成付費方案，並設定持久儲存：

1. 將服務方案改成 `Starter` 或你選擇的付費方案。
2. 在 Render 加一個 Persistent Disk。
3. Mount Path 設為 `/var/data/employee-training`
4. 在 Environment Variables 新增：
   - `TRAINING_DATA_DIR=/var/data/employee-training`

這樣員工帳號、題庫、練習紀錄、進度就不會因為重啟或重部署而消失。

## 上線後先測這幾件事

1. 管理員登入
2. 員工登入
3. 開始練習並完成一題
4. 後台新增員工並儲存
5. 匯出報表 Excel

## 補充

如果 Render 部署後網址出來了，但手機打不開，先確認：

- Render 部署狀態是 `Live`
- 不是還在 `Building`
- 網址是 `https://...onrender.com`
