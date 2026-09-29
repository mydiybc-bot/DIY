/* 訂位管線「預約系統自動抓資料」GAS 專案（綁定「DIYBC 訂位資料」試算表）的部分備份。
 * 本檔＝程式碼.gs 內 fetchRange_ 區塊（原 393–403 行）2026-09-29 版：加暫時性錯誤自動重試。
 * 完整程式碼只存在 Google 端（編輯器文字含登入 cookie 字樣，Chrome 擴充讀取會被擋，無法整檔備份）。
 */

/* ========== 暫時性錯誤自動重試（2026-09-29 新增） ==========
 * 後台偶發 502／503／504 或連線逾時 → 等一下再試，最多 3 次。
 * 時間保護：預估重試做完會超過「本次執行開始後 330 秒」就不重試，直接照舊報錯
 *（GAS 360 秒硬殺，不可以為了重試把整輪拖死）。
 * 只重試伺服器暫時錯誤；302（登入失效）、404 等一律不重試，行為與舊版相同。
 */
const FETCH_RETRY_T0_ = Date.now();            // 本次執行開始時間（每次執行都會重算）
const FETCH_RETRY_WAITS_ = [10000, 30000];     // 第 2、3 次嘗試前各等幾毫秒
const FETCH_RETRY_CODES_ = [500, 502, 503, 504];
const FETCH_RETRY_LIMIT_MS_ = 330000;

function fetchRetry_(url, opts, label) {
  let r = null, err = null;
  for (let i = 0; i <= FETCH_RETRY_WAITS_.length; i++) {
    const t = Date.now();
    r = null; err = null;
    try {
      r = UrlFetchApp.fetch(url, opts);
      if (FETCH_RETRY_CODES_.indexOf(r.getResponseCode()) < 0) {
        if (i > 0) Logger.log(label + '：第 ' + (i + 1) + ' 次嘗試成功');
        return r;
      }
    } catch (e) {
      err = e;
    }
    if (i === FETCH_RETRY_WAITS_.length) break;
    const why = r ? 'HTTP ' + r.getResponseCode() : String(err && err.message || err);
    const wait = FETCH_RETRY_WAITS_[i];
    const cost = Date.now() - t;   // 用這次花的時間估下一次
    if (Date.now() - FETCH_RETRY_T0_ + wait + cost > FETCH_RETRY_LIMIT_MS_) {
      Logger.log(label + '：' + why + '，本次執行時間不夠，不重試');
      break;
    }
    Logger.log(label + '：' + why + '，' + (wait / 1000) + ' 秒後第 ' + (i + 2) + ' 次嘗試');
    Utilities.sleep(wait);
  }
  if (err) throw err;
  return r;
}

function fetchRange_(cookie, ds, ds2) {
  const url = CFG.BASE + '/Reservations?uid=&date=' + ds + '&date2=' + ds2 +
              '&storeId=&type=&status=&memo=&userName=';
  const r = fetchRetry_(url, {
    muteHttpExceptions: true, followRedirects: false,
    headers: { 'User-Agent': CFG.UA, 'Cookie': cookie }
  }, '查詢 ' + ds + '~' + ds2);
  if (r.getResponseCode() === 302) throw new Error('查詢被轉址回登入頁：登入態失效');
  if (r.getResponseCode() !== 200) throw new Error('查詢失敗 HTTP ' + r.getResponseCode() + '（' + ds + '~' + ds2 + '）');
  return parseTable_(r.getContentText(), ds);
}
