/* ⚠️ 本檔必須最後載入（.clasp.json filePushOrder 最後一個）：程式碼.js 也有一支 doGet（排班），後載入的才生效。2026-10-24 */
var REVIEWS_SHEET_ID = '1kzJEBP0CVR7LL2Jeoxd1L2AdF345z_zTe5ytG_djOTw';
var REVIEWS_SHEET_NAME = '工作表1';

function doGet(e) {
  var payload;
  try {
    if (e && e.parameter && e.parameter.action === 'aistatus') {   /* AI 判讀進度（2026-10-24） */
      payload = JSON.stringify(airStatus_());
      var cb0 = e.parameter.callback;
      if (cb0) return ContentService.createTextOutput(cb0 + '(' + payload + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
      return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
    }
    if (e && e.parameter && e.parameter.action === 'aisample') {   /* 抽查 AI 判讀結果 */
      payload = JSON.stringify(airSample_(e.parameter.n));
      return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
    }
    if (e && e.parameter && e.parameter.action === 'recent') {   /* 2026-10-30 決策中心店長看板／門市今日：近期評論（含評論時間），只回本月、上月與近 40 天 */
      payload = JSON.stringify(recentReviews_());
      var cb1 = e.parameter.callback;
      if (cb1) return ContentService.createTextOutput(cb1 + '(' + payload + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
      return ContentService.createTextOutput(payload).setMimeType(ContentService.MimeType.JSON);
    }
    var ai = {};
    try { ai = airMap_(); } catch (eAi) { ai = {}; }   /* AI 分頁讀不到也不影響原本資料 */
    var sh = SpreadsheetApp.openById(REVIEWS_SHEET_ID).getSheetByName(REVIEWS_SHEET_NAME);
    if (!sh) throw new Error('找不到分頁：' + REVIEWS_SHEET_NAME);
    var lastRow = sh.getLastRow();
    var rows = [];
    if (lastRow >= 2) {
      var vals = sh.getRange(2, 1, lastRow - 1, 11).getValues();
      for (var i = 0; i < vals.length; i++) {
        var r = vals[i];
        if (r[0] === '' || r[0] === null) continue;
        var a = ai[String(r[5] || '')] || ['', '', '', '', ''];
        /* 0–7 原欄位不變；8 起為 AI 判讀：people(Y/N)、夥伴、稱讚重點、負評類型、回覆草稿 */
        rows.push([r[0], r[1], fmtYM_(r[6]), r[7], r[8], r[9], r[10], r[4], a[0], a[1], a[2], a[3], a[4]]);
      }
    }
    payload = JSON.stringify({ ok: true, count: rows.length, rows: rows });
  } catch (err) {
    payload = JSON.stringify({ ok: false, error: String(err) });
  }
  var cb = e && e.parameter && e.parameter.callback;
  if (cb) return ContentService.createTextOutput(cb + '(' + payload + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

function fmtYM_(val) {
  if (val === null || val === '') return '';
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return val.getFullYear() + '-' + ('0' + (val.getMonth() + 1)).slice(-2);
  }
  var mt = String(val).trim().match(/^(\d{4})[-\/](\d{1,2})/);
  return mt ? mt[1] + '-' + ('0' + mt[2]).slice(-2) : String(val).trim();
}
/* 2026-10-30：近期評論（給決策中心 GAS action=reviews 用）。每列 [0門市,1評論者,2評論時間 yyyy-MM-dd HH:mm,3年月,4星等數字,5內容,6 AI people,7 AI夥伴,8 AI負評分類]
   範圍：年月＝本月或上月，或評論時間在 40 天內。結果暫存 10 分鐘。 */
function recentReviews_() {
  var cache = CacheService.getScriptCache(), ck = 'recent|v1', hit = cache.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch (eC) { } }
  var tz = Session.getScriptTimeZone() || 'Asia/Taipei', now = new Date();
  var cur = Utilities.formatDate(now, tz, 'yyyy-MM'), pd = new Date(now.getFullYear(), now.getMonth() - 1, 1), prv = Utilities.formatDate(pd, tz, 'yyyy-MM');
  var from = Utilities.formatDate(new Date(now.getTime() - 40 * 86400000), tz, 'yyyy-MM-dd');
  var ai = {};
  try { ai = airMap_(); } catch (eAi) { ai = {}; }
  var sh = SpreadsheetApp.openById(REVIEWS_SHEET_ID).getSheetByName(REVIEWS_SHEET_NAME), last = sh.getLastRow(), rows = [];
  if (last >= 2) {
    var vals = sh.getRange(2, 1, last - 1, 8).getValues();
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];
      if (r[0] === '' || r[0] === null) continue;
      var ym = fmtYM_(r[6]), t = r[2], ts = (Object.prototype.toString.call(t) === '[object Date]') ? Utilities.formatDate(t, tz, 'yyyy-MM-dd HH:mm') : String(t || '').replace('T', ' ').slice(0, 16);
      if (ym !== cur && ym !== prv && !(ts.slice(0, 10) >= from)) continue;
      var a = ai[String(r[5] || '')] || ['', '', '', '', ''];
      rows.push([r[0], r[1], ts, ym, r[7], String(r[4] || "").slice(0, 400), a[0], a[1], a[3]]);
    }
  }
  var res = { ok: true, at: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm'), cur: cur, prv: prv, rows: rows };
  try { var js = JSON.stringify(res); if (js.length < 95000) cache.put(ck, js, 600); } catch (eP) { }
  return res;
}
