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