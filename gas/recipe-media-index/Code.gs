/**
 * 食譜素材索引（2026-10-05 建）
 * 綁在 Google 試算表「食譜素材索引」上的接收程式：
 * 門市食譜後台（diybc.azurewebsites.net）掃出來的「哪支食譜哪一步用哪個圖片影片」寫進本試算表。
 * 只讀寫自己這一張試算表（權限 spreadsheets.currentonly）；通行碼存在指令碼屬性 TOKEN，不寫在程式裡。
 */
var VERSION = 'recipe-media-index-v2';

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function out_(o, cb) {
  var s = JSON.stringify(o);
  return cb
    ? ContentService.createTextOutput(cb + '(' + s + ')').setMimeType(ContentService.MimeType.JAVASCRIPT)
    : ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
}

function token_() { return PropertiesService.getScriptProperties().getProperty('TOKEN'); }
function auth_(t) { var k = token_(); return !!k && t === k; }

/** 在編輯器執行一次，用來跳出授權視窗 */
function authorize() { return ss_().getName(); }

function doGet(e) {
  var p = (e && e.parameter) || {};
  var cb = /^[A-Za-z_][A-Za-z0-9_]{0,40}$/.test(p.callback || '') ? p.callback : '';
  try {
    var a = p.action || 'ping';
    if (a === 'ping') return out_({ ok: true, version: VERSION, ready: !!token_() }, cb);
    if (a === 'init') {
      if (token_()) return out_({ ok: false, msg: 'already initialised' }, cb);
      if (!p.token || p.token.length < 24) return out_({ ok: false, msg: 'token too short' }, cb);
      PropertiesService.getScriptProperties().setProperty('TOKEN', p.token);
      return out_({ ok: true, name: ss_().getName() }, cb);
    }
    if (!auth_(p.token)) return out_({ ok: false, msg: 'auth' }, cb);
    if (a === 'count') return out_(count_(), cb);
    if (a === 'sample') return out_(sample_(p.sheet, Number(p.row || 2), Number(p.n || 3)), cb);
    if (a === 'blobmap') return out_(blobmap_(), cb);
    return out_({ ok: false, msg: 'unknown action' }, cb);
  } catch (err) {
    return out_({ ok: false, msg: String(err) }, cb);
  }
}

function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents);
    if (!auth_(b.token)) return out_({ ok: false, msg: 'auth' });
    if (b.action === 'prepare') return out_(prepare_(b));
    if (b.action === 'rows') return out_(rows_(b));
    if (b.action === 'finish') return out_(finish_(b));
    return out_({ ok: false, msg: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, msg: String(err) });
  }
}

/** 建分頁、清空、寫表頭。b.sheets=[{name, header[], rows, text[欄號]}] */
function prepare_(b) {
  var ss = ss_(), res = [];
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    b.sheets.forEach(function (d) {
      var sh = ss.getSheetByName(d.name) || ss.insertSheet(d.name);
      if (sh.getFilter()) sh.getFilter().remove();
      sh.clear();
      var needR = d.rows + 1, needC = d.header.length;
      if (sh.getMaxRows() < needR) sh.insertRowsAfter(sh.getMaxRows(), needR - sh.getMaxRows());
      if (sh.getMaxColumns() < needC) sh.insertColumnsAfter(sh.getMaxColumns(), needC - sh.getMaxColumns());
      (d.text || []).forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
      sh.getRange(1, 1, 1, needC).setValues([d.header]);
      res.push({ name: d.name, maxRows: sh.getMaxRows() });
    });
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { ok: true, sheets: res };
}

/** 寫一批列到固定位置（重送不會多寫）。b={sheet, start(第幾筆資料，從 1 起), rows[][]} */
function rows_(b) {
  var sh = ss_().getSheetByName(b.sheet);
  if (!sh) return { ok: false, msg: 'no sheet ' + b.sheet };
  var n = b.rows.length;
  if (!n) return { ok: true, sheet: b.sheet, start: b.start, n: 0 };
  var lock = LockService.getScriptLock();
  lock.waitLock(60000);
  try {
    sh.getRange(b.start + 1, 1, n, b.rows[0].length).setValues(b.rows);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { ok: true, sheet: b.sheet, start: b.start, n: n };
}

/** 收尾排版。b={sheets:[{name, widths[], wrap, filter}], hide[], order[], remove[]} */
function finish_(b) {
  var ss = ss_(), notes = [];
  (b.sheets || []).forEach(function (d) {
    var sh = ss.getSheetByName(d.name);
    if (!sh) return;
    var lr = sh.getLastRow(), lc = sh.getLastColumn();
    if (!lr || !lc) return;
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, lc).setFontWeight('bold').setBackground('#fdd35d');
    (d.widths || []).forEach(function (w, i) { if (w && i < lc) sh.setColumnWidth(i + 1, w); });
    sh.getRange(1, 1, lr, lc)
      .setWrapStrategy(d.wrap ? SpreadsheetApp.WrapStrategy.WRAP : SpreadsheetApp.WrapStrategy.CLIP)
      .setVerticalAlignment('top');
    if (d.filter && lr > 1) {
      if (sh.getFilter()) sh.getFilter().remove();
      sh.getRange(1, 1, lr, lc).createFilter();
    }
    if (sh.getMaxRows() > lr + 1) sh.deleteRows(lr + 2, sh.getMaxRows() - lr - 1);
    if (sh.getMaxColumns() > lc) sh.deleteColumns(lc + 1, sh.getMaxColumns() - lc);
  });
  (b.hide || []).forEach(function (n) { var s = ss.getSheetByName(n); if (s) s.hideSheet(); });
  (b.order || []).forEach(function (n, i) {
    try { var s = ss.getSheetByName(n); if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(i + 1); } }
    catch (err) { notes.push('order ' + n + ': ' + err); }
  });
  (b.remove || []).forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s);
  });
  SpreadsheetApp.flush();
  var c = count_();
  c.notes = notes;
  return c;
}

function count_() {
  return {
    ok: true, version: VERSION, name: ss_().getName(),
    sheets: ss_().getSheets().map(function (s) {
      var lr = s.getLastRow(), filled = 0;
      if (lr > 1) filled = s.getRange(2, 1, lr - 1, 1).getValues().filter(function (r) { return r[0] !== ''; }).length;
      return { name: s.getName(), rows: lr, cols: s.getLastColumn(), filledA: filled, hidden: s.isSheetHidden() };
    })
  };
}

function sample_(name, row, n) {
  var sh = ss_().getSheetByName(name);
  if (!sh) return { ok: false, msg: 'no sheet' };
  n = Math.max(1, Math.min(n, 20));
  return { ok: true, values: sh.getRange(row, 1, n, sh.getLastColumn()).getDisplayValues() };
}

/** 雲端空間檔案清單（檔名、大小、內容指紋、日期），由本機上傳到隱藏分頁 _blob，給掃描頁比對用 */
function blobmap_() {
  var sh = ss_().getSheetByName('_blob');
  if (!sh) return { ok: false, msg: 'no _blob' };
  var lr = sh.getLastRow();
  return { ok: true, rows: lr > 1 ? sh.getRange(2, 1, lr - 1, 4).getValues() : [] };
}
