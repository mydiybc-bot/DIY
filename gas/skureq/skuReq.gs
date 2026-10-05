/**
 * diybc-skureq-api ｜ skuReq.gs v1.1（2026-10-05：經營者《1005》申請表拿掉 用途／使用頻率／用在哪裡／標配數量 → 這四欄改成選填；
 *   舊版網頁照樣送得進來。審查改每週三，審完由出貨中心通知各店）
 *   v1.0（2026-09-24 cg 批）
 * 門市「新品項申請」的寫入端。只寫 fact_sku_req（BOM 本＋儀表板專用檔兩邊同步寫）。
 * ⛔ 不碰 dim_sku／dim_store_par：核准轉入由網頁「總部」身分走主程式 type:sku／type:param。
 *
 * 部署步驟：
 *  1. 新建獨立 GAS 專案 diybc-skureq-api，貼上本檔（取代預設 Code.gs 內容）
 *  2. 部署 → 新增部署作業 → 網頁應用程式：執行身分＝我、存取權＝任何人 → 複製 /exec 網址
 *  3. 把網址貼到下面 SR.SELF_URL，存檔
 *  4. 執行 srSetup()（第一次會要求授權）→ 兩個檔各建 fact_sku_req，W1＝網址
 *  5. 執行 aaaSrCheck()，Log 全部 ✅ 才算完成
 */
var SR = {
  VER: '1.1',
  BOM_ID: '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao',
  DASH_ID: '1FF7lW3JINR0-Id7MMYSRkoRzA1BYdqktO94NbzAFYG0',
  TAB: 'fact_sku_req',
  TOKEN: 'dbc-skureq-Hv6pL3',
  SELF_URL: 'https://script.google.com/macros/s/AKfycbxal-Xme2BKxE8Fi5Xdh3QVPIqc9-s6X_RBrnIOYXUqnyPtpoN0aifjRee7GGUc0K6uAQ/exec',   // ← 部署後貼 /exec 網址
  MAX_ROWS: 5000,
  HEAD: ['申請單號','申請時間','門市','申請人','品名','類別','用途','用途說明','頻率','其他店',
         '廠商','使用單位','採購單位','每採購單位內容量','單價','標配或週用量','需要日期','參考連結',
         '狀態','處理結果sku_id','處理說明','處理時間'],
  NUM: {'每採購單位內容量':1,'單價':1,'標配或週用量':1},
  CATS: ['食材','耗材','器具','模具'],
  USES: ['做甜點／半成品','店務使用'],
  FREQS: ['長期固定使用','一次性／短期'],
  DECIDE: ['已核准','臨時採購','退回']
};

/* ---------- 入口 ---------- */
function doGet() {
  return srOut_({ ok: true, api: 'skureq', ver: SR.VER });
}

function doPost(e) {
  var out;
  try {
    var raw = (e && e.postData && e.postData.contents) || '';
    if (raw.length > 8000) return srOut_({ ok: false, err: 'too_large' });
    var p = JSON.parse(raw || '{}');
    if (p.token !== SR.TOKEN) return srOut_({ ok: false, err: 'token' });
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) return srOut_({ ok: false, err: 'busy' });
    try {
      if (p.action === 'submit') out = srSubmit_(p.req || {});
      else if (p.action === 'decide') out = srDecide_(p);
      else out = { ok: false, err: 'action' };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    out = { ok: false, err: String((err && err.message) || err) };
  }
  srLog_(out);
  return srOut_(out);
}

/* ---------- 送出申請（店長） ---------- */
function srSubmit_(r) {
  var s = function (v, max) { return srText_(v, max); };
  var id = s(r.id, 30);
  if (!/^R\d{2}-\d{10}-[A-Z0-9]{2}$/.test(id)) return { ok: false, err: 'id' };
  var st = s(r.st, 3);
  if (!/^\d{1,2}$/.test(st)) return { ok: false, err: 'store' };
  var req = {
    who: s(r.who, 30), name: s(r.name, 40), cat: s(r.cat, 4), use: s(r.use, 10), useNote: s(r.useNote, 100),
    freq: s(r.freq, 10), others: s(r.others, 40), vendor: s(r.vendor, 30), useUnit: s(r.useUnit, 10),
    buyUnit: s(r.buyUnit, 10), need: s(r.need, 10), link: s(r.link, 300)
  };
  var need = ['who','name','cat','vendor','useUnit','buyUnit'];   /* v1.1：用途／用在哪裡／頻率改選填 */
  for (var i = 0; i < need.length; i++) if (!req[need[i]]) return { ok: false, err: 'missing_' + need[i] };
  if (SR.CATS.indexOf(req.cat) < 0) return { ok: false, err: 'cat' };
  if (req.use && SR.USES.indexOf(req.use) < 0) return { ok: false, err: 'use' };
  if (req.freq && SR.FREQS.indexOf(req.freq) < 0) return { ok: false, err: 'freq' };
  var pq = srNum_(r.pq), price = srNum_(r.price), qty = srNum_(r.qty);
  if (r.qty === '' || r.qty == null) qty = 0;   /* v1.1：標配數量改選填（沒填＝0，核准後各店自己到店別參數設） */
  if (!(pq > 0) || !(price > 0) || !(qty >= 0)) return { ok: false, err: 'number' };
  if (req.need && !/^\d{4}-\d{2}-\d{2}$/.test(req.need)) req.need = '';
  if (req.link && !/^https?:\/\//i.test(req.link)) req.link = '';
  req.others = req.others.split(/[,，、;；\s]+/).filter(function (x) { return /^\d{1,2}$/.test(x); }).join(',');

  var sheets = srSheets_();
  var bom = sheets[0];
  var ids = srColVals_(bom, 1);
  if (ids.indexOf(id) >= 0) return { ok: true, dup: true, id: id };   // 前端重送＝冪等
  if (ids.length >= SR.MAX_ROWS) return { ok: false, err: 'full' };

  var row = [id, srNow_(), st, req.who, req.name, req.cat, req.use, req.useNote, req.freq, req.others,
             req.vendor, req.useUnit, req.buyUnit, pq, price, qty, req.need, req.link, '待審', '', '', ''];
  sheets.forEach(function (sh) { srAppend_(sh, row); });
  return { ok: true, id: id };
}

/* ---------- 處理結果（總部）：只改 待審 的列 ---------- */
function srDecide_(p) {
  var ids = (Array.isArray(p.ids) ? p.ids : []).map(function (x) { return srText_(x, 30); })
    .filter(function (x) { return /^R\d{2}-\d{10}-[A-Z0-9]{2}$/.test(x); }).slice(0, 50);
  var status = srText_(p.status, 10);
  var sku = srText_(p.sku, 12);
  var note = srText_(p.note, 200);
  if (!ids.length) return { ok: false, err: 'ids' };
  if (SR.DECIDE.indexOf(status) < 0) return { ok: false, err: 'status' };
  if (status === '已核准' && !/^[A-Za-z]{1,3}\d{1,5}$/.test(sku)) return { ok: false, err: 'sku' };
  if (status !== '已核准') sku = '';

  var now = srNow_(), res = [];
  srSheets_().forEach(function (sh) {
    var col = srColVals_(sh, 1), n = 0, skip = 0;
    var cStat = SR.HEAD.indexOf('狀態') + 1;
    ids.forEach(function (id) {
      var i = col.indexOf(id);
      if (i < 0) return;
      var rowNo = i + 2;
      var cur = String(sh.getRange(rowNo, cStat).getValue() || '').trim();
      if (cur && cur !== '待審') { skip++; return; }   // 已處理過的不覆蓋
      sh.getRange(rowNo, cStat, 1, 4).setValues([[status, sku, note, now]]);
      n++;
    });
    res.push(n + (skip ? '（略過 ' + skip + '）' : ''));
  });
  return { ok: true, updated: res };
}

/* ---------- 建表（部署後執行一次） ---------- */
function srSetup() {
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^\/]+\/exec$/.test(SR.SELF_URL))
    throw new Error('請先把網頁應用程式的 /exec 網址貼到 SR.SELF_URL');
  [SR.BOM_ID, SR.DASH_ID].forEach(function (fid) {
    var ss = SpreadsheetApp.openById(fid);
    var sh = ss.getSheetByName(SR.TAB);
    if (!sh) {
      sh = ss.insertSheet(SR.TAB);
    } else {
      var h = sh.getRange(1, 1, 1, SR.HEAD.length).getValues()[0].map(function (x) { return String(x).trim(); });
      var empty = h.every(function (x) { return !x; });
      if (!empty && h.join('|') !== SR.HEAD.join('|'))
        throw new Error(ss.getName() + ' 的 ' + SR.TAB + ' 表頭與程式不同，已停止（不覆蓋）');
    }
    var need = SR.HEAD.length + 1;
    if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
    if (sh.getMaxRows() < 200) sh.insertRowsAfter(sh.getMaxRows(), 200 - sh.getMaxRows());
    srFormat_(sh);
    sh.getRange(1, 1, 1, SR.HEAD.length).setValues([SR.HEAD]).setFontWeight('bold');
    sh.getRange(1, need).setNumberFormat('@').setValue(SR.SELF_URL);
    sh.setFrozenRows(1);
    Logger.log('✅ ' + ss.getName() + '：' + SR.TAB + ' 就緒（W1＝網址）');
  });
}

/* ---------- 唯讀檢查 ---------- */
function aaaSrCheck() {
  var ok = true, data = [];
  [SR.BOM_ID, SR.DASH_ID].forEach(function (fid) {
    var ss = SpreadsheetApp.openById(fid), sh = ss.getSheetByName(SR.TAB);
    if (!sh) { Logger.log('❌ ' + ss.getName() + ' 沒有 ' + SR.TAB); ok = false; data.push(null); return; }
    var head = sh.getRange(1, 1, 1, SR.HEAD.length + 1).getValues()[0];
    var hOk = head.slice(0, SR.HEAD.length).join('|') === SR.HEAD.join('|');
    var uOk = String(head[SR.HEAD.length]).trim() === SR.SELF_URL && !!SR.SELF_URL;
    Logger.log((hOk ? '✅' : '❌') + ' ' + ss.getName() + ' 表頭');
    Logger.log((uOk ? '✅' : '❌') + ' ' + ss.getName() + ' W1 網址' + (uOk ? '' : '（現值：' + head[SR.HEAD.length] + '）'));
    ok = ok && hOk && uOk;
    var last = sh.getLastRow();
    data.push(last > 1 ? sh.getRange(2, 1, last - 1, SR.HEAD.length).getDisplayValues().map(function (r) { return r.join('|'); }) : []);
  });
  if (data[0] && data[1]) {
    var same = data[0].length === data[1].length && data[0].every(function (r, i) { return r === data[1][i]; });
    Logger.log((same ? '✅' : '❌') + ' 兩檔內容一致（' + data[0].length + ' 筆 vs ' + data[1].length + ' 筆）');
    ok = ok && same;
  }
  Logger.log(ok ? '✅ 全部通過' : '❌ 有問題，見上方');
  return ok;
}

/* ---------- 工具 ---------- */
function srSheets_() {
  return [SR.BOM_ID, SR.DASH_ID].map(function (fid) {
    var sh = SpreadsheetApp.openById(fid).getSheetByName(SR.TAB);
    if (!sh) throw new Error('setup');
    return sh;
  });
}
function srFormat_(sh) {
  var rows = sh.getMaxRows();
  SR.HEAD.forEach(function (h, i) {
    sh.getRange(1, i + 1, rows, 1).setNumberFormat(SR.NUM[h] ? '0.###' : '@');
  });
}
function srAppend_(sh, row) {
  var r = Math.max(sh.getLastRow(), 1) + 1;
  if (r > sh.getMaxRows()) {
    sh.insertRowsAfter(sh.getMaxRows(), 200);
    srFormat_(sh);
  }
  var fmts = [SR.HEAD.map(function (h) { return SR.NUM[h] ? '0.###' : '@'; })];
  sh.getRange(r, 1, 1, row.length).setNumberFormats(fmts).setValues([row]);
}
function srColVals_(sh, col) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, col, last - 1, 1).getValues().map(function (r) { return String(r[0]).trim(); });
}
function srText_(v, max) {
  var t = String(v == null ? '' : v).replace(/[\r\n\t]+/g, ' ').trim().slice(0, max);
  if (/^[=+\-@]/.test(t)) t = "'" + t;   // 防公式注入
  return t;
}
function srNum_(v) { var n = Number(v); return isFinite(n) ? Math.round(n * 1000) / 1000 : NaN; }
function srNow_() { return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'); }
function srOut_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function srLog_(o) { try { console.log(JSON.stringify(o)); } catch (e) {} }