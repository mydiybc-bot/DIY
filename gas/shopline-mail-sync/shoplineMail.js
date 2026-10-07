/* =====================================================================
 * diybc-shopline-mail-sync v1.4（2026-10-07）
 * Shopline 門市叫貨單 → 採購系統月檔 fact_shopline（每天自動，雲端執行）
 *
 * 為什麼用 Email：Shopline 沒開 Open API、後台要登入；但每張訂單都會寄「[新訂單]」通知到 mydiybc@gmail.com，
 *   通知信裡有訂單號碼、日期、收件人（門市「N店」）、每個品項的數量與單價。
 * 權限：Gmail 只用「唯讀」（Gmail 進階服務 Gmail.Users.Messages.list/get），不能寄信／刪信／改信。
 * 規則（與採購大平台「上傳出貨中心訂單報表」同口徑）：
 *   店號＝收件人／訂購人的「N店」（零售客人沒有 N店 → 不寫）；日期＝訂單日期（台北）；訂單號碼前加「#」；
 *   品項＝商品名＋「｜」＋規格（用 fact_shopline 歷史寫法、其次 Shopline 商品目錄還原）；分類＝食材／器模具（其他＝未分類）。
 * 只新增、不覆蓋：fact_shopline 已有的訂單（含 A/B 拆單子單）一律略過——
 *   通知信是「下單當下」的內容；用後台 API 補進來的訂單已是最新狀態，不能被舊內容蓋掉。
 *   ⚠️ 已知限制（2026-10-07 實測 9/19～10/5）：①出貨中心下單後才加的品項、改的數量不會反映（Shopline 改單不寄信；
 *      34 張有信訂單中 25 張被改過，漏約 15% 金額）②總部在後台代建的訂單（大貨到店）不寄信，完全抓不到（45 張中 11 張）。
 *      → 11/4 寄提醒信，請 Claude 用後台 API 整批補最終版本（會覆蓋本程式寫入的訂單）。
 *   取消：Shopline 60 天內沒有寄任何「取消」信，取消處理實際上不會觸發，保留備用。
 * 取消：主旨含「#單號…已取消」的通知信。
 *   - 新增時：單號（或其母單）已取消 → 不新增。
 *   - 表上已有：只刪「訂單號碼完全相同」的列（取消 #…B 不會動到 #… 或 #…A）；
 *     母單取消但表上有它的 A/B 子單 → 不自動刪子單，記 log 並寄信。
 *   - 刪除用「由下往上刪列」，不整表重寫；刪前把被刪的列記進 log。
 * 寫入前重新讀一次表，縮短與「採購大平台手動上傳」撞在一起的時間。
 * 每次執行寫一列到月檔 shopline_sync_log（只記數字、單號、品名，不記顧客資料）。
 * 11/5 之後自動刪除自己的排程（Shopline 11/3 到期）。
 * ===================================================================== */
var SLM = {
  VER: 'v1.4',
  MONTH_SS: '19tG-AMYtiZPTYQcz85TGiGS2G4M5U1gZw9Mab2UFFU0',
  SHEET: 'fact_shopline',
  LOG: 'shopline_sync_log',
  NEW_Q: 'from:orders@shopline.com subject:新訂單',
  CANCEL_Q: 'from:orders@shopline.com subject:已取消',
  DAYS: 21,          // 新訂單信往回看幾天（只新增不覆蓋，重看無害）
  CANCEL_DAYS: 60,   // 取消信看更久，保證「新訂單信還在範圍內」時取消信一定也在
  STOP_AFTER: '2026-11-05',
  FINAL_REMIND: '2026-11-04',   // 這天寄一次提醒：請 Claude 用後台做最後補單
  HOUR: 6,
  NOTIFY: 'mydiybc@gmail.com',
  CATS: {'食材': 1, '器模具': 1},
  HDR: ['店號', '日期', '訂單號碼', '品項', '分類', '數量', '單價']
};

/* ---------- 對外函式 ---------- */
function slmDaily() {
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  if (today > SLM.STOP_AFTER) { slmRemoveTriggers_(); slmLog_({mode: 'stop', note: 'Shopline 已到期，自動停用排程'}); return; }
  var st = slmRun_({mode: 'daily', write: true});
  if (today === SLM.FINAL_REMIND) {
    try {
      MailApp.sendEmail(SLM.NOTIFY, '提醒：Shopline 到期前請叫 Claude 做最後一次補單',
        'Shopline 門市訂單每天是靠「新訂單」通知信自動補進採購系統，但出貨中心事後加的品項、改的數量，' +
        '以及總部在後台代建的訂單（例如大貨到店）都不會寄信，所以沒進來。\n\n' +
        '請在 Shopline 到期前跟 Claude 說：「Shopline 最後補單」。Claude 會用後台把 10/6 之後的門市訂單整批換成最終版本。');
    } catch (e) {}
  }
  return st;
}
function slmDryRun() { return slmRun_({mode: 'dryrun', write: false}); }   // 只算不寫
function slmRunNow() { return slmRun_({mode: 'manual', write: true}); }    // 手動補跑（寫入）
function slmSetup() {
  slmRemoveTriggers_();
  ScriptApp.newTrigger('slmDaily').timeBased().everyDays(1).atHour(SLM.HOUR).inTimezone('Asia/Taipei').create();
  slmLog_({mode: 'setup', note: '已裝每日 ' + SLM.HOUR + ' 點排程，' + SLM.STOP_AFTER + ' 後自動停'});
}
function slmRemoveTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'slmDaily') ScriptApp.deleteTrigger(t); });
}

/* ---------- Gmail（唯讀，進階服務） ---------- */
function slmListMsgs_(q, max) {
  var ids = [], tok = null;
  do {
    var r = Gmail.Users.Messages.list('me', {q: q, maxResults: 100, pageToken: tok});
    (r.messages || []).forEach(function (m) { ids.push(m.id); });
    tok = r.nextPageToken;
  } while (tok && ids.length < (max || 1000));
  return ids;
}
function slmGetMsg_(id) {
  var m = Gmail.Users.Messages.get('me', id, {format: 'full'});
  var hs = (m.payload && m.payload.headers) || [], h = {};
  hs.forEach(function (x) { h[String(x.name).toLowerCase()] = x.value; });
  var html = '';
  (function walk(p) {
    if (!p || html) return;
    if (/^text\/html/i.test(p.mimeType || '') && p.body && p.body.data) { html = slmB64_(p.body.data); return; }
    (p.parts || []).forEach(walk);
  })(m.payload);
  if (!html) (function walk2(p) {
    if (!p || html) return;
    if (/^text\/plain/i.test(p.mimeType || '') && p.body && p.body.data) { html = slmB64_(p.body.data); return; }
    (p.parts || []).forEach(walk2);
  })(m.payload);
  return {id: id, from: h['from'] || '', subject: h['subject'] || '', time: Number(m.internalDate || 0), html: html};
}
function slmB64_(d) {
  // Apps Script 的 Gmail 進階服務會把 body.data 直接給成「位元組陣列」；字串時才是 base64url
  if (d && typeof d !== 'string') return Utilities.newBlob(d).getDataAsString('UTF-8');
  var s = String(d || '').replace(/-/g, '+').replace(/_/g, '/');
  while (s.length % 4) s += '=';
  return Utilities.newBlob(Utilities.base64Decode(s)).getDataAsString('UTF-8');
}

/* 取消信 → {單號(含尾碼): 1} */
function slmCancelSet_(st) {
  var set = {};
  slmListMsgs_(SLM.CANCEL_Q + ' newer_than:' + SLM.CANCEL_DAYS + 'd', 1000).forEach(function (id) {
    var m = slmGetMsg_(id);
    if (m.from.indexOf('orders@shopline.com') < 0) return;
    var mm = m.subject.match(/#(\d{10,}[A-Z]*)[^#]*已取消/);
    if (mm) { set[mm[1]] = 1; if (st) st.cancelMails++; }
  });
  return set;
}

/* ---------- 主流程 ---------- */
function slmRun_(opt) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { slmLog_({mode: opt.mode, note: '上一次還在跑，略過'}); return; }
  var st = {mode: opt.mode, mails: 0, store: 0, retail: 0, skipped: 0, cancelled: 0, newOrders: 0, newRows: 0, removed: 0,
            cancelMails: 0, unknown: [], fail: [], warn: [], removedRows: [], note: ''};
  var writeErr = false;
  try {
    var ss = SpreadsheetApp.openById(SLM.MONTH_SS), sh = ss.getSheetByName(SLM.SHEET);
    if (!sh) throw new Error('找不到分頁 ' + SLM.SHEET);
    slmCheckHeader_(sh);
    var vals0 = sh.getDataRange().getValues();
    var dict = slmBuildDict_(vals0, null);

    // 1) 先讀取消信
    var cancel = slmCancelSet_(st);

    // 2) 新訂單信（同一單號只取最早一封）
    var msgs = slmListMsgs_(SLM.NEW_Q + ' newer_than:' + SLM.DAYS + 'd', 1000).map(slmGetMsg_)
      .filter(function (m) { return m.from.indexOf('orders@shopline.com') >= 0 && m.subject.indexOf('新訂單') >= 0; })
      .sort(function (a, b) { return a.time - b.time; });
    var cand = {}, order = [];
    msgs.forEach(function (m) {
      st.mails++;
      var o;
      try { o = slmParseOrder_(m.html, m.subject); } catch (e) { st.fail.push('解析錯誤：' + String(e && e.message || e).slice(0, 80)); return; }
      if (!o.ord) { st.fail.push('讀不到訂單號碼'); return; }
      if (!o.store) { st.retail++; return; }
      st.store++;
      if (cand[o.ord]) return;
      if (!o.items.length) { st.fail.push('#' + o.ord + ' 讀不到品項'); return; }
      if (!o.date) { st.fail.push('#' + o.ord + ' 讀不到日期'); return; }
      cand[o.ord] = o; order.push(o.ord);
    });

    // 3) 寫入前重新讀表（縮短與手動上傳撞車的時間）
    var vals = sh.getDataRange().getValues();
    var have = slmHaveBases_(vals), haveExact = slmHaveExact_(vals);
    var add = [];
    order.forEach(function (ord) {
      var o = cand[ord], base = ord.replace(/[A-Z]+$/, '');
      if (cancel[ord] || cancel[base]) { st.cancelled++; return; }
      if (have[base]) { st.skipped++; return; }
      st.newOrders++;
      o.items.forEach(function (it) {
        var c = slmCanon_(it.parts, dict);
        if (c.unknown) st.unknown.push(c.name);
        add.push([Number(o.store), slmDate_(o.date), '#' + ord, c.name, c.cat, it.qty, it.price]);
      });
    });
    st.newRows = add.length;

    // 4) 表上已有、但已取消 → 只刪完全相同的單號
    var delIdx = [];
    Object.keys(cancel).forEach(function (c) {
      var rows = haveExact['#' + c] || [];
      rows.forEach(function (i) { delIdx.push(i); });
      if (!/[A-Z]$/.test(c)) {
        var kids = Object.keys(haveExact).filter(function (k) { return k.indexOf('#' + c) === 0 && k !== '#' + c; });
        if (kids.length) st.warn.push('母單 #' + c + ' 已取消，但表上還有子單 ' + kids.join('、') + '（未自動刪，請確認）');
      }
    });
    delIdx.sort(function (a, b) { return b - a; });

    if (opt.write) {
      try {
        if (delIdx.length) {
          var before = sh.getLastRow();
          delIdx.forEach(function (i) {   // i 是 vals 的索引（0＝表頭）→ 列號 i+1；由下往上刪
            var r = vals[i];
            st.removedRows.push([r[0], slmFmt_(r[1]), r[2], String(r[3]).slice(0, 40), r[5], r[6]]);
          });
          slmDeleteRows_(sh, delIdx.map(function (i) { return i + 1; }));
          st.removed = delIdx.length;
          if (sh.getLastRow() !== before - delIdx.length) st.warn.push('刪列後列數不對：刪前 ' + before + '、刪 ' + delIdx.length + '、刪後 ' + sh.getLastRow());
        }
        if (add.length) {
          var start = sh.getLastRow() + 1, need = start + add.length - 1 - sh.getMaxRows();
          if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need);
          sh.getRange(start, 1, add.length, SLM.HDR.length).setValues(add);
        }
      } catch (we) { writeErr = true; throw we; }
    } else {
      st.removed = delIdx.length + ' 列(試跑未刪)';
    }
    st.note = (add.length ? '新增 ' + st.newOrders + ' 張（' + slmDates_(add) + '）' : '沒有新的門市訂單') + (st.removed ? '；取消移除 ' + st.removed : '');
  } catch (e) {
    st.note = (writeErr ? '寫入失敗：' : '錯誤：') + (e && e.message || e);
    st.fail.push(st.note);
  } finally {
    lock.releaseLock();
  }
  slmLog_(st);
  if ((st.fail.length || st.warn.length) && opt.mode !== 'dryrun') slmMail_(st, writeErr);
  return st;
}

function slmCheckHeader_(sh) {
  var h = sh.getRange(1, 1, 1, SLM.HDR.length).getValues()[0].map(function (x) { return String(x).trim(); });
  if (h.join('|') !== SLM.HDR.join('|')) throw new Error('fact_shopline 表頭跟預期不同（' + h.join('、') + '），停止寫入');
}

/* 刪多列：把相連的列合併成一次 deleteRows，由下往上 */
function slmDeleteRows_(sh, rowNums) {
  var rs = rowNums.slice().sort(function (a, b) { return b - a; }), i = 0;
  while (i < rs.length) {
    var hi = rs[i], lo = hi;
    while (i + 1 < rs.length && rs[i + 1] === lo - 1) { i++; lo = rs[i]; }
    sh.deleteRows(lo, hi - lo + 1);
    i++;
  }
}

/* ---------- 解析一封「[新訂單]」通知信 ---------- */
function slmParseOrder_(html, subject) {
  var lines = slmHtmlToLines_(html);
  var text = lines.join('\n');
  var o = {ord: '', date: '', store: '', items: []};
  var m = (subject || '').match(/#(\d{10,}[A-Z]*)/) || text.match(/訂單號碼\s*(\d{10,}[A-Z]*)/);
  if (m) o.ord = m[1];
  m = text.match(/訂單日期\s*(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) o.date = m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
  o.store = slmStoreOf_(text);
  var a = -1, b = lines.length;
  for (var i = 0; i < lines.length; i++) {
    if (a < 0 && /訂單詳情/.test(lines[i])) a = i;
    else if (a >= 0 && /^小計\s*[:：]?/.test(lines[i])) { b = i; break; }
  }
  if (a < 0) return o;
  var pending = [];
  var QTY = /^(?:(.*?)\s+)?(\d+)\s*[x×X]\s*NT\$\s*([\d,]+(?:\.\d+)?)$/;
  for (var j = a; j < b; j++) {
    var L = lines[j];
    if (j === a) { L = L.replace(/^[\s\S]*?訂單詳情\s*/, ''); }
    L = L.trim();
    if (!L) continue;
    var q = L.match(QTY);
    if (q) {
      var parts = pending.slice();
      if (q[1] && q[1].trim()) parts.push(q[1].trim());
      pending = [];
      if (!parts.length) continue;
      o.items.push({parts: parts, qty: Number(q[2]), price: Number(q[3].replace(/,/g, ''))});
      continue;
    }
    if (/^-?NT\$\s*[\d,]+(?:\.\d+)?$/.test(L)) { pending = []; continue; }   // 這一筆的小計金額＝品項結束
    pending.push(L);
  }
  return o;
}

/* 店號：先看收件人（地址那行開頭），再看訂購人；只收 1–20 */
function slmStoreOf_(text) {
  var re = /(?:^|[^\d])(\d{1,2})\s*[號号]?\s*店/;
  var pick = function (s) { var k = s.match(re); if (!k) return ''; var n = +k[1]; return (n >= 1 && n <= 20) ? String(n) : ''; };
  var m = text.match(/地址\s*([^\n]*)/);
  if (m) { var who = m[1].split(/\s+台灣|\s+\d{3}\s/)[0]; var s1 = pick(who); if (s1) return s1; }
  m = text.match(/訂購人\s*([^\n]*)/);
  if (m) { var s2 = pick(m[1].replace(/\(.*?@.*?\)/g, '')); if (s2) return s2; }
  return '';
}

function slmHtmlToLines_(html) {
  // Shopline 通知信版面（2026-10-07 實測）：品名一個 <p>、規格每一欄各一個 <p>、「1 x\n NT$100」在同一個 <p> 裡換行；
  // 還有 <!-- … --> 註解（裡面有「x =」）。所以：先去註解 → 把所有空白（含換行）壓成一格 → 只在區塊標籤斷行。
  var s = String(html || '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<(style|script|head)[\s\S]*?<\/\1>/gi, '')
    .replace(/[\r\n\t ]+/g, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/?(p|div|tr|td|th|li|ul|ol|table|tbody|thead|h[1-6])(\s[^>]*)?>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ').replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"').replace(/&#39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(+n); })
    .replace(/&#x([0-9a-f]+);/gi, function (_, n) { return String.fromCharCode(parseInt(n, 16)); })
    .replace(/&amp;/gi, '&');
  return s.split('\n').map(function (x) { return x.replace(/[ \t\u00a0\u3000]+/g, ' ').trim(); }).filter(function (x) { return x; });
}

/* ---------- 品名還原 ---------- */
function slmKey_(s) { return String(s || '').replace(/[\s｜|,，]+/g, ''); }
function slmCatOk_(c) { return SLM.CATS[c] ? c : ''; }

/* vals：fact_shopline 全表；beforeDate（'yyyy-MM-dd'）有給時只用這天以前的列（驗證用） */
function slmBuildDict_(vals, beforeDate) {
  var cnt = {};
  for (var i = 1; i < vals.length; i++) {
    if (beforeDate && slmFmt_(vals[i][1]) >= beforeDate) continue;
    var nm = String(vals[i][3] || ''); if (!nm) continue;
    var k = slmKey_(nm), id = nm + '\u0001' + String(vals[i][4] || '');
    cnt[k] = cnt[k] || {}; cnt[k][id] = (cnt[k][id] || 0) + 1;
  }
  var hist = {};
  Object.keys(cnt).forEach(function (k) {
    var best = null, bn = -1;
    Object.keys(cnt[k]).forEach(function (id) { if (cnt[k][id] > bn) { bn = cnt[k][id]; best = id; } });
    var p = best.split('\u0001'); hist[k] = [p[0], p[1]];
  });
  return {hist: hist, cat: (typeof SLM_CATALOG !== 'undefined') ? SLM_CATALOG : {}};
}

function slmCanon_(parts, dict) {
  var k = slmKey_(parts.join(' '));
  var h = dict.hist[k], c = dict.cat[k];
  var catOf = function () { return slmCatOk_(h && h[1]) || slmCatOk_(c && c[1]) || '未分類'; };
  if (h) return {name: h[0], cat: catOf()};
  if (c) return {name: c[0], cat: catOf()};
  var name = parts.length > 1 ? parts[0] + '｜' + parts.slice(1).join(' ') : parts[0];
  return {name: name, cat: '未分類', unknown: true};
}

/* ---------- 小工具 ---------- */
function slmHaveBases_(vals) {
  var have = {};
  for (var i = 1; i < vals.length; i++) { var b = String(vals[i][2] || '').replace(/^#/, '').replace(/[A-Z]+$/, ''); if (b) have[b] = 1; }
  return have;
}
function slmHaveExact_(vals) {
  var m = {};
  for (var i = 1; i < vals.length; i++) { var o = String(vals[i][2] || '').trim(); if (o) (m[o] = m[o] || []).push(i); }
  return m;
}
function slmDate_(s) { var m = s.split('-'); return new Date(+m[0], +m[1] - 1, +m[2]); }
function slmFmt_(d) { return d instanceof Date ? Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd') : String(d || '').slice(0, 10); }
function slmDates_(rows) {
  var mn = null, mx = null;
  rows.forEach(function (r) { var d = r[1]; if (!mn || d < mn) mn = d; if (!mx || d > mx) mx = d; });
  var f = function (d) { return Utilities.formatDate(d, 'Asia/Taipei', 'MM/dd'); };
  return f(mn) + '～' + f(mx);
}

function slmLog_(st) {
  try {
    var ss = SpreadsheetApp.openById(SLM.MONTH_SS);
    var sh = ss.getSheetByName(SLM.LOG);
    if (!sh) {
      sh = ss.insertSheet(SLM.LOG);
      sh.getRange(1, 1, 1, 13).setValues([['時間', '模式', '版本', '新訂單信', '門市單', '零售單', '已存在略過', '已取消略過', '新增訂單', '新增列', '取消移除', '說明', '明細']]);
      try { if (sh.getMaxColumns() > 13) sh.deleteColumns(14, sh.getMaxColumns() - 13); } catch (e1) {}
    }
    var detail = JSON.stringify({unknown: (st.unknown || []).slice(0, 40), fail: (st.fail || []).slice(0, 40), warn: (st.warn || []).slice(0, 20),
      cancelMails: st.cancelMails || 0, removedRows: (st.removedRows || []).slice(0, 50), check: st.check || null});
    sh.insertRowBefore(2);
    sh.getRange(2, 1, 1, 13).setValues([[new Date(), st.mode || '', SLM.VER, st.mails || 0, st.store || 0, st.retail || 0, st.skipped || 0,
      st.cancelled || 0, st.newOrders || 0, st.newRows || 0, String(st.removed || 0), st.note || '', detail.slice(0, 45000)]]);
    var n = sh.getLastRow(); if (n > 300) sh.deleteRows(301, n - 300);
  } catch (e) { Logger.log('寫 log 失敗：' + e); }
}

function slmMail_(st, writeErr) {
  try {
    var subj = writeErr ? '❌ Shopline 門市訂單自動同步：寫入失敗' : '⚠️ Shopline 門市訂單自動同步：需要確認';
    MailApp.sendEmail(SLM.NOTIFY, subj,
      'Shopline → 採購系統 fact_shopline 今天的自動同步：\n\n' + (st.fail || []).concat(st.warn || []).join('\n') +
      '\n\n新增訂單 ' + (st.newOrders || 0) + ' 張、取消移除 ' + (st.removed || 0) + ' 列。\n明細在月檔「shopline_sync_log」分頁。請把這封信轉給 Claude 處理。');
  } catch (e) { Logger.log('寄信失敗：' + e); }
}

/* ---------- 上線前驗證（只讀不寫） ----------
 * fact_shopline 中 FROM～TO 的門市訂單（母單＋A/B 子單合併）＝標準答案（後台 API 補進來的最新狀態）。
 * 用同一套每日流程的規則（同單號取最早一封信、品名字典只用 FROM 以前的資料）解析通知信，逐品項比數量、單價、分類。
 * 另列「信件判定為門市單、表上卻沒有」的單（可能誤判）。結果寫進 shopline_sync_log（mode=check）。 */
function slmCheck() {
  var FROM = '2026-09-19', TO = '2026-10-05';
  var vals = SpreadsheetApp.openById(SLM.MONTH_SS).getSheetByName(SLM.SHEET).getDataRange().getValues();
  var dict = slmBuildDict_(vals, FROM), truth = {}, tcat = {}, tstore = {}, tdate = {};
  for (var i = 1; i < vals.length; i++) {
    var ds = slmFmt_(vals[i][1]); if (ds < FROM || ds > TO) continue;
    var b = String(vals[i][2]).replace(/^#/, '').replace(/[A-Z]+$/, '');
    var key = String(vals[i][3]) + ' @' + Number(vals[i][6]);
    truth[b] = truth[b] || {}; truth[b][key] = (truth[b][key] || 0) + Number(vals[i][5]);
    tcat[String(vals[i][3])] = String(vals[i][4]);
    tstore[b] = String(vals[i][0]); tdate[b] = tdate[b] && tdate[b] < ds ? tdate[b] : ds;
  }
  var msgs = slmListMsgs_(SLM.NEW_Q + ' after:2026/09/17 before:2026/10/08', 1000).map(slmGetMsg_)
    .filter(function (m) { return m.from.indexOf('orders@shopline.com') >= 0 && m.subject.indexOf('新訂單') >= 0; })
    .sort(function (a, b) { return a.time - b.time; });
  var got = {}, gstore = {}, gdate = {}, gcat = {}, unknown = {}, fail = [], extra = [], viaCatalog = 0, viaHist = 0;
  msgs.forEach(function (m) {
    var o; try { o = slmParseOrder_(m.html, m.subject); } catch (e) { fail.push(String(e)); return; }
    var b = String(o.ord).replace(/[A-Z]+$/, '');
    if (got[b]) return;
    if (!truth[b]) { if (o.store && o.date >= FROM && o.date <= TO) extra.push(o.ord + '(' + o.store + '店 ' + o.date + ')'); return; }
    gstore[b] = o.store; gdate[b] = o.date; got[b] = {};
    o.items.forEach(function (it) {
      var k = slmKey_(it.parts.join(' '));
      if (dict.hist[k]) viaHist++; else if (dict.cat[k]) viaCatalog++;
      var c = slmCanon_(it.parts, dict); if (c.unknown) unknown[c.name] = 1;
      gcat[c.name] = c.cat;
      var key = c.name + ' @' + it.price; got[b][key] = (got[b][key] || 0) + it.qty;
    });
  });
  var res = {orders: Object.keys(truth).length, mailFound: 0, same: 0, storeOk: 0, dateOk: 0, viaHist: viaHist, viaCatalog: viaCatalog,
    catDiff: [], diff: [], noMail: [], extra: extra};
  Object.keys(truth).forEach(function (b) {
    if (!got[b]) { res.noMail.push(b + '(' + tstore[b] + '店 ' + tdate[b] + ')'); return; }
    res.mailFound++;
    if (gstore[b] === tstore[b]) res.storeOk++;
    if (gdate[b] === tdate[b]) res.dateOk++;
    var t = truth[b], g = got[b], d = [], seen = {};
    Object.keys(t).concat(Object.keys(g)).forEach(function (k) {
      if (seen[k]) return; seen[k] = 1;
      if ((t[k] || 0) !== (g[k] || 0)) d.push({k: k.slice(0, 60), api: t[k] || 0, mail: g[k] || 0});
    });
    if (!d.length) res.same++; else res.diff.push({ord: b, store: tstore[b], d: d.slice(0, 10)});
  });
  Object.keys(gcat).forEach(function (n) { if (tcat[n] && tcat[n] !== gcat[n]) res.catDiff.push(n.slice(0, 50) + '：表上 ' + tcat[n] + '／信 ' + gcat[n]); });
  var cset = slmCancelSet_(null);
  res.cancelSample = Object.keys(cset).slice(0, 20);
  var st = {mode: 'check', note: '比對 ' + res.orders + ' 張：有信 ' + res.mailFound + '、完全相同 ' + res.same + '、店號對 ' + res.storeOk +
    '、日期對 ' + res.dateOk + '、表上沒有卻判為門市單 ' + extra.length + '；取消信抓到 ' + res.cancelSample.length + ' 張',
    unknown: Object.keys(unknown), fail: fail, check: res};
  slmLog_(st);
  return st;
}

/* ---------- 診斷：只看一封門市單「訂單詳情～小計」之間（只有品名、價格；長數字遮掉） ---------- */
function slmDiag() {
  var ids = slmListMsgs_(SLM.NEW_Q + ' newer_than:5d', 50), out = null;
  for (var i = 0; i < ids.length && !out; i++) {
    var m = slmGetMsg_(ids[i]);
    if (m.from.indexOf('orders@shopline.com') < 0) continue;
    var o = slmParseOrder_(m.html, m.subject);
    if (!o.store) continue;
    var html = m.html, a = html.indexOf('訂單詳情'), b = html.indexOf('小計', a);
    var seg = a >= 0 ? html.slice(a, b > a ? Math.min(b + 20, a + 12000) : a + 12000) : '(找不到 訂單詳情)';
    seg = seg.replace(/<img[^>]*>/gi, '<img>').replace(/\s(style|class|width|height|align|valign|bgcolor|cellpadding|cellspacing|border)="[^"]*"/gi, '').replace(/\d{8,}/g, '########');
    var lines = slmHtmlToLines_(html), li = -1;
    for (var k = 0; k < lines.length; k++) if (/訂單詳情/.test(lines[k])) { li = k; break; }
    out = {ord: o.ord, lineIdx: li, lines: (li >= 0 ? lines.slice(li, li + 40) : []).map(function (x) { return x.replace(/\d{8,}/g, '########').slice(0, 80); }),
      seg: seg.slice(0, 6000)};
  }
  slmLog_({mode: 'diag', note: out ? '診斷 #' + out.ord : '找不到門市單', check: out});
}

/* ---------- 刪掉指定訂單號碼（完全相同才刪；刪前把列記進 log；刪後核對列數） ---------- */
function slmRemoveOrders_(nums, why) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('鎖不到');
  var st = {mode: 'remove', removed: 0, removedRows: [], fail: [], warn: [], note: ''};
  try {
    var sh = SpreadsheetApp.openById(SLM.MONTH_SS).getSheetByName(SLM.SHEET);
    slmCheckHeader_(sh);
    var vals = sh.getDataRange().getValues(), want = {}, idx = [];
    nums.forEach(function (n) { want[String(n).trim()] = 1; });
    for (var i = 1; i < vals.length; i++) if (want[String(vals[i][2]).trim()]) {
      idx.push(i + 1);
      st.removedRows.push([vals[i][0], slmFmt_(vals[i][1]), vals[i][2], String(vals[i][3]).slice(0, 40), vals[i][5], vals[i][6]]);
    }
    var before = sh.getLastRow();
    if (idx.length) slmDeleteRows_(sh, idx);
    st.removed = idx.length;
    if (sh.getLastRow() !== before - idx.length) st.warn.push('刪後列數不對：刪前 ' + before + '、刪 ' + idx.length + '、刪後 ' + sh.getLastRow());
    st.note = (why || '') + '：刪除 ' + nums.join('、') + ' 共 ' + idx.length + ' 列';
  } catch (e) { st.fail.push(String(e && e.message || e)); st.note = '錯誤：' + st.fail[0]; }
  finally { lock.releaseLock(); }
  slmLog_(st);
  return st;
}
/* 2026-10-07 全面核對：7 號店 9/13 訂單在 Shopline 已取消，fact_shopline 還留 2 列 */
function slmFix20261007() { return slmRemoveOrders_(['#20260913103735151'], '全面核對（Shopline 已取消）'); }
