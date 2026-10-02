/*****************************************************************
 * repair_20261002.gs — POS 營收與肚肚對帳修復（一次性）
 * 2026-10-02｜經營者核可「照計畫全部修」
 *
 *   ① 先備份  ② 有問題的日期從肚肚重新匯入（試算表 POS資料 ＋ BigQuery）
 *   ③ 補 2024/12/1–12/19 明細與 2024/12 整月折扣  ④ 逐店逐月再和肚肚對一次
 *
 * 分工：本檔只做「需要試算表權限或肚肚登入」的事。
 *       比對、產生刪補計畫、寫 BigQuery 正式表，由 Claude 在本機用 SQL 做（可審、可重跑）。
 *
 * 執行順序（下拉選單由上往下）：
 *   rpr0_status          唯讀：印目前進度
 *   rpr1_snapshotBefore  POS資料 全表 → BQ rpr20261002_sheet_pre（＝試算表完整備份，含列號）
 *   rpr2_exports         逐日抓肚肚銷售明細（單日）→ BQ rpr20261002_exp
 *   rpr3_discounts       逐日抓肚肚 2024/12 專案折扣（單日）→ BQ rpr20261002_disc
 *   rpr4_applySheet      ★唯一會改 POS資料：依 BQ 計畫表刪列＋補列（成功後不能再跑）
 *   rpr5_snapshotAfter   改後全表 → BQ rpr20261002_sheet_post（核對用）
 *
 * 1/2/3/5 只寫 rpr20261002_* 暫存表，不碰正式資料；都可重跑、逾時會自動接續。
 * 本檔不改任何既有函式；肚肚登入沿用指令碼屬性裡既有的帳密（與每日匯入相同）。
 * 避開清晨 04:30–08:30 排程時段。
 *****************************************************************/

var RPR = {
  SS_ID: '1EyDihj4LPok_dvv3ZkAzDhsHqs7kDi5RTCXPF5Lt1ao',
  TAB: 'POS資料',
  PROJECT: 'diybc-make-sync',
  DS: 'diybc_pos',
  BUDGET_MS: 270000,   // 4.5 分鐘收工，留時間給 BigQuery 載入與記進度
  FETCH_BUDGET_MS: 180000, // 抓肚肚：3 分鐘後不再開新的一天（單日最慢約 1.5 分鐘）
  CHUNK: 25000,        // 快照每批列數
  ADD_CHUNK: 4000,     // 補列每批列數
  EXP_DAYS: [
    '2024-12-01','2024-12-02','2024-12-03','2024-12-04','2024-12-05','2024-12-06','2024-12-07',
    '2024-12-08','2024-12-09','2024-12-10','2024-12-11','2024-12-12','2024-12-13','2024-12-14',
    '2024-12-15','2024-12-16','2024-12-17','2024-12-18','2024-12-19',
    '2025-01-04','2025-03-09','2025-03-15','2025-03-22','2025-04-16','2025-04-23',
    '2025-05-05','2025-05-09','2025-05-10','2025-05-17','2025-05-18','2025-05-20','2025-05-21',
    '2025-05-22','2025-05-23','2025-05-27','2025-05-28','2025-06-14','2025-06-18','2025-07-05',
    '2025-07-15','2025-07-19','2025-07-24','2025-08-01','2025-08-02','2025-08-03','2025-08-06',
    '2025-08-13','2025-08-15','2025-08-19','2025-08-20','2025-08-21','2025-09-01','2025-09-06',
    '2025-09-14','2025-09-18','2025-09-20','2025-09-29','2025-10-02','2025-10-12','2025-10-13',
    '2025-10-14','2025-10-16','2025-10-17','2025-10-19','2025-10-20','2025-10-21','2025-10-23',
    '2025-10-25','2025-10-26','2025-10-27','2025-11-02','2025-11-03','2025-11-11','2025-11-14',
    '2025-11-17','2025-11-18','2025-11-24','2025-11-27','2025-11-28','2025-11-30','2025-12-03',
    '2025-12-11','2025-12-12','2025-12-16','2025-12-19','2025-12-31','2026-01-05','2026-07-29',
    '2026-08-08','2026-08-09','2026-08-10'
  ],
  DISC_FROM: '2024-12-01',
  DISC_TO: '2024-12-31'
};

/* ==================== 0. 唯讀：進度 ==================== */

function rpr0_status() {
  var p = PropertiesService.getScriptProperties().getProperties();
  Object.keys(p).filter(function (k) { return k.indexOf('RPR20261002') === 0; }).sort()
    .forEach(function (k) { Logger.log(k + ' = ' + String(p[k]).slice(0, 300)); });
  ['sheet_pre', 'exp', 'disc', 'plan_del', 'plan_add', 'plan_meta', 'sheet_post'].forEach(function (t) {
    try { Logger.log('rpr20261002_' + t + '：' + rprQuery_('SELECT COUNT(*) FROM ' + rprT_('rpr20261002_' + t))[0][0] + ' 列'); }
    catch (e) { Logger.log('rpr20261002_' + t + '：（尚未建立）'); }
  });
  var sh = SpreadsheetApp.openById(RPR.SS_ID).getSheetByName(RPR.TAB);
  Logger.log('POS資料 目前最後一列：' + sh.getLastRow());
}

/* ==================== 1／5. 全表快照 ==================== */

function rpr1_snapshotBefore() { return rprSnapshot_('rpr20261002_sheet_pre'); }
function rpr5_snapshotAfter()  { return rprSnapshot_('rpr20261002_sheet_post'); }

var RPR_SNAP_FIELDS = [['rownum', 'INTEGER'], ['a', 'STRING'], ['b', 'STRING'], ['c', 'STRING'], ['c_kind', 'STRING'],
  ['d', 'STRING'], ['e', 'STRING'], ['f', 'STRING'], ['g', 'STRING'], ['h', 'STRING'], ['i', 'STRING'], ['j', 'STRING'], ['k', 'STRING']];

function rprSnapshot_(table) {
  var props = PropertiesService.getScriptProperties();
  var key = 'RPR20261002_SNAP_' + table;
  var st = JSON.parse(props.getProperty(key) || 'null');
  var sh = SpreadsheetApp.openById(RPR.SS_ID).getSheetByName(RPR.TAB);
  var last = sh.getLastRow();
  if (st && st.done && st.last === last) { Logger.log(table + ' 已完成：' + st.rows + ' 列（試算表最後一列 ' + last + '），不重拍。'); return; }
  if (!st || st.last !== last || st.done) {
    if (st) Logger.log('試算表列數已變（' + st.last + ' → ' + last + '）或要重拍，從頭開始');
    st = { next: 2, last: last, rows: 0, done: false };
  }
  var t0 = Date.now();
  while (st.next <= last && Date.now() - t0 < RPR.BUDGET_MS) {
    var n = Math.min(RPR.CHUNK, last - st.next + 1);
    var vals = sh.getRange(st.next, 1, n, 11).getValues();
    var lines = new Array(n);
    for (var i = 0; i < n; i++) {
      var r = vals[i], c = '', kind = '';
      if (r[2]) {   // 與 syncSheetToBigQuery_batch_v2 相同的日期換算
        var d = r[2] instanceof Date ? r[2] : new Date(r[2]);
        kind = r[2] instanceof Date ? 'date' : 'text';
        c = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2);
      }
      lines[i] = [st.next + i, rprS_(r[0]), rprS_(r[1]), c, kind, rprS_(r[3]), rprS_(r[4]), rprS_(r[5]),
        rprS_(r[6]), rprS_(r[7]), rprS_(r[8]), rprS_(r[9]), rprS_(r[10])].map(rprCsv_).join(',');
    }
    var got = rprLoad_(table, RPR_SNAP_FIELDS, lines.join('\n'), st.next === 2);
    if (got !== n) throw new Error('快照載入列數不符：預期 ' + n + '、實際 ' + got);
    st.rows += n; st.next += n;
    props.setProperty(key, JSON.stringify(st));
    Logger.log(table + '：已拍到第 ' + (st.next - 1) + ' 列（累計 ' + st.rows + '）');
  }
  if (st.next > last) {
    st.done = true; props.setProperty(key, JSON.stringify(st));
    var bq = Number(rprQuery_('SELECT COUNT(*) FROM ' + rprT_(table))[0][0]);
    Logger.log((bq === last - 1 ? '✅ ' : '🔴 ') + table + ' 完成：BQ ' + bq + ' 列／試算表資料列 ' + (last - 1));
  } else {
    Logger.log('⏸ 時間到，請再執行一次 ' + (table.indexOf('post') > 0 ? 'rpr5_snapshotAfter' : 'rpr1_snapshotBefore') + ' 接續');
  }
}

/* ==================== 2. 肚肚銷售明細（單日） ==================== */

var RPR_EXP_FIELDS = [['day', 'STRING'], ['seq', 'INTEGER'], ['c0', 'STRING'], ['c1', 'STRING'], ['c2', 'STRING'],
  ['c3', 'STRING'], ['c4', 'STRING'], ['c11', 'STRING'], ['c18', 'STRING'], ['c19', 'STRING'], ['c20', 'STRING'],
  ['c21', 'STRING'], ['c22', 'STRING'], ['c23', 'STRING'], ['c24', 'STRING'], ['ncol', 'INTEGER']];

function rpr2_exports() {
  rprFetchDays_('EXP', 'rpr20261002_exp', RPR.EXP_DAYS, RPR_EXP_FIELDS, function (tok, day) {
    var payload = 'hierarchy_id=' + DUDOO_CONFIG_.HIERARCHY_ID + '&start_date=' + day + '&end_date=' + day;
    for (var i = 0; i < DUDOO_CONFIG_.COMPANY_IDS.length; i++) payload += '&company_id%5B%5D=' + DUDOO_CONFIG_.COMPANY_IDS[i];
    var resp = UrlFetchApp.fetch(DUDOO_CONFIG_.API_BASE + '/reports/getSaleDetailsAnalysis/export', {
      method: 'post', contentType: 'application/x-www-form-urlencoded', payload: payload,
      headers: { 'Access-Token': tok }, muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error(day + ' 銷售明細 HTTP ' + resp.getResponseCode());
    var rows = dudooPOS_parseCSV_(resp.getContentText());   // 與每日匯入同一支解析
    return rows.map(function (r, k) {
      return [day, k + 1, r[0], r[1], r[2], r[3], r[4], r[11], r[18], r[19], r[20], r[21], r[22], r[23], r[24], r.length];
    });
  });
}

/* ==================== 3. 肚肚專案折扣（單日） ==================== */

var RPR_DISC_FIELDS = [['day', 'STRING'], ['seq', 'INTEGER'], ['c0', 'STRING'], ['c1', 'STRING'], ['c2', 'STRING'],
  ['c3', 'STRING'], ['c4', 'STRING'], ['c5', 'STRING'], ['c6', 'STRING'], ['c7', 'STRING'], ['c8', 'STRING'],
  ['c9', 'STRING'], ['c10', 'STRING'], ['c11', 'STRING'], ['c12', 'STRING'], ['ncol', 'INTEGER']];

function rpr3_discounts() {
  var days = [], d = new Date(RPR.DISC_FROM + 'T00:00:00+08:00'), end = new Date(RPR.DISC_TO + 'T00:00:00+08:00');
  while (d <= end) { days.push(Utilities.formatDate(d, 'Asia/Taipei', 'yyyy-MM-dd')); d = new Date(d.getTime() + 86400000); }
  rprFetchDays_('DISC', 'rpr20261002_disc', days, RPR_DISC_FIELDS, function (tok, day) {
    var payload = 'access_token=' + encodeURIComponent(tok) + '&hierarchy_id=' + DUDOO_CONFIG_.HIERARCHY_ID +
      '&start_date=' + day + '&end_date=' + day + '&export_type=csv&params=';
    for (var i = 0; i < DUDOO_CONFIG_.COMPANY_IDS.length; i++) payload += '&company_id%5B%5D=' + DUDOO_CONFIG_.COMPANY_IDS[i];
    var resp = UrlFetchApp.fetch(DUDOO_CONFIG_.API_BASE + '/reports/getProjectDiscount?type=reports', {
      method: 'post', contentType: 'application/x-www-form-urlencoded', payload: payload,
      headers: { 'Access-Token': tok }, muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) throw new Error(day + ' 專案折扣 HTTP ' + resp.getResponseCode());
    var txt = resp.getContentText(), csv = txt;
    try { var j = JSON.parse(txt); if (j && typeof j.data === 'string') csv = j.data; } catch (e) {}
    var rows = dudooPOS_parseCSV_(csv);
    return rows.map(function (r, k) {
      var out = [day, k + 1];
      for (var c = 0; c < 13; c++) out.push(r[c] === undefined ? '' : r[c]);
      out.push(r.length);
      return out;
    });
  });
}

/** 逐日抓、每 8 天載入一次 BQ；已完成的日期記在指令碼屬性，可重跑接續。 */
function rprFetchDays_(tag, table, days, fields, fetchOne) {
  var props = PropertiesService.getScriptProperties();
  var key = 'RPR20261002_' + tag + '_DONE';
  var done = JSON.parse(props.getProperty(key) || '[]');
  // 清掉上次「已載入但沒記到進度」的殘留，避免重複
  if (done.length) {
    try {
      rprQuery_('DELETE FROM ' + rprT_(table) + ' WHERE day NOT IN (' + done.map(function (x) { return "'" + x + "'"; }).join(',') + ')');
    } catch (e) { Logger.log('（清殘留略過：' + e.message + '）'); }
  }
  var todo = days.filter(function (x) { return done.indexOf(x) < 0; });
  Logger.log(tag + '：共 ' + days.length + ' 天，已完成 ' + done.length + '，待抓 ' + todo.length);
  if (!todo.length) { Logger.log('✅ ' + tag + ' 全部完成'); return; }
  var tok = rprLogin_();
  var t0 = Date.now(), buf = [], bufDays = [], first = done.length === 0;
  function flush() {
    if (!bufDays.length) return;
    if (buf.length) {
      var got = rprLoad_(table, fields, buf.map(function (r) { return r.map(rprCsv_).join(','); }).join('\n'), first);
      if (got !== buf.length) throw new Error('載入列數不符：預期 ' + buf.length + '、實際 ' + got);
      first = false;
    } else if (first) {
      rprLoad_(table, fields, '', true); first = false;   // 只為建表
    }
    done = done.concat(bufDays);
    props.setProperty(key, JSON.stringify(done));
    Logger.log(tag + '：已載入 ' + bufDays.join('、') + '（' + buf.length + ' 列）');
    buf = []; bufDays = [];
  }
  for (var i = 0; i < todo.length; i++) {
    // 單日匯出偶爾要 1 分多鐘，超過 3 分鐘就不再開新的一天，確保 6 分鐘內收工
    if (Date.now() - t0 > RPR.FETCH_BUDGET_MS) { flush(); Logger.log('⏸ 時間到，請再執行一次接續'); return; }
    var t1 = Date.now();
    var rows = fetchOne(tok, todo[i]);
    Logger.log(tag + ' ' + todo[i] + '：' + rows.length + ' 列（' + ((Date.now() - t1) / 1000).toFixed(1) + ' 秒）');
    buf = buf.concat(rows); bufDays.push(todo[i]);
    if (bufDays.length >= 2) flush();
    Utilities.sleep(800);   // 對肚肚每秒不超過 1 次
  }
  flush();
  Logger.log('✅ ' + tag + ' 全部完成（' + done.length + ' 天）');
}

/* ==================== 4. 依計畫改 POS資料（唯一寫入） ==================== */

/**
 * 計畫表（Claude 在 BQ 產生、經營者核可的範圍）：
 *   rpr20261002_plan_del (rownum, a, c, d, i)：要刪的列＋該列應有內容（不符就整批中止）
 *   rpr20261002_plan_add (seq, a, b, c, d, g, h, i, j)：要補的列（欄位同肚肚匯出）
 *   rpr20261002_plan_meta (k, v)：last＝產生計畫時試算表最後一列、del_n、add_n
 * 流程：核對 → 一次刪完（Sheets API batchUpdate，全有或全無）→ 分批補列（沿用 dudooPOS_appendToSheet_）。
 * 進度記在 RPR20261002_APPLY；中途斷掉再執行會從下一批補列接續；完成後拒絕再跑。
 */
function rpr4_applySheet() {
  var props = PropertiesService.getScriptProperties();
  var KEY = 'RPR20261002_APPLY';
  var st = JSON.parse(props.getProperty(KEY) || 'null');
  if (st && st.stage === 'done') { Logger.log('🔴 已完成過（' + JSON.stringify(st) + '），不可重跑'); return; }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('有其他程式正在執行，稍後再試');
  try {
    var meta = {};
    rprQuery_('SELECT k, v FROM ' + rprT_('rpr20261002_plan_meta')).forEach(function (r) { meta[r[0]] = r[1]; });
    var add = rprQuery_('SELECT seq, a, b, c, d, g, h, i, j FROM ' + rprT_('rpr20261002_plan_add') + ' ORDER BY seq');
    if (String(add.length) !== String(meta.add_n)) throw new Error('補列數 ' + add.length + ' ≠ 計畫 ' + meta.add_n);

    var ss = SpreadsheetApp.openById(RPR.SS_ID), sh = ss.getSheetByName(RPR.TAB);

    if (!st) {
      var del = rprQuery_('SELECT rownum, a, c, d, i FROM ' + rprT_('rpr20261002_plan_del') + ' ORDER BY rownum');
      if (String(del.length) !== String(meta.del_n)) throw new Error('刪列數 ' + del.length + ' ≠ 計畫 ' + meta.del_n);
      var last = sh.getLastRow();
      if (String(last) !== String(meta.last)) throw new Error('試算表最後一列 ' + last + ' ≠ 產生計畫時 ' + meta.last + '（有人動過），中止');

      // 核對每一列要刪的內容
      var bad = [];
      var cache = {}, CH = 20000;
      del.forEach(function (r) {
        var rn = Number(r[0]), base = Math.floor((rn - 2) / CH) * CH + 2;
        if (!cache[base]) cache[base] = sh.getRange(base, 1, Math.min(CH, last - base + 1), 9).getValues();
        var v = cache[base][rn - base];
        var c = '';
        if (v[2]) { var d = v[2] instanceof Date ? v[2] : new Date(v[2]); c = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2) + '-' + ('0' + d.getDate()).slice(-2); }
        if (rprS_(v[0]) !== (r[1] || '') || c !== (r[2] || '') || rprS_(v[3]) !== (r[3] || '') || rprS_(v[8]) !== (r[4] || '')) {
          if (bad.length < 10) bad.push(rn + '：表內 [' + [rprS_(v[0]), c, rprS_(v[3]), rprS_(v[8])].join('|') + '] 計畫 [' + r.slice(1).join('|') + ']');
          else bad.push(rn);
        }
      });
      if (bad.length) { Logger.log('🔴 有 ' + bad.length + ' 列內容與計畫不符，未刪任何列：\n' + bad.slice(0, 10).join('\n')); return; }
      Logger.log('✅ 核對通過：' + del.length + ' 列內容皆與計畫相同');

      // 合併成連續區塊，由下往上一次刪完
      var rows = del.map(function (r) { return Number(r[0]); }).sort(function (a, b) { return a - b; });
      var blocks = [], s = rows[0], p = rows[0];
      for (var k = 1; k < rows.length; k++) {
        if (rows[k] === p + 1) p = rows[k]; else { blocks.push([s, p]); s = rows[k]; p = rows[k]; }
      }
      if (rows.length) blocks.push([s, p]);
      var sheetId = sh.getSheetId();
      var reqs = blocks.reverse().map(function (b) {
        return { deleteDimension: { range: { sheetId: sheetId, dimension: 'ROWS', startIndex: b[0] - 1, endIndex: b[1] } } };
      });
      if (reqs.length) Sheets.Spreadsheets.batchUpdate({ requests: reqs }, RPR.SS_ID);
      var after = SpreadsheetApp.openById(RPR.SS_ID).getSheetByName(RPR.TAB).getLastRow();
      Logger.log('刪列：' + blocks.length + ' 個區塊、' + rows.length + ' 列；最後一列 ' + last + ' → ' + after);
      if (after !== last - rows.length) {
        props.setProperty(KEY, JSON.stringify({ stage: 'deleted_mismatch', last: last, after: after, del: rows.length, appended: 0, ts: new Date().toISOString() }));
        throw new Error('刪後列數不符（預期 ' + (last - rows.length) + '），已停在刪除後、尚未補列，請人工檢查');
      }
      st = { stage: 'deleted', last: last, afterDel: after, del: rows.length, appended: 0, ts: new Date().toISOString() };
      props.setProperty(KEY, JSON.stringify(st));
    } else if (st.stage !== 'deleted') {
      Logger.log('🔴 狀態 ' + JSON.stringify(st) + '，需要人工檢查，未動作'); return;
    }

    // 分批補列（沿用每日匯入的寫法：Sheets API append＋清 E:F 讓 ARRAYFORMULA 展開）
    var t0 = Date.now();
    while (st.appended < add.length) {
      if (Date.now() - t0 > RPR.BUDGET_MS) { Logger.log('⏸ 時間到，已補 ' + st.appended + '／' + add.length + '，請再執行一次接續'); return; }
      var part = add.slice(st.appended, st.appended + RPR.ADD_CHUNK).map(function (r) {
        var raw = new Array(25);
        for (var x = 0; x < 25; x++) raw[x] = '';
        raw[0] = r[1] || ''; raw[1] = r[2] || ''; raw[2] = r[3] || ''; raw[11] = r[4] || '';
        raw[18] = r[5] || ''; raw[19] = r[6] || ''; raw[22] = r[7] || ''; raw[24] = r[8] || '';
        return raw;
      });
      dudooPOS_appendToSheet_(part);
      st.appended += part.length;
      props.setProperty(KEY, JSON.stringify(st));
      Logger.log('補列：' + st.appended + '／' + add.length);
    }
    var fin = SpreadsheetApp.openById(RPR.SS_ID).getSheetByName(RPR.TAB).getLastRow();
    st.stage = 'done'; st.finalLast = fin; st.doneTs = new Date().toISOString();
    props.setProperty(KEY, JSON.stringify(st));
    Logger.log((fin === st.afterDel + add.length ? '✅ ' : '🔴 ') + '完成：最後一列 ' + st.last + ' → 刪後 ' + st.afterDel + ' → 補後 ' + fin + '（預期 ' + (st.afterDel + add.length) + '）');
  } finally {
    lock.releaseLock();
  }
}

/* ==================== 6. 收掉補列時多長出來的空白列 ==================== */

/**
 * rpr4 用 Sheets API 分批 append 時，表尾原有的 10 列雜列被往下推，並多出 9,833 列全空白列
 * （2026-10-02 快照 sheet_post 證實：資料列 2～333967 連續；333968～333977＝原本的 10 列雜列；
 *   333978～343810 全空白）。經營者 10/02 核可刪除。
 * 本函式只刪「A～K 全空白」且位於 333978 之後的列；最後一列、表尾樣子、任何一格有值，對不上就整批中止。
 */
function rpr6_trimBlankTail() {
  var props = PropertiesService.getScriptProperties();
  var KEY = 'RPR20261002_TRIM';
  if (props.getProperty(KEY)) { Logger.log('🔴 已執行過：' + props.getProperty(KEY)); return; }
  var DATA_END = 333967, KEEP_END = 333977, LAST = 343810;
  var sh = SpreadsheetApp.openById(RPR.SS_ID).getSheetByName(RPR.TAB);
  var last = sh.getLastRow(), maxRows = sh.getMaxRows();
  Logger.log('刪前：最後一列 ' + last + '，表格總列數 ' + maxRows);
  if (last !== LAST) { Logger.log('🔴 最後一列 ' + last + ' ≠ 預期 ' + LAST + '，未動作'); return; }
  var head = sh.getRange(DATA_END, 1, KEEP_END - DATA_END + 1, 11).getValues();
  if (!head[0][2] || head[1].join('') !== '' || String(head[2][6]) !== '346' || String(head[3][6]) !== '0' || String(head[4][6]) !== '0') {
    Logger.log('🔴 表尾樣子與快照不同，未動作：' + JSON.stringify(head.slice(0, 5))); return;
  }
  var n = LAST - KEEP_END;
  var vals = sh.getRange(KEEP_END + 1, 1, n, 11).getValues();
  for (var i = 0; i < vals.length; i++) {
    if (vals[i].join('') !== '') { Logger.log('🔴 第 ' + (KEEP_END + 1 + i) + ' 列有值，未動作：' + JSON.stringify(vals[i])); return; }
  }
  Logger.log('✅ 核對通過：第 ' + (KEEP_END + 1) + '～' + LAST + ' 列共 ' + n + ' 列，A～K 全空白');
  sh.deleteRows(KEEP_END + 1, n);
  SpreadsheetApp.flush();
  var sh2 = SpreadsheetApp.openById(RPR.SS_ID).getSheetByName(RPR.TAB);
  var after = sh2.getLastRow(), maxAfter = sh2.getMaxRows();
  props.setProperty(KEY, JSON.stringify({ deleted: n, from: KEEP_END + 1, lastBefore: last, lastAfter: after, maxBefore: maxRows, maxAfter: maxAfter, ts: new Date().toISOString() }));
  Logger.log((after === KEEP_END ? '✅ ' : '🔴 ') + '刪除 ' + n + ' 列；最後一列 ' + last + ' → ' + after + '；表格總列數 ' + maxRows + ' → ' + maxAfter);
}

/* ==================== 共用小工具 ==================== */

function rprT_(t) { return '`' + RPR.PROJECT + '.' + RPR.DS + '.' + t + '`'; }
function rprS_(v) { return (v === null || v === undefined) ? '' : String(v); }
function rprCsv_(v) { return '"' + rprS_(v).replace(/"/g, '""') + '"'; }

function rprLogin_() {
  var p = PropertiesService.getScriptProperties();
  var resp = UrlFetchApp.fetch(DUDOO_CONFIG_.API_BASE + '/auth/login', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    payload: JSON.stringify({ code: p.getProperty('DUDOO_CODE'), username: p.getProperty('DUDOO_USERNAME'), password: p.getProperty('DUDOO_PASSWORD') })
  });
  var j = {};
  try { j = JSON.parse(resp.getContentText()); } catch (e) {}
  var tok = j.data && j.data.access_token;
  if (!tok) throw new Error('肚肚登入失敗 HTTP ' + resp.getResponseCode());
  return tok;
}

function rprQuery_(sql) {
  var r = BigQuery.Jobs.query({ query: sql, useLegacySql: false, timeoutMs: 60000, maxResults: 10000 }, RPR.PROJECT);
  var id = r.jobReference.jobId, loc = r.jobReference.location;
  while (!r.jobComplete) { Utilities.sleep(1000); r = BigQuery.Jobs.getQueryResults(RPR.PROJECT, id, { location: loc, maxResults: 10000 }); }
  var out = (r.rows || []).map(function (row) { return row.f.map(function (c) { return c.v; }); });
  var tok = r.pageToken;
  while (tok) {
    var pg = BigQuery.Jobs.getQueryResults(RPR.PROJECT, id, { location: loc, pageToken: tok, maxResults: 10000 });
    (pg.rows || []).forEach(function (row) { out.push(row.f.map(function (c) { return c.v; })); });
    tok = pg.pageToken;
  }
  return out;
}

function rprLoad_(table, fields, csv, truncate) {
  var job = { configuration: { load: {
    destinationTable: { projectId: RPR.PROJECT, datasetId: RPR.DS, tableId: table },
    schema: { fields: fields.map(function (f) { return { name: f[0], type: f[1] }; }) },
    sourceFormat: 'CSV', allowQuotedNewlines: true, createDisposition: 'CREATE_IF_NEEDED',
    writeDisposition: truncate ? 'WRITE_TRUNCATE' : 'WRITE_APPEND'
  } } };
  var j = BigQuery.Jobs.insert(job, RPR.PROJECT, Utilities.newBlob(csv || '', 'application/octet-stream'));
  var id = j.jobReference.jobId, loc = j.jobReference.location;
  for (var t = 0; t < 180 && j.status.state !== 'DONE'; t++) { Utilities.sleep(1000); j = BigQuery.Jobs.get(RPR.PROJECT, id, { location: loc }); }
  if (j.status.state !== 'DONE') throw new Error('BQ 載入逾時：' + table);
  if (j.status.errorResult) throw new Error('BQ 載入失敗 ' + table + '：' + JSON.stringify(j.status.errors || j.status.errorResult).slice(0, 500));
  return Number((j.statistics && j.statistics.load && j.statistics.load.outputRows) || 0);
}
