/**
 * dudooGuard.gs — 肚肚匯入守門員＋每日對帳自動修正（2026-10-05 新增；v2 同日擴充）
 * 專案：「一鍵追加新品」（綁 BOM 本，script ID 1PkavxV6r3GZbEbS_B5ZWGL1ovm30T55XhllAuI1aglBe_RPERdP-DJnJ）
 * 經營者 2026-10-05 裁示：POS 儀表板要跟肚肚「業績概況」的營業額一致；「以後不要再發生、不要又要我處理」。
 *
 * 會讓兩邊不一致的狀況（2026-10-05 實查）與本檔的處理：
 *   ① 跨日退款：肚肚算在「退款那天」→ 匯入時把建立日期≠匯入日的列改記匯入日。
 *   ② 幽靈品項：銷售明細匯出檔多一列帳單裡沒有的品項 → 拿「交易紀錄」逐張比，能唯一認定就剔除。
 *   ③ 肚肚自己對不上：帳單總額≠品項加總（例 2025-05-17 店7 第 68 張單 1,900 但品項只有 600）
 *      → 補一列「肚肚對帳調整」（主類別 入場&共廚&其他：不算甜點數、不算來客數），讓銷售總額＝業績概況。
 *   ④ 折扣漏匯（Make 漏跑）→ 每日對帳自動從肚肚「專案折扣明細」重抓當天覆蓋 pos_discounts。
 *   ⑤ 肚肚事後作廢／改單 → 每日對帳自動重抓那一店那一天（守門員上線日以後才重抓；之前的日子用調整列）。
 *
 * 入口：
 *   dudooGuard_apply_(rows, ds, token)：dudooPOS_loginAndImportDate_ 寫進 POS資料 之前呼叫（①②③）。出錯不擋匯入。
 *   dudooCheck_daily()：每天 05:00 觸發（2026-10-07 由 07:30 提早）。近 7 天逐店日＋本月／上月逐店總額 vs 業績概況，對不上就自動修（④⑤③），
 *     修完再比一次；結果寫 BOM 本「肚肚對帳」分頁（POS 儀表板頂部燈號讀這裡）。沒事不寄信，修好寄「不用處理」。
 *   aaaDudooGuardDryRun()／aaaDudooCheckDryRun()：唯讀預演，不寫入、不寄信。
 *   dudooCheck_installTrigger()：建立 05:00 觸發器（已存在就不重複建）。
 * 每次自動修正前，BigQuery 先備份到 pos_transactions_autobak／pos_discounts_autobak（多一欄 bak_at）。
 */

var DGUARD = {
  TZ: 'Asia/Taipei',
  STORES: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],   // 13 西門已歇業、100 測試店不比
  CHECK_DAYS: 7,
  LIVE_FROM: '2026-10-05',       // 守門員上線後的第一個營業日：此日（含）以後的店日才整天重抓
  MAX_FIX_STOREDAYS: 8,          // 每次最多修幾個店日（6 分鐘上限；其餘隔天再修）
  ADJ_MAX_ABS: 3000,             // 調整列安全上限：差額超過 3,000 且超過該店當日 20% 就不自動補（多半是 API 異常）
  ADJ_MAX_RATIO: 0.2,
  TIME_BUDGET_MS: 240000,        // 開始修正前已用超過 4 分鐘就不再修
  ADJ_NAME: '肚肚對帳調整',
  ADJ_CATEGORY: '入場&共廚&其他',
  RESULT_TAB: '肚肚對帳',
  MAP_TAB: '產品名稱對照表',
  PROP_ALERTED: 'DGUARD_ALERTED',
  BQ: '`diybc-make-sync.diybc_pos.',
  DRYRUN_DATES: ['2026-07-29', '2026-01-05', '2026-10-04']
};

/* ═══════════ 純計算（不呼叫任何 Google 服務，可單元測試） ═══════════ */

/** 日期字串 → 'yyyy-MM-dd'；看不懂回 '' */
function dguardNormDate_(v) {
  var m = String(v || '').trim().match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

/** ① 跨日退款改記匯入日。直接改 rows[i][2]，回傳說明文字陣列 */
function dguardRedate_(rows, ds) {
  var out = [];
  rows.forEach(function (r) {
    var d = dguardNormDate_(r[2]);
    if (d && d !== ds) {
      out.push(r[0] + '店 ' + r[11] + ' ×' + r[19] + ' 單價 ' + r[18] + '（匯出檔日期 ' + d + '）');
      r[2] = ds;
    }
  });
  return out;
}

function dguardLine_(r) { return (parseFloat(r[18]) || 0) * (parseFloat(r[19]) || 0); }

/** 匯出檔逐店銷售總額 */
function dguardExpGross_(rows) {
  var exp = {};
  rows.forEach(function (r) {
    var st = parseInt(r[0], 10);
    if (!isNaN(st)) exp[st] = (exp[st] || 0) + dguardLine_(r);
  });
  return exp;
}

/**
 * ② 幽靈品項判定。
 *   rows：匯出檔欄位陣列（0 店號、4 交易序號、11 品名、18 單價、19 數量）
 *   dash：{店號: {cid, gross}}（業績概況）
 *   billsOf(cid)：回該店當日帳單陣列（交易紀錄：sale_code、sale_amount、discount、form_no）
 * 回傳 {rows: 剔除後的列, removed: [說明], unresolved: [{store, diff, why}]}
 */
function dguardGhost_(rows, dash, billsOf) {
  var res = { rows: rows, removed: [], unresolved: [] };
  var exp = dguardExpGross_(rows);
  var drop = {};
  DGUARD.STORES.forEach(function (st) {
    var dg = dash[st];
    var diff = Math.round((exp[st] || 0) - (dg ? dg.gross : 0));
    if (diff === 0) return;
    if (!dg) { res.unresolved.push({ store: st, diff: diff, why: '業績概況沒有這家店' }); return; }

    var billGross = {}, billNo = {};
    billsOf(dg.cid).forEach(function (b) {
      var code = String(b.sale_code || '').replace(/\D/g, '');
      if (!code) return;
      billGross[code] = (billGross[code] || 0) + (Number(b.sale_amount) || 0) + (Number(b.discount) || 0);
      billNo[code] = b.form_no;
    });
    var bySerial = {};
    rows.forEach(function (r, i) {
      if (parseInt(r[0], 10) !== st) return;
      var code = String(r[4] || '').replace(/\D/g, '');
      (bySerial[code] = bySerial[code] || []).push(i);
    });

    if (diff < 0) {   // 匯出檔比較少：沒有可剔除的，交給③調整列；這裡只留說明
      var short = Object.keys(billGross).map(function (code) {
        var e = (bySerial[code] || []).reduce(function (a, i) { return a + dguardLine_(rows[i]); }, 0);
        return { code: code, gap: Math.round(billGross[code] - e) };
      }).filter(function (x) { return x.gap > 0 && (bySerial[x.code] || []).length > 0; });
      res.unresolved.push({ store: st, diff: diff, why: '匯出檔比業績概況少' +
        (short.length ? '（帳單總額比品項多：' + short.map(function (x) { return '第 ' + (billNo[x.code] || '?') + ' 張 ' + x.gap; }).join('、') + '）' : '') });
      return;
    }

    var cand = [], candSum = 0;
    Object.keys(bySerial).forEach(function (code) {
      var idx = bySerial[code];
      var eSum = idx.reduce(function (a, i) { return a + dguardLine_(rows[i]); }, 0);
      if (billGross[code] === undefined) {
        var allPos = idx.every(function (i) { return (parseFloat(rows[i][19]) || 0) > 0; });
        if (allPos && eSum > 0) { cand.push({ code: code, idx: idx, amt: Math.round(eSum), why: '交易紀錄沒有這張單' }); candSum += Math.round(eSum); }
        return;
      }
      var excess = Math.round(eSum - billGross[code]);
      if (excess <= 0) return;
      var hit = -1;
      for (var j = 0; j < idx.length; j++) { if (Math.round(dguardLine_(rows[idx[j]])) === excess) { hit = idx[j]; break; } }
      if (hit >= 0) { cand.push({ code: code, idx: [hit], amt: excess, why: '帳單裡沒有這一列' }); candSum += excess; }
      else { cand.push({ code: code, idx: [], amt: excess, why: '帳單少 ' + excess + ' 但找不到同金額的單一品項' }); candSum += excess; }
    });

    var clean = cand.length > 0 && candSum === diff && cand.every(function (c) { return c.idx.length > 0; });
    if (clean) {
      cand.forEach(function (c) {
        c.idx.forEach(function (i) {
          drop[i] = true;
          var r = rows[i];
          res.removed.push(st + '店 單號 ' + r[4] + ' ' + r[11] + ' ×' + r[19] + ' 單價 ' + r[18] + '（' + c.why + '）');
        });
      });
    } else {
      res.unresolved.push({ store: st, diff: diff, why: '比對交易紀錄後無法唯一認定' +
        (cand.length ? '（' + cand.map(function (c) { return c.code + ' ' + c.amt + ' ' + c.why; }).join('；') + '）' : '') });
    }
  });
  if (Object.keys(drop).length) res.rows = rows.filter(function (r, i) { return !drop[i]; });
  return res;
}

/** 調整列（匯出檔欄位格式；dudooPOS_appendToSheet_ 會取 0,1,2,11,18,19,22,24） */
function dguardAdjRow_(st, name, ds, amt) {
  var r = [];
  for (var i = 0; i < 26; i++) r.push('');
  r[0] = String(st); r[1] = name; r[2] = ds; r[11] = DGUARD.ADJ_NAME;
  r[18] = String(amt); r[19] = '1'; r[22] = String(amt); r[24] = '0';
  return r;
}

/** 店名：先用匯出檔同店的列，沒有就用業績概況名稱去掉店號 */
function dguardStoreName_(rows, st, dash) {
  for (var i = 0; i < rows.length; i++) if (parseInt(rows[i][0], 10) === st && rows[i][1]) return String(rows[i][1]);
  return dash[st] && dash[st].name ? String(dash[st].name).replace(/^\s*\d+\s*/, '') : String(st);
}

/** 差額是否在可自動補的範圍（避免 API 異常時補一大筆） */
function dguardAdjOk_(amt, base) {
  return Math.abs(amt) <= DGUARD.ADJ_MAX_ABS || Math.abs(amt) <= Math.abs(base) * DGUARD.ADJ_MAX_RATIO;
}

/** ③ 剔除後仍有差額的店補一列調整，讓該店銷售總額＝業績概況。回傳 {rows, adjusted:[{store, amt}], skipped:[{store, amt}]} */
function dguardAdjust_(rows, ds, dash) {
  var exp = dguardExpGross_(rows), adjusted = [], skipped = [], out = rows.slice();
  DGUARD.STORES.forEach(function (st) {
    if (!dash[st]) return;
    var amt = Math.round(dash[st].gross - (exp[st] || 0));
    if (amt === 0) return;
    if (!dguardAdjOk_(amt, exp[st] || 0)) { skipped.push({ store: st, amt: amt }); return; }
    out.push(dguardAdjRow_(st, dguardStoreName_(rows, st, dash), ds, amt));
    adjusted.push({ store: st, amt: amt });
  });
  return { rows: out, adjusted: adjusted, skipped: skipped };
}

/** 肚肚「專案折扣明細」CSV → pos_discounts 列（店別空白＝13 西門；100 測試店略過） */
function dguardParseDiscounts_(csv, ds) {
  var lines = String(csv || '').split(/\r?\n/), out = [];
  for (var i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    var c = dudooPOS_splitCSVLine_(lines[i]);
    if (c.length < 12) continue;
    var d = dguardNormDate_(c[0]);
    if (!d) continue;
    if (String(c[1]) === '100') continue;
    out.push({ date: d, store: c[1] === '' ? 13 : parseInt(c[1], 10), project: c[2], serial: c[3],
               qty: parseInt(c[6], 10) || 0, total: Number(c[7]) || 0, discount: Number(c[9]) || 0, actual: Number(c[11]) || 0 });
  }
  return out;
}

/** 逐店日比對：ours {店|日: {g,d,n}}，dashOf(ds) → 業績概況；回 [{ds, st, o, x}] */
function dguardCompare_(days, stores, ours, dashOf) {
  var bad = [];
  days.forEach(function (ds) {
    var dash = dashOf(ds);
    stores.forEach(function (st) {
      var o = ours[st + '|' + ds] || { g: 0, d: 0, n: 0 };
      var x = dash[st] || { gross: 0, discount: 0, amount: 0 };
      if (o.n !== x.amount || o.g !== x.gross || o.d !== x.discount) bad.push({ ds: ds, st: st, o: o, x: x });
    });
  });
  return bad;
}

/** 依日期與上線日決定每個不一致店日要怎麼修 */
function dguardPlan_(bad) {
  var discDates = {}, repull = [], adjust = [];
  bad.forEach(function (b) {
    if (b.o.d !== b.x.discount) discDates[b.ds] = true;
    if (b.o.g !== b.x.gross) {
      if (b.ds >= DGUARD.LIVE_FROM) repull.push({ ds: b.ds, st: b.st });
      else adjust.push({ ds: b.ds, st: b.st, amt: Math.round(b.x.gross - b.o.g) });
    }
  });
  return { discDates: Object.keys(discDates).sort(), repull: repull, adjust: adjust };
}

function dguardSql_(s) {
  if (s === null || s === undefined) return 'NULL';
  return "'" + String(s).replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/\n/g, '\\n') + "'";
}

/** 'yyyy-MM-dd' 加減天數（純字串運算，避開時區） */
function dguardAddDays_(ds, n) {
  var p = ds.split('-'), d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}

/* ═══════════ 肚肚 API ═══════════ */

function dguardLogin_() {
  var props = PropertiesService.getScriptProperties();
  var resp = UrlFetchApp.fetch(DUDOO_CONFIG_.API_BASE + '/auth/login', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ code: props.getProperty('DUDOO_CODE'), username: props.getProperty('DUDOO_USERNAME'),
                              password: props.getProperty('DUDOO_PASSWORD') }),
    muteHttpExceptions: true
  });
  var j = {};
  try { j = JSON.parse(resp.getContentText() || '{}'); } catch (e) { }
  var tok = j.data && j.data.access_token;
  if (!tok) throw new Error('肚肚登入失敗 HTTP ' + resp.getResponseCode());
  return tok;
}

function dguardPostRaw_(path, token, pairs) {
  var body = pairs.map(function (p) { return encodeURIComponent(p[0]) + '=' + encodeURIComponent(p[1]); }).join('&');
  var resp = UrlFetchApp.fetch(DUDOO_CONFIG_.API_BASE + path, {
    method: 'post', contentType: 'application/x-www-form-urlencoded', payload: body,
    headers: { 'Access-Token': token }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error(path + ' HTTP ' + resp.getResponseCode());
  return resp.getContentText();
}

function dguardPost_(path, token, pairs) { return JSON.parse(dguardPostRaw_(path, token, pairs)); }

function dguardCompanyPairs_(pairs, cids) {
  (cids || DUDOO_CONFIG_.COMPANY_IDS).forEach(function (c) { pairs.push(['company_id[]', c]); });
  return pairs;
}

/** 業績概況（from～to 合計）→ {店號: {cid, name, amount, discount, gross}} */
function dguardDashboard_(token, from, to) {
  var j = dguardPost_('/reports/getGroupDashboard?type=reports', token, dguardCompanyPairs_(
    [['access_token', token], ['start_date', from], ['end_date', to], ['hierarchy_id', DUDOO_CONFIG_.HIERARCHY_ID]]));
  if (!j || !j.data || !j.data.length) throw new Error('業績概況沒有回資料');
  var out = {};
  j.data.forEach(function (d) {
    var st = parseInt(d.company_name, 10);
    if (isNaN(st)) return;
    out[st] = { cid: String(d.company_id), name: String(d.company_name || ''), amount: Math.round(Number(d.amount) || 0),
                discount: Math.round(Number(d.discount_amount) || 0), gross: Math.round(Number(d.total_sale_amount) || 0) };
  });
  return out;
}

/** 交易紀錄：某店某日全部帳單 */
function dguardBills_(token, cid, ds) {
  var all = [], start = 0;
  for (var page = 0; page < 10; page++) {
    var j = dguardPost_('/reports/getBillInvoicesList?type=reports', token, [
      ['access_token', token], ['draw', '1'], ['start', String(start)], ['length', '500'],
      ['hierarchy_id', DUDOO_CONFIG_.HIERARCHY_ID], ['company_id', cid], ['start_date', ds], ['end_date', ds],
      ['status', ''], ['time_filter', 'bill_create_time'], ['order_source_filter', ''], ['trans_type_filter', '']]);
    if (!j || !j.success) throw new Error('交易紀錄讀取失敗（店 ' + cid + '）');
    var got = j.data || [];
    all = all.concat(got);
    start += got.length;
    if (!got.length || start >= (Number(j.recordsTotal) || 0)) break;
  }
  return all;
}

/** 銷售明細匯出（單日；cids 不給＝全部店） */
function dguardExport_(token, ds, cids) {
  var csv = dguardPostRaw_('/reports/getSaleDetailsAnalysis/export', token, dguardCompanyPairs_(
    [['hierarchy_id', DUDOO_CONFIG_.HIERARCHY_ID], ['start_date', ds], ['end_date', ds]], cids));
  return dudooPOS_parseCSV_(csv);
}

/** 專案折扣明細（單日全部店） */
function dguardDiscounts_(token, ds) {
  var csv = dguardPostRaw_('/reports/getProjectDiscount?type=reports', token, dguardCompanyPairs_(
    [['access_token', token], ['hierarchy_id', DUDOO_CONFIG_.HIERARCHY_ID], ['start_date', ds], ['end_date', ds],
     ['export_type', 'csv'], ['params', '']]));
  return dguardParseDiscounts_(csv, ds);
}

function dguardMail_(subject, lines) {
  try { MailApp.sendEmail(DUDOO_CONFIG_.NOTIFY_EMAIL, subject, lines.join('\n')); }
  catch (e) { Logger.log('[守門員] 寄信失敗：' + e); }
}

/** BigQuery 寫入（允許多敘述） */
function dguardBqRun_(sql) {
  var res = BigQuery.Jobs.query({ query: sql, useLegacySql: false, timeoutMs: 120000 }, 'diybc-make-sync');
  if (!res.jobComplete) throw new Error('BigQuery 逾時');
  return res;
}

/* ═══════════ 匯入守門員（dudooPOS_loginAndImportDate_ 呼叫） ═══════════ */

/** 對一天的匯出列跑 ①②③；回 {rows, redated, removed, adjusted, notes} */
function dguardProcessDay_(rows, ds, token, dash, onlyStores) {
  var redated = dguardRedate_(rows, ds);
  var d = dash;
  if (onlyStores) { d = {}; onlyStores.forEach(function (st) { if (dash[st]) d[st] = dash[st]; }); }
  var g = dguardGhost_(rows, d, function (cid) { return dguardBills_(token, cid, ds); });
  var a = dguardAdjust_(g.rows, ds, d);
  return { rows: a.rows, redated: redated, removed: g.removed, adjusted: a.adjusted, skipped: a.skipped,
           notes: g.unresolved.map(function (u) { return u.store + '店 ' + u.why; }) };
}

function dudooGuard_apply_(rows, ds, token, dry) {
  try {
    var dash = dguardDashboard_(token, ds, ds);
    var p = dguardProcessDay_(rows, ds, token, dash);
    if (p.redated.length) Logger.log('[守門員] 跨日退款改記 ' + ds + '：' + p.redated.length + ' 列\n  ' + p.redated.join('\n  '));
    if (p.skipped.length) Logger.log('[守門員] 差額過大不自動補（交給 05:00 每日對帳）：' + p.skipped.map(function (x) { return x.store + '店 ' + x.amt; }).join('、'));
    if (!p.removed.length && !p.adjusted.length) {
      if (!p.skipped.length) Logger.log('[守門員] ' + ds + ' 12 店銷售總額＝業績概況 ✅');
      return p.rows;
    }
    var lines = ['肚肚匯入守門員 ' + ds + '：已自動處理，寫入後 12 店銷售總額＝肚肚業績概況。不用處理。', ''];
    if (p.removed.length) { lines.push('剔除匯出檔多出來的品項（帳單裡沒有）：'); p.removed.forEach(function (s) { lines.push('  ' + s); }); lines.push(''); }
    if (p.adjusted.length) {
      lines.push('補「' + DGUARD.ADJ_NAME + '」列（肚肚自己兩份報表對不上，補到＝業績概況；不算甜點數與來客數）：');
      p.adjusted.forEach(function (x) { lines.push('  ' + x.store + '店 ' + (x.amt > 0 ? '+' : '') + x.amt); });
      if (p.notes.length) p.notes.forEach(function (s) { lines.push('    ' + s); });
      lines.push('');
    }
    if (p.redated.length) { lines.push('跨日退款改記 ' + ds + '：'); p.redated.forEach(function (s) { lines.push('  ' + s); }); }
    Logger.log(lines.join('\n'));
    if (!dry) dguardMail_('[POS守門] ' + ds + ' ✅ 已自動對齊肚肚（不用處理）', lines);
    return p.rows;
  } catch (e) {
    dguardRedate_(rows, ds);   // 比對失敗也要照退款日規則
    Logger.log('[守門員] ⚠️ 比對失敗（照原資料匯入）：' + e);
    if (!dry) dguardMail_('[POS守門] ' + ds + ' 比對暫時失敗（資料已照常匯入，05:00 每日對帳會自動補修，不用處理）',
      ['肚肚匯入守門員 ' + ds, '', '錯誤：' + e, '', '資料已照常寫入；05:00 每日對帳會再比一次並自動修正。']);
    return rows;
  }
}

/** 唯讀預演：抓 DRYRUN_DATES 每天的匯出檔跑守門員，只寫執行記錄。 */
function aaaDudooGuardDryRun() {
  var token = dguardLogin_();
  DGUARD.DRYRUN_DATES.forEach(function (ds) {
    var rows = dguardExport_(token, ds);
    var before = rows.length;
    var out = dudooGuard_apply_(rows, ds, token, true);
    Logger.log('[預演] ' + ds + '：匯出 ' + before + ' 列 → 寫入 ' + out.length + ' 列');
  });
  Logger.log('[預演] 結束：沒有寫入試算表、沒有寄信');
}

/* ═══════════ 每日對帳＋自動修正（05:00 觸發） ═══════════ */

function dudooCheck_daily() { return dguardCheck_(false); }

/** 唯讀預演：照常比對、列出「會怎麼修」，不寫入、不寄信。 */
function aaaDudooCheckDryRun() { return dguardCheck_(true); }

function dguardOurs_(from, to) {
  var rows = audit_bqAll_('SELECT store_code, CAST(sale_date AS STRING), CAST(SUM(gross) AS INT64), ' +
    'CAST(SUM(total_discount) AS INT64), CAST(SUM(net) AS INT64) FROM ' + DGUARD.BQ + "v_daily_net` " +
    "WHERE sale_date BETWEEN '" + from + "' AND '" + to + "' GROUP BY 1, 2");
  var ours = {};
  rows.forEach(function (r) { ours[parseInt(r[0], 10) + '|' + r[1]] = { g: Number(r[2]) || 0, d: Number(r[3]) || 0, n: Number(r[4]) || 0 }; });
  return ours;
}

function dguardOursRange_(from, to) {
  var rows = audit_bqAll_('SELECT store_code, CAST(SUM(net) AS INT64) FROM ' + DGUARD.BQ + "v_daily_net` " +
    "WHERE sale_date BETWEEN '" + from + "' AND '" + to + "' GROUP BY 1");
  var out = {};
  rows.forEach(function (r) { out[parseInt(r[0], 10)] = Number(r[1]) || 0; });
  return out;
}

function dguardDays_(from, to) {
  var out = [], d = from;
  while (d <= to) { out.push(d); d = dguardAddDays_(d, 1); }
  return out;
}

/** ④ 折扣：重抓肚肚當天全部店的專案折扣明細，備份後覆蓋；各店合計再對不上業績概況就補折扣調整列 */
function dguardFixDiscounts_(token, ds, dash) {
  var list = dguardDiscounts_(token, ds);
  var sum = {};
  list.forEach(function (x) { sum[x.store] = (sum[x.store] || 0) + x.discount; });
  var anyDash = DGUARD.STORES.some(function (st) { return dash[st] && dash[st].discount > 0; });
  if (!list.length && anyDash) throw new Error(ds + ' 折扣明細回 0 筆但業績概況有折扣，不覆蓋');
  var adj = [];
  DGUARD.STORES.forEach(function (st) {
    var want = dash[st] ? dash[st].discount : 0, got = Math.round(sum[st] || 0);
    if (want !== got) {
      list.push({ date: ds, store: st, project: DGUARD.ADJ_NAME, serial: '', qty: 1, total: 0, discount: want - got, actual: 0 });
      adj.push(st + '店 ' + (want - got > 0 ? '+' : '') + (want - got));
    }
  });
  var T = DGUARD.BQ + 'pos_discounts`', B = DGUARD.BQ + 'pos_discounts_autobak`';
  var values = list.map(function (x) {
    return "(DATE '" + ds + "', " + x.store + ', ' + dguardSql_(x.project) + ', ' + dguardSql_(x.serial) + ', ' + x.qty +
      ", NUMERIC '" + x.total + "', NUMERIC '" + x.discount + "', NUMERIC '" + x.actual + "', CURRENT_TIMESTAMP())";
  });
  dguardBqRun_(
    'CREATE TABLE IF NOT EXISTS ' + B + ' AS SELECT *, CURRENT_TIMESTAMP() AS bak_at FROM ' + T + ' WHERE FALSE;\n' +
    'INSERT INTO ' + B + ' SELECT *, CURRENT_TIMESTAMP() FROM ' + T + " WHERE sale_date = '" + ds + "';\n" +
    'BEGIN TRANSACTION;\n' +
    'DELETE FROM ' + T + " WHERE sale_date = '" + ds + "';\n" +
    (values.length ? 'INSERT INTO ' + T + ' (sale_date, store_code, project_name, serial_no, quantity, total_amount, discount, actual_amount, imported_at) VALUES\n' + values.join(',\n') + ';\n' : '') +
    'COMMIT TRANSACTION;');
  return list.length + ' 筆' + (adj.length ? '（另補折扣調整：' + adj.join('、') + '）' : '');
}

/** 確保「肚肚對帳調整」在產品名稱對照表（主類別 入場&共廚&其他），否則會被當成「找不到」算進甜點數 */
function dguardEnsureMapping_() {
  var sh = SpreadsheetApp.openById(RECON_POS_SS_ID).getSheetByName(DGUARD.MAP_TAB);
  if (!sh) throw new Error('找不到分頁「' + DGUARD.MAP_TAB + '」');
  var last = sh.getLastRow();
  var names = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
  var lastA = 1;
  for (var i = 0; i < names.length; i++) {
    var v = String(names[i][0]).trim();
    if (v === DGUARD.ADJ_NAME) return false;
    if (v) lastA = i + 2;
  }
  var r = lastA + 1;   // 只寫 A（品名）、G（主類別）、H（次類別），不碰其他欄（C 欄是偵測新產品清單）
  sh.getRange(r, 1).setValue(DGUARD.ADJ_NAME);
  sh.getRange(r, 7, 1, 2).setValues([[DGUARD.ADJ_CATEGORY, '無']]);
  SpreadsheetApp.flush();
  return true;
}

/**
 * ⑤／③ 品項：tasks = [{ds, st, mode:'repull'|'adjust', amt}]。
 * 一次讀 POS資料 A～C 找要換掉的列 → 由下往上刪 → 一次補列 → 一次重寫 BigQuery。
 */
function dguardFixItems_(token, tasks, dashOf) {
  if (!tasks.length) return [];
  dguardEnsureMapping_();
  var newRows = [], done = [], repullKeys = {};
  var byDate = {};
  tasks.forEach(function (t) { (byDate[t.ds] = byDate[t.ds] || []).push(t); });
  Object.keys(byDate).sort().forEach(function (ds) {
    var ts = byDate[ds], dash = dashOf(ds);
    var rp = ts.filter(function (t) { return t.mode === 'repull'; });
    if (rp.length) {
      var stores = rp.map(function (t) { return t.st; });
      var cids = stores.map(function (st) { return dash[st].cid; });
      var rows = dguardExport_(token, ds, cids).filter(function (r) { return stores.indexOf(parseInt(r[0], 10)) >= 0; });
      var p = dguardProcessDay_(rows, ds, token, dash, stores);
      var unsafe = {};
      p.skipped.forEach(function (x) { unsafe[x.store] = '差額 ' + x.amt + ' 過大'; });
      stores.forEach(function (st) {
        var real = p.rows.filter(function (r) { return parseInt(r[0], 10) === st && r[11] !== DGUARD.ADJ_NAME; }).length;
        if (!real && dash[st] && dash[st].gross) unsafe[st] = '重抓到 0 列';
      });
      newRows = newRows.concat(p.rows.filter(function (r) { return !unsafe[parseInt(r[0], 10)]; }));
      rp.forEach(function (t) {
        if (unsafe[t.st]) { done.push(ds + ' ' + t.st + '店：⚠️ ' + unsafe[t.st] + '，保留原資料不動'); return; }
        repullKeys[t.st + '|' + ds] = true;
        var adj = p.adjusted.filter(function (x) { return x.store === t.st; });
        done.push(ds + ' ' + t.st + '店：整天重抓 ' + p.rows.filter(function (r) { return parseInt(r[0], 10) === t.st; }).length + ' 列' +
          (p.removed.length ? '，剔除幽靈品項 ' + p.removed.length : '') + (adj.length ? '，補調整 ' + adj[0].amt : ''));
      });
    }
    ts.filter(function (t) { return t.mode === 'adjust'; }).forEach(function (t) {
      if (!dguardAdjOk_(t.amt, dash[t.st] ? dash[t.st].gross : 0)) { done.push(ds + ' ' + t.st + '店：⚠️ 差額 ' + t.amt + ' 過大，不自動補'); return; }
      newRows.push(dguardAdjRow_(t.st, (dash[t.st] && dash[t.st].name ? dash[t.st].name.replace(/^\s*\d+\s*/, '') : String(t.st)), ds, t.amt));
      done.push(ds + ' ' + t.st + '店：補調整列 ' + (t.amt > 0 ? '+' : '') + t.amt + '（' + DGUARD.LIVE_FROM + ' 前的日子不整天重抓）');
    });
  });

  // BigQuery 備份要動到的店日
  var keys = tasks.map(function (t) { return "(sale_date = '" + t.ds + "' AND store_code = " + t.st + ')'; });
  var T = DGUARD.BQ + 'pos_transactions`', B = DGUARD.BQ + 'pos_transactions_autobak`';
  dguardBqRun_('CREATE TABLE IF NOT EXISTS ' + B + ' AS SELECT *, CURRENT_TIMESTAMP() AS bak_at FROM ' + T + ' WHERE FALSE;\n' +
    'INSERT INTO ' + B + ' SELECT *, CURRENT_TIMESTAMP() FROM ' + T + ' WHERE ' + keys.join(' OR ') + ';');

  // 先記下要刪的舊列（由目前資料找），先補新列、最後才刪舊列：中途失敗也不會少資料
  var del = [], sh = SpreadsheetApp.openById(RECON_POS_SS_ID).getSheetByName(RECON_POS_TAB);
  if (Object.keys(repullKeys).length) {
    var last = sh.getLastRow();
    var vals = sh.getRange(2, 1, last - 1, 3).getValues(), cache = {};
    for (var i = 0; i < vals.length; i++) {
      var st = parseInt(vals[i][0], 10);
      if (isNaN(st)) continue;
      var v = vals[i][2], d;
      if (v instanceof Date) { d = cache[v.getTime()]; if (d === undefined) d = cache[v.getTime()] = Utilities.formatDate(v, DGUARD.TZ, 'yyyy-MM-dd'); }
      else d = dguardNormDate_(v);
      if (repullKeys[st + '|' + d]) del.push(i + 2);
    }
  }
  if (newRows.length) dudooPOS_appendToSheet_(newRows);
  SpreadsheetApp.flush();
  for (var k = del.length - 1; k >= 0;) {   // 新列在表尾，刪上面的舊列不影響新列
    var end = del[k], start = end;
    while (k - 1 >= 0 && del[k - 1] === start - 1) { k--; start--; }
    sh.deleteRows(start, end - start + 1);
    k--;
  }
  SpreadsheetApp.flush();
  syncSheetToBigQuery_batch_v2(tasks.map(function (t) { return { date: t.ds, store: t.st }; }));
  return done;
}

/** 寫 BOM 本「肚肚對帳」分頁：第 2 列摘要（POS 儀表板燈號讀這列），第 5 列起明細 */
function dguardWriteResult_(summary, detail) {
  var ss = SpreadsheetApp.openById(RECON_POS_SS_ID);
  var sh = ss.getSheetByName(DGUARD.RESULT_TAB) || ss.insertSheet(DGUARD.RESULT_TAB);
  sh.clearContents();
  sh.getRange(1, 1, 200, 2).setNumberFormat('@');   // 時間、日期存純文字，POS 儀表板燈號才讀得到原字串
  sh.getRange(1, 1, 2, 7).setValues([
    ['檢查時間', '檢查範圍', '狀態', '不一致店日', '已自動修正', '未修好', '說明'],
    summary]);
  sh.getRange(4, 1).setValue('明細（只列曾經不一致的店日）');
  var hdr = ['日期', '店號', '儀表板營業額', '肚肚營業額', '差額', '處理'];
  var body = [hdr].concat(detail.length ? detail : [['', '', '', '', '', '全部一致']]);
  sh.getRange(5, 1, body.length, hdr.length).setValues(body);
}

function dguardCheck_(dry) {
  var t0 = Date.now(), tz = DGUARD.TZ;
  var today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  var to = dguardAddDays_(today, -1), from = dguardAddDays_(today, -DGUARD.CHECK_DAYS);
  var monthFrom = to.substring(0, 8) + '01';
  var prevTo = dguardAddDays_(monthFrom, -1), prevFrom = prevTo.substring(0, 8) + '01';
  var lock = null;
  if (!dry) {
    lock = LockService.getScriptLock();
    if (!lock.tryLock(120000)) { Logger.log('[每日對帳] 拿不到鎖，略過'); return { skipped: true }; }
  }
  try {
    var token = dguardLogin_();
    var dashCache = {};
    var dashOf = function (ds) { return dashCache[ds] || (dashCache[ds] = dguardDashboard_(token, ds, ds)); };

    // 1) 近 7 天逐店日
    var days = dguardDays_(from, to);
    var bad = dguardCompare_(days, DGUARD.STORES, dguardOurs_(from, to), dashOf);

    // 2) 本月（到昨天）＋上月逐店總額；不一致才逐日掃那個月（7 天內已比過的日子跳過）
    var monthNotes = [];
    [[monthFrom, to], [prevFrom, prevTo]].forEach(function (m) {
      if (m[0] > m[1]) return;
      var mDash = dguardDashboard_(token, m[0], m[1]), mOurs = dguardOursRange_(m[0], m[1]);
      var badSt = DGUARD.STORES.filter(function (st) { return (mOurs[st] || 0) !== (mDash[st] ? mDash[st].amount : 0); });
      monthNotes.push(m[0].substring(0, 7) + (badSt.length ? ' 有 ' + badSt.length + ' 店對不上（' + badSt.join('、') + '）' : ' 12 店一致'));
      if (!badSt.length) return;
      var mDays = dguardDays_(m[0], m[1]).filter(function (d) { return days.indexOf(d) < 0; });
      if (!mDays.length) return;
      bad = bad.concat(dguardCompare_(mDays, badSt, dguardOurs_(mDays[0], mDays[mDays.length - 1]), dashOf));
    });

    var plan = dguardPlan_(bad), fixed = [], failed = [];
    var itemTasks = plan.repull.map(function (t) { return { ds: t.ds, st: t.st, mode: 'repull' }; })
      .concat(plan.adjust.map(function (t) { return { ds: t.ds, st: t.st, mode: 'adjust', amt: t.amt }; }));
    var overflow = itemTasks.slice(DGUARD.MAX_FIX_STOREDAYS);
    itemTasks = itemTasks.slice(0, DGUARD.MAX_FIX_STOREDAYS);

    if (dry) {
      plan.discDates.forEach(function (ds) { fixed.push('（預演）會重抓 ' + ds + ' 折扣明細覆蓋'); });
      itemTasks.forEach(function (t) { fixed.push('（預演）' + t.ds + ' ' + t.st + '店：' + (t.mode === 'repull' ? '整天重抓' : '補調整列 ' + t.amt)); });
    } else {
      plan.discDates.forEach(function (ds) {
        if (Date.now() - t0 > DGUARD.TIME_BUDGET_MS) { failed.push(ds + ' 折扣：時間不夠，明天再修'); return; }
        try { fixed.push(ds + ' 折扣重抓覆蓋 ' + dguardFixDiscounts_(token, ds, dashOf(ds))); }
        catch (e) { failed.push(ds + ' 折扣：' + e); }
      });
      if (itemTasks.length) {
        if (Date.now() - t0 > DGUARD.TIME_BUDGET_MS) itemTasks.forEach(function (t) { failed.push(t.ds + ' ' + t.st + '店：時間不夠，明天再修'); });
        else {
          try { fixed = fixed.concat(dguardFixItems_(token, itemTasks, dashOf)); }
          catch (e) { failed.push('品項修正失敗：' + e); }
        }
      }
    }
    overflow.forEach(function (t) { failed.push(t.ds + ' ' + t.st + '店：超過單次上限，明天再修'); });

    // 3) 修完再比一次
    var still = bad;
    if (!dry && bad.length) {
      var touched = {};
      bad.forEach(function (b) { touched[b.ds] = true; });
      var tds = Object.keys(touched).sort();
      var ours2 = dguardOurs_(tds[0], tds[tds.length - 1]);
      still = bad.filter(function (b) {
        var o = ours2[b.st + '|' + b.ds] || { g: 0, d: 0, n: 0 }, x = dashOf(b.ds)[b.st] || { gross: 0, discount: 0, amount: 0 };
        return o.n !== x.amount || o.g !== x.gross || o.d !== x.discount;
      });
    }

    var stillKey = {};
    still.forEach(function (b) { stillKey[b.ds + '|' + b.st] = true; });
    var detail = bad.map(function (b) {
      return [b.ds, b.st, b.o.n, b.x.amount, b.o.n - b.x.amount, dry ? '（預演）待修' : (stillKey[b.ds + '|' + b.st] ? '⚠️ 未修好' : '🔧 已自動修正')];
    });
    var status = !bad.length ? '✅ 一致' : (dry ? '（預演）有 ' + bad.length + ' 店日待修' : (still.length ? '⚠️ 有 ' + still.length + ' 店日未修好' : '🔧 已自動修正'));
    var range = days[0] + '～' + to + ' 逐日；' + monthNotes.join('；');
    var summary = [Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'), range, status, bad.length,
                   dry ? 0 : bad.length - still.length, dry ? 0 : still.length, fixed.concat(failed).join('；').substring(0, 45000)];

    Logger.log('[每日對帳] ' + status + '｜' + range + (bad.length ? '\n  ' + detail.map(function (d) { return d.join(' '); }).join('\n  ') : '') +
               (fixed.length ? '\n  處理：' + fixed.join('\n        ') : '') + (failed.length ? '\n  未處理：' + failed.join('\n          ') : ''));
    if (dry) return { status: status, bad: bad.length, plan: fixed };

    dguardWriteResult_(summary, detail);

    if (bad.length) {
      var props = PropertiesService.getScriptProperties(), seen = {};
      try { seen = JSON.parse(props.getProperty(DGUARD.PROP_ALERTED) || '{}'); } catch (e) { seen = {}; }
      var lines = ['POS 儀表板 vs 肚肚業績概況 每日對帳（' + range + '）', ''];
      if (!still.length) {
        lines.push('發現 ' + bad.length + ' 店日不一致，已全部自動修正，修完再比一次確認一致。不用處理。', '');
        detail.forEach(function (d) { lines.push('  ' + d[0] + ' ' + d[1] + '店：儀表板 ' + d[2] + ' vs 肚肚 ' + d[3] + '（差 ' + d[4] + '）'); });
        lines.push('', '處理：'); fixed.forEach(function (s) { lines.push('  ' + s); });
        dguardMail_('[POS對帳] ✅ 已自動修正 ' + bad.length + ' 店日（不用處理）', lines);
      } else {
        var fresh = still.filter(function (b) { return !seen[b.ds + '|' + b.st + '|' + (b.o.n - b.x.amount)]; });
        if (fresh.length) {
          lines.push('有 ' + still.length + ' 店日自動修正後仍不一致，明天 05:00 會再試一次。', '連續兩天都出現同一店日，請把這封信轉給 Claude。', '');
          still.forEach(function (b) { lines.push('  ' + b.ds + ' ' + b.st + '店：儀表板 ' + b.o.n + ' vs 肚肚 ' + b.x.amount + '（差 ' + (b.o.n - b.x.amount) + '）'); });
          if (fixed.length) { lines.push('', '已處理：'); fixed.forEach(function (s) { lines.push('  ' + s); }); }
          if (failed.length) { lines.push('', '未處理原因：'); failed.forEach(function (s) { lines.push('  ' + s); }); }
          dguardMail_('[POS對帳] ⚠️ ' + still.length + ' 店日自動修正後仍不一致', lines);
        }
      }
      var keep = {};
      still.forEach(function (b) { keep[b.ds + '|' + b.st + '|' + (b.o.n - b.x.amount)] = 1; });
      props.setProperty(DGUARD.PROP_ALERTED, JSON.stringify(keep));
    }
    return { status: status, bad: bad.length, still: still.length };
  } catch (e) {
    Logger.log('[每日對帳] 失敗：' + e + '\n' + (e && e.stack));
    if (!dry) {
      try { dguardWriteResult_([Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm'), '', '❌ 執行失敗', '', '', '', String(e).substring(0, 500)], []); } catch (e2) { }
      dguardMail_('[POS對帳] ❌ 每日對帳執行失敗（請把這封信轉給 Claude）', ['錯誤：' + e, '', (e && e.stack) || '']);
    }
    return { error: String(e) };
  } finally {
    if (lock) lock.releaseLock();
  }
}

/** 一次性：建立每日 05:00 觸發器（已存在就不重複建立）。 */
function dudooCheck_installTrigger() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'dudooCheck_daily'; });
  if (has) { Logger.log('dudooCheck_daily 觸發器已存在，不重複建立'); return '已存在'; }
  ScriptApp.newTrigger('dudooCheck_daily').timeBased().everyDays(1).atHour(5).nearMinute(0).inTimezone(DGUARD.TZ).create();
  Logger.log('✅ 已建立 dudooCheck_daily 每日 05:00 觸發器');
  return '已建立';
}
