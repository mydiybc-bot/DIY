/**
 * dudooGuard.gs — 肚肚匯入守門員＋每日對帳（2026-10-05 新增）
 * 專案：「一鍵追加新品」（綁 BOM 本，script ID 1PkavxV6r3GZbEbS_B5ZWGL1ovm30T55XhllAuI1aglBe_RPERdP-DJnJ）
 * 經營者 2026-10-05 裁示：POS 儀表板要跟肚肚「業績概況」的營業額一致，「以後不要再發生這種問題」。
 *
 * 2026-10-05 查到兩種讓兩邊不一致的狀況：
 *   ① 跨日退款：肚肚把退款算在「退款那天」；舊匯入照匯出檔的建立日期把退款列記回「原交易日」（bq-sync 地雷 9）。
 *      例：5 店 2025-12-26 的單 2026-01-05 退 809 → 12 月、1 月各差 809。
 *   ② 幽靈品項：肚肚「銷售明細」匯出檔偶爾多一列帳單裡沒有的品項，看起來跟正常列一模一樣。
 *      例：2026-07-29 10 店第 4 張單，交易紀錄只收 644，匯出檔卻多一列檸檬雙重奏 680。
 *
 * 本檔做三件事：
 *   dudooGuard_apply_(rows, ds, token)：dudooPOS_loginAndImportDate_ 寫進 POS資料 之前呼叫。
 *     ① 建立日期不是匯入日的列（跨日退款）一律改記匯入日（＝肚肚的算法）。
 *     ② 逐店比「匯出檔銷售總額」vs「業績概況銷售總額」；對不上才讀「交易紀錄」逐張帳單比，
 *        找出帳單裡沒有的列。這些列加起來「剛好等於」差額才剔除，否則照原樣寫入並寄信。
 *     ②出任何錯都不擋匯入（照①處理後的資料寫入）並寄信。
 *   dudooCheck_daily()：每天 07:30 觸發。BigQuery v_daily_net 近 7 天逐店日 vs 業績概況，
 *     營業額／銷售總額／折扣任一對不上就寄信（同一店日同一差額只寄一次）。
 *     折扣對不上＝Make 折扣管線漏跑（bq-sync 地雷 13／15），要人工補。
 *   aaaDudooGuardDryRun()／aaaDudooCheckDryRun()：唯讀預演，不寫入、不寄信。
 */

var DGUARD = {
  TZ: 'Asia/Taipei',
  STORES: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],   // 13 西門已歇業、100 測試店不比
  CHECK_DAYS: 7,
  PROP_ALERTED: 'DGUARD_ALERTED',
  DRYRUN_DATES: ['2026-07-29', '2026-01-05', '2026-10-04']
};

/* ═══════════ 純計算（不呼叫任何 Google 服務，可單元測試） ═══════════ */

/** 匯出檔日期字串 → 'yyyy-MM-dd'；看不懂回 '' */
function dguardNormDate_(v) {
  var m = String(v || '').trim().match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

/** ① 跨日退款改記匯入日。直接改 rows[i][2]，回傳改了哪些列（說明文字） */
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

/**
 * ② 幽靈品項判定。
 *   rows：匯出檔欄位陣列（0 店號、4 交易序號、11 品名、18 單價、19 數量）
 *   dash：{店號: {cid, gross}}（業績概況）
 *   billsOf(cid)：回該店當日帳單陣列（交易紀錄：sale_code、sale_amount、discount）
 * 回傳 {rows: 剔除後的列, removed: [說明], unresolved: [{store, diff, why}]}
 */
function dguardGhost_(rows, dash, billsOf) {
  var res = { rows: rows, removed: [], unresolved: [] };
  var exp = {};
  rows.forEach(function (r) {
    var st = parseInt(r[0], 10);
    if (isNaN(st)) return;
    exp[st] = (exp[st] || 0) + (parseFloat(r[18]) || 0) * (parseFloat(r[19]) || 0);
  });
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

    var bySerial = {};   // 交易序號 → [列索引]
    rows.forEach(function (r, i) {
      if (parseInt(r[0], 10) !== st) return;
      var code = String(r[4] || '').replace(/\D/g, '');
      (bySerial[code] = bySerial[code] || []).push(i);
    });

    if (diff < 0) {   // 匯出檔少了品項：不知道少了什麼，無法自動補，只列出哪幾張帳單比匯出檔多
      var short = Object.keys(billGross).map(function (code) {
        var e = (bySerial[code] || []).reduce(function (a, i) { return a + (parseFloat(rows[i][18]) || 0) * (parseFloat(rows[i][19]) || 0); }, 0);
        return { code: code, gap: Math.round(billGross[code] - e) };
      }).filter(function (x) { return x.gap > 0; });
      res.unresolved.push({ store: st, diff: diff, why: '匯出檔比業績概況少，無法自動補' +
        (short.length ? '；帳單比匯出檔多：' + short.map(function (x) { return '第 ' + (billNo[x.code] || '?') + ' 張（' + x.code + '）' + x.gap; }).join('、') : '') });
      return;
    }

    var cand = [], candSum = 0;
    Object.keys(bySerial).forEach(function (code) {
      var idx = bySerial[code];
      var line = function (i) { return (parseFloat(rows[i][18]) || 0) * (parseFloat(rows[i][19]) || 0); };
      var eSum = idx.reduce(function (a, i) { return a + line(i); }, 0);
      if (billGross[code] === undefined) {
        // 帳單不存在：只有「整張都是正數量」才視為幽靈單（退款列的原帳單在別天，不算）
        var allPos = idx.every(function (i) { return (parseFloat(rows[i][19]) || 0) > 0; });
        if (allPos && eSum > 0) { cand.push({ code: code, idx: idx, amt: Math.round(eSum), why: '交易紀錄沒有這張單' }); candSum += Math.round(eSum); }
        return;
      }
      var excess = Math.round(eSum - billGross[code]);
      if (excess <= 0) return;
      var hit = -1;
      for (var j = 0; j < idx.length; j++) { if (Math.round(line(idx[j])) === excess) { hit = idx[j]; break; } }
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
      res.unresolved.push({ store: st, diff: diff, why: '比對交易紀錄後無法唯一認定：' +
        (cand.length ? cand.map(function (c) { return c.code + ' ' + c.amt + ' ' + c.why; }).join('；') : '每張帳單都對得上') });
    }
  });
  if (Object.keys(drop).length) res.rows = rows.filter(function (r, i) { return !drop[i]; });
  return res;
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

function dguardPost_(path, token, pairs) {
  var body = pairs.map(function (p) { return encodeURIComponent(p[0]) + '=' + encodeURIComponent(p[1]); }).join('&');
  var resp = UrlFetchApp.fetch(DUDOO_CONFIG_.API_BASE + path, {
    method: 'post', contentType: 'application/x-www-form-urlencoded', payload: body,
    headers: { 'Access-Token': token }, muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error(path + ' HTTP ' + resp.getResponseCode());
  return JSON.parse(resp.getContentText());
}

/** 業績概況（from～to 合計）→ {店號: {cid, amount, discount, gross}} */
function dguardDashboard_(token, from, to) {
  var pairs = [['access_token', token], ['start_date', from], ['end_date', to], ['hierarchy_id', DUDOO_CONFIG_.HIERARCHY_ID]];
  DUDOO_CONFIG_.COMPANY_IDS.forEach(function (c) { pairs.push(['company_id[]', c]); });
  var j = dguardPost_('/reports/getGroupDashboard?type=reports', token, pairs);
  if (!j || !j.data || !j.data.length) throw new Error('業績概況沒有回資料');
  var out = {};
  j.data.forEach(function (d) {
    var st = parseInt(d.company_name, 10);
    if (isNaN(st)) return;
    out[st] = { cid: String(d.company_id), amount: Math.round(Number(d.amount) || 0),
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

function dguardMail_(subject, lines) {
  try { MailApp.sendEmail(DUDOO_CONFIG_.NOTIFY_EMAIL, subject, lines.join('\n')); }
  catch (e) { Logger.log('[守門員] 寄信失敗：' + e); }
}

/* ═══════════ 匯入守門員（dudooPOS_loginAndImportDate_ 呼叫） ═══════════ */

function dudooGuard_apply_(rows, ds, token, dry) {
  var redated = dguardRedate_(rows, ds);
  if (redated.length) Logger.log('[守門員] 跨日退款改記 ' + ds + '：' + redated.length + ' 列\n  ' + redated.join('\n  '));
  try {
    var dash = dguardDashboard_(token, ds, ds);
    var res = dguardGhost_(rows, dash, function (cid) { return dguardBills_(token, cid, ds); });
    if (!res.removed.length && !res.unresolved.length) {
      Logger.log('[守門員] ' + ds + ' 12 店銷售總額＝業績概況 ✅');
      return res.rows;
    }
    var lines = ['肚肚匯入守門員 ' + ds, ''];
    if (res.removed.length) {
      lines.push('已剔除匯出檔多出來的品項（帳單裡沒有，剔除後該店＝業績概況）：');
      res.removed.forEach(function (s) { lines.push('  ' + s); });
      lines.push('');
    }
    if (res.unresolved.length) {
      lines.push('⚠️ 對不上、無法自動處理（照原樣寫入，請人工確認）：');
      res.unresolved.forEach(function (u) { lines.push('  ' + u.store + '店 匯出檔−業績概況＝' + u.diff + '：' + u.why); });
      lines.push('');
    }
    if (redated.length) { lines.push('跨日退款改記 ' + ds + '：'); redated.forEach(function (s) { lines.push('  ' + s); }); }
    Logger.log(lines.join('\n'));
    if (!dry) dguardMail_('[POS守門] ' + ds + (res.unresolved.length ? ' ⚠️ 有店對不上' : ' 已剔除幽靈品項'), lines);
    return res.rows;
  } catch (e) {
    Logger.log('[守門員] ⚠️ 比對失敗（照原資料匯入）：' + e);
    if (!dry) dguardMail_('[POS守門] ' + ds + ' 比對失敗（資料照常匯入）', ['肚肚匯入守門員 ' + ds, '', '錯誤：' + e,
      '', '資料已照常寫入（跨日退款改記 ' + redated.length + ' 列）。07:30 的每日對帳會再比一次。']);
    return rows;
  }
}

/** 唯讀預演：登入肚肚、抓 DRYRUN_DATES 每天的匯出檔跑守門員，只寫執行記錄（不寫試算表、不寄信）。 */
function aaaDudooGuardDryRun() {
  var token = dguardLogin_();
  DGUARD.DRYRUN_DATES.forEach(function (ds) {
    var payload = 'hierarchy_id=' + DUDOO_CONFIG_.HIERARCHY_ID + '&start_date=' + ds + '&end_date=' + ds;
    DUDOO_CONFIG_.COMPANY_IDS.forEach(function (c) { payload += '&company_id%5B%5D=' + c; });
    var resp = UrlFetchApp.fetch(DUDOO_CONFIG_.API_BASE + '/reports/getSaleDetailsAnalysis/export', {
      method: 'post', contentType: 'application/x-www-form-urlencoded', payload: payload,
      headers: { 'Access-Token': token }, muteHttpExceptions: true });
    var rows = dudooPOS_parseCSV_(resp.getContentText());
    var before = rows.length;
    var out = dudooGuard_apply_(rows, ds, token, true);
    Logger.log('[預演] ' + ds + '：匯出 ' + before + ' 列 → 寫入 ' + out.length + ' 列');
  });
  Logger.log('[預演] 結束：沒有寫入試算表、沒有寄信');
}

/* ═══════════ 每日對帳（07:30 觸發） ═══════════ */

function dudooCheck_daily() { return dguardCheck_(false); }

/** 唯讀預演：照常比對並寫執行記錄，不寄信、不記「已通知」。 */
function aaaDudooCheckDryRun() { return dguardCheck_(true); }

function dguardCheck_(dry) {
  var tz = DGUARD.TZ, days = [];
  for (var i = DGUARD.CHECK_DAYS; i >= 1; i--) days.push(Utilities.formatDate(new Date(Date.now() - i * 86400000), tz, 'yyyy-MM-dd'));
  try {
    var token = dguardLogin_();
    var bq = audit_bqAll_('SELECT store_code, CAST(sale_date AS STRING), CAST(SUM(gross) AS INT64), ' +
      'CAST(SUM(total_discount) AS INT64), CAST(SUM(net) AS INT64) FROM `diybc-make-sync.diybc_pos.v_daily_net` ' +
      "WHERE sale_date BETWEEN '" + days[0] + "' AND '" + days[days.length - 1] + "' GROUP BY 1, 2");
    var ours = {};
    bq.forEach(function (r) { ours[parseInt(r[0], 10) + '|' + r[1]] = { g: Number(r[2]) || 0, d: Number(r[3]) || 0, n: Number(r[4]) || 0 }; });

    var bad = [];
    days.forEach(function (ds) {
      var dash = dguardDashboard_(token, ds, ds);
      DGUARD.STORES.forEach(function (st) {
        var o = ours[st + '|' + ds] || { g: 0, d: 0, n: 0 };
        var x = dash[st] || { gross: 0, discount: 0, amount: 0 };
        if (o.n !== x.amount || o.g !== x.gross || o.d !== x.discount) {
          var why = o.g !== x.gross ? (o.d !== x.discount ? '品項與折扣都不同' : '品項金額不同（匯入／匯出檔問題）')
                                    : '折扣不同（Make 折扣管線，地雷 13／15）';
          bad.push({ key: st + '|' + ds + '|' + (o.n - x.amount), text: ds + ' ' + st + '店：儀表板 ' + o.n + '（銷售 ' + o.g + '／折扣 ' + o.d +
            '） vs 肚肚 ' + x.amount + '（銷售 ' + x.gross + '／折扣 ' + x.discount + '），差 ' + (o.n - x.amount) + '；' + why });
        }
      });
    });

    var props = PropertiesService.getScriptProperties();
    var seen = {};
    try { seen = JSON.parse(props.getProperty(DGUARD.PROP_ALERTED) || '{}'); } catch (e) { seen = {}; }
    var fresh = bad.filter(function (b) { return !seen[b.key]; });

    Logger.log('[每日對帳] ' + days[0] + '～' + days[days.length - 1] + '：不一致 ' + bad.length + ' 店日（新 ' + fresh.length + '）' +
      (bad.length ? '\n  ' + bad.map(function (b) { return b.text; }).join('\n  ') : ' ✅'));
    if (dry) return { checked: days.length * DGUARD.STORES.length, bad: bad.length, fresh: fresh.length };

    if (fresh.length) {
      var lines = ['POS 儀表板 vs 肚肚業績概況 每日對帳（' + days[0] + '～' + days[days.length - 1] + '，12 店逐日）', '',
                   '新發現的不一致：'];
      fresh.forEach(function (b) { lines.push('  ' + b.text); });
      if (bad.length > fresh.length) lines.push('', '（另有 ' + (bad.length - fresh.length) + ' 店日之前已通知、尚未修好）');
      lines.push('', '同一店日同一差額只通知一次。');
      dguardMail_('[POS對帳] 儀表板與肚肚不一致 ' + fresh.length + ' 店日', lines);
    }
    var keep = {};
    bad.forEach(function (b) { keep[b.key] = 1; });   // 只記窗內仍存在的差異，修好或滑出 7 天就自動忘掉
    props.setProperty(DGUARD.PROP_ALERTED, JSON.stringify(keep));
    return { checked: days.length * DGUARD.STORES.length, bad: bad.length, fresh: fresh.length };
  } catch (e) {
    Logger.log('[每日對帳] 失敗：' + e);
    if (!dry) dguardMail_('[POS對帳] 每日對帳執行失敗 ' + days[days.length - 1], ['錯誤：' + e, '', (e && e.stack) || '']);
    return { error: String(e) };
  }
}

/** 一次性：建立每日 07:30 觸發器（已存在就不重複建）。 */
function dudooCheck_installTrigger() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'dudooCheck_daily'; });
  if (has) { Logger.log('dudooCheck_daily 觸發器已存在，不重複建立'); return '已存在'; }
  ScriptApp.newTrigger('dudooCheck_daily').timeBased().everyDays(1).atHour(7).nearMinute(30).inTimezone(DGUARD.TZ).create();
  Logger.log('✅ 已建立 dudooCheck_daily 每日 07:30 觸發器');
  return '已建立';
}
