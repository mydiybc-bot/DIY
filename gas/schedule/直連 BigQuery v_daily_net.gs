// ====================================================================
// bq_connector.gs — 直連 BigQuery v_daily_net（方案 B 單一真相）
// 用 Service Account 授權，查 diybc-make-sync.diybc_pos.v_daily_net
// ====================================================================

// ★ 把下載的 SA JSON 整包貼進 Script Properties，不要硬編在程式碼裡
//   設定 → 指令碼屬性 → 新增：key = SA_KEY_JSON, value = {整包JSON}
var BQ_PROJECT_ID = 'diybc-make-sync';

// ★ 2026-10-04 效能第 2 批：通行證（access token）存 CacheService 50 分鐘（token 本身 60 分鐘有效），
//   期間內直接重用、過期才重新簽章換 token；快取任何環節失敗都退回原本「每次重簽」的路。
var BQ_TOKEN_CACHE_KEY_ = 'BQ_SA_TOKEN_V1_';   // 後面接 SA 的 client_email，換金鑰就自然換 key
var BQ_TOKEN_CACHE_SEC_ = 50 * 60;

function _getBqAccessToken_() {
  var raw = PropertiesService.getScriptProperties().getProperty('SA_KEY_JSON');
  if (!raw) throw new Error('缺 SA_KEY_JSON 指令碼屬性');
  var sa = JSON.parse(raw);

  var cache = null;
  var cacheKey = BQ_TOKEN_CACHE_KEY_ + String(sa.client_email || '').replace(/[^\w.@\-]/g, '_').slice(0, 200);
  try {
    cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) return cached;
  } catch (eCache) { cache = null; }

  var now = Math.floor(Date.now() / 1000);
  var header = Utilities.base64EncodeWebSafe(JSON.stringify({alg:'RS256', typ:'JWT'}));
  var claim = Utilities.base64EncodeWebSafe(JSON.stringify({
    iss: sa.client_email,
    scope: 'https://www.googleapis.com/auth/bigquery.readonly',
    aud: 'https://oauth2.googleapis.com/token',
    exp: now + 3600,
    iat: now
  }));
  var signatureInput = header + '.' + claim;
  var signature = Utilities.computeRsaSha256Signature(signatureInput, sa.private_key);
  var jwt = signatureInput + '.' + Utilities.base64EncodeWebSafe(signature);

  var tokenRes = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    payload: {
      grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
      assertion: jwt
    },
    muteHttpExceptions: true
  });
  var tokenData = JSON.parse(tokenRes.getContentText());
  if (!tokenData.access_token) throw new Error('取 token 失敗: ' + tokenRes.getContentText());
  try {
    if (cache) {
      var life = Number(tokenData.expires_in) || 3600;                       // Google 回的有效秒數（通常 3599）
      var ttl = Math.min(BQ_TOKEN_CACHE_SEC_, Math.max(60, life - 600));    // 最多 50 分、至少留 10 分安全邊際
      cache.put(cacheKey, tokenData.access_token, ttl);
    }
  } catch (ePut) { }
  return tokenData.access_token;
}

/**
 * 查 v_daily_net，回傳 [{sale_date, store_code, gross, total_discount, net}, ...]
 * @param {string} startDate 'YYYY-MM-DD'（可選，不傳則全部）
 * @param {string} endDate   'YYYY-MM-DD'（可選）
 */
function queryDailyNet(startDate, endDate) {
  var token = _getBqAccessToken_();
  var where = '';
  if (startDate && endDate) {
    where = "WHERE sale_date BETWEEN '" + startDate + "' AND '" + endDate + "'";
  }
  var sql = 'SELECT FORMAT_DATE("%Y-%m-%d", sale_date) AS sale_date, ' +
            'store_code, gross, total_discount, net ' +
            'FROM `diybc-make-sync.diybc_pos.v_daily_net` ' + where +
            ' ORDER BY sale_date, store_code';

  var res = UrlFetchApp.fetch(
    'https://bigquery.googleapis.com/bigquery/v2/projects/' + BQ_PROJECT_ID + '/queries',
    {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ query: sql, useLegacySql: false, timeoutMs: 30000 }),
      muteHttpExceptions: true
    }
  );
  var data = JSON.parse(res.getContentText());
  if (data.error) throw new Error('BQ 查詢錯誤: ' + JSON.stringify(data.error));

  var out = [];
  if (data.rows) {
    for (var i = 0; i < data.rows.length; i++) {
      var f = data.rows[i].f;
      out.push({
        sale_date: f[0].v,
        store_code: Number(f[1].v),
        gross: Number(f[2].v),
        total_discount: Number(f[3].v),
        net: Number(f[4].v)
      });
    }
  }
  return out;
}

// 唯讀測試：跑這個確認直連通
function testBqConnection() {
  var rows = queryDailyNet('2026-05-31', '2026-05-31');
  Logger.log('回傳 ' + rows.length + ' 列');
  rows.forEach(function(r){
    Logger.log(r.sale_date + ' 店' + r.store_code + ' net=' + r.net);
  });
}