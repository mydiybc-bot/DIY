/* ================= 🎬 步驟影片批次上傳（2026-11-01） =================
   經營者 11/01：每個月大量步驟影片 → Claude 在本機剪好（資料交換/工具/食譜影片剪輯），再「自動上傳」到自己做食譜系統。
   做法：本機一步送一次（POST，影片 base64），這裡用本專案的後台登入（rb_backend.gs）打開那一步的「編輯步驟」頁，
   **只換圖片／影片**：標題、說明、食材工具、容器、計時器照頁面原樣送回（食材列不刪不加），送出後再讀一次逐項核對。
   安全：
     ‧ 只上傳到「停用」的食譜（試做版）；啟用中的一律擋（客人正在用的食譜，照 10/17 規定要先問經營者）。
     ‧ 第幾步以後台步驟清單為準；本機可帶該步標題，對不上就停（避免放錯步）。
     ‧ 格式／大小沿用 nrMediaBlob_（MP4 ≤ 5MB、圖片 ≤ 300KB、檢查檔頭）。
     ‧ 同一支食譜同時只跑一個請求（CacheService 擋重複）。
     ‧ 不新增、不刪除後台任何步驟或食譜；每次寫入記在 fact_recipe_import（req_id＝VU-…，imp_id＝VU）。
   身分：不用申請表身分密碼（Claude 沒有），改用專用通行碼——本機 ~/.config/diybc/video-upload.token（不進版控），
         這裡只放它的 SHA-256（公開程式庫也安全）。要換：本機重產 token，改 VU_TOKEN_SHA 再部署。
   動作（都掛 NOAUTH，函式自己驗通行碼）：
     vuFind（只讀）：名稱（包含）→ 後台繁中食譜清單 [{id,title,status}]
     vuPeek（只讀）：食譜編號 → 名稱、啟用、公開、步驟清單 [{n,sid,title}]
     vuUpload（寫後台那一步的圖片／影片）：{bid, n, title?, type, b64, vu_id, dry?}；dry＝只檢查不送 */
var VU_TOKEN_SHA = '4d803ba819fc5b308e4c2673c83ae96049f36c65268780ad8f6d922f78aa3cf0';
var VU_ZH = '42575226-23b1-4ac3-b123-3cba50e3cee8';   /* 繁體中文 */
var VU_GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
var VU_WHO = { kind: 'token', role: 'Claude 影片上傳' };

function vuHex_(bytes) { return bytes.map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function vuAuth_(p) {
  var t = String((p && p.token) || '');
  var ok = t.length >= 20 && vuHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, t, Utilities.Charset.UTF_8)) === VU_TOKEN_SHA;
  if (!ok) { Utilities.sleep(1000); throw fail_('影片上傳通行碼不對', 'auth'); }
}
function vuBid_(p) { var b = str_(p.bid).trim(); if (!VU_GUID.test(b)) throw fail_('食譜編號格式不對：' + b, 'arg'); return b; }
function vuNorm_(s) { return String(s || '').replace(/\r\n/g, '\n').replace(/[ \s]+/g, ' ').trim(); }

function vuFind_(p) {
  vuAuth_(p);
  var title = str_(p.title).trim();
  if (!title) throw fail_('請給食譜名稱', 'arg');
  var sess = rbSession_();
  var r = rbAuthedGet_(sess, '/Recipes?Title=' + encodeURIComponent(title) + '&LanguageId=' + encodeURIComponent(VU_ZH) + '&GroupId=');
  if (r.code !== 200) throw fail_('讀不到後台食譜清單（HTTP ' + r.code + '）', 'rbnet');
  rbSaveSession_(sess);
  var rows = rbListRows_(r.text).map(function (x) { return { id: x.id, title: x.title, status: x.status, exact: rbTitleNorm_(x.title) === rbTitleNorm_(title), link: rbLink_(x.id) }; });
  return { id: '', data: rows };
}

function vuPeek_(p) {
  vuAuth_(p);
  var bid = vuBid_(p), sess = rbSession_();
  var pg = rbRecipeEditPage_(sess, bid), idx = rbStepIndex_(sess, bid);
  rbSaveSession_(sess);
  return { id: bid, data: { bid: bid, title: rbFieldGet_(pg.fields, 'Title') || '', inUse: pg.inUse, pub: pg.pub, link: rbLink_(bid),
    steps: idx.map(function (s, i) { return { n: i + 1, sid: s.sid, title: s.title }; }) } };
}

/* 讀一步：解析＋原本的說明 HTML */
function vuStepRead_(sess, sid) {
  var r = rbAuthedGet_(sess, '/Steps/Edit/' + encodeURIComponent(sid));
  var pg = rbStepEditParse_(r, sid), f = rbFormHtml_(r.text, /^\/Steps\/Edit\//i);
  pg.content = rbFieldGet_(rbFormFields_(f), 'Content') || '';
  return pg;
}
/* 只換圖片／影片：其他欄位照原樣送回；deletedIngred 空白、不送 Ingredients＝食材列不動 */
function vuPost_(sess, pg, blob) {
  var payload = {
    '__RequestVerificationToken': pg.token, RecipeId: pg.rid, StepId: pg.sid, Order: pg.order,
    StepMedia: pg.media, StepTitle: pg.title, Content: pg.content,
    StopClock: pg.stop ? 'true' : 'false', Timer: pg.timer || '', deletedIngred: '', Image: blob
  };
  var res = rbReq_(sess, 'post', '/Steps/Edit/' + encodeURIComponent(pg.sid), { payload: payload });
  if (rbIsLogin_(res)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  if (res.code === 302 || res.code === 301) return true;
  if (res.code === 200) throw rbErr_('食譜後台沒有接受這一步的影片：' + (rbFormErrors_(res.text) || '表單有欄位不合格'), 'rbreject');
  throw rbErr_('食譜後台回覆異常（HTTP ' + res.code + '）', 'rbunknown');
}
function vuMediaBytes_(url) {
  try {
    var r = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, followRedirects: true, headers: { Range: 'bytes=0-0' } });
    var h = r.getAllHeaders(), cr = String(h['Content-Range'] || h['content-range'] || '');
    var m = cr.match(/\/(\d+)$/);
    if (m) return +m[1];
    if (r.getResponseCode() === 200) return r.getBlob().getBytes().length;
  } catch (e) {}
  return -1;
}

function vuUpload_(p) {
  vuAuth_(p);
  var bid = vuBid_(p), n = parseInt(p.n, 10), dry = !!p.dry, vid = str_(p.vu_id).trim() || ('VU-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss'));
  if (!(n >= 1)) throw fail_('第幾步要是 1 以上的數字', 'arg');
  if (!/^VU-[\w-]{1,60}$/.test(vid)) throw fail_('上傳編號格式不對', 'arg');
  var cache = CacheService.getScriptCache(), ck = 'vu:busy:' + bid;
  if (!dry) { if (cache.get(ck)) throw fail_('這支食譜正在上傳另一步，請等上一步完成再送', 'busy'); cache.put(ck, '1', 180); }
  try {
    var sess = rbSession_(), relogin = function () { rbDropSession_(); sess = rbFreshSession_(); };
    var rp = rbRecipeEditPage_(sess, bid), rtitle = rbFieldGet_(rp.fields, 'Title') || '';
    if (rp.inUse) throw fail_('「' + rtitle + '」是啟用中的食譜，影片只能自動上傳到停用的試做版（啟用中的要先問經營者）', 'active');
    var idx = rbStepIndex_(sess, bid), s = idx[n - 1];
    if (!s) throw fail_('「' + rtitle + '」後台只有 ' + idx.length + ' 步，沒有第 ' + n + ' 步', 'nostep');
    if (str_(p.title).trim() && rbTitleNorm_(s.title) !== rbTitleNorm_(p.title)) throw fail_('第 ' + n + ' 步標題對不上：後台「' + s.title + '」、本機「' + p.title + '」，先停下來', 'mismatch');
    var md = nrMediaBlob_({ i: n - 1, type: p.type, b64: p.b64 }, n - 1);
    var pg = vuStepRead_(sess, s.sid);
    if (pg.rid !== bid) throw fail_('後台第 ' + n + ' 步不是這支食譜的（頁面可能改版）', 'mismatch');
    var before = { title: pg.title, content: vuNorm_(pg.content), ings: pg.ings.slice().sort().join(','), timer: pg.timer, stop: pg.stop, media: pg.mediaUrl };
    var base = { n: n, sid: s.sid, title: s.title, recipe: rtitle, bytes: md.bytes, type: md.type, old_media: pg.mediaUrl, ings: pg.ings.length };
    if (dry) { rbSaveSession_(sess); base.dry = true; return { id: vid, data: base }; }
    try { vuPost_(sess, pg, md.blob); }
    catch (e) { if (e && e.code === 'rbrelogin') { relogin(); pg = vuStepRead_(sess, s.sid); vuPost_(sess, pg, md.blob); } else throw e; }
    /* 核對：再讀一次，只有圖片／影片可以變 */
    var aft = vuStepRead_(sess, s.sid), bad = [];
    if (aft.title !== before.title) bad.push('標題');
    if (vuNorm_(aft.content) !== before.content) bad.push('說明');
    if (aft.ings.slice().sort().join(',') !== before.ings) bad.push('食材工具');
    if (aft.timer !== before.timer || aft.stop !== before.stop) bad.push('計時器');
    if (!aft.mediaUrl) bad.push('圖片影片沒有存上');
    var got = aft.mediaUrl ? vuMediaBytes_(aft.mediaUrl) : -1;
    if (got >= 0 && got !== md.bytes) bad.push('雲端檔案大小 ' + got + '≠上傳 ' + md.bytes);
    rbSaveSession_(sess);
    base.media = aft.mediaUrl; base.size_check = got; base.verified = !bad.length;
    rbLog_(vid, 'VU', bad.length ? 'fail' : 'media', bid, String(n), bad.length ? 'N' : 'Y',
      (bad.length ? '核對不符：' + bad.join('、') + '｜' : '') + '第 ' + n + ' 步「' + s.title + '」換' + (md.type.indexOf('video') === 0 ? '影片' : '圖片') + ' ' + md.bytes + ' bytes（' + rtitle + '）', VU_WHO);
    if (bad.length) throw fail_('第 ' + n + ' 步上傳後核對不符：' + bad.join('、') + '。請食譜負責夥伴看一下後台這一步', 'verify', base);
    return { id: vid, data: base };
  } finally { if (!dry) cache.remove(ck); }
}

/* ---------- 掛上 API（本檔排在 Code.js 後面；包一層＝呼叫時才找函式） ---------- */
ACTIONS.vuFind = function (p) { return vuFind_(p); };
ACTIONS.vuPeek = function (p) { return vuPeek_(p); };
ACTIONS.vuUpload = function (p) { return vuUpload_(p); };
NOAUTH.vuFind = 1; NOAUTH.vuPeek = 1; NOAUTH.vuUpload = 1;   /* 函式自己驗通行碼 */
READS.vuFind = 1; READS.vuPeek = 1;                           /* 只讀：不排全系統的鎖 */
SELFLOCK.vuUpload = 1;                                        /* 要連後台、一次 10～30 秒：自己擋同一支食譜重複，不佔全系統的鎖 */
