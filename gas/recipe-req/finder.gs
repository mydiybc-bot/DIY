/* ================= 🗂 找食譜：勾選的後台食譜下載成 Word（2026-10-18） =================
   經營者 10/18：「後台食譜類別很亂，在新品申請表做一個功能重新整理類別、輕易找到甜點食譜，還要能勾選食譜下載」。
   分類、搜尋都在前端做（資料沿用 🧬 食譜演化的 lineage）；這裡只負責「下載」：
   用後台本來就有的「批次匯出勾選食譜」（POST /Recipes/ExportRecipesZip，ids＝逗號分隔的食譜 ID）→ 回傳 Word 壓縮檔（每支一個 .doc）。
   ⚠ 只讀：這個後台網址只產生檔案，不改、不刪後台任何資料（10/18 實測不需要驗證碼；2 支 0.4 秒、60 支 16 秒、約 5.6KB／支，圖片是連結不內嵌）。
   ⚠ 要身分密碼（任何身分都可以；完整食譜內容不給只知道網址的人），每次下載寫一筆紀錄（哪個身分、幾支、哪些 ID）。
   ⚠ 一次最多 FD_MAX 支（前端自動分批再合成一個壓縮檔）。 */
var FD_MAX = 40;
var FD_ID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function fdExport_(sess, ids) {
  var r = rbReq_(sess, 'post', '/Recipes/ExportRecipesZip', { payload: { ids: ids.join(',') }, binary: true });
  if (rbIsLogin_(r)) throw rbErr_('食譜後台登入逾時', 'rbrelogin');
  return r;
}
function rcpZip_(p, auth) {
  var seen = {}, ids = (Array.isArray(p.ids) ? p.ids : String(p.ids || '').split(','))
    .map(function (s) { return String(s || '').trim().toLowerCase(); })
    .filter(function (s) { if (!FD_ID.test(s) || seen[s]) return false; seen[s] = 1; return true; });
  if (!ids.length) throw fail_('請先勾選要下載的食譜', 'bad');
  if (ids.length > FD_MAX) throw fail_('一次最多下載 ' + FD_MAX + ' 支（頁面會自動分批）', 'bad');
  var sess = rbSession_(), r;
  try { r = fdExport_(sess, ids); }
  catch (e) {
    if (!e || e.code !== 'rbrelogin') throw e;
    rbDropSession_(); sess = rbFreshSession_(); r = fdExport_(sess, ids);   /* 登入逾時 → 重新登入再試一次（這個網址不改資料，重送無害） */
  }
  var h = r.resp.getAllHeaders(), ct = String(h['Content-Type'] || h['content-type'] || '');
  if (r.code !== 200 || !/zip|octet-stream/i.test(ct)) throw rbErr_('食譜後台沒有給檔案（HTTP ' + r.code + '，' + (ct || '沒有檔案類型') + '）', 'rbnet');
  var bytes = r.resp.getContent();
  if (!bytes || bytes.length < 22) throw rbErr_('食譜後台給的壓縮檔是空的', 'rbnet');
  rbSaveSession_(sess);
  return { id: '', msg: ids.length + ' 支：' + ids.join(',').slice(0, 1500), data: { b64: Utilities.base64Encode(bytes), bytes: bytes.length, count: ids.length } };
}
ACTIONS.rcpZip = function (p, a) { return rcpZip_(p, a); };
READS.rcpZip = 1;          /* 不排鎖（後台匯出 40 支約 10 秒，不能卡住別人送件） */
AUDIT_READS.rcpZip = 1;    /* 成功也寫一筆紀錄：誰、幾支、哪些食譜 */
/* 編輯器手動驗證：下載 2 支測試（只讀），看紀錄的大小 */
function fdTestNow() {
  var r = rcpZip_({ ids: ['6abd0c79-4257-4321-8e4c-970eb78538f4', '3762cdd1-2073-4e82-a89f-9f3f3477bdc9'] }, {});
  Logger.log(JSON.stringify({ bytes: r.data.bytes, count: r.data.count, zipHead: r.data.b64.slice(0, 4) }));
}
