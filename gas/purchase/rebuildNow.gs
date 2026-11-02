/** rebuildNow.gs v1.0 — 立即重算 API（2026-08-25 q 批）
 * 用途：前端總覽頁「🔄 立即重算」按鈕的後端。
 * 部署位置：diybc-purchase-agg 專案（第 7 支檔案；必須在本專案內，才呼叫得到
 *          rebuildPurchaseAgg / rebuildEventLayer 本尊）。
 * ⚠️ 貼上前務必全專案搜尋「doGet」——必須查無結果才能貼；若已存在 doGet，停手回報。
 *   （doPost 完全不動，與既有 inv/param/sku/season/conv/booking/order/event 寫入無關）
 * 全域一律 RB_ 前綴（Apps Script 各檔共用全域）；引用 Code.gs 既有常數 TOKEN。
 * 流程：doGet(action=rebuild) 驗 token → 排入 1 秒後一次性觸發器 RB_run → 立即回應
 *      {ok:1,id}。RB_run 跑 rebuildPurchaseAgg＋rebuildEventLayer，狀態寫
 *      ScriptProperties，前端以 doGet(action=status) 輪詢、比對 id。
 * 注意：月度聚合 rebuildMonthAgg 不在此（僅影響查詢頁月報，維持每日 07:40）。
 * v1.1（2026-11-02）兩個問題一起修：
 *   ① Google 一次性觸發器幾乎不跑（10/27～11/2 七天只跑過 1 次；11/2 16:28 排入 20 分鐘都沒跑）
 *      → 加巡檢 RB_tick（每 1 分鐘；編輯器執行 RB_setupTick 建一次）：排入超過 1 分鐘沒開跑就由它接手。
 *   ② 兩段重算塞在同一次執行會撞 Google 6 分鐘上限（rebuildPurchaseAgg 104～269 秒＋rebuildEventLayer
 *      59～117 秒，最慢 386 秒；11/2 17:02 實際逾時）→ 拆兩段：第一段跑完已用不到 3 分鐘就接著跑第二段，
 *      否則標「mid」交給下一次巡檢（另開一次執行，重新有 6 分鐘）。
 *   狀態：queued → running → (mid → running2) → done／error。每段開頭短暫上鎖看狀態，觸發器與巡檢同時到也只跑一次。
 *   一段跑超過 7 分鐘還沒結束＝被 Google 中斷 → 巡檢標 error，前端會顯示失敗、可以再按一次。
 */
var RB_PROP_KEY = 'RB_STATE';
var RB_STALE_MS = 15 * 60 * 1000;   // 排入/執行中超過 15 分鐘視為僵死，允許重新排入（v1.1：兩段＋巡檢，6→15 分鐘）
var RB_TICK_GRACE_MS = 60 * 1000;   // v1.1：排入超過 1 分鐘一次性觸發器還沒跑 → 巡檢接手
var RB_SEG_DEAD_MS = 7 * 60 * 1000; // v1.1：一段跑超過 7 分鐘＝已被 Google 6 分鐘上限中斷
var RB_INLINE_MS = 3 * 60 * 1000;   // v1.1：第一段跑完已用不到 3 分鐘 → 同一次執行接著跑第二段（最慢 117 秒）
var RB_BUSY = { queued: 1, running: 1, mid: 1, running2: 1 };

function doGet(e){
  var p = (e && e.parameter) || {};
  var out;
  if (String(p.token || '') !== TOKEN) {
    out = { ok: 0, msg: 'token' };
  } else if (p.action === 'rebuild') {
    out = RB_queue_();
  } else if (p.action === 'status') {
    out = RB_status_();
  } else {
    out = { ok: 0, msg: 'unknown action' };
  }
  return RB_json_(out, p.callback);
}

function RB_json_(obj, cb){
  var txt = JSON.stringify(obj);
  if (cb && /^[A-Za-z0-9_$]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + txt + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

function RB_state_(){
  try {
    return JSON.parse(PropertiesService.getScriptProperties().getProperty(RB_PROP_KEY) || '{}');
  } catch (err) { return {}; }
}
function RB_save_(st){
  PropertiesService.getScriptProperties().setProperty(RB_PROP_KEY, JSON.stringify(st));
}


function RB_queue_(){
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: 0, busy: 1, msg: 'lock' };
  try {
    var st = RB_state_(), now = Date.now();
    if (RB_BUSY[st.status] && now - (st.id || 0) < RB_STALE_MS) {
      return { ok: 0, busy: 1, age: Math.round((now - st.id) / 1000) };
    }
    RB_clearTriggers_();
    ScriptApp.newTrigger('RB_run').timeBased().after(1000).create();   // 快的話 1 分鐘內跑；沒跑由 RB_tick 接手
    RB_save_({ status: 'queued', id: now });
    return { ok: 1, id: now };
  } finally {
    lock.releaseLock();
  }
}

function RB_status_(){
  var st = RB_state_();
  return { ok: 1, status: st.status || 'idle', id: st.id || 0, doneAt: st.doneAt || 0, msg: st.msg || '' };
}

function RB_clearTriggers_(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'RB_run') ScriptApp.deleteTrigger(t);
  });
}

/* v1.1：上鎖看狀態，符合 from 才改成 to 並記開始時間；回傳 null＝不該由這一路跑。
   from 為 null＝編輯器手動執行，除了另一路正在跑以外一律可跑。 */
function RB_claim_(from, to){
  var lock = LockService.getScriptLock();   // 只鎖這一下，重算本身不佔鎖（不擋盤點寫入）
  if (!lock.tryLock(10000)) return null;
  try {
    var st = RB_state_(), now = Date.now();
    var live = (st.status === 'running' || st.status === 'running2') && now - (st.segAt || st.id || 0) < RB_SEG_DEAD_MS;
    if (from ? st.status !== from : live) return null;
    if (!st.id) st.id = now;
    st.status = to; st.segAt = now;
    RB_save_(st);
    return st;
  } finally {
    lock.releaseLock();
  }
}

/* 第一段：rebuildPurchaseAgg；時間夠就接著跑第二段。觸發器、巡檢、編輯器手動都走這裡。 */
function RB_run(){
  var t0 = Date.now();
  var st = RB_claim_(RB_state_().status === 'queued' ? 'queued' : null, 'running');
  if (!st) return;
  RB_clearTriggers_();               // 先自清一次性觸發器，失敗也不殘留
  try {
    rebuildPurchaseAgg();
  } catch (err) {
    st.status = 'error'; st.doneAt = Date.now(); st.msg = String((err && err.message) || err);
    RB_save_(st);
    return;
  }
  st.status = 'mid'; st.segAt = Date.now();
  RB_save_(st);
  if (Date.now() - t0 < RB_INLINE_MS) RB_run2();
}

/* 第二段：rebuildEventLayer。 */
function RB_run2(){
  var st = RB_claim_('mid', 'running2');
  if (!st) return;
  try {
    rebuildEventLayer();
    st.status = 'done'; st.doneAt = Date.now(); st.msg = '';
  } catch (err) {
    st.status = 'error'; st.doneAt = Date.now(); st.msg = String((err && err.message) || err);
  }
  RB_save_(st);
}

/* v1.1：巡檢（每 1 分鐘）。沒事時只讀一次狀態就結束。 */
function RB_tick(){
  var st = RB_state_(), now = Date.now();
  if (st.status === 'queued' && now - (st.id || 0) >= RB_TICK_GRACE_MS) { RB_run(); return; }
  if (st.status === 'mid') { RB_run2(); return; }
  if ((st.status === 'running' || st.status === 'running2') && now - (st.segAt || st.id || 0) >= RB_SEG_DEAD_MS) {
    st.msg = '重算超過 6 分鐘被 Google 中斷（' + (st.status === 'running2' ? '檔期' : '補貨') + '那段），請再按一次大更新';
    st.status = 'error'; st.doneAt = now;
    RB_save_(st);
  }
}

/* v1.1：在編輯器執行一次即可（重複執行不會疊加）。 */
function RB_setupTick(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if (t.getHandlerFunction() === 'RB_tick') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('RB_tick').timeBased().everyMinutes(1).create();
}
