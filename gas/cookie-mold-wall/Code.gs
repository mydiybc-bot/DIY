/**
 * 餅乾模範例牆 API（2026-11-18 新建；AI 專屬餅乾模・線上試玩 static/cookie-mold.html 的「範例牆」後台）
 *
 * 用法（給經營者）：把圖片丟進雲端硬碟資料夾「餅乾模範例牆」就會出現在試玩頁，最慢 5 分鐘更新。
 *   檔名規則：編號-種類-標題.jpg　例：1-照片-柴犬麻糬.jpg、1-模具.jpg、1-餅乾.jpg
 *     編號：同一組用同一個數字，數字小的排前面
 *     種類：照片／模具／餅乾（三張不必到齊，缺的那格顯示圖示）
 *     標題：可不寫；寫在任一張上即可
 *   要下架：把檔案移出資料夾，或檔名最前面加「x」。
 *
 * 讀取：?action=list（可加 &callback=函式名）→ { ok, items:[{no,title,photo,mold,cookie}], folder }
 * 第一次使用：在編輯器執行 setup()（建立資料夾、設成「知道連結的人可檢視」、放一份說明）。
 * 這支程式只讀資料夾裡的檔名與檔案編號，不碰雲端硬碟其他地方。
 */
var VER = '2026-11-18a';
var FOLDER_NAME = '餅乾模範例牆';
var KIND = { '照片': 'photo', '模具': 'mold', '餅乾': 'cookie' };

function doGet(e) {
  var p = (e && e.parameter) || {}, o;
  try {
    if (p.action === 'list') o = list_();
    else o = { ok: true, ver: VER };
  } catch (err) { o = { ok: false, msg: String(err && err.message || err) }; }
  if (p.callback && /^[A-Za-z_$][\w$]*$/.test(p.callback)) {
    return ContentService.createTextOutput(p.callback + '(' + JSON.stringify(o) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function folder_() {
  var id = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!id) throw new Error('還沒執行 setup()');
  return DriveApp.getFolderById(id);
}

function list_() {
  var cache = CacheService.getScriptCache(), hit = cache.get('list');
  if (hit) return JSON.parse(hit);
  var it = folder_().getFiles(), groups = {};
  while (it.hasNext()) {
    var f = it.next(), name = f.getName();
    if (!/^image\//.test(f.getMimeType())) continue;
    var m = name.match(/^(\d+)\s*[-_－—–\s]\s*(照片|模具|餅乾)(?:\s*[-_－—–\s]\s*(.+?))?\.[A-Za-z0-9]+$/);
    if (!m) continue;
    var no = parseInt(m[1], 10), g = groups[no] || (groups[no] = { no: no, title: '' });
    g[KIND[m[2]]] = 'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w800';
    if (m[3] && !g.title) g.title = m[3];
  }
  var items = Object.keys(groups).map(function (k) { return groups[k]; }).sort(function (a, b) { return a.no - b.no; });
  items.forEach(function (g) { if (!g.title) g.title = '範例 ' + g.no; });
  var o = { ok: true, ver: VER, items: items };
  cache.put('list', JSON.stringify(o), 300);
  return o;
}

/** 第一次使用時在編輯器執行一次：建立資料夾並記住它。重複執行不會多建。 */
function setup() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('FOLDER_ID'), folder;
  if (id) folder = DriveApp.getFolderById(id);
  else {
    folder = DriveApp.createFolder(FOLDER_NAME);
    props.setProperty('FOLDER_ID', folder.getId());
    folder.createFile('＿使用說明＿.txt',
      '把圖片丟進這個資料夾，就會出現在「AI 專屬餅乾模」試玩頁的範例牆，最慢 5 分鐘更新。\n\n' +
      '檔名規則：編號-種類-標題.jpg\n' +
      '例：1-照片-柴犬麻糬.jpg、1-模具.jpg、1-餅乾.jpg\n\n' +
      '・編號：同一組用同一個數字，數字小的排前面\n' +
      '・種類：照片／模具／餅乾（三張不必到齊）\n' +
      '・標題：可不寫；寫在任一張上即可\n' +
      '・要下架：把檔案移出資料夾，或檔名最前面加 x\n\n' +
      '注意：這個資料夾裡的圖片任何人都看得到，請只放可以公開的照片（要先取得客人同意）。\n', 'text/plain');
  }
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  CacheService.getScriptCache().remove('list');
  Logger.log('資料夾：' + folder.getUrl());
  return folder.getUrl();
}

/** 丟完圖片想馬上更新（不等 5 分鐘）時執行。 */
function refreshNow() { CacheService.getScriptCache().remove('list'); return list_(); }
