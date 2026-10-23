/* 🌏 平板外語食譜：食譜畫面（2026-10-24）
   tablet-recipe.html（店內平板）與 tablet-recipe-review.html（總部對照核准的「用平板畫面看」）共用。
   版型照門市食譜後台的平板頁（/steps/recipedetails/{id}）：左右滑動一頁一頁——
   開場說明 → 封面 → 注意事項（要勾「已讀」才能往下）→ 拿托盤 → 步驟（黃色標題列、影片只在翻到那頁才載入＋重新載入鈕、
   大圖、內文、步驟清單）／食材頁（容器說明圖、依取放區分組＋「看位置」照片、塑膠碗／瓷碗圖示）→ 保存方式 → 完成。
   資料＝申請單 API i18nGet／i18nPreview 回的內容（i18n.gs i18nBuild_）。圖片影片直接用後台雲端同一個網址。
   用法：TRView.mount(容器, 資料, 語言, { onExit: fn, preview: true/false }) → 回傳 { destroy } */
(function () {
  'use strict';
  var BACK = 'https://diybc.azurewebsites.net';
  var AREA_ORDER = ['5e1bd8a2-04b1-457a-8638-fef023c223f3', 'c1b9eaae-f1f0-4152-8392-d45507cbafa1', '5c1f4ecf-2b35-45aa-bc7d-a2cd7710b405',
    'e796c6da-1b35-4a0a-8f9d-78e7619ad5ed', 'f7b9b79a-448f-44d1-8154-979dcaf68fed', '971ea9b5-095b-4fe8-a29e-fce186544e21', '796e6d72-bf7d-40eb-bbaf-fd678b64b871'];   /* 跟後台平板頁同順序 */
  var T = {
    ja: { intro: '【はじめての方へ】', introBody: 'このタブレットが今日の先生です。\n画面の説明に沿って、1ステップずつ進めてください。\nわからないことがあれば、いつでもスタッフにお声がけください。', ready: '準備はいいですか？\n画面を左にスワイプして始めましょう ►►►',
      noticeTitle: 'ご注意ください', notice: ['このレシピは細かく設計されています。手順を変えると仕上がりに影響します。', 'レシピにない材料を使いたい場合は、スタッフにお声がけください（追加の材料はご購入となります）。', '使った道具や容器は洗ってください。', 'このレシピは著作権で保護されています。写真・動画の撮影、メモ、その他の方法での複製はご遠慮ください。'],
      agree: '上記の内容を確認しました。', agreeFirst: '先に「確認しました」にチェックしてください', tray: 'トレーを1枚取って、次のページから始めましょう。',
      bowlP: 'プラスチックボウル', bowlC: '陶器のボウル', oneBowl: '材料1つにつき容器1つ', where: '場所を見る', menu: 'メニュー', next: '次へ', prev: '前へ', steps: 'ステップ一覧',
      done: '完成です！おつかれさまでした。', doneSub: 'わからないことがあれば、スタッフにお声がけください。', backMenu: 'メニューに戻る', reload: '動画を再読み込み', pres: '保存方法', close: '閉じる' },
    ko: { intro: '[처음 오신 분은 꼭 읽어 주세요]', introBody: '오늘은 이 태블릿이 선생님입니다.\n화면 안내에 따라 한 단계씩 진행해 주세요.\n궁금한 점은 언제든지 직원에게 물어보세요.', ready: '준비되셨나요?\n화면을 왼쪽으로 넘겨 시작하세요 ►►►',
      noticeTitle: '주의 사항', notice: ['이 레시피는 세심하게 설계되어 있습니다. 임의로 바꾸면 결과에 영향을 줍니다.', '레시피에 없는 재료를 사용하고 싶으시면 직원에게 문의해 주세요(추가 재료는 구매하셔야 합니다).', '사용한 도구와 용기는 세척해 주세요.', '이 레시피는 저작권으로 보호됩니다. 사진·동영상 촬영, 메모, 기타 방법으로 복제하는 것을 금지합니다.'],
      agree: '위 내용을 확인했습니다.', agreeFirst: "먼저 '확인했습니다'에 체크해 주세요", tray: '트레이를 하나 가져오세요. 다음 페이지부터 시작합니다.',
      bowlP: '플라스틱 볼', bowlC: '도자기 볼', oneBowl: '재료 하나당 용기 하나', where: '위치 보기', menu: '메뉴', next: '다음', prev: '이전', steps: '단계 목록',
      done: '완성! 수고하셨습니다.', doneSub: '궁금한 점은 직원에게 물어보세요.', backMenu: '메뉴로 돌아가기', reload: '동영상 다시 불러오기', pres: '보관 방법', close: '닫기' },
    vi: { intro: '[Dành cho khách lần đầu]', introBody: 'Hôm nay chiếc máy tính bảng này là người hướng dẫn của bạn.\nHãy làm theo từng bước trên màn hình.\nNếu có thắc mắc, hãy hỏi nhân viên bất cứ lúc nào.', ready: 'Bạn đã sẵn sàng chưa?\nVuốt màn hình sang trái để bắt đầu ►►►',
      noticeTitle: 'Lưu ý', notice: ['Công thức này được thiết kế kỹ lưỡng; mọi thay đổi đều ảnh hưởng đến thành phẩm.', 'Nếu muốn dùng nguyên liệu không có trong công thức, vui lòng hỏi nhân viên (nguyên liệu thêm phải mua thêm).', 'Vui lòng rửa sạch dụng cụ và đồ đựng đã dùng.', 'Nội dung công thức được bảo hộ bản quyền; không chụp ảnh, quay phim, ghi chép hoặc sao chép dưới bất kỳ hình thức nào.'],
      agree: 'Tôi đã đọc kỹ các nội dung trên.', agreeFirst: 'Vui lòng đánh dấu "Tôi đã đọc" trước', tray: 'Hãy lấy một khay, bắt đầu từ trang tiếp theo.',
      bowlP: 'Tô nhựa', bowlC: 'Tô sứ', oneBowl: 'Mỗi nguyên liệu một đồ đựng', where: 'Xem vị trí', menu: 'Menu', next: 'Tiếp', prev: 'Trước', steps: 'Danh sách bước',
      done: 'Hoàn thành! Cảm ơn bạn.', doneSub: 'Nếu có thắc mắc, hãy hỏi nhân viên.', backMenu: 'Về menu', reload: 'Tải lại video', pres: 'Cách bảo quản', close: 'Đóng' },
    id: { intro: '[Wajib dibaca untuk pengunjung pertama]', introBody: 'Tablet ini adalah instruktur Anda hari ini.\nIkuti petunjuk di layar langkah demi langkah.\nJika ada pertanyaan, silakan tanya staf kami kapan saja.', ready: 'Sudah siap?\nGeser layar ke kiri untuk mulai ►►►',
      noticeTitle: 'Perhatian', notice: ['Resep ini dirancang dengan cermat; setiap perubahan akan memengaruhi hasilnya.', 'Jika ingin memakai bahan yang tidak ada di resep, silakan minta bantuan staf (bahan tambahan harus dibeli).', 'Harap cuci alat dan wadah yang sudah dipakai.', 'Isi resep ini dilindungi hak cipta; dilarang memotret, merekam video, mencatat, atau menyalin dengan cara apa pun.'],
      agree: 'Saya sudah membaca semua hal di atas.', agreeFirst: 'Centang "Saya sudah membaca" terlebih dahulu', tray: 'Ambil satu nampan, lalu mulai dari halaman berikutnya.',
      bowlP: 'Mangkuk plastik', bowlC: 'Mangkuk keramik', oneBowl: 'Satu wadah untuk setiap bahan', where: 'Lihat lokasi', menu: 'Menu', next: 'Lanjut', prev: 'Kembali', steps: 'Daftar langkah',
      done: 'Selesai! Terima kasih.', doneSub: 'Jika ada yang kurang jelas, tanyakan staf.', backMenu: 'Kembali ke menu', reload: 'Muat ulang video', pres: 'Cara penyimpanan', close: 'Tutup' },
    th: { intro: '[สำหรับผู้มาครั้งแรก]', introBody: 'วันนี้แท็บเล็ตเครื่องนี้คือผู้สอนของคุณ\nทำตามคำแนะนำบนหน้าจอทีละขั้นตอน\nหากมีข้อสงสัย สอบถามพนักงานได้ตลอดเวลา', ready: 'พร้อมหรือยัง?\nปัดหน้าจอไปทางซ้ายเพื่อเริ่ม ►►►',
      noticeTitle: 'ข้อควรทราบ', notice: ['สูตรนี้ออกแบบมาอย่างละเอียด การเปลี่ยนแปลงใด ๆ จะมีผลต่อผลลัพธ์', 'หากต้องการใช้วัตถุดิบที่ไม่มีในสูตร กรุณาติดต่อพนักงาน (วัตถุดิบเพิ่มเติมต้องซื้อเพิ่ม)', 'กรุณาล้างอุปกรณ์และภาชนะที่ใช้แล้ว', 'เนื้อหาสูตรนี้ได้รับการคุ้มครองลิขสิทธิ์ ห้ามถ่ายภาพ ถ่ายวิดีโอ จดบันทึก หรือคัดลอกด้วยวิธีใด ๆ'],
      agree: 'ฉันอ่านข้อความข้างต้นแล้ว', agreeFirst: 'กรุณาติ๊ก "ฉันอ่านแล้ว" ก่อน', tray: 'หยิบถาดหนึ่งใบ แล้วเริ่มจากหน้าถัดไป',
      bowlP: 'ชามพลาสติก', bowlC: 'ชามเซรามิก', oneBowl: 'วัตถุดิบหนึ่งอย่างต่อหนึ่งภาชนะ', where: 'ดูตำแหน่ง', menu: 'เมนู', next: 'ถัดไป', prev: 'ก่อนหน้า', steps: 'รายการขั้นตอน',
      done: 'เสร็จแล้ว! ขอบคุณ', doneSub: 'หากมีข้อสงสัย สอบถามพนักงานได้เลย', backMenu: 'กลับไปที่เมนู', reload: 'โหลดวิดีโออีกครั้ง', pres: 'วิธีเก็บรักษา', close: 'ปิด' }
  };
  var HTML_LANG = { ja: 'ja', ko: 'ko', vi: 'vi', id: 'id', th: 'th' };

  var CSS = '.trv{position:fixed;inset:0;z-index:9000;background:#fff;color:#40210f;font-family:system-ui,-apple-system,"Hiragino Sans","Noto Sans",sans-serif;--u:min(1vw,0.78vh);overflow:hidden;-webkit-user-select:none;user-select:none;-webkit-touch-callout:none}'
    + '.trv *{box-sizing:border-box}'
    + '.trv-track{position:absolute;left:0;top:0;bottom:calc(var(--u)*9.6);display:flex;transition:transform .32s ease;will-change:transform}'
    + '.trv-slide{flex:0 0 100vw;width:100vw;height:100%;overflow-y:auto;overflow-x:hidden;-webkit-overflow-scrolling:touch;position:relative;background:#fff}'
    + '.trv-bar{position:absolute;left:0;right:0;bottom:0;height:calc(var(--u)*9.6);display:flex;background:#efefef;border-top:2px solid #fff;z-index:5}'
    + '.trv-arrow{flex:0 0 calc(var(--u)*10);display:flex;align-items:center;justify-content:center;font-size:calc(var(--u)*4.5);font-weight:800;color:#40210f;background:#efefef;border:none;cursor:pointer}'
    + '.trv-arrow[disabled]{color:#ccc;cursor:default}'
    + '.trv-nums{flex:1;display:flex;overflow-x:auto;scrollbar-width:none;align-items:center}.trv-nums::-webkit-scrollbar{display:none}'
    + '.trv-num{flex:0 0 calc(var(--u)*10);text-align:center;font-size:calc(var(--u)*3.2);padding:calc(var(--u)*2.4) 0;color:#8a7a70;cursor:pointer;border:none;background:none}'
    + '.trv-num.on{background:#fdd35d;color:#40210f;font-weight:800}'
    + '.trv-menu{flex:0 0 auto;border:none;border-right:2px solid #fff;background:#40210f;color:#fff;font-size:calc(var(--u)*2.8);font-weight:800;padding:0 calc(var(--u)*3);cursor:pointer}'
    + '.trv-pv{position:absolute;right:calc(var(--u)*2);top:calc(var(--u)*2);z-index:6;background:#c43030;color:#fff;border-radius:999px;padding:calc(var(--u)*1) calc(var(--u)*2.6);font-size:calc(var(--u)*2.4);font-weight:800}'
    + '.trv-intro{min-height:100%;padding:calc(var(--u)*10);font-size:calc(var(--u)*3.7);white-space:pre-wrap;line-height:1.6;background:url(' + BACK + '/img/bottom.jpg) no-repeat right bottom/auto 35vh}'
    + '.trv-intro h2{font-size:calc(var(--u)*7);margin:calc(var(--u)*12) 0 calc(var(--u)*4)}'
    + '.trv-cover{height:calc(100% - var(--u)*25);background:#f6f1ec center bottom/cover no-repeat}'
    + '.trv-cname{height:calc(var(--u)*25);display:flex;align-items:center;justify-content:center;text-align:center;font-size:calc(var(--u)*5);font-weight:700;padding:0 calc(var(--u)*5)}'
    + '.trv-nhead{background:#efdfcf;height:calc(var(--u)*18);display:flex;align-items:center;justify-content:center;font-size:calc(var(--u)*4.5);font-weight:800}'
    + '.trv-nbody{padding:calc(var(--u)*8) calc(var(--u)*10) calc(var(--u)*30);font-size:calc(var(--u)*3.7);line-height:1.7}.trv-nbody ol{padding-left:1.2em;margin:0}.trv-nbody li{margin-bottom:.8em}'
    + '.trv-nbody li.r{color:#ff0000;font-weight:800}'
    + '.trv-agree{position:absolute;left:0;right:0;bottom:0;min-height:calc(var(--u)*25);background:#fdd35d;display:flex;align-items:center;justify-content:center;gap:calc(var(--u)*3);font-size:calc(var(--u)*4.2);font-weight:800;padding:0 calc(var(--u)*6);cursor:pointer}'
    + '.trv-agree input{width:calc(var(--u)*7);height:calc(var(--u)*7);accent-color:#40210f;flex:0 0 auto}'
    + '.trv-tray{height:34%;display:flex;align-items:flex-end;justify-content:center;text-align:center;font-size:calc(var(--u)*7);font-weight:800;padding:0 10%}'
    + '.trv-trayimg{height:66%;background:url(' + BACK + '/img/tray.jpg) no-repeat center top/100% auto}'
    + '.trv-video{width:100%;background:#000;display:block;max-height:62vh}'
    + '.trv-vwrap{position:relative}.trv-reload{flex:0 0 auto;border:none;background:rgba(255,255,255,.55);border-radius:6px;font-size:calc(var(--u)*2.8);padding:.2em .6em;color:#402100;font-weight:800;cursor:pointer}'
    + '.trv-img{width:100%;display:block;cursor:zoom-in}'
    + '.trv-title{position:sticky;top:0;z-index:3;background:#fdd35d;color:#40210f;font-weight:800;font-size:calc(var(--u)*3.2);padding:.75em 1em;display:flex;align-items:center;gap:.6em}'
    + '.trv-title .n{flex:0 0 auto;opacity:.75}.trv-title .t{flex:1}'
    + '.trv-lbtn{flex:0 0 auto;border:none;background:rgba(255,255,255,.55);border-radius:6px;font-size:calc(var(--u)*2.8);padding:.2em .6em;cursor:pointer;color:#40210f}'
    + '.trv-body{font-size:calc(var(--u)*3.2);padding:1em 1em calc(var(--u)*12);line-height:1.65}.trv-body p{margin:0 0 .6em}.trv-body p.c{text-align:center}.trv-body img{max-width:100%;height:auto;cursor:zoom-in}'
    + '.trv-cont{background:url(' + BACK + '/img/container.png) no-repeat center top/100% auto;height:calc(var(--u)*40);display:flex;flex-wrap:wrap;font-weight:800;text-align:center}'
    + '.trv-cont .b{width:50%;font-size:calc(var(--u)*3.4);padding-top:calc(var(--u)*20)}'
    + '.trv-cont .o{width:100%;font-size:calc(var(--u)*5);background:#efdfcf;margin-top:calc(var(--u)*2.5);padding:.2em 0}'
    + '.trv-area{padding:calc(var(--u)*3) calc(var(--u)*4) 0}.trv-aname{font-weight:800;font-size:calc(var(--u)*3.6);border-bottom:2px solid #efdfcf;padding-bottom:.3em;display:flex;align-items:center;gap:.5em}'
    + '.trv-where{border:1px solid #d8c8be;background:#fff;border-radius:999px;font-size:calc(var(--u)*2.6);padding:.2em .8em;cursor:pointer;color:#40210f;font-weight:700}'
    + '.trv-ings{display:flex;flex-wrap:wrap;padding:.5em 0 0}.trv-ing{width:50%;display:flex;align-items:center;gap:.5em;font-size:calc(var(--u)*3.2);padding:.35em 0}'
    + '.trv-ing img{width:calc(var(--u)*8);flex:0 0 auto}.trv-ing .q{font-weight:800;white-space:nowrap}'
    + '.trv-pres{position:absolute;left:0;right:0;bottom:0;background:#fff;padding:.8em 1.2em;font-size:calc(var(--u)*3);white-space:pre-wrap;line-height:1.6;max-height:55%;overflow:auto}'
    + '.trv-done{height:100%;display:flex;flex-direction:column;align-items:center;justify-content:center;text-align:center;gap:calc(var(--u)*4);padding:10%}'
    + '.trv-done h2{font-size:calc(var(--u)*7);margin:0}.trv-done p{font-size:calc(var(--u)*3.6);margin:0}'
    + '.trv-big{background:#fdd35d;border:none;border-radius:14px;font-size:calc(var(--u)*4.5);font-weight:800;padding:.6em 1.6em;color:#40210f;cursor:pointer}'
    + '.trv-list{position:absolute;right:0;top:0;bottom:calc(var(--u)*9.6);width:min(70vw,560px);background:#efefef;z-index:8;overflow-y:auto;box-shadow:-4px 0 18px rgba(0,0,0,.18);font-size:calc(var(--u)*3.2)}'
    + '.trv-list h3{margin:0;background:#fdd35d;padding:.7em 1em;font-size:1em;position:sticky;top:0}'
    + '.trv-list button{display:flex;gap:.6em;width:100%;text-align:left;border:none;background:none;padding:.65em 1em;font-size:1em;color:#40210f;cursor:pointer;border-bottom:1px solid #e2d8d0}'
    + '.trv-list button.on{background:#fff3c4;font-weight:800}.trv-list .n{flex:0 0 2.2em;opacity:.7}'
    + '.trv-lb{position:absolute;inset:0;z-index:20;background:rgba(0,0,0,.86);display:flex;align-items:center;justify-content:center;flex-direction:column;gap:12px}'
    + '.trv-lb img{max-width:96vw;max-height:84vh;object-fit:contain}.trv-lb button{font-size:18px;padding:8px 22px;border-radius:999px;border:none;font-weight:800}'
    + '.trv-toast{position:absolute;left:50%;top:40%;transform:translateX(-50%);background:#40210f;color:#fff;padding:.7em 1.4em;border-radius:999px;font-size:calc(var(--u)*3.2);z-index:30;font-weight:700}';

  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function abs(u) { u = String(u || ''); return /^https?:/i.test(u) ? u : u ? BACK + (u.charAt(0) === '/' ? '' : '/') + u : ''; }
  /* 譯文內文：只放行 p／br／b／u／i／font color／img（後台雲端圖片）；其他標籤當文字 */
  function safeHtml(h) {
    return String(h || '').replace(/<(\/?)([a-z0-9]+)([^>]*)>/gi, function (m, cl, tag, at) {
      tag = tag.toLowerCase();
      if (tag === 'br') return '<br>';
      if (tag === 'b' || tag === 'u' || tag === 'i') return '<' + cl + tag + '>';
      if (tag === 'p') { if (cl) return '</p>'; return /class="c"/.test(at) ? '<p class="c">' : '<p>'; }
      if (tag === 'font') { if (cl) return '</font>'; var c = (at.match(/color="(#[0-9a-f]{3,6}|[a-z]+)"/i) || [])[1]; return c ? '<font color="' + c + '">' : '<font>'; }
      if (tag === 'img' && !cl) { var s = (at.match(/src="([^"]+)"/i) || [])[1] || ''; return /^https:\/\/diybcstorage\.blob\.core\.windows\.net\//.test(s) ? '<img src="' + esc(s) + '" alt="">' : ''; }
      return '';
    });
  }
  function qtyTxt(q, u) { q = String(q == null ? '' : q); u = String(u || ''); return /\d/.test(u) ? u + ' × ' + q : (q + (u ? ' ' + u : '')); }
  function injectCss() { if (document.getElementById('trvCss')) return; var st = document.createElement('style'); st.id = 'trvCss'; st.textContent = CSS; document.head.appendChild(st); }

  function mount(host, data, lang, opt) {
    opt = opt || {}; injectCss();
    var L = T[lang] || T.ja, root = document.createElement('div'), slides = [], cur = 0, agreed = false, destroyed = false;
    root.className = 'trv'; root.lang = HTML_LANG[lang] || 'en';
    /* 版面 */
    var intro = '<div class="trv-intro"><h2>' + esc(L.intro) + '</h2>' + esc(L.introBody) + '\n\n' + esc(L.ready) + '</div>';
    slides.push({ kind: 'fix', html: intro });
    slides.push({ kind: 'fix', html: '<div class="trv-cover" style="background-image:url(\'' + esc(abs(data.cover)) + '\')"></div><div class="trv-cname">' + esc(data.name) + '</div>' });
    slides.push({ kind: 'notice', html: '<div class="trv-nhead">' + esc(L.noticeTitle) + '</div><div class="trv-nbody"><ol>' + L.notice.map(function (t, i) { return '<li' + (i === 1 ? ' class="r"' : '') + '>' + esc(t) + '</li>'; }).join('') + '</ol></div>'
      + '<label class="trv-agree"><input type="checkbox" class="trv-agcb">' + esc(L.agree) + '</label>' });
    slides.push({ kind: 'fix', html: '<div class="trv-tray">' + esc(L.tray) + '</div><div class="trv-trayimg"></div>' });
    var N = data.slides.length, first = slides.length;
    data.slides.forEach(function (s, i) {
      var vid = !(s.ings && s.ings.length) && !!s.video;   /* 影片頁：重新載入鈕放標題列右邊（跟後台一樣；第 3 頁說明圖畫的就是那個位置） */
      var head = '<div class="trv-title"><span class="n">' + (i + 1) + ' / ' + N + '</span><span class="t">' + esc(s.t) + '</span>'
        + (vid ? '<button type="button" class="trv-reload" data-reload="1" title="' + esc(L.reload) + '">↻</button>' : '')
        + '<button type="button" class="trv-lbtn" data-list="1">☰ ' + esc(L.steps) + '</button></div>';
      var h = '';
      if (s.ings && s.ings.length) {
        h += '<div class="trv-cont"><div class="b">[ ' + esc(L.bowlP) + ' ]</div><div class="b">[ ' + esc(L.bowlC) + ' ]</div><div class="o">' + esc(L.oneBowl) + '</div></div>' + head;
        var groups = {}, order = [];
        s.ings.forEach(function (g) { var a = g.a || '_'; if (!groups[a]) { groups[a] = []; order.push(a); } groups[a].push(g); });
        order.sort(function (a, b) { var ia = AREA_ORDER.indexOf(a), ib = AREA_ORDER.indexOf(b); return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib); });
        order.forEach(function (a) {
          var ar = (data.areas || {})[a] || {}, list = groups[a].slice().sort(function (x, y) { var r = function (c) { return c === 'container1' ? 1 : c === 'container2' ? 2 : 0; }; return r(x.c) - r(y.c); });
          h += '<div class="trv-area"><div class="trv-aname">' + esc(ar.name || '') + (ar.img ? ' <button type="button" class="trv-where" data-zoom="' + esc(abs(ar.img)) + '">📍 ' + esc(L.where) + '</button>' : '') + '</div><div class="trv-ings">'
            + list.map(function (g) { var ic = g.c === 'container1' ? BACK + '/img/pbowl.png' : g.c === 'container2' ? BACK + '/img/cbowl.png' : ''; return '<div class="trv-ing">' + (ic ? '<img src="' + ic + '" alt="">' : '') + '<span>' + esc(g.n) + ' <span class="q">' + esc(qtyTxt(g.q, g.u)) + '</span></span></div>'; }).join('') + '</div></div>';
        });
        h += '<div style="height:calc(var(--u)*14)"></div>';
      } else {
        if (s.video) h += '<div class="trv-vwrap"><video class="trv-video" controls playsinline preload="none" data-src="' + esc(s.video) + '"></video></div>';
        else if (s.img) h += '<img class="trv-img" src="' + esc(s.img) + '" alt="" loading="lazy" data-zoom="' + esc(s.img) + '">';
        h += head + '<div class="trv-body">' + safeHtml(s.h) + '</div>';
      }
      slides.push({ kind: 'step', html: h, n: i });
    });
    if (data.pres) slides.push({ kind: 'fix', html: '<div class="trv-cover" style="height:100%;background-position:center;background-image:url(\'' + esc(abs(data.cover)) + '\')"></div><div class="trv-pres">' + esc(data.pres) + '</div>' });
    slides.push({ kind: 'fix', html: '<div class="trv-done"><h2>🎉 ' + esc(L.done) + '</h2><p>' + esc(L.doneSub) + '</p><button type="button" class="trv-big" data-exit="1">' + esc(L.backMenu) + '</button></div>' });

    root.innerHTML = '<div class="trv-track">' + slides.map(function (s) { return '<section class="trv-slide">' + s.html + '</section>'; }).join('') + '</div>'
      + (opt.preview ? '<div class="trv-pv">PREVIEW 預覽（未放行）</div>' : '')
      + '<div class="trv-bar"><button type="button" class="trv-menu" data-exit="1">☰ ' + esc(L.menu) + '</button><button type="button" class="trv-arrow" data-go="-1" aria-label="' + esc(L.prev) + '">‹</button><div class="trv-nums">'
      + slides.map(function (s, i) { return s.kind === 'step' ? '<button type="button" class="trv-num" data-to="' + i + '">' + ('0' + (s.n + 1)).slice(-2) + '</button>' : ''; }).join('')
      + '</div><button type="button" class="trv-arrow" data-go="1" aria-label="' + esc(L.next) + '">›</button></div>';
    host.appendChild(root);
    var track = root.querySelector('.trv-track'), secs = root.querySelectorAll('.trv-slide'), nums = root.querySelector('.trv-nums');

    function toast(t) { var d = document.createElement('div'); d.className = 'trv-toast'; d.textContent = t; root.appendChild(d); setTimeout(function () { if (d.parentNode) d.parentNode.removeChild(d); }, 1800); }
    function blocked(to) { return to > 2 && !agreed; }   /* 注意事項（第 3 頁）沒勾不能往後 */
    function videos() {
      secs.forEach(function (sec, i) {
        var v = sec.querySelector('video'); if (!v) return;
        if (i === cur) { if (!v.getAttribute('src')) { v.setAttribute('src', v.getAttribute('data-src')); v.load(); } }
        else if (v.getAttribute('src')) { try { v.pause(); } catch (e) { } v.removeAttribute('src'); v.load(); }   /* 釋放解碼器（後台也是這樣做，避免黑畫面） */
      });
    }
    function go(to) {
      to = Math.max(0, Math.min(slides.length - 1, to));
      if (blocked(to)) { to = Math.min(to, 2); if (cur === 2) toast(L.agreeFirst); }
      cur = to;
      track.style.transform = 'translateX(' + (-100 * cur) + 'vw)';
      root.querySelectorAll('.trv-num').forEach(function (b) { var on = Number(b.getAttribute('data-to')) === cur; b.classList.toggle('on', on); if (on) b.scrollIntoView({ block: 'nearest', inline: 'center', behavior: 'smooth' }); });
      root.querySelector('[data-go="-1"]').disabled = cur === 0;
      root.querySelector('[data-go="1"]').disabled = cur === slides.length - 1;
      closeList(); videos();
      if (opt.onPage) opt.onPage(cur, slides.length);
    }
    var listEl = null;
    function closeList() { if (listEl && listEl.parentNode) listEl.parentNode.removeChild(listEl); listEl = null; }
    function openList() {
      if (listEl) { closeList(); return; }
      listEl = document.createElement('div'); listEl.className = 'trv-list';
      listEl.innerHTML = '<h3>' + esc(data.name) + '</h3>' + slides.map(function (s, i) { return s.kind === 'step' ? '<button type="button" data-to="' + i + '"' + (i === cur ? ' class="on"' : '') + '><span class="n">' + ('0' + (s.n + 1)).slice(-2) + '</span><span>' + esc(data.slides[s.n].t) + '</span></button>' : ''; }).join('');
      root.appendChild(listEl);
      var on = listEl.querySelector('.on'); if (on) on.scrollIntoView({ block: 'center' });
    }
    function zoom(src) {
      var d = document.createElement('div'); d.className = 'trv-lb';
      d.innerHTML = '<img src="' + esc(src) + '" alt=""><button type="button">' + esc(L.close) + '</button>';
      d.onclick = function () { d.parentNode.removeChild(d); };
      root.appendChild(d);
    }
    root.addEventListener('click', function (e) {
      var t = e.target.closest ? e.target.closest('[data-go],[data-to],[data-exit],[data-list],[data-reload],[data-zoom]') : null;
      if (!t || !root.contains(t)) return;
      if (t.hasAttribute('data-go')) go(cur + Number(t.getAttribute('data-go')));
      else if (t.hasAttribute('data-to')) go(Number(t.getAttribute('data-to')));
      else if (t.hasAttribute('data-exit')) { destroy(); if (opt.onExit) opt.onExit(); }
      else if (t.hasAttribute('data-list')) openList();
      else if (t.hasAttribute('data-reload')) { var v = t.closest('.trv-slide').querySelector('video'); v.setAttribute('src', v.getAttribute('data-src')); v.load(); try { v.play(); } catch (e2) { } }
      else if (t.hasAttribute('data-zoom')) zoom(t.getAttribute('data-zoom'));
    });
    root.querySelector('.trv-agcb').addEventListener('change', function () { agreed = this.checked; if (agreed) go(3); });
    /* 左右滑動（水平明顯大於垂直才算；在影片上也可以滑） */
    var sx = 0, sy = 0, st = 0;
    root.addEventListener('touchstart', function (e) { var p = e.touches[0]; sx = p.clientX; sy = p.clientY; st = Date.now(); }, { passive: true });
    root.addEventListener('touchend', function (e) {
      if (listEl || root.querySelector('.trv-lb')) return;
      var p = e.changedTouches[0], dx = p.clientX - sx, dy = p.clientY - sy;
      if (Math.abs(dx) > 60 && Math.abs(dx) > Math.abs(dy) * 1.5 && Date.now() - st < 900) go(cur + (dx < 0 ? 1 : -1));
    }, { passive: true });
    function onKey(e) { if (e.key === 'ArrowRight') go(cur + 1); else if (e.key === 'ArrowLeft') go(cur - 1); else if (e.key === 'Escape') { if (root.querySelector('.trv-lb')) root.querySelector('.trv-lb').click(); else closeList(); } }
    document.addEventListener('keydown', onKey);
    /* 禁止長按存圖、右鍵（注意事項第 4 條：不可拍攝或複製；擋不住截圖，只是不方便） */
    root.addEventListener('contextmenu', function (e) { e.preventDefault(); });
    function destroy() { if (destroyed) return; destroyed = true; document.removeEventListener('keydown', onKey); root.querySelectorAll('video').forEach(function (v) { try { v.pause(); } catch (e) { } v.removeAttribute('src'); }); if (root.parentNode) root.parentNode.removeChild(root); }
    go(0);
    return { destroy: destroy, go: go, count: slides.length, firstStep: first, root: root };
  }
  window.TRView = { mount: mount, T: T, BACK: BACK, abs: abs, safeHtml: safeHtml, qtyTxt: qtyTxt, esc: esc };
})();
