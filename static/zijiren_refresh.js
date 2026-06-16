/* 自己人一鍵刷新腳本 — 放在 static/zijiren_refresh.js
   由書籤載入後在「已登入的 Azure 後台分頁」執行：
   抓最新全部 → 覆蓋 188uu → 自動重建 agg_member_month + 4 張聚合表 → 儀表板更新
   過程約 20 分鐘，右上角會顯示進度；請勿關閉或重新整理該分頁。 */
(async () => {
  if (window.__diybcRefreshing) { alert('刷新已在進行中，請稍候'); return; }
  window.__diybcRefreshing = true;

  // 同一個 /exec（GET=查詢、POST=刷新）
  const GAS = 'https://script.google.com/macros/s/AKfycbx61-Ww1W3CCAA1luawfKaO3lsc7kVrDw5sB0_z21HWtAbJ9BMy_Mm8T4fnkqU91QsP/exec';

  // ── 進度框 ──
  const box = document.createElement('div');
  box.style.cssText = 'position:fixed;top:16px;right:16px;z-index:2147483647;background:#fff;border:2px solid #C4412F;border-radius:10px;padding:14px 18px;font:14px/1.6 system-ui,sans-serif;color:#333;box-shadow:0 6px 20px rgba(0,0,0,.25);max-width:340px';
  box.innerHTML = '<b style="color:#C4412F">🔄 自己人資料刷新</b><div id="__rfMsg" style="margin-top:6px">準備中…</div>';
  document.body.appendChild(box);
  const msg = (t) => { const m = document.getElementById('__rfMsg'); if (m) m.innerHTML = t; };

  const sleep = (ms) => new Promise(r => setTimeout(r, ms));
  const post = async (o) => {
    let last;
    for (let k = 0; k < 5; k++) {
      try { const j = JSON.parse(await (await fetch(GAS, { method: 'POST', body: JSON.stringify(o) })).text()); if (j.ok !== false) return j; last = j; }
      catch (e) { last = { err: String(e) }; }
      await sleep(1500 * Math.pow(2, k));
    }
    throw new Error('GAS 失敗：' + JSON.stringify(last));
  };
  const getR = async (u) => {
    let last;
    for (let k = 0; k < 5; k++) {
      try { return await (await fetch(u, { credentials: 'include' })).text(); }
      catch (e) { last = String(e); await sleep(1500 * Math.pow(2, k)); }
    }
    throw new Error('抓取失敗：' + last);
  };

  // 視窗：2024-05 → 本月（半月一段，連續不重疊）
  const pad = (n) => String(n).padStart(2, '0');
  const now = new Date(), ey = now.getFullYear(), em = now.getMonth() + 1;
  const W = [];
  for (let y = 2024; y <= ey; y++) {
    const m0 = (y === 2024) ? 5 : 1, m1 = (y === ey) ? em : 12;
    for (let m = m0; m <= m1; m++) {
      const last = new Date(y, m, 0).getDate();
      W.push([`${y}-${pad(m)}-01`, `${y}-${pad(m)}-15`]);
      W.push([`${y}-${pad(m)}-16`, `${y}-${pad(m)}-${last}`]);
    }
  }

  try {
    msg('① 清空舊資料…');
    await post({ action: 'refreshInit' });

    let total = 0;
    for (let i = 0; i < W.length; i++) {
      const [b, e] = W[i];
      const html = await getR(`/VIPHis/storeindex?UserId=&HisStoreId=all&StoreId=all&KeyWord=&Begin=${b}&End=${e}`);
      const doc = new DOMParser().parseFromString(html, 'text/html');
      const trs = [...doc.querySelectorAll('table tbody tr')];
      const rows = trs.map(tr => {
        const c = tr.children;
        return [
          (c[0]?.textContent || '').trim(), (c[1]?.textContent || '').trim(), (c[3]?.textContent || '').trim(),
          (c[4]?.textContent || '').trim(), (c[5]?.textContent || '').trim(), (c[6]?.textContent || '').trim(),
          (c[2]?.getAttribute('data-v') || '').trim()
        ];
      });
      for (let j = 0; j < rows.length; j += 8000) {
        const r = await post({ action: 'append', rows: rows.slice(j, j + 8000) });
        total = r.total;
      }
      msg(`② 抓取中　${i + 1}/${W.length}<br>已寫入 ${total.toLocaleString()} 筆`);
      await sleep(300);
    }

    msg(`③ 重建會員月表（約 1.5 分鐘）…<br>共 ${total.toLocaleString()} 筆`);
    await post({ action: 'memberMonth' });

    msg('④ 重算儀表板聚合（約 1 分鐘）…');
    await post({ action: 'runAgg' });

    msg(`✅ 完成！共 <b>${total.toLocaleString()}</b> 筆<br>儀表板數字已更新（重新整理 zijiren 頁查看）`);
    box.style.borderColor = '#2e9e5b';
  } catch (err) {
    msg('❌ 失敗：' + err + '<br>可再點一次重試；若一直失敗回報 Claude');
    box.style.borderColor = '#c0392b';
  } finally {
    window.__diybcRefreshing = false;
  }
})();
