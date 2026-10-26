/* 104 班表讀取（門市排班工作台每天比對 104 用；2026-10-26）
   用法：在「內建瀏覽器」已登入 104 的分頁（pro.104.com.tw 任一頁）用 javascript_tool 整段執行。
   只讀：班表清單 grid/multi、整份班表 getStep2Schedule、匯入範本 getScheduleTemplate（回 xlsx，不會跳下載）。不按任何按鈕、不改 104。
   回傳：'SYNC104_BEGIN' + JSON + 'SYNC104_END'（內容很大，工具會自動存成檔案，再用 tools/sync104/push.py 送到工作台）。 */
await (async () => {
  const STORE_KEYS = [[1, '精明'], [2, '草悟道'], [3, '南京'], [4, '士林'], [5, 'Focus'], [6, '新竹'], [7, '板橋'],
    [8, '新店店'], [9, '中壢'], [10, '藝文'], [11, 'A13'], [12, 'SKM']];
  const post = (u, b) => fetch(u, { method: 'POST', credentials: 'include', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(b) });
  const login = await fetch('/hrm2/api/auth/checkLogin', { credentials: 'include' });
  if (location.pathname.indexOf('signin') >= 0 || login.status !== 200) return 'SYNC104_BEGIN' + JSON.stringify({ ok: false, error: '104 沒有登入' }) + 'SYNC104_END';
  const list = (await (await post('/hrm2/api/shift/grid/multi', { status: -1, $criteriaType: 'basic' })).json()).data.data;
  const now = new Date(Date.now() + 8 * 3600e3);
  const yms = [0, 1].map(k => { const d = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth() + k, 1)); return d.getUTCFullYear() * 100 + d.getUTCMonth() + 1; });
  const pick = {};
  for (const s of list) {
    const m = /^(\d{3})(\d{2})\s/.exec(s.name || '');
    if (!m) continue;
    const ym = (parseInt(m[1], 10) + 1911) * 100 + parseInt(m[2], 10);
    if (yms.indexOf(ym) < 0) continue;
    const st = STORE_KEYS.find(([, k]) => s.name.indexOf(k) >= 0);
    if (!st) continue;   // 區主管、出貨中心等不是門市的班表略過
    const key = st[0] + '-' + ym;
    if (!pick[key] || s.modifyDate > pick[key].modifyDate) pick[key] = Object.assign({ store: st[0], ym: String(ym) }, s);
  }
  const items = [];
  for (const key of Object.keys(pick)) {
    const s = pick[key];
    try {
      const d = (await (await fetch('/hrm2/api/shift/getStep2Schedule/multi/' + s.id, { credentials: 'include' })).json()).data;
      const emps = d.scheduledEmps.map(e => ({ empId: e.empId, name: e.name, mustHour: e.mustHour, annualHour: e.annualHour, compHour: e.compHour }));
      const tags = {}; d.tagList.forEach(t => { tags[t.id] = t.name; });
      const pos = {}; d.positionList.forEach(p => { pos[p.id] = p.name; });
      const cls = {}; d.allClassList.forEach(c => { cls[c.id] = c.name; });
      const shifts = d.empShifts.filter(x => x.scheduleId === s.id && !x.destroyed)
        .map(x => [x.empId, x.baseDate, x.type, x.fromTime, x.toTime, x.tagIds || [], x.positionId || 0, x.classId || 0]);
      let tpl = '';
      if (emps.length) {
        const r = await post('/hrm2/api/shift/multiShift/getScheduleTemplate/', { id: s.id, type: 2, empIds: emps.map(e => e.empId) });
        if (r.ok) {
          const u = new Uint8Array(await r.arrayBuffer());
          let bin = ''; for (let i = 0; i < u.length; i += 8192) bin += String.fromCharCode.apply(null, u.subarray(i, i + 8192));
          tpl = btoa(bin);
        }
      }
      items.push({ store: s.store, ym: s.ym, id: String(s.id), name: s.name, status: s.status, modify: s.modifyDate,
        start: s.duration[0], end: s.duration[1], emps, tags, pos, cls, shifts, template_b64: tpl });
    } catch (e) {
      items.push({ store: s.store, ym: s.ym, id: String(s.id), name: s.name, error: String(e) });
    }
  }
  return 'SYNC104_BEGIN' + JSON.stringify({ ok: true, at: Date.now(), items }) + 'SYNC104_END';
})()
