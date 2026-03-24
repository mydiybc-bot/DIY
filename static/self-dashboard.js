const refreshButton = document.getElementById("refresh-button");
const applyButton = document.getElementById("apply-button");
const refreshStatus = document.getElementById("refresh-status");
const sourcePath = document.getElementById("source-path");
const startDateInput = document.getElementById("start-date");
const endDateInput = document.getElementById("end-date");
const storeFilter = document.getElementById("store-filter");
const couponFilter = document.getElementById("coupon-filter");
const filterNote = document.getElementById("filter-note");
const metricGrid = document.getElementById("metric-grid");
const trendChart = document.getElementById("trend-chart");
const joinChart = document.getElementById("join-chart");
const couponChart = document.getElementById("coupon-chart");
const overviewBody = document.getElementById("overview-body");
const pairMetrics = document.getElementById("pair-metrics");
const pairList = document.getElementById("pair-list");
const matrixWrap = document.getElementById("matrix-wrap");
const dessertGrid = document.getElementById("dessert-grid");
const couponList = document.getElementById("coupon-list");
const couponStoreChart = document.getElementById("coupon-store-chart");
const trendCaption = document.getElementById("trend-caption");
const joinCaption = document.getElementById("join-caption");
const couponCaption = document.getElementById("coupon-caption");
const overviewCaption = document.getElementById("overview-caption");
const pairCaption = document.getElementById("pair-caption");
const dessertCaption = document.getElementById("dessert-caption");
const couponStoreCaption = document.getElementById("coupon-store-caption");

const numberFormatter = new Intl.NumberFormat("zh-TW");
const percentFormatter = new Intl.NumberFormat("zh-TW", {
  style: "percent",
  minimumFractionDigits: 1,
  maximumFractionDigits: 1,
});
const dateFormatter = new Intl.DateTimeFormat("zh-TW", {
  year: "numeric",
  month: "2-digit",
  day: "2-digit",
  hour: "2-digit",
  minute: "2-digit",
});

const state = {
  payload: null,
};

async function fetchDashboardData() {
  refreshButton.disabled = true;
  applyButton.disabled = true;
  refreshStatus.textContent = "正在同步自己人儀表板...";

  try {
    const params = new URLSearchParams();
    if (startDateInput.value) params.set("start_date", startDateInput.value);
    if (endDateInput.value) params.set("end_date", endDateInput.value);
    if (storeFilter.value) params.set("store", storeFilter.value);
    if (couponFilter.value) params.set("coupon", couponFilter.value);

    const response = await fetch(`/api/self-dashboard?${params.toString()}`, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`讀取失敗 (${response.status})`);
    }

    state.payload = await response.json();
    syncControls();
    renderDashboard();

    const generatedAt = state.payload.generated_at ? dateFormatter.format(new Date(state.payload.generated_at)) : "剛剛";
    refreshStatus.textContent = `已載入網頁儀表板，更新時間 ${generatedAt}。`;
    sourcePath.textContent = state.payload.source_path || "";
  } catch (error) {
    const message = error.message || "讀取失敗";
    refreshStatus.textContent = message;
    sourcePath.textContent = "";
    renderError(message);
  } finally {
    refreshButton.disabled = false;
    applyButton.disabled = false;
  }
}

function syncControls() {
  const payload = state.payload;
  if (!payload) return;

  const { filters, meta } = payload;
  startDateInput.value = filters.start_date || meta.default_start || "";
  endDateInput.value = filters.end_date || meta.default_end || "";

  storeFilter.innerHTML = ["<option value=\"全部\">全部店別</option>"]
    .concat(meta.active_stores.map((store) => `<option value="${escapeHtml(store)}">${escapeHtml(store)}</option>`))
    .join("");
  storeFilter.value = filters.store || "全部";

  couponFilter.innerHTML = meta.available_coupons
    .map((coupon) => `<option value="${escapeHtml(coupon)}">${escapeHtml(coupon === "全部" ? "全部優惠券" : coupon)}</option>`)
    .join("");
  couponFilter.value = filters.coupon || "全部";

  filterNote.textContent = `主分析固定排除已於 2025 年 4 月關閉的 ${meta.closed_store}。月營收目前只更新到 ${String(meta.revenue_latest_month).slice(0, 4)}-${String(meta.revenue_latest_month).slice(4)}。`;
}

function renderDashboard() {
  const payload = state.payload;
  if (!payload) return;

  renderMetrics(payload.summary);
  renderGroupedChart(trendChart, payload.monthly_trends);
  renderBarList(joinChart, payload.join_ranking.slice(0, 12), "store", ["join_self", "upgrade_self"], ["加入", "升等"]);
  renderCouponBarChart(couponChart, payload.coupon_analysis.coupon_ranking);
  renderOverviewTable(payload.store_overview);
  renderPairSection(payload.pair_analysis);
  renderDessertGrid(payload.dessert_rankings);
  renderCouponSection(payload.coupon_analysis);

  const currentStore = payload.filters.store === "全部" ? "12 家店" : payload.filters.store;
  trendCaption.textContent = `${currentStore} 的加入、甜點與優惠使用月趨勢`;
  joinCaption.textContent = `目前區間內共比較 ${payload.join_ranking.length} 家現行店別`;
  couponCaption.textContent = payload.coupon_analysis.coupon_ranking.length
    ? `最多使用的是 ${payload.coupon_analysis.coupon_ranking[0].coupon}`
    : "目前區間沒有優惠券使用紀錄";
  overviewCaption.textContent = `日期 ${payload.filters.start_date} 到 ${payload.filters.end_date}`;
  pairCaption.textContent = payload.filters.store === "全部"
    ? "全部加入店別的消費流向"
    : `只看 ${payload.filters.store} 加入會員的消費流向`;
  dessertCaption.textContent = payload.filters.store === "全部"
    ? "每張卡片是 1 家店的前 5 名甜點"
    : `目前篩選店別是 ${payload.filters.store}，甜點卡片已同步聚焦`;
  couponStoreCaption.textContent = payload.filters.coupon === "全部"
    ? "目前顯示所有優惠券使用在各店的分布"
    : `目前顯示 ${payload.filters.coupon} 在各店的使用排行`;
}

function renderMetrics(summary) {
  const cards = [
    {
      label: "加入自己人",
      value: formatNumber(summary.join_self_total),
      note: `升等自己人 ${formatNumber(summary.upgrade_total)}`,
    },
    {
      label: "甜點數量",
      value: formatNumber(summary.dessert_total),
      note: `加入群友 ${formatNumber(summary.group_join_total)}`,
    },
    {
      label: "優惠使用",
      value: formatNumber(summary.coupon_total),
      note: `100 元甜點券 ${formatNumber(summary.coupon_100_total)}`,
    },
    {
      label: "預估自己人營業額",
      value: formatCurrency(summary.estimated_revenue),
      note: `以每份甜點 NT$550 推估`,
    },
    {
      label: "可對照月營收",
      value: formatCurrency(summary.comparable_revenue),
      note: summary.revenue_share == null ? "目前區間沒有對應月營收" : `營收占比 ${formatPercent(summary.revenue_share)}`,
    },
    {
      label: "同店消費甜點數",
      value: formatNumber(summary.same_store_desserts),
      note: `跨店消費 ${formatNumber(summary.cross_store_desserts)}`,
    },
    {
      label: "跨店消費率",
      value: summary.cross_store_rate == null ? "0%" : formatPercent(summary.cross_store_rate),
      note: "以加入店別對消費店別計算",
    },
    {
      label: "資料區間",
      value: `${state.payload.filters.start_date} → ${state.payload.filters.end_date}`,
      note: "主分析只保留現行 12 家店",
      emphasis: true,
    },
  ];

  metricGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="metric-card">
          <p class="metric-label">${escapeHtml(card.label)}</p>
          <h2 class="${card.emphasis ? "metric-emphasis" : ""}">${escapeHtml(card.value)}</h2>
          <p class="metric-note">${escapeHtml(card.note)}</p>
        </article>
      `
    )
    .join("");
}

function renderGroupedChart(container, rows) {
  if (!rows.length) {
    container.innerHTML = buildEmptyState("目前沒有月份資料。");
    return;
  }

  const maxValue = Math.max(
    ...rows.flatMap((row) => [row.join_self + row.upgrade_self, row.desserts, row.coupon_use]),
    1
  );
  container.innerHTML = `
    <div class="grouped-bars">
      ${rows
        .map((row) => {
          const joinTotal = row.join_self + row.upgrade_self;
          return `
            <div class="group-row">
              <div class="group-label">
                <span>${escapeHtml(row.month)}</span>
                <span>${formatNumber(joinTotal)} / ${formatNumber(row.desserts)} / ${formatNumber(row.coupon_use)}</span>
              </div>
              <div class="group-bars">
                <div class="group-bar">
                  <div class="segment-rail"><div class="segment-fill join" style="width:${toPercent(joinTotal, maxValue)}%"></div></div>
                </div>
                <div class="group-bar">
                  <div class="segment-rail"><div class="segment-fill dessert" style="width:${toPercent(row.desserts, maxValue)}%"></div></div>
                </div>
                <div class="group-bar">
                  <div class="segment-rail"><div class="segment-fill coupon" style="width:${toPercent(row.coupon_use, maxValue)}%"></div></div>
                </div>
              </div>
            </div>
          `;
        })
        .join("")}
    </div>
    <div class="legend">
      <span class="legend-item"><span class="legend-swatch join"></span>加入+升等</span>
      <span class="legend-item"><span class="legend-swatch dessert"></span>甜點數量</span>
      <span class="legend-item"><span class="legend-swatch coupon"></span>優惠使用</span>
    </div>
  `;
}

function renderBarList(container, rows, labelKey, valueKeys, legends) {
  if (!rows.length) {
    container.innerHTML = buildEmptyState("目前沒有排行資料。");
    return;
  }

  const maxValue = Math.max(...rows.map((row) => valueKeys.reduce((sum, key) => sum + Number(row[key] || 0), 0)), 1);
  container.innerHTML = `
    <div class="bar-stack">
      ${rows
        .map((row) => {
          const total = valueKeys.reduce((sum, key) => sum + Number(row[key] || 0), 0);
          const fills = valueKeys
            .map((key, index) => {
              const colorClass = index === 0 ? "join" : index === 1 ? "dessert" : "coupon";
              return `<div class="segment-fill ${colorClass}" style="width:${toPercent(row[key], maxValue)}%"></div>`;
            })
            .join("");

          return `
            <div class="bar-item">
              <div class="bar-label">
                <span>${escapeHtml(row[labelKey])}</span>
                <span>${formatNumber(total)}</span>
              </div>
              <div class="group-bars">
                ${valueKeys
                  .map((key, index) => `
                    <div class="group-bar">
                      <div class="segment-rail"><div class="segment-fill ${index === 0 ? "join" : "dessert"}" style="width:${toPercent(row[key], maxValue)}%"></div></div>
                    </div>
                  `)
                  .join("")}
              </div>
            </div>
          `;
        })
        .join("")}
    </div>
    <div class="legend">
      ${legends
        .map(
          (legend, index) => `
            <span class="legend-item"><span class="legend-swatch ${index === 0 ? "join" : "dessert"}"></span>${escapeHtml(legend)}</span>
          `
        )
        .join("")}
    </div>
  `;
}

function renderCouponBarChart(container, rows) {
  if (!rows.length) {
    container.innerHTML = buildEmptyState("目前區間沒有優惠券使用。");
    return;
  }
  const maxValue = Math.max(...rows.map((row) => row.count), 1);
  container.innerHTML = `
    <div class="bar-stack">
      ${rows
        .map(
          (row) => `
            <div class="bar-item">
              <div class="bar-label">
                <span>${escapeHtml(row.coupon)}</span>
                <span>${formatNumber(row.count)}</span>
              </div>
              <div class="bar-rail"><div class="bar-fill" style="width:${toPercent(row.count, maxValue)}%"></div></div>
            </div>
          `
        )
        .join("")}
    </div>
  `;
}

function renderOverviewTable(rows) {
  overviewBody.innerHTML = rows.length
    ? rows
        .map(
          (row) => `
            <tr>
              <td>${escapeHtml(row.store)}</td>
              <td>${formatNumber(row.join_self)}</td>
              <td>${formatNumber(row.upgrade_self)}</td>
              <td>${formatNumber(row.join_total)}</td>
              <td>${formatNumber(row.desserts)}</td>
              <td>${formatCurrency(row.estimated_revenue)}</td>
              <td>${formatCurrency(row.revenue)}</td>
              <td>${row.revenue_share == null ? '<span class="muted-cell">-</span>' : formatPercent(row.revenue_share)}</td>
              <td>${formatNumber(row.coupon_use)}</td>
            </tr>
          `
        )
        .join("")
    : `<tr><td colspan="9" class="muted-cell">目前沒有符合條件的店別資料。</td></tr>`;
}

function renderPairSection(pairAnalysis) {
  pairMetrics.innerHTML = [
    {
      label: "同店消費",
      value: formatNumber(pairAnalysis.same_store_desserts),
      note: "加入店別與消費店別相同",
    },
    {
      label: "跨店消費",
      value: formatNumber(pairAnalysis.cross_store_desserts),
      note: "加入店別與消費店別不同",
    },
    {
      label: "跨店率",
      value: pairAnalysis.cross_store_rate == null ? "0%" : formatPercent(pairAnalysis.cross_store_rate),
      note: "用甜點數量計算",
    },
  ]
    .map(
      (item) => `
        <div class="mini-card">
          <p class="mini-label">${escapeHtml(item.label)}</p>
          <p class="mini-value">${escapeHtml(item.value)}</p>
          <p class="metric-note">${escapeHtml(item.note)}</p>
        </div>
      `
    )
    .join("");

  pairList.innerHTML = pairAnalysis.top_pairs.length
    ? pairAnalysis.top_pairs
        .map(
          (item, index) => `
            <div class="list-card">
              <div class="list-title">
                <span><span class="rank-badge">${index + 1}</span>${escapeHtml(item.join_store)} → ${escapeHtml(item.consume_store)}</span>
                <span>${formatNumber(item.desserts)}</span>
              </div>
              <p class="list-note">這條流向在目前區間累積了 ${formatNumber(item.desserts)} 份甜點。</p>
            </div>
          `
        )
        .join("")
    : buildEmptyState("目前沒有店別流向資料。");

  renderMatrix(pairAnalysis.matrix);
}

function renderMatrix(rows) {
  if (!rows.length) {
    matrixWrap.innerHTML = buildEmptyState("目前沒有矩陣資料。");
    return;
  }
  const maxValue = Math.max(...rows.flatMap((row) => row.totals.map((item) => item.desserts)), 1);
  const header = rows[0].totals.map((item) => `<th>${escapeHtml(item.consume_store)}</th>`).join("");
  const body = rows
    .map((row) => {
      const cells = row.totals
        .map((item) => {
          const alpha = item.desserts ? (0.15 + item.desserts / maxValue * 0.55).toFixed(3) : "0";
          return `<td class="heat" style="background: rgba(164,67,29,${alpha});">${item.desserts ? formatNumber(item.desserts) : ""}</td>`;
        })
        .join("");
      return `<tr><th>${escapeHtml(row.join_store)}</th>${cells}</tr>`;
    })
    .join("");
  matrixWrap.innerHTML = `<table class="matrix-table"><thead><tr><th>加入店別</th>${header}</tr></thead><tbody>${body}</tbody></table>`;
}

function renderDessertGrid(rankings) {
  dessertGrid.innerHTML = rankings.length
    ? rankings
        .map((ranking) => {
          const items = ranking.items.length
            ? ranking.items
                .map(
                  (item, index) => `
                    <div class="dessert-row">
                      <span class="rank-badge">${index + 1}</span>
                      <span class="dessert-name">${escapeHtml(item.dessert)}</span>
                      <span class="dessert-count">${formatNumber(item.count)}</span>
                    </div>
                  `
                )
                .join("")
            : `<div class="empty-state">目前區間沒有甜點紀錄。</div>`;
          return `
            <article class="dessert-card">
              <div class="dessert-store">
                <span>${escapeHtml(ranking.store)}</span>
                <span class="metric-note">Top 5</span>
              </div>
              <div class="dessert-list">${items}</div>
            </article>
          `;
        })
        .join("")
    : buildEmptyState("目前沒有甜點排行資料。");
}

function renderCouponSection(couponAnalysis) {
  couponList.innerHTML = couponAnalysis.coupon_ranking.length
    ? couponAnalysis.coupon_ranking
        .map(
          (item, index) => `
            <div class="list-card">
              <div class="list-title">
                <span><span class="rank-badge">${index + 1}</span>${escapeHtml(item.coupon)}</span>
                <span>${formatNumber(item.count)}</span>
              </div>
            </div>
          `
        )
        .join("")
    : buildEmptyState("目前區間沒有優惠券使用紀錄。");

  const rows = couponAnalysis.store_ranking.filter((item) => item.count > 0);
  if (!rows.length) {
    couponStoreChart.innerHTML = buildEmptyState("目前沒有店別優惠分布。");
    return;
  }
  const maxValue = Math.max(...rows.map((row) => row.count), 1);
  couponStoreChart.innerHTML = `
    <div class="bar-stack">
      ${rows
        .map(
          (row) => `
            <div class="bar-item">
              <div class="bar-label">
                <span>${escapeHtml(row.store)}</span>
                <span>${formatNumber(row.count)}</span>
              </div>
              <div class="bar-rail"><div class="bar-fill" style="width:${toPercent(row.count, maxValue)}%"></div></div>
            </div>
          `
        )
        .join("")}
    </div>
  `;
}

function renderError(message) {
  const empty = buildEmptyState(message);
  metricGrid.innerHTML = "";
  trendChart.innerHTML = empty;
  joinChart.innerHTML = empty;
  couponChart.innerHTML = empty;
  overviewBody.innerHTML = `<tr><td colspan="9" class="muted-cell">${escapeHtml(message)}</td></tr>`;
  pairMetrics.innerHTML = "";
  pairList.innerHTML = empty;
  matrixWrap.innerHTML = empty;
  dessertGrid.innerHTML = empty;
  couponList.innerHTML = empty;
  couponStoreChart.innerHTML = empty;
}

function formatNumber(value) {
  return numberFormatter.format(Number(value || 0));
}

function formatCurrency(value) {
  return `NT$ ${numberFormatter.format(Number(value || 0))}`;
}

function formatPercent(value) {
  return percentFormatter.format(Number(value || 0));
}

function toPercent(value, maxValue) {
  if (!maxValue) return 0;
  return Math.max(2, Number(value || 0) / maxValue * 100);
}

function buildEmptyState(message) {
  return `<div class="empty-state">${escapeHtml(message)}</div>`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

refreshButton.addEventListener("click", fetchDashboardData);
applyButton.addEventListener("click", fetchDashboardData);
startDateInput.addEventListener("change", fetchDashboardData);
endDateInput.addEventListener("change", fetchDashboardData);
storeFilter.addEventListener("change", fetchDashboardData);
couponFilter.addEventListener("change", fetchDashboardData);

fetchDashboardData();
