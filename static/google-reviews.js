const refreshButton = document.getElementById("refresh-button");
const refreshStatus = document.getElementById("refresh-status");
const sourcePath = document.getElementById("source-path");
const storeFilter = document.getElementById("store-filter");
const monthFilter = document.getElementById("month-filter");
const ratingFilter = document.getElementById("rating-filter");
const keywordFilter = document.getElementById("keyword-filter");
const metricGrid = document.getElementById("metric-grid");
const monthChart = document.getElementById("month-chart");
const storeChart = document.getElementById("store-chart");
const starChart = document.getElementById("star-chart");
const positiveChart = document.getElementById("positive-chart");
const issueChart = document.getElementById("issue-chart");
const qualityGrid = document.getElementById("quality-grid");
const storesBody = document.getElementById("stores-body");
const lowReviewsBody = document.getElementById("low-reviews-body");
const filesBody = document.getElementById("files-body");
const recordsBody = document.getElementById("records-body");
const monthCaption = document.getElementById("month-caption");
const storeCaption = document.getElementById("store-caption");
const starCaption = document.getElementById("star-caption");
const positiveCaption = document.getElementById("positive-caption");
const issueCaption = document.getElementById("issue-caption");
const qualityCaption = document.getElementById("quality-caption");
const storesTableCaption = document.getElementById("stores-table-caption");
const lowReviewCaption = document.getElementById("low-review-caption");
const filesCaption = document.getElementById("files-caption");
const recordsCaption = document.getElementById("records-caption");

const percentFormatter = new Intl.NumberFormat("zh-TW", {
  style: "percent",
  maximumFractionDigits: 1,
});
const decimalFormatter = new Intl.NumberFormat("zh-TW", {
  minimumFractionDigits: 2,
  maximumFractionDigits: 2,
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
  filters: {
    store: "",
    month: "",
    rating: "",
    keyword: "",
  },
};

async function fetchDashboardData() {
  refreshButton.disabled = true;
  refreshStatus.textContent = "正在同步 Google 評論資料...";

  try {
    const response = await fetch("/api/google-reviews", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`讀取失敗 (${response.status})`);
    }

    state.payload = await response.json();
    syncFilterOptions();
    renderDashboard();

    const generatedAt = state.payload.generated_at ? formatDateTime(state.payload.generated_at) : "剛剛";
    refreshStatus.textContent = `已載入 ${state.payload.file_count} 份檔案，去重後 ${state.payload.deduped_row_count} 筆評論，更新時間 ${generatedAt}。`;
    sourcePath.textContent = state.payload.source_dir || "";
  } catch (error) {
    refreshStatus.textContent = error.message || "讀取失敗";
    metricGrid.innerHTML = "";
    monthChart.innerHTML = buildEmptyState("目前無法讀取 Google 評論資料。");
    storeChart.innerHTML = "";
    starChart.innerHTML = "";
    positiveChart.innerHTML = "";
    issueChart.innerHTML = "";
    qualityGrid.innerHTML = "";
    storesBody.innerHTML = `<tr><td colspan="7" class="empty-cell">${escapeHtml(error.message || "讀取失敗")}</td></tr>`;
    lowReviewsBody.innerHTML = `<tr><td colspan="5" class="empty-cell">目前沒有資料。</td></tr>`;
    filesBody.innerHTML = `<tr><td colspan="5" class="empty-cell">目前沒有資料。</td></tr>`;
    recordsBody.innerHTML = `<tr><td colspan="7" class="empty-cell">目前沒有資料。</td></tr>`;
    sourcePath.textContent = "";
  } finally {
    refreshButton.disabled = false;
  }
}

function syncFilterOptions() {
  const records = state.payload?.records || [];
  const stores = Array.from(new Set(records.map((record) => record.store_label))).sort();
  const months = Array.from(new Set(records.map((record) => record.month))).sort();

  storeFilter.innerHTML = [`<option value="">全部門市</option>`]
    .concat(stores.map((store) => `<option value="${escapeHtml(store)}">${escapeHtml(store)}</option>`))
    .join("");
  monthFilter.innerHTML = [`<option value="">全部月份</option>`]
    .concat(months.map((month) => `<option value="${escapeHtml(month)}">${escapeHtml(month)}</option>`))
    .join("");

  storeFilter.value = stores.includes(state.filters.store) ? state.filters.store : "";
  monthFilter.value = months.includes(state.filters.month) ? state.filters.month : "";
  ratingFilter.value = state.filters.rating;
  keywordFilter.value = state.filters.keyword;
  state.filters.store = storeFilter.value;
  state.filters.month = monthFilter.value;
}

function renderDashboard() {
  const records = getFilteredRecords(state.payload?.records || []);
  const summary = buildSummary(records, state.payload);

  renderMetrics(summary, state.payload);
  renderBarChart(monthChart, summary.months, "month", "review_count", "review_count");
  renderBarChart(storeChart, summary.stores.slice(0, 8), "store_label", "health_score", "health_score");
  renderBarChart(starChart, summary.starDistribution, "label", "review_count", "review_count");
  renderBarChart(positiveChart, summary.positiveThemes.slice(0, 6), "theme", "review_count", "review_count");
  renderBarChart(issueChart, summary.issueThemes.slice(0, 6), "theme", "review_count", "review_count", true);
  renderQuality(summary, state.payload);
  renderStoresTable(summary.stores);
  renderLowReviews(summary.lowReviews);
  renderFiles(state.payload?.files || []);
  renderRecords(records);

  monthCaption.textContent = summary.months.length
    ? `共 ${summary.months.length} 個月份，最高評論量 ${summary.highestMonth.month}`
    : "目前沒有符合條件的月份資料";
  storeCaption.textContent = summary.stores.length
    ? `最佳健康分數 ${summary.stores[0].store_label}｜${summary.stores[0].health_score} 分`
    : "目前沒有符合條件的門市資料";
  starCaption.textContent = summary.lowStarCount
    ? `低星評論 ${summary.lowStarCount} 筆，佔 ${formatPercent(summary.lowStarRate)}`
    : "目前沒有低星評論";
  positiveCaption.textContent = summary.positiveThemes.length
    ? `最常被提到的是 ${summary.positiveThemes[0].theme}`
    : "目前沒有正向主題資料";
  issueCaption.textContent = summary.issueThemes.length
    ? `最常見低星問題是 ${summary.issueThemes[0].theme}`
    : "目前沒有低星問題主題";
  qualityCaption.textContent = `原始 ${state.payload.raw_row_count} 筆，去重後 ${state.payload.deduped_row_count} 筆`;
  storesTableCaption.textContent = `共 ${summary.stores.length} 家門市`;
  lowReviewCaption.textContent = `目前篩選條件下 ${summary.lowReviews.length} 筆`;
  filesCaption.textContent = `共 ${state.payload.file_count} 份來源檔案`;
  recordsCaption.textContent = `共 ${records.length} 筆評論明細`;
}

function getFilteredRecords(records) {
  const keyword = state.filters.keyword.trim().toLowerCase();

  return records.filter((record) => {
    if (state.filters.store && record.store_label !== state.filters.store) {
      return false;
    }
    if (state.filters.month && record.month !== state.filters.month) {
      return false;
    }
    if (state.filters.rating === "low" && Number(record.stars) > 3) {
      return false;
    }
    if ((state.filters.rating === "4" || state.filters.rating === "5") && String(record.stars) !== state.filters.rating) {
      return false;
    }
    if (!keyword) {
      return true;
    }

    const haystack = [
      record.store_label,
      record.title,
      record.name,
      record.text,
      ...(record.positive_themes || []),
      ...(record.issue_themes || []),
    ]
      .join(" ")
      .toLowerCase();

    return haystack.includes(keyword);
  });
}

function buildSummary(records, payload) {
  const groupedByMonth = new Map();
  const groupedByStore = new Map();
  const starCounter = new Map();
  const positiveThemes = new Map();
  const issueThemes = new Map();

  records.forEach((record) => {
    accumulate(groupedByMonth, record.month, "month", record.stars <= 3);
    accumulateStore(groupedByStore, record);
    starCounter.set(record.stars, (starCounter.get(record.stars) || 0) + 1);
    (record.positive_themes || []).forEach((theme) => {
      positiveThemes.set(theme, (positiveThemes.get(theme) || 0) + 1);
    });
    (record.issue_themes || []).forEach((theme) => {
      issueThemes.set(theme, (issueThemes.get(theme) || 0) + 1);
    });
  });

  const months = Array.from(groupedByMonth.values()).sort((left, right) => left.month.localeCompare(right.month));
  const stores = Array.from(groupedByStore.values())
    .map((item) => ({
      ...item,
      average_star: item.review_count ? item.star_sum / item.review_count : 0,
      low_star_rate: item.review_count ? item.low_star_count / item.review_count : 0,
      text_rate: item.review_count ? item.text_count / item.review_count : 0,
    }))
    .sort((left, right) => right.health_score - left.health_score || right.review_count - left.review_count);

  const maxReviews = Math.max(...stores.map((store) => store.review_count), 1);
  stores.forEach((store) => {
    store.health_score = calculateHealthScore(store, maxReviews);
    store.average_star = round(store.average_star);
    store.low_star_rate = round(store.low_star_rate, 4);
    store.text_rate = round(store.text_rate, 4);
  });
  stores.sort((left, right) => right.health_score - left.health_score || right.review_count - left.review_count);

  const lowReviews = records.filter((record) => Number(record.stars) <= 3).slice(0, 40);
  const starDistribution = [1, 2, 3, 4, 5].map((star) => ({
    label: `${star} 星`,
    review_count: starCounter.get(star) || 0,
  }));
  const positiveThemeList = Array.from(positiveThemes.entries())
    .map(([theme, review_count]) => ({ theme, review_count }))
    .sort((left, right) => right.review_count - left.review_count);
  const issueThemeList = Array.from(issueThemes.entries())
    .map(([theme, review_count]) => ({ theme, review_count }))
    .sort((left, right) => right.review_count - left.review_count);
  const highestMonth = months.reduce((best, item) => (!best || item.review_count > best.review_count ? item : best), null);
  const lowStarCount = lowReviews.length;

  return {
    reviewCount: records.length,
    averageStar: records.length ? round(records.reduce((sum, record) => sum + Number(record.stars || 0), 0) / records.length) : 0,
    textRate: records.length ? round(records.filter((record) => record.has_text).length / records.length, 4) : 0,
    fiveStarRate: records.length ? round(records.filter((record) => Number(record.stars) === 5).length / records.length, 4) : 0,
    lowStarCount,
    lowStarRate: records.length ? round(lowStarCount / records.length, 4) : 0,
    months,
    highestMonth,
    stores,
    starDistribution,
    positiveThemes: positiveThemeList,
    issueThemes: issueThemeList,
    lowReviews,
    duplicateRate: payload?.raw_row_count ? round((payload.raw_row_count - payload.deduped_row_count) / payload.raw_row_count, 4) : 0,
  };
}

function accumulate(map, key, keyName, isLowStar) {
  if (!map.has(key)) {
    map.set(key, { [keyName]: key, review_count: 0, low_star_count: 0 });
  }

  const entry = map.get(key);
  entry.review_count += 1;
  if (isLowStar) {
    entry.low_star_count += 1;
  }
}

function accumulateStore(map, record) {
  if (!map.has(record.store_label)) {
    map.set(record.store_label, {
      store_label: record.store_label,
      review_count: 0,
      star_sum: 0,
      low_star_count: 0,
      text_count: 0,
      latest_review_at: record.published_at,
      health_score: 0,
    });
  }

  const entry = map.get(record.store_label);
  entry.review_count += 1;
  entry.star_sum += Number(record.stars || 0);
  entry.low_star_count += Number(record.stars) <= 3 ? 1 : 0;
  entry.text_count += record.has_text ? 1 : 0;
  if (record.published_at > entry.latest_review_at) {
    entry.latest_review_at = record.published_at;
  }
}

function calculateHealthScore(store, maxReviews) {
  const starScore = (store.average_star / 5) * 55;
  const lowStarScore = (1 - store.low_star_rate) * 20;
  const textScore = store.text_rate * 10;
  const volumeScore = (store.review_count / maxReviews) * 15;
  return Math.round(starScore + lowStarScore + textScore + volumeScore);
}

function renderMetrics(summary, payload) {
  const cards = [
    {
      label: "去重後評論數",
      value: formatNumber(summary.reviewCount),
      note: `原始 ${formatNumber(payload.raw_row_count)} 筆`,
    },
    {
      label: "平均星等",
      value: summary.reviewCount ? `${summary.averageStar.toFixed(2)} 星` : "0.00 星",
      note: `5 星率 ${formatPercent(summary.fiveStarRate)}`,
    },
    {
      label: "低星評論",
      value: `${formatNumber(summary.lowStarCount)} 筆`,
      note: `佔比 ${formatPercent(summary.lowStarRate)}`,
    },
    {
      label: "有文字評論率",
      value: formatPercent(summary.textRate),
      note: "可用來衡量內容豐富度",
    },
    {
      label: "重複評論列",
      value: formatNumber(payload.duplicate_count),
      note: `重複率 ${formatPercent(summary.duplicateRate)}`,
    },
    {
      label: "門市數",
      value: `${formatNumber(payload.summary.store_count)} 家`,
      note: payload.summary.latest_review_at ? `最新評論 ${formatDateTime(payload.summary.latest_review_at)}` : "尚無資料",
    },
  ];

  metricGrid.innerHTML = cards
    .map(
      (card) => `
        <article class="metric-card panel">
          <p class="metric-label">${escapeHtml(card.label)}</p>
          <h2>${escapeHtml(card.value)}</h2>
          <p class="metric-note">${escapeHtml(card.note)}</p>
        </article>
      `
    )
    .join("");
}

function renderBarChart(container, items, labelKey, valueKey, countKey = "review_count", negative = false) {
  if (!items.length) {
    container.innerHTML = buildEmptyState("目前沒有符合條件的資料。");
    return;
  }

  const maxValue = Math.max(...items.map((item) => Number(item[valueKey] || 0)), 1);
  container.innerHTML = items
    .map((item) => {
      const value = Number(item[valueKey] || 0);
      const ratio = Math.max(value / maxValue, 0.04);
      const displayValue = valueKey === "health_score"
        ? `${Math.round(value)} 分`
        : valueKey === "review_count"
          ? `${formatNumber(value)} 筆`
          : escapeHtml(String(value));
      return `
        <article class="bar-row ${negative ? "negative" : ""}">
          <div class="bar-copy">
            <p>${escapeHtml(item[labelKey] || "")}</p>
            <span>${formatMeta(item, countKey, valueKey)}</span>
          </div>
          <div class="bar-track">
            <div class="bar-fill" style="width: ${ratio * 100}%"></div>
          </div>
          <strong>${displayValue}</strong>
        </article>
      `;
    })
    .join("");
}

function renderQuality(summary, payload) {
  const qualityCards = [
    ["原始列數", formatNumber(payload.raw_row_count)],
    ["去重後列數", formatNumber(payload.deduped_row_count)],
    ["重複列數", formatNumber(payload.duplicate_count)],
    ["重複率", formatPercent(summary.duplicateRate)],
  ];

  qualityGrid.innerHTML = qualityCards
    .map(
      ([label, value]) => `
        <article class="quality-card">
          <p>${escapeHtml(label)}</p>
          <strong>${escapeHtml(value)}</strong>
        </article>
      `
    )
    .join("");
}

function renderStoresTable(stores) {
  if (!stores.length) {
    storesBody.innerHTML = `<tr><td colspan="7" class="empty-cell">目前沒有符合條件的門市資料。</td></tr>`;
    return;
  }

  storesBody.innerHTML = stores
    .map(
      (store) => `
        <tr>
          <td>${escapeHtml(store.store_label)}</td>
          <td>${formatNumber(store.review_count)}</td>
          <td>${store.average_star.toFixed(2)}</td>
          <td>${formatPercent(store.low_star_rate)}</td>
          <td>${formatPercent(store.text_rate)}</td>
          <td>${formatNumber(store.health_score)}</td>
          <td>${escapeHtml(formatDateTime(store.latest_review_at))}</td>
        </tr>
      `
    )
    .join("");
}

function renderLowReviews(records) {
  if (!records.length) {
    lowReviewsBody.innerHTML = `<tr><td colspan="5" class="empty-cell">目前沒有符合條件的低星評論。</td></tr>`;
    return;
  }

  lowReviewsBody.innerHTML = records
    .map(
      (record) => `
        <tr>
          <td>${escapeHtml(formatDateTime(record.published_at))}</td>
          <td>${escapeHtml(record.store_label)}</td>
          <td>${escapeHtml(`${record.stars} 星`)}</td>
          <td>${escapeHtml((record.issue_themes || []).join("、") || "未分類")}</td>
          <td>${escapeHtml(truncate(record.text || "無文字內容", 140))}</td>
        </tr>
      `
    )
    .join("");
}

function renderFiles(files) {
  if (!files.length) {
    filesBody.innerHTML = `<tr><td colspan="5" class="empty-cell">目前沒有來源檔案資料。</td></tr>`;
    return;
  }

  filesBody.innerHTML = files
    .map(
      (file) => `
        <tr>
          <td>${escapeHtml(file.file_name)}</td>
          <td>${formatNumber(file.raw_row_count)}</td>
          <td>${formatNumber(file.deduped_row_count)}</td>
          <td>${formatNumber(file.duplicate_count)}</td>
          <td>${escapeHtml(formatDateTime(file.modified_at))}</td>
        </tr>
      `
    )
    .join("");
}

function renderRecords(records) {
  if (!records.length) {
    recordsBody.innerHTML = `<tr><td colspan="7" class="empty-cell">目前沒有符合條件的評論。</td></tr>`;
    return;
  }

  recordsBody.innerHTML = records
    .slice(0, 250)
    .map(
      (record) => `
        <tr>
          <td>${escapeHtml(formatDateTime(record.published_at))}</td>
          <td>${escapeHtml(record.store_label)}</td>
          <td>${escapeHtml(`${record.stars} 星`)}</td>
          <td>${escapeHtml(record.name || "未署名")}</td>
          <td>${escapeHtml([...(record.positive_themes || []), ...(record.issue_themes || [])].join("、") || "未分類")}</td>
          <td>${escapeHtml(truncate(record.text || "無文字內容", 180))}</td>
          <td>${escapeHtml(record.file_name)}</td>
        </tr>
      `
    )
    .join("");
}

function formatMeta(item, countKey, valueKey) {
  if (valueKey === "health_score") {
    return `${formatNumber(item.review_count)} 筆｜${Number(item.average_star || 0).toFixed(2)} 星`;
  }
  if (countKey in item) {
    return `${formatNumber(item[countKey])} 筆`;
  }
  return "";
}

function formatNumber(value) {
  return new Intl.NumberFormat("zh-TW").format(Number(value || 0));
}

function formatPercent(value) {
  return percentFormatter.format(Number(value || 0));
}

function formatDateTime(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : dateFormatter.format(date);
}

function truncate(text, maxLength) {
  return text.length > maxLength ? `${text.slice(0, maxLength)}...` : text;
}

function buildEmptyState(text) {
  return `<div class="empty-state">${escapeHtml(text)}</div>`;
}

function round(value, digits = 3) {
  const factor = 10 ** digits;
  return Math.round(Number(value || 0) * factor) / factor;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

storeFilter.addEventListener("change", () => {
  state.filters.store = storeFilter.value;
  renderDashboard();
});

monthFilter.addEventListener("change", () => {
  state.filters.month = monthFilter.value;
  renderDashboard();
});

ratingFilter.addEventListener("change", () => {
  state.filters.rating = ratingFilter.value;
  renderDashboard();
});

keywordFilter.addEventListener("input", () => {
  state.filters.keyword = keywordFilter.value;
  renderDashboard();
});

refreshButton.addEventListener("click", fetchDashboardData);

fetchDashboardData();
window.setInterval(fetchDashboardData, 60000);
