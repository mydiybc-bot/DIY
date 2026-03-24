const refreshButton = document.getElementById("refresh-button");
const refreshStatus = document.getElementById("refresh-status");
const sourcePath = document.getElementById("source-path");
const unitFilter = document.getElementById("unit-filter");
const monthFilter = document.getElementById("month-filter");
const keywordFilter = document.getElementById("keyword-filter");
const includeZeroFilter = document.getElementById("include-zero-filter");
const metricGrid = document.getElementById("metric-grid");
const monthChart = document.getElementById("month-chart");
const unitChart = document.getElementById("unit-chart");
const categoryChart = document.getElementById("category-chart");
const fileGrid = document.getElementById("file-grid");
const recordsBody = document.getElementById("records-body");
const monthCaption = document.getElementById("month-caption");
const unitCaption = document.getElementById("unit-caption");
const categoryCaption = document.getElementById("category-caption");
const filesCaption = document.getElementById("files-caption");
const tableCaption = document.getElementById("table-caption");

const currencyFormatter = new Intl.NumberFormat("zh-TW");
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
    unit: "",
    month: "",
    keyword: "",
    includeZero: true,
  },
};

async function fetchDashboardData() {
  refreshButton.disabled = true;
  refreshStatus.textContent = "正在同步最新請款資料...";

  try {
    const response = await fetch("/api/reimbursements", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`讀取失敗 (${response.status})`);
    }

    state.payload = await response.json();
    syncFilterOptions();
    renderDashboard();

    const generatedAt = state.payload.generated_at ? formatDateTime(state.payload.generated_at) : "剛剛";
    const fileCount = state.payload.file_count || 0;
    const recordCount = state.payload.record_count || 0;
    const errorCount = (state.payload.errors || []).length;
    refreshStatus.textContent = `已載入 ${fileCount} 份檔案、${recordCount} 筆明細，更新時間 ${generatedAt}${errorCount ? `，另有 ${errorCount} 份檔案讀取失敗` : ""}。`;
    sourcePath.textContent = state.payload.source_dir || "";
  } catch (error) {
    metricGrid.innerHTML = "";
    monthChart.innerHTML = buildEmptyState("目前無法讀取儀表板資料。");
    unitChart.innerHTML = buildEmptyState("請稍後再試。");
    categoryChart.innerHTML = "";
    fileGrid.innerHTML = "";
    recordsBody.innerHTML = `<tr><td colspan="7" class="empty-cell">${escapeHtml(error.message || "讀取失敗")}</td></tr>`;
    refreshStatus.textContent = error.message || "讀取失敗";
    sourcePath.textContent = "";
  } finally {
    refreshButton.disabled = false;
  }
}

function syncFilterOptions() {
  const payload = state.payload;
  if (!payload) return;

  const units = Array.from(new Set((payload.records || []).map((record) => record.unit).filter(Boolean))).sort();
  const months = Array.from(new Set((payload.records || []).map((record) => record.month).filter(Boolean))).sort();

  const currentUnit = state.filters.unit;
  const currentMonth = state.filters.month;

  unitFilter.innerHTML = [`<option value="">全部單位</option>`]
    .concat(units.map((unit) => `<option value="${escapeHtml(unit)}">${escapeHtml(unit)}</option>`))
    .join("");
  monthFilter.innerHTML = [`<option value="">全部月份</option>`]
    .concat(months.map((month) => `<option value="${escapeHtml(month)}">${escapeHtml(month)}</option>`))
    .join("");

  if (units.includes(currentUnit)) {
    unitFilter.value = currentUnit;
  } else {
    state.filters.unit = "";
    unitFilter.value = "";
  }

  if (months.includes(currentMonth)) {
    monthFilter.value = currentMonth;
  } else {
    state.filters.month = "";
    monthFilter.value = "";
  }

  keywordFilter.value = state.filters.keyword;
  includeZeroFilter.checked = state.filters.includeZero;
}

function renderDashboard() {
  const payload = state.payload;
  if (!payload) return;

  const records = getFilteredRecords(payload.records || []);
  const files = buildFilteredFileSummaries(records);
  const summary = buildSummary(records);

  renderMetrics(summary);
  renderBarChart(monthChart, summary.months, "month", "amount", "month");
  renderBarChart(unitChart, summary.units, "unit", "amount", "unit");
  renderBarChart(categoryChart, summary.categories.slice(0, 8), "summary", "amount", "category");
  renderFiles(files);
  renderTable(records);

  monthCaption.textContent = summary.highestMonth
    ? `最高月份 ${summary.highestMonth.month}，共 ${formatCurrency(summary.highestMonth.amount)}`
    : "目前沒有符合條件的月份資料";
  unitCaption.textContent = summary.units.length
    ? `共 ${summary.units.length} 個單位，最高為 ${summary.units[0].unit}`
    : "目前沒有符合條件的單位資料";
  categoryCaption.textContent = summary.categories.length
    ? `前 ${Math.min(summary.categories.length, 8)} 名依金額排序`
    : "目前沒有符合條件的項目資料";
  filesCaption.textContent = files.length
    ? `目前篩選後涵蓋 ${files.length} 份檔案`
    : "目前沒有符合條件的檔案";
  tableCaption.textContent = `共 ${records.length} 筆明細`;
}

function getFilteredRecords(records) {
  const keyword = state.filters.keyword.trim().toLowerCase();

  return records.filter((record) => {
    if (state.filters.unit && record.unit !== state.filters.unit) {
      return false;
    }
    if (state.filters.month && record.month !== state.filters.month) {
      return false;
    }
    if (!state.filters.includeZero && Number(record.amount) === 0) {
      return false;
    }
    if (!keyword) {
      return true;
    }

    const haystack = [
      record.summary,
      record.note,
      record.invoice,
      record.file_name,
      record.payee,
      record.code,
    ]
      .join(" ")
      .toLowerCase();

    return haystack.includes(keyword);
  });
}

function buildSummary(records) {
  const totalAmount = records.reduce((sum, record) => sum + Number(record.amount || 0), 0);
  const nonzeroRecords = records.filter((record) => Number(record.amount || 0) > 0);
  const zeroAmountCount = records.length - nonzeroRecords.length;

  const byMonth = new Map();
  const byUnit = new Map();
  const byCategory = new Map();

  records.forEach((record) => {
    accumulate(byMonth, record.month || "未分類月份", record.amount, "month");
    accumulate(byUnit, record.unit || "未分類單位", record.amount, "unit");
    accumulate(byCategory, record.summary || "未填項目", record.amount, "summary");
  });

  const months = Array.from(byMonth.values()).sort((left, right) => left.month.localeCompare(right.month));
  const units = Array.from(byUnit.values()).sort((left, right) => right.amount - left.amount);
  const categories = Array.from(byCategory.values()).sort((left, right) => right.amount - left.amount);
  const highestMonth = months.reduce((best, item) => (!best || item.amount > best.amount ? item : best), null);
  const largestRecord = records.reduce((best, item) => (!best || item.amount > best.amount ? item : best), null);

  return {
    totalAmount,
    recordCount: records.length,
    nonzeroRecordCount: nonzeroRecords.length,
    zeroAmountCount,
    averageAmount: nonzeroRecords.length ? totalAmount / nonzeroRecords.length : 0,
    highestMonth,
    largestRecord,
    months,
    units,
    categories,
  };
}

function accumulate(map, key, amount, labelKey) {
  if (!map.has(key)) {
    map.set(key, { [labelKey]: key, amount: 0, recordCount: 0 });
  }

  const entry = map.get(key);
  entry.amount += Number(amount || 0);
  entry.recordCount += 1;
}

function renderMetrics(summary) {
  const topUnit = summary.units[0];
  const largestRecord = summary.largestRecord;
  const highestMonth = summary.highestMonth;

  const cards = [
    {
      label: "總請款金額",
      value: formatCurrency(summary.totalAmount),
      note: `目前篩選條件下共 ${summary.recordCount} 筆`,
    },
    {
      label: "有效請款筆數",
      value: `${summary.nonzeroRecordCount} 筆`,
      note: summary.zeroAmountCount ? `另有 ${summary.zeroAmountCount} 筆 0 元項目` : "沒有 0 元項目",
    },
    {
      label: "平均每筆金額",
      value: formatCurrency(summary.averageAmount),
      note: "以非 0 元明細計算",
    },
    {
      label: "最大單筆",
      value: largestRecord ? formatCurrency(largestRecord.amount) : "NT$0",
      note: largestRecord ? `${largestRecord.summary}｜${largestRecord.unit}` : "目前沒有資料",
    },
    {
      label: "最高月份",
      value: highestMonth ? highestMonth.month : "無資料",
      note: highestMonth ? formatCurrency(highestMonth.amount) : "目前沒有資料",
    },
    {
      label: "最高單位",
      value: topUnit ? topUnit.unit : "無資料",
      note: topUnit ? formatCurrency(topUnit.amount) : "目前沒有資料",
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

function renderBarChart(container, items, labelKey, valueKey, style) {
  if (!items.length) {
    container.innerHTML = buildEmptyState("目前沒有符合條件的資料。");
    return;
  }

  const maxValue = Math.max(...items.map((item) => Number(item[valueKey] || 0)), 1);
  container.innerHTML = items
    .map((item) => {
      const value = Number(item[valueKey] || 0);
      const ratio = Math.max(value / maxValue, 0.04);
      return `
        <article class="bar-row bar-row-${style}">
          <div class="bar-copy">
            <p>${escapeHtml(item[labelKey])}</p>
            <span>${item.recordCount} 筆</span>
          </div>
          <div class="bar-track">
            <div class="bar-fill" style="width: ${ratio * 100}%"></div>
          </div>
          <strong>${escapeHtml(formatCurrency(value))}</strong>
        </article>
      `;
    })
    .join("");
}

function buildFilteredFileSummaries(records) {
  const grouped = new Map();

  records.forEach((record) => {
    const key = record.file_name;
    if (!grouped.has(key)) {
      grouped.set(key, {
        file_name: record.file_name,
        unit: record.unit,
        date: record.date,
        total_amount: 0,
        record_count: 0,
        zero_amount_count: 0,
      });
    }

    const entry = grouped.get(key);
    entry.total_amount += Number(record.amount || 0);
    entry.record_count += 1;
    if (Number(record.amount || 0) === 0) {
      entry.zero_amount_count += 1;
    }
  });

  return Array.from(grouped.values()).sort((left, right) => {
    const dateCompare = (right.date || "").localeCompare(left.date || "");
    if (dateCompare !== 0) return dateCompare;
    return right.total_amount - left.total_amount;
  });
}

function renderFiles(files) {
  if (!files.length) {
    fileGrid.innerHTML = buildEmptyState("目前沒有符合條件的檔案。");
    return;
  }

  fileGrid.innerHTML = files
    .map(
      (file) => `
        <article class="file-card">
          <p class="file-date">${escapeHtml(file.date || "未標示日期")}</p>
          <h3>${escapeHtml(file.unit || "未標示單位")}</h3>
          <p class="file-total">${escapeHtml(formatCurrency(file.total_amount))}</p>
          <p class="file-meta">${file.record_count} 筆明細${file.zero_amount_count ? `，${file.zero_amount_count} 筆 0 元` : ""}</p>
          <p class="file-name">${escapeHtml(file.file_name)}</p>
        </article>
      `
    )
    .join("");
}

function renderTable(records) {
  if (!records.length) {
    recordsBody.innerHTML = `<tr><td colspan="7" class="empty-cell">目前沒有符合條件的明細。</td></tr>`;
    return;
  }

  recordsBody.innerHTML = records
    .map(
      (record) => `
        <tr>
          <td>${escapeHtml(record.date || "")}</td>
          <td>${escapeHtml(record.unit || "")}</td>
          <td>
            <div class="primary-cell">${escapeHtml(record.summary || "")}</div>
            <div class="secondary-cell">${escapeHtml(record.payee || "")}</div>
          </td>
          <td class="amount-cell">${escapeHtml(formatCurrency(record.amount))}</td>
          <td>${escapeHtml(record.invoice || "")}</td>
          <td>${escapeHtml(record.note || "")}</td>
          <td>${escapeHtml(record.file_name || "")}</td>
        </tr>
      `
    )
    .join("");
}

function formatCurrency(value) {
  return `NT$${currencyFormatter.format(Math.round(Number(value || 0)))}`;
}

function formatDateTime(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? value : dateFormatter.format(date);
}

function buildEmptyState(text) {
  return `<div class="empty-state">${escapeHtml(text)}</div>`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

unitFilter.addEventListener("change", () => {
  state.filters.unit = unitFilter.value;
  renderDashboard();
});

monthFilter.addEventListener("change", () => {
  state.filters.month = monthFilter.value;
  renderDashboard();
});

keywordFilter.addEventListener("input", () => {
  state.filters.keyword = keywordFilter.value;
  renderDashboard();
});

includeZeroFilter.addEventListener("change", () => {
  state.filters.includeZero = includeZeroFilter.checked;
  renderDashboard();
});

refreshButton.addEventListener("click", fetchDashboardData);

fetchDashboardData();
window.setInterval(fetchDashboardData, 60000);
