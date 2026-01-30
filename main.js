const state = {
  rawRows: [],
  columns: [],
  mapping: {},
  normalized: [],
  campaigns: [],
  selectedCampaign: null,
  window: "1d",
  chart: null,
  dateRange: { start: null, end: null },
  view: "dashboard",
  keywordFilter: "",
};

const STORAGE_KEY = "coupang-dashboard::autosave";

const REQUIRED_FIELDS = [
  {
    key: "campaign",
    label: "캠페인",
    required: true,
    guesses: ["캠페인명", "캠페인", "campaign"],
  },
  {
    key: "impressions",
    label: "노출수",
    required: true,
    guesses: ["노출수", "노출", "impression"],
  },
  {
    key: "clicks",
    label: "클릭",
    required: true,
    guesses: ["클릭수", "클릭", "click"],
  },
  {
    key: "cost",
    label: "광고비(VAT포함)",
    required: true,
    guesses: ["광고비", "비용", "cost", "spend"],
  },
  {
    key: "orders1",
    label: "총 주문수(1일)",
    required: false,
    guesses: [
      "총주문수1일",
      "총주문수(1일)",
      "총 주문수(1일)",
      "주문수1일",
      "주문수(1일)",
      "주문 수(1일)",
    ],
  },
  {
    key: "orders14",
    label: "총 주문수(14일)",
    required: false,
    guesses: ["총주문수14일", "총주문수(14일)", "총 주문수(14일)"],
  },
  {
    key: "revenue1",
    label: "총 전환매출액(1일)",
    required: false,
    guesses: ["총전환매출액1일", "총전환매출액(1일)", "총 전환매출액(1일)"],
  },
  {
    key: "revenue14",
    label: "총 전환매출액(14일)",
    required: false,
    guesses: ["총전환매출액14일", "총전환매출액(14일)", "총 전환매출액(14일)"],
  },
  {
    key: "sales1",
    label: "총 판매수량(1일)",
    required: false,
    guesses: ["총판매수량1일", "총판매수량(1일)", "총 판매수량(1일)"],
  },
  {
    key: "sales14",
    label: "총 판매수량(14일)",
    required: false,
    guesses: ["총판매수량14일", "총판매수량(14일)", "총 판매수량(14일)"],
  },
  {
    key: "orders",
    label: "주문(대체)",
    required: false,
    guesses: ["주문", "전환", "conversion", "order"],
  },
  {
    key: "revenue",
    label: "광고매출(대체)",
    required: false,
    guesses: ["광고매출", "매출", "sales", "revenue"],
  },
  {
    key: "sales",
    label: "판매량(대체)",
    required: false,
    guesses: ["판매량", "수량", "판매수량"],
  },
  {
    key: "area",
    label: "노출영역(검색/비검색)",
    required: false,
    guesses: ["광고노출지면", "노출지면", "노출영역", "영역", "placement"],
  },
  {
    key: "keyword",
    label: "키워드",
    required: false,
    guesses: ["키워드", "검색어", "keyword"],
  },
  {
    key: "optionId",
    label: "옵션ID",
    required: false,
    guesses: [
      "광고전환매출발생옵션id",
      "광고전환매출발생옵션ID",
      "광고집행옵션id",
      "광고집행옵션ID",
      "옵션id",
      "옵션 ID",
      "옵션번호",
      "option id",
    ],
  },
  {
    key: "optionName",
    label: "상품명",
    required: false,
    guesses: [
      "광고전환매출발생상품명",
      "광고전환매출발생 상품명",
      "광고집행상품명",
      "광고집행 상품명",
      "상품명",
      "옵션명",
      "옵션",
      "option name",
    ],
  },
  {
    key: "saleDate",
    label: "날짜",
    required: false,
    guesses: ["날짜", "판매일", "주문일", "집계일", "집계일자", "일자", "기간", "date"],
  },
];

const els = {
  fileInput: document.getElementById("fileInput"),
  loadJsonInput: document.getElementById("loadJsonInput"),
  saveBtn: document.getElementById("saveBtn"),
  windowSelect: document.getElementById("windowSelect"),
  emptyState: document.getElementById("emptyState"),
  dashboardWrap: document.getElementById("dashboardWrap"),
  kpiGrid: document.getElementById("kpiGrid"),
  trendChart: document.getElementById("trendChart"),
  startDate: document.getElementById("startDate"),
  endDate: document.getElementById("endDate"),
  applyDate: document.getElementById("applyDate"),
  dashboardView: document.getElementById("dashboardView"),
  campaignView: document.getElementById("campaignView"),
  campaignSelect: document.getElementById("campaignSelect"),
  campaignTable: document.getElementById("campaignTable"),
  campaignTabs: document.getElementById("campaignTabs"),
  campaignSectionTitle: document.getElementById("campaignSectionTitle"),
  campaignSectionDesc: document.getElementById("campaignSectionDesc"),
  tabBase: document.getElementById("tab-base"),
  tabSold: document.getElementById("tab-sold"),
  tabOptions: document.getElementById("tab-options"),
  tabKeywords: document.getElementById("tab-keywords"),
  tabExcluded: document.getElementById("tab-excluded"),
  tabManual: document.getElementById("tab-manual"),
  excludedCount: document.getElementById("excludedCount"),
  startDateCampaign: document.getElementById("startDateCampaign"),
  endDateCampaign: document.getElementById("endDateCampaign"),
  applyDateCampaign: document.getElementById("applyDateCampaign"),
  campaignQuickGroup: document.getElementById("campaignQuickGroup"),
};

const fmtNumber = (value, digits = 0) =>
  new Intl.NumberFormat("ko-KR", {
    maximumFractionDigits: digits,
    minimumFractionDigits: digits,
  }).format(isFinite(value) ? value : 0);

const fmtPercent = (value) => `${fmtNumber(value, 2)}%`;

const toNumber = (value) => {
  if (value === null || value === undefined) return 0;
  if (typeof value === "number") return value;
  const cleaned = String(value).replace(/[^0-9eE.+-]/g, "");
  const parsed = parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
};

const getCellSortValue = (cell) => {
  const raw = (cell?.textContent || "").trim();
  if (!raw) return { type: "string", value: "" };
  const numeric = toNumber(raw);
  if (Number.isFinite(numeric) && /[0-9]/.test(raw)) {
    return { type: "number", value: numeric };
  }
  return { type: "string", value: raw.toLowerCase() };
};

const attachTableSort = (table) => {
  if (!table || table.dataset.sortReady === "true") return;
  const headers = Array.from(table.querySelectorAll("thead th"));
  if (!headers.length) return;
  table.dataset.sortReady = "true";

  headers.forEach((th, index) => {
    th.classList.add("sortable");
    th.dataset.sortDir = "none";
    th.setAttribute("aria-sort", "none");
    th.style.cursor = "pointer";
    th.addEventListener("click", () => {
      const tbody = table.querySelector("tbody");
      if (!tbody) return;
      const rows = Array.from(tbody.querySelectorAll("tr"));
      if (!rows.length) return;

      const totalRows = rows.filter((row) =>
        row.textContent?.includes("합계")
      );
      const dataRows = rows.filter((row) => !totalRows.includes(row));

      const current = th.dataset.sortDir || "none";
      const next = current === "asc" ? "desc" : "asc";
      headers.forEach((h) => {
        h.dataset.sortDir = "none";
        h.setAttribute("aria-sort", "none");
      });
      th.dataset.sortDir = next;
      th.setAttribute("aria-sort", next === "asc" ? "ascending" : "descending");

      dataRows.sort((a, b) => {
        const aCell = a.children[index];
        const bCell = b.children[index];
        const aVal = getCellSortValue(aCell);
        const bVal = getCellSortValue(bCell);
        if (aVal.type === "number" && bVal.type === "number") {
          return next === "asc" ? aVal.value - bVal.value : bVal.value - aVal.value;
        }
        if (aVal.value < bVal.value) return next === "asc" ? -1 : 1;
        if (aVal.value > bVal.value) return next === "asc" ? 1 : -1;
        return 0;
      });

      tbody.innerHTML = "";
      dataRows.forEach((row) => tbody.appendChild(row));
      totalRows.forEach((row) => tbody.appendChild(row));
    });
  });
};

const normalizeHeader = (value) =>
  String(value || "")
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[()]/g, "");

const normalizeArea = (value) => {
  const text = String(value || "").trim();
  if (!text) return "기타";
  if (text.includes("비검색") || text.toLowerCase().includes("non")) return "비검색";
  if (text.includes("검색") || text.toLowerCase().includes("search")) return "검색";
  return text;
};

const buildMetric = (totals) => {
  const impressions = totals.impressions || 0;
  const clicks = totals.clicks || 0;
  const orders = totals.orders || 0;
  const cost = totals.cost || 0;
  const revenue = totals.revenue || 0;

  return {
    roas: cost ? (revenue / cost) * 100 : 0,
    cpc: clicks ? cost / clicks : 0,
    ctr: impressions ? (clicks / impressions) * 100 : 0,
    cvr: clicks ? (orders / clicks) * 100 : 0,
    cpa: orders ? cost / orders : 0,
    aov: orders ? revenue / orders : 0,
    cpm: impressions ? (cost / impressions) * 1000 : 0,
  };
};

const getStorageKey = (campaignName) => `excludedKeywords::${campaignName}`;

const loadExcluded = (campaignName) => {
  try {
    const raw = localStorage.getItem(getStorageKey(campaignName));
    return raw ? JSON.parse(raw) : [];
  } catch {
    return [];
  }
};

const saveExcluded = (campaignName, list) => {
  localStorage.setItem(getStorageKey(campaignName), JSON.stringify(list));
};

const parseDate = (value) => {
  if (!value) return null;
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const intVal = Math.trunc(value);
    if (intVal === value && intVal >= 19000101 && intVal <= 21001231) {
      const text = String(intVal);
      const y = parseInt(text.slice(0, 4), 10);
      const m = parseInt(text.slice(4, 6), 10);
      const d = parseInt(text.slice(6, 8), 10);
      return new Date(y, m - 1, d);
    }
    // Excel serial date support
    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(excelEpoch.getTime() + value * 86400000);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const text = String(value).trim().replace(/^'+/, "");
  if (!text) return null;
  if (/[eE]/.test(text)) {
    const num = Number(text);
    if (Number.isFinite(num)) {
      const intVal = Math.trunc(num);
      if (intVal >= 19000101 && intVal <= 21001231) {
        const s = String(intVal);
        const y = parseInt(s.slice(0, 4), 10);
        const m = parseInt(s.slice(4, 6), 10);
        const d = parseInt(s.slice(6, 8), 10);
        return new Date(y, m - 1, d);
      }
      if (intVal > 30000) {
        const excelEpoch = new Date(1899, 11, 30);
        const date = new Date(excelEpoch.getTime() + intVal * 86400000);
        return Number.isNaN(date.getTime()) ? null : date;
      }
    }
  }
  if (/^\\d+$/.test(text) && parseInt(text, 10) > 30000) {
    const serial = parseInt(text, 10);
    const excelEpoch = new Date(1899, 11, 30);
    const date = new Date(excelEpoch.getTime() + serial * 86400000);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  if (/^\\d{8}$/.test(text)) {
    const y = parseInt(text.slice(0, 4), 10);
    const m = parseInt(text.slice(4, 6), 10);
    const d = parseInt(text.slice(6, 8), 10);
    return new Date(y, m - 1, d);
  }
  const korean = text.match(/(\\d{4})\\s*년\\s*(\\d{1,2})\\s*월\\s*(\\d{1,2})\\s*일/);
  if (korean) {
    const y = parseInt(korean[1], 10);
    const m = parseInt(korean[2], 10);
    const d = parseInt(korean[3], 10);
    return new Date(y, m - 1, d);
  }
  const normalized = text.replace(/[./]/g, "-");
  const parts = normalized.split("-").map((v) => parseInt(v, 10));
  if (parts.length >= 3 && parts.every((v) => Number.isFinite(v))) {
    const [y, m, d] = parts;
    return new Date(y, m - 1, d);
  }
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
};

const parseInputDate = (value) => {
  if (!value) return null;
  const parts = String(value).split("-").map((v) => parseInt(v, 10));
  if (parts.length !== 3 || parts.some((v) => !Number.isFinite(v))) return null;
  const [y, m, d] = parts;
  return new Date(y, m - 1, d);
};

const dateToKey = (date) => {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
};

const handleFile = async (file) => {
  if (!file) return;
  const data = await file.arrayBuffer();
  const workbook = XLSX.read(data, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!rows.length) return;

  const headers = rows[0].map((h) => String(h).trim());
  const body = rows.slice(1).filter((row) => row.some((cell) => String(cell).trim()));

  state.columns = headers;
  state.rawRows = body.map((row) => {
    const obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });

  processData();
};

const saveToStorage = () => {
  if (!state.rawRows.length || !state.columns.length) return;
  const payload = {
    version: 2,
    columns: state.columns,
    rawRows: state.rawRows,
    mapping: state.mapping,
    window: state.window,
    dateRange: state.dateRange,
    selectedCampaign: state.selectedCampaign?.name || null,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
};

const loadFromStorage = () => {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return false;
    const payload = JSON.parse(raw);
    if (!payload.rawRows || !payload.columns) return false;
    state.columns = payload.columns;
    state.rawRows = payload.rawRows;
    state.window = payload.window || "1d";
    els.windowSelect.value = state.window;
    processData();
    if (payload.dateRange?.start && payload.dateRange?.end) {
      const start = new Date(payload.dateRange.start);
      const end = new Date(payload.dateRange.end);
      setRange(start, end);
    }
    if (payload.selectedCampaign) {
      selectCampaign(payload.selectedCampaign);
    }
    return true;
  } catch {
    return false;
  }
};

const autoMapColumns = () => {
  const normalized = state.columns.map((col) => normalizeHeader(col));
  const mapping = {};

  REQUIRED_FIELDS.forEach((field) => {
    let matchedIndex = -1;
    for (const guess of field.guesses) {
      const normGuess = normalizeHeader(guess);
      matchedIndex = normalized.findIndex((col) => col.includes(normGuess));
      if (matchedIndex !== -1) break;
    }
    if (matchedIndex !== -1) mapping[field.key] = state.columns[matchedIndex];
  });

  return mapping;
};

const buildMapping = () => {
  state.mapping = autoMapColumns();
};

const normalizeRows = () => {
  const mapped = state.mapping;
  const use14d = state.window === "14d";
  const ordersCol =
    mapped[use14d ? "orders14" : "orders1"] || mapped.orders || "";
  const revenueCol =
    mapped[use14d ? "revenue14" : "revenue1"] || mapped.revenue || "";
  const salesCol =
    mapped[use14d ? "sales14" : "sales1"] || mapped.sales || "";
  const applyVat = (value) => toNumber(value) * 1.1;
  const inferDateValue = (row) => {
    for (const key of state.columns) {
      const value = row[key];
      if (value instanceof Date) return value;
      if (typeof value === "number" && value > 30000) return value;
      if (typeof value === "string") {
        const text = value.trim();
        if (!text) continue;
        if (/^\\d{8}$/.test(text)) return text;
        if (/^\\d{4}[-/.]\\d{1,2}[-/.]\\d{1,2}/.test(text)) return text;
        if (/^\\d+$/.test(text) && parseInt(text, 10) > 30000) return text;
        if (/[eE]/.test(text)) return text;
      }
    }
    return "";
  };

  return state.rawRows.map((row) => ({
    campaign: String(row[mapped.campaign] || "미분류").trim(),
    impressions: toNumber(row[mapped.impressions]),
    clicks: toNumber(row[mapped.clicks]),
    orders: toNumber(row[ordersCol]),
    cost: applyVat(row[mapped.cost]),
    revenue: toNumber(row[revenueCol]),
    sales: toNumber(row[salesCol]) || toNumber(row[ordersCol]),
    area: normalizeArea(row[mapped.area]),
    keyword: (() => {
      const raw = String(row[mapped.keyword] || "").trim();
      return raw === "-" ? "" : raw;
    })(),
    optionId: String(row[mapped.optionId] || "").trim(),
    optionName: String(row[mapped.optionName] || "").trim(),
    saleDate: (() => {
      let value = row[mapped.saleDate];
      if (value === undefined || value === null || value === "") {
        value = inferDateValue(row);
      }
      if (typeof value === "string") return value.trim();
      return value || "";
    })(),
  }));
};

const buildCampaigns = () => {
  const map = new Map();
  state.normalized.forEach((row) => {
    const key = row.campaign || "미분류";
    if (!map.has(key)) {
      map.set(key, { name: key, rows: [] });
    }
    map.get(key).rows.push(row);
  });
  state.campaigns = Array.from(map.values());
};

const buildTotals = (rows) => {
  return rows.reduce(
    (acc, row) => {
      acc.impressions += row.impressions;
      acc.clicks += row.clicks;
      acc.orders += row.orders;
      acc.cost += row.cost;
      acc.revenue += row.revenue;
      acc.sales += row.sales;
      return acc;
    },
    { impressions: 0, clicks: 0, orders: 0, cost: 0, revenue: 0, sales: 0 }
  );
};

const filterRowsByDate = (rows) => {
  const { start, end } = state.dateRange;
  if (!start || !end) return rows;
  return rows.filter((row) => {
    const date = parseDate(row.saleDate);
    if (!date) return false;
    return date >= start && date <= end;
  });
};

const renderKpis = (totals) => {
  const metrics = buildMetric(totals);
  const cards = [
    { label: "광고비 지출", value: `${fmtNumber(totals.cost)}원` },
    { label: "광고매출", value: `${fmtNumber(totals.revenue)}원` },
    { label: "광고매출 ROAS", value: fmtPercent(metrics.roas) },
  ];
  els.kpiGrid.innerHTML = cards
    .map(
      (card) => `
        <div class="kpi-card">
          <div class="label">${card.label}</div>
          <div class="value">${card.value}</div>
        </div>`
    )
    .join("");
};

const renderChart = (rows) => {
  const map = new Map();
  rows.forEach((row) => {
    const date = parseDate(row.saleDate);
    if (!date) return;
    const key = dateToKey(date);
    if (!map.has(key)) {
      map.set(key, { date: key, cost: 0, revenue: 0 });
    }
    const item = map.get(key);
    item.cost += row.cost;
    item.revenue += row.revenue;
  });

  const points = Array.from(map.values()).sort((a, b) => a.date.localeCompare(b.date));
  const labels = points.map((p) => p.date.slice(5));
  const costs = points.map((p) => p.cost);
  const revenues = points.map((p) => p.revenue);
  const roas = points.map((p) => (p.cost ? (p.revenue / p.cost) * 100 : 0));

  if (state.chart) {
    state.chart.destroy();
  }

  state.chart = new Chart(els.trendChart, {
    type: "bar",
    data: {
      labels,
      datasets: [
        {
          type: "bar",
          label: "광고비",
          data: costs,
          backgroundColor: "#ff7a00",
          borderRadius: 6,
        },
        {
          type: "bar",
          label: "광고매출",
          data: revenues,
          backgroundColor: "#c00084",
          borderRadius: 6,
        },
        {
          type: "line",
          label: "광고매출 ROAS",
          data: roas,
          borderColor: "#2f6cff",
          borderWidth: 3,
          tension: 0.35,
          yAxisID: "y1",
        },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        y: {
          beginAtZero: true,
          ticks: {
            callback: (value) => fmtNumber(value),
          },
        },
        y1: {
          beginAtZero: true,
          position: "right",
          grid: { drawOnChartArea: false },
          ticks: {
            callback: (value) => `${value}%`,
          },
        },
      },
      plugins: {
        legend: {
          position: "top",
        },
      },
    },
  });
};

const setDefaultRangeFromData = () => {
  const dates = state.normalized
    .map((row) => parseDate(row.saleDate))
    .filter(Boolean)
    .sort((a, b) => a - b);

  if (!dates.length) {
    state.dateRange = { start: null, end: null };
    els.startDate.value = "";
    els.endDate.value = "";
    els.startDateCampaign.value = "";
    els.endDateCampaign.value = "";
    return;
  }

  const start = dates[0];
  const end = dates[dates.length - 1];
  state.dateRange = {
    start,
    end: new Date(end.getFullYear(), end.getMonth(), end.getDate(), 23, 59, 59, 999),
  };
  els.startDate.value = dateToKey(start);
  els.endDate.value = dateToKey(end);
  els.startDateCampaign.value = dateToKey(start);
  els.endDateCampaign.value = dateToKey(end);
};

const renderDashboard = () => {
  const rows = filterRowsByDate(state.normalized);
  const totals = buildTotals(rows);

  els.emptyState.classList.toggle("hidden", rows.length > 0);
  els.dashboardWrap.classList.toggle("hidden", rows.length === 0);

  if (!rows.length) return;
  renderKpis(totals);
  renderChart(rows);
};

const renderCampaignSelect = () => {
  els.campaignSelect.innerHTML = "";
  state.campaigns.forEach((campaign) => {
    const option = document.createElement("option");
    option.value = campaign.name;
    option.textContent = campaign.name;
    if (state.selectedCampaign?.name === campaign.name) {
      option.selected = true;
    }
    els.campaignSelect.appendChild(option);
  });
};

const getDateRangeLabel = () => {
  if (!state.dateRange.start || !state.dateRange.end) return "-";
  return `${dateToKey(state.dateRange.start)} ~ ${dateToKey(state.dateRange.end)}`;
};

const renderCampaignTable = () => {
  if (!els.campaignTable) return;
  const tbody = els.campaignTable.querySelector("tbody");
  tbody.innerHTML = "";

  state.campaigns.forEach((campaign) => {
    const rows = filterRowsByDate(campaign.rows);
    const totals = buildTotals(rows);
    const metrics = buildMetric(totals);

    const tr = document.createElement("tr");
    if (state.selectedCampaign?.name === campaign.name) {
      tr.classList.add("active-row");
    }
    tr.innerHTML = `
      <td>${getDateRangeLabel()}</td>
      <td>${campaign.name}</td>
      <td>${fmtPercent(metrics.roas)}</td>
      <td>${fmtNumber(metrics.cpc, 0)}원</td>
      <td>${fmtPercent(metrics.ctr)}</td>
      <td>${fmtPercent(metrics.cvr)}</td>
      <td>${fmtNumber(totals.sales)}</td>
      <td>${fmtNumber(totals.cost)}원</td>
      <td>${fmtNumber(totals.revenue)}원</td>
      <td>${fmtNumber(metrics.cpa, 0)}원</td>
    `;
    tr.addEventListener("click", () => {
      selectCampaign(campaign.name);
      showView("campaign");
    });
    tbody.appendChild(tr);
  });
};

const renderBaseStats = (rows) => {
  const areaMap = new Map();
  rows.forEach((row) => {
    const key = row.area || "기타";
    if (!areaMap.has(key)) {
      areaMap.set(key, {
        area: key,
        totals: { impressions: 0, clicks: 0, orders: 0, cost: 0, revenue: 0 },
      });
    }
    const item = areaMap.get(key);
    item.totals.impressions += row.impressions;
    item.totals.clicks += row.clicks;
    item.totals.orders += row.orders;
    item.totals.cost += row.cost;
    item.totals.revenue += row.revenue;
  });

  const rowsHtml = Array.from(areaMap.values()).map((item) => {
    const metric = buildMetric({
      impressions: item.totals.impressions,
      clicks: item.totals.clicks,
      orders: item.totals.orders,
      cost: item.totals.cost,
      revenue: item.totals.revenue,
    });
    return `
      <tr>
        <td>${item.area}</td>
        <td>${fmtNumber(item.totals.impressions)}</td>
        <td>${fmtNumber(item.totals.clicks)}</td>
        <td>${fmtNumber(item.totals.orders)}</td>
        <td>${fmtPercent(metric.ctr)}</td>
        <td>${fmtPercent(metric.cvr)}</td>
        <td>${fmtNumber(metric.cpm, 0)}</td>
        <td>${fmtNumber(metric.cpc, 0)}</td>
        <td>${fmtNumber(item.totals.cost)}원</td>
        <td>${fmtNumber(item.totals.revenue)}원</td>
        <td>${fmtPercent(metric.roas)}</td>
        <td>${fmtNumber(metric.cpa, 0)}원</td>
        <td>${fmtNumber(metric.aov, 0)}원</td>
      </tr>`;
  });

  const totals = buildTotals(rows);
  const totalMetric = buildMetric(totals);

  const totalRow = `
    <tr>
      <td><strong>합계</strong></td>
      <td><strong>${fmtNumber(totals.impressions)}</strong></td>
      <td><strong>${fmtNumber(totals.clicks)}</strong></td>
      <td><strong>${fmtNumber(totals.orders)}</strong></td>
      <td><strong>${fmtPercent(totalMetric.ctr)}</strong></td>
      <td><strong>${fmtPercent(totalMetric.cvr)}</strong></td>
      <td><strong>${fmtNumber(totalMetric.cpm, 0)}</strong></td>
      <td><strong>${fmtNumber(totalMetric.cpc, 0)}</strong></td>
      <td><strong>${fmtNumber(totals.cost)}원</strong></td>
      <td><strong>${fmtNumber(totals.revenue)}원</strong></td>
      <td><strong>${fmtPercent(totalMetric.roas)}</strong></td>
      <td><strong>${fmtNumber(totalMetric.cpa, 0)}원</strong></td>
      <td><strong>${fmtNumber(totalMetric.aov, 0)}원</strong></td>
    </tr>`;

  return `
    <div class="table-wrap">
      <table class="table">
        <thead>
          <tr>
            <th>노출 영역</th>
            <th>노출수</th>
            <th>클릭</th>
            <th>주문</th>
            <th>클릭률</th>
            <th>전환율</th>
            <th>CPM</th>
            <th>CPC</th>
            <th>광고비</th>
            <th>광고매출</th>
            <th>ROAS</th>
            <th>전환당비용</th>
            <th>객단가</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml.join("")}
          ${totalRow}
        </tbody>
      </table>
    </div>`;
};

const renderSoldOptions = (rows) => {
  const filtered = rows.filter((row) => row.sales > 0 || row.revenue > 0);
  const map = new Map();

  filtered.forEach((row) => {
    const key = `${row.saleDate}||${row.optionName}||${row.keyword}`;
    if (!map.has(key)) {
      map.set(key, {
        saleDate: row.saleDate || "-",
        optionName: row.optionName || "-",
        keyword: row.keyword || "-",
        revenue: 0,
      });
    }
    map.get(key).revenue += row.revenue;
  });

  const rowsHtml = Array.from(map.values()).map((item) => `
      <tr>
        <td>${item.saleDate}</td>
        <td>${item.optionName}</td>
        <td>${item.keyword}</td>
        <td class="text-right">${fmtNumber(item.revenue)}원</td>
      </tr>`);

  const totals = buildTotals(filtered);
  const metrics = buildMetric(totals);

  return `
    <div class="table-wrap">
      <table class="table">
        <thead>
          <tr>
            <th>판매일</th>
            <th>옵션명</th>
            <th>키워드</th>
            <th>광고매출</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td colspan="4">
              광고비 ${fmtNumber(totals.cost)}원 · 광고매출 ROAS ${fmtPercent(metrics.roas)}
            </td>
          </tr>
          ${rowsHtml.join("") || `<tr><td colspan="4">데이터 없음</td></tr>`}
        </tbody>
      </table>
    </div>`;
};

const renderOptions = (rows) => {
  const map = new Map();
  rows.forEach((row) => {
    const key = `${row.optionId}||${row.optionName}`;
    if (!map.has(key)) {
      map.set(key, {
        optionId: row.optionId || "-",
        optionName: row.optionName || "-",
        totals: {
          sales: 0,
          cost: 0,
          revenue: 0,
          impressions: 0,
          clicks: 0,
          orders: 0,
        },
      });
    }
    const item = map.get(key);
    item.totals.sales += row.sales;
    item.totals.cost += row.cost;
    item.totals.revenue += row.revenue;
    item.totals.impressions += row.impressions;
    item.totals.clicks += row.clicks;
    item.totals.orders += row.orders;
  });

  const rowsHtml = Array.from(map.values()).map((item) => {
    const metric = buildMetric({
      impressions: item.totals.impressions,
      clicks: item.totals.clicks,
      orders: item.totals.orders,
      cost: item.totals.cost,
      revenue: item.totals.revenue,
    });

    return `
      <tr>
        <td>${item.optionId}</td>
        <td>${item.optionName}</td>
        <td>${fmtNumber(item.totals.orders)}</td>
        <td>${fmtNumber(item.totals.cost)}원</td>
        <td>${fmtNumber(item.totals.revenue)}원</td>
        <td>${fmtPercent(metric.roas)}</td>
        <td>${fmtNumber(item.totals.impressions)}</td>
        <td>${fmtNumber(item.totals.clicks)}</td>
        <td>${fmtPercent(metric.ctr)}</td>
        <td>${fmtPercent(metric.cvr)}</td>
      </tr>`;
  });

  return `
    <div class="table-wrap">
      <table class="table">
        <thead>
          <tr>
            <th>옵션ID</th>
            <th>옵션명</th>
            <th>주문</th>
            <th>광고비</th>
            <th>광고매출</th>
            <th>ROAS</th>
            <th>노출</th>
            <th>클릭</th>
            <th>클릭률</th>
            <th>전환율</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml.join("") || `<tr><td colspan="10">데이터 없음</td></tr>`}
        </tbody>
      </table>
    </div>`;
};

const renderKeywords = (rows) => {
  const map = new Map();
  rows.forEach((row) => {
    if (!row.keyword) return;
    if (state.keywordFilter && !row.keyword.includes(state.keywordFilter)) return;
    const key = row.keyword;
    if (!map.has(key)) {
      map.set(key, {
        keyword: key,
        totals: { impressions: 0, clicks: 0, orders: 0, cost: 0, revenue: 0 },
      });
    }
    const item = map.get(key);
    item.totals.impressions += row.impressions;
    item.totals.clicks += row.clicks;
    item.totals.orders += row.orders;
    item.totals.cost += row.cost;
    item.totals.revenue += row.revenue;
  });

  const excluded = new Set(loadExcluded(state.selectedCampaign?.name || ""));
  const rowsHtml = Array.from(map.values()).map((item) => {
    const metric = buildMetric({
      impressions: item.totals.impressions,
      clicks: item.totals.clicks,
      orders: item.totals.orders,
      cost: item.totals.cost,
      revenue: item.totals.revenue,
    });

    return `
      <tr data-keyword="${item.keyword}">
        <td>${item.keyword}</td>
        <td>${fmtNumber(item.totals.impressions)}</td>
        <td>${fmtNumber(item.totals.clicks)}</td>
        <td>${fmtPercent(metric.ctr)}</td>
        <td>${fmtNumber(item.totals.orders)}</td>
        <td>${fmtPercent(metric.cvr)}</td>
        <td>${fmtNumber(metric.cpc, 0)}</td>
        <td>${fmtNumber(item.totals.cost)}원</td>
        <td>${fmtNumber(item.totals.revenue)}원</td>
        <td>${fmtPercent(metric.roas)}</td>
        <td><input type="checkbox" ${excluded.has(item.keyword) ? "checked" : ""} /></td>
      </tr>`;
  });

  return `
    <div class="keyword-actions">
      <input id="keywordSearch" type="search" placeholder="키워드명으로 검색하세요." value="${state.keywordFilter}" />
      <button id="keywordSearchBtn" class="primary">검색</button>
      <div class="action-group">
        <button id="excludeBtn" class="action-btn">제외키워드 등록하기</button>
        <button class="action-btn secondary">수동 입찰가 관리 등록하기</button>
      </div>
    </div>
    <div class="table-wrap">
      <table class="table">
        <thead>
          <tr>
            <th>키워드</th>
            <th>노출</th>
            <th>클릭</th>
            <th>클릭률</th>
            <th>주문</th>
            <th>전환율</th>
            <th>CPC</th>
            <th>광고비</th>
            <th>광고매출</th>
            <th>ROAS</th>
            <th>제외</th>
          </tr>
        </thead>
        <tbody>
          ${rowsHtml.join("") || `<tr><td colspan="11">데이터 없음</td></tr>`}
        </tbody>
      </table>
    </div>`;
};

const renderExcluded = () => {
  const excluded = loadExcluded(state.selectedCampaign?.name || "");
  els.excludedCount.textContent = excluded.length;

  return `
    <div class="table-wrap">
      <table class="table">
        <thead>
          <tr>
            <th>키워드</th>
            <th>작업</th>
          </tr>
        </thead>
        <tbody>
          ${excluded
            .map(
              (keyword) => `
                <tr data-keyword="${keyword}">
                  <td>${keyword}</td>
                  <td><button class="ghost remove-excluded">삭제</button></td>
                </tr>`
            )
            .join("") || `<tr><td colspan="2">제외 키워드 없음</td></tr>`}
        </tbody>
      </table>
    </div>`;
};

const renderManual = () => `
  <div class="empty">
    <h3>수동 입찰가 관리</h3>
    <p class="muted">등록된 항목이 없습니다.</p>
  </div>`;

const renderCampaignTabs = (tab) => {
  const tabs = {
    base: els.tabBase,
    sold: els.tabSold,
    options: els.tabOptions,
    keywords: els.tabKeywords,
    excluded: els.tabExcluded,
    manual: els.tabManual,
  };

  Object.entries(tabs).forEach(([key, panel]) => {
    panel.classList.toggle("hidden", key !== tab);
  });

  els.campaignTabs.querySelectorAll("button").forEach((button) => {
    button.classList.toggle("active", button.dataset.tab === tab);
  });

  const descMap = {
    base: "ROAS가 높은 노출영역으로 예산이 잘 쓰이고 있는지 확인하세요.",
    sold: "해당 광고캠페인에서 어떤 옵션이 팔렸는지 한눈에 확인할 수 있습니다.",
    options: "판매가 잘 일어나지 않는 옵션은 광고 캠페인에서 OFF 해주세요.",
    keywords: "성과가 없는 키워드는 제외키워드로 등록하세요.",
    excluded: "제외 키워드를 관리할 수 있습니다.",
    manual: "수동 입찰가 관리 항목을 관리합니다.",
  };

  const titleMap = {
    base: "기본통계",
    sold: "팔린 옵션",
    options: "전체 옵션",
    keywords: "전체 키워드",
    excluded: "제외 키워드",
    manual: "수동 입찰가 관리",
  };

  els.campaignSectionTitle.textContent = titleMap[tab];
  els.campaignSectionDesc.textContent = descMap[tab];
};

const renderCampaignView = (tab = "base") => {
  if (!state.selectedCampaign) return;
  const rows = filterRowsByDate(state.selectedCampaign.rows);

  renderCampaignTabs(tab);

  if (tab === "base") {
    els.tabBase.innerHTML = renderBaseStats(rows);
    attachTableSort(els.tabBase.querySelector("table"));
  } else if (tab === "sold") {
    els.tabSold.innerHTML = renderSoldOptions(rows);
    attachTableSort(els.tabSold.querySelector("table"));
  } else if (tab === "options") {
    els.tabOptions.innerHTML = renderOptions(rows);
    attachTableSort(els.tabOptions.querySelector("table"));
  } else if (tab === "keywords") {
    els.tabKeywords.innerHTML = renderKeywords(rows);
    attachTableSort(els.tabKeywords.querySelector("table"));

    const searchInput = document.getElementById("keywordSearch");
    const searchBtn = document.getElementById("keywordSearchBtn");
    const excludeBtn = document.getElementById("excludeBtn");

    if (searchBtn) {
      searchBtn.addEventListener("click", () => {
        state.keywordFilter = searchInput.value.trim();
        renderCampaignView("keywords");
      });
    }

    if (excludeBtn) {
      excludeBtn.addEventListener("click", () => {
        const checked = Array.from(els.tabKeywords.querySelectorAll("tbody input[type='checkbox']"))
          .filter((input) => input.checked)
          .map((input) => input.closest("tr")?.dataset.keyword)
          .filter(Boolean);
        const current = loadExcluded(state.selectedCampaign.name);
        const next = Array.from(new Set([...current, ...checked]));
        saveExcluded(state.selectedCampaign.name, next);
        renderCampaignView("excluded");
      });
    }
  } else if (tab === "excluded") {
    els.tabExcluded.innerHTML = renderExcluded();
    attachTableSort(els.tabExcluded.querySelector("table"));
    els.tabExcluded.querySelectorAll(".remove-excluded").forEach((btn) => {
      btn.addEventListener("click", (event) => {
        const keyword = event.target.closest("tr")?.dataset.keyword;
        if (!keyword) return;
        const current = loadExcluded(state.selectedCampaign.name);
        const next = current.filter((item) => item !== keyword);
        saveExcluded(state.selectedCampaign.name, next);
        renderCampaignView("excluded");
      });
    });
  } else if (tab === "manual") {
    els.tabManual.innerHTML = renderManual();
  }

  els.excludedCount.textContent = loadExcluded(state.selectedCampaign.name).length;
  renderCampaignTable();
};

const showView = (view) => {
  state.view = view;
  els.dashboardView.classList.toggle("hidden", view !== "dashboard");
  els.campaignView.classList.toggle("hidden", view !== "campaign");
  document.querySelectorAll(".nav-item").forEach((item) => {
    item.classList.toggle("active", item.dataset.view === view);
  });
};

const selectCampaign = (name) => {
  state.selectedCampaign = state.campaigns.find((c) => c.name === name) || null;
  if (!state.selectedCampaign) return;
  renderCampaignSelect();
  renderCampaignView("base");
  showView("campaign");
};

const processData = () => {
  buildMapping();
  state.normalized = normalizeRows();
  buildCampaigns();
  setDefaultRangeFromData();
  renderDashboard();
  if (state.campaigns.length) {
    selectCampaign(state.campaigns[0].name);
  }
  renderCampaignTable();
  saveToStorage();
};

const handleWindowChange = () => {
  state.window = els.windowSelect.value;
  if (!state.rawRows.length) return;
  processData();
};

const applyDate = (startInput, endInput) => {
  const start = parseInputDate(startInput.value);
  const end = parseInputDate(endInput.value);
  if (!start || !end) return;
  setRange(start, end);
};

const setRange = (start, end) => {
  const startKey = dateToKey(start);
  const endKey = dateToKey(end);
  els.startDate.value = startKey;
  els.endDate.value = endKey;
  els.startDateCampaign.value = startKey;
  els.endDateCampaign.value = endKey;
  end.setHours(23, 59, 59, 999);
  state.dateRange = { start, end };
  renderDashboard();
  renderCampaignView("base");
  renderCampaignTable();
};

const handleQuickRange = (event, group) => {
  const btn = event.target.closest("button");
  if (!btn) return;
  const range = btn.dataset.range;
  if (!range) return;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  let start = new Date(today);
  let end = new Date(today);

  if (range === "prev") {
    if (state.dateRange.start && state.dateRange.end) {
      const diff = state.dateRange.end.getTime() - state.dateRange.start.getTime();
      end = new Date(state.dateRange.start.getTime() - 86400000);
      start = new Date(end.getTime() - diff);
    }
  } else if (range === "next") {
    if (state.dateRange.start && state.dateRange.end) {
      const diff = state.dateRange.end.getTime() - state.dateRange.start.getTime();
      start = new Date(state.dateRange.end.getTime() + 86400000);
      end = new Date(start.getTime() + diff);
    }
  } else if (range === "yesterday") {
    start.setDate(start.getDate() - 1);
    end = new Date(start);
  } else if (range === "3d") {
    start.setDate(start.getDate() - 2);
  } else if (range === "1w") {
    start.setDate(start.getDate() - 6);
  } else if (range === "2w") {
    start.setDate(start.getDate() - 13);
  } else if (range === "1m") {
    start.setMonth(start.getMonth() - 1);
    start.setDate(start.getDate() + 1);
  } else if (range === "2m") {
    start.setMonth(start.getMonth() - 2);
    start.setDate(start.getDate() + 1);
  } else if (range === "this") {
    start = new Date(today.getFullYear(), today.getMonth(), 1);
  } else if (range === "last") {
    start = new Date(today.getFullYear(), today.getMonth() - 1, 1);
    end = new Date(today.getFullYear(), today.getMonth(), 0);
  }

  setRange(start, end);

  group.querySelectorAll("button").forEach((button) => {
    button.classList.toggle("active", button === btn);
  });
};

const handleSave = () => {
  if (!state.normalized.length) {
    alert("저장할 분석 데이터가 없습니다.");
    return;
  }

  const payload = {
    version: 1,
    createdAt: new Date().toISOString(),
    mapping: state.mapping,
    rows: state.normalized,
    dateRange: state.dateRange,
    window: state.window,
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `coupang-dashboard-${Date.now()}.json`;
  a.click();
  URL.revokeObjectURL(url);
};

const handleLoad = async (file) => {
  if (!file) return;
  const text = await file.text();
  const payload = JSON.parse(text);

  if (!payload.rows || !Array.isArray(payload.rows)) {
    alert("올바른 저장 파일이 아닙니다.");
    return;
  }

  state.mapping = payload.mapping || {};
  state.normalized = payload.rows;
  state.window = payload.window || "1d";
  els.windowSelect.value = state.window;

  if (payload.dateRange?.start && payload.dateRange?.end) {
    const start = new Date(payload.dateRange.start);
    const end = new Date(payload.dateRange.end);
    state.dateRange = { start, end };
    els.startDate.value = dateToKey(start);
    els.endDate.value = dateToKey(end);
    els.startDateCampaign.value = dateToKey(start);
    els.endDateCampaign.value = dateToKey(end);
  } else {
    setDefaultRangeFromData();
  }

  buildCampaigns();
  renderDashboard();
  if (state.campaigns.length) {
    selectCampaign(state.campaigns[0].name);
  }
  renderCampaignTable();
  saveToStorage();
};

els.fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  handleFile(file);
});

els.saveBtn.addEventListener("click", handleSave);
els.loadJsonInput.addEventListener("change", (event) => handleLoad(event.target.files[0]));
els.windowSelect.addEventListener("change", handleWindowChange);

els.applyDate.addEventListener("click", () => applyDate(els.startDate, els.endDate));
els.applyDateCampaign.addEventListener("click", () =>
  applyDate(els.startDateCampaign, els.endDateCampaign)
);

els.campaignQuickGroup.addEventListener("click", (event) =>
  handleQuickRange(event, els.campaignQuickGroup)
);

document.querySelector(".quick-group").addEventListener("click", (event) =>
  handleQuickRange(event, document.querySelector(".quick-group"))
);

els.campaignSelect.addEventListener("change", (event) => {
  selectCampaign(event.target.value);
});

els.campaignTabs.addEventListener("click", (event) => {
  const btn = event.target.closest("button");
  if (!btn) return;
  renderCampaignView(btn.dataset.tab || "base");
});

document.querySelectorAll(".nav-item").forEach((item) => {
  item.addEventListener("click", () => {
    if (!item.dataset.view) return;
    showView(item.dataset.view);
  });
});

const setupDragAndDrop = () => {
  const target = els.emptyState;
  if (!target) return;
  const prevent = (event) => {
    event.preventDefault();
    event.stopPropagation();
  };
  ["dragenter", "dragover"].forEach((type) => {
    target.addEventListener(type, (event) => {
      prevent(event);
      target.classList.add("drag-over");
    });
  });
  ["dragleave", "drop"].forEach((type) => {
    target.addEventListener(type, (event) => {
      prevent(event);
      target.classList.remove("drag-over");
    });
  });
  target.addEventListener("drop", (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) handleFile(file);
  });
};

setupDragAndDrop();
loadFromStorage();
