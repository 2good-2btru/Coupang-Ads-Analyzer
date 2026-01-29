const state = {
  rawRows: [],
  columns: [],
  mapping: {},
  normalized: [],
  campaigns: [],
  selectedCampaign: null,
  window: "1d",
};

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
    guesses: ["총주문수1일", "총주문수(1일)", "총 주문수(1일)"],
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
    guesses: ["광고전환매출발생옵션id", "광고전환매출발생옵션ID", "광고집행옵션id", "광고집행옵션ID", "옵션id", "옵션 ID", "옵션번호", "option id"],
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
    label: "판매일",
    required: false,
    guesses: ["날짜", "판매일", "주문일", "date"],
  },
];

const els = {
  fileInput: document.getElementById("fileInput"),
  loadJsonInput: document.getElementById("loadJsonInput"),
  saveBtn: document.getElementById("saveBtn"),
  windowSelect: document.getElementById("windowSelect"),
  mappingPanel: document.getElementById("mappingPanel"),
  mappingGrid: document.getElementById("mappingGrid"),
  applyMapping: document.getElementById("applyMapping"),
  campaignList: document.getElementById("campaignList"),
  campaignSearch: document.getElementById("campaignSearch"),
  emptyState: document.getElementById("emptyState"),
  detailWrap: document.getElementById("detailWrap"),
  campaignTitle: document.getElementById("campaignTitle"),
  campaignMeta: document.getElementById("campaignMeta"),
  summaryCards: document.getElementById("summaryCards"),
  tabs: document.getElementById("tabs"),
  tabBase: document.getElementById("tab-base"),
  tabSold: document.getElementById("tab-sold"),
  tabOptions: document.getElementById("tab-options"),
  tabKeywords: document.getElementById("tab-keywords"),
  tabExcluded: document.getElementById("tab-excluded"),
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
  const cleaned = String(value).replace(/[^0-9.-]/g, "");
  const parsed = parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
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

  buildMappingUI();
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

const buildMappingUI = () => {
  state.mapping = autoMapColumns();
  els.mappingPanel.classList.remove("hidden");
  els.mappingGrid.innerHTML = "";

  REQUIRED_FIELDS.forEach((field) => {
    const wrapper = document.createElement("div");
    wrapper.className = "mapping-item";
    const label = document.createElement("label");
    label.textContent = `${field.label}${field.required ? " *" : ""}`;

    const select = document.createElement("select");
    const none = document.createElement("option");
    none.value = "";
    none.textContent = "없음";
    select.appendChild(none);

    state.columns.forEach((col) => {
      const option = document.createElement("option");
      option.value = col;
      option.textContent = col;
      if (state.mapping[field.key] === col) option.selected = true;
      select.appendChild(option);
    });

    select.addEventListener("change", (event) => {
      state.mapping[field.key] = event.target.value;
    });

    wrapper.appendChild(label);
    wrapper.appendChild(select);
    els.mappingGrid.appendChild(wrapper);
  });
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

  return state.rawRows.map((row) => ({
    campaign: String(row[mapped.campaign] || "미분류").trim(),
    impressions: toNumber(row[mapped.impressions]),
    clicks: toNumber(row[mapped.clicks]),
    orders: toNumber(row[ordersCol]),
    cost: toNumber(row[mapped.cost]),
    revenue: toNumber(row[revenueCol]),
    sales: toNumber(row[salesCol]),
    area: normalizeArea(row[mapped.area]),
    keyword: String(row[mapped.keyword] || "").trim(),
    optionId: String(row[mapped.optionId] || "").trim(),
    optionName: String(row[mapped.optionName] || "").trim(),
    saleDate: String(row[mapped.saleDate] || "").trim(),
  }));
};

const buildCampaigns = () => {
  const map = new Map();

  state.normalized.forEach((row) => {
    const key = row.campaign || "미분류";
    if (!map.has(key)) {
      map.set(key, {
        name: key,
        rows: [],
        totals: { impressions: 0, clicks: 0, orders: 0, cost: 0, revenue: 0, sales: 0 },
      });
    }
    const campaign = map.get(key);
    campaign.rows.push(row);
    campaign.totals.impressions += row.impressions;
    campaign.totals.clicks += row.clicks;
    campaign.totals.orders += row.orders;
    campaign.totals.cost += row.cost;
    campaign.totals.revenue += row.revenue;
    campaign.totals.sales += row.sales;
  });

  state.campaigns = Array.from(map.values()).map((campaign) => ({
    ...campaign,
    metrics: buildMetric(campaign.totals),
  }));

  renderCampaignList();
};

const renderCampaignList = () => {
  const query = els.campaignSearch.value.trim().toLowerCase();
  els.campaignList.innerHTML = "";

  const filtered = state.campaigns.filter((campaign) =>
    campaign.name.toLowerCase().includes(query)
  );

  filtered.forEach((campaign) => {
    const item = document.createElement("div");
    item.className = "list-item";
    if (state.selectedCampaign && state.selectedCampaign.name === campaign.name) {
      item.classList.add("active");
    }
    const title = document.createElement("h3");
    title.textContent = campaign.name;
    const meta = document.createElement("div");
    meta.className = "meta";
    meta.textContent = `ROAS ${fmtNumber(campaign.metrics.roas, 2)}% · 광고비 ${fmtNumber(
      campaign.totals.cost
    )}원`;

    item.appendChild(title);
    item.appendChild(meta);
    item.addEventListener("click", () => selectCampaign(campaign.name));
    els.campaignList.appendChild(item);
  });
};

const renderSummary = (campaign) => {
  els.campaignTitle.textContent = campaign.name;
  els.campaignMeta.textContent = `행 수 ${campaign.rows.length} · 판매량 ${fmtNumber(
    campaign.totals.sales
  )}`;

  const metrics = campaign.metrics;
  const cards = [
    { label: "ROAS", value: fmtPercent(metrics.roas) },
    { label: "CPC", value: `${fmtNumber(metrics.cpc, 2)}원` },
    { label: "클릭률", value: fmtPercent(metrics.ctr) },
    { label: "전환율", value: fmtPercent(metrics.cvr) },
    { label: "광고비", value: `${fmtNumber(campaign.totals.cost)}원` },
    { label: "광고매출", value: `${fmtNumber(campaign.totals.revenue)}원` },
    { label: "전환당비용", value: `${fmtNumber(metrics.cpa, 2)}원` },
    { label: "객단가", value: `${fmtNumber(metrics.aov, 2)}원` },
  ];

  els.summaryCards.innerHTML = cards
    .map(
      (card) => `
        <div class="card">
          <div class="label">${card.label}</div>
          <div class="value">${card.value}</div>
        </div>`
    )
    .join("");
};

const renderBaseStats = (campaign) => {
  const areaMap = new Map();
  campaign.rows.forEach((row) => {
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

  const rows = Array.from(areaMap.values()).map((item) => {
    const metric = buildMetric({
      impressions: item.totals.impressions,
      clicks: item.totals.clicks,
      orders: item.totals.orders,
      cost: item.totals.cost,
      revenue: item.totals.revenue,
    });
    return [
      item.area,
      fmtNumber(item.totals.impressions),
      fmtNumber(item.totals.clicks),
      fmtNumber(item.totals.orders),
      fmtPercent(metric.ctr),
      fmtPercent(metric.cvr),
      fmtNumber(metric.cpm, 2),
      fmtNumber(metric.cpc, 2),
      fmtNumber(item.totals.cost),
      fmtNumber(item.totals.revenue),
      fmtPercent(metric.roas),
      fmtNumber(metric.cpa, 2),
      fmtNumber(metric.aov, 2),
    ];
  });

  els.tabBase.innerHTML = renderTable(
    [
      "노출영역",
      "노출수",
      "클릭",
      "주문",
      "클릭률",
      "전환율",
      "CPM",
      "CPC",
      "광고비(VAT포함)",
      "광고매출",
      "ROAS",
      "전환당비용",
      "객단가",
    ],
    rows
  );
};

const renderSoldOptions = (campaign) => {
  const filtered = campaign.rows.filter((row) => row.sales > 0);
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

  const rows = Array.from(map.values()).map((item) => [
    item.saleDate,
    item.optionName,
    item.keyword,
    fmtNumber(item.revenue),
  ]);

  els.tabSold.innerHTML = renderTable(
    ["판매일", "상품명", "판매된 키워드", "광고매출"],
    rows
  );
};

const renderOptions = (campaign) => {
  const map = new Map();
  campaign.rows.forEach((row) => {
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

  const rows = Array.from(map.values()).map((item) => {
    const metric = buildMetric({
      impressions: item.totals.impressions,
      clicks: item.totals.clicks,
      orders: item.totals.orders,
      cost: item.totals.cost,
      revenue: item.totals.revenue,
    });

    return [
      item.optionId,
      item.optionName,
      fmtNumber(item.totals.sales),
      fmtNumber(item.totals.cost),
      fmtNumber(item.totals.revenue),
      fmtPercent(metric.roas),
      fmtNumber(item.totals.impressions),
      fmtNumber(item.totals.clicks),
      fmtPercent(metric.ctr),
      fmtPercent(metric.cvr),
    ];
  });

  els.tabOptions.innerHTML = renderTable(
    [
      "옵션ID",
      "상품명",
      "판매량",
      "광고비(VAT포함)",
      "광고매출",
      "ROAS",
      "노출",
      "클릭",
      "클릭률",
      "전환율",
    ],
    rows
  );
};

const renderKeywords = (campaign) => {
  const map = new Map();
  campaign.rows.forEach((row) => {
    if (!row.keyword) return;
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

  const excluded = new Set(loadExcluded(campaign.name));

  const rows = Array.from(map.values()).map((item) => {
    const metric = buildMetric({
      impressions: item.totals.impressions,
      clicks: item.totals.clicks,
      orders: item.totals.orders,
      cost: item.totals.cost,
      revenue: item.totals.revenue,
    });

    return {
      keyword: item.keyword,
      cells: [
        item.keyword,
        fmtNumber(item.totals.impressions),
        fmtNumber(item.totals.clicks),
        fmtPercent(metric.ctr),
        fmtNumber(item.totals.orders),
        fmtPercent(metric.cvr),
        fmtNumber(metric.cpc, 2),
        fmtNumber(item.totals.cost),
        fmtNumber(item.totals.revenue),
        fmtPercent(metric.roas),
      ],
      excluded: excluded.has(item.keyword),
    };
  });

  rows.sort((a, b) => a.keyword.localeCompare(b.keyword));

  const tableRows = rows
    .map(
      (row) => `
      <tr data-keyword="${row.keyword}">
        <td><input type="checkbox" ${row.excluded ? "checked" : ""} /></td>
        ${row.cells.map((cell) => `<td>${cell}</td>`).join("")}
      </tr>`
    )
    .join("");

  els.tabKeywords.innerHTML = `
    <div class="table-actions">
      <button id="saveExcluded" class="ghost">선택 키워드 제외 저장</button>
      <span class="small">체크된 키워드가 제외 키워드로 저장됩니다.</span>
    </div>
    <div class="table-wrap">
      <table class="table">
        <thead>
          <tr>
            <th>제외</th>
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
          </tr>
        </thead>
        <tbody>
          ${tableRows || `<tr><td colspan="11">데이터 없음</td></tr>`}
        </tbody>
      </table>
    </div>
  `;

  const saveBtn = document.getElementById("saveExcluded");
  saveBtn.addEventListener("click", () => {
    const checked = Array.from(els.tabKeywords.querySelectorAll("tbody tr"))
      .filter((row) => row.querySelector("input")?.checked)
      .map((row) => row.dataset.keyword)
      .filter(Boolean);

    saveExcluded(campaign.name, checked);
    renderExcluded(campaign);
  });
};

const renderExcluded = (campaign) => {
  const excluded = loadExcluded(campaign.name);

  const rows = excluded.map(
    (keyword) => `
      <tr data-keyword="${keyword}">
        <td>${keyword}</td>
        <td><button class="ghost remove-excluded">삭제</button></td>
      </tr>`
  );

  els.tabExcluded.innerHTML = `
    <div class="table-wrap">
      <table class="table">
        <thead>
          <tr>
            <th>키워드</th>
            <th>작업</th>
          </tr>
        </thead>
        <tbody>
          ${rows.join("") || `<tr><td colspan="2">제외 키워드 없음</td></tr>`}
        </tbody>
      </table>
    </div>
  `;

  els.tabExcluded.querySelectorAll(".remove-excluded").forEach((btn) => {
    btn.addEventListener("click", (event) => {
      const keyword = event.target.closest("tr")?.dataset.keyword;
      if (!keyword) return;
      const next = excluded.filter((item) => item !== keyword);
      saveExcluded(campaign.name, next);
      renderExcluded(campaign);
      renderKeywords(campaign);
    });
  });
};

const renderTable = (headers, rows) => {
  const bodyRows = rows.length
    ? rows.map((cells) => `<tr>${cells.map((cell) => `<td>${cell}</td>`).join("")}</tr>`).join("")
    : `<tr><td colspan="${headers.length}">데이터 없음</td></tr>`;

  return `
    <div class="table-wrap">
      <table class="table">
        <thead>
          <tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr>
        </thead>
        <tbody>${bodyRows}</tbody>
      </table>
    </div>
  `;
};

const selectCampaign = (name) => {
  const campaign = state.campaigns.find((item) => item.name === name);
  if (!campaign) return;
  state.selectedCampaign = campaign;

  els.emptyState.classList.add("hidden");
  els.detailWrap.classList.remove("hidden");

  renderCampaignList();
  renderSummary(campaign);
  renderBaseStats(campaign);
  renderSoldOptions(campaign);
  renderOptions(campaign);
  renderKeywords(campaign);
  renderExcluded(campaign);
};

const handleApplyMapping = () => {
  const requiredKeys = ["campaign", "impressions", "clicks", "cost"];
  const windowOrders = state.window === "14d" ? "orders14" : "orders1";
  const windowRevenue = state.window === "14d" ? "revenue14" : "revenue1";

  const missing = requiredKeys
    .filter((key) => !state.mapping[key])
    .map((key) => REQUIRED_FIELDS.find((field) => field.key === key)?.label || key);

  if (!state.mapping[windowOrders] && !state.mapping.orders) {
    missing.push(state.window === "14d" ? "총 주문수(14일)" : "총 주문수(1일)");
  }
  if (!state.mapping[windowRevenue] && !state.mapping.revenue) {
    missing.push(state.window === "14d" ? "총 전환매출액(14일)" : "총 전환매출액(1일)");
  }

  if (missing.length) {
    alert(`필수 컬럼이 누락되었습니다: ${missing.join(", ")}`);
    return;
  }

  state.normalized = normalizeRows();
  buildCampaigns();

  if (state.campaigns.length) {
    selectCampaign(state.campaigns[0].name);
  }
};

const handleTabSwitch = (event) => {
  const btn = event.target.closest("button");
  if (!btn) return;
  const tab = btn.dataset.tab;
  if (!tab) return;

  els.tabs.querySelectorAll("button").forEach((button) => {
    button.classList.toggle("active", button === btn);
  });

  const panels = {
    base: els.tabBase,
    sold: els.tabSold,
    options: els.tabOptions,
    keywords: els.tabKeywords,
    excluded: els.tabExcluded,
  };

  Object.entries(panels).forEach(([key, panel]) => {
    panel.classList.toggle("hidden", key !== tab);
  });
};

const handleWindowChange = () => {
  state.window = els.windowSelect.value;
  if (!state.rawRows.length) return;
  state.normalized = normalizeRows();
  buildCampaigns();
  if (state.selectedCampaign) {
    selectCampaign(state.selectedCampaign.name);
  }
};

const handleSave = () => {
  if (!state.normalized.length) {
    alert("저장할 분석 데이터가 없습니다.");
    return;
  }

  const excludedByCampaign = {};
  state.campaigns.forEach((campaign) => {
    excludedByCampaign[campaign.name] = loadExcluded(campaign.name);
  });

  const payload = {
    version: 1,
    createdAt: new Date().toISOString(),
    mapping: state.mapping,
    rows: state.normalized,
    excludedByCampaign,
  };

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `coupang-report-${Date.now()}.json`;
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
  if (payload.excludedByCampaign && typeof payload.excludedByCampaign === "object") {
    Object.entries(payload.excludedByCampaign).forEach(([name, list]) => {
      if (Array.isArray(list)) saveExcluded(name, list);
    });
  }
  buildCampaigns();

  if (state.campaigns.length) {
    selectCampaign(state.campaigns[0].name);
  }
};

els.fileInput.addEventListener("change", (event) => {
  const file = event.target.files[0];
  handleFile(file);
});

els.applyMapping.addEventListener("click", handleApplyMapping);
els.campaignSearch.addEventListener("input", renderCampaignList);
els.tabs.addEventListener("click", handleTabSwitch);
els.saveBtn.addEventListener("click", handleSave);
els.loadJsonInput.addEventListener("change", (event) => handleLoad(event.target.files[0]));
els.windowSelect.addEventListener("change", handleWindowChange);
