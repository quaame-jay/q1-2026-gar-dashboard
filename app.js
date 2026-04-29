const DATA = window.GAR_DATA;
const MONTHS = DATA.metadata.monthOrder;
const COLORS = ["#2f75b5", "#70ad47", "#ffc000", "#ed7d31", "#c00000", "#8064a2", "#5b9bd5", "#a5a5a5"];
const STATE_COLORS = {
  Received: "#70ad47",
  "Partially Received": "#ffc000",
  "In Transit": "#5b9bd5",
  "Partially Shipped": "#ed7d31",
  "Vendor Rejected": "#c00000",
  Closed: "#a5a5a5",
  Open: "#7f7f7f",
};

const filterIds = ["month", "region", "district", "type", "supplier", "state", "consumed"];
const filters = {
  month: document.getElementById("monthFilter"),
  region: document.getElementById("regionFilter"),
  district: document.getElementById("districtFilter"),
  type: document.getElementById("typeFilter"),
  supplier: document.getElementById("supplierFilter"),
  state: document.getElementById("stateFilter"),
  consumed: document.getElementById("consumedFilter"),
  search: document.getElementById("siteSearch"),
};

const fmt = new Intl.NumberFormat("en-US");
const one = new Intl.NumberFormat("en-US", { maximumFractionDigits: 1 });
const pctFmt = new Intl.NumberFormat("en-US", { style: "percent", maximumFractionDigits: 1 });

function unique(values) {
  return [...new Set(values.filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b)));
}

function avg(values) {
  const nums = values.filter((v) => Number.isFinite(v));
  return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
}

function ratio(n, d) {
  return d ? n / d : 0;
}

function countBy(rows, keyFn) {
  const map = new Map();
  rows.forEach((row) => {
    const key = keyFn(row) || "Unknown";
    map.set(key, (map.get(key) || 0) + 1);
  });
  return [...map.entries()].map(([label, value]) => ({ label, value }));
}

function topN(items, n, sortKey = "value") {
  return [...items].sort((a, b) => b[sortKey] - a[sortKey]).slice(0, n);
}

function selected() {
  return {
    month: filters.month.value,
    region: filters.region.value,
    district: filters.district.value,
    type: filters.type.value,
    supplier: filters.supplier.value,
    state: filters.state.value,
    consumed: filters.consumed.value,
    search: filters.search.value.trim().toLowerCase(),
  };
}

function filterData() {
  const f = selected();
  const orders = DATA.orders.filter((o) => {
    if (f.month !== "All" && o.month !== f.month) return false;
    if (f.region !== "All" && o.region !== f.region) return false;
    if (f.district !== "All" && o.district !== f.district) return false;
    if (f.type !== "All" && o.type !== f.type) return false;
    if (f.supplier !== "All" && o.supplier !== f.supplier) return false;
    if (f.state !== "All" && o.state !== f.state) return false;
    if (f.search && !`${o.customer} ${o.site} ${o.orderNumber}`.toLowerCase().includes(f.search)) return false;
    return true;
  });

  const consumption = DATA.consumption.filter((c) => {
    if (f.region !== "All" && c.region !== f.region) return false;
    if (f.district !== "All" && c.district !== f.district) return false;
    if (f.type !== "All" && c.type !== f.type) return false;
    if (f.consumed !== "All" && String(c.consumed) !== f.consumed) return false;
    if (f.search && !c.facility.toLowerCase().includes(f.search)) return false;
    return true;
  });

  return { orders, consumption };
}

function populateSelect(select, values, allLabel = "All") {
  const prior = select.value || allLabel;
  select.innerHTML = "";
  [allLabel, ...values].forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
  select.value = [...select.options].some((o) => o.value === prior) ? prior : allLabel;
}

function initFilters() {
  populateSelect(filters.month, MONTHS);
  populateSelect(filters.region, unique([...DATA.orders.map((o) => o.region), ...DATA.consumption.map((c) => c.region)]));
  populateSelect(filters.district, unique([...DATA.orders.map((o) => o.district), ...DATA.consumption.map((c) => c.district)]));
  populateSelect(filters.type, unique([...DATA.orders.map((o) => o.type), ...DATA.consumption.map((c) => c.type)]));
  populateSelect(filters.supplier, unique(DATA.orders.map((o) => o.supplier)));
  populateSelect(filters.state, unique(DATA.orders.map((o) => o.state)));
  populateSelect(filters.consumed, ["0", "1", "2", "3"]);
}

function refreshDependentFilters() {
  const f = selected();
  const orderScope = DATA.orders.filter((o) => {
    if (f.month !== "All" && o.month !== f.month) return false;
    if (f.region !== "All" && o.region !== f.region) return false;
    if (f.type !== "All" && o.type !== f.type) return false;
    return true;
  });
  const consScope = DATA.consumption.filter((c) => {
    if (f.region !== "All" && c.region !== f.region) return false;
    if (f.type !== "All" && c.type !== f.type) return false;
    return true;
  });
  populateSelect(filters.district, unique([...orderScope.map((o) => o.district), ...consScope.map((c) => c.district)]));
}

function renderKpis(orders, consumption) {
  const states = new Map(countBy(orders, (o) => o.state).map((x) => [x.label, x.value]));
  const totalOrders = orders.length;
  const totalFacilities = consumption.length;
  const received = states.get("Received") || 0;
  const consumed3 = consumption.filter((c) => c.consumed === 3).length;
  const consumedCounts = new Map(countBy(consumption, (c) => String(c.consumed)).map((x) => [x.label, x.value]));
  const uniqueSites = new Set(orders.map((o) => o.site)).size;
  const kpis = [
    ["Total Orders", fmt.format(totalOrders), "Filtered Q1 order records"],
    ["Total Facilities", fmt.format(totalFacilities), "Filtered consumption facilities"],
    ["Onboarded Sites / Customers", fmt.format(uniqueSites), "Unique sites in filtered orders"],
    ["Orders Received", fmt.format(received), `${pctFmt.format(ratio(received, totalOrders))} of filtered orders`],
    ["Partially Received", fmt.format(states.get("Partially Received") || 0), "Needs receipt follow-up"],
    ["In Transit", fmt.format(states.get("In Transit") || 0), "Still moving through fulfillment"],
    ["Receipt Completion Rate", pctFmt.format(ratio(received, totalOrders)), "Received / total orders"],
    ["Avg Submission TAT", one.format(avg(orders.map((o) => o.submissionTat))), "Days"],
    ["Avg Approval TAT", one.format(avg(orders.map((o) => o.approvalTat))), "Days"],
    ["Avg Promising TAT", one.format(avg(orders.map((o) => o.promisingTat))), "Days"],
    ["Avg WH Processing TAT", one.format(avg(orders.map((o) => o.whTat))), "Days"],
    ["Avg Delivery TAT", one.format(avg(orders.map((o) => o.deliveryTat))), "Days"],
    ["Facilities Consuming 3 Months", fmt.format(consumed3), `${pctFmt.format(ratio(consumed3, totalFacilities))} coverage`],
    ["Consumed 0 / 1 / 2 Months", `${fmt.format(consumedCounts.get("0") || 0)} / ${fmt.format(consumedCounts.get("1") || 0)} / ${fmt.format(consumedCounts.get("2") || 0)}`, "Facilities needing attention"],
    ["Consumption Coverage Rate", pctFmt.format(ratio(consumed3, totalFacilities)), "Consumed all three months"],
  ];

  document.getElementById("kpiGrid").innerHTML = kpis.map(([label, value, detail]) => `
    <article class="kpi">
      <div class="label">${label}</div>
      <div class="value">${value}</div>
      <div class="detail">${detail}</div>
    </article>
  `).join("");
}

function renderCallouts(orders, consumption) {
  const states = new Map(countBy(orders, (o) => o.state).map((x) => [x.label, x.value]));
  const totalOrders = orders.length;
  const received = states.get("Received") || 0;
  const openRisk = ["Partially Received", "In Transit", "Partially Shipped"].reduce((sum, key) => sum + (states.get(key) || 0), 0);
  const regionCoverage = countBy(consumption, (c) => c.region).map((r) => {
    const rows = consumption.filter((c) => c.region === r.label);
    return { label: r.label, total: rows.length, consumed3: rows.filter((c) => c.consumed === 3).length, coverage: ratio(rows.filter((c) => c.consumed === 3).length, rows.length) };
  }).sort((a, b) => a.coverage - b.coverage);
  const deliveryByDistrict = groupDeliveryTat(orders);
  const slowest = deliveryByDistrict[0];
  const notes = [
    { type: "good", text: `Receipt completion is ${pctFmt.format(ratio(received, totalOrders))}: ${fmt.format(received)} of ${fmt.format(totalOrders)} filtered orders are fully received.` },
    { type: "warn", text: `${fmt.format(openRisk)} orders are partially received, in transit, or partially shipped and should stay on the follow-up list.` },
    { type: "risk", text: regionCoverage[0] ? `Lowest consumption coverage: ${regionCoverage[0].label} at ${pctFmt.format(regionCoverage[0].coverage)}.` : "No consumption rows match the current filter." },
    { type: "warn", text: slowest ? `Longest average delivery TAT: ${slowest.label} at ${one.format(slowest.value)} days.` : "No delivery TAT values match the current filter." },
    { type: "info", text: `${fmt.format(DATA.metadata.excludedNonQ1)} non-Q1 order records are flagged outside this dashboard's Jan-Mar 2026 scope.` },
  ];
  document.getElementById("callouts").innerHTML = notes.map((note) => `<div class="callout ${note.type}">${note.text}</div>`).join("");
}

function renderBarChart(el, rows, options = {}) {
  const max = Math.max(...rows.map((r) => r.value), 0);
  const heightClass = options.tall ? " tall" : "";
  if (!rows.length || max === 0) {
    el.innerHTML = `<div class="empty-state${heightClass}">No matching records</div>`;
    return;
  }
  el.innerHTML = rows.map((row, index) => {
    const width = Math.max(1, (row.value / max) * 100);
    const color = row.color || COLORS[index % COLORS.length];
    const label = row.label;
    const value = options.percent ? pctFmt.format(row.value) : options.decimal ? one.format(row.value) : fmt.format(row.value);
    return `
      <div class="bar-row">
        <div class="bar-label" title="${label}">${label}</div>
        <div class="bar-track"><div class="bar-fill" style="width:${width}%;background:${color}"></div></div>
        <div class="bar-value">${value}</div>
      </div>
    `;
  }).join("");
}

function renderLineChart(el, rows) {
  const w = 620;
  const h = 260;
  const pad = 38;
  const max = Math.max(...rows.map((r) => r.value), 1);
  const points = rows.map((r, i) => {
    const x = pad + (i * (w - pad * 2)) / Math.max(rows.length - 1, 1);
    const y = h - pad - (r.value / max) * (h - pad * 2);
    return { ...r, x, y };
  });
  const path = points.map((p, i) => `${i ? "L" : "M"} ${p.x} ${p.y}`).join(" ");
  el.innerHTML = `
    <svg class="line-chart" viewBox="0 0 ${w} ${h}" role="img" aria-label="Monthly order volume line chart">
      <line x1="${pad}" y1="${h - pad}" x2="${w - pad}" y2="${h - pad}" stroke="#c9d4df" />
      <line x1="${pad}" y1="${pad}" x2="${pad}" y2="${h - pad}" stroke="#c9d4df" />
      <path d="${path}" fill="none" stroke="#2f75b5" stroke-width="4" stroke-linecap="round" />
      ${points.map((p) => `
        <circle cx="${p.x}" cy="${p.y}" r="6" fill="#17365d" />
        <text x="${p.x}" y="${p.y - 13}" text-anchor="middle" font-size="13" font-weight="700" fill="#344054">${fmt.format(p.value)}</text>
        <text x="${p.x}" y="${h - 11}" text-anchor="middle" font-size="13" font-weight="700" fill="#667085">${p.label}</text>
      `).join("")}
    </svg>
  `;
}

function renderDonut(el, rows) {
  const total = rows.reduce((sum, row) => sum + row.value, 0);
  if (!total) {
    el.innerHTML = `<div class="empty-state">No matching records</div>`;
    return;
  }
  let offset = 25;
  const radius = 70;
  const circumference = 2 * Math.PI * radius;
  const segments = rows.map((row, index) => {
    const fraction = row.value / total;
    const dash = fraction * circumference;
    const color = row.color || STATE_COLORS[row.label] || COLORS[index % COLORS.length];
    const segment = `<circle r="${radius}" cx="110" cy="110" fill="transparent" stroke="${color}" stroke-width="28" stroke-dasharray="${dash} ${circumference - dash}" stroke-dashoffset="${offset}" transform="rotate(-90 110 110)" />`;
    offset -= dash;
    return segment;
  }).join("");
  const legend = rows.map((row, index) => {
    const color = row.color || STATE_COLORS[row.label] || COLORS[index % COLORS.length];
    return `
      <div class="legend-item">
        <span class="legend-swatch" style="background:${color}"></span>
        <span class="legend-label">${row.label}</span>
        <span class="legend-value">${fmt.format(row.value)} (${pctFmt.format(row.value / total)})</span>
      </div>
    `;
  }).join("");
  el.innerHTML = `
    <div class="donut-wrap">
      <svg class="donut" viewBox="0 0 220 220" role="img" aria-label="Order state donut chart">
        <circle r="${radius}" cx="110" cy="110" fill="transparent" stroke="#edf2f7" stroke-width="28" />
        ${segments}
        <text x="110" y="105" text-anchor="middle" font-size="24" font-weight="800" fill="#17365d">${fmt.format(total)}</text>
        <text x="110" y="128" text-anchor="middle" font-size="12" font-weight="700" fill="#667085">orders</text>
      </svg>
      <div>${legend}</div>
    </div>
  `;
}

function groupDeliveryTat(orders) {
  const map = new Map();
  orders.forEach((o) => {
    if (!Number.isFinite(o.deliveryTat)) return;
    if (!map.has(o.district)) map.set(o.district, []);
    map.get(o.district).push(o.deliveryTat);
  });
  return [...map.entries()]
    .filter(([, values]) => values.length >= 3)
    .map(([label, values]) => ({ label, value: avg(values), records: values.length }))
    .sort((a, b) => b.value - a.value);
}

function buildMatrix(orders, consumption) {
  const groups = new Map();
  consumption.forEach((c) => {
    const key = `${c.region}|${c.district}|${c.type}`;
    if (!groups.has(key)) groups.set(key, { region: c.region, district: c.district, type: c.type, facilities: [], orders: [] });
    groups.get(key).facilities.push(c);
  });
  orders.forEach((o) => {
    const key = `${o.region}|${o.district}|${o.type}`;
    if (!groups.has(key)) groups.set(key, { region: o.region, district: o.district, type: o.type, facilities: [], orders: [] });
    groups.get(key).orders.push(o);
  });
  return [...groups.values()].map((g) => {
    const consumed3 = g.facilities.filter((f) => f.consumed === 3).length;
    const received = g.orders.filter((o) => o.state === "Received").length;
    return {
      ...g,
      totalFacilities: g.facilities.length,
      consumed3,
      coverage: ratio(consumed3, g.facilities.length),
      totalOrders: g.orders.length,
      receiptCompletion: ratio(received, g.orders.length),
      avgDeliveryTat: avg(g.orders.map((o) => o.deliveryTat)),
    };
  }).sort((a, b) => b.avgDeliveryTat - a.avgDeliveryTat);
}

function renderTable(table, headers, rows, rowFn, limit = 50) {
  const visible = rows.slice(0, limit);
  table.innerHTML = `
    <thead><tr>${headers.map((h) => `<th>${h}</th>`).join("")}</tr></thead>
    <tbody>${visible.map(rowFn).join("")}</tbody>
  `;
}

function statusClass(value, goodHigh = true) {
  if (goodHigh) {
    if (value >= 0.8) return "status-good";
    if (value >= 0.5) return "status-warn";
    return "status-risk";
  }
  if (value <= 3) return "status-good";
  if (value <= 7) return "status-warn";
  return "status-risk";
}

function renderAll() {
  refreshDependentFilters();
  const { orders, consumption } = filterData();
  renderKpis(orders, consumption);
  renderCallouts(orders, consumption);
  document.getElementById("ordersScope").textContent = `${fmt.format(orders.length)} filtered orders`;

  renderLineChart(document.getElementById("monthlyOrdersChart"), MONTHS.map((m) => ({ label: m, value: orders.filter((o) => o.month === m).length })));
  renderDonut(document.getElementById("stateChart"), topN(countBy(orders, (o) => o.state).map((r) => ({ ...r, color: STATE_COLORS[r.label] })), 8));
  renderBarChart(document.getElementById("regionOrdersChart"), topN(countBy(orders, (o) => o.region), 8));
  renderBarChart(document.getElementById("typeOrdersChart"), topN(countBy(orders, (o) => o.type), 10));
  renderBarChart(document.getElementById("districtVolumeChart"), topN(countBy(orders, (o) => o.district), 10), { tall: true });

  const tatRows = [
    { label: "Submission", value: avg(orders.map((o) => o.submissionTat)), color: "#2f75b5" },
    { label: "Approval", value: avg(orders.map((o) => o.approvalTat)), color: "#70ad47" },
    { label: "Promising", value: avg(orders.map((o) => o.promisingTat)), color: "#ffc000" },
    { label: "WH Processing", value: avg(orders.map((o) => o.whTat)), color: "#ed7d31" },
    { label: "Delivery", value: avg(orders.map((o) => o.deliveryTat)), color: "#c00000" },
  ];
  renderBarChart(document.getElementById("tatStageChart"), tatRows, { decimal: true });
  renderBarChart(document.getElementById("deliveryTatChart"), topN(groupDeliveryTat(orders), 10, "value"), { decimal: true, tall: true });

  const regionCoverage = countBy(consumption, (c) => c.region).map((r) => {
    const rows = consumption.filter((c) => c.region === r.label);
    return { label: r.label, value: ratio(rows.filter((c) => c.consumed === 3).length, rows.length) };
  }).sort((a, b) => b.value - a.value);
  renderBarChart(document.getElementById("coverageRegionChart"), regionCoverage, { percent: true });
  renderBarChart(document.getElementById("consumedCountChart"), [0, 1, 2, 3].map((n) => ({ label: `${n} month${n === 1 ? "" : "s"}`, value: consumption.filter((c) => c.consumed === n).length })));

  const matrix = buildMatrix(orders, consumption);
  document.getElementById("matrixCount").textContent = `Showing ${fmt.format(Math.min(matrix.length, 60))} of ${fmt.format(matrix.length)} rows`;
  renderTable(document.getElementById("matrixTable"), ["Region", "District", "Facility type", "Facilities", "Consumed 3", "Coverage", "Orders", "Receipt completion", "Avg delivery TAT"], matrix, (m) => `
    <tr>
      <td>${m.region}</td><td>${m.district}</td><td>${m.type}</td>
      <td>${fmt.format(m.totalFacilities)}</td><td>${fmt.format(m.consumed3)}</td>
      <td class="${statusClass(m.coverage)}">${pctFmt.format(m.coverage)}</td>
      <td>${fmt.format(m.totalOrders)}</td>
      <td class="${statusClass(m.receiptCompletion)}">${pctFmt.format(m.receiptCompletion)}</td>
      <td class="${statusClass(m.avgDeliveryTat, false)}">${one.format(m.avgDeliveryTat)}</td>
    </tr>
  `, 60);

  const districtCoverage = countBy(consumption, (c) => `${c.region}|${c.district}`).map((r) => {
    const [region, district] = r.label.split("|");
    const rows = consumption.filter((c) => c.region === region && c.district === district);
    const consumed3 = rows.filter((c) => c.consumed === 3).length;
    return { region, district, total: rows.length, consumed3, coverage: ratio(consumed3, rows.length) };
  }).sort((a, b) => a.coverage - b.coverage || b.total - a.total);
  document.getElementById("coverageTableCount").textContent = `Showing ${fmt.format(Math.min(districtCoverage.length, 80))} of ${fmt.format(districtCoverage.length)} districts`;
  renderTable(document.getElementById("coverageTable"), ["Region", "District", "Facilities", "Consumed 3 months", "Coverage"], districtCoverage, (r) => `
    <tr><td>${r.region}</td><td>${r.district}</td><td>${fmt.format(r.total)}</td><td>${fmt.format(r.consumed3)}</td><td class="${statusClass(r.coverage)}">${pctFmt.format(r.coverage)}</td></tr>
  `, 80);

  const sortedOrders = orders.slice(0, 100);
  document.getElementById("orderTableCount").textContent = `Showing ${fmt.format(sortedOrders.length)} of ${fmt.format(orders.length)} orders`;
  renderTable(document.getElementById("orderTable"), ["Month", "Region", "District", "Site", "Type", "State", "Supplier", "Order #", "Delivery TAT"], sortedOrders, (o) => `
    <tr><td>${o.month}</td><td>${o.region}</td><td>${o.district}</td><td>${o.site}</td><td>${o.type}</td><td>${o.state}</td><td>${o.supplier}</td><td>${o.orderNumber}</td><td class="${statusClass(o.deliveryTat || 0, false)}">${Number.isFinite(o.deliveryTat) ? one.format(o.deliveryTat) : ""}</td></tr>
  `, 100);

  renderQualityNotes();
}

function renderQualityNotes() {
  const notes = [
    ["Source workbook", DATA.metadata.sourceFile],
    ["Q1 scope", "Jan-26, Feb-26, and Mar-26 order records are included. Apr-26 records are excluded from dashboard KPIs."],
    ["Non-Q1 records flagged", fmt.format(DATA.metadata.excludedNonQ1)],
    ["Duplicate order-number rows", fmt.format(DATA.metadata.duplicateOrderRows)],
    ["Blank values", "Blank Region, District, Type, Supplier, Transporter, State, Customer, and Facility values are shown as Unknown."],
    ["Facility names", "Leading/trailing spaces and tab characters were removed."],
    ["On-time departure non-binary values", fmt.format(DATA.metadata.invalidOnTimeDeparture)],
    ["On-time delivery non-binary values", fmt.format(DATA.metadata.invalidOnTimeDelivery)],
    ["On-time calculation approach", "Only valid 0/1 values should be used for on-time rates. Date/text values are treated as source data-quality issues."],
    ["TAT calculation approach", "Averages use numeric TAT values only; blanks or text values are excluded."],
  ];
  document.getElementById("qualityNotes").innerHTML = notes.map(([title, text]) => `
    <div class="quality-item"><strong>${title}</strong><span>${text}</span></div>
  `).join("");
}

function bindEvents() {
  Object.values(filters).forEach((control) => control.addEventListener("input", renderAll));
  document.getElementById("resetFilters").addEventListener("click", () => {
    filterIds.forEach((id) => filters[id].value = "All");
    filters.search.value = "";
    renderAll();
  });
  document.getElementById("printDashboard").addEventListener("click", () => window.print());
  document.querySelectorAll(".tab-button").forEach((button) => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab-button").forEach((b) => b.classList.remove("active"));
      document.querySelectorAll(".page").forEach((p) => p.classList.remove("active"));
      button.classList.add("active");
      document.getElementById(button.dataset.page).classList.add("active");
    });
  });
}

initFilters();
bindEvents();
renderAll();
