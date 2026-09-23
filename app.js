/*
 * Water Meter Readings Tracker -- app logic.
 *
 * Data model: an array of "months", each with its own set of daily rows.
 * Everything is editable in the browser and persisted to localStorage.
 * A comparison view overlays the same metric across every tracked month
 * (by day-of-month) and flags statistical outliers to help spot anomalies.
 */

// -----------------------------------------------------------------------
// Seed data -- the original transcription (matches Water_Meter_Readings.xlsx)
// -----------------------------------------------------------------------
const SEED_ROWS = [
  {day:1,  lakeInput:35244, cityWater:29446, wellMeter:196000, wellFlow:33.9, lakeUsage:4.25, note:""},
  {day:2,  lakeInput:35294, cityWater:29472, wellMeter:48660,  wellFlow:33.9, lakeUsage:4.5,  note:""},
  {day:3,  lakeInput:35328, cityWater:29483, wellMeter:48000,  wellFlow:33.8, lakeUsage:4.75, note:""},
  {day:4,  lakeInput:35347, cityWater:29483, wellMeter:48540,  wellFlow:33.9, lakeUsage:4.5,  note:""},
  {day:5,  lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:6,  lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:7,  lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:8,  lakeInput:35463, cityWater:29500, wellMeter:196000, wellFlow:34.2, lakeUsage:4.5,  note:""},
  {day:9,  lakeInput:35489, cityWater:29501, wellMeter:48700,  wellFlow:34,   lakeUsage:4.5,  note:""},
  {day:10, lakeInput:35526, cityWater:29518, wellMeter:49100,  wellFlow:34,   lakeUsage:4.75, note:""},
  {day:11, lakeInput:35545, cityWater:29518, wellMeter:49350,  wellFlow:null, lakeUsage:4.5,  note:"Well Flow illegible (obscured by pen in photo)"},
  {day:12, lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:13, lakeInput:null,  cityWater:2,     wellMeter:null,   wellFlow:null, lakeUsage:null, note:"Only a partial '2' visible for City water meter -- verify original"},
  {day:14, lakeInput:35638, cityWater:29531, wellMeter:146000, wellFlow:34.2, lakeUsage:null, note:"Lake Usage illegible (obscured by pen in photo)"},
  {day:15, lakeInput:35665, cityWater:29535, wellMeter:52370,  wellFlow:33.8, lakeUsage:5,    note:""},
  {day:16, lakeInput:35686, cityWater:29538, wellMeter:45320,  wellFlow:34.2, lakeUsage:5,    note:""},
  {day:17, lakeInput:35716, cityWater:29541, wellMeter:49000,  wellFlow:34,   lakeUsage:5.75, note:""},
  {day:18, lakeInput:35760, cityWater:29541, wellMeter:48900,  wellFlow:34,   lakeUsage:6,    note:""},
  {day:19, lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:20, lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:21, lakeInput:35854, cityWater:29543, wellMeter:146600, wellFlow:34,   lakeUsage:6.75, note:""},
  {day:22, lakeInput:35858, cityWater:29543, wellMeter:22380,  wellFlow:null, lakeUsage:6.5,  note:"Well Flow shown as dash (—) in original -- not recorded"},
  {day:23, lakeInput:35868, cityWater:29543, wellMeter:28130,  wellFlow:null, lakeUsage:6,    note:'City water shown as ditto mark (") = same as row above; Well Flow shown as dash (—) -- not recorded'},
  {day:24, lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:25, lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:26, lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:27, lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:28, lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:29, lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
  {day:30, lakeInput:null,  cityWater:null,  wellMeter:null,   wellFlow:null, lakeUsage:null, note:""},
];

const STORAGE_KEY = "waterMeterTracker.v2";
const OLD_STORAGE_KEY = "waterMeterTracker.v1"; // pre-multi-month format
const MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const FIELDS = ["lakeInput","cityWater","wellMeter","wellFlow","lakeUsage"];
const FIELD_LABELS = {
  lakeInput: "Lake Input meter reading (gal)",
  cityWater: "City water meter reading (CCF)",
  wellMeter: "Well meter reading (gal)",
  wellFlow: "Well Flow (gal/min)",
  lakeUsage: "Lake Usage (gal)",
};
// Distinct colors for overlaying multiple months on the comparison chart.
const MONTH_COLOR_POOL = ["#1584a8", "#c9822f", "#2a8703", "#7c3aed", "#c92f5e", "#0b3d5c", "#0b7a75", "#b45309"];

let state = loadState();
let charts = {};

// -----------------------------------------------------------------------
// Data helpers
// -----------------------------------------------------------------------
function daysInMonth(monthName, year) {
  const idx = MONTH_NAMES.indexOf(monthName);
  return new Date(year, idx + 1, 0).getDate();
}

function blankRows(count) {
  return Array.from({ length: count }, (_, i) => ({
    day: i + 1, lakeInput: null, cityWater: null, wellMeter: null,
    wellFlow: null, lakeUsage: null, note: "",
  }));
}

function monthId(monthName, year) {
  return `${year}-${String(MONTH_NAMES.indexOf(monthName) + 1).padStart(2, "0")}`;
}

function loadState() {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (raw) {
    try {
      const parsed = JSON.parse(raw);
      if (parsed && Array.isArray(parsed.months) && parsed.months.length) return parsed;
    } catch (e) { /* fall through */ }
  }
  // Migrate the old single-month format if it's all we have.
  const oldRaw = localStorage.getItem(OLD_STORAGE_KEY);
  if (oldRaw) {
    try {
      const old = JSON.parse(oldRaw);
      if (old && old.rows) {
        const id = monthId(old.month, old.year);
        return { months: [{ id, month: old.month, year: old.year, rows: old.rows }], activeMonthId: id };
      }
    } catch (e) { /* fall through */ }
  }
  const id = monthId("April", 2026);
  return {
    months: [{ id, month: "April", year: 2026, rows: JSON.parse(JSON.stringify(SEED_ROWS)) }],
    activeMonthId: id,
  };
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  const status = document.getElementById("saveStatus");
  if (status) status.textContent = "Saved " + new Date().toLocaleTimeString();
}

function activeMonth() {
  return state.months.find(m => m.id === state.activeMonthId) || state.months[0];
}

// -----------------------------------------------------------------------
// Month tabs + editing controls
// -----------------------------------------------------------------------
function renderMonthTabs() {
  const bar = document.getElementById("monthTabs");
  bar.innerHTML = "";
  state.months.forEach(m => {
    const btn = document.createElement("button");
    btn.textContent = `${m.month} ${m.year}`;
    const active = m.id === state.activeMonthId;
    btn.className = "px-3 py-1.5 rounded-full text-sm font-medium border " +
      (active ? "text-white border-transparent" : "bg-white text-gray-700 border-gray-300 hover:bg-gray-100");
    if (active) btn.style.background = "var(--mid)";
    btn.setAttribute("aria-pressed", active ? "true" : "false");
    btn.addEventListener("click", () => {
      state.activeMonthId = m.id;
      saveState();
      renderAll();
    });
    bar.appendChild(btn);
  });

  const addBtn = document.createElement("button");
  addBtn.textContent = "+ Add Month";
  addBtn.className = "px-3 py-1.5 rounded-full text-sm font-medium border border-dashed border-gray-400 text-gray-600 hover:bg-gray-100";
  addBtn.addEventListener("click", addNewMonth);
  bar.appendChild(addBtn);
}

function addNewMonth() {
  const monthName = prompt(
    "Which month do you want to add? (type full name, e.g. May)\n" +
    "Existing months: " + state.months.map(m => `${m.month} ${m.year}`).join(", ")
  );
  if (!monthName) return;
  const matched = MONTH_NAMES.find(m => m.toLowerCase() === monthName.trim().toLowerCase());
  if (!matched) { alert("Didn't recognize that month name -- please use the full name, e.g. 'May'."); return; }
  const yearStr = prompt("Which year?", String(new Date().getFullYear()));
  if (!yearStr || isNaN(Number(yearStr))) { alert("Please enter a valid year."); return; }
  const year = Number(yearStr);
  const id = monthId(matched, year);
  if (state.months.some(m => m.id === id)) { alert(`${matched} ${year} is already being tracked.`); return; }

  state.months.push({ id, month: matched, year, rows: blankRows(daysInMonth(matched, year)) });
  state.activeMonthId = id;
  saveState();
  renderAll();
}

function renderEditingControls() {
  const m = activeMonth();
  const monthSel = document.getElementById("activeMonthSelect");
  monthSel.innerHTML = MONTH_NAMES.map(name => `<option value="${name}" ${name === m.month ? "selected" : ""}>${name}</option>`).join("");
  document.getElementById("activeYearInput").value = m.year;
  document.getElementById("deleteMonthBtn").disabled = state.months.length <= 1;
  document.getElementById("deleteMonthBtn").classList.toggle("opacity-40", state.months.length <= 1);
}

function relabelActiveMonth(newMonthName, newYear) {
  const m = activeMonth();
  const newId = monthId(newMonthName, newYear);
  if (newId !== m.id && state.months.some(mm => mm.id === newId)) {
    alert(`${newMonthName} ${newYear} already exists -- pick a different month/year.`);
    renderEditingControls();
    return;
  }
  const newCount = daysInMonth(newMonthName, newYear);
  if (newCount !== m.rows.length) {
    // Grow or shrink the row set to match the new month's real day count,
    // preserving whatever overlapping day data already exists.
    const preserved = blankRows(newCount);
    m.rows.forEach(r => { if (r.day <= newCount) preserved[r.day - 1] = r; });
    m.rows = preserved;
  }
  m.month = newMonthName;
  m.year = newYear;
  m.id = newId;
  state.activeMonthId = newId;
  saveState();
  renderAll();
}

function deleteActiveMonth() {
  if (state.months.length <= 1) return;
  const m = activeMonth();
  if (!confirm(`Delete all data for ${m.month} ${m.year}? This can't be undone.`)) return;
  state.months = state.months.filter(mm => mm.id !== m.id);
  state.activeMonthId = state.months[0].id;
  saveState();
  renderAll();
}

// -----------------------------------------------------------------------
// Editable table (active month only)
// -----------------------------------------------------------------------
function renderTable() {
  const m = activeMonth();
  const tbody = document.getElementById("tableBody");
  tbody.innerHTML = "";
  m.rows.forEach((row, idx) => {
    const tr = document.createElement("tr");
    tr.className = idx % 2 === 0 ? "bg-white" : "bg-gray-50";
    if (row.note) tr.classList.add("flagged");

    const dateLabel = `${m.month} ${row.day}, ${m.year}`;
    const dateTd = document.createElement("td");
    dateTd.className = "px-3 py-1.5 font-medium whitespace-nowrap";
    dateTd.textContent = dateLabel;
    tr.appendChild(dateTd);

    FIELDS.forEach(field => {
      const td = document.createElement("td");
      td.className = "px-2 py-1";
      const input = document.createElement("input");
      input.type = "number";
      input.step = "any";
      input.value = row[field] === null || row[field] === undefined ? "" : row[field];
      input.className = "w-full text-right rounded border border-gray-300 px-2 py-1";
      input.setAttribute("aria-label", `${FIELD_LABELS[field]} for ${dateLabel}`);
      input.addEventListener("input", () => {
        const v = input.value.trim();
        row[field] = v === "" ? null : Number(v);
        saveState();
        renderCharts();
        renderComparison();
      });
      td.appendChild(input);
      tr.appendChild(td);
    });

    const noteTd = document.createElement("td");
    noteTd.className = "px-3 py-1.5 text-xs text-yellow-900";
    const noteInput = document.createElement("input");
    noteInput.type = "text";
    noteInput.value = row.note || "";
    noteInput.className = "w-full rounded border border-gray-300 px-2 py-1 text-xs";
    noteInput.setAttribute("aria-label", `Notes for ${dateLabel}`);
    noteInput.addEventListener("input", () => {
      row.note = noteInput.value;
      tr.classList.toggle("flagged", !!row.note);
      saveState();
    });
    noteTd.appendChild(noteInput);
    tr.appendChild(noteTd);

    tbody.appendChild(tr);
  });
}

// -----------------------------------------------------------------------
// Per-month charts (overview + 4 detail charts, active month only)
// -----------------------------------------------------------------------
function chartLabels() {
  return activeMonth().rows.map(r => r.day);
}

function makeLineDataset(field, label, color) {
  return {
    label,
    data: activeMonth().rows.map(r => r[field]),
    borderColor: color,
    backgroundColor: color,
    spanGaps: true,
    tension: 0.25,
  };
}

function renderCharts() {
  const labels = chartLabels();
  renderOverviewChart(labels);
  buildOrUpdate("meterChart", labels, [
    makeLineDataset("lakeInput", "Lake Input (gal)", "#1584a8"),
    makeLineDataset("wellMeter", "Well Meter (gal)", "#0b3d5c"),
  ]);
  buildOrUpdate("cityChart", labels, [makeLineDataset("cityWater", "City Water (CCF)", "#2fb4c9")]);
  buildOrUpdate("flowChart", labels, [makeLineDataset("wellFlow", "Well Flow (gal/min)", "#0b7a75")]);
  buildOrUpdate("usageChart", labels, [makeLineDataset("lakeUsage", "Lake Usage (gal)", "#1584a8")]);
}

function renderOverviewChart(labels) {
  const datasets = [
    { ...makeLineDataset("lakeInput", "Lake Input (gal)", "#1584a8"), yAxisID: "y" },
    { ...makeLineDataset("wellMeter", "Well Meter (gal)", "#0b3d5c"), yAxisID: "y" },
    { ...makeLineDataset("cityWater", "City Water (CCF)", "#2fb4c9"), yAxisID: "y" },
    { ...makeLineDataset("wellFlow", "Well Flow (gal/min)", "#c9822f"), yAxisID: "y1", borderDash: [6, 4] },
    { ...makeLineDataset("lakeUsage", "Lake Usage (gal)", "#c92f5e"), yAxisID: "y1", borderDash: [6, 4] },
  ];
  if (charts.overviewChart) {
    charts.overviewChart.data.labels = labels;
    charts.overviewChart.data.datasets = datasets;
    charts.overviewChart.update();
    return;
  }
  const ctx = document.getElementById("overviewChart").getContext("2d");
  charts.overviewChart = new Chart(ctx, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      scales: {
        x: { title: { display: true, text: "Day of month" } },
        y: { type: "linear", position: "left", title: { display: true, text: "Meter readings (gal / CCF)" } },
        y1: { type: "linear", position: "right", title: { display: true, text: "Well Flow (gal/min) / Lake Usage (gal)" }, grid: { drawOnChartArea: false } },
      },
      plugins: { legend: { position: "bottom" } },
    },
  });
}

function buildOrUpdate(canvasId, labels, datasets) {
  if (charts[canvasId]) {
    charts[canvasId].data.labels = labels;
    charts[canvasId].data.datasets = datasets;
    charts[canvasId].update();
    return;
  }
  const ctx = document.getElementById(canvasId).getContext("2d");
  charts[canvasId] = new Chart(ctx, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: { x: { title: { display: true, text: "Day of month" } }, y: { beginAtZero: false } },
      plugins: { legend: { position: "bottom" } },
    },
  });
}

// -----------------------------------------------------------------------
// Cross-month comparison + anomaly detection
// -----------------------------------------------------------------------
function computeAnomalies(field) {
  const maxDays = Math.max(...state.months.map(m => m.rows.length));
  const anomalies = [];
  for (let day = 1; day <= maxDays; day++) {
    const points = state.months
      .map(m => ({ m, val: (m.rows.find(r => r.day === day) || {})[field] }))
      .filter(p => p.val !== null && p.val !== undefined && !isNaN(p.val));
    if (points.length < 2) continue;
    const vals = points.map(p => p.val);
    const mean = vals.reduce((a, b) => a + b, 0) / vals.length;
    const variance = vals.reduce((a, b) => a + (b - mean) ** 2, 0) / vals.length;
    const std = Math.sqrt(variance);
    points.forEach(p => {
      const zScore = std > 0 ? Math.abs(p.val - mean) / std : 0;
      const pctDiff = mean !== 0 ? (Math.abs(p.val - mean) / Math.abs(mean)) * 100 : 0;
      // With exactly 2 data points, z-score is mathematically always 1.0 for
      // both (population std with n=2 equals the deviation itself), so a
      // z-score threshold can never fire. Fall back to percent-diff alone
      // when there are too few months for z-score to be meaningful.
      const isAnomaly = points.length <= 2 ? pctDiff >= 25 : (zScore >= 1.5 && pctDiff >= 15);
      if (isAnomaly) {
        anomalies.push({ monthLabel: `${p.m.month} ${p.m.year}`, monthId: p.m.id, day, value: p.val, mean, pctDiff });
      }
    });
  }
  return anomalies.sort((a, b) => b.pctDiff - a.pctDiff);
}

function populateMetricSelect() {
  const sel = document.getElementById("metricSelect");
  if (sel.options.length) return; // only needs building once
  sel.innerHTML = FIELDS.map(f => `<option value="${f}">${FIELD_LABELS[f]}</option>`).join("");
}

function renderComparison() {
  populateMetricSelect();
  const field = document.getElementById("metricSelect").value || FIELDS[0];
  const maxDays = Math.max(...state.months.map(m => m.rows.length));
  const labels = Array.from({ length: maxDays }, (_, i) => i + 1);
  const anomalies = computeAnomalies(field);

  const datasets = state.months.map((m, idx) => {
    const color = MONTH_COLOR_POOL[idx % MONTH_COLOR_POOL.length];
    const data = labels.map(day => {
      const row = m.rows.find(r => r.day === day);
      return row ? row[field] : null;
    });
    const pointBackgroundColor = labels.map(day =>
      anomalies.some(a => a.monthId === m.id && a.day === day) ? "#dc2626" : color
    );
    const pointRadius = labels.map(day =>
      anomalies.some(a => a.monthId === m.id && a.day === day) ? 6 : 3
    );
    return {
      label: `${m.month} ${m.year}`,
      data, borderColor: color, backgroundColor: color,
      pointBackgroundColor, pointRadius, spanGaps: true, tension: 0.25,
    };
  });

  if (charts.compareChart) charts.compareChart.destroy();
  const ctx = document.getElementById("compareChart").getContext("2d");
  charts.compareChart = new Chart(ctx, {
    type: "line",
    data: { labels, datasets },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: { x: { title: { display: true, text: "Day of month" } }, y: { title: { display: true, text: FIELD_LABELS[field] } } },
      plugins: { legend: { position: "bottom" } },
    },
  });

  renderAnomalyList(anomalies, field);
}

function renderAnomalyList(anomalies, field) {
  const box = document.getElementById("anomalyList");
  if (state.months.length < 2) {
    box.innerHTML = `<p class="text-sm text-gray-600">Add at least one more month above to enable anomaly detection.</p>`;
    return;
  }
  if (!anomalies.length) {
    box.innerHTML = `<p class="text-sm text-green-700">No anomalies detected for ${FIELD_LABELS[field]} -- all tracked months look consistent day-to-day.</p>`;
    return;
  }
  const rows = anomalies.slice(0, 20).map(a => `
    <tr class="border-t border-gray-200">
      <td class="px-3 py-1.5">${a.monthLabel}</td>
      <td class="px-3 py-1.5 text-right">Day ${a.day}</td>
      <td class="px-3 py-1.5 text-right">${a.value}</td>
      <td class="px-3 py-1.5 text-right">${a.mean.toFixed(2)}</td>
      <td class="px-3 py-1.5 text-right font-semibold text-red-700">${a.pctDiff.toFixed(0)}%</td>
    </tr>`).join("");
  box.innerHTML = `
    <p class="text-sm text-red-700 font-medium mb-2">${anomalies.length} anomal${anomalies.length === 1 ? "y" : "ies"} found for ${FIELD_LABELS[field]}:</p>
    <div class="overflow-x-auto rounded border border-gray-200">
      <table class="min-w-full text-sm">
        <thead class="bg-gray-100"><tr>
          <th class="px-3 py-1.5 text-left">Month</th>
          <th class="px-3 py-1.5 text-right">Day</th>
          <th class="px-3 py-1.5 text-right">Value</th>
          <th class="px-3 py-1.5 text-right">Avg of other months</th>
          <th class="px-3 py-1.5 text-right">Deviation</th>
        </tr></thead>
        <tbody>${rows}</tbody>
      </table>
    </div>`;
}

// -----------------------------------------------------------------------
// Toolbar actions
// -----------------------------------------------------------------------
function exportMonthCSV(m) {
  const header = ["Date", ...FIELDS.map(f => FIELD_LABELS[f]), "Notes"];
  const lines = [header.join(",")];
  m.rows.forEach(row => {
    const dateLabel = `${m.month} ${row.day}, ${m.year}`;
    const vals = [dateLabel, ...FIELDS.map(f => row[f]), `"${(row.note || "").replace(/"/g, '""')}"`];
    lines.push(vals.map(v => (v === null || v === undefined ? "" : v)).join(","));
  });
  downloadCSV(lines.join("\n"), `water_meter_${m.month}_${m.year}.csv`);
}

function exportAllMonthsCSV() {
  const header = ["Month", "Year", "Day", ...FIELDS.map(f => FIELD_LABELS[f]), "Notes"];
  const lines = [header.join(",")];
  state.months.forEach(m => {
    m.rows.forEach(row => {
      const vals = [m.month, m.year, row.day, ...FIELDS.map(f => row[f]), `"${(row.note || "").replace(/"/g, '""')}"`];
      lines.push(vals.map(v => (v === null || v === undefined ? "" : v)).join(","));
    });
  });
  downloadCSV(lines.join("\n"), `water_meter_all_months.csv`);
}

function downloadCSV(text, filename) {
  const blob = new Blob([text], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
}

function wireToolbar() {
  document.getElementById("resetBtn").addEventListener("click", () => {
    if (!confirm("Reset EVERYTHING back to just the original transcription? This deletes any other months and edits you've added.")) return;
    localStorage.removeItem(STORAGE_KEY);
    localStorage.removeItem(OLD_STORAGE_KEY);
    state = loadState();
    saveState();
    renderAll();
  });
  document.getElementById("exportBtn").addEventListener("click", () => exportMonthCSV(activeMonth()));
  document.getElementById("exportAllBtn").addEventListener("click", exportAllMonthsCSV);

  document.getElementById("activeMonthSelect").addEventListener("change", (e) => {
    relabelActiveMonth(e.target.value, activeMonth().year);
  });
  document.getElementById("activeYearInput").addEventListener("change", (e) => {
    const y = Number(e.target.value);
    if (!y) return;
    relabelActiveMonth(activeMonth().month, y);
  });
  document.getElementById("deleteMonthBtn").addEventListener("click", deleteActiveMonth);
  document.getElementById("metricSelect").addEventListener("change", renderComparison);
}

// -----------------------------------------------------------------------
// Init
// -----------------------------------------------------------------------
function renderAll() {
  renderMonthTabs();
  renderEditingControls();
  renderTable();
  renderCharts();
  renderComparison();
}

wireToolbar();
renderAll();
