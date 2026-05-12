const initialPayload = JSON.parse(document.getElementById("initial-dashboard-data").textContent);
const chartCanvas = document.getElementById("trend-chart");
const chartContext = chartCanvas.getContext("2d");

const elements = {
  connectionText: document.getElementById("connection-text"),
  analyzerOnStatus: document.getElementById("analyzer-on-status"),
  analyzerErrorStatus: document.getElementById("analyzer-error-status"),
  analyzerFailureStatus: document.getElementById("analyzer-failure-status"),
  titrolyzerAppStatus: document.getElementById("titrolyzer-app-status"),
  titrolyzerAppStatusText: document.getElementById("titrolyzer-app-status-text"),
  titrolyzerAppStatusDetail: document.getElementById("titrolyzer-app-status-detail"),
  sequenceStatus: document.getElementById("sequence-status"),
  programAStatus: document.getElementById("program-a-status"),
  programBStatus: document.getElementById("program-b-status"),
  analyzerTimestamp: document.getElementById("analyzer-timestamp"),
  r1Value: document.getElementById("r1-value"),
  r2Value: document.getElementById("r2-value"),
  r1Trend: document.getElementById("r1-trend"),
  r2Trend: document.getElementById("r2-trend"),
  readingAge: document.getElementById("reading-age"),
  historyBody: document.getElementById("history-body"),
  chlorineBody: document.getElementById("chlorine-body"),
  commandLogList: document.getElementById("command-log-list"),
  addReading: document.getElementById("add-reading"),
};

const appState = {
  activeSeries: "both",
  history: initialPayload.history,
  selectedRowId: "",
  editingNote: false,
  hoveringHistory: false,
};

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#39;");
}

function formatTitratorValue(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return "—";
  const text = number.toFixed(6).replace(/0+$/, "").replace(/\.$/, "");
  return text || "0";
}

function setCardAccent(element, accent) {
  const card = element?.closest(".metric-card");
  if (!card) return;
  card.classList.remove("accent-blue", "accent-gold", "accent-green", "accent-red", "accent-slate");
  card.classList.add(`accent-${accent}`);
}

function updateAnalyzerAccents(summary) {
  setCardAccent(
    elements.analyzerOnStatus,
    summary.analyzer_on_status === "On" ? "green" : "slate"
  );
  setCardAccent(
    elements.analyzerErrorStatus,
    summary.analyzer_error_status === "Error" ? "red" : summary.analyzer_error_status === "Clear" ? "green" : "slate"
  );
  setCardAccent(
    elements.analyzerFailureStatus,
    summary.analyzer_failure_status === "Failure" ? "red" : summary.analyzer_failure_status === "Clear" ? "green" : "slate"
  );
}


function selectRow(rowId) {
  appState.selectedRowId = rowId;
  document.querySelectorAll(".history-row").forEach((row) => {
    row.classList.toggle("selected", row.dataset.rowId === rowId);
  });
}

function buildNoteCellContent(td, noteValue, timestampKey) {
  td.textContent = "";
  const hasNote = noteValue && noteValue !== "—";

  if (hasNote) {
    const span = document.createElement("span");
    span.className = "note-text";
    span.textContent = noteValue;
    td.appendChild(span);
  }

  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "note-add-btn";
  btn.textContent = hasNote ? "✎" : "+ Add note";
  btn.addEventListener("click", (e) => {
    e.stopPropagation();
    startInlineNoteEdit(td, timestampKey, noteValue);
  });
  td.appendChild(btn);
}

function startInlineNoteEdit(td, timestampKey, noteValue) {
  if (td.querySelector("input")) return;
  const currentNote = noteValue && noteValue !== "—" ? noteValue : "";

  appState.editingNote = true;

  td.textContent = "";
  const input = document.createElement("input");
  input.type = "text";
  input.value = currentNote;
  input.className = "note-inline-input";
  input.placeholder = "Add a note…";
  input.maxLength = 200;
  td.appendChild(input);
  input.focus();
  input.select();

  let committed = false;

  function commit() {
    if (committed) return;
    committed = true;
    appState.editingNote = false;
    const note = input.value.trim();
    buildNoteCellContent(td, note, timestampKey);
    fetch("/api/titrator-note", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ timestamp: timestampKey, note }),
    })
      .then((r) => (r.ok ? r.json() : Promise.reject()))
      .then((payload) => {
        appState.history = payload.history;
        renderHistory(payload.history);
      })
      .catch(() => buildNoteCellContent(td, currentNote, timestampKey));
  }

  function cancel() {
    if (committed) return;
    committed = true;
    appState.editingNote = false;
    buildNoteCellContent(td, currentNote, timestampKey);
  }

  input.addEventListener("blur", commit);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      input.removeEventListener("blur", commit);
      commit();
    } else if (e.key === "Escape") {
      e.preventDefault();
      input.removeEventListener("blur", commit);
      cancel();
    }
  });
}

function drawChart(rows) {
  const { width, height } = chartCanvas;
  chartContext.clearRect(0, 0, width, height);

  const mode = appState.activeSeries;
  const keys = mode === "both" ? ["r1", "r2"] : [mode];

  const allValues = rows.flatMap((row) => keys.map((k) => Number(row[k]))).filter(Number.isFinite);
  if (!allValues.length) return;

  const low = Math.min(...allValues);
  const high = Math.max(...allValues);
  const spread = Math.max(high - low, Math.abs(high) * 0.05, 0.000001);
  const min = low - spread * 0.1;
  const max = high + spread * 0.1;
  const left = 70;
  const right = width - 20;
  const top = 24;
  const bottom = height - 58;

  chartContext.strokeStyle = "rgba(31, 41, 51, 0.12)";
  chartContext.lineWidth = 1;
  for (let i = 0; i <= 4; i += 1) {
    const y = top + ((bottom - top) / 4) * i;
    chartContext.beginPath();
    chartContext.moveTo(left, y);
    chartContext.lineTo(right, y);
    chartContext.stroke();
    const value = formatTitratorValue(max - ((max - min) / 4) * i);
    chartContext.fillStyle = "#6b7280";
    chartContext.font = "12px Bahnschrift";
    chartContext.textAlign = "right";
    chartContext.fillText(value, left - 4, y + 4);
  }

  const seriesConfig = {
    r1: { color: "#0077b6", dotColor: "rgba(0, 119, 182, 0.22)" },
    r2: { color: "#d97706", dotColor: "rgba(217, 119, 6, 0.22)" },
  };

  keys.forEach((key) => {
    const activeRows = rows.filter((row) => row[key] != null);
    if (!activeRows.length) return;
    const { color, dotColor } = seriesConfig[key];

    const toPoint = (value, index) => ({
      x: left + ((right - left) / Math.max(activeRows.length - 1, 1)) * index,
      y: bottom - ((value - min) / (max - min)) * (bottom - top),
    });

    chartContext.beginPath();
    activeRows.forEach((row, index) => {
      const point = toPoint(Number(row[key]), index);
      if (index === 0) chartContext.moveTo(point.x, point.y);
      else chartContext.lineTo(point.x, point.y);
    });
    chartContext.strokeStyle = color;
    chartContext.lineWidth = 4;
    chartContext.stroke();

    activeRows.forEach((row, index) => {
      const point = toPoint(Number(row[key]), index);
      chartContext.beginPath();
      chartContext.arc(point.x, point.y, 4.5, 0, Math.PI * 2);
      chartContext.fillStyle = dotColor;
      chartContext.fill();
    });

    if (key === keys[0]) {
      const maxLabels = Math.max(2, Math.floor((right - left) / 64));
      const step = Math.max(1, Math.ceil(activeRows.length / maxLabels));
      chartContext.fillStyle = "#6b7280";
      chartContext.font = "13px Bahnschrift";
      chartContext.textAlign = "center";
      activeRows.forEach((row, index) => {
        if (index % step === 0) {
          const point = toPoint(Number(row[key]), index);
          chartContext.fillText(row.chart_time, point.x, bottom + 16);
        }
      });
    }
  });

  chartContext.save();
  chartContext.fillStyle = "#94a3b8";
  chartContext.font = "11px Bahnschrift";
  chartContext.textAlign = "center";
  chartContext.translate(8, (top + bottom) / 2);
  chartContext.rotate(-Math.PI / 2);
  chartContext.fillText("R1 / R2 Value", 0, 0);
  chartContext.restore();

  chartContext.fillStyle = "#94a3b8";
  chartContext.font = "11px Bahnschrift";
  chartContext.textAlign = "center";
  chartContext.fillText("Time", (left + right) / 2, height - 10);
}

function updateSummary(summary) {
  if (!summary) {
    return;
  }
  if (elements.connectionText) elements.connectionText.textContent = summary.connection_text;
  if (elements.analyzerOnStatus) elements.analyzerOnStatus.textContent = summary.analyzer_on_status;
  if (elements.analyzerErrorStatus) elements.analyzerErrorStatus.textContent = summary.analyzer_error_status;
  if (elements.analyzerFailureStatus) elements.analyzerFailureStatus.textContent = summary.analyzer_failure_status;
  updateAnalyzerAccents(summary);
  if (elements.titrolyzerAppStatusText) elements.titrolyzerAppStatusText.textContent = summary.titrolyzer_app_status;
  if (elements.titrolyzerAppStatusDetail) elements.titrolyzerAppStatusDetail.textContent = summary.titrolyzer_app_detail;
  if (elements.titrolyzerAppStatus) {
    elements.titrolyzerAppStatus.classList.toggle("running", Boolean(summary.titrolyzer_app_running));
    elements.titrolyzerAppStatus.classList.toggle("stopped", !summary.titrolyzer_app_running);
  }
  if (elements.sequenceStatus) elements.sequenceStatus.textContent = summary.sequence_status;
  if (elements.programAStatus) elements.programAStatus.textContent = summary.program_a_status;
  if (elements.programBStatus) elements.programBStatus.textContent = summary.program_b_status;
  if (elements.analyzerTimestamp) elements.analyzerTimestamp.textContent = summary.analyzer_timestamp;
  if (elements.r1Value) elements.r1Value.textContent = summary.r1_value;
  if (elements.r2Value) elements.r2Value.textContent = summary.r2_value;
  if (elements.r1Trend) elements.r1Trend.textContent = summary.r1_trend;
  if (elements.r2Trend) elements.r2Trend.textContent = summary.r2_trend;
  if (elements.readingAge) elements.readingAge.textContent = summary.reading_age;
}

function makeHistoryRow(row, type) {
  const tr = document.createElement("tr");
  tr.className = "history-row";
  tr.dataset.rowId = `${row.timestamp_key}_${type.toLowerCase()}`;
  tr.dataset.timestamp = row.timestamp_key;

  const tsTd = document.createElement("td");
  tsTd.textContent = row.timestamp;

  const typeTd = document.createElement("td");
  typeTd.textContent = type;

  const valTd = document.createElement("td");
  valTd.textContent = type === "R1"
    ? (row.r1_display || formatTitratorValue(row.r1))
    : (row.r2_display || formatTitratorValue(row.r2));

  const noteTd = document.createElement("td");
  noteTd.className = "note-cell";
  buildNoteCellContent(noteTd, row.note, row.timestamp_key);

  tr.appendChild(tsTd);
  tr.appendChild(typeTd);
  tr.appendChild(valTd);
  tr.appendChild(noteTd);

  tr.addEventListener("click", () => selectRow(tr.dataset.rowId));
  return tr;
}

function renderHistory(rows) {
  if (!elements.historyBody) return;
  if (appState.editingNote || appState.hoveringHistory) return;
  elements.historyBody.innerHTML = "";
  [...rows].reverse().forEach((row) => {
    if (row.r1 != null) elements.historyBody.appendChild(makeHistoryRow(row, "R1"));
    if (row.r2 != null) elements.historyBody.appendChild(makeHistoryRow(row, "R2"));
  });
  document.querySelectorAll(".history-row").forEach((row) => {
    row.classList.toggle("selected", row.dataset.rowId === appState.selectedRowId);
  });
}

function renderChlorine(rows) {
  if (!elements.chlorineBody) return;
  elements.chlorineBody.innerHTML = "";
  rows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${row.timestamp}</td>
      <td>${row.value}</td>
      <td>${row.unit}</td>
      <td>${row.note}</td>
    `;
    elements.chlorineBody.appendChild(tr);
  });
}

function renderCommands(rows) {
  if (!elements.commandLogList) return;
  elements.commandLogList.innerHTML = "";
  rows.forEach((row) => {
    const item = document.createElement("li");
    item.innerHTML = `<strong>${row.command}</strong><br><span>${row.time}</span>`;
    elements.commandLogList.appendChild(item);
  });
}

async function refreshDashboard(withReading = false) {
  const response = await fetch(`/api/dashboard?refresh=${withReading ? "1" : "0"}`);
  const payload = await response.json();
  appState.history = payload.history;
  updateSummary(payload.summary);
  renderHistory(payload.history);
  renderChlorine(payload.chlorine_measurements);
  renderCommands(payload.command_log);
  drawChart(payload.history);
}

async function sendCommand(command, address) {
  const confirmed = window.confirm(`Send "${command}"?`);
  if (!confirmed) {
    return;
  }

  const response = await fetch("/api/command", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ command, address }),
  });

  if (!response.ok) {
    window.alert("The command could not be sent.");
    return;
  }

  const payload = await response.json();
  appState.history = payload.history;
  updateSummary(payload.summary);
  renderHistory(payload.history);
  renderCommands(payload.command_log);
  drawChart(payload.history);
}

document.querySelectorAll("[data-command]").forEach((button) => {
  button.addEventListener("click", () => {
    sendCommand(button.dataset.command, button.dataset.address);
  });
});

const toggleSeries = document.getElementById("toggle-series");
const seriesCycle = { both: "r1", r1: "r2", r2: "both" };
const seriesButtonLabel = { both: "R1 only", r1: "R2 only", r2: "Show both" };
const seriesLegend = {
  both: '<i class="legend-swatch swatch-r1"></i>R1 <i class="legend-swatch swatch-r2"></i>R2',
  r1:   '<i class="legend-swatch swatch-r1"></i>R1',
  r2:   '<i class="legend-swatch swatch-r2"></i>R2',
};

if (toggleSeries) {
  toggleSeries.addEventListener("click", () => {
    appState.activeSeries = seriesCycle[appState.activeSeries];
    toggleSeries.textContent = seriesButtonLabel[appState.activeSeries];
    const legend = document.getElementById("active-series-legend");
    if (legend) legend.innerHTML = seriesLegend[appState.activeSeries];
    drawChart(appState.history);
  });
}

if (elements.addReading) {
  elements.addReading.addEventListener("click", () => {
    refreshDashboard(true);
  });
}

const saveManualReading = document.getElementById("save-manual-reading");
if (saveManualReading) {
  saveManualReading.addEventListener("click", () => {
    const r1 = parseFloat(document.getElementById("manual-r1").value);
    const r2 = parseFloat(document.getElementById("manual-r2").value);
    const note = document.getElementById("manual-note").value.trim();
    if (isNaN(r1) || isNaN(r2)) {
      window.alert("Please enter both R1 and R2 values.");
      return;
    }
    fetch("/api/manual-reading", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ r1, r2, note }),
    }).then(() => {
      document.getElementById("manual-r1").value = "";
      document.getElementById("manual-r2").value = "";
      document.getElementById("manual-note").value = "";
      refreshDashboard(false);
    });
  });
}


window.addEventListener("resize", () => {
  drawChart(appState.history);
});

setInterval(() => {
  refreshDashboard(true);
}, 4000);

setInterval(() => {
  refreshDashboard(false);
}, 1000);

drawChart(appState.history);
renderHistory(appState.history);

const historyPanel = document.querySelector(".panel-history");
if (historyPanel) {
  historyPanel.addEventListener("mouseenter", () => { appState.hoveringHistory = true; });
  historyPanel.addEventListener("mouseleave", () => { appState.hoveringHistory = false; });
}
