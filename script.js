const historyBody = document.getElementById("history-body");
const commandLogList = document.getElementById("command-log-list");
const chartCanvas = document.getElementById("trend-chart");
const chartContext = chartCanvas.getContext("2d");

const elements = {
  sequenceStatus: document.getElementById("sequence-status"),
  programAStatus: document.getElementById("program-a-status"),
  programBStatus: document.getElementById("program-b-status"),
  analyzerTimestamp: document.getElementById("analyzer-timestamp"),
  r1Value: document.getElementById("r1-value"),
  r2Value: document.getElementById("r2-value"),
  r1Trend: document.getElementById("r1-trend"),
  r2Trend: document.getElementById("r2-trend"),
  readingAge: document.getElementById("reading-age"),
  connectionText: document.getElementById("connection-text"),
  toggleSeries: document.getElementById("toggle-series"),
};

const appState = {
  showR2: true,
  connectionOk: true,
  sequence: "Idle",
  programA: "Standby",
  programB: "Standby",
  history: [],
  commands: [],
};

function seededReading(index) {
  const base = Date.now() - (11 - index) * 4 * 60 * 1000;
  const r1 = 62 + Math.sin(index / 1.8) * 7 + index * 0.35;
  const r2 = 43 + Math.cos(index / 1.6) * 5 + index * 0.25;

  return {
    timestamp: new Date(base),
    r1: Number(r1.toFixed(2)),
    r2: Number(r2.toFixed(2)),
    sequence: index > 8 ? "Running" : "Idle",
    programA: index % 4 === 0 ? "Active" : "Standby",
    programB: index % 5 === 0 ? "Active" : "Standby",
  };
}

function formatTime(date) {
  return new Intl.DateTimeFormat("en-US", {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
    second: "2-digit",
  }).format(date);
}

function updateSummary() {
  const latest = appState.history[appState.history.length - 1];
  const previous = appState.history[appState.history.length - 2] || latest;
  const ageSeconds = Math.max(0, Math.round((Date.now() - latest.timestamp.getTime()) / 1000));

  elements.sequenceStatus.textContent = appState.sequence;
  elements.programAStatus.textContent = appState.programA;
  elements.programBStatus.textContent = appState.programB;
  elements.analyzerTimestamp.textContent = latest.timestamp.toLocaleTimeString();
  elements.r1Value.textContent = latest.r1.toFixed(2);
  elements.r2Value.textContent = latest.r2.toFixed(2);
  elements.r1Trend.textContent = latest.r1 > previous.r1 ? "Upward trend" : latest.r1 < previous.r1 ? "Slight dip" : "Stable";
  elements.r2Trend.textContent = latest.r2 > previous.r2 ? "Upward trend" : latest.r2 < previous.r2 ? "Slight dip" : "Stable";
  elements.readingAge.textContent = `${ageSeconds} sec`;
  elements.connectionText.textContent = appState.connectionOk ? "Connected to simulator" : "Connection interrupted";
}

function renderHistory() {
  historyBody.innerHTML = "";

  [...appState.history]
    .reverse()
    .forEach((row) => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${formatTime(row.timestamp)}</td>
        <td>${row.r1.toFixed(2)}</td>
        <td>${row.r2.toFixed(2)}</td>
        <td>${row.sequence}</td>
        <td>${row.programA}</td>
        <td>${row.programB}</td>
      `;
      historyBody.appendChild(tr);
    });
}

function renderCommands() {
  commandLogList.innerHTML = "";

  appState.commands.slice(0, 6).forEach((entry) => {
    const item = document.createElement("li");
    item.innerHTML = `<strong>${entry.command}</strong><br><span>${entry.time} • address ${entry.address}</span>`;
    commandLogList.appendChild(item);
  });
}

function drawChart() {
  const { width, height } = chartCanvas;
  chartContext.clearRect(0, 0, width, height);

  const rows = appState.history;
  const values = rows.flatMap((row) => [row.r1, row.r2]);
  const min = Math.min(...values) - 5;
  const max = Math.max(...values) + 5;
  const left = 52;
  const right = width - 20;
  const top = 24;
  const bottom = height - 42;

  chartContext.strokeStyle = "rgba(31, 41, 51, 0.12)";
  chartContext.lineWidth = 1;

  for (let i = 0; i <= 4; i += 1) {
    const y = top + ((bottom - top) / 4) * i;
    chartContext.beginPath();
    chartContext.moveTo(left, y);
    chartContext.lineTo(right, y);
    chartContext.stroke();

    const value = (max - ((max - min) / 4) * i).toFixed(0);
    chartContext.fillStyle = "#6b7280";
    chartContext.font = "14px Bahnschrift";
    chartContext.fillText(value, 10, y + 4);
  }

  const toPoint = (value, index) => {
    const x = left + ((right - left) / Math.max(rows.length - 1, 1)) * index;
    const y = bottom - ((value - min) / (max - min)) * (bottom - top);
    return { x, y };
  };

  const drawSeries = (key, color, fillColor) => {
    chartContext.beginPath();
    rows.forEach((row, index) => {
      const point = toPoint(row[key], index);
      if (index === 0) {
        chartContext.moveTo(point.x, point.y);
      } else {
        chartContext.lineTo(point.x, point.y);
      }
    });
    chartContext.strokeStyle = color;
    chartContext.lineWidth = 4;
    chartContext.stroke();

    rows.forEach((row, index) => {
      const point = toPoint(row[key], index);
      chartContext.beginPath();
      chartContext.arc(point.x, point.y, 4.5, 0, Math.PI * 2);
      chartContext.fillStyle = fillColor;
      chartContext.fill();
    });
  };

  drawSeries("r1", "#1e5eff", "rgba(30, 94, 255, 0.22)");

  if (appState.showR2) {
    drawSeries("r2", "#c07b12", "rgba(192, 123, 18, 0.22)");
  }

  chartContext.fillStyle = "#6b7280";
  chartContext.font = "13px Bahnschrift";
  rows.forEach((row, index) => {
    const point = toPoint(row.r1, index);
    chartContext.fillText(row.timestamp.toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }), point.x - 18, height - 16);
  });
}

function addReading() {
  const last = appState.history[appState.history.length - 1];
  const step = appState.sequence === "Running" ? 1.4 : 0.5;
  const modifier = Math.random() * 1.6 - 0.8;
  const newRow = {
    timestamp: new Date(),
    r1: Number((last.r1 + step + modifier).toFixed(2)),
    r2: Number((last.r2 + modifier * 1.2 + (appState.programB === "Active" ? 0.7 : 0.2)).toFixed(2)),
    sequence: appState.sequence,
    programA: appState.programA,
    programB: appState.programB,
  };

  appState.history = [...appState.history.slice(-11), newRow];
  updateSummary();
  renderHistory();
  drawChart();
}

function logCommand(command, address) {
  appState.commands.unshift({
    command,
    address,
    time: new Date().toLocaleTimeString(),
  });

  renderCommands();
}

function handleCommand(command, address) {
  const confirmed = window.confirm(`Send "${command}" to coil address ${address}?`);
  if (!confirmed) {
    return;
  }

  if (command === "Start Sequence") {
    appState.sequence = "Running";
  }

  if (command === "Stop Sequence") {
    appState.sequence = "Stopped";
  }

  if (command === "Break Sequence") {
    appState.sequence = "Paused";
  }

  if (command === "Trigger Program A") {
    appState.programA = "Active";
  }

  if (command === "Trigger Program B") {
    appState.programB = "Active";
  }

  logCommand(command, address);
  addReading();
}

function exportData() {
  const header = ["Timestamp", "R1", "R2", "Sequence", "Program A", "Program B"];
  const rows = appState.history.map((row) => [
    formatTime(row.timestamp),
    row.r1.toFixed(2),
    row.r2.toFixed(2),
    row.sequence,
    row.programA,
    row.programB,
  ]);
  const csv = [header, ...rows].map((row) => row.join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = "team-operations-export.csv";
  link.click();
  URL.revokeObjectURL(url);
}

function initialize() {
  appState.history = Array.from({ length: 12 }, (_, index) => seededReading(index));
  appState.sequence = appState.history[appState.history.length - 1].sequence;
  appState.programA = appState.history[appState.history.length - 1].programA;
  appState.programB = appState.history[appState.history.length - 1].programB;
  appState.commands = [
    { command: "Dashboard initialized", address: "--", time: new Date().toLocaleTimeString() },
    { command: "Monitoring active", address: "--", time: new Date().toLocaleTimeString() },
  ];

  updateSummary();
  renderHistory();
  renderCommands();
  drawChart();
}

document.querySelectorAll("[data-command]").forEach((button) => {
  button.addEventListener("click", () => {
    handleCommand(button.dataset.command, button.dataset.address);
  });
});

document.getElementById("add-reading").addEventListener("click", addReading);
document.getElementById("export-data").addEventListener("click", exportData);

elements.toggleSeries.addEventListener("click", () => {
  appState.showR2 = !appState.showR2;
  elements.toggleSeries.textContent = appState.showR2 ? "Hide R2" : "Show R2";
  drawChart();
});

setInterval(() => {
  addReading();
}, 4000);

setInterval(() => {
  updateSummary();
}, 1000);

window.addEventListener("resize", drawChart);

initialize();
