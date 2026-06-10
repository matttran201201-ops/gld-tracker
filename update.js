/**
 * GLD Tracker — downloads the official SPDR Gold Shares historical archive,
 * computes daily net holdings changes (buy/sell), and generates:
 *   - data.json       full processed dataset
 *   - gld_flows.csv   daily flows export (Excel-friendly)
 *   - dashboard.html  self-contained interactive dashboard
 *
 * Source: https://api.spdrgoldshares.com/api/v1/historical-archive?product=gld&exchange=NYSE&lang=en
 * Run: node update.js
 */
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const SRC_URL = "https://api.spdrgoldshares.com/api/v1/historical-archive?product=gld&exchange=NYSE&lang=en";
const DIR = __dirname;
const XLSX_FILE = path.join(DIR, "gld_archive.xlsx");

const OZ_PER_TONNE = 32150.7466;

async function download() {
  console.log("Downloading SPDR GLD archive...");
  const res = await fetch(SRC_URL, { headers: { "User-Agent": "Mozilla/5.0" } });
  if (!res.ok) throw new Error(`Download failed: HTTP ${res.status}`);
  const buf = Buffer.from(await res.arrayBuffer());
  fs.writeFileSync(XLSX_FILE, buf);
  console.log(`Saved ${(buf.length / 1024).toFixed(0)} KB -> gld_archive.xlsx`);
}

function parse() {
  const wb = XLSX.readFile(XLSX_FILE);
  const sheetName = wb.SheetNames.find((n) => /historical archive/i.test(n));
  if (!sheetName) throw new Error("Archive sheet not found. Sheets: " + wb.SheetNames.join(", "));
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });

  const records = [];
  for (const r of rows.slice(1)) {
    if (!r || !r[0]) continue;
    const totalOz = r[8], tonnes = r[9];
    if (typeof totalOz !== "number" || typeof tonnes !== "number") continue; // skip holidays/NYSE closed
    // parse "09-Jun-2026" without timezone shifts
    const m = String(r[0]).match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (!m) continue;
    const MONTHS = { Jan: "01", Feb: "02", Mar: "03", Apr: "04", May: "05", Jun: "06", Jul: "07", Aug: "08", Sep: "09", Oct: "10", Nov: "11", Dec: "12" };
    const mon = MONTHS[m[2]];
    if (!mon) continue;
    records.push({
      date: `${m[3]}-${mon}-${m[1].padStart(2, "0")}`,
      close: typeof r[1] === "number" ? r[1] : null,
      volume: typeof r[7] === "number" ? r[7] : null,
      totalOz,
      tonnes,
      nav: typeof r[10] === "number" ? r[10] : null,
    });
  }
  records.sort((a, b) => a.date.localeCompare(b.date));

  // daily net flow vs previous trading day
  for (let i = 0; i < records.length; i++) {
    const prev = records[i - 1];
    records[i].flowOz = prev ? +(records[i].totalOz - prev.totalOz).toFixed(2) : 0;
    records[i].flowTonnes = prev ? +(records[i].tonnes - prev.tonnes).toFixed(2) : 0;
  }
  return records;
}

function sumFlow(records, fromDate) {
  return +records.filter((r) => r.date >= fromDate).reduce((s, r) => s + r.flowTonnes, 0).toFixed(2);
}

function buildOutputs(records) {
  fs.writeFileSync(path.join(DIR, "data.json"), JSON.stringify(records));

  // CSV export
  const header = "Date,Tonnes,Total Ounces,Net Flow (tonnes),Net Flow (oz),GLD Close (USD),NAV (USD),Share Volume\n";
  const csv = header + records.map((r) =>
    [r.date, r.tonnes, r.totalOz, r.flowTonnes, r.flowOz, r.close ?? "", r.nav ?? "", r.volume ?? ""].join(",")
  ).join("\n");
  fs.writeFileSync(path.join(DIR, "gld_flows.csv"), csv);

  const last = records[records.length - 1];
  const year = last.date.slice(0, 4), month = last.date.slice(0, 7);
  const summary = {
    updated: new Date().toISOString(),
    latestDate: last.date,
    tonnes: last.tonnes,
    totalOz: last.totalOz,
    dayFlowTonnes: last.flowTonnes,
    mtdTonnes: sumFlow(records, month + "-01"),
    ytdTonnes: sumFlow(records, year + "-01-01"),
    close: last.close,
  };

  const html = buildHtml(records, summary);
  fs.writeFileSync(path.join(DIR, "dashboard.html"), html);
  return summary;
}

function buildHtml(records, s) {
  const dataJs = JSON.stringify(records);
  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>GLD Holdings Tracker — SPDR Gold Trust</title>
<meta name="viewport" content="width=device-width, initial-scale=1">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>
<style>
  :root { --bg:#0f1419; --card:#1a2129; --text:#e6e8ea; --muted:#8b949e; --green:#2ea05f; --red:#d9534f; --gold:#d4a843; --border:#2d3640; }
  * { box-sizing:border-box; margin:0; }
  body { background:var(--bg); color:var(--text); font:14px/1.5 "Segoe UI",system-ui,sans-serif; padding:24px; }
  h1 { font-size:20px; font-weight:600; } h1 .gold { color:var(--gold); }
  .sub { color:var(--muted); font-size:12px; margin:4px 0 20px; }
  .cards { display:grid; grid-template-columns:repeat(auto-fit,minmax(190px,1fr)); gap:12px; margin-bottom:20px; }
  .card { background:var(--card); border:1px solid var(--border); border-radius:10px; padding:14px 16px; }
  .card .label { color:var(--muted); font-size:11px; text-transform:uppercase; letter-spacing:.5px; }
  .card .value { font-size:24px; font-weight:700; margin-top:4px; }
  .card .detail { color:var(--muted); font-size:12px; margin-top:2px; }
  .pos { color:var(--green); } .neg { color:var(--red); }
  .panel { background:var(--card); border:1px solid var(--border); border-radius:10px; padding:18px; margin-bottom:20px; }
  .panel h2 { font-size:14px; font-weight:600; margin-bottom:10px; }
  .ranges { display:flex; gap:6px; margin-bottom:14px; flex-wrap:wrap; }
  .ranges button { background:transparent; color:var(--muted); border:1px solid var(--border); border-radius:6px; padding:4px 12px; cursor:pointer; font-size:12px; }
  .ranges button.active { background:var(--gold); color:#1a1408; border-color:var(--gold); font-weight:600; }
  .chartbox { position:relative; height:340px; }
  table { width:100%; border-collapse:collapse; font-size:13px; }
  th { text-align:right; color:var(--muted); font-weight:600; padding:8px 10px; border-bottom:1px solid var(--border); font-size:11px; text-transform:uppercase; letter-spacing:.4px; }
  td { text-align:right; padding:7px 10px; border-bottom:1px solid var(--border); font-variant-numeric:tabular-nums; }
  th:first-child, td:first-child { text-align:left; }
  tr:hover td { background:rgba(212,168,67,.05); }
  .src { color:var(--muted); font-size:11px; margin-top:8px; }
  a { color:var(--gold); }
</style>
</head>
<body>
<h1><span class="gold">GLD</span> Holdings Tracker — SPDR&reg; Gold Trust</h1>
<div class="sub">Data: official SPDR Gold Shares historical archive &middot; Latest trading day: <b>${s.latestDate}</b> &middot; Generated: ${s.updated.slice(0, 16).replace("T", " ")} UTC</div>

<div class="cards">
  <div class="card"><div class="label">Total holdings</div><div class="value">${s.tonnes.toLocaleString("en-US", { minimumFractionDigits: 2 })} t</div><div class="detail">${s.totalOz.toLocaleString("en-US", { maximumFractionDigits: 0 })} oz</div></div>
  <div class="card"><div class="label">Last day net flow</div><div class="value ${s.dayFlowTonnes >= 0 ? "pos" : "neg"}">${s.dayFlowTonnes >= 0 ? "+" : ""}${s.dayFlowTonnes.toFixed(2)} t</div><div class="detail">${s.dayFlowTonnes >= 0 ? "net BUY" : "net SELL"} on ${s.latestDate}</div></div>
  <div class="card"><div class="label">Month-to-date flow</div><div class="value ${s.mtdTonnes >= 0 ? "pos" : "neg"}">${s.mtdTonnes >= 0 ? "+" : ""}${s.mtdTonnes.toFixed(2)} t</div><div class="detail">since ${s.latestDate.slice(0, 7)}-01</div></div>
  <div class="card"><div class="label">Year-to-date flow</div><div class="value ${s.ytdTonnes >= 0 ? "pos" : "neg"}">${s.ytdTonnes >= 0 ? "+" : ""}${s.ytdTonnes.toFixed(2)} t</div><div class="detail">since ${s.latestDate.slice(0, 4)}-01-01</div></div>
  <div class="card"><div class="label">GLD close</div><div class="value">$${(s.close ?? 0).toFixed(2)}</div><div class="detail">NYSE Arca</div></div>
</div>

<div class="panel">
  <h2>Total holdings (tonnes) vs GLD price</h2>
  <div class="ranges" id="ranges1"></div>
  <div class="chartbox"><canvas id="holdChart"></canvas></div>
</div>

<div class="panel">
  <h2>Daily net flow (tonnes) — green = net buy, red = net sell</h2>
  <div class="ranges" id="ranges2"></div>
  <div class="chartbox"><canvas id="flowChart"></canvas></div>
</div>

<div class="panel">
  <h2>Last 30 trading days</h2>
  <table id="tbl">
    <thead><tr><th>Date</th><th>Tonnes</th><th>Total oz</th><th>Net flow (t)</th><th>Net flow (oz)</th><th>GLD close</th><th>Volume</th></tr></thead>
    <tbody></tbody>
  </table>
  <div class="src">Source: <a href="https://www.spdrgoldshares.com/usa/historical-data/" target="_blank">spdrgoldshares.com — historical data</a>. Net flow = change in Total Ounces of Gold in the Trust vs previous trading day.</div>
</div>

<script>
const DATA = ${dataJs};
const RANGES = { "1M": 22, "3M": 66, "6M": 132, "1Y": 252, "3Y": 756, "5Y": 1260, "All": Infinity };
const fmt = (n, d = 2) => n.toLocaleString("en-US", { minimumFractionDigits: d, maximumFractionDigits: d });

function slice(n) { return n === Infinity ? DATA : DATA.slice(-n); }

let holdChart, flowChart;
function drawHold(rangeKey) {
  const d = slice(RANGES[rangeKey]);
  if (holdChart) holdChart.destroy();
  holdChart = new Chart(document.getElementById("holdChart"), {
    type: "line",
    data: { labels: d.map(r => r.date), datasets: [
      { label: "Holdings (t)", data: d.map(r => r.tonnes), borderColor: "#d4a843", backgroundColor: "rgba(212,168,67,.12)", fill: true, pointRadius: 0, borderWidth: 1.6, yAxisID: "y", tension: .1 },
      { label: "GLD close ($)", data: d.map(r => r.close), borderColor: "#5b9bd5", pointRadius: 0, borderWidth: 1.2, yAxisID: "y2", tension: .1 },
    ]},
    options: { responsive: true, maintainAspectRatio: false, interaction: { mode: "index", intersect: false },
      plugins: { legend: { labels: { color: "#8b949e", boxWidth: 12 } } },
      scales: {
        x: { ticks: { color: "#8b949e", maxTicksLimit: 10 }, grid: { color: "#2d3640" } },
        y: { title: { display: true, text: "tonnes", color: "#8b949e" }, ticks: { color: "#d4a843" }, grid: { color: "#2d3640" } },
        y2: { position: "right", title: { display: true, text: "USD", color: "#8b949e" }, ticks: { color: "#5b9bd5" }, grid: { display: false } },
      } }
  });
}
function drawFlow(rangeKey) {
  const d = slice(RANGES[rangeKey]);
  if (flowChart) flowChart.destroy();
  flowChart = new Chart(document.getElementById("flowChart"), {
    type: "bar",
    data: { labels: d.map(r => r.date), datasets: [{ label: "Net flow (t)", data: d.map(r => r.flowTonnes),
      backgroundColor: d.map(r => r.flowTonnes >= 0 ? "rgba(46,160,95,.8)" : "rgba(217,83,79,.8)") }] },
    options: { responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => (c.raw >= 0 ? "+" : "") + fmt(c.raw) + " t (" + (c.raw >= 0 ? "net buy" : "net sell") + ")" } } },
      scales: { x: { ticks: { color: "#8b949e", maxTicksLimit: 10 }, grid: { display: false } },
                y: { title: { display: true, text: "tonnes", color: "#8b949e" }, ticks: { color: "#8b949e" }, grid: { color: "#2d3640" } } } }
  });
}
function buildRanges(el, draw, def) {
  Object.keys(RANGES).forEach(k => {
    const b = document.createElement("button"); b.textContent = k;
    if (k === def) b.classList.add("active");
    b.onclick = () => { el.querySelectorAll("button").forEach(x => x.classList.remove("active")); b.classList.add("active"); draw(k); };
    el.appendChild(b);
  });
  draw(def);
}
buildRanges(document.getElementById("ranges1"), drawHold, "1Y");
buildRanges(document.getElementById("ranges2"), drawFlow, "3M");

const tbody = document.querySelector("#tbl tbody");
DATA.slice(-30).reverse().forEach(r => {
  const tr = document.createElement("tr");
  const cls = r.flowTonnes >= 0 ? "pos" : "neg", sign = r.flowTonnes >= 0 ? "+" : "";
  tr.innerHTML = "<td>" + r.date + "</td><td>" + fmt(r.tonnes) + "</td><td>" + fmt(r.totalOz, 0) + "</td>" +
    "<td class='" + cls + "'>" + sign + fmt(r.flowTonnes) + "</td><td class='" + cls + "'>" + sign + fmt(r.flowOz, 0) + "</td>" +
    "<td>" + (r.close != null ? "$" + fmt(r.close) : "—") + "</td><td>" + (r.volume != null ? fmt(r.volume, 0) : "—") + "</td>";
  tbody.appendChild(tr);
});
</script>
</body>
</html>`;
}

(async () => {
  await download();
  const records = parse();
  const s = buildOutputs(records);
  console.log(`Parsed ${records.length} trading days (${records[0].date} -> ${s.latestDate})`);
  console.log(`Holdings: ${s.tonnes} t | day flow: ${s.dayFlowTonnes >= 0 ? "+" : ""}${s.dayFlowTonnes} t | MTD: ${s.mtdTonnes} t | YTD: ${s.ytdTonnes} t`);
  console.log("Outputs: dashboard.html, gld_flows.csv, data.json");
})().catch((e) => { console.error("ERROR:", e.message); process.exit(1); });
