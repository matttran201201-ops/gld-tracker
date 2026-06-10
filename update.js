/**
 * Gold ETF Holdings Tracker — downloads official daily data for:
 *   GLD  — SPDR Gold Shares          (SPDR archive API, XLSX)
 *   GLDM — SPDR Gold MiniShares      (SPDR archive API, XLSX)
 *   SGLN — iShares Physical Gold ETC (iShares fund file, SpreadsheetML XML)
 * computes daily net holdings changes (buy/sell) per fund, and generates:
 *   - data.json       processed dataset { gld: [...], gldm: [...], sgln: [...] }
 *   - gld_flows.csv   combined daily flows export (wide format)
 *
 * SGLN tonnes are derived as Securities In Issue x metal entitlement, where the
 * entitlement history is reconstructed from the Fund/Benchmark return ratio
 * (the ratio decays at exactly the TER rate), anchored to the published
 * "Daily Metal Entitlement per Security" and verified against the published
 * "Tonnes in Trust".
 *
 * Run: node update.js
 */
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const DIR = __dirname;
const OZ_PER_TONNE = 32150.7466;
const UA = { "User-Agent": "Mozilla/5.0" };

const SPDR_URL = (product) =>
  `https://api.spdrgoldshares.com/api/v1/historical-archive?product=${product}&exchange=NYSE&lang=en`;
const SGLN_URL =
  "https://www.ishares.com/uk/individual/en/products/258441/ishares-physical-gold-etc-fund/1535604580409.ajax?fileType=xls&fileName=iShares-Physical-Gold-ETC_fund&dataType=fund";

const MONTHS = { Jan: "01", Feb: "02", Mar: "03", Apr: "04", May: "05", Jun: "06", Jul: "07", Aug: "08", Sep: "09", Oct: "10", Nov: "11", Dec: "12" };

async function fetchBuffer(url) {
  const res = await fetch(url, { headers: UA });
  if (!res.ok) throw new Error(`Download failed (${url.slice(0, 60)}...): HTTP ${res.status}`);
  return Buffer.from(await res.arrayBuffer());
}

function addFlows(records) {
  records.sort((a, b) => a.date.localeCompare(b.date));
  for (let i = 0; i < records.length; i++) {
    const prev = records[i - 1];
    records[i].flowOz = prev ? +(records[i].totalOz - prev.totalOz).toFixed(2) : 0;
    records[i].flowTonnes = prev ? +(records[i].tonnes - prev.tonnes).toFixed(2) : 0;
  }
  return records;
}

// ---------- SPDR funds (GLD, GLDM) ----------
async function fetchSpdr(product) {
  const buf = await fetchBuffer(SPDR_URL(product));
  const wb = XLSX.read(buf, { type: "buffer" });
  const sheetName = wb.SheetNames.find((n) => /historical archive/i.test(n));
  if (!sheetName) throw new Error(`${product}: archive sheet not found`);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1 });

  const records = [];
  for (const r of rows.slice(1)) {
    if (!r || !r[0]) continue;
    const totalOz = r[8], tonnes = r[9];
    if (typeof totalOz !== "number" || typeof tonnes !== "number") continue; // holidays
    const m = String(r[0]).match(/^(\d{1,2})-([A-Za-z]{3})-(\d{4})$/);
    if (!m || !MONTHS[m[2]]) continue;
    records.push({
      date: `${m[3]}-${MONTHS[m[2]]}-${m[1].padStart(2, "0")}`,
      close: typeof r[1] === "number" ? r[1] : null,
      volume: typeof r[7] === "number" ? r[7] : null,
      totalOz, tonnes,
    });
  }
  return addFlows(records);
}

// ---------- iShares SGLN (SpreadsheetML 2003 XML) ----------
function parseSpreadsheetML(xml) {
  const sheets = {};
  const wsRe = /<ss:Worksheet ss:Name="([^"]+)"[\s\S]*?(?=<ss:Worksheet |<\/ss:Workbook>)/g;
  let ws;
  while ((ws = wsRe.exec(xml))) {
    const rows = [];
    const rowRe = /<ss:Row[^>]*>([\s\S]*?)<\/ss:Row>/g;
    let row;
    while ((row = rowRe.exec(ws[0]))) {
      const cells = [];
      const cellRe = /<ss:Data[^>]*ss:Type="(String|Number)"[^>]*>([\s\S]*?)<\/ss:Data>/g;
      let c;
      while ((c = cellRe.exec(row[1]))) {
        cells.push(c[1] === "Number" ? parseFloat(c[2]) : c[2].trim());
      }
      rows.push(cells);
    }
    sheets[ws[1]] = rows;
  }
  return sheets;
}

async function fetchSgln() {
  const buf = await fetchBuffer(SGLN_URL);
  const sheets = parseSpreadsheetML(buf.toString("utf8"));
  const ov = (sheets["Overview"] || []).flat();
  const hist = sheets["Historical"] || [];
  if (!hist.length) throw new Error("SGLN: Historical sheet not found");

  const lookup = (label) => {
    const i = ov.findIndex((x) => typeof x === "string" && x.includes(label));
    return i >= 0 ? ov[i + 1] : null;
  };
  const entitlementNow = parseFloat(lookup("Daily Metal Entitlement"));
  const publishedTonnes = parseFloat(String(lookup("Tonnes in Trust")).replace(/,/g, ""));
  if (!entitlementNow) throw new Error("SGLN: metal entitlement not found in Overview");

  // Historical columns: As Of, Currency, NAV, Securities In Issue, Net Assets, Fund Return Series, Benchmark Return Series
  const header = hist[0].map(String);
  const col = (re) => header.findIndex((h) => re.test(h));
  const cDate = col(/As Of/i), cShares = col(/Securities In Issue/i),
        cNav = col(/^NAV/i), cFund = col(/Fund Return/i), cBench = col(/Benchmark Return/i);

  const raw = [];
  for (const r of hist.slice(1)) {
    const m = String(r[cDate] || "").match(/^(\d{1,2})\/([A-Za-z]{3})\/(\d{4})$/);
    const shares = r[cShares];
    if (!m || !MONTHS[m[2]] || typeof shares !== "number") continue;
    raw.push({
      date: `${m[3]}-${MONTHS[m[2]]}-${m[1].padStart(2, "0")}`,
      shares,
      nav: typeof r[cNav] === "number" ? r[cNav] : null,
      ratio: typeof r[cFund] === "number" && typeof r[cBench] === "number" && r[cBench] !== 0
        ? r[cFund] / r[cBench] : null,
    });
  }
  raw.sort((a, b) => a.date.localeCompare(b.date));

  // anchor: latest row with a valid Fund/Benchmark ratio
  const anchor = [...raw].reverse().find((r) => r.ratio != null);
  if (!anchor) throw new Error("SGLN: no Fund/Benchmark return data");

  const records = [];
  let lastRatio = null;
  for (const r of raw) {
    if (r.ratio != null) lastRatio = r.ratio;
    // entitlement(t) = entitlement_now * ratio(t) / ratio(anchor); fee drag is the only divergence
    const ent = entitlementNow * ((lastRatio ?? anchor.ratio) / anchor.ratio);
    const totalOz = r.shares * ent;
    records.push({
      date: r.date,
      close: r.nav, // NAV per security (USD) — SGLN trades on LSE; NAV used in lieu of close
      volume: null,
      totalOz: +totalOz.toFixed(2),
      tonnes: +(totalOz / OZ_PER_TONNE).toFixed(2),
    });
  }
  addFlows(records);

  // sanity check vs published tonnes (latest day)
  const last = records[records.length - 1];
  if (publishedTonnes && Math.abs(last.tonnes - publishedTonnes) / publishedTonnes > 0.005) {
    console.warn(`WARN SGLN: computed ${last.tonnes} t vs published ${publishedTonnes} t`);
  }
  return records;
}

// ---------- outputs ----------
function sumFlow(records, fromDate) {
  return +records.filter((r) => r.date >= fromDate).reduce((s, r) => s + r.flowTonnes, 0).toFixed(2);
}

function buildCsv(funds) {
  const dates = [...new Set(Object.values(funds).flat().map((r) => r.date))].sort();
  const byDate = {};
  for (const [key, recs] of Object.entries(funds)) {
    byDate[key] = Object.fromEntries(recs.map((r) => [r.date, r]));
  }
  const cols = ["GLD", "GLDM", "SGLN"];
  const header = "Date," + cols.map((c) => `${c} Tonnes,${c} Flow (t)`).join(",") + ",Total Tonnes,Total Flow (t)\n";
  const lines = dates.map((d) => {
    let totalT = 0, totalF = 0, hasAny = false;
    const cells = [];
    let lastKnown = {};
    for (const c of cols) {
      const r = byDate[c.toLowerCase()][d];
      cells.push(r ? r.tonnes : "", r ? r.flowTonnes : "");
      if (r) { totalF += r.flowTonnes; hasAny = true; }
    }
    // total tonnes = sum of last known per fund up to d (fill-forward)
    for (const c of cols) {
      const recs = funds[c.toLowerCase()];
      let v = null;
      for (let i = recs.length - 1; i >= 0; i--) if (recs[i].date <= d) { v = recs[i].tonnes; break; }
      if (v != null) totalT += v;
    }
    return [d, ...cells, +totalT.toFixed(2), +totalF.toFixed(2)].join(",");
  });
  return header + lines.join("\n");
}

(async () => {
  console.log("Downloading GLD, GLDM (SPDR) and SGLN (iShares)...");
  const [gld, gldm, sgln] = await Promise.all([fetchSpdr("gld"), fetchSpdr("gldm"), fetchSgln()]);
  const funds = { gld, gldm, sgln };

  fs.writeFileSync(path.join(DIR, "data.json"), JSON.stringify({
    updated: new Date().toISOString(),
    funds,
  }));
  fs.writeFileSync(path.join(DIR, "gld_flows.csv"), buildCsv(funds));

  for (const [k, recs] of Object.entries(funds)) {
    const last = recs[recs.length - 1];
    const ytd = sumFlow(recs, last.date.slice(0, 4) + "-01-01");
    console.log(`${k.toUpperCase().padEnd(4)} ${recs.length} days  ${recs[0].date} -> ${last.date}  ` +
      `holdings ${last.tonnes} t  day ${last.flowTonnes >= 0 ? "+" : ""}${last.flowTonnes} t  YTD ${ytd >= 0 ? "+" : ""}${ytd} t`);
  }
  console.log("Outputs: data.json, gld_flows.csv");
})().catch((e) => { console.error("ERROR:", e.message); process.exit(1); });
